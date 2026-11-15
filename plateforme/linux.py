"""Fonctions propres à Linux (session X11) : clavier, presse-papier, fenêtre active, overlay.

Tout passe par le serveur X : raccourci global, collage simulé et forme de l'overlay. Une
session Wayland ne permet ni l'écoute du clavier de tout le système ni l'envoi de touches
aux autres applications ; l'application le signale au démarrage.
"""
import math
import os
import shutil
import subprocess
import threading

import pyperclip
import Xlib.display
import Xlib.X
import Xlib.XK
from Xlib.ext import shape
from pynput import keyboard as pynput_keyboard
from pynput.keyboard._xorg import Listener as _XorgListener

NAME = "Linux"
SETUP_SCRIPT = "setup.sh"

# Une seule connexion au serveur X pour ces fonctions, toutes appelées depuis Tk.
_display = None
_display_lock = threading.Lock()


def _x():
    global _display
    if _display is None:
        _display = Xlib.display.Display()
    return _display


# ── Clavier ──────────────────────────────────────────────────────────────────
# Une touche est désignée par son keycode X : il suit la touche physique, donc le raccourci
# reste le même en AZERTY comme en QWERTY. Keycode X = code evdev du noyau + 8.

MODIFIER_LABELS = {'ctrl': 'Ctrl', 'alt': 'Alt', 'shift': 'Maj', 'win': 'Super'}

_KEY_LABELS = {
    9: 'Échap', 22: 'Retour arrière', 23: 'Tab', 36: 'Entrée', 65: 'Espace',
    66: 'Verr. Maj', 107: 'Impr. écran', 127: 'Pause', 118: 'Inser', 119: 'Suppr',
    110: 'Origine', 115: 'Fin', 112: 'Page préc.', 117: 'Page suiv.',
    111: '↑', 113: '←', 114: '→', 116: '↓',
    63: 'Pavé num *', 82: 'Pavé num -', 86: 'Pavé num +', 91: 'Pavé num .',
    106: 'Pavé num /', 104: 'Pavé num Entrée',
}
# Rangée des chiffres, nommée par le chiffre comme sur un clavier QWERTY ou AZERTY avec Maj.
_KEY_LABELS.update({10 + i: str((i + 1) % 10) for i in range(10)})
_KEYPAD_DIGITS = {79: 7, 80: 8, 81: 9, 83: 4, 84: 5, 85: 6, 87: 1, 88: 2, 89: 3, 90: 0}
_KEY_LABELS.update({code: 'Pavé num %d' % n for code, n in _KEYPAD_DIGITS.items()})
_FUNCTION_KEYS = {67 + i: i + 1 for i in range(10)}
_FUNCTION_KEYS.update({95: 11, 96: 12})
_FUNCTION_KEYS.update({191 + i: 13 + i for i in range(12)})               # F13-F24
_KEY_LABELS.update({code: 'F%d' % n for code, n in _FUNCTION_KEYS.items()})

SAFE_BARE_CODES = set(_FUNCTION_KEYS)

DEFAULT_HOTKEY = {
    'modifiers': ['ctrl', 'alt'],
    # 9 de la rangée du haut (18) et du pavé numérique (81).
    'vks': [18, 81],
}

_MODIFIER_KEYSYMS = ('Control_L', 'Control_R', 'Alt_L', 'Alt_R', 'Shift_L', 'Shift_R',
                     'Super_L', 'Super_R', 'Meta_L', 'Meta_R', 'ISO_Level3_Shift')
_modifier_codes = None


def key_label(code):
    """Nom d'une touche ; une lettre prend le nom qu'elle a dans la disposition du clavier."""
    if code in _KEY_LABELS:
        return _KEY_LABELS[code]
    try:
        with _display_lock:
            keysym = _x().keycode_to_keysym(code, 0)
        name = Xlib.XK.keysym_to_string(keysym) if keysym else None
    except Exception:
        name = None
    if name and len(name) == 1:
        return name.upper()
    return 'Touche %d' % code


class _Listener(_XorgListener):
    """Écoute de pynput qui garde aussi le keycode X de la touche, que pynput ne transmet pas."""
    keycode = None

    def _handle_message(self, display, event, injected):
        self.keycode = event.detail
        try:
            super()._handle_message(display, event, injected)
        finally:
            self.keycode = None


def start_key_listener(on_press, on_release):
    """Écoute le clavier de tout le système ; on_press(touche, code) et on_release(touche, code)."""
    listener = None

    def press(key, *_):
        on_press(key, listener.keycode)

    def release(key, *_):
        on_release(key, listener.keycode)

    listener = _Listener(on_press=press, on_release=release)
    listener.start()
    return listener


def keys_held(codes):
    """Un modificateur ou l'une des touches `codes` est-il enfoncé en ce moment ?"""
    global _modifier_codes
    with _display_lock:
        display = _x()
        if _modifier_codes is None:
            _modifier_codes = {code for sym in _MODIFIER_KEYSYMS
                               for code, _ in display.keysym_to_keycodes(Xlib.XK.string_to_keysym(sym))}
        keymap = display.query_keymap()
    return any(keymap[code // 8] & (1 << (code % 8)) for code in _modifier_codes | set(codes))


_controller = pynput_keyboard.Controller()


def send_paste(terminal):
    """Envoie Ctrl+V, ou Ctrl+Maj+V pour un terminal."""
    keys = [pynput_keyboard.Key.ctrl]
    if terminal:
        keys.append(pynput_keyboard.Key.shift)
    for key in keys:
        _controller.press(key)
    try:
        _controller.press('v')
        _controller.release('v')
    finally:
        for key in reversed(keys):
            _controller.release(key)


# ── Fenêtre active ───────────────────────────────────────────────────────────

# Classes de fenêtre (WM_CLASS, en minuscules) des terminaux qui collent avec Ctrl+Maj+V. Les
# terminaux intégrés des éditeurs (Cursor, VS Code) partagent la classe de l'éditeur : pour
# eux, l'option manuelle.
_TERMINAL_CLASSES = {
    'gnome-terminal', 'gnome-terminal-server', 'mate-terminal', 'xfce4-terminal', 'tilix',
    'terminator', 'konsole', 'kitty', 'alacritty', 'wezterm', 'org.wezfurlong.wezterm',
    'guake', 'ptyxis', 'org.gnome.ptyxis', 'kgx', 'org.gnome.console', 'foot', 'lxterminal',
    'qterminal', 'sakura', 'terminology', 'tabby', 'hyper', 'ghostty', 'com.mitchellh.ghostty',
    'warp',
}


def terminal_focused():
    """La fenêtre au premier plan est-elle un terminal connu ?"""
    with _display_lock:
        display = _x()
        root = display.screen().root
        active = root.get_full_property(display.intern_atom('_NET_ACTIVE_WINDOW'),
                                        Xlib.X.AnyPropertyType)
        if not active or not active.value or not active.value[0]:
            return False
        window = display.create_resource_object('window', active.value[0])
        wm_class = window.get_wm_class() or ()
    return any(name.lower() in _TERMINAL_CLASSES for name in wm_class)


# ── Overlay ──────────────────────────────────────────────────────────────────
# Extension Shape de X11 : la zone qui reçoit les clics (Input) et la forme visible
# (Bounding) d'une fenêtre se décrivent par des rectangles.

def _window(win):
    return _x().create_resource_object('window', int(win.wm_frame(), 16))


def _shape(win, kind, rectangles):
    with _display_lock:
        _window(win).shape_rectangles(shape.SO.Set, kind, Xlib.X.Unsorted, 0, 0, rectangles)
        _x().sync()


def _rounded_rows(w, h, radius):
    """Rectangles, une ligne par pixel dans les coins, qui dessinent un rectangle arrondi."""
    r = min(radius, w // 2, h // 2)
    rows = []
    for y in range(r):
        inset = r - int(math.sqrt(max(0.0, r * r - (r - y - 0.5) ** 2)))
        rows.append((inset, y, w - 2 * inset, 1))
        rows.append((inset, h - 1 - y, w - 2 * inset, 1))
    if h - 2 * r > 0:
        rows.append((0, r, w, h - 2 * r))
    return rows


def overlay_click_through(win):
    """Zone de clic vide : les clics atteignent l'application derrière."""
    _shape(win, shape.SK.Input, [])


def overlay_round_corners(win, radius):
    w, h = win.winfo_width(), win.winfo_height()
    if w <= 1 or h <= 1:
        return
    # CreateRoundRectRgn de Windows prend le diamètre des coins, Shape le rayon.
    _shape(win, shape.SK.Bounding, _rounded_rows(w, h, radius // 2))


def overlay_circle(win):
    w, h = win.winfo_width(), win.winfo_height()
    if w <= 1 or h <= 1:
        return
    _shape(win, shape.SK.Bounding, _rounded_rows(w, h, min(w, h) // 2))


# ── Micro ────────────────────────────────────────────────────────────────────

# Greffons ALSA listés comme entrées mais qui ne sont pas des micros, ou qui doublonnent
# « default », « pipewire » et « pulse », lesquels passent par le son du système.
_ALSA_PLUGINS = {'sysdefault', 'front', 'iec958', 'spdif', 'dsnoop', 'dmix', 'samplerate',
                 'speexrate', 'lavrate', 'upmix', 'vdownmix', 'jack', 'oss', 'hdmi'}


def mic_hidden(name):
    """Entrée audio à cacher de la liste des micros (le filtre commun est dans app.py)."""
    lower = name.lower()
    return lower in _ALSA_PLUGINS or lower.startswith('surround')


# ── Session ──────────────────────────────────────────────────────────────────

def environment_problem():
    """Raison pour laquelle le raccourci ou le collage ne peuvent pas marcher, sinon None."""
    if os.environ.get('XDG_SESSION_TYPE') == 'wayland':
        return ("Cette session est en Wayland : le raccourci global et le collage automatique "
                "n'y fonctionnent pas.\n\nÀ l'écran de connexion, choisissez une session X11 "
                "(sous Linux Mint, Cinnamon tout court plutôt que Cinnamon sur Wayland).")
    if shutil.which('xclip') is None:
        return ("xclip est absent : la copie et le collage ne fonctionneront pas.\n\n"
                "Installez-le avec « sudo apt install xclip », ou relancez setup.sh.")
    return None


HOTKEY_FAILURE_HINT = "Vérifiez que la session est en X11 et non en Wayland."


# ── Presse-papier ────────────────────────────────────────────────────────────
# Sous X11, le presse-papier n'est pas stocké : l'application qui a copié le sert à la demande,
# dans chacun des formats (« cibles ») qu'elle propose. xclip ne sait servir qu'une cible à la
# fois : on garde la plus utile, image d'abord, puis fichiers, puis texte.

# Le texte passe aussi par xclip, plutôt que par ce que pyperclip devinerait seul (xsel,
# wl-clipboard, ou clip.exe s'il se croit sous WSL).
pyperclip.set_clipboard('xclip')

_RESTORE_TARGETS = (
    'image/png', 'x-special/gnome-copied-files', 'text/uri-list',
    'UTF8_STRING', 'text/plain;charset=utf-8', 'STRING', 'text/plain',
)
_TEXT_TARGETS = ('UTF8_STRING', 'text/plain;charset=utf-8', 'STRING', 'text/plain')


def _xclip_read(target):
    result = subprocess.run(['xclip', '-selection', 'clipboard', '-o', '-t', target],
                            capture_output=True, timeout=2)
    return result.stdout if result.returncode == 0 else None


def clipboard_snapshot():
    """Contenu du presse-papier : (cible, octets), ou None s'il est vide."""
    targets = _xclip_read('TARGETS')
    if not targets:
        return None
    offered = set(targets.decode('utf-8', 'replace').split())
    for target in _RESTORE_TARGETS:
        if target in offered:
            data = _xclip_read(target)
            if data is not None:
                return target, data
    return None


def clipboard_restore(snapshot):
    """Remet un contenu pris par clipboard_snapshot()."""
    if snapshot is None:
        return
    target, data = snapshot
    # xclip reste en arrière-plan pour servir le contenu jusqu'à la prochaine copie.
    subprocess.run(['xclip', '-selection', 'clipboard', '-t', target, '-i'],
                   input=data, check=True, timeout=2)


def clipboard_sequence():
    """Empreinte du texte du presse-papier : elle change quand autre chose y est copié.

    X11 n'a pas de compteur de copies comme Windows ; le texte suffit à savoir si le
    presse-papier tient encore la dictée qu'on vient de coller.
    """
    targets = _xclip_read('TARGETS')
    offered = set(targets.decode('utf-8', 'replace').split()) if targets else set()
    for target in _TEXT_TARGETS:
        if target in offered:
            return hash(_xclip_read(target))
    return hash(targets)
