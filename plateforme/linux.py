"""Fonctions propres à Linux (session X11) : clavier, presse-papier, fenêtre active, overlay.

Tout passe par le serveur X : raccourci global, collage simulé et forme de l'overlay. En
session Wayland, wayland.py prend le relais et reprend d'ici ce qui reste valable sous
XWayland (noms des touches, overlay, micros).
"""
import math
import os
import queue
import select
import shutil
import subprocess
import threading

import pyperclip
import Xlib.display
import Xlib.error
import Xlib.X
import Xlib.XK
from Xlib.ext import shape
from pynput import keyboard as pynput_keyboard
from pynput import mouse as pynput_mouse
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


# ── Raccourci réservé ────────────────────────────────────────────────────────
# L'écoute (XRecord) voit les touches sans les retenir : seule, elle laisserait le 9 du
# pavé numérique s'écrire dans la fenêtre active. Le raccourci est donc aussi réservé par
# XGrabKey (ou XGrabButton pour la souris) : X11 le remet à l'application seule, et
# l'écoute le voit toujours.

GRABS_HOTKEY = True
SUPPORTS_MOUSE = True
# Boutons 1 à 3 : clics ; 4 à 7 : molette. À partir de 8 : boutons latéraux et suivants.
MOUSE_BUTTON_MIN = 8
MOUSE_BUTTON_MAX = 31
BARE_KEY_WARNING = ("Cette touche seule est réservée à la dictée : elle n'aura plus son "
                    "effet habituel tant que ce raccourci est actif.")
BARE_BUTTON_WARNING = ("Ce bouton est réservé à la dictée : il n'aura plus son effet habituel "
                       "(page précédente ou suivante dans le navigateur, par exemple).")

_BUTTON_LABELS = {8: 'Bouton souris arrière', 9: 'Bouton souris avant'}
_GRAB_MODIFIERS = ('ctrl', 'alt', 'shift', 'win')


def button_label(button):
    return _BUTTON_LABELS.get(button, 'Bouton souris %d' % button)


def start_mouse_listener(on_click):
    """Écoute les clics de tout le système ; on_click(numéro du bouton, enfoncé)."""
    def click(x, y, button, pressed):
        number = getattr(button, 'value', None)
        if isinstance(number, int):
            on_click(number, pressed)

    listener = pynput_mouse.Listener(on_click=click)
    listener.start()
    return listener


class _Grabber(threading.Thread):
    """Connexion X dédiée aux réservations : tous ses appels passent par son propre fil."""

    def __init__(self):
        super().__init__(daemon=True)
        self.display = Xlib.display.Display()
        self.root = self.display.screen().root
        self.grabbed = []
        self.on_button = None
        self._commands = queue.Queue()
        self._wake_r, self._wake_w = os.pipe()
        self.start()

    def call(self, fn):
        reply = queue.Queue(maxsize=1)
        self._commands.put((fn, reply))
        os.write(self._wake_w, b'.')
        ok, value = reply.get(timeout=5)
        if not ok:
            raise value
        return value

    def run(self):
        while True:
            ready, _, _ = select.select([self.display, self._wake_r], [], [])
            if self._wake_r in ready:
                os.read(self._wake_r, 64)
            try:
                while True:
                    fn, reply = self._commands.get_nowait()
                    try:
                        reply.put((True, fn()))
                    except Exception as e:
                        reply.put((False, e))
            except queue.Empty:
                pass
            while self.display.pending_events():
                event = self.display.next_event()
                if event.type in (Xlib.X.ButtonPress, Xlib.X.ButtonRelease) and self.on_button:
                    try:
                        self.on_button(event.detail, event.type == Xlib.X.ButtonPress)
                    except Exception as e:
                        print(f"⚠️ Bouton du raccourci : {e}")

    def _modifier_masks(self):
        """Masque X de chaque modificateur, et celui de Verr. Num, d'après la disposition."""
        mapping = self.display.get_modifier_mapping()

        def mask_of(*names):
            codes = {code for name in names
                     for code, _ in self.display.keysym_to_keycodes(Xlib.XK.string_to_keysym(name))}
            for i, row in enumerate(mapping):
                if codes & (set(row) - {0}):
                    return 1 << i
            return 0

        masks = {'ctrl': Xlib.X.ControlMask, 'shift': Xlib.X.ShiftMask,
                 'alt': mask_of('Alt_L', 'Alt_R') or Xlib.X.Mod1Mask,
                 'win': mask_of('Super_L', 'Super_R') or Xlib.X.Mod4Mask}
        return masks, mask_of('Num_Lock')

    def set(self, hotkey, on_button):
        for kind, code, mask in self.grabbed:
            if kind == 'key':
                self.root.ungrab_key(code, mask)
            else:
                self.root.ungrab_button(code, mask)
        self.grabbed = []
        self.on_button = on_button

        masks, numlock = self._modifier_masks()
        # Verr. Maj et Verr. Num comptent comme des modificateurs pour X : chaque
        # combinaison est réservée avec et sans eux.
        locks = {0, Xlib.X.LockMask, numlock, Xlib.X.LockMask | numlock}
        required = 0
        for name in hotkey['modifiers']:
            required |= masks[name]
        # Comme l'écoute, un raccourci avec modificateurs accepte des modificateurs en plus
        # (certaines macros ajoutent Maj) ; une touche seule n'en accepte aucun.
        extra = [masks[m] for m in _GRAB_MODIFIERS if m not in hotkey['modifiers']] \
            if hotkey['modifiers'] else []
        combos = set()
        for subset in range(1 << len(extra)):
            mask = required
            for i, m in enumerate(extra):
                if subset >> i & 1:
                    mask |= m
            combos.update(mask | lock for lock in locks)

        if hotkey.get('button'):
            targets = [('button', hotkey['button'])]
        else:
            targets = [('key', code) for code in hotkey['vks']]
        catch = Xlib.error.CatchError(Xlib.error.BadAccess)
        for kind, code in targets:
            for mask in combos:
                if kind == 'key':
                    self.root.grab_key(code, mask, False, Xlib.X.GrabModeAsync,
                                       Xlib.X.GrabModeAsync, onerror=catch)
                else:
                    self.root.grab_button(code, mask, False,
                                          Xlib.X.ButtonPressMask | Xlib.X.ButtonReleaseMask,
                                          Xlib.X.GrabModeAsync, Xlib.X.GrabModeAsync,
                                          Xlib.X.NONE, Xlib.X.NONE, onerror=catch)
                self.grabbed.append((kind, code, mask))
        self.display.sync()
        if catch.get_error():
            return ("Ce raccourci est déjà réservé par une autre application ou par un "
                    "raccourci du système. Il lance quand même la dictée, mais la touche "
                    "parvient aussi à la fenêtre active.")
        return None


_grabber = None


def grab_hotkey(hotkey, on_button=None):
    """Réserve le raccourci à l'application ; on_button(bouton, enfoncé) reçoit les clics du
    bouton réservé. Renvoie None, ou la raison pour laquelle la réservation a échoué."""
    global _grabber
    try:
        if _grabber is None:
            _grabber = _Grabber()
        return _grabber.call(lambda: _grabber.set(hotkey, on_button))
    except Exception as e:
        return f"Réservation du raccourci impossible : {e}"


_controller = None


def send_paste(terminal):
    """Envoie Ctrl+V, ou Ctrl+Maj+V pour un terminal."""
    global _controller
    if _controller is None:
        _controller = pynput_keyboard.Controller()
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
