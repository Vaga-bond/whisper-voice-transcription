"""Fonctions propres à Linux en session Wayland : mêmes fonctions que linux.py, autres moyens.

Wayland interdit aux applications de lire le clavier des autres et de leur envoyer des
touches. Chaque besoin passe donc par un moyen autorisé :

- raccourci, Échap et boutons de la souris : processus entrees_wayland.py, qui lit les
  périphériques du noyau (groupe « input ») et réserve le raccourci ;
- collage : un clavier virtuel (uinput) envoie Ctrl+V ;
- presse-papier : wl-clipboard ;
- fenêtre active : un script KWin la signale à chaque changement (KDE Plasma).

L'interface Tk tourne sous XWayland : noms des touches et forme de l'overlay viennent de
linux.py, inchangés. Les codes de touches restent ceux de X11 (code evdev + 8), si bien qu'un
raccourci enregistré dans une session X11 garde son sens ici.
"""
import atexit
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time

import evdev
import pyperclip
from evdev import ecodes
from pynput import keyboard as pynput_keyboard

from .linux import (  # noqa: F401  (réexportés pour app.py)
    BARE_BUTTON_WARNING, BARE_KEY_WARNING, DEFAULT_HOTKEY, MODIFIER_LABELS, MOUSE_BUTTON_MAX,
    MOUSE_BUTTON_MIN, SAFE_BARE_CODES, _TERMINAL_CLASSES, _display_lock, _x, button_label,
    key_label, mic_hidden, overlay_circle, overlay_click_through, overlay_round_corners,
)

NAME = "Linux (Wayland)"
SETUP_SCRIPT = "setup.sh"

GRABS_HOTKEY = True
SUPPORTS_MOUSE = True

HOTKEY_FAILURE_HINT = ("Relancez setup.sh, puis fermez la session et rouvrez-la pour que "
                       "l'accès au clavier soit pris en compte.")

_X_OFFSET = 8
# Modificateurs et Échap, par keycode X, sous la forme que pynput leur donne sous X11.
_SPECIAL_KEYS = {
    37: pynput_keyboard.Key.ctrl_l, 105: pynput_keyboard.Key.ctrl_r,
    64: pynput_keyboard.Key.alt_l, 108: pynput_keyboard.Key.alt_gr,
    50: pynput_keyboard.Key.shift_l, 62: pynput_keyboard.Key.shift_r,
    133: pynput_keyboard.Key.cmd_l, 134: pynput_keyboard.Key.cmd_r,
    9: pynput_keyboard.Key.esc,
}
_MODIFIER_CODES = {
    'ctrl': {37, 105}, 'alt': {64, 108}, 'shift': {50, 62}, 'win': {133, 134},
}
_ALL_MODIFIER_CODES = set().union(*_MODIFIER_CODES.values())

_HELPER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'entrees_wayland.py')
_HELPER_START_TIMEOUT = 5.0


# ── Clavier et souris ────────────────────────────────────────────────────────

class _Entrees:
    """Le processus d'écoute, et ce que l'application y a branché."""

    def __init__(self):
        self.process = subprocess.Popen(
            [sys.executable, '-u', _HELPER], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            text=True, bufsize=1)
        self.held = set()                 # keycodes X enfoncés
        self.key_handlers = []            # (on_press, on_release)
        self.mouse_handlers = []          # on_click(bouton, enfoncé)
        self.reserved = None              # raccourci réservé
        self.on_button = None
        self.button_down = False
        self.denied = False
        self._ready = threading.Event()
        threading.Thread(target=self._read, daemon=True).start()
        if not self._ready.wait(_HELPER_START_TIMEOUT):
            raise RuntimeError("le processus d'écoute du clavier ne répond pas")
        if self.denied:
            raise RuntimeError("accès au clavier refusé (groupe « input »)")

    def _read(self):
        for line in self.process.stdout:
            kind, *values = line.split()
            try:
                if kind == 'k':
                    self._on_key(int(values[0]), int(values[1]))
                elif kind == 'b':
                    self._on_button(int(values[0]), values[1] == '1')
                elif kind == 'ok':
                    self._ready.set()
                elif kind == 'e':
                    if values[:1] == ['permission']:
                        self.denied = True
                    else:
                        print("⚠️ Écoute du clavier :", ' '.join(values))
            except Exception as e:
                print(f"⚠️ Écoute du clavier : {e}")
        self._ready.set()

    def _on_key(self, code, value):
        if value == 0:
            self.held.discard(code)
        else:
            self.held.add(code)
        key = _SPECIAL_KEYS.get(code) or pynput_keyboard.KeyCode.from_vk(code)
        for on_press, on_release in list(self.key_handlers):
            (on_release if value == 0 else on_press)(key, code)

    def _modifiers_match(self, required):
        held = {name for name, codes in _MODIFIER_CODES.items() if self.held & codes}
        return set(required) <= held if required else not held

    def _on_button(self, button, pressed):
        for on_click in list(self.mouse_handlers):
            on_click(button, pressed)
        hotkey = self.reserved
        if self.on_button is None or not hotkey or hotkey.get('button') != button:
            return
        # Comme le bouton réservé sous X11 : l'appui compte avec les bons modificateurs,
        # et son relâchement suit toujours.
        if pressed and self._modifiers_match(hotkey['modifiers']):
            self.button_down = True
            self.on_button(button, True)
        elif not pressed and self.button_down:
            self.button_down = False
            self.on_button(button, False)

    def reserve(self, hotkey, on_button):
        self.reserved, self.on_button = hotkey, on_button
        if hotkey.get('button'):
            wanted = {'modifiers': hotkey['modifiers'], 'button': hotkey['button']}
        elif hotkey.get('vks'):
            wanted = {'modifiers': hotkey['modifiers'], 'key': hotkey['vks'][0]}
        else:
            wanted = None
        self.process.stdin.write(json.dumps(wanted) + '\n')
        self.process.stdin.flush()

    def alive(self):
        return self.process.poll() is None


_entrees = None
_entrees_lock = threading.Lock()


def _get_entrees():
    global _entrees
    with _entrees_lock:
        if _entrees is None or not _entrees.alive():
            _entrees = _Entrees()
        return _entrees


class _Handle:
    """Ce que start_key_listener et start_mouse_listener renvoient : de quoi arrêter."""

    def __init__(self, items, item):
        self._items, self._item = items, item
        items.append(item)

    def stop(self):
        if self._item in self._items:
            self._items.remove(self._item)


def start_key_listener(on_press, on_release):
    """Écoute le clavier de tout le système ; on_press(touche, code) et on_release(touche, code)."""
    entrees = _get_entrees()
    _paste_keyboard()       # créé d'avance : le bureau met un instant à adopter un clavier
    _watch_active_window()
    return _Handle(entrees.key_handlers, (on_press, on_release))


def start_mouse_listener(on_click):
    """Écoute les boutons latéraux de toutes les souris ; on_click(bouton, enfoncé)."""
    return _Handle(_get_entrees().mouse_handlers, on_click)


def keys_held(codes):
    """Un modificateur ou l'une des touches `codes` est-il enfoncé en ce moment ?"""
    held = _get_entrees().held
    return bool(held & (_ALL_MODIFIER_CODES | set(codes)))


def grab_hotkey(hotkey, on_button=None):
    """Réserve le raccourci à l'application ; on_button(bouton, enfoncé) reçoit les clics du
    bouton réservé. Renvoie None, ou la raison pour laquelle la réservation a échoué."""
    try:
        _get_entrees().reserve(hotkey, on_button)
    except Exception as e:
        return f"Réservation du raccourci impossible : {e}"
    return None


# ── Collage ──────────────────────────────────────────────────────────────────

_keyboard = None
_PASTE_STEP = 0.008   # le temps que le bureau traite chaque touche, dans l'ordre


def _paste_keyboard():
    global _keyboard
    if _keyboard is None:
        _keyboard = evdev.UInput(
            {ecodes.EV_KEY: [ecodes.KEY_LEFTCTRL, ecodes.KEY_LEFTSHIFT, ecodes.KEY_V,
                             _v_code()]},
            name='Transcription Vocale (collage)')
    return _keyboard


def _v_code():
    """Code de la touche qui donne « v » dans la disposition du clavier (AZERTY, Dvorak…)."""
    try:
        with _display_lock:
            import Xlib.XK
            code = _x().keysym_to_keycode(Xlib.XK.string_to_keysym('v'))
        if code:
            return code - _X_OFFSET
    except Exception:
        pass
    return ecodes.KEY_V


def send_paste(terminal):
    """Envoie Ctrl+V, ou Ctrl+Maj+V pour un terminal."""
    keyboard = _paste_keyboard()
    keys = [ecodes.KEY_LEFTCTRL] + ([ecodes.KEY_LEFTSHIFT] if terminal else [])
    v = _v_code()

    def tap(code, value):
        keyboard.write(ecodes.EV_KEY, code, value)
        keyboard.syn()
        time.sleep(_PASTE_STEP)

    try:
        for key in keys:
            tap(key, 1)
        tap(v, 1)
        tap(v, 0)
    finally:
        for key in reversed(keys):
            tap(key, 0)


# ── Fenêtre active ───────────────────────────────────────────────────────────
# KWin ne dit pas quelle fenêtre est active à qui le demande, mais un script KWin peut
# l'annoncer : on lui fait appeler une méthode D-Bus de l'application à chaque changement.

_WAYLAND_TERMINAL_CLASSES = _TERMINAL_CLASSES | {'org.kde.konsole', 'org.kde.yakuake', 'yakuake'}
_KWIN_SCRIPT_NAME = 'transcription-vocale-fenetre-active'
_DBUS_INTERFACE = 'org.transcriptionvocale.FenetreActive'
_DBUS_XML = ('<node><interface name="%s"><method name="Changer">'
             '<arg type="s" name="classe" direction="in"/></method></interface></node>'
             % _DBUS_INTERFACE)

_active_class = None
_watch_started = False


def _watch_active_window():
    global _watch_started
    if _watch_started:
        return
    _watch_started = True
    try:
        _start_window_watch()
    except Exception as e:
        print(f"⚠️ Fenêtre active non suivie (KDE Plasma requis) : {e}")


def _start_window_watch():
    import gi
    gi.require_version('Gio', '2.0')
    from gi.repository import Gio, GLib

    ready = threading.Event()
    failure = []

    def on_call(_conn, _sender, _path, _iface, _method, params, invocation):
        global _active_class
        _active_class = params.unpack()[0].lower()
        invocation.return_value(None)

    def serve():
        context = GLib.MainContext.new()
        context.push_thread_default()
        try:
            bus = Gio.bus_get_sync(Gio.BusType.SESSION)
            interface = Gio.DBusNodeInfo.new_for_xml(_DBUS_XML).interfaces[0]
            register = getattr(bus, 'register_object_with_closures2', None)
            if register is not None:
                register('/FenetreActive', interface, on_call, None, None)
            else:
                bus.register_object('/FenetreActive', interface, on_call, None, None)
            _load_kwin_script(Gio, GLib, bus)
        except Exception as e:
            failure.append(e)
            ready.set()
            return
        ready.set()
        GLib.MainLoop(context).run()

    threading.Thread(target=serve, daemon=True).start()
    ready.wait(5)
    if failure:
        raise failure[0]


def _load_kwin_script(Gio, GLib, bus):
    script = (
        'function signaler(fenetre) {\n'
        '    callDBus("%s", "/FenetreActive", "%s", "Changer",\n'
        '             fenetre ? String(fenetre.resourceClass) : "");\n'
        '}\n'
        'workspace.windowActivated.connect(signaler);\n'
        'signaler(workspace.activeWindow);\n'
    ) % (bus.get_unique_name(), _DBUS_INTERFACE)
    folder = os.environ.get('XDG_RUNTIME_DIR') or tempfile.gettempdir()
    path = os.path.join(folder, _KWIN_SCRIPT_NAME + '.js')
    with open(path, 'w', encoding='utf-8') as f:
        f.write(script)

    scripting = Gio.DBusProxy.new_sync(bus, Gio.DBusProxyFlags.NONE, None, 'org.kde.KWin',
                                       '/Scripting', 'org.kde.kwin.Scripting', None)

    def unload():
        scripting.call_sync('unloadScript', GLib.Variant('(s)', (_KWIN_SCRIPT_NAME,)),
                            Gio.DBusCallFlags.NONE, -1, None)

    unload()   # celui d'une session précédente de l'application, s'il était resté
    scripting.call_sync('loadScript', GLib.Variant('(ss)', (path, _KWIN_SCRIPT_NAME)),
                        Gio.DBusCallFlags.NONE, -1, None)
    scripting.call_sync('start', None, Gio.DBusCallFlags.NONE, -1, None)
    atexit.register(lambda: _quietly(unload))


def _quietly(fn):
    try:
        fn()
    except Exception:
        pass


def terminal_focused():
    """La fenêtre au premier plan est-elle un terminal connu ?"""
    _watch_active_window()
    return _active_class in _WAYLAND_TERMINAL_CLASSES


# ── Session ──────────────────────────────────────────────────────────────────

def environment_problem():
    """Raison pour laquelle le raccourci ou le collage ne peuvent pas marcher, sinon None."""
    devices = glob.glob('/dev/input/event*')
    if devices and not any(os.access(path, os.R_OK) for path in devices):
        return ("L'accès au clavier est refusé : le raccourci global ne fonctionnera pas.\n\n"
                "Relancez setup.sh, puis fermez la session et rouvrez-la.")
    if not os.access('/dev/uinput', os.W_OK):
        return ("Le clavier virtuel (/dev/uinput) est inaccessible : le collage automatique "
                "ne fonctionnera pas.\n\nRelancez setup.sh, puis redémarrez l'ordinateur.")
    if shutil.which('wl-copy') is None or shutil.which('wl-paste') is None:
        return ("wl-clipboard est absent : la copie et le collage ne fonctionneront pas.\n\n"
                "Installez-le avec « sudo apt install wl-clipboard », ou relancez setup.sh.")
    return None


# ── Presse-papier ────────────────────────────────────────────────────────────
# wl-clipboard parle au presse-papier de Wayland sans avoir de fenêtre active. Comme xclip
# sous X11, wl-copy reste en arrière-plan pour servir le contenu jusqu'à la prochaine copie.

pyperclip.set_clipboard('wl-clipboard')

_RESTORE_TYPES = (
    'image/png', 'x-special/gnome-copied-files', 'text/uri-list',
    'text/plain;charset=utf-8', 'UTF8_STRING', 'text/plain', 'STRING', 'TEXT',
)
_TEXT_TYPES = ('text/plain;charset=utf-8', 'UTF8_STRING', 'text/plain', 'STRING', 'TEXT')


def _offered_types():
    result = subprocess.run(['wl-paste', '--list-types'], capture_output=True, timeout=2)
    if result.returncode != 0:
        return []
    return result.stdout.decode('utf-8', 'replace').split()


def _wl_read(mime):
    result = subprocess.run(['wl-paste', '--no-newline', '--type', mime],
                            capture_output=True, timeout=2)
    return result.stdout if result.returncode == 0 else None


def clipboard_snapshot():
    """Contenu du presse-papier : (type, octets), ou None s'il est vide."""
    offered = set(_offered_types())
    for mime in _RESTORE_TYPES:
        if mime in offered:
            data = _wl_read(mime)
            if data is not None:
                return mime, data
    return None


def clipboard_restore(snapshot):
    """Remet un contenu pris par clipboard_snapshot()."""
    if snapshot is None:
        return
    mime, data = snapshot
    # Sorties détachées : wl-copy reste vivant pour servir le contenu, et run() attendrait
    # sinon la fermeture de ses sorties.
    subprocess.run(['wl-copy', '--type', mime], input=data, check=True, timeout=2,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def clipboard_sequence():
    """Empreinte du texte du presse-papier : elle change quand autre chose y est copié."""
    offered = _offered_types()
    for mime in _TEXT_TYPES:
        if mime in offered:
            return hash(_wl_read(mime))
    return hash(tuple(offered))
