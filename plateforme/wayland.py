"""Fonctions propres à Linux en session Wayland (KDE Plasma) : mêmes besoins que linux.py,
par les moyens que Wayland autorise.

Wayland n'envoie le clavier et la souris qu'à la fenêtre au premier plan. Une application
en arrière-plan comme celle-ci demande donc au bureau :

- raccourci de dictée : portail de raccourcis globaux (org.freedesktop.portal.GlobalShortcuts).
  KDE intercepte la combinaison, l'empêche d'atteindre l'application active et signale
  l'appui puis le relâchement. Le portail ne connaît que le clavier : pour un bouton de la
  souris, KWin transforme le bouton en une touche qu'aucun clavier ne porte (« Lancer 5 »),
  et cette touche est le raccourci du portail. La transformation n'existe que tant que la
  dictée est active : en pause, le bouton retrouve son effet habituel. Aucun périphérique
  n'est lu par l'application ;
- annulation : pendant un enregistrement seulement, le bouton arrière devient de même une
  seconde touche fantôme (« Lancer 6 »), raccourci « annuler » du portail ;
- collage : un clavier virtuel (uinput) envoie Ctrl+V, en écriture seule ;
- presse-papier : wl-clipboard ;
- fenêtre active : un script KWin la signale à chaque changement.

L'interface Tk tourne sous XWayland : noms des touches et forme de l'overlay viennent de
linux.py, inchangés.
"""
import atexit
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
import uuid

import evdev
import pyperclip
from evdev import ecodes

from .linux import (  # noqa: F401  (réexportés pour app.py)
    DEFAULT_PAUSE_HOTKEY, MODIFIER_LABELS, MOUSE_BUTTON_MAX, MOUSE_BUTTON_MIN,
    SAFE_BARE_CODES, _TERMINAL_CLASSES, _display_lock, _x, button_label, key_label,
    mic_hidden, overlay_circle, overlay_click_through, overlay_round_corners,
)

NAME = "Linux (Wayland)"
SETUP_SCRIPT = "setup.sh"

# Le raccourci passe par le bureau : l'application ne lit ni ne réserve aucune touche.
HOTKEY_VIA_DESKTOP = True
GRABS_HOTKEY = False
SUPPORTS_MOUSE = True

DEFAULT_HOTKEY = {'modifiers': [], 'vks': [], 'button': 9}   # bouton avant de la souris
BARE_KEY_WARNING = ""
BARE_BUTTON_WARNING = ("Tant que la dictée est active, ce bouton ne sert qu'à elle. En pause, "
                       "il retrouve son effet habituel (page suivante, jeux…). Pendant un "
                       "enregistrement, le bouton arrière l'annule.")

HOTKEY_FAILURE_HINT = ("Le raccourci passe par le portail de raccourcis de KDE "
                       "(xdg-desktop-portal-kde) : vérifiez qu'il est installé, puis relancez "
                       "l'application.")

APP_ID = 'transcription-vocale'     # nom du fichier .desktop installé par setup.sh


# ── Raccourci par le portail ─────────────────────────────────────────────────

_PORTAL = 'org.freedesktop.portal.Desktop'
_PORTAL_PATH = '/org/freedesktop/portal/desktop'
_SHORTCUTS = 'org.freedesktop.portal.GlobalShortcuts'
_REQUEST = 'org.freedesktop.portal.Request'

SHORTCUT_MOUSE = 'dictee-souris'
SHORTCUT_KEYBOARD = 'dictee-clavier'
SHORTCUT_CANCEL = 'dictee-annuler'
_DICTATION_SHORTCUTS = {SHORTCUT_MOUSE, SHORTCUT_KEYBOARD}
# Touche fantôme du bouton de souris. Elle doit exister dans la disposition du clavier (F24
# n'y est pas : KWin ignorerait la transformation) et n'être prise par rien d'autre
# (« Outils » ouvre la Configuration du système).
PHANTOM_TRIGGER = 'XF86Launch5'
PHANTOM_KEY = 'Launch (5)'           # même touche, sous le nom qu'en donne Qt à KWin
CANCEL_TRIGGER = 'XF86Launch6'       # touche fantôme du bouton d'annulation
CANCEL_KEY = 'Launch (6)'
CANCEL_BUTTON = 8                    # bouton arrière

_REQUEST_TIMEOUT = 120               # la première fois, KDE attend que l'on accepte


class _Portail:
    """Une session du portail, servie par un fil qui fait tourner sa boucle GLib."""

    def __init__(self, shortcuts, on_change, on_ready):
        """Ne bloque pas : la première fois, KDE attend que l'on accepte les raccourcis.
        on_ready(None ou exception) est appelé, hors du fil de Tk, une fois la session prête."""
        import gi
        gi.require_version('Gio', '2.0')
        from gi.repository import Gio, GLib
        self.Gio, self.GLib = Gio, GLib
        self.shortcuts = shortcuts
        self.on_change = on_change
        self.on_ready = on_ready
        self.held = set()                 # raccourcis enfoncés en ce moment
        self.session = None
        self.context = None
        self.conn = None
        self.ready = False
        threading.Thread(target=self._serve, daemon=True).start()

    # Fil du portail

    def _serve(self):
        Gio, GLib = self.Gio, self.GLib
        self.context = GLib.MainContext.new()
        self.context.push_thread_default()
        try:
            # Connexion à part : l'identité de l'application se déclare une seule fois par
            # connexion, avant tout autre appel au portail.
            address = Gio.dbus_address_get_for_bus_sync(Gio.BusType.SESSION, None)
            self.conn = Gio.DBusConnection.new_for_address_sync(
                address, Gio.DBusConnectionFlags.AUTHENTICATION_CLIENT
                | Gio.DBusConnectionFlags.MESSAGE_BUS_CONNECTION, None, None)
            self._register()
            self.session = self._request('CreateSession', lambda token: GLib.Variant(
                '(a{sv})', ({'handle_token': GLib.Variant('s', token),
                             'session_handle_token': GLib.Variant('s', 's' + token)},)
            ))['session_handle']
            for signal in ('Activated', 'Deactivated'):
                self.conn.signal_subscribe(_PORTAL, _SHORTCUTS, signal, _PORTAL_PATH, None,
                                           Gio.DBusSignalFlags.NONE, self._on_signal)
            self._bind()
        except Exception as e:
            self.on_ready(e)
            return
        self.ready = True
        self.on_ready(None)
        GLib.MainLoop(self.context).run()

    def _register(self):
        try:
            self.conn.call_sync(_PORTAL, _PORTAL_PATH, 'org.freedesktop.host.portal.Registry',
                                'Register', self.GLib.Variant('(sa{sv})', (APP_ID, {})),
                                None, self.Gio.DBusCallFlags.NONE, -1, None)
        except self.GLib.Error as e:
            # Sans fichier .desktop, KDE range les raccourcis sous un nom générique.
            print(f"⚠️ Portail : application non identifiée ({e.message})")

    def _request(self, method, build_params):
        """Appelle une méthode qui répond par un objet Request et attend sa réponse."""
        Gio, GLib = self.Gio, self.GLib
        token = 'tv' + uuid.uuid4().hex
        sender = self.conn.get_unique_name()[1:].replace('.', '_')
        path = '/org/freedesktop/portal/desktop/request/%s/%s' % (sender, token)
        loop = GLib.MainLoop(self.context)
        result = {}

        def on_response(_conn, _sender, _path, _iface, _signal, params):
            result['code'], result['data'] = params.unpack()
            loop.quit()

        def on_timeout():
            loop.quit()
            return False

        subscription = self.conn.signal_subscribe(_PORTAL, _REQUEST, 'Response', path, None,
                                                  Gio.DBusSignalFlags.NONE, on_response)
        try:
            self.conn.call_sync(_PORTAL, _PORTAL_PATH, _SHORTCUTS, method, build_params(token),
                                GLib.VariantType('(o)'), Gio.DBusCallFlags.NONE, -1, None)
            timeout = GLib.timeout_source_new_seconds(_REQUEST_TIMEOUT)
            timeout.set_callback(on_timeout)
            timeout.attach(self.context)
            loop.run()
            timeout.destroy()
        finally:
            self.conn.signal_unsubscribe(subscription)
        if 'code' not in result:
            raise RuntimeError("%s : pas de réponse du portail" % method)
        if result['code'] != 0:
            raise RuntimeError("%s refusé par le portail (code %d)" % (method, result['code']))
        return result['data']

    def _bind(self):
        GLib = self.GLib
        shortcuts = []
        for ident, description, trigger in self.shortcuts:
            options = {'description': GLib.Variant('s', description)}
            if trigger:
                options['preferred_trigger'] = GLib.Variant('s', trigger)
            shortcuts.append((ident, options))
        self._request('BindShortcuts', lambda token: GLib.Variant(
            '(oa(sa{sv})sa{sv})',
            (self.session, shortcuts, '', {'handle_token': GLib.Variant('s', token)})))

    def _on_signal(self, _conn, _sender, _path, _iface, signal, params):
        session, ident, _timestamp, _options = params.unpack()
        if session != self.session:
            return
        if signal == 'Activated':
            self.held.add(ident)
            self.on_change(ident, True)
        else:
            self.held.discard(ident)
            self.on_change(ident, False)

    # Appels depuis les autres fils

    def _in_thread(self, fn):
        """Exécute fn dans le fil du portail et en rend le résultat (ou l'exception)."""
        done = threading.Event()
        outcome = {}

        def run():
            try:
                outcome['value'] = fn()
            except Exception as e:
                outcome['error'] = e
            done.set()
            return False

        self.context.invoke_full(self.GLib.PRIORITY_DEFAULT, run)
        if not done.wait(_REQUEST_TIMEOUT + 10):
            raise RuntimeError("le portail de raccourcis ne répond pas")
        if 'error' in outcome:
            raise outcome['error']
        return outcome.get('value')

    def triggers(self):
        """{identifiant: description du déclencheur choisi dans KDE}."""
        def call():
            data = self._request('ListShortcuts', lambda token: self.GLib.Variant(
                '(oa{sv})', (self.session, {'handle_token': self.GLib.Variant('s', token)})))
            return {ident: options.get('trigger_description', '')
                    for ident, options in data.get('shortcuts', [])}
        return self._in_thread(call)

    def configure(self):
        """Ouvre la fenêtre de KDE où l'on choisit les touches des raccourcis."""
        def call():
            self.conn.call_sync(_PORTAL, _PORTAL_PATH, _SHORTCUTS, 'ConfigureShortcuts',
                                self.GLib.Variant('(osa{sv})', (self.session, '', {})),
                                None, self.Gio.DBusCallFlags.NONE, -1, None)
        self._in_thread(call)


_portail = None


def start_desktop_hotkey(on_change, on_ready):
    """Demande au bureau les raccourcis de dictée, sans bloquer.

    on_change(action, enfoncé) à chaque appui et relâchement, action valant 'dictee' ou
    'annuler' ; on_ready(None ou raison de l'échec)
    une fois la session prête, après acceptation dans la fenêtre de KDE la première fois.
    Les deux sont appelés hors du fil de Tk."""
    global _portail
    if _portail is not None:
        return
    try:
        _clear_phantom_rebinds()      # restes d'une session précédente interrompue
    except (OSError, subprocess.SubprocessError) as e:
        print(f"⚠️ Réglages des boutons de KDE illisibles : {e}")
    shortcuts = [
        (SHORTCUT_MOUSE, "Dictée par un bouton de la souris (touche fantôme, ne pas modifier)",
         PHANTOM_TRIGGER),
        (SHORTCUT_KEYBOARD, "Dictée par une combinaison du clavier", None),
        (SHORTCUT_CANCEL, "Annulation par le bouton arrière (touche fantôme, ne pas modifier)",
         CANCEL_TRIGGER),
    ]

    def changed(ident, pressed):
        on_change('annuler' if ident == SHORTCUT_CANCEL else 'dictee', pressed)

    def ready(error):
        on_ready(None if error is None else f"Raccourci de dictée indisponible : {error}")

    try:
        _portail = _Portail(shortcuts, changed, ready)
    except Exception as e:
        on_ready(f"Raccourci de dictée indisponible : {e}")


def configure_desktop_keyboard():
    """Ouvre la fenêtre de KDE où l'on choisit la combinaison du clavier. Renvoie None, ou la
    raison d'un échec."""
    if _portail is None or not _portail.ready:
        return "Le portail de raccourcis n'est pas prêt."
    try:
        _portail.configure()
    except Exception as e:
        return f"Fenêtre des raccourcis de KDE impossible à ouvrir : {e}"
    return None


def desktop_keyboard_trigger():
    """Combinaison du clavier choisie dans KDE pour la dictée, ou chaîne vide."""
    if _portail is None or not _portail.ready:
        return ''
    try:
        return _portail.triggers().get(SHORTCUT_KEYBOARD, '')
    except Exception as e:
        print(f"⚠️ Raccourcis du portail illisibles : {e}")
        return ''


def keys_held(codes):
    """Le raccourci de dictée est-il encore tenu (bouton ou combinaison) ?"""
    return bool(_portail is not None and _portail.held & _DICTATION_SHORTCUTS)


# ── Bouton de souris transformé par KWin ─────────────────────────────────────
# Réglage natif de KDE (Configuration du système > Souris > boutons supplémentaires) : le
# groupe [ButtonRebinds][Mouse] de kcminputrc associe ExtraButtonN à une touche. KWin le
# relit dès qu'on l'écrit avec --notify. ExtraButton1 est le bouton X 8 (arrière),
# ExtraButton2 le bouton 9 (avant), etc.

_PHANTOM_VALUE = 'Key,' + PHANTOM_KEY
_CANCEL_VALUE = 'Key,' + CANCEL_KEY
_OUR_VALUES = {_PHANTOM_VALUE, _CANCEL_VALUE}
_REBIND_GROUP = ['--file', 'kcminputrc', '--group', 'ButtonRebinds', '--group', 'Mouse']
_REBIND_HEADER = '[ButtonRebinds][Mouse]'


def _rebind_name(button):
    return 'ExtraButton%d' % (button - MOUSE_BUTTON_MIN + 1)


def _read_rebinds():
    """{ExtraButtonN: valeur} du groupe, lu d'un coup dans kcminputrc (un appel à kreadconfig6
    par bouton retarderait chaque début d'enregistrement)."""
    base = os.environ.get('XDG_CONFIG_HOME') or os.path.expanduser('~/.config')
    try:
        with open(os.path.join(base, 'kcminputrc'), encoding='utf-8') as f:
            lines = f.read().splitlines()
    except FileNotFoundError:
        return {}
    rebinds, inside = {}, False
    for line in lines:
        line = line.strip()
        if line.startswith('['):
            inside = line == _REBIND_HEADER
        elif inside and '=' in line:
            key, value = line.split('=', 1)
            rebinds[key.strip()] = value.strip()
    return rebinds


def _read_rebind(name):
    return _read_rebinds().get(name, '')


def _write_rebind(name, value):
    args = ['kwriteconfig6', *_REBIND_GROUP, '--key', name, '--notify']
    args += ['--delete'] if value is None else [value]
    subprocess.run(args, check=True, capture_output=True, timeout=5)


def _clear_phantom_rebinds(keep=()):
    """Retire nos transformations de bouton, sauf celles de `keep` ; ne touche à aucune autre."""
    for name, value in _read_rebinds().items():
        if name not in keep and value in _OUR_VALUES:
            _write_rebind(name, None)


def desktop_cancel_button(hotkey):
    """Bouton qui annule un enregistrement, ou None quand il sert déjà à la dictée."""
    return None if hotkey.get('button') == CANCEL_BUTTON else CANCEL_BUTTON


def apply_desktop_hotkey(hotkey, active, recording=False):
    """Pose ou retire les transformations de boutons selon l'état de la dictée : bouton de
    dictée tant qu'elle est active, bouton d'annulation pendant un enregistrement.

    Renvoie None, ou la raison pour laquelle le bouton de dictée ne peut pas servir."""
    wanted = {}
    if active and hotkey.get('button'):
        wanted[_rebind_name(hotkey['button'])] = _PHANTOM_VALUE
    cancel = desktop_cancel_button(hotkey)
    if active and recording and cancel:
        wanted[_rebind_name(cancel)] = _CANCEL_VALUE
    problem = None
    try:
        current = _read_rebinds()
        for name, value in list(wanted.items()):
            taken = current.get(name)
            if taken and taken not in _OUR_VALUES:
                del wanted[name]
                if value == _PHANTOM_VALUE:
                    problem = ("Ce bouton a déjà une autre fonction dans les réglages de la souris "
                               "de KDE (%s) : retirez-la, ou choisissez un autre bouton." % taken)
                else:
                    print("⚠️ Bouton arrière déjà réglé dans KDE (%s) : pas d'annulation par la "
                          "souris." % taken)
        for name, value in current.items():
            if value in _OUR_VALUES and wanted.get(name) != value:
                _write_rebind(name, None)
        for name, value in wanted.items():
            if current.get(name) != value:
                _write_rebind(name, value)
    except (OSError, subprocess.SubprocessError) as e:
        return f"Réglage du bouton dans KDE impossible : {e}"
    return problem


def _restore_buttons():
    try:
        _clear_phantom_rebinds()
    except Exception:
        pass


atexit.register(_restore_buttons)
# À la fermeture de session, le système arrête l'application par SIGTERM : sans ce relais,
# atexit ne passerait pas et le bouton resterait transformé jusqu'au prochain lancement.
if (threading.current_thread() is threading.main_thread()
        and signal.getsignal(signal.SIGTERM) is signal.SIG_DFL):
    signal.signal(signal.SIGTERM, lambda *_: sys.exit(0))


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


def prepare_paste():
    """Crée d'avance le clavier de collage et le suivi de la fenêtre active : le bureau met un
    instant à adopter un nouveau clavier."""
    try:
        _paste_keyboard()
    except OSError as e:
        print(f"⚠️ Clavier de collage indisponible : {e}")
    _watch_active_window()


def _v_code():
    """Code de la touche qui donne « v » dans la disposition du clavier (AZERTY, Dvorak…)."""
    try:
        with _display_lock:
            import Xlib.XK
            code = _x().keysym_to_keycode(Xlib.XK.string_to_keysym('v'))
        if code:
            return code - 8
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


def bring_to_front(title):
    """Met au premier plan la fenêtre de l'application, par un script KWin d'un instant :
    KDE refuse sinon qu'une application d'arrière-plan prenne le premier plan."""
    import json
    try:
        import gi
        gi.require_version('Gio', '2.0')
        from gi.repository import Gio, GLib
        script = (
            'for (const fenetre of workspace.windowList()) {\n'
            '    if (fenetre.pid === %d && fenetre.caption === %s) {\n'
            '        workspace.activeWindow = fenetre;\n'
            '        break;\n'
            '    }\n'
            '}\n'
        ) % (os.getpid(), json.dumps(title))
        name = _KWIN_SCRIPT_NAME + '-premier-plan'
        folder = os.environ.get('XDG_RUNTIME_DIR') or tempfile.gettempdir()
        path = os.path.join(folder, name + '.js')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(script)
        bus = Gio.bus_get_sync(Gio.BusType.SESSION)
        scripting = Gio.DBusProxy.new_sync(bus, Gio.DBusProxyFlags.NONE, None, 'org.kde.KWin',
                                           '/Scripting', 'org.kde.kwin.Scripting', None)
        unload = GLib.Variant('(s)', (name,))
        scripting.call_sync('unloadScript', unload, Gio.DBusCallFlags.NONE, -1, None)
        scripting.call_sync('loadScript', GLib.Variant('(ss)', (path, name)),
                            Gio.DBusCallFlags.NONE, -1, None)
        scripting.call_sync('start', None, Gio.DBusCallFlags.NONE, -1, None)
        time.sleep(1)     # KWin lit le fichier du script en différé : le laisser s'exécuter
        scripting.call_sync('unloadScript', unload, Gio.DBusCallFlags.NONE, -1, None)
    except Exception as e:
        print(f"⚠️ Fenêtre non mise au premier plan : {e}")


def terminal_focused():
    """La fenêtre au premier plan est-elle un terminal connu ?"""
    _watch_active_window()
    return _active_class in _WAYLAND_TERMINAL_CLASSES


# ── Session ──────────────────────────────────────────────────────────────────

def environment_problem():
    """Raison pour laquelle le raccourci ou le collage ne peuvent pas marcher, sinon None."""
    if shutil.which('kwriteconfig6') is None or shutil.which('kreadconfig6') is None:
        return ("Les outils de configuration de KDE (kwriteconfig6) sont absents : le bouton "
                "de souris ne pourra pas servir de raccourci.\n\nCette version Wayland "
                "demande KDE Plasma 6.")
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
