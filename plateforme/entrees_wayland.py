"""Processus d'écoute du clavier et de la souris sous Wayland, lancé par plateforme/wayland.py.

Wayland ne laisse aucune application lire les touches des autres : ce processus lit donc les
périphériques du noyau (/dev/input, accessibles au groupe « input »). Il rapporte chaque
touche et chaque bouton latéral sur sa sortie, une ligne par événement :

    k <keycode X> <0 relâchée | 1 enfoncée | 2 répétée>
    b <bouton X> <0 | 1>

Pour réserver le raccourci, comme XGrabKey sous X11, il prend le périphérique pour lui seul
(EVIOCGRAB) et en rejoue tous les autres événements sur un clone virtuel (uinput) qui porte le
même nom et les mêmes identifiants, pour que KDE lui applique les mêmes réglages. Il tourne à
part de l'application : une transcription qui occupe Python ne ralentit jamais la frappe, et
s'il s'arrête, le noyau rend aussitôt les périphériques.

Il lit sur son entrée la réservation voulue, une ligne JSON :

    {"modifiers": ["ctrl", "alt"], "key": <keycode X>}   ou   {"modifiers": [], "button": 8}
    null : rien de réservé

Il s'arrête quand son entrée se ferme, c'est-à-dire quand l'application se termine.
"""
import errno
import glob
import json
import os
import selectors
import signal
import struct
import sys
import time

import evdev
from evdev import ecodes

# Keycode X = code evdev + 8 : les raccourcis enregistrés sous X11 restent valables.
X_OFFSET = 8
# Boutons latéraux : BTN_SIDE est le bouton 8 de X11, BTN_EXTRA le 9, et ainsi de suite.
X_FIRST_SIDE_BUTTON = 8
SIDE_BUTTONS = 24

MODIFIERS = {
    'ctrl': {ecodes.KEY_LEFTCTRL, ecodes.KEY_RIGHTCTRL},
    'alt': {ecodes.KEY_LEFTALT, ecodes.KEY_RIGHTALT},
    'shift': {ecodes.KEY_LEFTSHIFT, ecodes.KEY_RIGHTSHIFT},
    'win': {ecodes.KEY_LEFTMETA, ecodes.KEY_RIGHTMETA},
}

# Voyants, sons et vibrations vont du bureau au périphérique : le noyau renvoie aussi à ses
# lecteurs ceux qu'on écrit sur le vrai clavier, et les rejouer sur le clone ferait tourner
# les voyants en boucle entre le bureau et lui.
OUTPUT_TYPES = {ecodes.EV_LED, ecodes.EV_SND, ecodes.EV_FF}

RESCAN_SECONDS = 2.0
_EVENT = struct.Struct('llHHi')   # struct input_event sur 64 bits


def _is_virtual(path):
    """Périphérique créé par un logiciel (nos clones, Steam, bureau à distance…) ?"""
    sys_path = '/sys/class/input/%s/device' % os.path.basename(path)
    return '/virtual/' in os.path.realpath(sys_path)


class Device:
    """Un périphérique physique, éventuellement réservé et rejoué sur son clone."""

    def __init__(self, path):
        self.path = path
        self.dev = evdev.InputDevice(path)
        self.keys = set(self.dev.capabilities().get(ecodes.EV_KEY, []))
        self.clone = None
        self.swallowed = set()

    def fileno(self):
        return self.dev.fd

    def busy(self):
        """Une touche est tenue : la prendre ou la rendre maintenant la laisserait coincée."""
        try:
            return bool(self.dev.active_keys())
        except OSError:
            return False

    def grab(self):
        info = self.dev.info
        clone = evdev.UInput.from_device(
            self.dev, name=self.dev.name, vendor=info.vendor, product=info.product,
            version=info.version, bustype=info.bustype)
        try:
            self.dev.grab()
        except OSError:
            clone.close()
            raise
        self.clone = clone

    def ungrab(self):
        try:
            self.dev.ungrab()
        except OSError:
            pass
        self.clone.close()
        self.clone = None
        self.swallowed.clear()


class Listener:
    def __init__(self):
        self.selector = selectors.DefaultSelector()
        self.devices = {}          # chemin -> Device
        self.held = set()          # touches enfoncées, tous claviers confondus
        self.reserve = None        # (modificateurs requis, code evdev visé)
        self.pending = False       # une prise ou un retour attend qu'on relâche les touches
        self.denied = False
        self.stdin_buffer = b''
        self.selector.register(sys.stdin.fileno(), selectors.EVENT_READ, ('stdin', None))

    def emit(self, line):
        try:
            sys.stdout.write(line + '\n')
            sys.stdout.flush()
        except BrokenPipeError:
            raise SystemExit(0)

    # ── Périphériques ───────────────────────────────────────────────────────

    def scan(self):
        present = set(glob.glob('/dev/input/event*'))
        for path in sorted(present - set(self.devices)):
            if _is_virtual(path):
                continue
            try:
                device = Device(path)
            except PermissionError:
                if not self.denied:
                    self.denied = True
                    self.emit('e permission')
                continue
            except OSError:
                continue
            if not device.keys:
                device.dev.close()
                continue
            self.devices[path] = device
            self.selector.register(device, selectors.EVENT_READ, ('device', device))
        for path in set(self.devices) - present:
            self.drop(self.devices[path])
        self.apply_reserve()

    def drop(self, device):
        self.selector.unregister(device)
        if device.clone is not None:
            self.release(device)
        del self.devices[device.path]
        try:
            device.dev.close()
        except OSError:
            pass

    def take(self, device):
        device.grab()
        self.selector.register(device.clone.fd, selectors.EVENT_READ, ('clone', device))

    def release(self, device):
        self.selector.unregister(device.clone.fd)
        device.ungrab()

    def apply_reserve(self):
        """Prend les périphériques qui portent la touche réservée, rend les autres."""
        self.pending = False
        target = self.reserve[1] if self.reserve else None
        for device in list(self.devices.values()):
            want = target is not None and target in device.keys
            if want == (device.clone is not None):
                continue
            if device.busy():
                self.pending = True
                continue
            try:
                self.take(device) if want else self.release(device)
            except OSError as e:
                self.emit('e grab %s' % e)

    # ── Événements ──────────────────────────────────────────────────────────

    def modifiers_match(self, required):
        held = {name for name, codes in MODIFIERS.items() if self.held & codes}
        return required <= held if required else not held

    def swallow(self, device, event):
        """L'événement est-il le raccourci réservé, à ne pas rejouer ?"""
        if self.reserve is None or event.code != self.reserve[1]:
            return False
        if event.value == 1:
            if self.modifiers_match(self.reserve[0]):
                device.swallowed.add(event.code)
                return True
            return False
        if event.code in device.swallowed:
            if event.value == 0:
                device.swallowed.discard(event.code)
            return True
        return False

    def on_device(self, device):
        try:
            events = list(device.dev.read())
        except OSError as e:
            if e.errno in (errno.ENODEV, errno.EIO):
                self.drop(device)
                return
            raise
        for event in events:
            if event.type == ecodes.EV_KEY:
                code, value = event.code, event.value
                if code < ecodes.BTN_MISC:
                    if value == 1:
                        self.held.add(code)
                    elif value == 0:
                        self.held.discard(code)
                    self.emit('k %d %d' % (code + X_OFFSET, value))
                elif 0 <= code - ecodes.BTN_SIDE < SIDE_BUTTONS and value in (0, 1):
                    self.emit('b %d %d' % (code - ecodes.BTN_SIDE + X_FIRST_SIDE_BUTTON, value))
                if device.clone is not None and self.swallow(device, event):
                    continue
            if device.clone is not None and event.type not in OUTPUT_TYPES:
                device.clone.write_event(event)
        if self.pending:
            self.apply_reserve()

    def on_clone(self, device):
        """Le bureau allume les voyants (Verr. Num…) du clone : on les reporte sur le vrai."""
        try:
            data = os.read(device.clone.fd, _EVENT.size * 16)
        except OSError:
            return
        for offset in range(0, len(data) - _EVENT.size + 1, _EVENT.size):
            _, _, type_, code, value = _EVENT.unpack_from(data, offset)
            if type_ == ecodes.EV_LED:
                try:
                    device.dev.set_led(code, value)
                except OSError:
                    pass

    def on_stdin(self):
        data = os.read(sys.stdin.fileno(), 4096)
        if not data:
            raise SystemExit(0)
        self.stdin_buffer += data
        *lines, self.stdin_buffer = self.stdin_buffer.split(b'\n')
        for line in lines:
            try:
                wanted = json.loads(line)
            except ValueError:
                continue
            if not wanted:
                self.reserve = None
            elif wanted.get('button'):
                code = int(wanted['button']) - X_FIRST_SIDE_BUTTON + ecodes.BTN_SIDE
                self.reserve = (frozenset(wanted.get('modifiers') or ()), code)
            else:
                code = int(wanted['key']) - X_OFFSET
                self.reserve = (frozenset(wanted.get('modifiers') or ()), code)
            self.apply_reserve()

    def run(self):
        self.scan()
        self.emit('ok')
        next_scan = time.monotonic() + RESCAN_SECONDS
        while True:
            timeout = max(0.0, next_scan - time.monotonic())
            for key, _ in self.selector.select(timeout):
                kind, device = key.data
                if kind == 'stdin':
                    self.on_stdin()
                elif device.path in self.devices:
                    if kind == 'device':
                        self.on_device(device)
                    elif device.clone is not None:
                        self.on_clone(device)
            if time.monotonic() >= next_scan:
                self.scan()
                next_scan = time.monotonic() + RESCAN_SECONDS

    def close(self):
        for device in list(self.devices.values()):
            if device.clone is not None:
                device.ungrab()
            device.dev.close()


def main():
    signal.signal(signal.SIGTERM, lambda *_: sys.exit(0))
    signal.signal(signal.SIGINT, signal.SIG_IGN)
    listener = Listener()
    try:
        listener.run()
    finally:
        listener.close()


if __name__ == '__main__':
    main()
