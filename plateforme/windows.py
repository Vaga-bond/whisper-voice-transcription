"""Fonctions propres à Windows : clavier, presse-papier, fenêtre active, forme de l'overlay."""
import ctypes
import time
from ctypes import wintypes

import pyautogui
from pynput import keyboard as pynput_keyboard

NAME = "Windows"
SETUP_SCRIPT = "setup.bat"

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)


# ── Clavier ──────────────────────────────────────────────────────────────────
# Une touche est désignée par son code virtuel Windows (VK) : il suit la touche physique,
# donc le raccourci reste le même en AZERTY comme en QWERTY.

MODIFIER_LABELS = {'ctrl': 'Ctrl', 'alt': 'Alt', 'shift': 'Maj', 'win': 'Win'}

_KEY_LABELS = {
    8: 'Retour arrière', 9: 'Tab', 13: 'Entrée', 19: 'Pause', 20: 'Verr. Maj',
    27: 'Échap', 32: 'Espace', 33: 'Page préc.', 34: 'Page suiv.', 35: 'Fin',
    36: 'Origine', 37: '←', 38: '↑', 39: '→', 40: '↓', 44: 'Impr. écran',
    45: 'Inser', 46: 'Suppr',
    106: 'Pavé num *', 107: 'Pavé num +', 109: 'Pavé num -',
    110: 'Pavé num .', 111: 'Pavé num /',
}
_KEY_LABELS.update({vk: chr(vk) for vk in range(48, 58)})                 # 0-9
_KEY_LABELS.update({vk: chr(vk) for vk in range(65, 91)})                 # A-Z
_KEY_LABELS.update({vk: 'Pavé num %d' % (vk - 96) for vk in range(96, 106)})
_KEY_LABELS.update({vk: 'F%d' % (vk - 111) for vk in range(112, 136)})

# Touches qu'on peut prendre seules sans qu'elles se déclenchent pendant la frappe.
SAFE_BARE_CODES = set(range(112, 136))                                    # F1-F24

DEFAULT_HOTKEY = {
    'modifiers': ['ctrl', 'alt'],
    # 9 de la rangée du haut (VK 57) et du pavé numérique (VK 105) : selon le
    # matériel et le logiciel de macro, l'un ou l'autre est émis.
    'vks': [57, 105],
}

# Maj, Ctrl, Alt, Win gauche et droite : enfoncés, ils détournent Ctrl+V.
_MODIFIER_VKS = (0x10, 0x11, 0x12, 0x5B, 0x5C)


def key_label(code):
    return _KEY_LABELS.get(code, 'VK %d' % code)


def _key_code(key):
    """Code VK d'une touche pynput (touche ordinaire ou touche spéciale comme F9)."""
    vk = getattr(key, 'vk', None)
    if vk is None and isinstance(key, pynput_keyboard.Key):
        vk = getattr(key.value, 'vk', None)
    return vk


def start_key_listener(on_press, on_release):
    """Écoute le clavier de tout le système ; on_press(touche, code) et on_release(touche, code)."""
    listener = pynput_keyboard.Listener(
        on_press=lambda key: on_press(key, _key_code(key)),
        on_release=lambda key: on_release(key, _key_code(key)))
    listener.start()
    return listener


def keys_held(codes):
    """Un modificateur ou l'une des touches `codes` est-il enfoncé en ce moment ?"""
    return any(_user32.GetAsyncKeyState(vk) & 0x8000 for vk in _MODIFIER_VKS + tuple(codes))


# Sous Windows, le raccourci n'est pas réservé et la souris n'est pas proposée.
GRABS_HOTKEY = False
SUPPORTS_MOUSE = False


def grab_hotkey(hotkey, on_button=None):
    return None


def send_paste(terminal):
    """Envoie Ctrl+V, ou Ctrl+Maj+V pour un terminal."""
    if terminal:
        pyautogui.hotkey('ctrl', 'shift', 'v')
    else:
        pyautogui.hotkey('ctrl', 'v')


# ── Fenêtre active ───────────────────────────────────────────────────────────

# Classes de fenêtre natives des terminaux connus. Les terminaux intégrés des éditeurs
# Electron (Cursor, VS Code) partagent la classe de l'éditeur : pour eux, l'option manuelle.
_TERMINAL_WINDOW_CLASSES = {
    'CASCADIA_HOSTING_WINDOW_CLASS',   # Windows Terminal
    'ConsoleWindowClass',              # cmd, PowerShell classique
    'PuTTY',
    'mintty',                          # Git Bash, msys2
}


def terminal_focused():
    """La fenêtre au premier plan est-elle un terminal natif ?"""
    hwnd = _user32.GetForegroundWindow()
    if not hwnd:
        return False
    buf = ctypes.create_unicode_buffer(256)
    _user32.GetClassNameW(hwnd, buf, 256)
    return buf.value in _TERMINAL_WINDOW_CLASSES


# ── Overlay ──────────────────────────────────────────────────────────────────

def _hwnd(win):
    return int(win.wm_frame(), 16)


def overlay_click_through(win):
    """Les clics traversent la fenêtre. Tk a déjà posé WS_EX_LAYERED via `-alpha` : on n'y
    touche pas pour éviter l'état « fenêtre noire ». SWP_FRAMECHANGED valide la modification."""
    GWL_EXSTYLE = -20
    WS_EX_TRANSPARENT = 0x00000020
    SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_FRAMECHANGED = 0x0002, 0x0001, 0x0004, 0x0020
    hwnd = _hwnd(win)
    style = _user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    _user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_TRANSPARENT)
    _user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                         SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)


def overlay_round_corners(win, radius):
    w, h = win.winfo_width(), win.winfo_height()
    if w <= 1 or h <= 1:
        return
    region = _gdi32.CreateRoundRectRgn(0, 0, w + 1, h + 1, radius, radius)
    _user32.SetWindowRgn(_hwnd(win), region, True)


def overlay_circle(win):
    w, h = win.winfo_width(), win.winfo_height()
    if w <= 1 or h <= 1:
        return
    region = _gdi32.CreateEllipticRgn(0, 0, w + 1, h + 1)
    _user32.SetWindowRgn(_hwnd(win), region, True)


# ── Micro ────────────────────────────────────────────────────────────────────

def mic_hidden(name):
    """Entrée audio à cacher de la liste des micros (le filtre commun est dans app.py)."""
    return False


# ── Session ──────────────────────────────────────────────────────────────────

def environment_problem():
    """Raison pour laquelle le raccourci ou le collage ne peuvent pas marcher, sinon None."""
    return None


HOTKEY_FAILURE_HINT = "Essayez de lancer l'application en tant qu'administrateur."


# ── Presse-papier ────────────────────────────────────────────────────────────
# pyperclip ne lit que le texte : pour rendre le presse-papier après un collage, il faut en
# garder chaque format, sinon une image copiée avant la dictée serait perdue.

_user32.OpenClipboard.argtypes = [wintypes.HWND]
_user32.OpenClipboard.restype = wintypes.BOOL
_user32.CloseClipboard.restype = wintypes.BOOL
_user32.EmptyClipboard.restype = wintypes.BOOL
_user32.EnumClipboardFormats.argtypes = [wintypes.UINT]
_user32.EnumClipboardFormats.restype = wintypes.UINT
_user32.GetClipboardData.argtypes = [wintypes.UINT]
_user32.GetClipboardData.restype = wintypes.HANDLE
_user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
_user32.SetClipboardData.restype = wintypes.HANDLE
_user32.GetClipboardSequenceNumber.restype = wintypes.DWORD
_kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
_kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
_kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
_kernel32.GlobalLock.restype = wintypes.LPVOID
_kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
_kernel32.GlobalUnlock.restype = wintypes.BOOL
_kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
_kernel32.GlobalSize.restype = ctypes.c_size_t
_kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
_kernel32.GlobalFree.restype = wintypes.HGLOBAL

_GMEM_MOVEABLE = 0x0002
# Formats portés par un objet graphique et non par un bloc mémoire (CF_BITMAP, CF_PALETTE,
# métafichiers, formats « display ») : Windows les recrée depuis leurs équivalents en mémoire,
# CF_DIB pour une image.
_HANDLE_FORMATS = {2, 3, 9, 14, 0x80, 0x82, 0x83, 0x8E}


def clipboard_sequence():
    """Numéro qui change à chaque écriture dans le presse-papier, par n'importe quelle application."""
    return _user32.GetClipboardSequenceNumber()


def _open():
    # Une autre application peut tenir le presse-papier quelques millisecondes.
    for _ in range(50):
        if _user32.OpenClipboard(None):
            return
        time.sleep(0.01)
    raise OSError("presse-papier occupé par une autre application")


def clipboard_snapshot():
    """Contenu du presse-papier : liste de (format, octets)."""
    _open()
    try:
        items = []
        fmt = _user32.EnumClipboardFormats(0)
        while fmt:
            if fmt not in _HANDLE_FORMATS:
                handle = _user32.GetClipboardData(fmt)
                ptr = _kernel32.GlobalLock(handle) if handle else None
                if ptr:
                    try:
                        items.append((fmt, ctypes.string_at(ptr, _kernel32.GlobalSize(handle))))
                    finally:
                        _kernel32.GlobalUnlock(handle)
            fmt = _user32.EnumClipboardFormats(fmt)
        return items
    finally:
        _user32.CloseClipboard()


def clipboard_restore(items):
    """Remet un contenu pris par clipboard_snapshot()."""
    _open()
    try:
        _user32.EmptyClipboard()
        for fmt, data in items:
            handle = _kernel32.GlobalAlloc(_GMEM_MOVEABLE, max(len(data), 1))
            if not handle:
                raise MemoryError("mémoire insuffisante pour rendre le presse-papier")
            ptr = _kernel32.GlobalLock(handle)
            ctypes.memmove(ptr, data, len(data))
            _kernel32.GlobalUnlock(handle)
            # Accepté, le bloc appartient au système ; refusé, il reste à libérer.
            if not _user32.SetClipboardData(fmt, handle):
                _kernel32.GlobalFree(handle)
    finally:
        _user32.CloseClipboard()
