"""Copie complète du presse-papier Windows, tous formats compris (texte, image, fichiers…).

pyperclip ne lit que le texte : pour rendre le presse-papier après un collage, il faut en
garder chaque format, sinon une image copiée avant la dictée serait perdue.
"""
import ctypes
import time
from ctypes import wintypes

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

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


def sequence():
    """Numéro qui change à chaque écriture dans le presse-papier, par n'importe quelle application."""
    return _user32.GetClipboardSequenceNumber()


def _open():
    # Une autre application peut tenir le presse-papier quelques millisecondes.
    for _ in range(50):
        if _user32.OpenClipboard(None):
            return
        time.sleep(0.01)
    raise OSError("presse-papier occupé par une autre application")


def snapshot():
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


def restore(items):
    """Remet un contenu pris par snapshot()."""
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
