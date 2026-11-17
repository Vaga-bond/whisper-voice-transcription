"""Ce qui dépend du système : chaque module expose les mêmes fonctions.

À importer avant pynput : sous Linux, pynput doit passer par X11 (XWayland en session
Wayland, où l'interface tourne aussi). En session Wayland, clavier, souris, collage et
presse-papier passent par wayland.py.
"""
import os
import sys

if sys.platform == "win32":
    from .windows import *  # noqa: F401,F403
elif sys.platform.startswith("linux"):
    os.environ.setdefault("PYNPUT_BACKEND", "xorg")
    if os.environ.get("XDG_SESSION_TYPE") == "wayland":
        from .wayland import *  # noqa: F401,F403
    else:
        from .linux import *  # noqa: F401,F403
else:
    raise ImportError(f"Système non pris en charge : {sys.platform}")
