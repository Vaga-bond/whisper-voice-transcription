"""Ce qui dépend du système : chaque module expose les mêmes fonctions.

À importer avant pynput : sous Linux, l'écoute du clavier doit passer par X11.
"""
import os
import sys

if sys.platform == "win32":
    from .windows import *  # noqa: F401,F403
elif sys.platform.startswith("linux"):
    os.environ.setdefault("PYNPUT_BACKEND", "xorg")
    from .linux import *  # noqa: F401,F403
else:
    raise ImportError(f"Système non pris en charge : {sys.platform}")
