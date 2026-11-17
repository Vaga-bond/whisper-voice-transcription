#!/usr/bin/env python3
"""
Transcription vocale, locale ou OpenAI, pour Windows et Linux (X11).
Ce qui dépend du système est dans le paquet `plateforme`.
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import socket
import sys
import os
import json
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path
import plateforme   # avant pynput : il choisit le moteur clavier sous Linux
import pyperclip
from dotenv import load_dotenv
from openai import OpenAI
import sounddevice as sd
import numpy as np
import wave
from pynput import keyboard as pynput_keyboard
import pygame

from local_engine import (LocalEngine, LOCAL_MODELS, download as download_model, find_model,
                          missing_packages, missing_requirement)

# Sans ça, les emojis des prints crashent une console Windows en cp1252.
for flux in (sys.stdout, sys.stderr):
    if flux is not None:
        flux.reconfigure(encoding='utf-8', errors='replace')

load_dotenv()

# Modèles compatibles v1/audio/transcriptions : (libellé UI, identifiant API).
# gpt-live-transcribe et gpt-realtime-whisper sont exclus, ils ne répondent que
# sur les sessions Realtime WebSocket.
MODEL_OPTIONS = [
    ("GPT Transcribe", "gpt-transcribe"),
    ("GPT-4o Transcribe", "gpt-4o-transcribe"),
    ("GPT-4o Mini Transcribe", "gpt-4o-mini-transcribe"),
    ("Whisper-1", "whisper-1"),
    *[(spec.label, model_id) for model_id, spec in LOCAL_MODELS.items()],
]

# Tarifs OpenAI. "duration" = prix ferme par minute, "tokens" = prix par million
# de tokens. `est_per_minute` ne sert que de repli si l'API ne renvoie pas `usage`.
MODEL_PRICING = {
    "gpt-transcribe": {
        "mode": "duration", "per_minute": 0.0045, "est_per_minute": 0.0045,
    },
    "gpt-4o-transcribe": {
        "mode": "tokens", "input_per_1m": 2.50, "output_per_1m": 10.00,
        "est_per_minute": 0.0033,
    },
    "gpt-4o-mini-transcribe": {
        "mode": "tokens", "input_per_1m": 1.25, "output_per_1m": 5.00,
        "est_per_minute": 0.0016,
    },
    "whisper-1": {
        "mode": "duration", "per_minute": 0.006, "est_per_minute": 0.006,
    },
    **{model_id: {"mode": "local", "est_per_minute": 0.0} for model_id in LOCAL_MODELS},
}

def compute_transcription_cost(model, usage, duration_sec):
    """Retourne (coût USD, exact) d'après le `usage` renvoyé par l'API."""
    pricing = MODEL_PRICING.get(model)
    if pricing is None:
        return 0.0, False
    if pricing["mode"] == "local":
        return 0.0, True

    usage_type = getattr(usage, "type", None)
    try:
        if usage_type == "duration":
            return usage.seconds / 60 * pricing["per_minute"], True
        if usage_type == "tokens":
            cost = (usage.input_tokens / 1_000_000 * pricing["input_per_1m"]
                    + usage.output_tokens / 1_000_000 * pricing["output_per_1m"])
            return cost, True
    except (AttributeError, TypeError, KeyError):
        pass

    return duration_sec / 60 * pricing["est_per_minute"], False

HISTORY_FILE = Path(__file__).parent / "transcription_history.json"

# Option de lancement : démarrer rangée dans la barre système, sans ouvrir la fenêtre.
START_IN_TRAY = "--barre" in sys.argv[1:]

PREFS_FILE = Path(__file__).parent / "user_preferences.json"


def _build_modifier_map():
    """Associe chaque variante de modificateur pynput (ctrl_l, ctrl_r…) à un nom canonique."""
    variants = {
        'ctrl': ('ctrl', 'ctrl_l', 'ctrl_r'),
        'alt': ('alt', 'alt_l', 'alt_r', 'alt_gr'),
        'shift': ('shift', 'shift_l', 'shift_r'),
        'win': ('cmd', 'cmd_l', 'cmd_r'),
    }
    mapping = {}
    for canonical, attrs in variants.items():
        for attr in attrs:
            key = getattr(pynput_keyboard.Key, attr, None)
            if key is not None:
                mapping[key] = canonical
    return mapping

MODIFIER_NAMES = _build_modifier_map()
MODIFIER_ORDER = ('ctrl', 'alt', 'shift', 'win')
MODIFIER_LABELS = plateforme.MODIFIER_LABELS

# Une touche est désignée par un code qui suit sa position physique (code VK sous Windows,
# keycode X sous Linux) : le raccourci reste le même en AZERTY comme en QWERTY. Le champ
# s'appelle « vks » dans les préférences depuis la version Windows.

# Le système réémet « touche enfoncée » tant que la touche reste pressée : sans ces
# deux gardes, un appui un peu long bascule deux fois et coupe l'enregistrement.
HOTKEY_REPEAT_GAP = 1.0      # au-delà, deux appuis ne peuvent plus être une répétition
HOTKEY_MIN_INTERVAL = 0.35   # intervalle minimum entre deux bascules réelles

# Le collage attend que le raccourci soit relâché : tenu, Ctrl+V deviendrait
# Ctrl+Alt+V et Alt + pavé numérique un code Alt (« ○ » pour Alt+9).
PASTE_KEY_POLL_MS = 10
PASTE_RELEASE_TIMEOUT = 2.0   # au-delà, le texte attend dans le presse-papier
# L'application cible lit le presse-papier en traitant Ctrl+V, parfois avec retard.
CLIPBOARD_RESTORE_DELAY_MS = 300
DEFAULT_HOTKEY = plateforme.DEFAULT_HOTKEY

def describe_hotkey(hotkey):
    """Rend un raccourci lisible : {'modifiers': ['ctrl','alt'], 'vks': [57]} -> 'Ctrl+Alt+9'."""
    parts = [MODIFIER_LABELS[m] for m in MODIFIER_ORDER if m in hotkey['modifiers']]
    vks = hotkey['vks']
    if hotkey.get('button'):
        parts.append(plateforme.button_label(hotkey['button']))
    else:
        parts.append(plateforme.key_label(vks[0]) if vks else '?')
    return '+'.join(parts)

def normalize_hotkey(value, default=DEFAULT_HOTKEY):
    """Valide un raccourci lu depuis les préférences, sinon retourne `default`.

    Un bouton de souris est noté {'button': 8} avec 'vks' vide : sur un système sans souris
    pour le raccourci, il retombe sur le défaut."""
    try:
        modifiers = [m for m in MODIFIER_ORDER if m in value['modifiers']]
        button = value.get('button')
        if button is not None and plateforme.SUPPORTS_MOUSE:
            button = int(button)
            if plateforme.MOUSE_BUTTON_MIN <= button <= plateforme.MOUSE_BUTTON_MAX:
                return {'modifiers': modifiers, 'vks': [], 'button': button}
        vks = [int(v) for v in value['vks'] if 0 < int(v) < 256]
        if vks:
            return {'modifiers': modifiers, 'vks': vks}
    except (KeyError, TypeError, ValueError):
        pass
    return {'modifiers': list(default['modifiers']), 'vks': list(default['vks'])}

DEFAULT_PREFS = {
    "selected_model": "whisper-large-v3-turbo-local",
    "selected_device_name": None,   # Nom du micro (pas l'index — plus stable entre sessions)
    "max_recording_duration": 240,
    "auto_copy": True,
    "auto_paste": True,
    "sound_enabled": True,
    "overlay_enabled": True,
    "minimize_to_tray_on_close": True,
    "terminal_paste": False,
    "overlay_position": None,
    "hotkey": DEFAULT_HOTKEY,
    "pause_hotkey": plateforme.DEFAULT_PAUSE_HOTKEY,
    "hotkey_paused": False,
}

ENV_FILE = Path(__file__).parent / ".env"

# Le script d'installation copie .env.example vers .env : OPENAI_API_KEY est donc défini avec la
# valeur d'exemple avant toute saisie de clé.
API_KEY_PLACEHOLDERS = {
    "votre_cle_api_ici", "your_api_key_here", "sk-xxx", "sk-...",
}

def read_api_key():
    """Retourne la clé API si elle est réellement renseignée, sinon None."""
    key = (os.getenv('OPENAI_API_KEY') or '').strip()
    if not key or key in API_KEY_PLACEHOLDERS:
        return None
    return key

SINGLE_INSTANCE_PORT = 49731

# pystray + Pillow sont optionnels : l'app démarre aussi sans, la tray est juste désactivée.
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

class ToolTip:
    """Tooltip minimaliste en pur tkinter (pas de dep).
    Apparaît après 400ms de survol, disparaît sur Leave ou clic.

    `text` accepte une chaîne, ou un callable évalué à chaque survol pour un
    contenu qui dépend de l'état courant."""

    def __init__(self, widget, text, wraplength=240, delay_ms=400):
        self.widget = widget
        self.text = text
        self.wraplength = wraplength
        self.delay_ms = delay_ms
        self.tip_window = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self.tip_window is not None:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 2
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.overrideredirect(True)
        self.tip_window.attributes('-topmost', True)
        self.tip_window.geometry(f"+{x}+{y}")
        tk.Label(
            self.tip_window, text=self.text() if callable(self.text) else self.text,
            background="#ffffe0", relief="solid", borderwidth=1,
            font=("Arial", 8), wraplength=self.wraplength,
            justify="left", padx=6, pady=4,
        ).pack()

    def _hide(self, event=None):
        self._cancel()
        if self.tip_window is not None:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None

class FloatingOverlay:
    """Overlay translucide pour afficher l'état courant (enregistrement,
    transcription, résultat). Visible même quand l'app est réduite/dans la tray.

    Composé de deux Toplevels :
    - `self.win` : panneau message, click-through (clics passent derrière),
      coins arrondis, semi-transparent
    - `self.handle` : petite bulle circulaire non-click-through, attachée au coin
      haut-gauche de l'overlay. Permet de drag & drop pour déplacer l'ensemble.
      La position est persistée via le callback `on_position_saved`.
    """

    CORNER_RADIUS = 14
    HANDLE_DEBORD = 2
    ALPHA_OVERLAY = 0.68
    ALPHA_HANDLE_REST = 0.78
    ALPHA_HANDLE_ACTIVE = 0.95

    def __init__(self, root, on_position_saved=None):
        self.root = root
        self.enabled = True
        self.custom_position = None  # (x, y) coin haut-gauche de l'overlay, None = défaut centré-haut
        self.on_position_saved = on_position_saved
        self._hide_after_id = None
        self._handle_show_after_id = None
        self._drag_offset_x = 0
        self._drag_offset_y = 0
        self._drag_handle_dx = 0
        self._drag_handle_dy = 0
        self._is_dragging = False

        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes('-topmost', True)
        self.win.attributes('-alpha', self.ALPHA_OVERLAY)
        self.win.configure(bg='#c62828')
        self.label = tk.Label(
            self.win, text="", font=("Segoe UI", 13, "bold"),
            bg='#c62828', fg='white', padx=22, pady=8,
        )
        self.label.pack()

        self.handle = tk.Toplevel(root)
        self.handle.overrideredirect(True)
        self.handle.attributes('-topmost', True)
        self.handle.attributes('-alpha', self.ALPHA_HANDLE_REST)
        self.handle.configure(bg='#c62828')
        self.handle_label = tk.Label(
            self.handle, text="⠿", font=("Segoe UI", 10, "bold"),
            bg='#c62828', fg='#ffffff', cursor='fleur',
        )
        self.handle_label.pack(padx=2, pady=0)
        for widget in (self.handle, self.handle_label):
            widget.bind('<Button-1>', self._on_drag_start)
            widget.bind('<B1-Motion>', self._on_drag_motion)
            widget.bind('<ButtonRelease-1>', self._on_drag_release)
            widget.bind('<Enter>', self._on_handle_enter, add='+')
            widget.bind('<Leave>', self._on_handle_leave, add='+')

        self.win.withdraw()
        self.handle.withdraw()

        # La forme et le click-through exigent que les fenêtres existent déjà côté système.
        self.win.update_idletasks()
        self.handle.update_idletasks()
        self._enable_click_through()
        self._apply_rounded_corners_to_overlay()
        self._apply_circular_shape_to_handle()

    def _enable_click_through(self):
        """Les clics traversent l'overlay principal."""
        try:
            plateforme.overlay_click_through(self.win)
        except Exception as e:
            print(f"⚠️ Click-through overlay non activé: {e}")

    def _apply_rounded_corners_to_overlay(self):
        try:
            plateforme.overlay_round_corners(self.win, self.CORNER_RADIUS)
        except Exception as e:
            print(f"⚠️ Arrondi overlay non appliqué: {e}")

    def _apply_circular_shape_to_handle(self):
        try:
            plateforme.overlay_circle(self.handle)
        except Exception as e:
            print(f"⚠️ Poignée circulaire non appliquée: {e}")

    def _reposition(self):
        """Positionne l'overlay selon `custom_position` (sauvegardée) ou sinon
        au centre-haut de l'écran principal. Clampe aux bornes de l'écran pour
        éviter qu'il ne sorte après changement de résolution ou d'écran principal."""
        self.win.update_idletasks()
        self.handle.update_idletasks()
        screen_w = self.win.winfo_screenwidth()
        screen_h = self.win.winfo_screenheight()
        w = self.win.winfo_width()
        h = self.win.winfo_height()

        if self.custom_position is not None:
            x, y = self.custom_position
        else:
            x = max((screen_w - w) // 2, 10)
            y = 30

        x = max(0, min(x, max(0, screen_w - w)))
        y = max(0, min(y, max(0, screen_h - h)))

        self.win.geometry(f"+{x}+{y}")
        self._position_handle()

    def _position_handle(self):
        """Place la poignée en bas-droite de l'overlay, en très léger débord."""
        self.win.update_idletasks()
        self.handle.update_idletasks()
        overlay_x = self.win.winfo_rootx()
        overlay_y = self.win.winfo_rooty()
        overlay_w = self.win.winfo_width()
        overlay_h = self.win.winfo_height()
        handle_w = self.handle.winfo_width()
        handle_h = self.handle.winfo_height()
        hx = overlay_x + overlay_w - handle_w + self.HANDLE_DEBORD
        hy = overlay_y + overlay_h - handle_h + self.HANDLE_DEBORD
        self.handle.geometry(f"+{max(0, hx)}+{max(0, hy)}")

    def _on_handle_enter(self, event):
        """Au survol : poignée plus opaque pour retour visuel."""
        try:
            self.handle.attributes('-alpha', self.ALPHA_HANDLE_ACTIVE)
        except Exception:
            pass

    def _on_handle_leave(self, event):
        """Quand la souris quitte : retour à l'alpha discret (sauf pendant un drag)."""
        if self._is_dragging:
            return
        try:
            self.handle.attributes('-alpha', self.ALPHA_HANDLE_REST)
        except Exception:
            pass

    def _on_drag_start(self, event):
        self._drag_offset_x = event.x_root - self.handle.winfo_rootx()
        self._drag_offset_y = event.y_root - self.handle.winfo_rooty()
        self._drag_handle_dx = self.handle.winfo_rootx() - self.win.winfo_rootx()
        self._drag_handle_dy = self.handle.winfo_rooty() - self.win.winfo_rooty()
        self._is_dragging = True
        try:
            self.handle.attributes('-alpha', self.ALPHA_HANDLE_ACTIVE)
        except Exception:
            pass

    def _on_drag_motion(self, event):
        new_handle_x = event.x_root - self._drag_offset_x
        new_handle_y = event.y_root - self._drag_offset_y
        new_overlay_x = new_handle_x - self._drag_handle_dx
        new_overlay_y = new_handle_y - self._drag_handle_dy

        screen_w = self.win.winfo_screenwidth()
        screen_h = self.win.winfo_screenheight()
        w = self.win.winfo_width()
        h = self.win.winfo_height()
        new_overlay_x = max(0, min(new_overlay_x, max(0, screen_w - w)))
        new_overlay_y = max(0, min(new_overlay_y, max(0, screen_h - h)))

        self.win.geometry(f"+{new_overlay_x}+{new_overlay_y}")
        self._position_handle()

    def _on_drag_release(self, event):
        self._is_dragging = False
        try:
            self.handle.attributes('-alpha', self.ALPHA_HANDLE_REST)
        except Exception:
            pass
        x = self.win.winfo_rootx()
        y = self.win.winfo_rooty()
        self.custom_position = (x, y)
        if self.on_position_saved:
            self.on_position_saved(x, y)

    def show(self, text, bg='#c62828'):
        if not self.enabled:
            self.hide()
            return
        if self._hide_after_id is not None:
            self.root.after_cancel(self._hide_after_id)
            self._hide_after_id = None

        self.label.config(text=text, bg=bg)
        self.win.configure(bg=bg)
        self.handle.configure(bg=bg)
        self.handle_label.configure(bg=bg)

        # Tk ne finalise la géométrie d'un Toplevel qu'une fois mappé : au premier
        # show la fenêtre apparaîtrait à une taille pré-layout. On la matérialise
        # donc invisible et hors écran avant de la placer réellement.
        first_show = not self.win.winfo_viewable()
        if first_show:
            self.win.attributes('-alpha', 0.0)
            self.win.geometry("+-10000+-10000")
            self.win.deiconify()
            self.win.update()   # force la matérialisation complète du layout
            self._apply_rounded_corners_to_overlay()
            self._reposition()
            self.win.attributes('-alpha', self.ALPHA_OVERLAY)
        else:
            self.win.update_idletasks()
            self._apply_rounded_corners_to_overlay()
            self._reposition()

        self.win.attributes('-topmost', True)

        # Différé : laisse tk finaliser la géométrie avant de placer la poignée.
        if not self.handle.winfo_viewable():
            if self._handle_show_after_id is not None:
                try:
                    self.root.after_cancel(self._handle_show_after_id)
                except Exception:
                    pass
            self._handle_show_after_id = self.root.after(30, self._show_handle_delayed)

    def _show_handle_delayed(self):
        """Affiche la poignée après que l'overlay principal ait fini son layout."""
        self._handle_show_after_id = None
        if not self.enabled or not self.win.winfo_viewable():
            return
        self._position_handle()
        self.handle.deiconify()
        self.handle.attributes('-topmost', True)

    def show_briefly(self, text, bg, duration_ms=1500):
        if not self.enabled:
            return
        self.show(text, bg)
        self._hide_after_id = self.root.after(duration_ms, self.hide)

    def hide(self):
        if self._hide_after_id is not None:
            try:
                self.root.after_cancel(self._hide_after_id)
            except Exception:
                pass
            self._hide_after_id = None
        if self._handle_show_after_id is not None:
            try:
                self.root.after_cancel(self._handle_show_after_id)
            except Exception:
                pass
            self._handle_show_after_id = None
        self.win.withdraw()
        self.handle.withdraw()

class VoiceTranscriptionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Transcription Vocale - Whisper")
        self.root.geometry("800x800")
        self.root.resizable(True, True)
        self.root.minsize(600, 500)
        
        self.is_recording = False
        self.is_transcribing = False
        self.cancel_requested = False
        self.audio_frames = []
        self.recording_thread = None
        self.recording_start_time = None
        self.recording_model = None
        self._downloads = set()   # modèles locaux en cours de téléchargement
        self.selected_device_index = None
        self.local_engine = LocalEngine(language="fr")

        self.prefs = self._load_prefs()

        self.selected_model = self.prefs.get("selected_model", DEFAULT_PREFS["selected_model"])
        if self.selected_model not in MODEL_PRICING:
            self.selected_model = DEFAULT_PREFS["selected_model"]

        self.preferred_mic_name = self.prefs.get("selected_device_name")

        self.session_cost = 0.0
        self.session_transcriptions = 0
        self.session_started_at = datetime.now()

        self.history = self._load_history()

        self.hotkey_thread = None
        self.esc_thread = None

        self.CHANNELS = 1
        self.RATE = 44100
        self.DTYPE = np.int16

        self.max_recording_duration = int(self.prefs.get("max_recording_duration", 240))

        self.hotkey = normalize_hotkey(self.prefs.get("hotkey"))
        # Raccourci de pause : clavier seulement, modificateurs exacts (Ctrl+9 ne doit pas
        # répondre à Ctrl+Alt+9). En pause, le raccourci de dictée est rendu au système.
        self.pause_hotkey = self._keyboard_only(normalize_hotkey(
            self.prefs.get("pause_hotkey"), plateforme.DEFAULT_PAUSE_HOTKEY))
        self.hotkey_paused = bool(self.prefs.get("hotkey_paused", False))
        self._pause_down = False
        self._hotkey_capture = None

        self._hotkey_down = False
        self._hotkey_last_press = 0.0
        self._hotkey_last_toggle = 0.0

        self.microphone_name = "Chargement..."

        # BooleanVars créées ici (avant setup_ui) pour pouvoir attacher les traces
        # et appliquer directement les valeurs persistées.
        self.auto_copy_var = tk.BooleanVar(value=bool(self.prefs.get("auto_copy", True)))
        self.auto_paste_var = tk.BooleanVar(value=bool(self.prefs.get("auto_paste", True)))
        self.sound_enabled_var = tk.BooleanVar(value=bool(self.prefs.get("sound_enabled", True)))
        self.overlay_enabled_var = tk.BooleanVar(value=bool(self.prefs.get("overlay_enabled", True)))
        self.minimize_to_tray_var = tk.BooleanVar(value=bool(self.prefs.get("minimize_to_tray_on_close", True)))
        self.terminal_paste_var = tk.BooleanVar(value=bool(self.prefs.get("terminal_paste", False)))
        self.auto_copy_var.trace_add('write', lambda *_: self._save_prefs())
        self.auto_paste_var.trace_add('write', lambda *_: self._on_paste_toggles_changed())
        self.sound_enabled_var.trace_add('write', lambda *_: self._save_prefs())
        self.overlay_enabled_var.trace_add('write', lambda *_: self._on_overlay_toggle())
        self.minimize_to_tray_var.trace_add('write', lambda *_: self._save_prefs())
        self.terminal_paste_var.trace_add('write', lambda *_: self._on_paste_toggles_changed())

        # Non bloquant sans clé : elle peut être saisie depuis l'interface.
        self.client = None
        api_key = read_api_key()
        if api_key:
            try:
                self.client = OpenAI(api_key=api_key)
            except Exception as e:
                print(f"⚠️ Init OpenAI échouée (clé probablement invalide): {e}")

        self.sounds_dir = os.path.join(os.path.dirname(__file__), 'sounds')
        os.makedirs(self.sounds_dir, exist_ok=True)

        self.setup_ui()

        self.overlay = FloatingOverlay(
            self.root,
            on_position_saved=self._on_overlay_position_saved
        )
        self.overlay.enabled = self.overlay_enabled_var.get()
        saved_pos = self.prefs.get("overlay_position")
        if isinstance(saved_pos, (list, tuple)) and len(saved_pos) == 2:
            try:
                self.overlay.custom_position = (int(saved_pos[0]), int(saved_pos[1]))
            except (TypeError, ValueError):
                pass

        # Différées pour que la fenêtre s'affiche instantanément.
        self.tray_icon = None
        self.root.after(100, self._deferred_init)
    
    def setup_ui(self):
        """Crée l'interface utilisateur"""
        
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        left_panel = tk.Frame(main_container, width=240)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_panel.pack_propagate(False)
        
        main_frame = tk.Frame(main_container)
        main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        title_label = tk.Label(
            main_frame,
            text="🎙️ Transcription Vocale",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        record_button_frame = tk.Frame(main_frame)
        record_button_frame.pack(pady=10)
        
        self.record_button = tk.Button(
            record_button_frame,
            text="🎤 Démarrer l'enregistrement",
            font=("Arial", 12),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            command=self.toggle_recording,
            width=25,
            height=2
        )
        self.record_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = tk.Button(
            record_button_frame,
            text="❌ Annuler",
            font=("Arial", 12),
            bg="#ff9800",
            fg="white",
            activebackground="#f57c00",
            command=self.cancel_recording,
            width=15,
            height=2
        )
        
        self.hotkey_label = tk.Label(
            main_frame,
            text=self._hotkey_banner(),
            font=("Arial", 9),
            fg="blue"
        )
        self.hotkey_label.pack(pady=(5, 10))
        
        options_frame = tk.LabelFrame(
            left_panel,
            text="Options",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        options_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tk.Label(
            options_frame,
            text="Clé API OpenAI:",
            font=("Arial", 9, "bold"),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 2))

        self.api_key_status_label = tk.Label(
            options_frame,
            text="…",
            font=("Arial", 8),
            anchor="w"
        )
        self.api_key_status_label.pack(anchor=tk.W, fill=tk.X)

        tk.Button(
            options_frame,
            text="Modifier la clé…",
            font=("Arial", 8),
            command=self._open_api_key_dialog,
            anchor="w"
        ).pack(anchor=tk.W, pady=(2, 5))

        tk.Frame(options_frame, height=1, bg="#cccccc").pack(fill=tk.X, pady=(4, 8))

        tk.Label(
            options_frame,
            text="Raccourci:",
            font=("Arial", 9, "bold"),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 2))

        self.hotkey_pref_label = tk.Label(
            options_frame,
            text=describe_hotkey(self.hotkey),
            font=("Arial", 8),
            anchor="w"
        )
        self.hotkey_pref_label.pack(anchor=tk.W, fill=tk.X)

        tk.Button(
            options_frame,
            text="Modifier le raccourci…",
            font=("Arial", 8),
            command=self._open_hotkey_dialog,
            anchor="w"
        ).pack(anchor=tk.W, pady=(2, 5))

        tk.Label(
            options_frame,
            text="Pause du raccourci:",
            font=("Arial", 9, "bold"),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(4, 2))

        self.pause_pref_label = tk.Label(
            options_frame,
            text=self._pause_description(),
            font=("Arial", 8),
            anchor="w"
        )
        self.pause_pref_label.pack(anchor=tk.W, fill=tk.X)

        tk.Button(
            options_frame,
            text="Modifier la pause…",
            font=("Arial", 8),
            command=lambda: self._open_hotkey_dialog('pause'),
            anchor="w"
        ).pack(anchor=tk.W, pady=(2, 5))

        tk.Frame(options_frame, height=1, bg="#cccccc").pack(fill=tk.X, pady=(4, 8))

        tk.Label(
            options_frame,
            text="Micro:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(0, 2))
        
        self.mic_display_label = tk.Label(
            options_frame,
            text=getattr(self, 'microphone_name', 'Chargement...'),
            font=("Arial", 8),
            fg="gray",
            wraplength=210,
            justify=tk.LEFT,
            anchor="w"
        )
        self.mic_display_label.pack(anchor=tk.W, fill=tk.X, pady=(0, 5))
        
        tk.Label(
            options_frame,
            text="Changer:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(5, 2))
        
        self.mic_var = tk.StringVar()
        self.mic_dropdown = tk.OptionMenu(options_frame, self.mic_var, "Chargement...")
        self.mic_dropdown.config(width=24, font=("Arial", 8), anchor="w")
        self.mic_dropdown.pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        tk.Label(
            options_frame,
            text="Modèle:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(5, 2))

        default_model_label = next(
            (label for label, api in MODEL_OPTIONS if api == self.selected_model),
            MODEL_OPTIONS[0][0]
        )
        self.model_var = tk.StringVar(value=default_model_label)
        self.model_dropdown = tk.OptionMenu(options_frame, self.model_var, default_model_label)
        self.model_dropdown.config(width=24, font=("Arial", 8), anchor="w")
        self.model_dropdown.pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        model_menu = self.model_dropdown['menu']
        model_menu.delete(0, 'end')
        for label, api_name in MODEL_OPTIONS:
            model_menu.add_command(
                label=label,
                command=lambda l=label, n=api_name: self._change_model(l, n)
            )

        ToolTip(
            self.model_dropdown,
            "Coût par minute d'audio :\n\n"
            "• GPT Transcribe — 0,0045 $ ferme (précision la plus haute)\n"
            "• GPT-4o Transcribe — ≈ 0,0033 $ aux tokens\n"
            "• GPT-4o Mini Transcribe — ≈ 0,0016 $ aux tokens\n"
            "• Whisper-1 — 0,006 $ ferme (ancienne génération)\n"
            "• Modèles « local » — gratuits, sur l'ordinateur, sans clé API. Nemotron "
            "transcrit pendant que vous parlez, les autres à l'arrêt.\n\n"
            "Les modèles aux tokens varient selon la densité de parole ; "
            "le compteur ci-dessous utilise le coût réel renvoyé par l'API.",
            wraplength=280
        )

        tk.Label(
            options_frame,
            text="Durée max:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(5, 2))
        
        self.duration_var = tk.IntVar(value=self.max_recording_duration)
        self.duration_slider = tk.Scale(
            options_frame,
            from_=5,
            to=900,  # 15 minutes max
            orient=tk.HORIZONTAL,
            variable=self.duration_var,
            length=150,
            command=self._on_duration_change,
            resolution=30,  # Snap à 30 secondes (demi-minutes)
            showvalue=False
        )
        self.duration_slider.pack(fill=tk.X, pady=5, padx=5)
        
        self.duration_label = tk.Label(
            options_frame,
            text="4 min",
            font=("Arial", 9, "bold"),
            fg="blue"
        )
        self.duration_label.pack(pady=(0, 5))

        # Copie et collage sont décorrélés : coller sans copier restaure le
        # presse-papier à sa valeur d'origine après le collage.
        tk.Checkbutton(
            options_frame,
            text="Copier dans le presse-papier",
            variable=self.auto_copy_var,
            font=("Arial", 8),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(5, 0))

        self.auto_paste_checkbox = tk.Checkbutton(
            options_frame,
            text="Coller dans le champ actif (Ctrl+V)",
            variable=self.auto_paste_var,
            font=("Arial", 8),
            anchor="w"
        )
        self.auto_paste_checkbox.pack(anchor=tk.W, fill=tk.X)

        self.terminal_paste_checkbox = tk.Checkbutton(
            options_frame,
            text="Coller pour terminal (Ctrl+Maj+V)",
            variable=self.terminal_paste_var,
            font=("Arial", 8),
            anchor="w"
        )
        self.terminal_paste_checkbox.pack(anchor=tk.W, fill=tk.X)

        ToolTip(
            self.terminal_paste_checkbox,
            "Par défaut, le collage se fait avec Ctrl+V (fonctionne partout).\n\n"
            "À cocher si vous collez souvent dans un terminal intégré "
            "(Cursor, VS Code…) qui n'accepte que Ctrl+Maj+V. "
            "Les terminaux autonomes sont détectés automatiquement.\n\n"
            "⚠ Quelques vieilles applications (Notepad) ne reconnaissent pas "
            "Ctrl+Maj+V comme « coller » — à décocher dans ce cas."
        )

        tk.Checkbutton(
            options_frame,
            text="Sons activés",
            variable=self.sound_enabled_var,
            font=("Arial", 8),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X)

        tk.Checkbutton(
            options_frame,
            text="Afficher l'overlay",
            variable=self.overlay_enabled_var,
            font=("Arial", 8),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X)

        tk.Checkbutton(
            options_frame,
            text="Réduire à la fermeture",
            variable=self.minimize_to_tray_var,
            font=("Arial", 8),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X)

        tk.Frame(options_frame, height=1, bg="#cccccc").pack(fill=tk.X, pady=(12, 6))

        tk.Label(
            options_frame,
            text="Session actuelle:",
            font=("Arial", 9, "bold"),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 2))

        self.session_count_label = tk.Label(
            options_frame,
            text="0 transcription",
            font=("Arial", 8),
            fg="gray",
            anchor="w"
        )
        self.session_count_label.pack(anchor=tk.W, fill=tk.X)

        self.session_cost_label = tk.Label(
            options_frame,
            text="$0.0000",
            font=("Arial", 10, "bold"),
            fg="#2e7d32",
            anchor="w"
        )
        self.session_cost_label.pack(anchor=tk.W, fill=tk.X, pady=(0, 5))

        tk.Frame(options_frame, height=1, bg="#cccccc").pack(fill=tk.X, pady=(8, 6))

        month_header = tk.Frame(options_frame)
        month_header.pack(anchor=tk.W, fill=tk.X, pady=(0, 2))

        tk.Label(
            month_header,
            text="Ce mois:",
            font=("Arial", 9, "bold")
        ).pack(side=tk.LEFT)

        month_help = tk.Label(
            month_header,
            text="?",
            font=("Arial", 8, "bold"),
            fg="#1565c0",
            cursor="question_arrow"
        )
        month_help.pack(side=tk.LEFT, padx=(4, 0))
        ToolTip(month_help, self._month_breakdown_text, wraplength=300)

        self.month_count_label = tk.Label(
            options_frame,
            text="0 transcription",
            font=("Arial", 8),
            fg="gray",
            anchor="w"
        )
        self.month_count_label.pack(anchor=tk.W, fill=tk.X)

        self.month_cost_label = tk.Label(
            options_frame,
            text="$0.0000",
            font=("Arial", 10, "bold"),
            fg="#1565c0",
            anchor="w"
        )
        self.month_cost_label.pack(anchor=tk.W, fill=tk.X, pady=(0, 5))

        self.clear_history_button = tk.Button(
            options_frame,
            text="Effacer l'historique des coûts",
            font=("Arial", 8),
            anchor="w",
            relief=tk.FLAT,
            fg="#888888",
            cursor="hand2",
            command=self._clear_cost_history
        )
        self.clear_history_button.pack(anchor=tk.W, fill=tk.X)

        self.root.after(100, self._load_microphones)
        
        text_frame = tk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        tk.Label(
            text_frame,
            text="Transcription:",
            font=("Arial", 10, "bold")
        ).pack(anchor=tk.W)
        
        self.text_area = scrolledtext.ScrolledText(
            text_frame,
            wrap=tk.WORD,
            font=("Arial", 11),
            height=20,
            width=70,
            undo=True,
            maxundo=-1
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # Annulation native de Tk : chaque transcription, effacement ou collage
        # est une étape, Ctrl+Z la défait et Ctrl+Y la refait, sans limite.
        # Liés aussi sur la zone elle-même : son "break" court-circuite la liaison
        # de classe Text, qui ferait sinon une seconde annulation.
        for widget in (self.text_area, self.root):
            for seq in ('<Control-z>', '<Control-Z>'):
                widget.bind(seq, self.undo_text)
            for seq in ('<Control-y>', '<Control-Y>'):
                widget.bind(seq, self.redo_text)
        
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=15, fill=tk.X)
        
        self.clear_button = tk.Button(
            button_frame,
            text="🗑️ Effacer",
            command=self.clear_text,
            width=18,
            height=1,
            bg="#f44336",
            fg="white",
            activebackground="#da190b",
            activeforeground="white",
            font=("Arial", 10)
        )
        self.clear_button.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        self.copy_button = tk.Button(
            button_frame,
            text="📋 Copier tout",
            command=self.copy_to_clipboard,
            width=18,
            height=1,
            font=("Arial", 10)
        )
        self.copy_button.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        status_bar = tk.Frame(main_frame, bg="#f0f0f0", relief=tk.SUNKEN, bd=1)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        self.status_label = tk.Label(
            status_bar,
            text="Prêt",
            font=("Arial", 9),
            fg="gray",
            bg="#f0f0f0",
            anchor="w",
            padx=8,
            pady=4
        )
        self.status_label.pack(fill=tk.X)

        self._update_session_display()
        self._update_api_key_status(ok=(self.client is not None))
        self._on_duration_change(self.max_recording_duration)
        self._on_paste_toggles_changed()

    def _load_microphones(self):
        """Charge la liste des microphones disponibles (filtrés et dédupliqués)"""
        try:
            devices = sd.query_devices()
            
            exclude_keywords = [
                'mixage', 'mix', 'stereo mix', 'stéréo mix',
                'mapper', 'microsoft', 'mappeur',
                'entrée ligne', 'line in', 'line-in',
                'what u hear', 'wave out mix',
                'steam streaming',
                'pilote', 'driver', 'system32', '.sys',
                'input (speaker', 'input (headphone',
                'wave speaker', 'wave microphone headphone'
            ]
            
            input_devices = []
            seen_names = set()
            
            for idx, device in enumerate(devices):
                if device['max_input_channels'] > 0:
                    name = device['name']
                    name_lower = name.lower()
                    
                    if '@' in name or 'system32' in name_lower or '.sys' in name_lower:
                        continue
                    
                    if any(keyword in name_lower for keyword in exclude_keywords):
                        continue

                    if plateforme.mic_hidden(name):
                        continue
                    
                    normalized_name = name.strip()
                    if normalized_name.count('(') > normalized_name.count(')'):
                        normalized_name += ')'
                    
                    if normalized_name not in seen_names:
                        seen_names.add(normalized_name)
                        input_devices.append({
                            'index': idx,
                            'name': normalized_name,
                            'device': device
                        })
            
            if input_devices:
                input_devices.sort(key=lambda x: x['name'])
                
                device_names = [d['name'] for d in input_devices]
                current_mic = self.microphone_name if hasattr(self, 'microphone_name') else device_names[0]
                self.mic_var.set(current_mic)
                
                menu = self.mic_dropdown['menu']
                menu.delete(0, 'end')
                
                for device_info in input_devices:
                    name = device_info['name']
                    real_index = device_info['index']
                    menu.add_command(
                        label=name,
                        command=lambda idx=real_index, n=name: self._change_microphone(idx, n)
                    )
                
                chosen = None
                if self.preferred_mic_name:
                    for device_info in input_devices:
                        if device_info['name'] == self.preferred_mic_name:
                            chosen = device_info
                            break

                if chosen is None:
                    try:
                        default_input_idx = sd.default.device[0]
                        if default_input_idx is not None and default_input_idx >= 0:
                            for device_info in input_devices:
                                if device_info['index'] == default_input_idx:
                                    chosen = device_info
                                    break
                    except Exception:
                        pass

                if chosen is None:
                    chosen = input_devices[0]

                self.selected_device_index = chosen['index']
                self.microphone_name = chosen['name']
                self.mic_var.set(self.microphone_name)
                
                self.mic_display_label.config(text=self.microphone_name)
        except Exception as e:
            print(f"Erreur chargement microphones: {e}")
    
    def _change_microphone(self, device_index, device_name):
        """Change le microphone sélectionné (persisté dans les préférences)"""
        try:
            self.selected_device_index = device_index
            self.microphone_name = device_name
            self.mic_var.set(device_name)
            self.mic_display_label.config(text=device_name)
            self._save_prefs()
            print(f"Microphone changé: {device_name} (index: {device_index})")
        except Exception as e:
            print(f"Erreur changement microphone: {e}")

    def _change_model(self, label, api_name):
        """Change le modèle de transcription (persisté dans les préférences)"""
        self.selected_model = api_name
        self.model_var.set(label)
        self._save_prefs()
        print(f"Modèle changé: {label} ({api_name})")
        if api_name in LOCAL_MODELS and self._local_model_ready(api_name):
            self.local_engine.preload(api_name)

    def _local_model_ready(self, model_id):
        """Le modèle local est-il utilisable ? Sinon l'explique, ou propose de le télécharger."""
        problem = missing_packages()
        if problem:
            messagebox.showerror("Modèle local", problem)
            return False
        if find_model(model_id) is not None:
            return True
        spec = LOCAL_MODELS[model_id]
        if model_id in self._downloads:
            messagebox.showinfo("Modèle local", f"{spec.label} est en cours de téléchargement.")
        elif messagebox.askyesno(
                "Modèle local",
                f"{spec.label} n'est pas encore sur cet ordinateur.\n\n"
                f"Le télécharger maintenant ? Environ {spec.size_text} depuis Hugging Face, "
                f"dans le dossier « models » de l'application."):
            self._start_download(model_id)
        return False

    def _start_download(self, model_id):
        """Télécharge le modèle en arrière-plan, l'avancement dans le statut et l'overlay."""
        label = LOCAL_MODELS[model_id].label
        self._downloads.add(model_id)
        shown = [-1]

        def progress(done, total):
            percent = done * 100 // total if total else 0
            if percent != shown[0]:
                shown[0] = percent
                self.root.after(0, self._show_download_progress, label, percent)

        def run():
            try:
                download_model(model_id, progress)
                error = None
            except Exception as e:
                error = e
            self.root.after(0, self._download_finished, model_id, error)

        threading.Thread(target=run, daemon=True).start()

    def _show_download_progress(self, label, percent):
        text = f"⬇️ {label} : {percent} %"
        busy = self.is_recording or self.is_transcribing
        if not busy:
            self.status_label.config(text=text, fg="blue")
            if hasattr(self, 'overlay'):
                self.overlay.show(text, bg="#1565c0")

    def _download_finished(self, model_id, error):
        self._downloads.discard(model_id)
        label = LOCAL_MODELS[model_id].label
        busy = self.is_recording or self.is_transcribing
        if error is not None:
            if not busy and hasattr(self, 'overlay'):
                self.overlay.hide()
            self.status_label.config(text="❌ Téléchargement échoué", fg="red")
            messagebox.showerror("Téléchargement", f"{label} : téléchargement échoué.\n\n{error}")
            return
        if not busy:
            self.status_label.config(text=f"✅ {label} prêt", fg="green")
            if hasattr(self, 'overlay'):
                self.overlay.show_briefly(f"✅ {label} prêt", bg="#2e7d32", duration_ms=2500)
        if self.selected_model == model_id:
            self.local_engine.preload(model_id)

    def _deferred_init(self):
        """Initialisations lourdes différées pour accélérer l'apparition de la fenêtre"""
        try:
            pygame.mixer.init()
        except Exception as e:
            print(f"⚠️ Init pygame mixer échouée: {e}")

        problem = plateforme.environment_problem()
        if problem:
            print(f"⚠️ {problem}")
            messagebox.showwarning("Transcription Vocale", problem)

        self.setup_global_hotkey()
        self._setup_tray()
        # Lancée à l'ouverture de session (--barre), l'application démarre rangée dans la
        # barre système ; sans icône, la fenêtre reste le seul moyen de la retrouver.
        if START_IN_TRAY and self.tray_icon is None:
            self._restore_window()

        if self.selected_model in LOCAL_MODELS and missing_requirement(self.selected_model) is None:
            self.local_engine.preload(self.selected_model)

    def _make_tray_icon_image(self):
        """Génère programmatiquement l'icône de la tray (évite de shipper un binaire).

        Micro sur fond vert quand le raccourci est actif ; sur fond gris et barré quand il
        est en pause : l'état se lit d'un coup d'œil dans la barre système."""
        img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        background = (120, 120, 120, 255) if self.hotkey_paused else (46, 160, 67, 255)
        d.ellipse((4, 4, 60, 60), fill=background)
        try:
            d.rounded_rectangle((24, 16, 40, 40), radius=8, fill=(255, 255, 255, 255))
        except AttributeError:
            d.rectangle((24, 16, 40, 40), fill=(255, 255, 255, 255))
        d.rectangle((30, 40, 34, 48), fill=(255, 255, 255, 255))
        d.rectangle((22, 48, 42, 52), fill=(255, 255, 255, 255))
        if self.hotkey_paused:
            d.line((14, 50, 50, 14), fill=(120, 120, 120, 255), width=10)
            d.line((14, 50, 50, 14), fill=(255, 255, 255, 255), width=5)
        return img

    def _tray_title(self):
        if self.hotkey_paused:
            return "Transcription Vocale — dictée en pause (%s pour la réactiver)" % (
                describe_hotkey(self.pause_hotkey))
        return "Transcription Vocale — dictée active (%s : pause)" % (
            describe_hotkey(self.pause_hotkey))

    def _update_tray(self):
        """Reporte l'état de la pause sur l'icône, son infobulle et son menu."""
        if self.tray_icon is None:
            return
        try:
            self.tray_icon.icon = self._make_tray_icon_image()
            self.tray_icon.title = self._tray_title()
            self.tray_icon.update_menu()
        except Exception as e:
            print(f"⚠️ Icône système non mise à jour : {e}")

    def _setup_tray(self):
        """Crée l'icône de la barre système avec menu clic-droit (Afficher/Quitter).
        Si pystray/Pillow ne sont pas installés, la tray est simplement désactivée."""
        if not TRAY_AVAILABLE:
            print("ℹ️ pystray/Pillow non installés — fonctionnement sans tray.")
            return

        try:
            icon_image = self._make_tray_icon_image()

            def on_show(icon, item):
                self.root.after(0, self._restore_window)

            def on_quit(icon, item):
                self.root.after(0, self._quit_app)

            def on_pause(icon, item):
                self.root.after(0, self.toggle_hotkey_pause)

            menu = pystray.Menu(
                pystray.MenuItem("Afficher", on_show, default=True),
                pystray.MenuItem(
                    lambda item: ("Réactiver la dictée" if self.hotkey_paused
                                  else "Mettre la dictée en pause"),
                    on_pause),
                pystray.MenuItem("Quitter", on_quit),
            )

            self.tray_icon = pystray.Icon(
                "whisper-voice",
                icon_image,
                self._tray_title(),
                menu
            )

            def run_tray():
                # Sans icône, la croix doit quitter : sinon la fenêtre serait cachée sans retour.
                try:
                    self.tray_icon.run()
                except Exception as e:
                    print(f"⚠️ Icône système arrêtée : {e}")
                    self.tray_icon = None

            self.tray_thread = threading.Thread(target=run_tray, daemon=True)
            self.tray_thread.start()
            print("✅ Icône système (tray) activée")
        except Exception as e:
            print(f"⚠️ Impossible d'initialiser la tray: {e}")
            self.tray_icon = None

    def _restore_window(self):
        """Ré-affiche la fenêtre principale depuis la tray."""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _hide_to_tray(self):
        """Appelé sur clic X. Respecte la préférence utilisateur :
        - si tray disponible + option activée → masque dans la tray
        - sinon → quitte réellement l'application."""
        wants_hide = (hasattr(self, 'minimize_to_tray_var')
                      and self.minimize_to_tray_var.get())
        if self.tray_icon is not None and wants_hide:
            self.root.withdraw()
        else:
            self._quit_app()

    def _is_terminal_window_focused(self):
        """La fenêtre au premier plan est-elle un terminal connu ? En cas d'erreur, non."""
        try:
            return plateforme.terminal_focused()
        except Exception as e:
            print(f"⚠️ Fenêtre active illisible : {e}")
            return False

    def _on_overlay_toggle(self):
        """Réaction au changement du toggle overlay : applique + sauvegarde."""
        if hasattr(self, 'overlay'):
            self.overlay.enabled = self.overlay_enabled_var.get()
            if not self.overlay.enabled:
                self.overlay.hide()
        self._save_prefs()

    def _on_overlay_position_saved(self, x, y):
        """Appelé par l'overlay après un drag utilisateur : persiste la position."""
        self._save_prefs()

    def _on_paste_toggles_changed(self):
        """Garantit la cohérence entre auto_paste et terminal_paste :
        - terminal_paste activé → auto_paste forcé à True
        - auto_paste désactivé  → terminal_paste forcé à False
        Met à jour l'état visuel (disabled) des deux cases pour rendre
        la contrainte lisible à l'utilisateur."""
        if self.terminal_paste_var.get() and not self.auto_paste_var.get():
            self.auto_paste_var.set(True)
            return  # la trace auto_paste va re-déclencher la méthode
        if not self.auto_paste_var.get() and self.terminal_paste_var.get():
            self.terminal_paste_var.set(False)
            return

        if hasattr(self, 'auto_paste_checkbox') and hasattr(self, 'terminal_paste_checkbox'):
            auto_paste_on = self.auto_paste_var.get()
            terminal_on = self.terminal_paste_var.get()
            self.auto_paste_checkbox.config(state='disabled' if terminal_on else 'normal')
            self.terminal_paste_checkbox.config(state='normal' if auto_paste_on else 'disabled')

        self._save_prefs()

    def _quit_app(self):
        """Quitte proprement : sauvegarde, arrêt de la tray, destruction de la fenêtre.

        L'historique n'est pas réécrit ici, il est déjà sauvé à chaque transcription.
        """
        try:
            self._save_prefs()
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde à la fermeture: {e}")

        if self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
            self.tray_icon = None

        try:
            self.root.destroy()
        except Exception:
            pass

    def _update_api_key_status(self, ok):
        """Affiche l'état de la clé API dans le panneau d'options"""
        if not hasattr(self, 'api_key_status_label'):
            return
        if ok:
            self.api_key_status_label.config(text="✓ Configurée", fg="#2e7d32")
        else:
            self.api_key_status_label.config(text="✗ Non configurée", fg="#c62828")

    def _open_api_key_dialog(self):
        """Ouvre une boîte de dialogue pour saisir/remplacer la clé API OpenAI.
        La clé saisie est masquée (affichée en puces) et enregistrée dans .env."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Clé API OpenAI")
        dialog.geometry("440x200")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(
            dialog,
            text="Collez votre clé API OpenAI :",
            font=("Arial", 10, "bold")
        ).pack(pady=(15, 5))

        entry = tk.Entry(dialog, show="•", font=("Consolas", 10), width=48)
        entry.pack(pady=5, padx=20)
        entry.focus_set()

        tk.Label(
            dialog,
            text="La clé est enregistrée dans le fichier .env local.\n"
                 "Elle n'est jamais affichée en clair ni transmise ailleurs.\n"
                 "Obtenez une clé sur platform.openai.com/api-keys",
            font=("Arial", 8),
            fg="gray",
            justify=tk.CENTER
        ).pack(pady=(5, 5))

        def save():
            new_key = entry.get().strip()
            if not new_key:
                return
            if not new_key.startswith("sk-"):
                messagebox.showerror(
                    "Clé invalide",
                    "La clé API OpenAI commence normalement par « sk- »."
                )
                return
            if self._save_api_key(new_key):
                dialog.destroy()

        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        tk.Button(
            btn_frame, text="Enregistrer", command=save,
            bg="#4CAF50", fg="white", activebackground="#45a049",
            width=14
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            btn_frame, text="Annuler", command=dialog.destroy, width=14
        ).pack(side=tk.LEFT, padx=5)

        dialog.bind('<Return>', lambda e: save())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def _save_api_key(self, new_key):
        """Valide la clé auprès d'OpenAI, puis (seulement en cas de succès) l'enregistre
        dans .env et recharge le client. Évite de polluer .env avec une clé invalide."""
        try:
            temp_client = OpenAI(api_key=new_key)
            temp_client.models.list()
        except Exception as e:
            self._update_api_key_status(ok=(self.client is not None))
            messagebox.showerror(
                "Clé API rejetée",
                f"OpenAI a rejeté cette clé :\n{e}"
            )
            return False

        try:
            lines = []
            if ENV_FILE.exists():
                with open(ENV_FILE, 'r', encoding='utf-8') as f:
                    lines = f.readlines()

            found = False
            new_lines = []
            for line in lines:
                if line.strip().startswith('OPENAI_API_KEY='):
                    new_lines.append(f'OPENAI_API_KEY={new_key}\n')
                    found = True
                else:
                    new_lines.append(line)
            if not found:
                new_lines.append(f'OPENAI_API_KEY={new_key}\n')

            with open(ENV_FILE, 'w', encoding='utf-8') as f:
                f.writelines(new_lines)

            os.environ['OPENAI_API_KEY'] = new_key
            self.client = temp_client
            self._update_api_key_status(ok=True)
            messagebox.showinfo("Clé API", "✅ Clé API enregistrée et validée.")
            return True
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer la clé : {e}")
            return False

    def _load_prefs(self):
        """Charge les préférences utilisateur (merge avec défauts pour compat ascendante)"""
        try:
            if PREFS_FILE.exists():
                with open(PREFS_FILE, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        return {**DEFAULT_PREFS, **loaded}
        except Exception as e:
            print(f"⚠️ Erreur lecture préférences: {e}")
        return dict(DEFAULT_PREFS)

    def _save_prefs(self):
        """Sauvegarde les préférences actuelles sur disque"""
        try:
            prefs = {
                "selected_model": self.selected_model,
                "selected_device_name": (self.microphone_name
                                         if self.microphone_name not in (None, "Chargement...")
                                         else None),
                "max_recording_duration": self.max_recording_duration,
                "auto_copy": self.auto_copy_var.get() if hasattr(self, 'auto_copy_var') else True,
                "auto_paste": self.auto_paste_var.get() if hasattr(self, 'auto_paste_var') else True,
                "sound_enabled": self.sound_enabled_var.get() if hasattr(self, 'sound_enabled_var') else True,
                "overlay_enabled": self.overlay_enabled_var.get() if hasattr(self, 'overlay_enabled_var') else True,
                "minimize_to_tray_on_close": self.minimize_to_tray_var.get() if hasattr(self, 'minimize_to_tray_var') else True,
                "terminal_paste": self.terminal_paste_var.get() if hasattr(self, 'terminal_paste_var') else False,
                "overlay_position": (list(self.overlay.custom_position)
                                     if hasattr(self, 'overlay') and self.overlay.custom_position
                                     else None),
                "hotkey": self.hotkey,
                "pause_hotkey": self.pause_hotkey,
                "hotkey_paused": self.hotkey_paused,
            }
            with open(PREFS_FILE, 'w', encoding='utf-8') as f:
                json.dump(prefs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde préférences: {e}")

    def _load_history(self):
        """Charge l'historique persistant depuis le fichier JSON"""
        try:
            if HISTORY_FILE.exists():
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, dict) and 'transcriptions' in data:
                        return data
        except Exception as e:
            print(f"⚠️ Erreur lecture historique: {e}")
        return {"transcriptions": []}

    def _save_history(self):
        """Sauvegarde l'historique sur disque (appelé après chaque transcription et à la fermeture)"""
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde historique: {e}")

    def _log_transcription(self, model, duration_sec, cost_usd, cost_exact=True):
        """Enregistre une transcription dans l'historique persistant"""
        entry = {
            "at": datetime.now().isoformat(timespec='seconds'),
            "model": model,
            "duration_sec": round(duration_sec, 2),
            "cost_usd": round(cost_usd, 6),
            "cost_exact": cost_exact,
        }
        self.history.setdefault("transcriptions", []).append(entry)
        self._save_history()

    def _compute_month_stats(self):
        """Calcule le nombre de transcriptions et le coût total du mois en cours"""
        now = datetime.now()
        month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        count = 0
        total = 0.0
        for entry in self.history.get("transcriptions", []):
            try:
                at = datetime.fromisoformat(entry["at"])
            except (KeyError, ValueError):
                continue
            if at >= month_start:
                count += 1
                total += entry.get("cost_usd", 0.0)
        return count, total

    def _month_breakdown_text(self):
        """Détail du mois en cours par modèle, pour la tooltip du '?'."""
        now = datetime.now()
        month_start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)

        par_modele = {}
        for entry in self.history.get("transcriptions", []):
            try:
                at = datetime.fromisoformat(entry["at"])
            except (KeyError, ValueError):
                continue
            if at < month_start:
                continue
            stats = par_modele.setdefault(entry["model"], [0, 0.0, 0.0])
            stats[0] += 1
            stats[1] += entry.get("duration_sec", 0.0)
            stats[2] += entry.get("cost_usd", 0.0)

        if not par_modele:
            return "Aucune transcription ce mois-ci."

        lignes = [f"Détail depuis le {month_start.strftime('%d/%m')} :", ""]
        for model, (count, secondes, cout) in sorted(par_modele.items(),
                                                     key=lambda kv: -kv[1][2]):
            label = next((l for l, api in MODEL_OPTIONS if api == model), model)
            label = label.replace(" (recommandé)", "")
            minutes = secondes / 60
            lignes.append(f"{label}\n   {count} transcription{'s' if count > 1 else ''}"
                          f" · {minutes:.1f} min · ${cout:.4f}")
        return "\n".join(lignes)

    def _clear_cost_history(self):
        """Vide l'historique des coûts, après confirmation."""
        count = len(self.history.get("transcriptions", []))
        if count == 0:
            messagebox.showinfo("Historique", "L'historique est déjà vide.")
            return

        if not messagebox.askyesno(
            "Effacer l'historique",
            f"Effacer les {count} entrées du suivi des coûts ?\n\n"
            "Les compteurs de session et du mois repartent de zéro. "
            "Cette action est irréversible."
        ):
            return

        self.history["transcriptions"] = []
        self._save_history()
        self.session_cost = 0.0
        self.session_transcriptions = 0
        self._update_session_display()

    def _update_session_display(self):
        """Met à jour l'affichage des compteurs de session et du mois en cours"""
        count = self.session_transcriptions
        suffix = "s" if count > 1 else ""
        self.session_count_label.config(text=f"{count} transcription{suffix}")
        self.session_cost_label.config(text=f"${self.session_cost:.4f}")

        month_count, month_cost = self._compute_month_stats()
        month_suffix = "s" if month_count > 1 else ""
        self.month_count_label.config(text=f"{month_count} transcription{month_suffix}")
        self.month_cost_label.config(text=f"${month_cost:.4f}")

    def _update_recording_timer(self):
        """Met à jour le timer d'enregistrement (barre de statut + overlay flottant).
        Tick toutes les 500ms tant que l'enregistrement est actif."""
        if not self.is_recording or self.recording_start_time is None:
            return
        elapsed = time.time() - self.recording_start_time
        max_sec = self.max_recording_duration

        def fmt(secs):
            m = int(secs // 60)
            s = int(secs % 60)
            return f"{m:02d}:{s:02d}"

        timer_text = f"{fmt(elapsed)} / {fmt(max_sec)}"
        self.status_label.config(
            text=f"🔴 Enregistrement  {timer_text}",
            fg="red"
        )
        if hasattr(self, 'overlay'):
            self.overlay.show(f"🔴 {timer_text}", bg="#c62828")

        self.root.after(500, self._update_recording_timer)
    
    def _on_duration_change(self, value):
        """Appelé quand le curseur change"""
        raw_value = int(float(value))
        rounded_value = round(raw_value / 30) * 30
        
        if rounded_value != self.max_recording_duration:
            self.max_recording_duration = rounded_value
            self.duration_var.set(rounded_value)
            self._save_prefs()
        
        if self.max_recording_duration >= 60:
            minutes = self.max_recording_duration // 60
            seconds = self.max_recording_duration % 60
            if seconds == 0:
                display = f"{minutes} min"
            else:
                display = f"{minutes} min {seconds}s"
        else:
            display = f"{self.max_recording_duration}s"
        
        self.duration_label.config(text=display)
    
    def toggle_recording(self):
        """Bascule entre démarrer et arrêter l'enregistrement"""
        if self.is_transcribing:
            self.play_sound('error.wav', volume=0.3)
            return

        if not self.is_recording:
            if self.selected_model in LOCAL_MODELS:
                if not self._local_model_ready(self.selected_model):
                    return
            elif self.client is None:
                messagebox.showwarning(
                    "Clé API manquante",
                    "Configurez votre clé API OpenAI dans les options "
                    "(section « Clé API OpenAI » → bouton « Modifier la clé »)."
                )
                return
            self.start_recording()
        else:
            self.stop_recording()
    
    def play_sound(self, filename, volume=0.5):
        """Joue un son (WAV) avec volume réglable. Respecte le toggle 'Sons activés'."""
        # Toggle global — la variable peut ne pas encore exister pendant l'init UI
        if hasattr(self, 'sound_enabled_var') and not self.sound_enabled_var.get():
            return
        try:
            sound_path = os.path.join(self.sounds_dir, filename)
            if not os.path.exists(sound_path):
                sound_path = sound_path.replace('.mp3', '.wav')

            if os.path.exists(sound_path):
                sound = pygame.mixer.Sound(sound_path)
                sound.set_volume(volume)
                sound.play()
        except Exception as e:
            try:
                import winsound
                winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
            except:
                print(f"Erreur lecture son: {e}")
    
    def start_recording(self):
        """Démarre l'enregistrement audio"""
        if self.is_transcribing:
            self.play_sound('error.wav', volume=0.3)
            return
        
        self.is_recording = True
        self.cancel_requested = False
        self.audio_frames = []
        self.recording_start_time = time.time()
        # Le modèle local transcrit pendant l'enregistrement : il reçoit le son au fil de l'eau.
        self.recording_model = self.selected_model
        if self.recording_model in LOCAL_MODELS:
            self.local_engine.begin(self.recording_model, self.RATE)

        self.play_sound('start.wav')
        
        self.record_button.config(
            text="⏹️ Arrêter",
            bg="#f44336",
            activebackground="#da190b"
        )
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        self.status_label.config(text="🔴 Enregistrement en cours...", fg="red")
        
        self.recording_thread = threading.Thread(target=self._record_audio, daemon=True)
        self.recording_thread.start()

        self.root.after(100, self._update_recording_timer)
    
    def cancel_recording(self):
        """Annule l'enregistrement sans faire l'appel API"""
        if self.is_recording:
            self.is_recording = False
            self.cancel_requested = True
            self.audio_frames = []
            if self.recording_model in LOCAL_MODELS:
                self.local_engine.cancel()
            self.play_sound('error.wav', volume=0.3)
            self._reset_ui("Enregistrement annulé")
            self.status_label.config(text="❌ Enregistrement annulé", fg="orange")
            self.root.after(2000, lambda: self.status_label.config(text="Prêt", fg="gray"))
            if hasattr(self, 'overlay'):
                self.overlay.show_briefly("❌ Annulé", bg="#ff9800", duration_ms=1500)
    
    def stop_recording(self):
        """Arrête l'enregistrement et transcrit"""
        if not self.is_recording:
            return
        
        self.is_recording = False
        
        self.cancel_button.pack_forget()
        
        self.play_sound('recording_stopped.wav', volume=0.2)

        self.status_label.config(text="⏳ Traitement en cours...", fg="orange")
        self.record_button.config(state=tk.DISABLED)
        if hasattr(self, 'overlay'):
            self.overlay.show("⏳ Transcription…", bg="#1565c0")
        
        threading.Thread(target=self._process_recording, daemon=True).start()
    
    def _record_audio(self):
        """Enregistre l'audio en continu via un InputStream sounddevice.

        Le callback est appelé automatiquement par PortAudio sur chaque bloc
        capturé, sans gap entre les blocs — contrairement à la boucle précédente
        `sd.rec() + sd.wait()` qui laissait passer ~10-20ms entre chaque chunk.
        """
        device = self.selected_device_index if self.selected_device_index is not None else None
        local = self.recording_model in LOCAL_MODELS

        def audio_callback(indata, frames, time_info, status):
            # `indata` est un buffer numpy réutilisé — on copie avant de stocker.
            if self.is_recording and not self.cancel_requested:
                chunk = indata.copy().tobytes()
                self.audio_frames.append(chunk)
                if local:
                    self.local_engine.feed(chunk)

        try:
            # blocksize=0 laisse PortAudio choisir la taille optimale pour le périphérique.
            with sd.InputStream(
                samplerate=self.RATE,
                channels=self.CHANNELS,
                dtype=self.DTYPE,
                device=device,
                callback=audio_callback,
                blocksize=0,
            ):
                while self.is_recording and not self.cancel_requested:
                    if self.recording_start_time and (time.time() - self.recording_start_time) >= self.max_recording_duration:
                        self.root.after(0, self._auto_stop_recording)
                        break
                    time.sleep(0.05)
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: messagebox.showerror("Erreur", f"Erreur d'enregistrement: {err}"))
            self.is_recording = False
    
    def _auto_stop_recording(self):
        """Arrête automatiquement l'enregistrement après 1 minute"""
        if self.is_recording:
            self.stop_recording()
    
    def _process_recording(self):
        """Traite l'enregistrement et transcrit"""
        if self.cancel_requested:
            self.is_transcribing = False
            return
        
        self.is_transcribing = True
        
        try:
            model_used = self.recording_model
            if not self.audio_frames or self.cancel_requested:
                self.is_transcribing = False
                if model_used in LOCAL_MODELS:
                    self.local_engine.cancel()
                if not self.cancel_requested:
                    self.root.after(0, lambda: self._reset_ui("Aucun audio enregistré"))
                return

            if model_used in LOCAL_MODELS:
                self._finish_local_transcription()
                return

            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.wav')
            temp_filename = temp_file.name
            temp_file.close()
            
            wf = wave.open(temp_filename, 'wb')
            wf.setnchannels(self.CHANNELS)
            wf.setsampwidth(2)  # 16-bit = 2 bytes
            wf.setframerate(self.RATE)
            wf.writeframes(b''.join(self.audio_frames))
            wf.close()
            
            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return

            self.root.after(0, lambda: self.status_label.config(text="🔄 Transcription en cours...", fg="blue"))
            
            # Une fois l'appel API lancé, il n'est plus annulable.
            with open(temp_filename, 'rb') as audio_file:
                transcript = self.client.audio.transcriptions.create(
                    model=model_used,
                    file=audio_file,
                    language="fr"
                )

            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return

            text = transcript.text

            os.unlink(temp_filename)

            # Durée calculée depuis les octets réels (InputStream a une taille de
            # bloc variable) : octets / 2 (int16) / taux d'échantillonnage.
            total_bytes = sum(len(f) for f in self.audio_frames)
            audio_duration_seconds = total_bytes / 2 / self.RATE
            cost, exact = compute_transcription_cost(
                model_used, getattr(transcript, "usage", None), audio_duration_seconds
            )
            self.session_cost += cost
            self.session_transcriptions += 1
            self._log_transcription(model_used, audio_duration_seconds, cost, exact)
            self.root.after(0, self._update_session_display)

            self.root.after(0, lambda: self._display_text(text))
            
        except Exception as e:
            self.is_transcribing = False
            error_msg = f"Erreur lors de la transcription: {e}"
            self.root.after(0, lambda: messagebox.showerror("Erreur", error_msg))
            self.root.after(0, lambda: self._reset_ui("Erreur"))
    
    def _finish_local_transcription(self):
        """Récupère le texte du modèle local, qui a transcrit pendant l'enregistrement."""
        self.root.after(0, lambda: self.status_label.config(text="🔄 Transcription en cours...", fg="blue"))
        text = self.local_engine.finish()
        audio_duration_seconds = sum(len(f) for f in self.audio_frames) / 2 / self.RATE
        self.session_transcriptions += 1
        self._log_transcription(self.recording_model, audio_duration_seconds, 0.0, True)
        self.root.after(0, self._update_session_display)
        self.root.after(0, lambda: self._display_text(text))

    def _display_text(self, text):
        """Affiche le texte transcrit et gère copie/collage selon les toggles utilisateur.

        Les deux toggles sont indépendants :
        - auto_copy seul        → texte dans le presse-papier, pas d'injection
        - auto_paste seul       → injection dans le champ actif, presse-papier restauré
        - les deux              → texte collé ET conservé dans le presse-papier
        - aucun                 → texte uniquement visible dans la zone de transcription
        """
        self.text_area.edit_separator()
        self.text_area.insert(tk.END, text + "\n\n")
        self.text_area.edit_separator()
        self.text_area.see(tk.END)

        if self.auto_paste_var.get():
            self._paste_when_released(text, time.monotonic() + PASTE_RELEASE_TIMEOUT)
            return

        status_text = "✅ Transcription terminée"
        status_color = "green"
        if self.auto_copy_var.get():
            try:
                pyperclip.copy(text)
                status_text = "✅ Copié dans le presse-papier"
            except Exception:
                status_text = "✅ Transcription terminée (copie échouée)"
                status_color = "orange"
        self._finish_transcription(status_text, status_color)

    def _paste_keys_held(self):
        """Un modificateur ou une touche du raccourci est-il encore enfoncé ?"""
        try:
            return plateforme.keys_held(self.hotkey['vks'])
        except Exception as e:
            print(f"⚠️ État du clavier illisible : {e}")
            return False

    def _paste_when_released(self, text, deadline):
        """Colle dès que le clavier est libre, sans figer l'interface pendant l'attente."""
        if not self._paste_keys_held():
            self._paste_text(text)
            return
        if time.monotonic() < deadline:
            self.root.after(PASTE_KEY_POLL_MS, self._paste_when_released, text, deadline)
            return
        try:
            pyperclip.copy(text)
            status_text = "⚠️ Touches encore enfoncées : Ctrl+V pour coller"
        except Exception:
            status_text = "⚠️ Touches encore enfoncées : collage annulé"
        self._finish_transcription(status_text, "orange")

    def _paste_text(self, text):
        """Injecte le texte dans le champ actif, puis rend l'ancien presse-papier sauf si auto_copy."""
        auto_copy = self.auto_copy_var.get()
        saved_clipboard = None
        snapshot_error = None
        if not auto_copy:
            try:
                saved_clipboard = plateforme.clipboard_snapshot()
            except Exception as e:
                snapshot_error = e

        try:
            pyperclip.copy(text)
            pasted_sequence = plateforme.clipboard_sequence()
            # Ctrl+Maj+V si le toggle manuel est actif (terminaux intégrés type
            # Cursor) ou si la fenêtre active est un terminal connu.
            plateforme.send_paste(self.terminal_paste_var.get()
                                  or self._is_terminal_window_focused())
            if snapshot_error is not None:
                print(f"⚠️ Presse-papier non sauvegardé : {snapshot_error}")
                status_text, status_color = "⚠️ Collé, ancien presse-papier perdu", "orange"
            elif auto_copy:
                status_text, status_color = "✅ Collé et copié", "green"
            else:
                status_text, status_color = "✅ Collé (presse-papier préservé)", "green"
        except Exception as e:
            print(f"⚠️ Collage échoué : {e}")
            status_text = "✅ Transcription terminée (collage échoué)"
            status_color = "orange"
            pasted_sequence = None

        if not auto_copy and snapshot_error is None:
            self.root.after(CLIPBOARD_RESTORE_DELAY_MS,
                            self._restore_clipboard, pasted_sequence, saved_clipboard)
        self._finish_transcription(status_text, status_color)

    def _restore_clipboard(self, pasted_sequence, saved):
        """Remet l'ancien presse-papier, sauf si autre chose y a été copié entre-temps."""
        try:
            if pasted_sequence is None or plateforme.clipboard_sequence() == pasted_sequence:
                plateforme.clipboard_restore(saved)
        except Exception as e:
            print(f"⚠️ Presse-papier non rendu : {e}")
            if hasattr(self, 'overlay'):
                self.overlay.show_briefly("⚠️ Ancien presse-papier perdu", bg="#ff9800",
                                          duration_ms=1800)

    def _finish_transcription(self, status_text, status_color):
        self.play_sound('end.wav')
        self._reset_ui(status_text)
        self.status_label.config(text=status_text, fg=status_color)

        if hasattr(self, 'overlay'):
            overlay_bg = "#2e7d32" if status_color == "green" else "#ff9800"
            self.overlay.show_briefly(status_text, bg=overlay_bg, duration_ms=1800)

        self.is_transcribing = False
    
    def _reset_ui(self, status_text="Prêt"):
        """Réinitialise l'interface après l'enregistrement"""
        self.record_button.config(
            text="🎤 Démarrer l'enregistrement",
            bg="#4CAF50",
            activebackground="#45a049",
            state=tk.NORMAL
        )
        self.cancel_button.pack_forget()
        
        self.status_label.config(text=status_text, fg="gray")
        
        self.audio_frames = []
        self.recording_start_time = None
        self.cancel_requested = False
    
    def clear_text(self):
        """Efface le contenu de la zone de texte (Ctrl+Z le rend)"""
        self.text_area.edit_separator()
        self.text_area.delete(1.0, tk.END)
        self.text_area.edit_separator()
        self.status_label.config(text="✅ Texte effacé (Ctrl+Z pour annuler)", fg="gray")
        self.root.after(3000, lambda: self.status_label.config(text="Prêt", fg="gray"))

    def undo_text(self, event=None):
        """Défait la dernière modification de la zone de texte (Ctrl+Z)"""
        self._step_text(self.text_area.edit_undo, "↩️ Annulé", "Rien à annuler")
        return "break"

    def redo_text(self, event=None):
        """Refait la dernière modification annulée (Ctrl+Y)"""
        self._step_text(self.text_area.edit_redo, "↪️ Rétabli", "Rien à rétablir")
        return "break"

    def _step_text(self, action, done_text, empty_text):
        try:
            action()
            self.status_label.config(text=done_text, fg="green")
        except tk.TclError:
            self.status_label.config(text=empty_text, fg="gray")
        self.root.after(2000, lambda: self.status_label.config(text="Prêt", fg="gray"))
    
    def copy_to_clipboard(self):
        """Copie le contenu de la zone de texte dans le presse-papier"""
        text = self.text_area.get(1.0, tk.END).strip()
        if text:
            try:
                pyperclip.copy(text)
                self.status_label.config(text="✅ Texte copié !", fg="green")
                self.root.after(2000, lambda: self.status_label.config(text="Prêt", fg="gray"))
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de copier: {e}")
        else:
            messagebox.showinfo("Info", "Aucun texte à copier")
    
    def _hotkey_banner(self):
        """Texte du bandeau sous les boutons, reflétant le raccourci courant et la pause."""
        if self.hotkey_paused:
            return "⏸️ %s en pause — %s pour le réactiver" % (
                describe_hotkey(self.hotkey), describe_hotkey(self.pause_hotkey))
        return "⌨️ %s (global) — Échap pour annuler — %s : pause" % (
            describe_hotkey(self.hotkey), describe_hotkey(self.pause_hotkey))

    def _pause_description(self):
        state = "en pause" if self.hotkey_paused else "actif"
        return "%s (raccourci %s)" % (describe_hotkey(self.pause_hotkey), state)

    @staticmethod
    def _keyboard_only(hotkey):
        """Le raccourci de pause est une combinaison du clavier."""
        if hotkey.get('button'):
            return normalize_hotkey(None, plateforme.DEFAULT_PAUSE_HOTKEY)
        return hotkey

    def _refresh_hotkey_labels(self):
        if hasattr(self, 'hotkey_label'):
            self.hotkey_label.config(text=self._hotkey_banner())
        if hasattr(self, 'hotkey_pref_label'):
            self.hotkey_pref_label.config(text=describe_hotkey(self.hotkey))
        if hasattr(self, 'pause_pref_label'):
            self.pause_pref_label.config(text=self._pause_description())

    def _apply_hotkey(self, hotkey, target='hotkey'):
        """Enregistre un nouveau raccourci (de dictée ou de pause) et rafraîchit l'affichage."""
        if target == 'pause':
            self.pause_hotkey = hotkey
        else:
            self.hotkey = hotkey
        self._save_prefs()
        self._refresh_hotkey_labels()
        self._update_tray()
        self._grab_hotkey(notify=True)

    def _grab_hotkey(self, notify=False):
        """Réserve les raccourcis au système (Linux) : les touches ne s'écrivent plus ailleurs.

        En pause, seul le raccourci de pause reste réservé : celui de dictée retrouve son
        effet habituel (bouton de souris dans un jeu, par exemple)."""
        hotkeys = [dict(self.pause_hotkey, exact=True)]
        if not self.hotkey_paused:
            hotkeys.insert(0, self.hotkey)
        problem = plateforme.grab_hotkeys(hotkeys, self._on_hotkey_button)
        if problem:
            print(f"⚠️ {problem}")
            if notify:
                messagebox.showwarning("Raccourci", problem)

    def toggle_hotkey_pause(self):
        """Met le raccourci de dictée en pause, ou l'en sort."""
        if self.is_recording:
            if hasattr(self, 'overlay'):
                self.overlay.show_briefly("⏸️ Pause impossible pendant la dictée",
                                          bg="#ff9800", duration_ms=1500)
            return
        self.hotkey_paused = not self.hotkey_paused
        self._hotkey_down = False
        self._save_prefs()
        self._refresh_hotkey_labels()
        self._update_tray()
        self._grab_hotkey()
        print("⏸️ Raccourci en pause" if self.hotkey_paused else "▶️ Raccourci réactivé")
        if hasattr(self, 'overlay'):
            if self.hotkey_paused:
                self.overlay.show_briefly("⏸️ Dictée en pause", bg="#616161", duration_ms=1500)
            else:
                self.overlay.show_briefly("▶️ Dictée réactivée", bg="#2e7d32", duration_ms=1500)

    def _pause_matches(self, vk, held):
        """Le raccourci de pause, avec exactement ses modificateurs ?"""
        return vk in self.pause_hotkey['vks'] and held == set(self.pause_hotkey['modifiers'])

    def _on_hotkey_button(self, button, pressed):
        """Clic du bouton de souris réservé (appelé hors du fil de Tk)."""
        if self.hotkey_paused or button != self.hotkey.get('button'):
            return
        if not pressed:
            self._hotkey_down = False
        elif self._hotkey_capture is None and self._accept_hotkey_press():
            self.root.after(0, self.toggle_recording)

    def _accept_hotkey_press(self):
        """Filtre les répétitions du système : seule une transition relâché → enfoncé bascule."""
        now = time.monotonic()
        is_repeat = self._hotkey_down and (now - self._hotkey_last_press) < HOTKEY_REPEAT_GAP
        self._hotkey_last_press = now
        self._hotkey_down = True

        if is_repeat:
            return False
        if now - self._hotkey_last_toggle < HOTKEY_MIN_INTERVAL:
            return False

        self._hotkey_last_toggle = now
        return True

    def _hotkey_matches(self, vk, held):
        """Le raccourci est-il déclenché par `vk` avec les modificateurs `held` ?

        Les modificateurs requis doivent être enfoncés sans que les autres soient exclus,
        certaines macros ajoutant Maj. Sans modificateur requis, tout modificateur enfoncé
        disqualifie la frappe.
        """
        if vk not in self.hotkey['vks']:
            return False
        required = set(self.hotkey['modifiers'])
        return required <= held if required else not held

    def setup_global_hotkey(self):
        """Écoute le clavier globalement : raccourci d'enregistrement + Échap.

        Suivi manuel des modificateurs plutôt que `GlobalHotKeys`, sensible à l'ordre et
        au rythme d'arrivée des touches qu'une macro peut émettre en quelques millisecondes.
        """
        held = set()

        def on_press(key, vk):
            try:
                modifier = MODIFIER_NAMES.get(key)
                if modifier:
                    held.add(modifier)
                    if self._hotkey_capture is not None:
                        self.root.after(0, self._on_capture_modifiers, frozenset(held))
                    return

                # Neutralise le raccourci courant pendant l'essai.
                if self._hotkey_capture is not None:
                    if key == pynput_keyboard.Key.esc:
                        self.root.after(0, self._close_hotkey_dialog)
                    else:
                        if vk is not None:
                            self.root.after(0, self._on_capture_key, frozenset(held), vk)
                    return

                if key == pynput_keyboard.Key.esc:
                    if self.is_recording:
                        self.root.after(0, self.cancel_recording)
                    return

                if vk is not None and self._pause_matches(vk, held):
                    if not self._pause_down:
                        self._pause_down = True
                        self.root.after(0, self.toggle_hotkey_pause)
                    return

                if self.hotkey_paused:
                    return

                if vk is not None and self._hotkey_matches(vk, held):
                    if self._accept_hotkey_press():
                        self.root.after(0, self.toggle_recording)
            except Exception:
                pass

        def on_release(key, vk):
            try:
                modifier = MODIFIER_NAMES.get(key)
                if modifier:
                    held.discard(modifier)
                    if self._hotkey_capture is not None:
                        self.root.after(0, self._on_capture_modifiers, frozenset(held))
                    return

                if vk is not None and vk in self.pause_hotkey['vks']:
                    self._pause_down = False
                if vk is not None and vk in self.hotkey['vks']:
                    self._hotkey_down = False
            except Exception:
                pass

        try:
            self.hotkey_thread = plateforme.start_key_listener(on_press, on_release)
            self._grab_hotkey()
            print(f"✅ Raccourci {describe_hotkey(self.hotkey)} configuré")
        except Exception as e:
            print(f"⚠️ Impossible de configurer le raccourci: {e}")
            messagebox.showerror("Raccourci clavier",
                                 f"Le raccourci global ne fonctionnera pas :\n{e}\n\n"
                                 f"{plateforme.HOTKEY_FAILURE_HINT}")

    def _open_hotkey_dialog(self, target='hotkey'):
        """Capture une combinaison : les modificateurs s'accumulent, la première touche
        normale la fige — l'ordre de pression est donc sans importance.

        `target` : 'hotkey' pour le raccourci de dictée, 'pause' pour celui qui le met en
        pause (clavier seulement)."""
        if self._hotkey_capture is not None:
            return
        with_mouse = plateforme.SUPPORTS_MOUSE and target == 'hotkey'
        default = DEFAULT_HOTKEY if target == 'hotkey' else plateforme.DEFAULT_PAUSE_HOTKEY

        dialog = tk.Toplevel(self.root)
        dialog.title("Modifier le raccourci" if target == 'hotkey'
                     else "Modifier le raccourci de pause")
        dialog.geometry("440x280")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.protocol("WM_DELETE_WINDOW", self._close_hotkey_dialog)

        tk.Label(
            dialog,
            text="Appuyez sur la combinaison souhaitée",
            font=("Arial", 10, "bold")
        ).pack(pady=(14, 2))

        tk.Label(
            dialog,
            text="Maintenez Ctrl, Alt, Maj ou %s, puis appuyez sur la touche finale%s."
                 % (MODIFIER_LABELS['win'],
                    " ou sur un bouton latéral de la souris" if with_mouse else ""),
            font=("Arial", 8),
            fg="gray",
            wraplength=400
        ).pack()

        preview = tk.Label(dialog, text="…", font=("Arial", 14, "bold"), fg="#1565c0")
        preview.pack(pady=10)

        warning = tk.Label(dialog, text="", font=("Arial", 8), fg="#e65100", wraplength=400)
        warning.pack()

        buttons = tk.Frame(dialog)
        buttons.pack(pady=12)

        validate = tk.Button(
            buttons, text="Valider", width=12, state=tk.DISABLED,
            command=self._confirm_hotkey
        )
        validate.pack(side=tk.LEFT, padx=4)

        tk.Button(
            buttons, text="Annuler", width=12, command=self._close_hotkey_dialog
        ).pack(side=tk.LEFT, padx=4)

        tk.Button(
            dialog,
            text="Rétablir %s" % describe_hotkey(default),
            font=("Arial", 8),
            relief=tk.FLAT,
            fg="#1565c0",
            command=self._reset_hotkey
        ).pack()

        mouse = None
        if with_mouse:
            def on_click(button, pressed):
                if pressed and button >= plateforme.MOUSE_BUTTON_MIN:
                    self.root.after(0, self._on_capture_button, button)
            try:
                mouse = plateforme.start_mouse_listener(on_click)
            except Exception as e:
                print(f"⚠️ Écoute de la souris impossible : {e}")

        self._hotkey_capture = {
            'dialog': dialog,
            'preview': preview,
            'warning': warning,
            'validate': validate,
            'pending': None,
            'held': frozenset(),
            'mouse': mouse,
            'target': target,
        }

    def _on_capture_modifiers(self, held):
        """Aperçu vivant tant que l'utilisateur n'a pas figé la combinaison."""
        capture = self._hotkey_capture
        if capture is None:
            return
        capture['held'] = held
        if capture['pending'] is not None:
            return
        parts = [MODIFIER_LABELS[m] for m in MODIFIER_ORDER if m in held]
        capture['preview'].config(text='+'.join(parts + ['…']) if parts else '…')

    def _on_capture_key(self, held, vk):
        """Fige la combinaison sur la première touche non-modificateur."""
        if self._hotkey_capture is None:
            return
        hotkey = {'modifiers': [m for m in MODIFIER_ORDER if m in held], 'vks': [vk]}
        if hotkey['modifiers']:
            warning = ""
        elif plateforme.GRABS_HOTKEY:
            warning = plateforme.BARE_KEY_WARNING
        elif vk in plateforme.SAFE_BARE_CODES:
            warning = ""
        else:
            warning = ("Cette touche seule se déclenchera aussi pendant que vous "
                       "tapez du texte. Ajoutez Ctrl ou Alt pour l'éviter.")
        self._show_captured(hotkey, warning)

    def _on_capture_button(self, button):
        """Fige la combinaison sur un bouton de souris, avec les modificateurs tenus."""
        capture = self._hotkey_capture
        if capture is None:
            return
        held = capture['held']
        hotkey = {'modifiers': [m for m in MODIFIER_ORDER if m in held], 'vks': [],
                  'button': button}
        self._show_captured(hotkey, "" if hotkey['modifiers'] else plateforme.BARE_BUTTON_WARNING)

    def _show_captured(self, hotkey, warning):
        capture = self._hotkey_capture
        capture['pending'] = hotkey
        capture['preview'].config(text=describe_hotkey(hotkey))
        other = self.pause_hotkey if capture['target'] == 'hotkey' else self.hotkey
        clash = (not hotkey.get('button') and not other.get('button')
                 and set(hotkey['modifiers']) == set(other['modifiers'])
                 and set(hotkey['vks']) & set(other['vks']))
        if clash:
            capture['warning'].config(
                text="Cette combinaison est déjà celle %s." % (
                    "de la pause" if capture['target'] == 'hotkey' else "de la dictée"))
            capture['validate'].config(state=tk.DISABLED)
            return
        capture['warning'].config(text=warning)
        capture['validate'].config(state=tk.NORMAL)

    def _confirm_hotkey(self):
        """Applique la combinaison capturée."""
        capture = self._hotkey_capture
        if capture is None or capture['pending'] is None:
            return
        hotkey, target = capture['pending'], capture['target']
        self._close_hotkey_dialog()
        self._apply_hotkey(hotkey, target)

    def _reset_hotkey(self):
        """Rétablit le raccourci d'origine."""
        target = self._hotkey_capture['target'] if self._hotkey_capture else 'hotkey'
        self._close_hotkey_dialog()
        if target == 'pause':
            self._apply_hotkey(normalize_hotkey(None, plateforme.DEFAULT_PAUSE_HOTKEY), 'pause')
        else:
            self._apply_hotkey(normalize_hotkey(DEFAULT_HOTKEY))

    def _close_hotkey_dialog(self):
        """Ferme la fenêtre de capture et réactive le raccourci global."""
        capture = self._hotkey_capture
        if capture is None:
            return
        self._hotkey_capture = None
        if capture['mouse'] is not None:
            capture['mouse'].stop()
        try:
            capture['dialog'].grab_release()
            capture['dialog'].destroy()
        except Exception:
            pass

def acquire_single_instance_lock():
    """Retourne le socket verrou, ou None si l'app tourne déjà (et la ramène au premier plan)."""
    lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        # Pas de SO_REUSEADDR : on veut que le second bind échoue.
        lock.bind(("127.0.0.1", SINGLE_INSTANCE_PORT))
        lock.listen(1)
        return lock
    except OSError:
        lock.close()
        try:
            with socket.create_connection(("127.0.0.1", SINGLE_INSTANCE_PORT), timeout=2):
                pass
        except OSError:
            pass
        return None

def serve_single_instance_lock(lock, app):
    """Ré-affiche la fenêtre quand un second lancement se signale."""
    def boucle():
        while True:
            try:
                conn, _ = lock.accept()
                conn.close()
            except OSError:
                return
            app.root.after(0, app._restore_window)

    threading.Thread(target=boucle, daemon=True).start()

def main():
    lock = acquire_single_instance_lock()
    if lock is None:
        print("ℹ️ Application déjà lancée : fenêtre existante ramenée au premier plan.")
        return

    root = tk.Tk()
    if START_IN_TRAY:
        root.withdraw()
    app = VoiceTranscriptionApp(root)
    serve_single_instance_lock(lock, app)

    root.protocol("WM_DELETE_WINDOW", app._hide_to_tray)

    root.mainloop()

if __name__ == '__main__':
    main()
