#!/usr/bin/env python3
"""
Application simple de transcription vocale avec OpenAI Whisper
Interface Tkinter native Windows
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import os
import json
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path
import pyperclip
from dotenv import load_dotenv
from openai import OpenAI
import sounddevice as sd
import numpy as np
import wave
from pynput import keyboard as pynput_keyboard
import pygame
import pyautogui

# Charger les variables d'environnement
load_dotenv()

# Modèles de transcription disponibles : (libellé UI, identifiant API)
MODEL_OPTIONS = [
    ("GPT-4o Mini (recommandé)", "gpt-4o-mini-transcribe"),
    ("GPT-4o", "gpt-4o-transcribe"),
    ("Whisper-1", "whisper-1"),
]

# Tarifs OpenAI en USD par seconde d'audio
MODEL_PRICING_PER_SECOND = {
    "whisper-1": 0.006 / 60,
    "gpt-4o-transcribe": 0.006 / 60,
    "gpt-4o-mini-transcribe": 0.003 / 60,
}

# Fichier d'historique persistant (à côté du script)
HISTORY_FILE = Path(__file__).parent / "transcription_history.json"

# Fichier de préférences utilisateur (micro, modèle, toggles, durée)
PREFS_FILE = Path(__file__).parent / "user_preferences.json"

DEFAULT_PREFS = {
    "selected_model": "gpt-4o-mini-transcribe",
    "selected_device_name": None,   # Nom du micro (pas l'index — plus stable entre sessions)
    "max_recording_duration": 240,
    "auto_copy": True,    # Laisser le texte dans le presse-papier après transcription
    "auto_paste": True,   # Injecter le texte dans le champ actif (indépendant de auto_copy)
    "sound_enabled": True,
    "overlay_enabled": True,                 # Afficher l'overlay flottant pendant les états
    "minimize_to_tray_on_close": True,       # La croix réduit dans la tray au lieu de quitter
    "terminal_paste": False,                 # Coller avec Ctrl+Shift+V (compatible terminaux)
    "overlay_position": None,                # [x, y] de l'overlay après drag utilisateur
}

ENV_FILE = Path(__file__).parent / ".env"

# pystray + Pillow sont optionnels : l'app démarre aussi sans, la tray est juste désactivée.
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False


class ToolTip:
    """Tooltip minimaliste en pur tkinter (pas de dep).
    Apparaît après 400ms de survol, disparaît sur Leave ou clic."""

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
            self.tip_window, text=self.text,
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
    HANDLE_DEBORD = 2           # Débord visuel de la poignée hors de l'overlay (px)
    ALPHA_OVERLAY = 0.68
    ALPHA_HANDLE_REST = 0.78    # Un peu plus opaque que l'overlay, pour que les points soient lisibles
    ALPHA_HANDLE_ACTIVE = 0.95  # Au survol / pendant le drag

    def __init__(self, root, on_position_saved=None):
        self.root = root
        self.enabled = True
        self.custom_position = None  # (x, y) coin haut-gauche de l'overlay, None = défaut centré-haut
        self.on_position_saved = on_position_saved
        self._hide_after_id = None
        self._handle_show_after_id = None  # affichage différé de la poignée (anti-jitter)
        self._drag_offset_x = 0
        self._drag_offset_y = 0
        self._drag_handle_dx = 0  # delta poignée→overlay capturé au début du drag
        self._drag_handle_dy = 0
        self._is_dragging = False

        # --- Overlay principal (click-through) ---
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

        # --- Poignée de drag (NON click-through, cliquable) ---
        # Même fond que l'overlay (mis à jour dans show()), dots blancs pour le contraste.
        # Alpha dynamique : discret au repos, plus prononcé au survol/drag.
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

        # Cachés par défaut, apparaissent lors d'un état
        self.win.withdraw()
        self.handle.withdraw()

        # Appliquer styles Windows (click-through, arrondis) après que les
        # fenêtres existent effectivement côté OS.
        self.win.update_idletasks()
        self.handle.update_idletasks()
        self._enable_click_through()
        self._apply_rounded_corners_to_overlay()
        self._apply_circular_shape_to_handle()

    def _enable_click_through(self):
        """Ajoute WS_EX_TRANSPARENT à l'overlay principal (Windows uniquement).
        Tkinter a déjà posé WS_EX_LAYERED via `-alpha` ; on n'y touche pas pour
        éviter l'état « fenêtre noire ». SetWindowPos(SWP_FRAMECHANGED) commit la modif."""
        try:
            import ctypes
            hwnd = int(self.win.wm_frame(), 16)
            GWL_EXSTYLE = -20
            WS_EX_TRANSPARENT = 0x00000020
            SWP_NOMOVE = 0x0002; SWP_NOSIZE = 0x0001; SWP_NOZORDER = 0x0004
            SWP_FRAMECHANGED = 0x0020
            user32 = ctypes.windll.user32
            style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style | WS_EX_TRANSPARENT)
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        except Exception as e:
            print(f"⚠️ Click-through overlay non activé: {e}")

    def _apply_rounded_corners_to_overlay(self):
        """Applique des coins arrondis via CreateRoundRectRgn (Windows, pas de dep)."""
        try:
            import ctypes
            hwnd = int(self.win.wm_frame(), 16)
            w = self.win.winfo_width()
            h = self.win.winfo_height()
            if w <= 1 or h <= 1:
                return
            region = ctypes.windll.gdi32.CreateRoundRectRgn(
                0, 0, w + 1, h + 1, self.CORNER_RADIUS, self.CORNER_RADIUS
            )
            ctypes.windll.user32.SetWindowRgn(hwnd, region, True)
        except Exception as e:
            print(f"⚠️ Arrondi overlay non appliqué: {e}")

    def _apply_circular_shape_to_handle(self):
        """Rend la poignée parfaitement ronde via CreateEllipticRgn."""
        try:
            import ctypes
            hwnd = int(self.handle.wm_frame(), 16)
            w = self.handle.winfo_width()
            h = self.handle.winfo_height()
            if w <= 1 or h <= 1:
                return
            region = ctypes.windll.gdi32.CreateEllipticRgn(0, 0, w + 1, h + 1)
            ctypes.windll.user32.SetWindowRgn(hwnd, region, True)
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

        # Sécurité multi-écran : garder l'overlay visible
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
        # Coin bas-droite, avec ~2px de débord sur la droite et en bas
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
        # Offset du clic dans la poignée
        self._drag_offset_x = event.x_root - self.handle.winfo_rootx()
        self._drag_offset_y = event.y_root - self.handle.winfo_rooty()
        # Delta poignée → overlay capturé une fois au début (indépendant de la
        # géométrie, fonctionne qu'on place la poignée haut-gauche ou bas-droite)
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

        # Clamp aux bornes de l'écran (sécurité multi-écran)
        screen_w = self.win.winfo_screenwidth()
        screen_h = self.win.winfo_screenheight()
        w = self.win.winfo_width()
        h = self.win.winfo_height()
        new_overlay_x = max(0, min(new_overlay_x, max(0, screen_w - w)))
        new_overlay_y = max(0, min(new_overlay_y, max(0, screen_h - h)))

        self.win.geometry(f"+{new_overlay_x}+{new_overlay_y}")
        # Repositionner la poignée relativement à la nouvelle position de l'overlay
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
        # Poignée : même couleur que l'overlay pour qu'elle s'y intègre visuellement
        self.handle.configure(bg=bg)
        self.handle_label.configure(bg=bg)

        # Tk ne finalise la géométrie d'un Toplevel qu'une fois mappé. Au premier
        # show, un update_idletasks() ne suffit donc pas : la fenêtre apparaît
        # brièvement à une taille pré-layout puis se redimensionne (flash visible).
        # Parade : au premier show, on la rend invisible (alpha=0) et on la place
        # hors écran, on force un cycle d'événements pour que tk mesure vraiment,
        # puis on la replace à la bonne position et on rétablit l'alpha.
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

        # Poignée : affichage différé (~30ms) si pas déjà visible — laisse le
        # temps à tk de finaliser totalement avant qu'on la positionne.
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
        # Si entre-temps l'overlay a été masqué ou désactivé, on renonce
        if not self.enabled or not self.win.winfo_viewable():
            return
        # Re-positionner maintenant que les tailles sont réellement définitives
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
        # Annuler l'affichage différé de la poignée si en attente
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
        # Taille par défaut plus grande et adaptée à un écran standard
        self.root.geometry("800x800")
        self.root.resizable(True, True)
        # Taille minimale pour éviter que la fenêtre soit trop petite
        self.root.minsize(600, 500)
        
        # État de l'application
        self.is_recording = False
        self.is_transcribing = False  # Indique si une transcription est en cours
        self.cancel_requested = False  # Flag pour annuler la transcription
        self.audio_frames = []
        self.recording_thread = None
        self.recording_start_time = None  # Timestamp du début d'enregistrement
        self.selected_device_index = None  # Index du microphone sélectionné

        # Préférences utilisateur persistantes (micro, modèle, durée, toggles)
        self.prefs = self._load_prefs()

        # Modèle de transcription sélectionné (via prefs)
        self.selected_model = self.prefs.get("selected_model", "gpt-4o-mini-transcribe")
        # Sécurité : si le fichier de prefs contient un modèle inconnu, retomber sur le défaut
        if self.selected_model not in MODEL_PRICING_PER_SECOND:
            self.selected_model = "gpt-4o-mini-transcribe"

        # Nom du micro préféré (appliqué par _load_microphones si toujours présent)
        self.preferred_mic_name = self.prefs.get("selected_device_name")

        # Suivi des coûts de la session (réinitialisé à chaque démarrage)
        self.session_cost = 0.0
        self.session_transcriptions = 0
        self.session_started_at = datetime.now()

        # Historique persistant (chargé au démarrage)
        self.history = self._load_history()

        # Historique pour undo (Ctrl+Z)
        self.text_history = []
        self.history_index = -1

        # Thread pour les raccourcis clavier
        self.hotkey_thread = None
        self.esc_thread = None

        # Configuration audio
        self.CHANNELS = 1
        self.RATE = 44100
        self.DTYPE = np.int16

        # Durée maximum d'enregistrement (via prefs)
        self.max_recording_duration = int(self.prefs.get("max_recording_duration", 240))

        # Nom du micro affiché avant que _load_microphones ne remplisse la liste
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

        # Client OpenAI — non bloquant si la clé est absente (l'utilisateur peut la
        # renseigner via l'UI). Le bouton d'enregistrement affichera un message
        # d'erreur clair si on tente d'enregistrer sans clé.
        self.client = None
        api_key = os.getenv('OPENAI_API_KEY')
        if api_key:
            try:
                self.client = OpenAI(api_key=api_key)
            except Exception as e:
                print(f"⚠️ Init OpenAI échouée (clé probablement invalide): {e}")

        # Chemin des sons
        self.sounds_dir = os.path.join(os.path.dirname(__file__), 'sounds')
        os.makedirs(self.sounds_dir, exist_ok=True)

        # Interface (créée en premier pour affichage immédiat)
        self.setup_ui()

        # Overlay flottant (créé après la fenêtre principale pour héritage correct)
        self.overlay = FloatingOverlay(
            self.root,
            on_position_saved=self._on_overlay_position_saved
        )
        self.overlay.enabled = self.overlay_enabled_var.get()
        # Restaurer la position sauvegardée (si elle existe et est bien formée)
        saved_pos = self.prefs.get("overlay_position")
        if isinstance(saved_pos, (list, tuple)) and len(saved_pos) == 2:
            try:
                self.overlay.custom_position = (int(saved_pos[0]), int(saved_pos[1]))
            except (TypeError, ValueError):
                pass

        # Initialisations lourdes différées (pygame + hotkey + tray) pour que la
        # fenêtre s'affiche instantanément, sans frame noir ni délai visible.
        self.tray_icon = None
        self.root.after(100, self._deferred_init)
    
    def setup_ui(self):
        """Crée l'interface utilisateur"""
        
        # Frame principal avec layout horizontal
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Colonne de gauche (options) — élargie pour que les libellés tiennent
        left_panel = tk.Frame(main_container, width=240)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        left_panel.pack_propagate(False)
        
        # Colonne principale (contenu)
        main_frame = tk.Frame(main_container)
        main_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Titre
        title_label = tk.Label(
            main_frame,
            text="🎙️ Transcription Vocale",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        # Frame pour les boutons d'enregistrement
        record_button_frame = tk.Frame(main_frame)
        record_button_frame.pack(pady=10)
        
        # Bouton Enregistrer/Arrêter
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
        
        # Bouton Annuler (caché par défaut)
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
        # Ne pas packer pour l'instant, sera affiché pendant l'enregistrement
        
        # Affichage du raccourci clavier (sous les boutons)
        self.hotkey_label = tk.Label(
            main_frame,
            text="⌨️ Ctrl+Alt+9 (global) — Échap pour annuler",
            font=("Arial", 9),
            fg="blue"
        )
        self.hotkey_label.pack(pady=(5, 10))
        
        # Panel gauche : Options avec cadre
        options_frame = tk.LabelFrame(
            left_panel,
            text="Options",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        options_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Clé API OpenAI — section en haut pour les utilisateurs qui partent d'un repo nu
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

        # Microphone actuel
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
        
        # Sélection du microphone
        tk.Label(
            options_frame,
            text="Changer:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(5, 2))
        
        self.mic_var = tk.StringVar()
        self.mic_dropdown = tk.OptionMenu(options_frame, self.mic_var, "Chargement...")
        self.mic_dropdown.config(width=24, font=("Arial", 8), anchor="w")
        self.mic_dropdown.pack(anchor=tk.W, fill=tk.X, pady=(0, 10))

        # Sélection du modèle de transcription
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

        # Curseur pour la durée maximum (horizontal)
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
            showvalue=False  # Retirer le chiffre en secondes qui suit le curseur
        )
        self.duration_slider.pack(fill=tk.X, pady=5, padx=5)
        
        self.duration_label = tk.Label(
            options_frame,
            text="4 min",
            font=("Arial", 9, "bold"),
            fg="blue"
        )
        self.duration_label.pack(pady=(0, 5))

        # Toggles comportement (les BooleanVars sont créées dans __init__
        # pour attacher les traces d'auto-save avant le premier rendu).
        # Copie et collage sont décorrélés : coller sans copier est possible
        # (le presse-papier est restauré à sa valeur d'origine après collage).
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
            "(Cursor, VS Code, Windows Terminal…) qui n'accepte que Ctrl+Maj+V. "
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

        # Séparateur + suivi des coûts de la session
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

        # Statistiques mois en cours
        tk.Frame(options_frame, height=1, bg="#cccccc").pack(fill=tk.X, pady=(8, 6))

        tk.Label(
            options_frame,
            text="Ce mois:",
            font=("Arial", 9, "bold"),
            anchor="w"
        ).pack(anchor=tk.W, fill=tk.X, pady=(0, 2))

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

        # Charger la liste des microphones (après que l'interface soit créée)
        self.root.after(100, self._load_microphones)
        
        # Zone de texte pour la transcription
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
            width=70
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # Raccourci Ctrl+Z pour undo
        self.text_area.bind('<Control-z>', self.undo_clear)
        self.text_area.bind('<Control-Z>', self.undo_clear)
        self.root.bind('<Control-z>', self.undo_clear)
        self.root.bind('<Control-Z>', self.undo_clear)
        
        # Boutons d'action
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

        # Status bar en bas : enregistrement / transcription / feedback
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

        # Initialiser affichages dynamiques (stats mois + clé API + label durée)
        self._update_session_display()
        self._update_api_key_status(ok=(self.client is not None))
        # Forcer la mise à jour du label "durée" pour matcher la valeur persistée
        self._on_duration_change(self.max_recording_duration)
        # Appliquer l'état visuel initial des toggles de collage (enabled/disabled)
        self._on_paste_toggles_changed()

    def _load_microphones(self):
        """Charge la liste des microphones disponibles (filtrés et dédupliqués)"""
        try:
            devices = sd.query_devices()
            
            # Mots-clés à exclure (périphériques virtuels, mappers, etc.)
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
            
            # Filtrer les périphériques d'entrée réels
            input_devices = []
            seen_names = set()  # Pour dédupliquer par nom exact
            
            for idx, device in enumerate(devices):
                if device['max_input_channels'] > 0:
                    name = device['name']
                    name_lower = name.lower()
                    
                    # Exclure les périphériques avec patterns suspects
                    if '@' in name or 'system32' in name_lower or '.sys' in name_lower:
                        continue
                    
                    # Exclure les périphériques virtuels/non-microphones par mots-clés
                    if any(keyword in name_lower for keyword in exclude_keywords):
                        continue
                    
                    # Dédupliquer par nom exact (même nom = même périphérique, peu importe le hostapi)
                    # Normaliser le nom (enlever espaces en fin, parenthèses manquantes, etc.)
                    normalized_name = name.strip()
                    # Corriger les noms tronqués (ajouter parenthèse fermante si manquante)
                    if normalized_name.count('(') > normalized_name.count(')'):
                        normalized_name += ')'
                    
                    if normalized_name not in seen_names:
                        seen_names.add(normalized_name)
                        input_devices.append({
                            'index': idx,  # Index réel dans sd.query_devices()
                            'name': normalized_name,
                            'device': device
                        })
            
            if input_devices:
                # Trier par nom pour une meilleure lisibilité
                input_devices.sort(key=lambda x: x['name'])
                
                device_names = [d['name'] for d in input_devices]
                current_mic = self.microphone_name if hasattr(self, 'microphone_name') else device_names[0]
                self.mic_var.set(current_mic)
                
                # Recréer le menu avec les microphones filtrés
                menu = self.mic_dropdown['menu']
                menu.delete(0, 'end')
                
                for device_info in input_devices:
                    name = device_info['name']
                    real_index = device_info['index']  # Index réel du périphérique
                    menu.add_command(
                        label=name,
                        command=lambda idx=real_index, n=name: self._change_microphone(idx, n)
                    )
                
                # Priorité 1 : micro persisté dans les préférences (si toujours disponible)
                chosen = None
                if self.preferred_mic_name:
                    for device_info in input_devices:
                        if device_info['name'] == self.preferred_mic_name:
                            chosen = device_info
                            break

                # Priorité 2 : micro par défaut de Windows
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

                # Priorité 3 : premier micro disponible
                if chosen is None:
                    chosen = input_devices[0]

                self.selected_device_index = chosen['index']
                self.microphone_name = chosen['name']
                self.mic_var.set(self.microphone_name)
                
                # Mettre à jour l'affichage
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

    def _deferred_init(self):
        """Initialisations lourdes différées pour accélérer l'apparition de la fenêtre"""
        try:
            pygame.mixer.init()
        except Exception as e:
            print(f"⚠️ Init pygame mixer échouée: {e}")

        self.setup_global_hotkey()
        self._setup_tray()

    def _make_tray_icon_image(self):
        """Génère programmatiquement l'icône de la tray (évite de shipper un binaire)."""
        img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.ellipse((4, 4, 60, 60), fill=(220, 50, 50, 255))       # fond rouge
        # Corps du micro (rectangle arrondi, fallback si Pillow ancien)
        try:
            d.rounded_rectangle((24, 16, 40, 40), radius=8, fill=(255, 255, 255, 255))
        except AttributeError:
            d.rectangle((24, 16, 40, 40), fill=(255, 255, 255, 255))
        d.rectangle((30, 40, 34, 48), fill=(255, 255, 255, 255))  # pied
        d.rectangle((22, 48, 42, 52), fill=(255, 255, 255, 255))  # base
        return img

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

            menu = pystray.Menu(
                pystray.MenuItem("Afficher", on_show, default=True),
                pystray.MenuItem("Quitter", on_quit),
            )

            self.tray_icon = pystray.Icon(
                "whisper-voice",
                icon_image,
                "Transcription Vocale",
                menu
            )

            self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
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

    # Classes de fenêtre natives des terminaux Windows connus. Les terminaux
    # intégrés des éditeurs Electron (Cursor, VS Code) ne sont PAS ici car ils
    # partagent la classe de l'éditeur — pour ceux-là, activer le toggle manuel.
    _TERMINAL_WINDOW_CLASSES = {
        'CASCADIA_HOSTING_WINDOW_CLASS',   # Windows Terminal
        'ConsoleWindowClass',              # cmd, PowerShell classique
        'PuTTY',
        'mintty',                          # Git Bash, msys2
    }

    def _is_terminal_window_focused(self):
        """Retourne True si la fenêtre au premier plan est un terminal natif.
        Ignorer les erreurs : sur non-Windows ou si ctypes échoue, on renvoie False."""
        try:
            import ctypes
            user32 = ctypes.windll.user32
            hwnd = user32.GetForegroundWindow()
            if not hwnd:
                return False
            buf = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, buf, 256)
            return buf.value in self._TERMINAL_WINDOW_CLASSES
        except Exception:
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
        # Invariants
        if self.terminal_paste_var.get() and not self.auto_paste_var.get():
            self.auto_paste_var.set(True)
            return  # la trace auto_paste va re-déclencher la méthode
        if not self.auto_paste_var.get() and self.terminal_paste_var.get():
            self.terminal_paste_var.set(False)
            return

        # État visuel
        if hasattr(self, 'auto_paste_checkbox') and hasattr(self, 'terminal_paste_checkbox'):
            auto_paste_on = self.auto_paste_var.get()
            terminal_on = self.terminal_paste_var.get()
            # Si terminal est coché, auto_paste est verrouillé à True
            self.auto_paste_checkbox.config(state='disabled' if terminal_on else 'normal')
            # Si auto_paste est décoché, terminal_paste est verrouillé à False
            self.terminal_paste_checkbox.config(state='normal' if auto_paste_on else 'disabled')

        self._save_prefs()

    def _quit_app(self):
        """Quitte proprement : sauvegarde, arrêt de la tray, destruction de la fenêtre."""
        try:
            self._save_history()
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
        # 1) Validation : appel gratuit (models.list) pour vérifier que la clé est active
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

        # 2) Écriture dans .env en préservant les autres variables éventuelles
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

    def _log_transcription(self, model, duration_sec, cost_usd):
        """Enregistre une transcription dans l'historique persistant"""
        entry = {
            "at": datetime.now().isoformat(timespec='seconds'),
            "model": model,
            "duration_sec": round(duration_sec, 2),
            "cost_usd": round(cost_usd, 6),
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
        # Overlay flottant (visible même app réduite/cachée)
        if hasattr(self, 'overlay'):
            self.overlay.show(f"🔴 {timer_text}", bg="#c62828")

        self.root.after(500, self._update_recording_timer)
    
    def _on_duration_change(self, value):
        """Appelé quand le curseur change"""
        # Arrondir à la demi-minute la plus proche (snap à 30 secondes)
        raw_value = int(float(value))
        # Arrondir à la demi-minute la plus proche
        rounded_value = round(raw_value / 30) * 30
        
        # Si la valeur a changé, mettre à jour le slider + sauvegarder
        if rounded_value != self.max_recording_duration:
            self.max_recording_duration = rounded_value
            # Mettre à jour la valeur du slider sans déclencher le callback
            self.duration_var.set(rounded_value)
            self._save_prefs()
        
        # Afficher en minutes (toujours des multiples de 30 secondes)
        if self.max_recording_duration >= 60:
            minutes = self.max_recording_duration // 60
            seconds = self.max_recording_duration % 60
            if seconds == 0:
                display = f"{minutes} min"
            else:
                display = f"{minutes} min {seconds}s"
        else:
            # En dessous de 60 secondes, on garde les secondes mais on snap quand même
            display = f"{self.max_recording_duration}s"
        
        self.duration_label.config(text=display)
    
    def toggle_recording(self):
        """Bascule entre démarrer et arrêter l'enregistrement"""
        # Empêcher les appels parallèles
        if self.is_transcribing:
            # Une transcription est en cours, jouer un son d'erreur
            self.play_sound('error.wav', volume=0.3)
            return

        if not self.is_recording:
            # Vérifier que la clé API est configurée avant de démarrer
            if self.client is None:
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
            # Essayer WAV d'abord
            if not os.path.exists(sound_path):
                # Essayer sans extension
                sound_path = sound_path.replace('.mp3', '.wav')

            if os.path.exists(sound_path):
                # Utiliser pygame pour contrôler le volume
                sound = pygame.mixer.Sound(sound_path)
                sound.set_volume(volume)  # Volume réglable (0.0 à 1.0)
                sound.play()
        except Exception as e:
            # Fallback sur winsound si pygame échoue
            try:
                import winsound
                winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
            except:
                print(f"Erreur lecture son: {e}")
    
    def start_recording(self):
        """Démarre l'enregistrement audio"""
        # Double vérification (au cas où)
        if self.is_transcribing:
            self.play_sound('error.wav', volume=0.3)
            return
        
        self.is_recording = True
        self.cancel_requested = False
        self.audio_frames = []
        self.recording_start_time = time.time()
        
        # Jouer le son de début
        self.play_sound('start.wav')
        
        # Mettre à jour l'interface
        self.record_button.config(
            text="⏹️ Arrêter",
            bg="#f44336",
            activebackground="#da190b"
        )
        # Afficher le bouton Annuler
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        self.status_label.config(text="🔴 Enregistrement en cours...", fg="red")
        
        # Démarrer l'enregistrement dans un thread séparé
        self.recording_thread = threading.Thread(target=self._record_audio, daemon=True)
        self.recording_thread.start()

        # Lancer le timer visuel dans la barre de statut
        self.root.after(100, self._update_recording_timer)
    
    def cancel_recording(self):
        """Annule l'enregistrement sans faire l'appel API"""
        if self.is_recording:
            self.is_recording = False
            self.cancel_requested = True
            self.audio_frames = []
            # Jouer le son d'erreur pour feedback sonore
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
        
        # Cacher le bouton Annuler
        self.cancel_button.pack_forget()
        
        # Jouer le son léger indiquant que l'enregistrement est arrêté (avant transcription)
        self.play_sound('recording_stopped.wav', volume=0.2)

        self.status_label.config(text="⏳ Traitement en cours...", fg="orange")
        self.record_button.config(state=tk.DISABLED)
        if hasattr(self, 'overlay'):
            self.overlay.show("⏳ Transcription…", bg="#1565c0")
        
        # Arrêter l'enregistrement dans un thread séparé
        threading.Thread(target=self._process_recording, daemon=True).start()
    
    def _record_audio(self):
        """Enregistre l'audio en continu via un InputStream sounddevice.

        Le callback est appelé automatiquement par PortAudio sur chaque bloc
        capturé, sans gap entre les blocs — contrairement à la boucle précédente
        `sd.rec() + sd.wait()` qui laissait passer ~10-20ms entre chaque chunk.
        """
        device = self.selected_device_index if self.selected_device_index is not None else None

        def audio_callback(indata, frames, time_info, status):
            # `indata` est un buffer numpy réutilisé — on copie avant de stocker.
            if self.is_recording and not self.cancel_requested:
                self.audio_frames.append(indata.copy().tobytes())

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
        # Vérifier si l'annulation a été demandée
        if self.cancel_requested:
            self.is_transcribing = False
            return
        
        # Marquer qu'une transcription est en cours
        self.is_transcribing = True
        
        try:
            if not self.audio_frames or self.cancel_requested:
                self.is_transcribing = False
                if not self.cancel_requested:
                    self.root.after(0, lambda: self._reset_ui("Aucun audio enregistré"))
                return
            
            # Sauvegarder dans un fichier temporaire
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.wav')
            temp_filename = temp_file.name
            temp_file.close()
            
            # Écrire le fichier WAV
            wf = wave.open(temp_filename, 'wb')
            wf.setnchannels(self.CHANNELS)
            wf.setsampwidth(2)  # 16-bit = 2 bytes
            wf.setframerate(self.RATE)
            wf.writeframes(b''.join(self.audio_frames))
            wf.close()
            
            # Vérifier si l'annulation a été demandée avant de lancer l'API
            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return

            # Transcrit avec OpenAI
            self.root.after(0, lambda: self.status_label.config(text="🔄 Transcription en cours...", fg="blue"))
            
            # Note: Une fois l'appel API lancé, on ne peut pas l'annuler facilement
            # Mais on peut éviter de lancer l'appel si cancel_requested est True
            model_used = self.selected_model
            with open(temp_filename, 'rb') as audio_file:
                transcript = self.client.audio.transcriptions.create(
                    model=model_used,
                    file=audio_file,
                    language="fr"
                )

            # Vérifier à nouveau après l'appel
            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return

            text = transcript.text

            # Nettoyer le fichier temporaire
            os.unlink(temp_filename)

            # Mise à jour du suivi des coûts (uniquement sur transcription réussie).
            # Calcul de la durée à partir des bytes réels (plus fiable avec InputStream
            # dont la taille de bloc est variable) : bytes / 2 (int16) / taux = secondes.
            total_bytes = sum(len(f) for f in self.audio_frames)
            audio_duration_seconds = total_bytes / 2 / self.RATE
            rate = MODEL_PRICING_PER_SECOND.get(model_used, 0)
            cost = audio_duration_seconds * rate
            self.session_cost += cost
            self.session_transcriptions += 1
            self._log_transcription(model_used, audio_duration_seconds, cost)
            self.root.after(0, self._update_session_display)

            # Afficher le texte dans l'interface
            self.root.after(0, lambda: self._display_text(text))
            
        except Exception as e:
            self.is_transcribing = False
            error_msg = f"Erreur lors de la transcription: {e}"
            self.root.after(0, lambda: messagebox.showerror("Erreur", error_msg))
            self.root.after(0, lambda: self._reset_ui("Erreur"))
    
    def _display_text(self, text):
        """Affiche le texte transcrit et gère copie/collage selon les toggles utilisateur.

        Les deux toggles sont indépendants :
        - auto_copy seul        → texte dans le presse-papier, pas d'injection
        - auto_paste seul       → injection dans le champ actif, presse-papier restauré
        - les deux              → texte collé ET conservé dans le presse-papier
        - aucun                 → texte uniquement visible dans la zone de transcription
        """
        # Ajouter le texte à la zone de texte
        self.text_area.insert(tk.END, text + "\n\n")
        self.text_area.see(tk.END)

        # Réinitialiser l'historique après une nouvelle transcription
        self.text_history = []
        self.history_index = -1

        auto_copy = self.auto_copy_var.get()
        auto_paste = self.auto_paste_var.get()

        status_text = "✅ Transcription terminée"
        status_color = "green"

        if auto_paste:
            # Pour injecter dans le champ actif sans écraser durablement le presse-papier,
            # on sauvegarde l'ancien contenu si l'utilisateur n'a pas activé auto_copy.
            saved_clipboard = None
            if not auto_copy:
                try:
                    saved_clipboard = pyperclip.paste()
                except Exception:
                    saved_clipboard = None

            try:
                pyperclip.copy(text)
                time.sleep(0.15)
                # Utiliser Ctrl+Maj+V si :
                #  - l'option manuelle est activée (utile pour terminaux intégrés
                #    type Cursor/VS Code où la détection par classe de fenêtre échoue)
                #  - OU si la fenêtre active est un terminal connu (Windows Terminal,
                #    cmd classique, PuTTY, mintty…)
                if self.terminal_paste_var.get() or self._is_terminal_window_focused():
                    pyautogui.hotkey('ctrl', 'shift', 'v')
                else:
                    pyautogui.hotkey('ctrl', 'v')
                # Laisser le temps au collage de se terminer avant de restaurer
                time.sleep(0.15)
            except Exception:
                status_text = "✅ Transcription terminée (collage échoué)"
                status_color = "orange"

            if not auto_copy and saved_clipboard is not None:
                try:
                    pyperclip.copy(saved_clipboard)
                except Exception:
                    pass

            if status_color == "green":
                status_text = (
                    "✅ Collé et copié"
                    if auto_copy else
                    "✅ Collé (presse-papier préservé)"
                )
        elif auto_copy:
            try:
                pyperclip.copy(text)
                status_text = "✅ Copié dans le presse-papier"
            except Exception:
                status_text = "✅ Transcription terminée (copie échouée)"
                status_color = "orange"

        # Son de fin + status
        self.play_sound('end.wav')
        self._reset_ui(status_text)
        self.status_label.config(text=status_text, fg=status_color)

        # Overlay flottant : flash bref pour confirmer le résultat
        if hasattr(self, 'overlay'):
            overlay_bg = "#2e7d32" if status_color == "green" else "#ff9800"
            self.overlay.show_briefly(status_text, bg=overlay_bg, duration_ms=1800)

        # Marquer que la transcription est terminée
        self.is_transcribing = False
    
    def _reset_ui(self, status_text="Prêt"):
        """Réinitialise l'interface après l'enregistrement"""
        self.record_button.config(
            text="🎤 Démarrer l'enregistrement",
            bg="#4CAF50",
            activebackground="#45a049",
            state=tk.NORMAL
        )
        # Cacher le bouton Annuler
        self.cancel_button.pack_forget()
        
        self.status_label.config(text=status_text, fg="gray")
        
        self.audio_frames = []
        self.recording_start_time = None
        self.cancel_requested = False
    
    def clear_text(self):
        """Efface le contenu de la zone de texte"""
        # Sauvegarder dans l'historique avant d'effacer
        current_text = self.text_area.get(1.0, tk.END).strip()
        if current_text:
            self.text_history.append(current_text)
            self.history_index = len(self.text_history) - 1
            # Limiter l'historique à 10 entrées
            if len(self.text_history) > 10:
                self.text_history.pop(0)
                self.history_index = len(self.text_history) - 1
        
        self.text_area.delete(1.0, tk.END)
        self.status_label.config(text="✅ Texte effacé (Ctrl+Z pour annuler)", fg="gray")
        self.root.after(3000, lambda: self.status_label.config(text="Prêt", fg="gray"))
    
    def undo_clear(self, event=None):
        """Restaure le texte précédemment effacé (Ctrl+Z)"""
        # Vérifier si on a un historique
        if len(self.text_history) == 0:
            return None
        
        # Vérifier si la zone est vide ou presque vide
        current_text = self.text_area.get(1.0, tk.END).strip()
        
        # Si la zone est vide, restaurer le dernier texte effacé
        if not current_text and self.history_index >= 0:
            restored_text = self.text_history[self.history_index]
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(1.0, restored_text)
            self.status_label.config(text="↩️ Texte restauré", fg="green")
            self.root.after(2000, lambda: self.status_label.config(text="Prêt", fg="gray"))
            # Décrémenter l'index pour pouvoir restaurer plusieurs fois
            self.history_index -= 1
            if self.history_index < 0:
                self.history_index = -1
            return "break"  # Empêcher le comportement par défaut de Ctrl+Z
        
        return "break"  # Toujours empêcher le comportement par défaut
    
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
    
    def setup_global_hotkey(self):
        """Configure Ctrl+Alt+9 (toggle enregistrement) + Échap (annulation).

        On utilise un `Listener` avec suivi manuel des modificateurs plutôt que
        `GlobalHotKeys`, parce que certains logiciels de souris (Logitech G Hub,
        Razer Synapse, etc.) envoient les keystrokes synthétiques avec un timing
        ou un ordre que `GlobalHotKeys` ne reconnaît pas toujours. Le suivi
        manuel accepte n'importe quel ordre de pression et fonctionne avec les
        raccourcis souris.
        """
        modifiers = {'ctrl': False, 'alt': False}

        def is_key_9(key):
            """Reconnaît le 9 alphanumérique (VK=57) comme le pavé numérique (VK_NUMPAD9=105)."""
            vk = getattr(key, 'vk', None)
            if vk in (57, 105):
                return True
            if hasattr(key, 'char') and key.char == '9':
                return True
            return False

        def on_press(key):
            try:
                if key in (pynput_keyboard.Key.ctrl_l, pynput_keyboard.Key.ctrl_r):
                    modifiers['ctrl'] = True
                elif key in (pynput_keyboard.Key.alt_l, pynput_keyboard.Key.alt_r, pynput_keyboard.Key.alt_gr):
                    modifiers['alt'] = True
                elif key == pynput_keyboard.Key.esc:
                    if self.is_recording:
                        self.root.after(0, self.cancel_recording)
                elif is_key_9(key) and modifiers['ctrl'] and modifiers['alt']:
                    self.root.after(0, self.toggle_recording)
            except Exception:
                pass

        def on_release(key):
            try:
                if key in (pynput_keyboard.Key.ctrl_l, pynput_keyboard.Key.ctrl_r):
                    modifiers['ctrl'] = False
                elif key in (pynput_keyboard.Key.alt_l, pynput_keyboard.Key.alt_r, pynput_keyboard.Key.alt_gr):
                    modifiers['alt'] = False
            except Exception:
                pass

        try:
            def start_listener():
                with pynput_keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
                    listener.join()

            self.hotkey_thread = threading.Thread(target=start_listener, daemon=True)
            self.hotkey_thread.start()
            print("✅ Raccourci Ctrl+Alt+9 configuré")
        except Exception as e:
            print(f"⚠️ Impossible de configurer le raccourci: {e}")
            print("💡 Essayez de lancer l'application en tant qu'administrateur")
    

def main():
    root = tk.Tk()
    app = VoiceTranscriptionApp(root)

    # Le X de la fenêtre réduit dans la tray (si disponible), sinon quitte.
    # Pour réellement quitter, utiliser le menu clic-droit de la tray.
    root.protocol("WM_DELETE_WINDOW", app._hide_to_tray)

    # Démarrer la boucle principale
    root.mainloop()

if __name__ == '__main__':
    main()
