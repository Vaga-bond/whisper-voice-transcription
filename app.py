#!/usr/bin/env python3
"""
Application simple de transcription vocale avec OpenAI Whisper
Interface Tkinter native Windows
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import os
import tempfile
import time
from datetime import datetime
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

class VoiceTranscriptionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Transcription Vocale - Whisper")
        # Taille par défaut plus grande et adaptée à un écran standard
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        # Taille minimale pour éviter que la fenêtre soit trop petite
        self.root.minsize(600, 400)
        
        # État de l'application
        self.is_recording = False
        self.is_transcribing = False  # Indique si une transcription est en cours
        self.cancel_requested = False  # Flag pour annuler la transcription
        self.audio_frames = []
        self.audio_stream = None
        self.audio = None
        self.recording_thread = None
        self.recording_start_time = None  # Timestamp du début d'enregistrement
        self.selected_device_index = None  # Index du microphone sélectionné
        
        # Historique pour undo (Ctrl+Z)
        self.text_history = []
        self.history_index = -1
        
        # Thread pour les raccourcis clavier
        self.hotkey_thread = None
        self.pressed_keys = set()
        
        # Configuration audio
        self.CHANNELS = 1
        self.RATE = 44100
        self.DTYPE = np.int16
        
        # Durée maximum d'enregistrement (en secondes, modifiable via slider)
        self.max_recording_duration = 60  # Par défaut 60 secondes
        
        # Client OpenAI
        api_key = os.getenv('OPENAI_API_KEY')
        if not api_key:
            messagebox.showerror(
                "Erreur de configuration",
                "Clé API OpenAI non trouvée !\n\n"
                "1. Copiez .env.example vers .env\n"
                "2. Ajoutez votre clé API dans .env"
            )
            self.root.destroy()
            return
        
        try:
            self.client = OpenAI(api_key=api_key)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur d'initialisation OpenAI: {e}")
            self.root.destroy()
            return
        
        # Détecter le microphone par défaut
        try:
            default_input_idx = sd.default.device[0]
            if default_input_idx is not None and default_input_idx >= 0:
                device_info = sd.query_devices(default_input_idx)
                self.microphone_name = device_info.get('name', 'Micro inconnu')
                self.selected_device_index = default_input_idx
            else:
                devices = sd.query_devices()
                input_devices = [d for d in devices if d['max_input_channels'] > 0]
                if input_devices:
                    self.microphone_name = input_devices[0]['name']
                    self.selected_device_index = 0
                else:
                    self.microphone_name = "Aucun micro détecté"
                    self.selected_device_index = None
        except Exception as e:
            self.microphone_name = f"Micro par défaut (erreur: {str(e)[:30]})"
            self.selected_device_index = None
        
        # Chemin des sons
        self.sounds_dir = os.path.join(os.path.dirname(__file__), 'sounds')
        os.makedirs(self.sounds_dir, exist_ok=True)
        
        # Initialiser pygame pour les sons (avec contrôle du volume)
        try:
            pygame.mixer.init()
        except:
            pass
        
        # Interface
        self.setup_ui()
        
        # Raccourci clavier global Ctrl+Alt+9
        self.setup_global_hotkey()
    
    def setup_ui(self):
        """Crée l'interface utilisateur"""
        
        # Frame principal avec layout horizontal
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Colonne de gauche (options)
        left_panel = tk.Frame(main_container, width=180)
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
        
        # Indicateur d'enregistrement
        self.status_label = tk.Label(
            main_frame,
            text="",
            font=("Arial", 10),
            fg="gray"
        )
        self.status_label.pack()
        
        
        # Affichage du raccourci clavier
        self.hotkey_label = tk.Label(
            main_frame,
            text="⌨️ Ctrl+Alt+9 (global)",
            font=("Arial", 9),
            fg="blue"
        )
        self.hotkey_label.pack(pady=(0, 10))
        
        # Panel gauche : Options avec cadre
        options_frame = tk.LabelFrame(
            left_panel,
            text="Options",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        options_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
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
            wraplength=150,
            justify=tk.LEFT
        )
        self.mic_display_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Sélection du microphone
        tk.Label(
            options_frame,
            text="Changer:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(5, 2))
        
        self.mic_var = tk.StringVar()
        self.mic_dropdown = tk.OptionMenu(options_frame, self.mic_var, "Chargement...")
        self.mic_dropdown.config(width=18, font=("Arial", 8))
        self.mic_dropdown.pack(anchor=tk.W, pady=(0, 10))
        
        # Curseur pour la durée maximum (horizontal)
        tk.Label(
            options_frame,
            text="Durée max:",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(5, 2))
        
        self.duration_var = tk.IntVar(value=60)
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
            text="1 min",
            font=("Arial", 9, "bold"),
            fg="blue"
        )
        self.duration_label.pack(pady=(0, 0))
        
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
                
                # Sélectionner le microphone par défaut de Windows
                try:
                    default_input_idx = sd.default.device[0]
                    if default_input_idx is not None and default_input_idx >= 0:
                        # Trouver le périphérique par défaut dans la liste filtrée
                        default_device = None
                        for device_info in input_devices:
                            if device_info['index'] == default_input_idx:
                                default_device = device_info
                                break
                        
                        if default_device:
                            self.selected_device_index = default_device['index']
                            self.microphone_name = default_device['name']
                            self.mic_var.set(self.microphone_name)
                        else:
                            # Si le micro par défaut n'est pas dans la liste filtrée, prendre le premier
                            self.selected_device_index = input_devices[0]['index']
                            self.microphone_name = input_devices[0]['name']
                            self.mic_var.set(self.microphone_name)
                    else:
                        # Pas de micro par défaut, prendre le premier
                        self.selected_device_index = input_devices[0]['index']
                        self.microphone_name = input_devices[0]['name']
                        self.mic_var.set(self.microphone_name)
                except:
                    # En cas d'erreur, prendre le premier
                    self.selected_device_index = input_devices[0]['index']
                    self.microphone_name = input_devices[0]['name']
                    self.mic_var.set(self.microphone_name)
                
                # Mettre à jour l'affichage
                self.mic_display_label.config(text=self.microphone_name)
        except Exception as e:
            print(f"Erreur chargement microphones: {e}")
    
    def _change_microphone(self, device_index, device_name):
        """Change le microphone sélectionné"""
        try:
            self.selected_device_index = device_index
            self.microphone_name = device_name
            self.mic_var.set(device_name)
            self.mic_display_label.config(text=device_name)
            print(f"Microphone changé: {device_name} (index: {device_index})")
        except Exception as e:
            print(f"Erreur changement microphone: {e}")
    
    def _on_duration_change(self, value):
        """Appelé quand le curseur change"""
        # Arrondir à la demi-minute la plus proche (snap à 30 secondes)
        raw_value = int(float(value))
        # Arrondir à la demi-minute la plus proche
        rounded_value = round(raw_value / 30) * 30
        
        # Si la valeur a changé, mettre à jour le slider
        if rounded_value != self.max_recording_duration:
            self.max_recording_duration = rounded_value
            # Mettre à jour la valeur du slider sans déclencher le callback
            self.duration_var.set(rounded_value)
        
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
            self.start_recording()
        else:
            self.stop_recording()
    
    def play_sound(self, filename, volume=0.5):
        """Joue un son (WAV) avec volume réglable"""
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
            self.root.after(2000, lambda: self.status_label.config(text="", fg="gray"))
    
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
        
        # Arrêter l'enregistrement dans un thread séparé
        threading.Thread(target=self._process_recording, daemon=True).start()
    
    def _record_audio(self):
        """Enregistre l'audio dans un thread séparé (durée max configurable)"""
        try:
            # Enregistrer avec sounddevice
            # Utiliser le microphone sélectionné si disponible
            device = self.selected_device_index if self.selected_device_index is not None else None
            
            while self.is_recording and not self.cancel_requested:
                # Vérifier si on a dépassé la durée maximum
                if self.recording_start_time and (time.time() - self.recording_start_time) >= self.max_recording_duration:
                    # Arrêter automatiquement après la durée max
                    self.root.after(0, lambda: self._auto_stop_recording())
                    break
                
                # Enregistrer un chunk de 0.1 seconde
                chunk = sd.rec(
                    int(0.1 * self.RATE),
                    samplerate=self.RATE,
                    channels=self.CHANNELS,
                    dtype=self.DTYPE,
                    device=device
                )
                sd.wait()
                if not self.cancel_requested:
                    self.audio_frames.append(chunk.tobytes())
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erreur", f"Erreur d'enregistrement: {e}"))
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
            
            # Vérifier si l'annulation a été demandée avant de lancer l'API
            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return
            
            # Transcrit avec OpenAI Whisper
            self.root.after(0, lambda: self.status_label.config(text="🔄 Transcription en cours...", fg="blue"))
            
            # Note: Une fois l'appel API lancé, on ne peut pas l'annuler facilement
            # Mais on peut éviter de lancer l'appel si cancel_requested est True
            with open(temp_filename, 'rb') as audio_file:
                transcript = self.client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    language="fr"
                )
            
            # Vérifier à nouveau après l'appel
            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return
            
            # Vérifier à nouveau après l'appel (au cas où)
            if self.cancel_requested:
                os.unlink(temp_filename)
                self.is_transcribing = False
                return
            
            text = transcript.text
            
            # Nettoyer le fichier temporaire
            os.unlink(temp_filename)
            
            # Afficher le texte dans l'interface
            self.root.after(0, lambda: self._display_text(text))
            
        except Exception as e:
            self.is_transcribing = False
            error_msg = f"Erreur lors de la transcription: {e}"
            self.root.after(0, lambda: messagebox.showerror("Erreur", error_msg))
            self.root.after(0, lambda: self._reset_ui("Erreur"))
    
    def _display_text(self, text):
        """Affiche le texte transcrit et le copie dans le presse-papier"""
        # Ajouter le texte à la zone de texte
        self.text_area.insert(tk.END, text + "\n\n")
        self.text_area.see(tk.END)
        
        # Réinitialiser l'historique après une nouvelle transcription
        # (on ne veut pas restaurer un texte effacé avant une nouvelle transcription)
        if self.is_recording == False:  # Si on vient de finir l'enregistrement
            self.text_history = []
            self.history_index = -1
        
        # Copier automatiquement dans le presse-papier
        try:
            pyperclip.copy(text)
            
            # Essayer de coller dans le champ actif si possible
            self._paste_to_active_field(text)
            
            # Jouer le son de fin (une fois que le texte est copié)
            self.play_sound('end.wav')
            self.status_label.config(text="✅ Transcription terminée et copiée !", fg="green")
        except Exception as e:
            self.status_label.config(text="✅ Transcription terminée (copie échouée)", fg="orange")
    
    def _paste_to_active_field(self, text):
        """Colle le texte dans le champ actif (si possible)"""
        try:
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            pass
        
        # Réinitialiser l'interface
        self._reset_ui()
        
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
        
        if status_text == "Prêt":
            self.status_label.config(text="", fg="gray")
        else:
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
        self.root.after(3000, lambda: self.status_label.config(text="", fg="gray"))
    
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
            self.root.after(2000, lambda: self.status_label.config(text="", fg="gray"))
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
                self.root.after(2000, lambda: self.status_label.config(text="", fg="gray"))
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de copier: {e}")
        else:
            messagebox.showinfo("Info", "Aucun texte à copier")
    
    def setup_global_hotkey(self):
        """Configure le raccourci clavier global Ctrl+Alt+9 avec pynput"""
        # État des touches pour détecter la combinaison
        self.pressed_keys = set()
        
        def on_press(key):
            """Détecte quand une touche est pressée"""
            try:
                # Détecter Échap (seulement pendant l'enregistrement)
                if key == pynput_keyboard.Key.esc:
                    if self.is_recording:
                        self.root.after(0, self.cancel_recording)
                    return  # Ne pas traiter Échap comme une autre touche
                
                # Gérer les touches spéciales (Ctrl, Alt)
                if key == pynput_keyboard.Key.ctrl_l or key == pynput_keyboard.Key.ctrl_r:
                    self.pressed_keys.add('ctrl')
                elif key == pynput_keyboard.Key.alt_l or key == pynput_keyboard.Key.alt_r:
                    self.pressed_keys.add('alt')
                else:
                    # Détecter le 9 (pavé numérique ou normal)
                    is_9 = False
                    
                    # Méthode 1: Vérifier via le virtual key code (pavé numérique)
                    # VK_NUMPAD9 = 0x69 (105 en décimal)
                    try:
                        # Sur Windows, on peut accéder au vk via l'attribut
                        vk = getattr(key, 'vk', None)
                        if vk == 105:  # VK_NUMPAD9
                            is_9 = True
                    except:
                        pass
                    
                    # Méthode 2: Vérifier le 9 normal
                    if not is_9:
                        try:
                            if hasattr(key, 'char') and key.char == '9':
                                is_9 = True
                        except:
                            pass
                    
                    # Méthode 3: Vérifier via la représentation (pour debug)
                    if not is_9:
                        try:
                            key_repr = repr(key)
                            # Le pavé numérique peut être représenté différemment
                            # On accepte aussi le 9 normal si Ctrl+Alt sont pressés
                            if '9' in key_repr or (hasattr(key, 'char') and key.char and '9' in str(key.char)):
                                # Si Ctrl+Alt sont déjà pressés, on considère que c'est le bon 9
                                if 'ctrl' in self.pressed_keys and 'alt' in self.pressed_keys:
                                    is_9 = True
                        except:
                            pass
                    
                    if is_9:
                        self.pressed_keys.add('9')
                        
                        # Vérifier si Ctrl+Alt+9 est pressé
                        if 'ctrl' in self.pressed_keys and 'alt' in self.pressed_keys:
                            # Déclencher l'action
                            self.root.after(0, self.toggle_recording)
                            # Nettoyer les touches pour éviter les répétitions
                            self.pressed_keys.discard('9')
                    elif hasattr(key, 'char') and key.char:
                        # Autre touche normale
                        self.pressed_keys.add(key.char.lower())
            
            except Exception as e:
                pass
        
        def on_release(key):
            """Détecte quand une touche est relâchée"""
            try:
                # Retirer la touche de l'ensemble
                if key == pynput_keyboard.Key.ctrl_l or key == pynput_keyboard.Key.ctrl_r:
                    self.pressed_keys.discard('ctrl')
                elif key == pynput_keyboard.Key.alt_l or key == pynput_keyboard.Key.alt_r:
                    self.pressed_keys.discard('alt')
                else:
                    # Retirer le 9 (pavé numérique ou normal)
                    try:
                        if hasattr(key, 'vk') and key.vk == 105:  # VK_NUMPAD9
                            self.pressed_keys.discard('9')
                    except:
                        pass
                    
                    if hasattr(key, 'char') and key.char == '9':
                        self.pressed_keys.discard('9')
                    elif hasattr(key, 'char') and key.char:
                        self.pressed_keys.discard(key.char.lower())
            
            except Exception:
                pass
        
        try:
            # Démarrer le listener dans un thread séparé
            def start_listener():
                with pynput_keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
                    listener.join()
            
            self.hotkey_thread = threading.Thread(target=start_listener, daemon=True)
            self.hotkey_thread.start()
            print("✅ Raccourci Ctrl+Alt+9 configuré (détection manuelle)")
        except Exception as e:
            print(f"⚠️ Impossible de configurer le raccourci: {e}")
            print("💡 Essayez de lancer l'application en tant qu'administrateur")
    

def main():
    root = tk.Tk()
    app = VoiceTranscriptionApp(root)
    
    # Fermeture directe sans message
    def on_closing():
        """Gestion de la fermeture"""
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Démarrer la boucle principale
    root.mainloop()

if __name__ == '__main__':
    main()
