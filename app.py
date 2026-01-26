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
        self.audio_frames = []
        self.audio_stream = None
        self.audio = None
        self.recording_thread = None
        
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
        
        # Vérifier et afficher le microphone disponible
        try:
            # Obtenir l'index du périphérique d'entrée par défaut
            default_input_idx = sd.default.device[0]
            if default_input_idx is not None and default_input_idx >= 0:
                device_info = sd.query_devices(default_input_idx)
                self.microphone_name = device_info.get('name', 'Micro inconnu')
            else:
                # Essayer de trouver n'importe quel périphérique d'entrée
                devices = sd.query_devices()
                input_devices = [d for d in devices if d['max_input_channels'] > 0]
                if input_devices:
                    self.microphone_name = input_devices[0]['name']
                else:
                    self.microphone_name = "Aucun micro détecté"
        except Exception as e:
            self.microphone_name = f"Micro par défaut (erreur: {str(e)[:30]})"
        
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
        
        # Frame principal
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titre
        title_label = tk.Label(
            main_frame,
            text="🎙️ Transcription Vocale",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Bouton Enregistrer/Arrêter
        self.record_button = tk.Button(
            main_frame,
            text="🎤 Démarrer l'enregistrement",
            font=("Arial", 12),
            bg="#4CAF50",
            fg="white",
            activebackground="#45a049",
            command=self.toggle_recording,
            width=30,
            height=2
        )
        self.record_button.pack(pady=10)
        
        # Indicateur d'enregistrement
        self.status_label = tk.Label(
            main_frame,
            text="",
            font=("Arial", 10),
            fg="gray"
        )
        self.status_label.pack()
        
        # Affichage du microphone utilisé
        mic_name = getattr(self, 'microphone_name', 'Chargement...')
        self.mic_label = tk.Label(
            main_frame,
            text=f"🎤 Micro: {mic_name}",
            font=("Arial", 9),
            fg="gray"
        )
        self.mic_label.pack(pady=(0, 5))
        
        # Affichage du raccourci clavier
        self.hotkey_label = tk.Label(
            main_frame,
            text="⌨️ Raccourci: Ctrl+Alt+9 (pavé numérique) - Global",
            font=("Arial", 9),
            fg="blue"
        )
        self.hotkey_label.pack(pady=(0, 10))
        
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
            height=20,  # Augmenté de 15 à 20
            width=70    # Augmenté de 60 à 70
        )
        self.text_area.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # Raccourci Ctrl+Z pour undo (amélioré)
        self.text_area.bind('<Control-z>', self.undo_clear)
        self.text_area.bind('<Control-Z>', self.undo_clear)
        # Aussi depuis la fenêtre principale
        self.root.bind('<Control-z>', self.undo_clear)
        self.root.bind('<Control-Z>', self.undo_clear)
        
        # Boutons d'action
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        self.clear_button = tk.Button(
            button_frame,
            text="🗑️ Effacer",
            command=self.clear_text,
            width=15,
            bg="#f44336",  # Rouge
            fg="white",
            activebackground="#da190b",  # Rouge foncé au survol
            activeforeground="white"
        )
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        self.copy_button = tk.Button(
            button_frame,
            text="📋 Copier",
            command=self.copy_to_clipboard,
            width=15
        )
        self.copy_button.pack(side=tk.LEFT, padx=5)
    
    def toggle_recording(self):
        """Bascule entre démarrer et arrêter l'enregistrement"""
        if not self.is_recording:
            self.start_recording()
        else:
            self.stop_recording()
    
    def play_sound(self, filename):
        """Joue un son (WAV) avec volume réduit à 50%"""
        try:
            sound_path = os.path.join(self.sounds_dir, filename)
            # Essayer WAV d'abord
            if not os.path.exists(sound_path):
                # Essayer sans extension
                sound_path = sound_path.replace('.mp3', '.wav')
            
            if os.path.exists(sound_path):
                # Utiliser pygame pour contrôler le volume
                sound = pygame.mixer.Sound(sound_path)
                sound.set_volume(0.5)  # 50% du volume
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
        self.is_recording = True
        self.audio_frames = []
        
        # Jouer le son de début
        self.play_sound('start.wav')
        
        # Mettre à jour l'interface
        self.record_button.config(
            text="⏹️ Arrêter l'enregistrement",
            bg="#f44336",
            activebackground="#da190b"
        )
        self.status_label.config(text="🔴 Enregistrement en cours...", fg="red")
        
        # Démarrer l'enregistrement dans un thread séparé
        self.recording_thread = threading.Thread(target=self._record_audio, daemon=True)
        self.recording_thread.start()
    
    def stop_recording(self):
        """Arrête l'enregistrement et transcrit"""
        if not self.is_recording:
            return
        
        self.is_recording = False
        self.status_label.config(text="⏳ Traitement en cours...", fg="orange")
        self.record_button.config(state=tk.DISABLED)
        
        # Arrêter l'enregistrement dans un thread séparé
        threading.Thread(target=self._process_recording, daemon=True).start()
    
    def _record_audio(self):
        """Enregistre l'audio dans un thread séparé"""
        try:
            # Enregistrer avec sounddevice
            while self.is_recording:
                # Enregistrer un chunk de 0.1 seconde
                chunk = sd.rec(
                    int(0.1 * self.RATE),
                    samplerate=self.RATE,
                    channels=self.CHANNELS,
                    dtype=self.DTYPE
                )
                sd.wait()  # Attendre que l'enregistrement soit terminé
                self.audio_frames.append(chunk.tobytes())
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erreur", f"Erreur d'enregistrement: {e}"))
            self.is_recording = False
    
    def _process_recording(self):
        """Traite l'enregistrement et transcrit"""
        try:
            if not self.audio_frames:
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
            
            # Transcrit avec OpenAI Whisper
            self.root.after(0, lambda: self.status_label.config(text="🔄 Transcription en cours...", fg="blue"))
            
            with open(temp_filename, 'rb') as audio_file:
                transcript = self.client.audio.transcriptions.create(
                    model="whisper-1",
                    file=audio_file,
                    language="fr"  # Français par défaut, peut être changé
                )
            
            text = transcript.text
            
            # Nettoyer le fichier temporaire
            os.unlink(temp_filename)
            
            # Afficher le texte dans l'interface
            self.root.after(0, lambda: self._display_text(text))
            
        except Exception as e:
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
            # Attendre un peu pour s'assurer que la transcription est terminée
            import time
            time.sleep(0.2)
            
            # Simuler Ctrl+V pour coller dans le champ actif
            # pyautogui permet de simuler les raccourcis clavier
            pyautogui.hotkey('ctrl', 'v')
            
        except Exception as e:
            # Si ça échoue, ce n'est pas grave, le texte est déjà dans le presse-papier
            pass
        
        # Réinitialiser l'interface
        self._reset_ui()
    
    def _reset_ui(self, status_text="Prêt"):
        """Réinitialise l'interface après l'enregistrement"""
        self.record_button.config(
            text="🎤 Démarrer l'enregistrement",
            bg="#4CAF50",
            activebackground="#45a049",
            state=tk.NORMAL
        )
        if status_text == "Prêt":
            self.status_label.config(text="", fg="gray")
        else:
            self.status_label.config(text=status_text, fg="gray")
        
        self.audio_frames = []
    
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
