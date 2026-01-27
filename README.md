# Transcription Vocale - Whisper API

Application simple pour transcrire la voix en texte avec l'API OpenAI Whisper.

## Installation

1. Lancez `setup.bat` pour créer l'environnement virtuel et installer les dépendances
2. Éditez `.env` et ajoutez votre clé API OpenAI (obtenez-la sur https://platform.openai.com/api-keys)
3. Lancez `start.bat` pour démarrer l'application

## Utilisation

- **Ctrl+Alt+9** (pavé numérique) : Déclenche l'enregistrement depuis n'importe où
- **Échap** : Annule l'enregistrement en cours
- Le texte est transcrit, copié dans le presse-papier et collé automatiquement dans le champ actif
- **Ctrl+Z** : Restaure le texte effacé

## Fonctionnalités

- Enregistrement vocal avec sélection du microphone
- Transcription via OpenAI Whisper API
- Copie automatique dans le presse-papier
- Collage automatique dans le champ actif
- Raccourci clavier global (Ctrl+Alt+9)
- Durée maximum d'enregistrement configurable (5 secondes à 15 minutes)
- Annulation possible avant l'appel API
- Sons de feedback (début, fin, erreur)

## Fichiers

- `app.py` - Application principale
- `.env` - Clé API (ne pas versionner, utilisez `.env.example` comme modèle)
- `sounds/` - Sons de feedback (début, fin, erreur)
- `requirements.txt` - Dépendances Python
