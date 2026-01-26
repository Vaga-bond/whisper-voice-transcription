# Transcription Vocale - Whisper API

Application simple pour transcrire la voix en texte avec l'API OpenAI Whisper.

## Installation

1. `setup.bat` - Installe les dépendances
2. Éditez `.env` et ajoutez votre clé API OpenAI
3. `start.bat` - Lance l'application

## Utilisation

- **Ctrl+Alt+9** (pavé numérique) : Déclenche l'enregistrement depuis n'importe où
- Le texte est transcrit, copié dans le presse-papier et collé automatiquement dans le champ actif
- **Ctrl+Z** : Restaure le texte effacé

## Fichiers

- `app.py` - Application principale
- `config.py` - Configuration
- `.env` - Clé API (ne pas versionner)
- `sounds/` - Sons de début/fin
