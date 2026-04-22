# Transcription Vocale — OpenAI

Petit outil Windows pour transcrire la voix en texte via l'API OpenAI, déclenchable par raccourci clavier global et utilisable depuis n'importe quelle application. Le texte est collé directement dans le champ actif, la fenêtre peut tourner réduite dans la barre système.

## Installation

1. Lancer `setup.bat` (crée l'environnement virtuel et installe les dépendances)
2. Lancer `start.bat` — la clé API se configure via l'interface (bouton **« Modifier la clé… »** dans le panneau d'options), validée auprès d'OpenAI puis sauvegardée dans `.env` local. Alternative : copier `.env.example` en `.env` et renseigner `OPENAI_API_KEY` à la main.

## Utilisation

- **Ctrl+Alt+9** (pavé numérique ou rangée du haut) : démarre / arrête l'enregistrement, fonctionne depuis n'importe quelle application et avec les macros de souris qui synthétisent cette combinaison
- **Échap** pendant l'enregistrement : annule, aucun appel API, aucun coût
- **Ctrl+Z** dans la fenêtre : restaure le texte effacé

Une fois la transcription terminée, le texte est collé dans le champ actif et/ou copié dans le presse-papier selon les deux toggles indépendants (voir Options).

La croix de la fenêtre réduit dans la barre système (configurable). Clic droit sur l'icône rouge → **Afficher** ou **Quitter**.

## Fonctionnalités

### Transcription
- Trois modèles au choix, persistés entre sessions :
  - **GPT-4o Mini Transcribe** (défaut, ≈ 0,003 $/min)
  - **GPT-4o Transcribe** (≈ 0,006 $/min)
  - **Whisper-1** (≈ 0,006 $/min)
- Capture audio en continu via `sounddevice.InputStream` avec callback (pas de gap entre échantillons)
- Durée maximum configurable de 5 secondes à 15 minutes
- Annulation à tout moment avant la requête API → zéro facturation

### Overlay flottant
- Fenêtre translucide toujours au-dessus, avec minuteur live pendant l'enregistrement et statut pendant la transcription
- **Click-through** (`WS_EX_TRANSPARENT`) : les clics traversent l'overlay et atteignent l'application derrière
- Coins arrondis, couleur adaptée à l'état (rouge / bleu / vert / orange)
- Déplaçable via une poignée dans le coin, position mémorisée dans les préférences
- Visible même quand l'application est réduite dans la barre système
- Désactivable dans les options

### Copie et collage
- **Copier dans le presse-papier** et **Coller dans le champ actif** sont deux options indépendantes. Coller sans copier est possible : le contenu d'origine du presse-papier est restauré après le collage.
- Détection automatique des terminaux natifs (Windows Terminal, cmd, PuTTY, mintty) → `Ctrl+Maj+V` au lieu de `Ctrl+V`
- Toggle manuel **« Coller pour terminal »** pour les terminaux intégrés (Cursor, VS Code…) qui partagent la classe de fenêtre de leur éditeur et ne sont pas détectables automatiquement

### Suivi des coûts
- Session courante (compteur + total en USD)
- Mois en cours (compteur + total en USD)
- Historique complet dans `transcription_history.json` (date, modèle, durée audio, coût estimé)
- Calcul local basé sur la durée audio × tarif OpenAI du modèle utilisé

### Autres
- Sélection du microphone avec filtrage des périphériques virtuels, persistance par nom (stable après re-branchement USB)
- Sons de feedback activables/désactivables (début, fin, erreur)
- Minimisation dans la barre système (tray) avec icône dédiée

## Dépendances

Installées par `setup.bat`. Principales :

- `openai`, `python-dotenv` — API OpenAI + gestion `.env`
- `sounddevice`, `numpy` — capture audio
- `pynput` — raccourci clavier global
- `pyperclip`, `pyautogui` — presse-papier + injection de frappes
- `pygame` — sons de feedback
- `pystray`, `Pillow` — icône dans la barre système (optionnelles, l'app fonctionne sans mais la tray est alors désactivée)

Voir `requirements.txt` pour les versions.

## Fichiers

- `app.py` — application principale
- `.env` — clé API (non versionné)
- `sounds/` — sons de feedback
- `transcription_history.json` — historique des transcriptions (local, non versionné)
- `user_preferences.json` — préférences UI persistées (local, non versionné)
- `setup.bat`, `start.bat`, `start_console.bat`, `update_dependencies.bat` — scripts de gestion

## Notes

Clé API OpenAI requise, facturée à votre compte selon l'usage. Les tarifs ci-dessus sont indicatifs et peuvent évoluer — voir [platform.openai.com/pricing](https://platform.openai.com/pricing) pour les prix en vigueur.
