# Transcription Vocale, locale ou OpenAI

Petit outil pour Windows et Linux qui transcrit la voix en texte, sur l'ordinateur (modèles locaux, gratuits) ou via l'API OpenAI. Un raccourci clavier global lance la dictée depuis n'importe quelle application, et le texte est collé directement dans le champ actif. La fenêtre peut tourner réduite dans la barre système.

## Installation

### Windows

Prérequis : **Python 3.10 ou plus récent**, installé avec l'option « Add Python to PATH » ([python.org/downloads](https://www.python.org/downloads/)).

1. Lancer `setup.bat` (crée l'environnement virtuel et installe les dépendances, moteur local compris)
2. Lancer `start.bat`

### Linux

Pour Debian, Ubuntu, Linux Mint, Kubuntu et leurs dérivés, en session **X11** ou **Wayland** : l'application choisit d'elle-même.

1. `./setup.sh` : installe les paquets du système (mot de passe administrateur demandé), crée l'environnement virtuel et ajoute **Transcription Vocale** au menu des applications
2. Sous Wayland, la première fois seulement : fermer la session et la rouvrir
3. Lancer depuis le menu, ou `./start.sh` (`./start_console.sh` pour voir les messages de diagnostic)

Sous Wayland, le système interdit à une application de lire le clavier des autres et de leur envoyer des touches. L'application passe donc par le noyau : `setup.sh` ajoute l'utilisateur au groupe `input` (lecture du clavier et de la souris) et ouvre `/dev/uinput` à la session (clavier virtuel qui envoie Ctrl+V). Le presse-papier passe par `wl-clipboard`. La détection des terminaux, qui collent avec Ctrl+Maj+V, demande KDE Plasma ; ailleurs, l'option manuelle « terminal » reste disponible.

Le modèle par défaut, **Whisper large-v3 turbo local**, ne demande ni compte ni clé : à la première dictée, l'application propose de télécharger son fichier (≈ 890 Mo). Les autres modèles locaux se téléchargent de la même façon quand on les choisit. La carte graphique est utilisée par son pilote habituel (Vulkan, présent avec les pilotes NVIDIA, AMD et Intel récents) ; sans carte graphique compatible, le processeur prend le relais, plus lentement.

Pour les modèles OpenAI, la clé API se configure dans l'interface (bouton **« Modifier la clé… »** du panneau d'options), validée auprès d'OpenAI puis sauvegardée dans `.env`. On peut aussi copier `.env.example` en `.env` et y renseigner `OPENAI_API_KEY`.

## Utilisation

- **Ctrl+Alt+9** par défaut (pavé numérique ou rangée du haut) : démarre et arrête l'enregistrement, depuis n'importe quelle application
- **Échap** pendant l'enregistrement : annule, sans appel ni coût
- **Ctrl+Z / Ctrl+Y** dans la fenêtre : défait et refait les modifications de la transcription, sans limite tant que l'application reste ouverte

Le texte transcrit est collé dans le champ actif et/ou copié dans le presse-papier, selon deux options indépendantes. La croix de la fenêtre réduit l'application dans la barre système (configurable) ; clic droit sur l'icône pour **Afficher** ou **Quitter**.

### Changer le raccourci

Panneau d'options → **« Modifier le raccourci… »**, puis appuyer sur la combinaison voulue : une ou plusieurs touches parmi **Ctrl, Alt, Maj ou Win** (Super sous Linux), puis la touche finale. La touche est mémorisée par sa position physique : le raccourci reste le même en AZERTY comme en QWERTY.

Une touche seule, sans Ctrl ni Alt, est acceptée, mais elle se déclenche aussi pendant la frappe ; l'application le signale. Les touches F1 à F12 et celles du pavé numérique n'ont pas ce problème.

Le raccourci est écouté au niveau du système : clavier, souris à macros, pédale ou StreamDeck le déclenchent de la même façon.

Sous Linux, le raccourci est en plus **réservé** à l'application : la touche lance la dictée sans parvenir à la fenêtre active (le 9 du pavé numérique ne s'écrit plus nulle part), et une touche seule devient utilisable. Un **bouton latéral de la souris** (arrière, avant ou suivants), seul ou avec Ctrl, Alt, Maj ou Super, peut aussi servir de raccourci, sans outil tiers ; il perd alors son effet habituel. Si une autre application ou un raccourci du système tient déjà la combinaison, l'application le signale.

## Modèles

Neuf modèles au choix, le dernier choisi est gardé d'une session à l'autre.

| Modèle | Où | Coût par minute | Avis à l'usage |
|---|---|---|---|
| **Whisper large-v3 turbo** (par défaut) | local | 0 $ | Excellent : rapide, noms propres et mots anglais justes, chiffres en chiffres, garde les hésitations. Au niveau de GPT-4o Mini Transcribe, gratuitement. |
| Qwen3-ASR 1.7B | local | 0 $ | Juste, ponctuation propre ; se trompe sur certains nombres, plus lourd à charger. |
| Parakeet TDT 0.6B v3 | local | 0 $ | Très rapide, correct en français, mais bascule parfois en anglais au milieu d'une dictée. |
| Nemotron 3.5 ASR Streaming 0.6B | local | 0 $ | Transcrit pendant l'enregistrement, mais peu satisfaisant après quelques semaines d'usage : chiffres en lettres, noms propres écorchés. |
| Granite Speech 4.1 2B | local | 0 $ | Lent sur les longues dictées, chiffres en lettres, erreurs fréquentes. |
| **GPT-4o Mini Transcribe** | OpenAI | ≈ 0,0016 $ | Excellent et le moins cher : transcription très fidèle, hésitations et répétitions comprises. |
| GPT-4o Transcribe | OpenAI | ≈ 0,0033 $ | Nettoie le texte : hésitations et reprises retirées, phrases standardisées. Utile pour un texte directement lisible, moins pour une dictée qu'on veut fidèle. |
| GPT Transcribe | OpenAI | 0,0045 $ | Même comportement que GPT-4o Transcribe. |
| Whisper-1 | OpenAI | 0,006 $ | Ancienne génération, la plus chère de la liste. |

Avis tirés d'essais à la voix sur de la dictée française, avec le même texte lu pour chaque modèle (octobre 2026, RTX 3080 pour les modèles locaux). Les coûts des modèles GPT-4o sont mesurés sur une minute de dictée continue et varient avec la densité de parole ; les tarifs OpenAI évoluent, voir [platform.openai.com/pricing](https://platform.openai.com/pricing).

### Modèles locaux

Ils tournent sur l'ordinateur avec le moteur [transcribe.cpp](https://github.com/handy-computer/transcribe.cpp), sans clé API ni envoi de son. Un seul modèle est chargé à la fois, au démarrage s'il est sélectionné, et libéré après 10 minutes sans usage.

| Modèle local | Fichier (Q8_0) | Mémoire graphique | Attente pour 60 s de dictée (RTX 3080) |
|---|---|---|---|
| Whisper large-v3 turbo | [≈ 890 Mo](https://huggingface.co/handy-computer/whisper-large-v3-turbo-gguf/resolve/main/whisper-large-v3-turbo-Q8_0.gguf) | ≈ 0,9 Go | 1,2 s |
| Qwen3-ASR 1.7B | [≈ 2,2 Go](https://huggingface.co/handy-computer/Qwen3-ASR-1.7B-gguf/resolve/main/Qwen3-ASR-1.7B-Q8_0.gguf) | ≈ 3,1 Go | 1,6 s |
| Parakeet TDT 0.6B v3 | [≈ 740 Mo](https://huggingface.co/handy-computer/parakeet-tdt-0.6b-v3-gguf/resolve/main/parakeet-tdt-0.6b-v3-Q8_0.gguf) | ≈ 0,7 Go | 0,6 s |
| Nemotron 3.5 ASR Streaming 0.6B | [≈ 720 Mo](https://huggingface.co/handy-computer/nemotron-3.5-asr-streaming-0.6b-gguf/resolve/main/nemotron-3.5-asr-streaming-0.6b-Q8_0.gguf) | ≈ 0,7 Go | quasi nulle |
| Granite Speech 4.1 2B | [≈ 2,6 Go](https://huggingface.co/handy-computer/granite-speech-4.1-2b-gguf/resolve/main/granite-speech-4.1-2b-Q8_0.gguf) | ≈ 3,4 Go | 9,7 s ; 6 minutes de dictée au plus |

Un modèle absent est proposé au téléchargement quand on le choisit ou qu'on lance une dictée avec lui ; il est enregistré dans le dossier `models` à côté de `app.py`. Un fichier déjà présent dans le cache Hugging Face (`~/.cache/huggingface/hub`) est utilisé tel quel.

## Fonctionnalités

### Overlay flottant
- Fenêtre translucide toujours au-dessus, avec minuteur pendant l'enregistrement et statut pendant la transcription
- Les clics la traversent et atteignent l'application derrière
- Couleur selon l'état, position mémorisée, visible même quand l'application est réduite ; désactivable

### Copie et collage
- **Copier** et **Coller dans le champ actif** sont indépendants. Coller sans copier est possible : le presse-papier d'origine, images et fichiers compris, est restauré 0,3 seconde après, sauf si autre chose y a été copié entre-temps. Sous Linux, il est rendu dans un seul format, le plus utile : l'image, sinon les fichiers, sinon le texte (le texte mis en forme revient en texte simple).
- Le collage attend que les touches du raccourci soient relâchées ; au-delà de 2 secondes, le texte reste dans le presse-papier sans être collé.
- Les terminaux sont détectés et reçoivent `Ctrl+Maj+V` : Windows Terminal, cmd, PuTTY et mintty sous Windows ; GNOME Terminal, Tilix, Terminator, Konsole, Kitty, Alacritty et les autres terminaux courants sous Linux ; l'option **« Coller pour terminal »** couvre les terminaux intégrés aux éditeurs (Cursor, VS Code…).

### Suivi des coûts
- Session et mois en cours, avec le détail par modèle au survol du **`?`**
- Coût réel, calculé à partir de l'usage renvoyé par l'API (secondes ou tokens facturés)
- Historique dans `transcription_history.json`, effaçable depuis l'application

### Autres
- Durée maximum d'enregistrement réglable de 5 secondes à 15 minutes
- Choix du microphone, gardé par son nom (stable après un rebranchement USB)
- Sons de début, de fin et d'erreur, désactivables
- Une seule instance : relancer l'application ramène la fenêtre existante

## Dépendances

Installées par `setup.bat` ou `setup.sh` ; versions dans `requirements.txt`.

- `openai`, `python-dotenv` : API OpenAI et fichier `.env`
- `transcribe-cpp`, `soxr` : modèles locaux et rééchantillonnage du micro à 16 kHz
- `sounddevice`, `numpy` : capture audio
- `pynput` : raccourci clavier global
- `pyperclip` : presse-papier ; `pyautogui` : frappes sous Windows
- Sous Linux, paquets du système installés par `setup.sh` : `xclip` (presse-papier), `python3-tk`, `libportaudio2`, `python3-gi` et AyatanaAppIndicator (icône de la barre système), pilotes Vulkan
- `pygame` : sons
- `pystray`, `Pillow` : icône de la barre système (facultatives)

## Fichiers

- `app.py` : application principale
- `local_engine.py` : transcription locale
- `plateforme/` : ce qui dépend du système, un module par système avec les mêmes fonctions (`windows.py`, `linux.py`) : clavier, presse-papier, fenêtre active, forme de l'overlay
- `models/` : fichiers des modèles locaux (non versionné)
- `.env` : clé API (non versionné)
- `sounds/` : sons
- `transcription_history.json`, `user_preferences.json` : historique et préférences (locaux, non versionnés)
- `setup.bat`, `start.bat`, `start_console.bat` : installation et lancement sous Windows
- `setup.sh`, `start.sh`, `start_console.sh` : installation et lancement sous Linux
- `LICENSE` : licence MIT

## Remerciements

À [Handy](https://handy.computer), qui m'a donné envie d'essayer les modèles locaux récents, et dont l'équipe développe transcribe.cpp et publie les fichiers GGUF utilisés ici.

## Contribuer

Les retours et les propositions sont les bienvenus : une *issue* pour signaler un bug ou suggérer une idée, une *pull request* pour proposer une modification. Les forks sont libres, et les contributions sont publiées sous la licence du projet.

Pour toute autre question : [contact@knotix.fr](mailto:contact@knotix.fr).

## Licence

[MIT](LICENSE) : libre d'utilisation, de modification, de redistribution et d'intégration, y compris commerciale. La seule obligation est de conserver la mention de copyright et le texte de la licence dans les copies.
