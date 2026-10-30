# Transcription Vocale — locale ou OpenAI

Petit outil Windows pour transcrire la voix en texte, sur l'ordinateur (modèles locaux, gratuits) ou via l'API OpenAI, déclenchable par raccourci clavier global et utilisable depuis n'importe quelle application. Le texte est collé directement dans le champ actif, la fenêtre peut tourner réduite dans la barre système.

## Installation

Prérequis : **Python 3.10 ou plus récent**, installé avec l'option « Add Python to PATH » ([python.org/downloads](https://www.python.org/downloads/)).

1. Lancer `setup.bat` (crée l'environnement virtuel et installe les dépendances, moteur local compris)
2. Lancer `start.bat`.

Le modèle par défaut, **Whisper large-v3 turbo local**, ne demande ni compte ni clé : à la première dictée, l'application propose de télécharger son fichier (≈ 890 Mo). Les autres modèles locaux se téléchargent de la même façon quand on les choisit. Rien d'autre à installer : le moteur arrive avec `setup.bat`, et la carte graphique est utilisée par son pilote habituel (Vulkan, présent avec les pilotes NVIDIA, AMD et Intel récents) ; sans carte graphique compatible, le processeur prend le relais, plus lentement.

Pour les modèles OpenAI, la clé API se configure via l'interface (bouton **« Modifier la clé… »** dans le panneau d'options), validée auprès d'OpenAI puis sauvegardée dans `.env` local. Alternative : copier `.env.example` en `.env` et renseigner `OPENAI_API_KEY` à la main.

## Utilisation

- **Ctrl+Alt+9** par défaut (pavé numérique ou rangée du haut) : démarre / arrête l'enregistrement, depuis n'importe quelle application. Modifiable — voir ci-dessous.
- **Échap** pendant l'enregistrement : annule, aucun appel API, aucun coût
- **Ctrl+Z / Ctrl+Y** dans la fenêtre : défait / refait la dernière modification de la transcription (collage par-dessus, effacement, frappe), sans limite tant que l'appli reste ouverte

Une fois la transcription terminée, le texte est collé dans le champ actif et/ou copié dans le presse-papier selon les deux toggles indépendants (voir Options).

La croix de la fenêtre réduit dans la barre système (configurable). Clic droit sur l'icône rouge → **Afficher** ou **Quitter**.

### Changer le raccourci

Panneau d'options → **« Modifier le raccourci… »**, puis appuyez sur la combinaison voulue : maintenez **Ctrl, Alt, Maj ou Win** (une ou plusieurs) et appuyez sur la touche finale. L'ordre de pression n'a pas d'importance, et la touche est mémorisée par sa position physique — le raccourci reste donc le même en AZERTY comme en QWERTY. Le choix est conservé entre les sessions.

Une touche seule, sans Ctrl ni Alt, est acceptée mais se déclenchera aussi pendant que vous tapez du texte : `A` seul lancerait un enregistrement à chaque lettre A écrite. L'application le signale au moment du choix. Les touches de fonction (F1 à F12) et celles du pavé numérique n'ont pas ce problème.

Le raccourci est écouté au niveau du système, en amont des applications. Windows ne distingue pas une touche tapée au doigt d'une touche émise par un périphérique programmable : clavier, souris à macros, pédale ou StreamDeck déclenchent l'enregistrement de la même façon, sans configuration particulière.

## Fonctionnalités

### Transcription
- Neuf modèles au choix, persistés entre sessions :

  | Modèle | Facturation | Coût par minute |
  |---|---|---|
  | **GPT Transcribe** | durée | 0,0045 $ ferme |
  | **GPT-4o Transcribe** | tokens (2,50 $ / 10,00 $ par 1M) | ≈ 0,0033 $ |
  | **GPT-4o Mini Transcribe** | tokens (1,25 $ / 5,00 $ par 1M) | ≈ 0,0016 $ |
  | **Whisper-1** (ancienne génération) | durée | 0,006 $ ferme |
  | **Whisper large-v3 turbo** (défaut, recommandé), **Nemotron 3.5**, **Qwen3-ASR 1.7B**, **Granite Speech 4.1 2B**, **Parakeet TDT 0.6B v3** (locaux) | aucune, sur l'ordinateur | 0 $ |

  Les coûts par minute des modèles GPT-4o sont mesurés sur une minute de dictée
  française continue (≈ 600 tokens audio en entrée + ≈ 180 tokens de texte en sortie) ;
  ils varient avec la densité de parole. GPT Transcribe est plus cher que GPT-4o Mini
  mais moins cher que Whisper-1, en échange de la meilleure précision de la gamme.

  **Lequel choisir.** Observation d'usage en août 2026, sur de la dictée française :
  les modèles récents (GPT Transcribe, GPT-4o Transcribe) rendent un texte propre,
  débarrassé des hésitations et des reprises, mais coûtent plus cher. GPT-4o Mini
  Transcribe est le moins cher de la liste et transcrit très fidèlement, en gardant
  les hésitations et les répétitions telles quelles. Le choix dépend donc de ce que
  vous faites du texte : dictée à retravailler ensuite ou texte directement lisible.
- **Les modèles locaux** tournent avec le moteur de [Handy](https://handy.computer),
  transcribe.cpp : sur la carte graphique via Vulkan quand elle est disponible, sinon sur le
  processeur. Aucune clé API, aucun envoi de son. Un seul modèle est chargé à la fois : il l'est
  au démarrage quand il est sélectionné, et libéré après 10 minutes sans usage, comme dans Handy.

  | Modèle local | Fichier (Q8_0) | Mémoire graphique | Attente pour 60 s de dictée (RTX 3080) |
  |---|---|---|---|
  | Whisper large-v3 turbo (recommandé) | [≈ 890 Mo](https://huggingface.co/handy-computer/whisper-large-v3-turbo-gguf/resolve/main/whisper-large-v3-turbo-Q8_0.gguf) | ≈ 0,9 Go | 1,2 s |
  | Nemotron 3.5 ASR Streaming 0.6B | [≈ 720 Mo](https://huggingface.co/handy-computer/nemotron-3.5-asr-streaming-0.6b-gguf/resolve/main/nemotron-3.5-asr-streaming-0.6b-Q8_0.gguf) | ≈ 0,7 Go | quasi nulle, transcrit pendant l'enregistrement |
  | Qwen3-ASR 1.7B | [≈ 2,2 Go](https://huggingface.co/handy-computer/Qwen3-ASR-1.7B-gguf/resolve/main/Qwen3-ASR-1.7B-Q8_0.gguf) | ≈ 3,1 Go | 1,6 s |
  | Granite Speech 4.1 2B | [≈ 2,6 Go](https://huggingface.co/handy-computer/granite-speech-4.1-2b-gguf/resolve/main/granite-speech-4.1-2b-Q8_0.gguf) | ≈ 3,4 Go | 9,7 s ; 6 minutes de dictée au plus |
  | Parakeet TDT 0.6B v3 | [≈ 740 Mo](https://huggingface.co/handy-computer/parakeet-tdt-0.6b-v3-gguf/resolve/main/parakeet-tdt-0.6b-v3-Q8_0.gguf) | ≈ 0,7 Go | 0,6 s |

  Nemotron transcrit pendant que vous parlez : le texte arrive une fraction de seconde après
  l'arrêt, quelle que soit la durée de la dictée. Les autres transcrivent tout d'un bloc à
  l'arrêt.

  **Lesquels choisir.** Observation d'usage en octobre 2026, sur de la dictée française :
  Whisper large-v3 turbo est le plus rapide à l'usage, reconnaît bien les noms propres et les
  mots anglais, écrit les chiffres en chiffres et garde les hésitations. Qwen3-ASR est aussi
  juste, un peu plus lourd. Parakeet est rapide mais bascule parfois en anglais au milieu d'une
  dictée française. Granite est lent sur les longues dictées et écrit les chiffres en lettres.
  Nemotron écrit les chiffres en lettres et écorche davantage les noms propres.

  Un modèle absent est proposé au téléchargement quand on le choisit ou qu'on lance une dictée
  avec lui : il est enregistré dans un dossier `models` à côté de `app.py`. Un fichier déjà
  téléchargé par Handy, dans le cache Hugging Face (`~/.cache/huggingface/hub`), est utilisé
  tel quel.
- Les modèles `gpt-live-transcribe` et `gpt-realtime-whisper` ne sont pas proposés : ils
  n'existent que sur les sessions Realtime WebSocket et renvoient 404 sur
  `v1/audio/transcriptions`, l'endpoint fichier utilisé ici.
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
- **Copier dans le presse-papier** et **Coller dans le champ actif** sont deux options indépendantes. Coller sans copier est possible : le contenu d'origine du presse-papier, image ou fichiers compris, est restauré 0,3 seconde après le collage, sauf si autre chose y a été copié entre-temps.
- Le collage attend que les touches du raccourci soient relâchées : tenues, elles changeraient `Ctrl+V` en un autre raccourci. Au-delà de 2 secondes, le texte reste dans le presse-papier sans être collé.
- Détection automatique des terminaux natifs (Windows Terminal, cmd, PuTTY, mintty) → `Ctrl+Maj+V` au lieu de `Ctrl+V`
- Toggle manuel **« Coller pour terminal »** pour les terminaux intégrés (Cursor, VS Code…) qui partagent la classe de fenêtre de leur éditeur et ne sont pas détectables automatiquement

### Suivi des coûts
- Session courante (compteur + total en USD)
- Mois en cours (compteur + total en USD), avec un **`?`** qui détaille au survol le nombre
  de transcriptions, les minutes d'audio et le coût **par modèle**
- Bouton **Effacer l'historique des coûts** (avec confirmation) pour repartir de zéro
- Historique complet dans `transcription_history.json` (date, modèle, durée audio, coût)
- Coût **réel** et non estimé : chaque réponse de l'API contient un champ `usage` qui donne
  soit les secondes facturées (`type: "duration"`), soit les tokens consommés
  (`type: "tokens"`). Le montant est calculé à partir de ce champ ; l'estimation par minute
  ne sert que de repli si `usage` est absent. Le drapeau `cost_exact` de chaque entrée
  d'historique indique lequel des deux a été utilisé.

### Autres
- Raccourci global personnalisable par capture de frappe, persisté entre sessions
- Sélection du microphone avec filtrage des périphériques virtuels, persistance par nom (stable après re-branchement USB)
- Sons de feedback activables/désactivables (début, fin, erreur)
- Minimisation dans la barre système (tray) avec icône dédiée

## Dépendances

Installées par `setup.bat`. Principales :

- `openai`, `python-dotenv` — API OpenAI + gestion `.env`
- `transcribe-cpp`, `soxr` — modèle local (moteur de Handy) et rééchantillonnage du micro à 16 kHz
- `sounddevice`, `numpy` — capture audio
- `pynput` — raccourci clavier global
- `pyperclip`, `pyautogui` — presse-papier + injection de frappes
- `pygame` — sons de feedback
- `pystray`, `Pillow` — icône dans la barre système (optionnelles, l'app fonctionne sans mais la tray est alors désactivée)

Voir `requirements.txt` pour les versions.

## Fichiers

- `app.py` — application principale
- `local_engine.py` — transcription locale (modèles et moteur)
- `clipboard.py` — sauvegarde et restauration du presse-papier dans tous ses formats
- `models/` — fichiers des modèles locaux, téléchargés par l'application (non versionné)
- `.env` — clé API (non versionné)
- `sounds/` — sons de feedback
- `transcription_history.json` — historique des transcriptions (local, non versionné)
- `user_preferences.json` — préférences UI persistées (local, non versionné)
- `setup.bat`, `start.bat`, `start_console.bat` — scripts de gestion
- `LICENSE` — licence MIT

## Notes

Clé API OpenAI requise pour les modèles OpenAI, facturée à votre compte selon l'usage ; les modèles locaux s'en passent. Les tarifs ci-dessus sont indicatifs et peuvent évoluer — voir [platform.openai.com/pricing](https://platform.openai.com/pricing) pour les prix en vigueur.

## Contribuer

Les retours et les propositions sont les bienvenus : ouvrez une *issue* pour signaler un bug ou suggérer une idée, une *pull request* pour proposer directement une modification. Les forks sont libres.

Les contributions sont publiées sous la licence du projet.

Pour toute autre question : [contact@knotix.fr](mailto:contact@knotix.fr).

## Licence

[MIT](LICENSE) — libre d'utilisation, de modification, de redistribution et d'intégration, y compris commerciale. La seule obligation est de conserver la mention de copyright et le texte de la licence dans les copies.
