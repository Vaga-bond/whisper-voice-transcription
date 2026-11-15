#!/usr/bin/env bash
# Installation sous Linux (Debian, Ubuntu, Linux Mint et dérivés), session X11.
set -euo pipefail
cd "$(dirname "$(readlink -f "$0")")"
APP_DIR="$(pwd)"

if [ "${XDG_SESSION_TYPE:-}" = "wayland" ]; then
    echo "⚠️  Session Wayland détectée : le raccourci global et le collage automatique"
    echo "   n'y fonctionnent pas. Choisissez une session X11 à l'écran de connexion."
    echo
fi

echo "Paquets du système (mot de passe administrateur demandé)..."
# python3-tk : interface ; libportaudio2 : micro ; xclip : presse-papier ;
# python3-dev et build-essential : compilation d'evdev, requis par pynput ;
# python3-gi et AyatanaAppIndicator : icône de la barre système ;
# libvulkan1 et mesa-vulkan-drivers : carte graphique pour les modèles locaux.
sudo apt-get update
sudo apt-get install -y python3-venv python3-tk python3-dev build-essential \
    libportaudio2 xclip python3-gi gir1.2-ayatanaappindicator3-0.1 \
    libvulkan1 mesa-vulkan-drivers

echo
echo "Environnement virtuel..."
# --system-site-packages : donne accès à python3-gi, qui ne s'installe pas avec pip.
python3 -m venv --system-site-packages venv
venv/bin/pip install --upgrade pip
venv/bin/pip install -r requirements.txt

if [ ! -f .env ]; then
    cp .env.example .env
fi

chmod +x start.sh start_console.sh

echo
echo "Raccourci dans le menu des applications..."
mkdir -p "$HOME/.local/share/applications"
cat > "$HOME/.local/share/applications/transcription-vocale.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Transcription Vocale
Comment=Dictée vocale, locale ou OpenAI
Exec="$APP_DIR/start.sh"
Path=$APP_DIR
Icon=audio-input-microphone
Terminal=false
Categories=Utility;Accessibility;
EOF

echo
echo "Installation terminée."
echo
echo "Lancement : « Transcription Vocale » dans le menu, ou ./start.sh"
echo "Avec les messages de diagnostic : ./start_console.sh"
echo
echo "Les modèles locaux fonctionnent sans clé : le modèle choisi se télécharge depuis"
echo "l'application à la première dictée. Pour les modèles OpenAI, renseignez la clé API"
echo "dans l'application (bouton « Modifier la clé… ») ou dans le fichier .env."
