#!/usr/bin/env bash
# Installation sous Linux (Debian, Ubuntu, Linux Mint, Kubuntu et dérivés), session X11 ou
# Wayland.
set -euo pipefail
cd "$(dirname "$(readlink -f "$0")")"
APP_DIR="$(pwd)"

echo "Paquets du système (mot de passe administrateur demandé)..."
# python3-tk : interface ; libportaudio2 : micro ; xclip et wl-clipboard : presse-papier
# (X11 et Wayland) ; python3-dev et build-essential : compilation d'evdev, requis par pynput ;
# python3-gi et AyatanaAppIndicator : icône de la barre système ;
# libvulkan1 et mesa-vulkan-drivers : carte graphique pour les modèles locaux.
sudo apt-get update
sudo apt-get install -y python3-venv python3-tk python3-dev build-essential \
    libportaudio2 xclip wl-clipboard python3-gi gir1.2-ayatanaappindicator3-0.1 \
    libvulkan1 mesa-vulkan-drivers

# Sous Wayland, aucune application ne peut lire le clavier des autres ni leur envoyer de
# touches : l'application lit les périphériques du noyau (groupe « input ») et colle par un
# clavier virtuel (/dev/uinput, ouvert à l'utilisateur de la session).
NEW_SESSION_NEEDED=0
if [ "${XDG_SESSION_TYPE:-}" = "wayland" ]; then
    echo
    echo "Session Wayland : accès au clavier et clavier virtuel..."
    if ! id -nG "$USER" | tr ' ' '\n' | grep -qx input; then
        sudo usermod -aG input "$USER"
    fi
    # Le groupe ne vaut que pour les sessions ouvertes après l'ajout.
    if ! id -nG | tr ' ' '\n' | grep -qx input; then
        NEW_SESSION_NEEDED=1
    fi
    echo 'KERNEL=="uinput", SUBSYSTEM=="misc", OPTIONS+="static_node=uinput", TAG+="uaccess"' \
        | sudo tee /etc/udev/rules.d/70-transcription-vocale-uinput.rules >/dev/null
    echo uinput | sudo tee /etc/modules-load.d/transcription-vocale-uinput.conf >/dev/null
    sudo modprobe uinput
    sudo udevadm control --reload-rules
    sudo udevadm trigger --subsystem-match=misc --sysname-match=uinput
fi

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
echo "Lancement à l'ouverture de session, rangée dans la barre système..."
# Pour l'arrêter : Configuration du système > Démarrage automatique (KDE), ou supprimer ce
# fichier.
mkdir -p "$HOME/.config/autostart"
cat > "$HOME/.config/autostart/transcription-vocale.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Transcription Vocale
Comment=Dictée vocale, rangée dans la barre système
Exec="$APP_DIR/start.sh" --barre
Path=$APP_DIR
Icon=audio-input-microphone
Terminal=false
X-GNOME-Autostart-enabled=true
EOF

echo
echo "Installation terminée."
echo
if [ "$NEW_SESSION_NEEDED" = 1 ]; then
    echo "⚠️  Fermez la session et rouvrez-la avant le premier lancement : l'accès au"
    echo "   clavier (groupe « input ») ne compte qu'à partir de la prochaine session."
    echo
fi
echo "Lancement : « Transcription Vocale » dans le menu, ou ./start.sh"
echo "Avec les messages de diagnostic : ./start_console.sh"
echo
echo "Les modèles locaux fonctionnent sans clé : le modèle choisi se télécharge depuis"
echo "l'application à la première dictée. Pour les modèles OpenAI, renseignez la clé API"
echo "dans l'application (bouton « Modifier la clé… ») ou dans le fichier .env."
