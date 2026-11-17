#!/usr/bin/env bash
# Démarrage sous Linux, détaché du terminal.
cd "$(dirname "$(readlink -f "$0")")"
# --barre : démarrer rangée dans la barre système (ouverture de session).
nohup venv/bin/python app.py "$@" >/dev/null 2>&1 &
