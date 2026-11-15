#!/usr/bin/env bash
# Démarrage sous Linux avec les messages dans le terminal, pour le diagnostic.
cd "$(dirname "$(readlink -f "$0")")"
exec venv/bin/python app.py
