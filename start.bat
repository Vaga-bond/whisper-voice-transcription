@echo off
REM ========================================
REM Démarrage de l'application (sans console)
REM ========================================

REM Activation de l'environnement virtuel
call venv\Scripts\activate.bat

REM Lancer avec pythonw.exe (sans console) et fermer le batch immédiatement
start "" pythonw app.py
