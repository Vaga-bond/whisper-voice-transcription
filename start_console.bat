@echo off
REM ========================================
REM Démarrage de l'application (avec console pour debug)
REM ========================================

echo Activation de l'environnement virtuel...
call venv\Scripts\activate.bat

echo Démarrage de l'application...
python app.py

pause
