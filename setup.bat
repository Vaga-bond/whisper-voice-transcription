@echo off
REM ========================================
REM Installation de l'application
REM ========================================

echo Creation de l'environnement virtuel...
python -m venv venv

echo Activation de l'environnement...
call venv\Scripts\activate.bat

echo Installation des dependances...
pip install --upgrade pip
pip install -r requirements.txt

echo.
echo Configuration de la cle API...
if not exist .env (
    echo Copie de .env.example vers .env...
    copy .env.example .env
    echo.
    echo Cle API OpenAI requise. Deux options :
    echo   1. La renseigner via l'interface de l'app ^(bouton "Modifier la cle..."^)
    echo   2. L'editer manuellement dans le fichier .env
    echo Obtenez votre cle sur: https://platform.openai.com/api-keys
    echo.
) else (
    echo Fichier .env existe deja
)

echo.
echo Installation terminee.
echo.
echo Pour demarrer l'application:
echo   start.bat
echo   OU
echo   venv\Scripts\activate
echo   python app.py
echo.

pause
