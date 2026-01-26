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
echo Configuration de l'API...
if not exist .env (
    echo Copie de .env.example vers .env...
    copy .env.example .env
    echo.
    echo ⚠️ IMPORTANT: Editez .env et ajoutez votre clé API OpenAI
    echo    Obtenez votre clé sur: https://platform.openai.com/api-keys
    echo.
) else (
    echo Fichier .env existe deja
)

echo.
echo ✅ Installation terminée !
echo.
echo Pour démarrer l'application:
echo   start.bat
echo   OU
echo   venv\Scripts\activate
echo   python app.py
echo.

pause
