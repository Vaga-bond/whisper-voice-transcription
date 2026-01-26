@echo off
REM ========================================
REM Mise à jour des dépendances
REM (Remplace keyboard par pynput)
REM ========================================

echo Activation de l'environnement virtuel...
call venv\Scripts\activate.bat

echo Desinstallation de keyboard (ancienne dependance)...
pip uninstall -y keyboard

echo Installation des nouvelles dependances...
pip install -r requirements.txt

echo.
echo ✅ Mise à jour terminée !
echo keyboard a ete supprime et remplace par pynput.
echo.

pause
