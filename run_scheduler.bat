@echo off
title REMU-CI VisionExtract - Envoi automatique
echo ============================================================
echo REMU-CI VisionExtract - Envoi automatique des rapports
echo Date: %date% %time%
echo ============================================================

cd /d "C:\Users\ASSISTANT IT\Downloads\projet_python\Vision Extract"

echo Lancement de l'extraction et de l'envoi...
python email_sender.py

if %errorlevel% equ 0 (
    echo [SUCCES] Envoi termine
) else (
    echo [ERREUR] Code: %errorlevel%
)

echo.
echo Termine: %date% %time%
echo ============================================================
pause