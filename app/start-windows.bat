@echo off
title ASECNA - Service Budget et Facturation
color 0A

echo.
echo ╔════════════════════════════════════════════╗
echo ║   ASECNA - Service Budget et Facturation   ║
echo ╚════════════════════════════════════════════╝
echo.
echo Demarrage du serveur...
echo.

cd /d "%~dp0"

REM Vérifier si Node.js est installé
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERREUR] Node.js n'est pas installe sur cet ordinateur.
    echo.
    echo Veuillez telecharger et installer Node.js depuis:
    echo https://nodejs.org/
    echo.
    pause
    exit /b 1
)

REM Vérifier si les dépendances sont installées
if not exist "node_modules\" (
    echo Installation des dependances (premiere execution)...
    call npm install
    if %ERRORLEVEL% NEQ 0 (
        echo [ERREUR] L'installation a echoue.
        pause
        exit /b 1
    )
)

REM Vérifier si le build existe
if not exist "dist\index.html" (
    echo Build de l'application...
    call npm run build
    if %ERRORLEVEL% NEQ 0 (
        echo [ERREUR] Le build a echoue.
        pause
        exit /b 1
    )
    
    REM Copier les fichiers publics
    xcopy /Y /E "public\*" "dist\" >nul 2>nul
)

REM Démarrer le serveur
echo.
echo Le serveur va demarrer...
echo.
start http://localhost:3002
node server/standalone-commonjs.js

pause
