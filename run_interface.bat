@echo off
setlocal

:: Se place dans le dossier du batch
cd /d %~dp0

set VENV_DIR=.venv
set REQ=requirements.txt
set PY_SCRIPT=Interface.py

:: Vérifie que le fichier Interface.py existe
if not exist "%PY_SCRIPT%" (
    echo Le fichier %PY_SCRIPT% est introuvable dans ce dossier.
    pause
    exit /b 1
)

:: Crée l'environnement virtuel si besoin
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Creation d'un environnement virtuel dans %VENV_DIR% ...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Erreur : impossible de creer le venv. Verifie que "python" est dans le PATH.
        pause
        exit /b 1
    )
)

:: Active le venv
call "%VENV_DIR%\Scripts\activate.bat"

:: Met à jour pip
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip

:: Installe les dépendances si requirements.txt existe
if exist "%REQ%" (
    echo Installation des dependances depuis %REQ% ...
    "%VENV_DIR%\Scripts\python.exe" -m pip install -r "%REQ%"
)

echo Lancement de %PY_SCRIPT% ...
"%VENV_DIR%\Scripts\python.exe" "%PY_SCRIPT%"

if errorlevel 1 (
    echo Le script s'est termine avec une erreur.
    pause
)

endlocal
