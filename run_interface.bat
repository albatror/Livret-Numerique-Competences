@echo off
setlocal

:: se place dans le dossier du batch
cd /d %~dp0

set VENV_DIR=.venv
set REQ=requirements.txt
set PY_SCRIPT=

:: trouve le premier .py du dossier (sauf ce batch)
for %%f in (*.py) do (
  if /I not "%%~nxf"=="%~nx0" (
    set "PY_SCRIPT=%%~nxf"
    goto :found_script
  )
)
:found_script

if "%PY_SCRIPT%"=="" (
  echo Aucun script Python trouve dans le dossier. Place ton .py ici ou edite le bat pour indiquer PY_SCRIPT.
  pause
  exit /b 1
)

if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo Creation d'un environnement virtuel dans %VENV_DIR% ...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Erreur : impossible de creer le venv. Verifie que "python" est dans le PATH.
        pause
        exit /b 1
    )
)

:: active le venv
call "%VENV_DIR%\Scripts\activate.bat"

:: Mettre a jour pip via python -m pip pour eviter le message "To modify pip..."
"%VENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip

:: installer requirements si existe
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
