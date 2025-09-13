@echo off
setlocal enabledelayedexpansion

REM === PARAMÈTRES ===
REM Python.org installateur standard - plus petit et inclut tkinter
set PYTHON_URL=https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe
set INSTALL_DIR=%~dp0PythonLocal
set REQUIREMENTS=%~dp0requirements.txt
set INTERFACE_SCRIPT=%~dp0Interface.py
set PYTHON_INSTALLER=%~dp0python_installer.exe

REM --- Vérification de la présence de Python système ---
set "PYTHON_EXE="
for %%P in (python python3) do (
    where %%P >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_EXE=%%P"
        goto python_found
    )
)

REM --- Python non trouvé, on installe la version locale ---
echo Python système non trouvé. Installation locale nécessaire.
goto install_python_local

:python_found
echo Python trouvé sur le système : %PYTHON_EXE%
goto post_python_check

:install_python_local
echo.
echo === Installation Python locale (sans droits admin) ===

REM --- Téléchargement de Python officiel ---
echo Téléchargement de Python 3.12.8...
powershell -NoProfile -Command "try { $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%' } catch { Write-Error $_.Exception.Message; exit 1 }"

if not exist "%PYTHON_INSTALLER%" (
    echo ERREUR : Échec du téléchargement
    pause
    exit /b 1
)

REM --- Installation Python en mode utilisateur ---
echo Installation de Python en mode utilisateur (inclut tkinter)...
"%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=0 Include_test=0 SimpleInstall=1 TargetDir="%INSTALL_DIR%"

REM --- Attendre la fin de l'installation ---
:wait_install
timeout /t 3 /nobreak >nul 2>&1
if not exist "%INSTALL_DIR%\python.exe" (
    echo Installation en cours...
    goto wait_install
)

REM --- Supprimer l'installateur ---
del "%PYTHON_INSTALLER%" >nul 2>&1

set PYTHON_EXE=%INSTALL_DIR%\python.exe
set PIP_EXE=%INSTALL_DIR%\Scripts\pip.exe

:post_python_check
REM --- Vérification de l'installation ---
if not exist "%PYTHON_EXE%" (
    echo ERREUR : Python non installé correctement
    pause
    exit /b 1
)

echo Installation Python terminée !

REM --- Test obligatoire de tkinter ---
echo Test de tkinter...
"%PYTHON_EXE%" -c "import sys; print('Python version:', sys.version)"
"%PYTHON_EXE%" -c "import tkinter; print('✓ tkinter est disponible et fonctionnel!')"
if errorlevel 1 (
    echo ✗ ERREUR : tkinter non disponible
    echo Tentative de réinstallation avec tkinter forcé...
    "%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=0 Include_tcltk=1 TargetDir="%INSTALL_DIR%"
    timeout /t 10 /nobreak >nul 2>&1
    "%PYTHON_EXE%" -c "import tkinter; print('✓ tkinter maintenant disponible!')"
    if errorlevel 1 (
        echo ✗ ERREUR CRITIQUE : Impossible d'installer tkinter
        pause
        exit /b 1
    )
)

REM --- Mise à jour pip ---
echo Mise à jour de pip...
"%PYTHON_EXE%" -m pip install --upgrade pip --quiet

REM --- Installation des dépendances ---
if exist "%REQUIREMENTS%" (
    echo Installation des dépendances depuis requirements.txt...
    "%PYTHON_EXE%" -m pip install -r "%REQUIREMENTS%"
) else (
    echo Aucun requirements.txt trouvé
)

REM --- Création des raccourcis ---
echo @echo off > python_local.bat
echo cd /d "%%~dp0" >> python_local.bat
echo "%PYTHON_EXE%" %%* >> python_local.bat

echo @echo off > run_interface.bat
echo cd /d "%%~dp0" >> run_interface.bat
echo "%PYTHON_EXE%" Interface.py >> run_interface.bat
echo pause >> run_interface.bat

REM --- Test final ---
echo.
echo === Vérifications finales ===
"%PYTHON_EXE%" -c "import sys, tkinter; print('✓ Python + tkinter installés avec succès'); print('Localisation:', sys.executable)"

REM --- Lancement de Interface.py ---
if exist "%INTERFACE_SCRIPT%" (
    echo.
    echo === Lancement de Interface.py ===
    "%PYTHON_EXE%" "%INTERFACE_SCRIPT%"
) else (
    echo.
    echo ATTENTION : Interface.py non trouvé dans le dossier %~dp0
    echo.
    echo Python installé localement dans : %INSTALL_DIR%
    echo Raccourcis créés :
    echo - python_local.bat : pour exécuter Python
    echo - run_interface.bat : pour lancer Interface.py
)

echo.
echo === Installation terminée ===
echo Python local (avec tkinter) : %INSTALL_DIR%
echo Raccourcis disponibles dans le dossier courant
pause
endlocal
