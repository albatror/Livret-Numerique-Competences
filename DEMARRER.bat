@echo off
setlocal enabledelayedexpansion

REM === PARAMÈTRES ===
set "PY_VER=3.11.0"
set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip"
set "WORKDIR=%~dp0"
set "PY_DIR=%WORKDIR%PythonEmbed"
set "PY_ZIP=%WORKDIR%python-%PY_VER%-embed-amd64.zip"
set "INTERFACE_SCRIPT=%WORKDIR%Interface.py"
set "REQ_FILE=%WORKDIR%requirements.txt"

echo.
echo ==============================
echo === Vérification de Python ===
echo ==============================

where python >nul 2>&1
if %errorlevel%==0 (
    echo Python systeme détecté.
    set "PYTHON=python"
) else (
    echo Python non détecté, utilisation de la version portable...
    if not exist "%PY_DIR%\python.exe" (
        echo Téléchargement de Python portable...
        if exist "%PY_ZIP%" del /f /q "%PY_ZIP%"
        curl -L "%PY_URL%" -o "%PY_ZIP%"
        if %errorlevel% neq 0 (
            echo ERREUR: Téléchargement Python impossible.
            pause
            exit /b 1
        )
        echo Décompression de Python portable...
        if exist "%PY_DIR%" rd /s /q "%PY_DIR%"
        mkdir "%PY_DIR%"
        tar -xf "%PY_ZIP%" -C "%PY_DIR%"
    ) else (
        echo Version portable déjà présente.
    )
    set "PYTHON=%PY_DIR%\python.exe"
)

echo.
echo ==============================
echo === Vérification de pip    ===
echo ==============================

"%PYTHON%" -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Installation de pip...
    "%PYTHON%" -m ensurepip --default-pip
    if %errorlevel% neq 0 (
        echo Téléchargement get-pip.py...
        curl -L https://bootstrap.pypa.io/get-pip.py -o "%WORKDIR%get-pip.py"
        "%PYTHON%" "%WORKDIR%get-pip.py"
        del /f /q "%WORKDIR%get-pip.py"
    )
)

echo Mise à jour pip/setuptools/wheel...
"%PYTHON%" -m pip install --upgrade pip setuptools wheel

if exist "%REQ_FILE%" (
    echo Installation des dépendances depuis requirements.txt...
    "%PYTHON%" -m pip install -r "%REQ_FILE%"
)

echo.
echo ==============================
echo === Vérification de Git/Curl===
echo ==============================

where git >nul 2>&1
if %errorlevel% neq 0 (
    echo Git n'est pas installe. Téléchargez manuellement la version portable.
) else (
    echo Git présent.
)

where curl >nul 2>&1
if %errorlevel% neq 0 (
    echo Curl non trouvé. Windows 10+ possède curl nativement, sinon téléchargez-le.
) else (
    echo Curl présent.
)

echo.
echo ==============================
echo === Lancement Interface.py ===
echo ==============================
"%PYTHON%" "%INTERFACE_SCRIPT%"

echo Nettoyage du zip...
if exist "%PY_ZIP%" del /f /q "%PY_ZIP%"

echo.
echo ==============================
echo === Terminé ! ===============
echo ==============================
pause
endlocal
