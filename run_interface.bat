@echo off
setlocal ENABLEDELAYEDEXPANSION

REM ---- CONFIGURATION ----
set "PYPORTABLE_URL=https://github.com/portapps/python-portable/releases/download/3.11.5-11/python-portable-win64-3.11.5-11.7z"
set "PYPORTABLE_DIR=%~dp0python-portable"
set "PYPORTABLE_EXE=%PYPORTABLE_DIR%\python.exe"
set "SEVENZIP_URL=https://www.7-zip.org/a/7zr.exe"
set "SEVENZIP_EXE=%~dp07zr.exe"

REM ---- 1. DETECTER PYTHON EXISTANT ----
where python > nul 2> nul
if %ERRORLEVEL%==0 (
    set "PYTHON_CMD=python"
    goto pipsetup
)

REM ---- 2. DETECTER PYTHON PORTABLE LOCAL ----
if exist "%PYPORTABLE_EXE%" (
    set "PYTHON_CMD=%PYPORTABLE_EXE%"
    goto pipsetup
)

REM ---- 3. TELECHARGER PYTHON PORTABLE ----
echo Python non detecte. Telechargement de Python Portable...
if exist python-portable.7z del /f /q python-portable.7z

powershell -Command "Invoke-WebRequest -Uri '%PYPORTABLE_URL%' -OutFile 'python-portable.7z'"
if not exist python-portable.7z (
    echo ECHEC du telechargement de Python portable.
    pause
    exit /b 1
)

REM ---- 4. TELECHARGER 7-Zip PORTABLE SI BESOIN ----
if not exist "%SEVENZIP_EXE%" (
    echo Telechargement de 7-Zip portable...
    powershell -Command "Invoke-WebRequest -Uri '%SEVENZIP_URL%' -OutFile '7zr.exe'"
    if not exist "%SEVENZIP_EXE%" (
        echo ECHEC du telechargement de 7-Zip.
        pause
        exit /b 1
    )
)

REM ---- 5. EXTRAIRE PYTHON PORTABLE ----
echo Extraction de Python portable...
"%SEVENZIP_EXE%" x python-portable.7z -opython-portable > nul
if not exist "%PYPORTABLE_EXE%" (
    echo ECHEC de l'extraction.
    pause
    exit /b 1
)
del python-portable.7z

set "PYTHON_CMD=%PYPORTABLE_EXE%"

:pipsetup
REM ---- 6. INSTALLATION DE PIP SI MANQUANT ----
%PYTHON_CMD% -m pip --version > nul 2> nul
if %ERRORLEVEL% NEQ 0 (
    echo Installation de pip...
    %PYTHON_CMD% -m ensurepip --upgrade
    %PYTHON_CMD% -m pip install --upgrade pip
)

REM ---- 7. INSTALLER LES DEPENDANCES ----
if exist requirements.txt (
    echo Installation des dependances Python...
    %PYTHON_CMD% -m pip install --user -r requirements.txt
) else (
    REM Adapter ici si besoin d'autres modules
    %PYTHON_CMD% -m pip install --user tk
)

REM ---- 8. LANCEMENT DE L'INTERFACE ----
echo Lancement de l'interface...
%PYTHON_CMD% Interface.py

endlocal
pause
