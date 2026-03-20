@echo off
echo ======================================================
echo Build Livret Numérique des Compétences Pro (.EXE)
echo ======================================================

REM Check if python is installed
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Python not found. Please install Python 3.
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
python -m pip install -r requirements.txt

REM Build EXE
echo Building executable...
python -m PyInstaller --clean LivretCompetences.spec

echo.
echo ======================================================
echo Build Complete!
echo The executable can be found in the "dist" folder.
echo ======================================================
pause
