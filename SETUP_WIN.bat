@echo off
echo Configuration de vcpkg pour Visual Studio...
vcpkg integrate install
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERREUR: vcpkg n'est pas installe ou n'est pas dans le PATH.
    echo Veuillez l'installer depuis https://vcpkg.io/
    pause
    exit /b %ERRORLEVEL%
)
echo.
echo vcpkg est pret. Vous pouvez maintenant ouvrir LivretCompetences.sln
pause
