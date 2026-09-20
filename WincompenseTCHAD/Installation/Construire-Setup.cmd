@echo off
setlocal EnableDelayedExpansion

rem ============================================================================
rem   Wincompense TCHAD - fabrication du programme d'installation
rem ============================================================================
rem
rem   A QUOI SERT CE FICHIER
rem       Compiler Wincompense.iss sans ouvrir Inno Setup : un double-clic, et
rem       le setup apparait dans Installation\Sortie.
rem
rem   CE QU'IL FAUT AVANT
rem       1. la solution compilee en RELEASE dans Visual Studio
rem          (liste de la barre d'outils : Release, puis Generer > Generer la
rem          solution) ;
rem       2. Inno Setup 6 installe - https://jrsoftware.org/isdl.php
rem
rem   POUR COMPILER DEPUIS UN AUTRE EMPLACEMENT
rem       Construire-Setup.cmd "C:\chemin\vers\bin\Release"
rem
rem ============================================================================

cd /d "%~dp0"

echo.
echo   Wincompense TCHAD - fabrication du setup
echo   =======================================
echo.

rem --- 1. Le dossier Release ---------------------------------------------------
set "RELEASE=%~1"
if "%RELEASE%"=="" set "RELEASE=%~dp0..\WincompenseTCHAD\bin\Release"

if not exist "%RELEASE%\Wincompense.exe" (
    echo   [ARRET] Wincompense.exe est introuvable dans :
    echo           %RELEASE%
    echo.
    echo   La solution n'a pas ete compilee en Release.
    echo   Dans Visual Studio : choisir Release au lieu de Debug dans la liste
    echo   de la barre d'outils, puis Generer ^> Generer la solution.
    echo.
    pause
    exit /b 1
)

rem --- 2. Le compilateur Inno Setup --------------------------------------------
rem Cherche dans les deux Program Files, puis dans le PATH. La version 6 est
rem installee en 32 bits par defaut, d'ou l'ordre.
set "ISCC="
for %%C in (
    "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
    "%ProgramFiles%\Inno Setup 6\ISCC.exe"
    "%ProgramFiles(x86)%\Inno Setup 5\ISCC.exe"
) do if exist %%C set "ISCC=%%~C"

if "%ISCC%"=="" for %%C in (ISCC.exe) do set "ISCC=%%~$PATH:C"

if "%ISCC%"=="" (
    echo   [ARRET] ISCC.exe est introuvable : Inno Setup n'est pas installe.
    echo.
    echo   A telecharger sur https://jrsoftware.org/isdl.php
    echo.
    echo   ATTENTION A LA LICENCE : depuis la version 7, Inno Setup n'est plus
    echo   gratuit pour un usage commercial, et une banque l'est. Prendre la
    echo   6.4.x, gratuite pour tout usage. Ce script compile a l'identique.
    echo.
    pause
    exit /b 1
)

echo   Release  : %RELEASE%
echo   Inno     : %ISCC%
echo.

rem --- 3. Compilation -----------------------------------------------------------
"%ISCC%" /DDossierRelease="%RELEASE%" "Wincompense.iss"

if errorlevel 1 (
    echo.
    echo   [ECHEC] La compilation s'est arretee. Le message ci-dessus dit pourquoi.
    echo.
    pause
    exit /b 1
)

echo.
echo   [OK] Le setup est dans : %~dp0Sortie
echo.

if exist "%~dp0Sortie" start "" "%~dp0Sortie"

pause
