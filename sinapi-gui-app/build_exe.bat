@echo off
REM Run this from the repository root (double-click or run from PowerShell/CMD).
REM It will change directory to the src folder, install PyInstaller if missing, and build a one-file exe.
SETLOCAL ENABLEDELAYEDEXPANSION
cd /d "%~dp0\src"
REM Use `python` from PATH if available, otherwise try common install location.
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
  set PY=python
) else (
  set PY=C:\\Users\\%USERNAME%\\AppData\\Local\\Programs\\Python\\Python313\\python.exe
)
echo Using Python: %PY%
echo Installing/ensuring PyInstaller...
"%PY%" -m pip install --upgrade pyinstaller
if %ERRORLEVEL% NEQ 0 (
  echo Failed to install PyInstaller. Aborting.
  pause
  exit /b 1
)













ENDLOCALpausedir "%~dp0src\dist\" /b
necho Build finished. Executable should be in "%~dp0src\dist\SinapiApp.exe")  exit /b 1  pause  echo PyInstaller build failed.if %ERRORLEVEL% NEQ 0 ("%PY%" -m PyInstaller --noconfirm --onefile --name SinapiApp --add-data "LINKS_SICRO.txt;." main.pyREM Include LINKS_SICRO.txt as data (destination .) so code reading resource_path can find it.necho Building executable with PyInstaller...