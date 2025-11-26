@echo off
setlocal enableextensions
REM Ensure config.json is present at project root for PyInstaller add-data
IF NOT EXIST "config.json" (
  IF EXIST "dist\config.json" (
    COPY /Y "dist\config.json" "config.json" >NUL
  ) ELSE (
    CALL :missing_config
    GOTO :eof
  )
)

GOTO :continue

:missing_config
echo [ERROR] config.json not found. Place config.json in the project root (same folder as this script) and rerun.
echo        Or keep it next to the built EXE after build.
pause
EXIT /B 1

:continue
REM Build script for creating Audentes Automation Tool executable
REM This script packages the Python application into a standalone .exe file

echo ========================================
echo Building Audentes Automation Tool
echo ========================================
echo.

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller==6.16.0
)

echo.
echo Building executable...
echo.

REM Build the executable
REM --onefile: Creates a single executable file
REM --noconsole: Hides console window (GUI app)
REM --name: Sets the output executable name
REM --add-data: Includes config.json with the executable
REM --hidden-import: Ensures these modules are included

pyinstaller --onefile ^
    --noconsole ^
    --name "AudentesAutomationTool" ^
    --add-data "config.json;." ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=selenium ^
    --hidden-import=tkinter ^
    --hidden-import=webdriver_manager.chrome ^
    --hidden-import=webdriver_manager.core ^
    main.py

if errorlevel 1 (
    echo.
    echo Build failed! Check the error messages above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build Successful!
echo ========================================
echo.
echo Executable location: dist\AudentesAutomationTool.exe
echo.
echo Next steps:
echo 1. Copy AudentesAutomationTool.exe to deployment folder
echo 2. Copy config.json to the same folder
echo 3. Test the executable
echo 4. Create deployment package with:
echo    - AudentesAutomationTool.exe
echo    - config.json
echo    - inputs/ (empty folder)
echo    - outputs/ (empty folder)
echo    - logs/ (empty folder)
echo.
pause

