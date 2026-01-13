@echo off
setlocal enabledelayedexpansion
title StimuPop Portable Builder
color 0B

echo.
echo  ================================================
echo       StimuPop Portable Bundle Builder
echo  ================================================
echo.
echo  This script creates a portable distribution of
echo  StimuPop that can be shared with testers.
echo.
echo  Requirements:
echo    - Internet connection (to download Python)
echo    - ~500MB disk space
echo.
echo  ================================================
echo.
pause

set "BUILD_DIR=%~dp0dist"
set "PYTHON_VERSION=3.11.9"
set "PYTHON_EMBED_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-embed-amd64.zip"
set "GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py"

:: Create dist directory
echo [1/7] Creating build directory...
if exist "%BUILD_DIR%" (
    echo       Cleaning existing dist folder...
    rmdir /s /q "%BUILD_DIR%"
)
mkdir "%BUILD_DIR%"
mkdir "%BUILD_DIR%\python"

:: Download Python embeddable
echo.
echo [2/7] Downloading Python %PYTHON_VERSION% embeddable...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_EMBED_URL%' -OutFile '%BUILD_DIR%\python.zip'"
if errorlevel 1 (
    echo [ERROR] Failed to download Python. Check your internet connection.
    pause
    exit /b 1
)

:: Extract Python
echo.
echo [3/7] Extracting Python...
powershell -Command "Expand-Archive -Path '%BUILD_DIR%\python.zip' -DestinationPath '%BUILD_DIR%\python' -Force"
del "%BUILD_DIR%\python.zip"

:: Enable pip in embedded Python (modify python311._pth)
echo.
echo [4/7] Configuring Python for pip...
set "PTH_FILE=%BUILD_DIR%\python\python311._pth"
(
    echo python311.zip
    echo .
    echo ../src
    echo import site
) > "%PTH_FILE%"

:: Download and install pip
echo.
echo [5/7] Installing pip...
powershell -Command "Invoke-WebRequest -Uri '%GET_PIP_URL%' -OutFile '%BUILD_DIR%\python\get-pip.py'"
"%BUILD_DIR%\python\python.exe" "%BUILD_DIR%\python\get-pip.py" --no-warn-script-location
del "%BUILD_DIR%\python\get-pip.py"

:: Install dependencies
echo.
echo [6/7] Installing StimuPop dependencies...
"%BUILD_DIR%\python\python.exe" -m pip install --no-warn-script-location -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

:: Copy application files
echo.
echo [7/7] Copying application files...
copy "%~dp0app.py" "%BUILD_DIR%\"
copy "%~dp0requirements.txt" "%BUILD_DIR%\"
copy "%~dp0config.yaml" "%BUILD_DIR%\" 2>nul
xcopy /E /I /Y "%~dp0src" "%BUILD_DIR%\src"

:: Copy user guide if it exists
if exist "%~dp0StimuPop_User_Guide.docx" (
    copy "%~dp0StimuPop_User_Guide.docx" "%BUILD_DIR%\"
)

:: Create logs directory
mkdir "%BUILD_DIR%\logs" 2>nul

echo.
echo  ================================================
echo       BUILD COMPLETE!
echo  ================================================
echo.
echo  The portable bundle is ready in:
echo    %BUILD_DIR%
echo.
echo  To distribute:
echo    1. Zip the entire 'dist' folder
echo    2. Share the ZIP with your tester
echo    3. Tester extracts and runs StimuPop.bat
echo.
echo  Folder contents:
dir /b "%BUILD_DIR%"
echo.
echo  ================================================
pause
