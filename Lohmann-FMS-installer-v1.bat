@echo off
setlocal enabledelayedexpansion
title Lohmann FMS Installer

:: Step 1: Ensure Admin Privileges
:: -------------------------------
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo =====================================================
echo        Lohmann FMS - Automated Installer
echo =====================================================
echo.

:: Step 2: Install Node.js via winget (ignore if already installed)
:: ---------------------------------------------------------------
echo ====================================================
echo Checking if Node.js is already installed...
echo ====================================================

node -v >nul 2>&1
if %errorlevel% equ 0 (
    echo Node.js is already installed.
    goto verify_install
)

echo.
echo Node.js not found. Proceeding with installation...
echo ====================================================
echo Checking for winget installation...
echo ====================================================

:: Check if winget exists
winget --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Winget is not available on this system.
    echo Attempting to install Node.js manually...
    goto manual_install
)

echo.
echo Winget found. Installing Node.js LTS via winget...
winget source update
winget install --id OpenJS.NodeJS -e --version 24.10.0 -h --accept-source-agreements --accept-package-agreements

if %errorlevel% neq 0 (
    echo.
    echo Winget failed to install Node.js. Falling back to manual installation...
    goto manual_install
)

goto update_node_paths

:manual_install
echo.
echo ====================================================
echo Detecting system architecture...
echo ====================================================
set "ARCH="
if "%PROCESSOR_ARCHITECTURE%"=="AMD64" (
    set "ARCH=x64"
) else (
    set "ARCH=x86"
)

echo Architecture detected: %ARCH%

set "URL_64=https://nodejs.org/dist/v24.10.0/node-v24.10.0-x64.msi"
@REM set "URL_32=https://nodejs.org/dist/v24.10.0/node-v24.10.0-x86.msi"

if "%ARCH%"=="x64" (
    set "NODE_URL=%URL_64%"
) else (
    echo 32-bit windows is not supported.
    pause
    exit /b 1
)

set "INSTALLER=%TEMP%\node_installer.msi"

echo.
echo ====================================================
echo Downloading Node.js installer...
echo ====================================================
powershell -Command ^ "try { Start-BitsTransfer -Source '%NODE_URL%' -Destination '%INSTALLER%' -ErrorAction Stop } catch { exit 1 }"
if %errorlevel% neq 0 (
    echo.
    echo Failed to download Node.js installer.
    echo Please check your internet connection or try manually downloading from:
    echo %NODE_URL%
    pause
    exit /b 1
)

echo.
echo ====================================================
echo Running Node.js installer...
echo ====================================================
start /wait msiexec /i "%INSTALLER%" 

echo.
echo Installer finished. Verifying installation...

goto update_node_paths

:update_node_paths

echo Updating Node.js environment paths...
set "PATH=C:\Program Files\nodejs;%PATH%"

:verify_install
node -v >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ====================================================
    echo ERROR: Node.js installation failed or not found.
    echo ====================================================
    pause
    exit /b 1
)

for /f "delims=" %%v in ('node -v') do set NODE_VERSION=%%v
echo.
echo ====================================================
echo SUCCESS! Node.js is available.
echo Version: %NODE_VERSION%
echo ====================================================

:: Step 3: Download latest release zip from GitHub (dynamic)
:: ---------------------------------------------------------
set "REPO_OWNER=AstroSyncDev"
set "REPO_NAME=lohmann-fms-binaries"
set "ZIP_PATH=%TEMP%\lohmann-fms.zip"
set "EXTRACT_PATH=%TEMP%\lohmann-fms"

echo Cleaning up any previous downloads...
if exist "%ZIP_PATH%" (
    del /f /q "%ZIP_PATH%"
    echo Removed old zip file: %ZIP_PATH%
)

if exist "%EXTRACT_PATH%" (
    rmdir /s /q "%EXTRACT_PATH%"
    echo Removed old folder: %EXTRACT_PATH%
)

echo Fetching latest release info...
for /f "usebackq tokens=*" %%i in (`powershell -Command ^
    "(Invoke-RestMethod https://api.github.com/repos/%REPO_OWNER%/%REPO_NAME%/releases/latest).assets | Where-Object { $_.browser_download_url -like '*.zip' } | Select-Object -ExpandProperty browser_download_url"`) do set "DOWNLOAD_URL=%%i"

if not defined DOWNLOAD_URL (
    echo Could not find any ZIP assets in latest release.
    pause
    exit /b
)

echo Downloading latest release from !DOWNLOAD_URL! ...
powershell -Command ^ "try { Start-BitsTransfer -Source '!DOWNLOAD_URL!' -Destination '%ZIP_PATH%' -ErrorAction Stop } catch { exit 1 }"
if %errorLevel% neq 0 (
    echo Failed to download release. Check your internet connection or the repo URL.
    pause
    exit /b
)
echo Download complete.

:: Step 4: Unzip the release
:: -------------------------
echo Extracting the downloaded ZIP...
powershell -Command "Expand-Archive -Path '%ZIP_PATH%' -DestinationPath '%EXTRACT_PATH%' -Force"
if %errorLevel% neq 0 (
    echo Failed to extract the ZIP file.
    pause
    exit /b
)
echo Extraction complete.

:: Step 5: Copy extracted folder to Windows App directory
:: ------------------------------------------------------
set "APP_DIR=%LOCALAPPDATA%\Lohmann-FMS"
echo Copying files to %APP_DIR%...
if exist "%APP_DIR%" (
    echo Previous installation found. Removing old files...
    rmdir /s /q "%APP_DIR%"
)
mkdir "%APP_DIR%"
xcopy "%EXTRACT_PATH%\*" "%APP_DIR%\" /E /H /C /I /Y
echo Files copied successfully.

:: Step 6: Install dependencies
:: -----------------------------------
echo Installing production dependencies...
cd /d "%APP_DIR%"
call npm install --omit-dev
if %errorLevel% neq 0 (
    echo NPM install encountered warnings or errors.
    pause
    exit /b
)
echo Dependencies installed.

:: Step 7: Create startup scripts (VBS + BAT)
:: -------------------------------------------------

echo Creating startup scripts...

:: Enable delayed expansion
setlocal enabledelayedexpansion

:: 7A — Background Node runner (still needs a .bat internally)
(
    echo @echo off
    echo cd /d "%APP_DIR%"
    echo call npm run start
) > "%APP_DIR%\Lohmann-FMS-App-Run.bat"
echo Startup bat script created at "%APP_DIR%\Lohmann-FMS-App.bat"

:: 7B — VBS launcher that runs the BAT hidden
(
    echo Set shell = CreateObject("WScript.Shell"^)
    echo shell.Run "!APP_DIR!\Lohmann-FMS-App-Run.bat", 0, False
) > "%APP_DIR%\Lohmann-FMS-App.vbs"

echo Startup vbs script created at "%APP_DIR%\Lohmann-FMS-App.vbs"

:: Step 8: Create Desktop Shortcut
:: -------------------------------
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\Lohmann-FMS-App.lnk"
echo Creating desktop shortcut...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$s=(New-Object -COMObject WScript.Shell).CreateShortcut('%SHORTCUT_PATH%');" ^
    "$s.TargetPath='%APP_DIR%\Lohmann-FMS-App.vbs';" ^
    "$s.WorkingDirectory='%APP_DIR%';" ^
    "$s.IconLocation='%SystemRoot%\System32\shell32.dll,25';" ^
    "$s.Save();"
echo Shortcut created on desktop.

:: Step 9: Add to Windows Startup (Current User)
:: -----------------------------------------------
set "STARTUP_DIR=%ProgramData%\Microsoft\Windows\Start Menu\Programs\StartUp"
set "STARTUP_SHORTCUT_PATH=%STARTUP_DIR%\Lohmann-FMS-App.lnk"

echo Ensuring Startup folder shortcut...

:: If shortcut exists, delete it first
if exist "%STARTUP_SHORTCUT_PATH%" (
    echo Existing startup shortcut found. Removing...
    del /f /q "%STARTUP_SHORTCUT_PATH%"
)

echo Creating startup shortcut in "%STARTUP_DIR%"...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$s=(New-Object -COMObject WScript.Shell).CreateShortcut('%STARTUP_SHORTCUT_PATH%');" ^
    "$s.TargetPath='%APP_DIR%\Lohmann-FMS-App.vbs';" ^
    "$s.WorkingDirectory='%APP_DIR%';" ^
    "$s.IconLocation='%SystemRoot%\System32\shell32.dll,25';" ^
    "$s.Save();"
if exist "%STARTUP_SHORTCUT_PATH%" (
    echo Startup shortcut created.
) else (
    echo Failed to create startup shortcut. Check permissions.
)

:: Step 10: Run the app and exit
:: ----------------------------
echo Launching Lohmann FMS...
start "" "%APP_DIR%\Lohmann-FMS-App.vbs"

echo Installation complete!
exit /b 0