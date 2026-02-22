@echo off
setlocal enabledelayedexpansion

:: ============================================================================
:: Download 64-bit Office LTSC (2021 + 2024) using Office Deployment Tool
:: Run this on a Windows PC, then transfer the output folders to Linux
:: ============================================================================

set "WORK_DIR=%USERPROFILE%\Desktop\OfficeDownloads"
set "ODT_URL=https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
set "ODT_DIR=%WORK_DIR%\ODT"

echo ============================================================
echo  Office 64-bit LTSC Downloader for Wine
echo  Downloads Office 2021 LTSC and 2024 LTSC (64-bit)
echo ============================================================
echo.
echo Output will be saved to: %WORK_DIR%
echo.

:: Create working directories
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
if not exist "%ODT_DIR%" mkdir "%ODT_DIR%"

:: -------------------------------------------------------
:: Step 1: Download and extract ODT
:: -------------------------------------------------------
echo [1/5] Downloading Office Deployment Tool...
if not exist "%ODT_DIR%\setup.exe" (
    powershell -Command "Invoke-WebRequest -Uri '%ODT_URL%' -OutFile '%WORK_DIR%\odt_installer.exe'"
    if errorlevel 1 (
        echo ERROR: Failed to download ODT. Check your internet connection.
        pause
        exit /b 1
    )
    echo Extracting ODT...
    "%WORK_DIR%\odt_installer.exe" /quiet /extract:"%ODT_DIR%"
    timeout /t 5 /nobreak >nul
    if not exist "%ODT_DIR%\setup.exe" (
        echo ERROR: ODT extraction failed. Try running odt_installer.exe manually.
        pause
        exit /b 1
    )
    echo ODT extracted successfully.
) else (
    echo ODT already present, skipping download.
)
echo.

:: -------------------------------------------------------
:: Step 2: Create Office 2021 LTSC 64-bit config
:: -------------------------------------------------------
echo [2/5] Creating configuration files...

(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2021"^>
echo     ^<Product ID="ProPlus2021Volume"^>
echo       ^<Language ID="en-us" /^>
echo     ^</Product^>
echo   ^</Add^>
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
echo   ^<Logging Level="Standard" Path="%%temp%%" /^>
echo ^</Configuration^>
) > "%WORK_DIR%\config-2021-LTSC-64bit.xml"

:: -------------------------------------------------------
:: Step 3: Create Office 2024 LTSC 64-bit config
:: -------------------------------------------------------
(
echo ^<Configuration^>
echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^>
echo     ^<Product ID="ProPlus2024Volume"^>
echo       ^<Language ID="en-us" /^>
echo     ^</Product^>
echo   ^</Add^>
echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
echo   ^<Logging Level="Standard" Path="%%temp%%" /^>
echo ^</Configuration^>
) > "%WORK_DIR%\config-2024-LTSC-64bit.xml"

echo Configuration files created.
echo.

:: -------------------------------------------------------
:: Step 4: Download Office 2021 LTSC 64-bit source files
:: -------------------------------------------------------
echo [3/5] Downloading Office 2021 LTSC 64-bit...
echo       This may take a while depending on your connection.
echo.

if not exist "%WORK_DIR%\Office2021" mkdir "%WORK_DIR%\Office2021"
copy "%WORK_DIR%\config-2021-LTSC-64bit.xml" "%WORK_DIR%\Office2021\config.xml" >nul

"%ODT_DIR%\setup.exe" /download "%WORK_DIR%\Office2021\config.xml"
if errorlevel 1 (
    echo WARNING: Office 2021 download may have encountered errors.
    echo          Check %WORK_DIR%\Office2021\ for partial downloads.
) else (
    echo Office 2021 LTSC 64-bit download complete.
)
echo.

:: -------------------------------------------------------
:: Step 5: Download Office 2024 LTSC 64-bit source files
:: -------------------------------------------------------
echo [4/5] Downloading Office 2024 LTSC 64-bit...
echo       This may take a while depending on your connection.
echo.

if not exist "%WORK_DIR%\Office2024" mkdir "%WORK_DIR%\Office2024"
copy "%WORK_DIR%\config-2024-LTSC-64bit.xml" "%WORK_DIR%\Office2024\config.xml" >nul

"%ODT_DIR%\setup.exe" /download "%WORK_DIR%\Office2024\config.xml"
if errorlevel 1 (
    echo WARNING: Office 2024 download may have encountered errors.
    echo          Check %WORK_DIR%\Office2024\ for partial downloads.
) else (
    echo Office 2024 LTSC 64-bit download complete.
)
echo.

:: -------------------------------------------------------
:: Summary
:: -------------------------------------------------------
echo [5/5] Done!
echo ============================================================
echo.
echo Downloaded files are at: %WORK_DIR%
echo.
echo Contents:
echo   Office2021\    - Office 2021 LTSC 64-bit source files
echo   Office2024\    - Office 2024 LTSC 64-bit source files
echo   ODT\           - Office Deployment Tool (setup.exe)
echo.
echo NEXT STEPS:
echo   1. Copy the entire OfficeDownloads folder to your Linux PC
echo      (USB drive, network share, etc.)
echo   2. On Linux, use ODT setup.exe via Wine to install:
echo      wine setup.exe /configure config.xml
echo.
echo ============================================================
pause
