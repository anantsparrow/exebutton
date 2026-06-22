@echo off
echo =======================================================
echo Building Phishing Reporter Add-In MSI Setup Installer
echo =======================================================

:: Locate WiX toolset compiler bin path
set WIX_BIN="C:\Program Files (x86)\WiX Toolset v3.11\bin"
if not exist %WIX_BIN%\candle.exe set WIX_BIN="C:\Program Files\WiX Toolset v3.11\bin"

if not exist %WIX_BIN%\candle.exe (
    echo [ERROR] WiX Toolset v3.11 was not found in standard paths.
    echo Please install it from https://wixtoolset.org/ and verify path, or add it to system PATH.
    pause
    exit /b 1
)

echo.
echo [1/2] Compiling Setup.wxs...
%WIX_BIN%\candle.exe Setup.wxs -o Setup.wixobj
if %errorlevel% neq 0 (
    echo [ERROR] WiX Compilation failed!
    pause
    exit /b %errorlevel%
)

echo.
echo [2/2] Linking Setup.wixobj into MSI Installer...
%WIX_BIN%\light.exe Setup.wixobj -out PhishingReporterSetup.msi
if %errorlevel% neq 0 (
    echo [ERROR] WiX Linking failed!
    pause
    exit /b %errorlevel%
)

echo.
echo =======================================================
echo Success! MSI Package generated: PhishingReporterSetup.msi
echo =======================================================
pause
exit /b 0
