@echo off
setlocal
cd /d "%~dp0"

rem ============================================================
rem  OutlookTrayKeeper build script
rem  Compiles the utility with the in-box .NET Framework
rem  compiler (csc.exe). No downloads, no Visual Studio needed.
rem  Works on Windows 10/11 out of the box.
rem ============================================================

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo [ERROR] .NET Framework 4.x compiler csc.exe not found.
    echo Windows 10/11 normally include it. Install ".NET Framework 4.8"
    echo runtime from microsoft.com and run this script again.
    pause
    exit /b 1
)

echo Compiling OutlookTrayKeeper.exe ...
"%CSC%" /nologo /target:winexe /platform:anycpu /out:"OutlookTrayKeeper.exe" /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll "OutlookTrayKeeper.cs"
if errorlevel 1 (
    echo.
    echo [ERROR] Build failed - see the messages above.
    pause
    exit /b 1
)

echo.
echo [OK] Built: OutlookTrayKeeper.exe
echo.
echo Next steps:
echo   1. Double-click OutlookTrayKeeper.exe.
echo   2. On first run it adds itself to the current user's autostart.
echo   3. Click [X] in Outlook - it will minimize to the tray.
echo      Full exit: Outlook menu "File - Exit", or the utility's
echo      tray icon menu.
echo.
pause
