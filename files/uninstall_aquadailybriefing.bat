@echo off
TITLE Aqua Daily Briefing Uninstaller
echo ==================================================
echo      Uninstalling Aqua Daily Briefing...
echo ==================================================
echo.

REM 1. Remove Startup Shortcut
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\AquaDailyBriefing.lnk" (
    del /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\AquaDailyBriefing.lnk"
    echo [OK] Startup shortcut removed.
) else (
    echo [INFO] Startup shortcut not found.
)

REM 2. Remove Desktop Shortcut
if exist "%USERPROFILE%\Desktop\AquaDailyBriefing.lnk" (
    del /q "%USERPROFILE%\Desktop\AquaDailyBriefing.lnk"
    echo [OK] Desktop shortcut removed.
) else (
    REM Check OneDrive Desktop as Windows often moves it there
    if exist "%OneDrive%\Desktop\AquaDailyBriefing.lnk" (
        del /q "%OneDrive%\Desktop\AquaDailyBriefing.lnk"
        echo [OK] Desktop shortcut removed from OneDrive.
    ) else (
        echo [INFO] Desktop shortcut not found.
    )
)

REM 3. Remove Application Folder
if exist "%USERPROFILE%\AquaBriefing" (
    rmdir /s /q "%USERPROFILE%\AquaBriefing"
    echo [OK] AquaBriefing folder deleted.
) else (
    echo [INFO] AquaBriefing folder not found.
)

echo.
echo ==================================================
echo      Uninstallation Completed Successfully!
echo ==================================================
echo.
pause