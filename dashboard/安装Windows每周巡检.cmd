@echo off
chcp 65001 >nul
if "%FIRE_TV_SHARED_ROOT%"=="" (
  echo FIRE_TV_SHARED_ROOT is required.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\tools\Register-FireTvScheduledTask.ps1" -SharedRoot "%FIRE_TV_SHARED_ROOT%"
echo.
if errorlevel 1 (
  echo Windows scheduled task installation failed.
) else (
  echo Windows scheduled task installation succeeded.
)
pause
