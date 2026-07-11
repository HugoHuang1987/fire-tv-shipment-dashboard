@echo off
chcp 65001 >nul
if "%FIRE_TV_SHARED_ROOT%"=="" (
  echo FIRE_TV_SHARED_ROOT is required.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\tools\Update-FireTvTaskStatus.ps1" -SharedRoot "%FIRE_TV_SHARED_ROOT%" -Show
echo.
pause
