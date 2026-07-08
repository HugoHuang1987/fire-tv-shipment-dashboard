@echo off
chcp 65001 >nul
setlocal

set "PY=%FIRE_TV_PYTHON%"
if "%PY%"=="" set "PY=python"
set "NODE=%FIRE_TV_NODE%"
if "%NODE%"=="" set "NODE=node"
set "NODE_MODULES=%FIRE_TV_NODE_MODULES%"

"%PY%" --version >nul 2>nul
if errorlevel 1 (
  echo Missing Python runtime. Set FIRE_TV_PYTHON or add python to PATH.
  if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
  exit /b 1
)

"%NODE%" --version >nul 2>nul
if errorlevel 1 (
  echo Missing Node runtime. Set FIRE_TV_NODE or add node to PATH.
  if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
  exit /b 1
)

if not "%NODE_MODULES%"=="" if not exist "%~dp0node_modules" (
  mklink /J "%~dp0node_modules" "%NODE_MODULES%" >nul
  if errorlevel 1 (
    echo Failed to link local Node modules.
    if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
    exit /b 1
  )
)

set "RULE_DEST=%~1"
if "%RULE_DEST%"=="" (
  echo Shared root argument is required.
  if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
  exit /b 1
)
set "FIRE_TV_SHARED_ROOT=%RULE_DEST%"
set "FIRE_TV_RULE_PATH=%RULE_DEST%\Fire TV quantity - rule.xlsx"
set "FIRE_TV_CALCULATION_RULE_PATH=%RULE_DEST%\Fire TV quantity - calculation-rules.json"
set "FIRE_TV_RAW_PATH=%RULE_DEST%\Fire TV quantity - raw.xlsx"
set "FIRE_TV_PATCH_PATH=%RULE_DEST%\Fire TV quantity - patch.xlsx"
if "%FIRE_TV_DISABLE_PATCH%"=="" set "FIRE_TV_DISABLE_PATCH=1"

for /f %%I in ('powershell.exe -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "RUN_STAMP=%%I"
for /f %%I in ('powershell.exe -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "RUN_DAY=%%I"

for %%D in ("%RULE_DEST%\dashboard" "%RULE_DEST%\tools" "%RULE_DEST%\history\raw" "%RULE_DEST%\history\rules" "%RULE_DEST%\history\outputs" "%RULE_DEST%\logs" "%RULE_DEST%\logs\wiki-sync") do if not exist "%%~D" mkdir "%%~D"

if not exist "%RULE_DEST%\Fire TV quantity - calculation-rules.docx" (
  copy /Y "%~dp0..\fire_tv_rules\Fire TV quantity - calculation-rules.docx" "%RULE_DEST%\Fire TV quantity - calculation-rules.docx" >nul
  if errorlevel 1 (
    echo Failed to install calculation rules docx.
    if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
    exit /b 1
  )
)

if not exist "%RULE_DEST%\Fire TV quantity - calculation-rules.json" (
  copy /Y "%~dp0..\fire_tv_rules\Fire TV quantity - calculation-rules.json" "%RULE_DEST%\Fire TV quantity - calculation-rules.json" >nul
  if errorlevel 1 (
    echo Failed to install calculation rules json.
    if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
    exit /b 1
  )
)

if not exist "%RULE_DEST%\Fire TV quantity - raw.xlsx" (
  if exist "%RULE_DEST%\20260623.xlsx" copy /Y "%RULE_DEST%\20260623.xlsx" "%RULE_DEST%\Fire TV quantity - raw.xlsx" >nul
)

echo Checking wiki rule sync...
if /I not "%FIRE_TV_SKIP_WIKI%"=="1" (
  set "WIKI_CONSOLE_LOG=%TEMP%\fire_tv_wiki_sync_last.log"
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Sync-FireTvWikiRules.ps1" -RulePath "%FIRE_TV_RULE_PATH%" -RunRoot "%RULE_DEST%\logs\wiki-sync\runs" > "%WIKI_CONSOLE_LOG%" 2>&1
  if errorlevel 1 (
    echo Wiki sync failed. Continue with existing rule.
    echo See: %RULE_DEST%\Fire TV quantity - wiki-sync-alerts.xlsx
    echo Console log: %WIKI_CONSOLE_LOG%
    set "FIRE_TV_WIKI_WARNING=1"
  )
)

echo Building dashboard data...
"%PY%" "%~dp0build_data.py"
set "BUILD_STATUS=%ERRORLEVEL%"
if not "%BUILD_STATUS%"=="0" if not "%BUILD_STATUS%"=="2" (
  echo Build failed. Existing dashboard data was not overwritten.
  if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
  exit /b 1
)

echo Exporting workbook outputs...
del /Q "%~dp0.data-source-success" >nul 2>nul
"%NODE%" "%~dp0export_data_source.mjs"
if not exist "%~dp0.data-source-success" (
  echo Workbook export failed. Existing shared outputs were not overwritten.
  if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
  exit /b 1
)

if "%BUILD_STATUS%"=="2" (
  echo Blocking exceptions found. Shared formal outputs remain unchanged.
  echo See: %RULE_DEST%\Fire TV quantity - exceptions.xlsx
  >> "%RULE_DEST%\logs\refresh_%RUN_DAY%.log" echo [%DATE% %TIME%] BLOCKED
  if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
  exit /b 2
)

copy /Y "%~dp0index.html" "%RULE_DEST%\dashboard\index.html" >nul
copy /Y "%~dp0dashboard-data.json" "%RULE_DEST%\dashboard\dashboard-data.json" >nul
copy /Y "%~dp0dashboard-data.js" "%RULE_DEST%\dashboard\dashboard-data.js" >nul
if not exist "%RULE_DEST%\dashboard\assets" mkdir "%RULE_DEST%\dashboard\assets"
copy /Y "%~dp0assets\styles.css" "%RULE_DEST%\dashboard\assets\styles.css" >nul
copy /Y "%~dp0assets\app.js" "%RULE_DEST%\dashboard\assets\app.js" >nul
copy /Y "%~dp0assets\lucide.min.js" "%RULE_DEST%\dashboard\assets\lucide.min.js" >nul
copy /Y "%~dp0Fire TV quantity - file-guide.txt" "%RULE_DEST%\Fire TV quantity - file-guide.txt" >nul

set "ARCHIVE_DIR=%RULE_DEST%\history\outputs\%RUN_STAMP%"
mkdir "%ARCHIVE_DIR%" >nul 2>nul
copy /Y "%RULE_DEST%\Fire TV quantity - data.xlsx" "%ARCHIVE_DIR%\Fire TV quantity - data.xlsx" >nul
copy /Y "%RULE_DEST%\Fire TV quantity - exceptions.xlsx" "%ARCHIVE_DIR%\Fire TV quantity - exceptions.xlsx" >nul
copy /Y "%RULE_DEST%\Fire TV quantity.xlsx" "%ARCHIVE_DIR%\Fire TV quantity.xlsx" >nul
>> "%RULE_DEST%\logs\refresh_%RUN_DAY%.log" echo [%DATE% %TIME%] SUCCESS

echo Refresh completed.
if /I not "%FIRE_TV_NO_PAUSE%"=="1" pause
