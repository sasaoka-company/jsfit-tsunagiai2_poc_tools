@echo off
setlocal
set "RC=0"

cd /d %~dp0\..

if not exist "logs" mkdir "logs" >nul 2>&1

for /f %%i in ('powershell.exe -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"
set "RUNLOG=%CD%\logs\run_all_%TS%.log"

REM ---- create UTF-8 (with BOM) log file header ----
powershell.exe -NoProfile -Command ^
  "$p='%RUNLOG%';" ^
  "[IO.File]::WriteAllText($p, '', (New-Object System.Text.UTF8Encoding($true)))"

echo RUN_CMD_PATH=%~f0> "%RUNLOG%"
echo RUNLOG_PATH=%RUNLOG%>> "%RUNLOG%"
echo START_TS=%TS%>> "%RUNLOG%"
echo.>> "%RUNLOG%"

set "UV_CMD="
where uv >nul 2>&1
if errorlevel 1 (
  if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_CMD=%USERPROFILE%\.local\bin\uv.exe"
  if not defined UV_CMD if exist "%LOCALAPPDATA%\Programs\uv\uv.exe" set "UV_CMD=%LOCALAPPDATA%\Programs\uv\uv.exe"
  if not defined UV_CMD if exist "%LOCALAPPDATA%\uv\uv.exe" set "UV_CMD=%LOCALAPPDATA%\uv\uv.exe"
) else (
  set "UV_CMD=uv"
)

if not defined UV_CMD (
  echo ERROR: UV_NOT_FOUND>> "%RUNLOG%"
  set "RC=1"
  goto END
)

echo INFO: USING_UV="%UV_CMD%">> "%RUNLOG%"
"%UV_CMD%" --version>> "%RUNLOG%" 2>&1

if not exist "data\01_input" (
  echo ERROR: INPUT_FOLDER_NOT_FOUND data\01_input>> "%RUNLOG%"
  set "RC=1"
  goto END
)

echo INFO: UV_SYNC_START>> "%RUNLOG%"
if exist "uv.lock" (
  "%UV_CMD%" sync --frozen>> "%RUNLOG%" 2>&1
) else (
  "%UV_CMD%" sync>> "%RUNLOG%" 2>&1
)
if errorlevel 1 (
  echo ERROR: UV_SYNC_FAILED>> "%RUNLOG%"
  set "RC=1"
  goto END
)
echo INFO: UV_SYNC_OK>> "%RUNLOG%"

echo INFO: RUN_ALL_START>> "%RUNLOG%"
set PYTHONUNBUFFERED=1
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

"%UV_CMD%" run python -u run_all.py>> "%RUNLOG%" 2>&1
set "RC=%ERRORLEVEL%"

echo INFO: RUN_ALL_RC=%RC%>> "%RUNLOG%"

:END
echo.
echo DONE. RC=%RC%
echo LOG: %RUNLOG%
echo.
pause
exit /b %RC%
