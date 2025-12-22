@echo off
setlocal enabledelayedexpansion

REM =========================================================
REM run.cmd（uv + run_all.py 起動）
REM - ログファイルを 1 本に集約：logs\run_all_YYYYMMDD_HHMMSS.log
REM - uv / python の標準出力・標準エラーをすべてログに保存
REM - 文字化け回避のため、ログに出すメッセージは ASCII（英数字）中心にする
REM   （第三者配布で環境差が出やすいため）
REM =========================================================

cd /d %~dp0\..

REM ---- logs フォルダ作成 ----
if not exist "logs" mkdir "logs" >nul 2>&1

REM ---- タイムスタンプ（ロケール依存回避）----
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%i"
set "RUNLOG=%CD%\logs\run_all_%TS%.log"

REM ---- 実行した run.cmd のパスとログパスを記録（混乱防止）----
echo RUN_CMD_PATH=%~f0> "%RUNLOG%"
echo RUNLOG_PATH=%RUNLOG%>> "%RUNLOG%"
echo START_TS=%TS%>> "%RUNLOG%"
echo.>> "%RUNLOG%"

REM ---- uv コマンド決定（PATH → 代表パス探索）----
set "UV_CMD="
where uv >nul 2>&1
if %errorlevel%==0 (
  set "UV_CMD=uv"
) else (
  if exist "%USERPROFILE%\.local\bin\uv.exe" set "UV_CMD=%USERPROFILE%\.local\bin\uv.exe"
  if not defined UV_CMD if exist "%LOCALAPPDATA%\Programs\uv\uv.exe" set "UV_CMD=%LOCALAPPDATA%\Programs\uv\uv.exe"
  if not defined UV_CMD if exist "%LOCALAPPDATA%\uv\uv.exe" set "UV_CMD=%LOCALAPPDATA%\uv\uv.exe"
)

if not defined UV_CMD (
  echo ERROR: UV_NOT_FOUND>> "%RUNLOG%"
  echo ERROR: uv not found. See log: %RUNLOG%
  exit /b 1
)

echo INFO: USING_UV="%UV_CMD%">> "%RUNLOG%"
"%UV_CMD%" --version>> "%RUNLOG%" 2>&1

REM ---- 入力フォルダ存在チェック ----
if not exist "data\01_input" (
  echo ERROR: INPUT_FOLDER_NOT_FOUND data\01_input>> "%RUNLOG%"
  echo ERROR: input folder not found. See log: %RUNLOG%
  exit /b 1
)

REM ---- 依存関係同期（uv sync）----
echo INFO: UV_SYNC_START>> "%RUNLOG%"
if exist "uv.lock" (
  "%UV_CMD%" sync --frozen>> "%RUNLOG%" 2>&1
) else (
  "%UV_CMD%" sync>> "%RUNLOG%" 2>&1
)
if errorlevel 1 (
  echo ERROR: UV_SYNC_FAILED>> "%RUNLOG%"
  echo ERROR: uv sync failed. See log: %RUNLOG%
  exit /b 1
)
echo INFO: UV_SYNC_OK>> "%RUNLOG%"

REM ---- パイプライン実行（標準出力を確実にログへ。-u でバッファ抑止）----
echo INFO: RUN_ALL_START>> "%RUNLOG%"
set PYTHONUNBUFFERED=1
"%UV_CMD%" run python -u run_all.py>> "%RUNLOG%" 2>&1
set "RC=%ERRORLEVEL%"

echo INFO: RUN_ALL_RC=%RC%>> "%RUNLOG%"

if not "%RC%"=="0" (
  echo ERROR: RUN_ALL_FAILED RC=%RC%>> "%RUNLOG%"
  echo ERROR: run_all failed (RC=%RC%). See log: %RUNLOG%
  exit /b %RC%
)

echo OK: COMPLETED>> "%RUNLOG%"
echo OK: completed. Log: %RUNLOG%
exit /b 0
