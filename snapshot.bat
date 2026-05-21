@Echo Preparing the Daily Snapshot Emails

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%snapshot_main.py
uv run "%SCRIPT_PATH%"
ENDLOCAL