@Echo Running the Bots Execution Report Extraction

SETLOCAL
set FILE_PATH=%~dp0
set SCRIPT_PATH=%FILE_PATH%cs_ber.py
uv run "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%typeb_ber.py
uv run "%SCRIPT_PATH%"
set SCRIPT_PATH=%FILE_PATH%lij_cbo_ber_email.py
uv run "%SCRIPT_PATH%"
ENDLOCAL

@Echo Process Completed.
pause