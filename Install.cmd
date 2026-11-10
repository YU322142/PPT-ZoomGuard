@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
set "LZG_RESULT=%errorlevel%"
if not "%LZG_RESULT%"=="0" echo Installation did not complete. See the message above.
pause
exit /b %LZG_RESULT%
