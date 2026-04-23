@echo off
setlocal
cd /d "%~dp0"

"%~dp0JCBudgeting.Server.exe"
set EXITCODE=%ERRORLEVEL%

if not "%EXITCODE%"=="0" (
  echo.
  echo JCBudgeting Server exited with code %EXITCODE%.
  echo Press any key to close this window.
  pause >nul
)

endlocal
