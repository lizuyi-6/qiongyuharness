@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -3.12 -m venv .venv
  if errorlevel 1 goto :failed
)
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :failed
set "QO_DEMO=1"
set "QO_ENGINE=demo"
set "QO_DATA_DIR=data\demo"
set "QO_PUBLIC_ORIGIN=http://127.0.0.1:8765"
set "QO_DATABASE_URL="
set "QO_USERS_FILE="
set "QO_CLAMD="
set "QO_OIDC_ISSUER="
set "QO_OIDC_CLIENT_ID="
set "QO_OIDC_CLIENT_SECRET="
echo.
echo Open http://127.0.0.1:8765 in your browser or desktop client.
echo Press Ctrl+C to stop the local service.
".venv\Scripts\python.exe" -m uvicorn server.app:app --host 127.0.0.1 --port 8765 --no-access-log
if errorlevel 1 goto :failed
exit /b 0
:failed
echo Start failed. Check that Python 3.12 is installed and package downloads are reachable.
pause
exit /b 1
