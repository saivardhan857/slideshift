@echo off
:loop
echo Starting SlideShift...
echo Open your browser at: http://localhost:8765
echo (Close this window to stop the server)
echo.
cd /d "%~dp0backend"
python -m uvicorn main:app --host 127.0.0.1 --port 8765 >> "%~dp0slideshift.log" 2>&1
echo.
echo Server stopped (exit code: %errorlevel%). Restarting in 5 seconds...
echo Check slideshift.log for details.
timeout /t 5 /nobreak >nul
goto loop
