@echo off
echo Starting SlideShift...
echo Open your browser at: http://localhost:8765
cd /d "%~dp0backend"
python -m uvicorn main:app --host 127.0.0.1 --port 8765
pause
