@echo off
cd /d "%~dp0"

echo Starting backend on http://127.0.0.1:8888
start "TodSaMong Backend" cmd /k python -m backend.main

timeout /t 4 /nobreak >nul

echo Starting frontend on http://127.0.0.1:8501
start "TodSaMong Frontend" cmd /k python -m streamlit run frontend/app.py --server.port 8501

timeout /t 4 /nobreak >nul
start "" "http://127.0.0.1:8501"
