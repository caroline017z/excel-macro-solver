@echo off
title 38DN Macro Runner Dashboard
start http://localhost:8502
python -m streamlit run "%~dp0dashboard.py" --server.port 8502 --server.headless true
pause
