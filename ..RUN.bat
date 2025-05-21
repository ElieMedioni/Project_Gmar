@echo off
cd /d "%~dp0"
call venv\Scripts\activate
streamlit run stream.py
start http://localhost:8501

