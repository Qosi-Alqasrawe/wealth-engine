@echo off
cd /d "%~dp0"

echo ========================================
echo Wealth Engine - Starting...
echo ========================================

call conda activate wealth-engine

echo.
echo Checking required packages...
python -c "import streamlit, pandas, numpy, requests, optuna, openpyxl; print('All packages OK')" 2>nul

if errorlevel 1 (
    echo.
    echo Missing packages detected. Installing requirements...
    pip install -r requirements.txt
)

echo.
echo Starting Streamlit...
echo Open this link if it does not open automatically:
echo http://localhost:8501
echo.

streamlit run app.py

pause