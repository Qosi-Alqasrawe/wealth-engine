@echo off
cd /d "%~dp0"
echo Creating Anaconda environment: wealth-engine
conda create -n wealth-engine python=3.11 -y
call conda activate wealth-engine
pip install -r requirements.txt
echo.
echo Setup complete.
echo To run the app, double-click run_app.bat or run:
echo conda activate wealth-engine
echo streamlit run app.py
pause
