@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt
python download_models.py
python main.py
pause
