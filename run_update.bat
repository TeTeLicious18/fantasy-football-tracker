@echo off
cd /d "%~dp0"
pip install xlsxwriter openpyxl requests beautifulsoup4 >nul 2>&1
python update_scores.py
