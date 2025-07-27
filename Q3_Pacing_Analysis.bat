@echo off
title Q3 Pacing Analysis Tool
echo.
echo ========================================
echo    Q3 Pacing Analysis Tool
echo ========================================
echo.
echo Starting analysis...
echo.

REM Change to the script directory
cd /d "%~dp0"

REM Activate virtual environment and run the analysis
call .venv\Scripts\activate.bat && python pacing_analyzer.py

echo.
echo ========================================
echo Analysis complete!
echo.
echo Check the 'output' folder for your results.
echo The file will be named: Q3_Pacing_Analysis_[date1]_[date2].xlsx
echo.
echo Press any key to close this window...
pause >nul
