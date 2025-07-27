@echo off
echo Starting Q3 Pacing Analysis...
echo.

REM Activate virtual environment and run the analysis
call .venv\Scripts\activate.bat && python pacing_analyzer.py

echo.
echo Analysis complete! Check the output folder for results.
pause
