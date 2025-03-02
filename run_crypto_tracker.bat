@echo off
echo Cryptocurrency Live Tracker System
echo ================================
echo.
echo This script will:
echo 1. Generate an initial analysis report (HTML)
echo 2. Start the live Excel tracker
echo.
echo Note: The Excel file will be updated every 5 minutes.
echo Press Ctrl+C to stop the program at any time.
echo.
pause

echo.
echo Generating initial analysis report...
python generate_report.py
echo.
echo Starting live Excel tracker...
echo.
python excel_live_update.py 