@echo off
title Biocare Dispatch Portal
echo ========================================================
echo   Starting Biocare Dispatch Portal (Local Server)
echo ========================================================
echo.
echo Connecting to database: Biocare Dispatch Tracker.xlsx
echo Starting web server at http://localhost:5000...
echo.
start "" http://localhost:5000
python server.py
pause
