@echo off
chcp 65001 > nul
cd /d "C:\Users\ASTHA BAHETI\OneDrive\Documents\Desktop\lab_task"
echo.
echo  Starting eCourts Crawler...
echo.
python crawler.py
echo.
echo  Script finished. Press any key to close.
pause
