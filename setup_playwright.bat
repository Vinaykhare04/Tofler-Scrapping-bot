@echo off
REM Windows batch script to install Playwright browsers

echo =========================================
echo Installing Playwright browsers...
echo =========================================

python -m playwright install chromium

echo =========================================
echo Setup complete!
echo =========================================
pause
