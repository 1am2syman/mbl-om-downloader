@echo off
echo ===================================================
echo   OM Downloader - Installation Script
echo ===================================================
echo.

echo [1/2] Installing Node.js dependencies...
call npm install

echo.
echo [2/2] Installing Playwright Edge dependencies...
call npx playwright install msedge

echo.
echo ===================================================
echo   Installation Complete!
echo   You can now run: node automate_om.js
echo ===================================================
pause
