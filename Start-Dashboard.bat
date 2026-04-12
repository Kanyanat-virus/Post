@echo off
echo ========================================================
echo Starting Local Web Server for Customer Files Dashboard
echo ========================================================
echo.
echo Please wait, the dashboard will open in your browser automatically...
echo.

REM Try to run python http server
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Python found. Starting server...
    start http://localhost:8080
    python -m http.server 8080
    pause
    exit
)

REM Try Node.js npx serve
npx --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Node.js found. Starting server...
    start http://localhost:8080
    npx serve -p 8080
    pause
    exit
)

echo ERROR: Neither Python nor Node.js is installed on your machine!
echo Please install python or open the folder using VSCode "Live Server" extension.
pause
