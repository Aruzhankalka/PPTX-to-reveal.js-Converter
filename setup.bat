@echo off

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed. Get Node 18 LTS from nodejs.org and run this again.
  pause
  exit /b 1
)

for /f "delims=" %%v in ('node --version') do set NODE_VERSION=%%v
echo Node %NODE_VERSION% - ok

echo Installing dependencies...
call npm install
if errorlevel 1 (
  echo Install failed.
  pause
  exit /b 1
)

set /p ANSWER="Install Playwright browsers too? (only needed for the browser tests) [y/N] "
if /i "%ANSWER%"=="y" call npx playwright install chromium firefox

echo.
echo Done. Start the app with: npm run dev
echo Then open http://localhost:3000
echo Sample file to try: tests\fixtures\sample.pptx
echo.

set /p START="Start it now? [Y/n] "
if /i "%START%"=="n" exit /b 0

call npm run dev
