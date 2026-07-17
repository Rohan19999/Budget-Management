@echo off
title Finance Backend API
cd /d "%~dp0backend"

echo Starting Finance Backend API...
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo Node.js is not installed or not available in PATH.
  echo Please install Node.js, then run this file again.
  echo.
  pause
  exit /b 1
)

set PORT=5050

powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort 5060 -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"
if not errorlevel 1 (
  echo Backend is already running at http://localhost:5060
  echo Keep using the project in your browser.
  echo.
  pause
  exit /b 0
)

echo Checking backend dependencies...
npm.cmd install
if errorlevel 1 (
  echo.
  echo Dependency installation failed.
  pause
  exit /b 1
)

echo Backend URL: http://localhost:5060
echo Keep this window open while using the project.
echo Press Ctrl+C to stop the backend.
echo.
node server.js
echo.
echo Backend stopped.
pause
