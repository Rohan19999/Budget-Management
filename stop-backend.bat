@echo off
title Stop Finance Backend API

echo Stopping Finance Backend API on port 5050...
echo.

powershell -NoProfile -Command "$connection = Get-NetTCPConnection -LocalPort 5050 -ErrorAction SilentlyContinue; if ($connection) { $connection | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }; Write-Host 'Backend stopped.' } else { Write-Host 'Backend is not running.' }"

echo.
pause
