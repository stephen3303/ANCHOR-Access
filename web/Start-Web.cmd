@echo off
rem Double-click to open the apps in the browser (starts this PC's server if needed).
rem Settings: web.config.json next to this file (copy web.config.example.json).
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0kit\Start-KitWeb.ps1" -Config "%~dp0web.config.json"
if errorlevel 1 (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0kit\Start-KitWeb.ps1" -Config "%~dp0web.config.json"
  pause
)
