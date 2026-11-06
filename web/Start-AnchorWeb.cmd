@echo off
rem Double-click to open ANCHOR in the browser (starts this PC's server if needed).
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Start-AnchorWeb.ps1"
if errorlevel 1 (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Start-AnchorWeb.ps1"
  pause
)
