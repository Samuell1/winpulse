@echo off
:: WinPulse launcher
:: Run this from any terminal, or double-click it.
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0checkup.ps1" %*
pause
