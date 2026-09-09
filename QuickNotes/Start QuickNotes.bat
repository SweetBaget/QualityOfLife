@echo off
rem ============================================================
rem  QuickNotes launcher.
rem  Runs the app without keeping a console window open.
rem  Requires: Windows 10/11 (built-in PowerShell 5.1) - no installs needed.
rem ============================================================
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0QuickNotes.ps1"
