@echo off
pushd "%~dp0"
REM Launch STA + bypass so no per-file Unblock-File / execution-policy change is needed.
REM The script self-elevates (UAC) for admin + STA if required.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File ".\Session-Manager.ps1"
popd
