@echo off
setlocal
set "APP_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%APP_DIR%photo-book-runner.ps1" -InstructionPath "%APP_DIR%photo-book-instruction.json"
pause
endlocal
