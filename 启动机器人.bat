@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0RobotAdmin.exe" (
    echo RobotAdmin.exe was not found in this folder.
    pause
    exit /b 1
)

start "" "%~dp0RobotAdmin.exe"
exit /b 0
