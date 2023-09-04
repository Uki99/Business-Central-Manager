@echo off
setlocal enabledelayedexpansion

:: Get the path to the directory where this batch script is located
set "scriptDir=%~dp0"

:: Set the path to the JSON file relative to the batch script location
set "jsonFile=%scriptDir%..\data\settings.json"

:: Use PowerShell to read the HidePowerShellConsole setting from the JSON file
for /f %%a in ('powershell -command "(Get-Content \"%jsonFile%\" | ConvertFrom-Json).settings.HidePowerShellConsole"') do (
    set "HidePowerShellConsole=%%a"
)

:: Run PowerShell as administrator with the appropriate window style
if "%HidePowerShellConsole%"=="True" (
    :: Run PowerShell script in hidden mode with admin privileges
    powershell -ExecutionPolicy Unrestricted -WindowStyle Hidden -NoProfile -File "%scriptDir%BE-terna Business Central Manager.ps1" %*
) else (
    :: Run PowerShell script in windowed mode with admin privileges
    powershell -ExecutionPolicy Unrestricted -WindowStyle Normal -NoProfile -File "%scriptDir%BE-terna Business Central Manager.ps1" %*
)