@echo off
setlocal enabledelayedexpansion

:: Get the path to the directory where this batch script is located
set "scriptDir=%~dp0"

:: Set the path to the JSON file relative to the batch script location
set "jsonFile=%scriptDir%..\data\settings.json"

:: Use PowerShell to read the HidePowerShellConsole setting from the JSON file
for /f %%a in ('powershell -NoProfile -Command "(Get-Content -LiteralPath $env:jsonFile -Raw | ConvertFrom-Json).settings.hidePowerShellConsole"') do (
    set "hidePowerShellConsole=%%a"
)

:: Run PowerShell as administrator with the appropriate window style
if "%hidePowerShellConsole%"=="True" (
    :: Run PowerShell script in hidden mode with admin privileges
    powershell -ExecutionPolicy Unrestricted -WindowStyle Hidden -NoProfile -File "%scriptDir%Start-BusinessCentralManager.ps1" %*
) else (
    :: Run PowerShell script in windowed mode with admin privileges
    powershell -ExecutionPolicy Unrestricted -WindowStyle Normal -NoProfile -File "%scriptDir%Start-BusinessCentralManager.ps1" %*
)