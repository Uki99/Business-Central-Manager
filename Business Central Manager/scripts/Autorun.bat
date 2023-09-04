@echo off
set "psScript=Autorun.ps1"

:: ------------- This bat script delegates running of the whole application and should not be changed ------------- ::


:: Check if running as administrator
NET SESSION >NUL 2>&1
if %errorLevel% == 0 (
    :: Running as administrator, execute PowerShell script
    powershell -File "%~dp0%psScript%"
) else (
    :: Not running as administrator, elevate the script
    echo Running script as administrator...
    powershell -Command "Start-Process -FilePath '%0' -ArgumentList '%psScript%' -Verb RunAs"
)