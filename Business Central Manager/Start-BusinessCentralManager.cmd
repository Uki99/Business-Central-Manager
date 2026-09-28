@echo off
setlocal DisableDelayedExpansion
set "BC_MANAGER_SCRIPT=%~dp0scripts\Start-BusinessCentralManager.ps1"
set "BC_MANAGER_SETTINGS=%~dp0data\settings.json"
set "BC_MANAGER_ARGUMENTS=%*"
powershell.exe -NoProfile -NonInteractive -Command ^
	"$hideConsole = $false; try { $hideConsole = [bool](Get-Content -LiteralPath $env:BC_MANAGER_SETTINGS -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop).settings.hidePowerShellConsole } catch {}; $windowStyle = if ($hideConsole) { 'Hidden' } else { 'Normal' }; $launchArgs = '-NoProfile -ExecutionPolicy Unrestricted -File ' + [char]34 + $env:BC_MANAGER_SCRIPT + [char]34 + ' ' + $env:BC_MANAGER_ARGUMENTS; Start-Process -FilePath 'powershell.exe' -ArgumentList $launchArgs -WindowStyle $windowStyle -ErrorAction Stop"
