# Management script container that has functionality for BC Container Helper module update #

Set-ExecutionPolicy Unrestricted
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Used to check for updates or install the BCContainerHelper module
function Update-BCContainerHelperModule {
    try {
        # Check if BCContainerHelper module is installed
        $installedModule = Get-InstalledModule -Name BcContainerHelper -ErrorAction SilentlyContinue

        if ($installedModule) {
            $newestVersion = Get-Module -ListAvailable -Name BcContainerHelper | Select-Object -ExpandProperty Version | Sort-Object -Descending | Select-Object -First 1

            if ($newestVersion -gt $installedModule.Version) {		
                $ConfirmModuleUpdate = [System.Windows.Forms.MessageBox]::Show("PowerShell module BCContainerHelper found. Do you want to update the module?", "Confirm Module Update", "YesNo", "Question")

                if ($ConfirmModuleUpdate -eq "No") {
                    return
                }

                Write-Host "Updating BCContainerHelper module. Please wait...`n"

                try {
                    Update-Module BcContainerHelper -ErrorAction Stop
                    Write-Host "Module BCContainerHelper successfully updated" -ForegroundColor Green
                } catch {
                    $errorMessage = $_.ToString()
                    Write-Host "Error occurred during module update:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
                    return
                }
            } elseif ($newestVersion -eq $installedModule.Version) {
                Write-Host "BCContainerHelper module is already up to date.`n" -ForegroundColor Green
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("PowerShell module BCContainerHelper not found. Press OK to install the required module now.", "BCContainerHelper Install", "OK", "Warning") | Out-Null

            Write-Host "Installing BCContainerHelper module. Please wait...`n"

            try {
                Install-Module BCContainerHelper -Force -ErrorAction Stop
                Write-Host "Module BCContainerHelper successfully installed`n" -ForegroundColor Green
            } catch {
                $errorMessage = $_.ToString()
                Write-Host "Error occurred during module installation:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
                $null = Read-Host
            }
        }
    } catch {
        $errorMessage = $_.ToString()
        Write-Host "An error occurred during BCContainerHelper module update:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
        $null = Read-Host
    }
}