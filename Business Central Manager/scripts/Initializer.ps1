

                                                                                    ### Setup Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Set-ExecutionPolicy Unrestricted
Add-Type -AssemblyName System.Windows.Forms

# Load settings from json
try {
    $settings = Get-Content (($PSScriptRoot | Split-Path) + "\data\settings.json") -Raw | ConvertFrom-Json -ErrorAction Stop
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}


# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                  ### Function Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# Used to update/install the required module BcContainer helper neccesary for application
function Update-BcContainerHelper {
    if (-not $settings.settings.SearchForUpdateBcContainerHelper) {
        return
    }

    Write-Host "Checking for BcContainerHelper module updates. Please wait...`n"

    # Check if BCContainerHelper module is installed
    $installedModule = Get-InstalledModule -Name BcContainerHelper -ErrorAction SilentlyContinue

    if ($installedModule) {
        $newestVersion = Get-Module -ListAvailable -Name BcContainerHelper | Select-Object -ExpandProperty Version | Sort-Object -Descending | Select-Object -First 1

        if ($newestVersion -gt $installedModule.Version) {		
            $ConfirmModuleUpdate = [System.Windows.Forms.MessageBox]::Show("PowerShell module BCContainerHelper found. Do you want to update the module?", "Confirm Module Update", "YesNo", "Question") | Out-Null      
            if ($ConfirmModuleUpdate -eq "No") {
                return
            }
            
            Write-Host "Updating BCContainerHelper module. Please wait...`n" -ForegroundColor Green
            
            try {
                Update-Module BcContainerHelper -ErrorAction Stop
            } catch {
                $errorMessage = $_.ToString()
                Write-Host "Error occurred during module update:`n$errorMessage"
                Exit 1
            }

            [System.Windows.Forms.MessageBox]::Show("Module BCContainerHelper successfully updated!", "BcContainerHelper Update", "OK", "Asterisk") | Out-Null
        } elseif ($newestVersion -eq $installedModule.Version) {
            Write-Host "BCContainerHelper module is already up to date.`n" -ForegroundColor Green
            return
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("PowerShell module BCContainerHelper not found. Press OK to install the required module now...", "BCContainerHelper Install", "OK", "Warning") | Out-Null
        Write-Host "Installing BCContainerHelper module. Please wait...`n" -ForegroundColor Green

        try {
            Install-Module BCContainerHelper -Force -ErrorAction Stop
        } catch {
            $errorMessage = $_.ToString()
            Write-Host "Error occurred during module installation:`n$errorMessage"
            Exit 1
        }

        [System.Windows.Forms.MessageBox]::Show("Module BCContainerHelper successfully installed!", "Success", "OK", "Asterisk") | Out-Null
    }
}


# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                  ### Executive Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Write-Host "Running initializer...`n`n" -ForegroundColor Green
Update-BcContainerHelper

Exit 0