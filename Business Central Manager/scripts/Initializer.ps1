

                                                                                    ### Setup Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Set-ExecutionPolicy Unrestricted
Add-Type -AssemblyName System.Windows.Forms
Import-Module -Force (($PSScriptRoot | Split-Path) + "\scripts\Update-Management.ps1")

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

# Used to check and update application
function Update-CheckBCManager {
    if (-not $settings.settings.CheckForApplicationUpdateOnStart) {
        return
    }

    Write-Host "Checking for Business Central Manager updates. Please wait...`n"
    
    $owner = "Uki99"
    $repo = "Business-Central-Manager"
    
    try {
        Update-BCManager -owner $owner -repo $repo -version $settings.settings.verion -upToDateMessage $false
    } catch {
        $errorMessage = $_.ToString()
        Write-Host "Error occurred during application update:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
        $null = Read-Host
    }
}


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
            
            Write-Host "Updating BCContainerHelper module. Please wait...`n"
            
            try {
                Update-Module BcContainerHelper -ErrorAction Stop
            } catch {
                $errorMessage = $_.ToString()
                Write-Host "Error occurred during module update:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
                $null = Read-Host
            }

            Write-Host "Module BCContainerHelper successfully updated!", "BcContainerHelper Update" -ForegroundColor Green
        } elseif ($newestVersion -eq $installedModule.Version) {
            Write-Host "BCContainerHelper module is already up to date.`n"
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
            Write-Host "Error occurred during module installation:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
            $null = Read-Host
        }

        Write-Host "Module BCContainerHelper successfully installed!" -ForegroundColor Green
    }
}


# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                  ### Executive Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Write-Host "Running initializer...`n`n" -ForegroundColor Green
Update-CheckBCManager
Update-BcContainerHelper

Exit 0