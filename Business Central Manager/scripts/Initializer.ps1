

                                                                                    ### Setup Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Set-ExecutionPolicy Unrestricted

Import-Module -Force (($PSScriptRoot | Split-Path) + "\scripts\modules\BCManager-UpdateManagement.ps1")
Import-Module -Force (($PSScriptRoot | Split-Path) + "\scripts\modules\BCContainerHelper-UpdateManagement.ps1")

Add-Type -AssemblyName System.Windows.Forms

# Load settings from settings.json
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

# Used to check and update main application
function Update-BCManager {
    if (-not $settings.settings.CheckForApplicationUpdateOnStart) {
        return
    }

    Write-Host "Checking for Business Central Manager updates. Please wait...`n"
    
    $owner = "Uki99"
    $repo = "Business-Central-Manager"
    $exitCode = 0
    
    try {
        Update-BCManagerApplication -owner $owner -repo $repo -currentVersion $settings.settings.ApplicationVersion -upToDateMessage $false -exitCode $exitCode

        if ($exitCode = 200) {
            Exit 200
        }
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

    try {
        Update-BCContainerHelperModule
    } catch {
        $errorMessage = $_.ToString()
        Write-Host "Error occurred during BCContainerHelper update:`n$errorMessage`n`nPress any key to continue" -ForegroundColor Red
        $null = Read-Host
    }
}


# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                  ### Executive Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Write-Host "Running initializer...`n`n" -ForegroundColor Green
Update-BCManager
Update-BcContainerHelper
Write-Host "Initializer finishing...`n`n" -ForegroundColor Green

Exit 0