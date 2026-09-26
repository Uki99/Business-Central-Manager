

                                                                                    ### Setup Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force

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
function Start-BusinessCentralManagerUpdateCheck {
    if (-not $settings.settings.checkForApplicationUpdateOnStart) {
        return
    }

    Write-Host "Checking for Business Central Manager updates. Please wait...`n"
    
    $owner = "Uki99"
    $repo = "Business-Central-Manager"
    $exitCode = 0
    
    try {
        Update-BusinessCentralManager -Owner $owner -Repository $repo -CurrentVersion $settings.settings.applicationVersion -ShowUpToDateMessage $false
    } catch {
        $errorMessage = $_.ToString()
        Write-Host "Error occurred during application update:`n$errorMessage" -ForegroundColor Red
    }
}


# Used to update/install the required module BcContainer helper neccesary for application
function Start-BcContainerHelperUpdateCheck {
    if (-not $settings.settings.searchForUpdateBcContainerHelper) {
        return
    }

    Write-Host "Checking for BcContainerHelper module updates. Please wait...`n"

    try {
        Update-BcContainerHelper
    } catch {
        $errorMessage = $_.ToString()
        Write-Host "Error occurred during BCContainerHelper update:`n$errorMessage" -ForegroundColor Red
    }
}


# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                  ### Executive Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

Write-Host "Running initializer...`n`n" -ForegroundColor Green
Start-BusinessCentralManagerUpdateCheck
Start-BcContainerHelperUpdateCheck
Write-Host "Initializer finishing...`n`n" -ForegroundColor Green

Exit 0