# Management script container that has functionality for BC Container Helper module update #

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Used to check for updates or install the BCContainerHelper module
function Update-BcContainerHelper {
    param ([System.Windows.Forms.IWin32Window] $Owner)

    try {
        # Check if BCContainerHelper module is installed
        $availableModule = Get-Module -ListAvailable -Name BcContainerHelper | Sort-Object Version -Descending | Select-Object -First 1
        $installedModule = Get-InstalledModule -Name BcContainerHelper -ErrorAction SilentlyContinue

        if ($installedModule) {
            $currentVersion = $installedModule.Version
            if ($availableModule -and $availableModule.Version -gt [version]$currentVersion) {
                $currentVersion = $availableModule.Version
            }
            $repository = if ($installedModule.Repository) { $installedModule.Repository } else { 'PSGallery' }
            $latestVersion = (Find-Module -Name BcContainerHelper -Repository $repository -ErrorAction Stop).Version

            if ([version]$latestVersion -gt [version]$currentVersion) {
                $ConfirmModuleUpdate = [System.Windows.Forms.MessageBox]::Show($Owner, "PowerShell module BCContainerHelper found. Do you want to update the module?", "Confirm Module Update", "YesNo", "Question")

                if ($ConfirmModuleUpdate -eq "No") {
                    return
                }

                Write-Host "Updating BCContainerHelper module. Please wait...`n"
                Update-Module BcContainerHelper -ErrorAction Stop
                Write-Host "Module BCContainerHelper successfully updated" -ForegroundColor Green
            } else {
                Write-Host "BCContainerHelper module is already up to date.`n" -ForegroundColor Green
            }
        } elseif ($availableModule) {
            Write-Host "BCContainerHelper is available but was not installed with PowerShellGet; skipping update check.`n"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show($Owner, "PowerShell module BCContainerHelper not found. Press OK to install the required module now.", "BCContainerHelper Install", "OK", "Warning") | Out-Null

            Write-Host "Installing BCContainerHelper module. Please wait...`n"

            Install-Module BCContainerHelper -Force -ErrorAction Stop
            Write-Host "Module BCContainerHelper successfully installed`n" -ForegroundColor Green
        }
    } catch {
        throw "BCContainerHelper setup failed: $($_.Exception.Message)"
    }
}