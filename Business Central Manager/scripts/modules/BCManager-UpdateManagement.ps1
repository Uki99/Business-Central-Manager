# Management script container that has functionality for application update #

Set-ExecutionPolicy Unrestricted
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Used to update Business Central Application
function Update-BCManagerApplication {
    param (
        [Parameter(Mandatory = $true)] [string] $owner,
        [Parameter(Mandatory = $true)] [string] $repo,
        [Parameter(Mandatory = $true)] [string] $currentVersion,
        [boolean] $upToDateMessage,
        [int] $exitCode
    )

    # Step 1: Send a request to get the latest release information from GitHub
    $uri = "https://api.github.com/repos/$owner/$repo/releases/latest"
    $releaseInfo = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction Stop

    # Step 2: Extract the zipball URL from the release information
    $zipballUrl = $releaseInfo.zipball_url

    # Step 3: Download and extract the zipball to a temporary folder
    $tempFolder = Join-Path $env:TEMP "Business-Central-Manager-temp"
    $tempZipPath = Join-Path $env:TEMP "Business-Central-Manager-temp.zip"

    # Cleanup - remove temporary files and folders from before to avoid bugs, if there is any
    Remove-Item -Path $tempZipPath, $tempFolder -Force -Recurse -ErrorAction SilentlyContinue
    
    # Download the zipball
    Invoke-WebRequest -Uri $zipballUrl -OutFile $tempZipPath -ErrorAction Stop
    
    # Extract the zipball
    Expand-Archive -Path $tempZipPath -DestinationPath $tempFolder -Force -ErrorAction Stop

    # Search for the dynamically generated folder name
    $generatedFolder = Get-ChildItem -Path $tempFolder -Directory | Where-Object { $_.Name -like "$owner-$repo-*" }

    # Check if the folder was found
    if ($generatedFolder) {
        # Construct the full path to the generated folder
        $fullPathToGeneratedFolder = Join-Path -Path $tempFolder -ChildPath $generatedFolder.Name
    } else {
        throw "Temp path could not be resolved while updating Business Central manager."
    }

    $tempSettings = Get-Content ($fullPathToGeneratedFolder + "\Business Central Manager\data\settings.json") -Raw | ConvertFrom-Json -ErrorAction Stop

    $tempVersion = [version] $tempSettings.settings.ApplicationVersion
    $lcurrentVersion = [version] $currentVersion

    # Step 4: Check if update is needed
    if ($tempVersion -gt $lcurrentVersion) {
        $ConfirmApplicationUpdate = [System.Windows.Forms.MessageBox]::Show(("Updates for Business Central Manager were found.`n`nCurrent version: {0}`nLatest version: {1}`n`nDo you want to download updates now?" -f $lcurrentVersion, $tempVersion), "Confirm Application Update", "YesNo", "Question")      
        if ($ConfirmApplicationUpdate -eq "No") {
            return
        }

        Write-Host "Updating Business Central Manager application. Please wait...`n"

        # Step 5: Replace files in the running folder
        $applicationRootLocation = ($PSScriptRoot | Split-Path | Split-Path)
        $oldSettings = Get-Content ($applicationRootLocation + '\data\settings.json') -Raw | ConvertFrom-Json -ErrorAction Stop
        Copy-Item "$fullPathToGeneratedFolder\Business Central Manager\*" -Destination $applicationRootLocation -Recurse -Force -ErrorAction Stop

        # Step 6: Return old user settings
        $tempSettings.settings.HidePowerShellConsole = $oldSettings.settings.HidePowerShellConsole
        $tempSettings.settings.NavAdminTool = $oldSettings.settings.NavAdminTool
        $tempSettings.settings.MainWindowXAMLRelativePath = $oldSettings.settings.MainWindowXAMLRelativePath
        $tempSettings.settings.SupportedBusinessCentralVersion = $oldSettings.settings.SupportedBusinessCentralVersion
        $tempSettings.settings.UnpublishLastInstalledAppDuringUpgrade = $oldSettings.settings.UnpublishLastInstalledAppDuringUpgrade
        $tempSettings.settings.SearchForUpdateBcContainerHelper = $oldSettings.settings.SearchForUpdateBcContainerHelper
        $tempSettings.settings.DelayBcContainerHelperModuleImport = $oldSettings.settings.DelayBcContainerHelperModuleImport
        $tempSettings.settings.CheckForApplicationUpdateOnStart = $oldSettings.settings.CheckForApplicationUpdateOnStart

        # Save the updated settings.json file
        $tempSettings | ConvertTo-Json | Set-Content -Path ($applicationRootLocation | Join-Path -ChildPath "data\settings.json") -Force
        
        # Step 7: Cleanup - remove temporary files and folders
        Remove-Item -Path $tempZipPath, $tempFolder -Force -Recurse -ErrorAction Stop

        Write-Host "Successfuly updated Business Central Manager to version $tempVersion. Restarting application.`n" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show(("Successfuly updated Business Central Manager to the version {0}. Restarting application." -f $tempVersion), "Success", "OK", "Asterisk") | Out-Null

        $exitCode = 200
        Restart-BusinessCentralManager
    } else {
        Write-Host "Business Central Manager is up to date.`n"
        if ($upToDateMessage) {
            [System.Windows.Forms.MessageBox]::Show("Business Central Manager is already up to date.`n", "Success", "OK", "Asterisk") | Out-Null
        }

        # Step 8: Cleanup - remove temporary files and folders
        Remove-Item -Path $tempZipPath, $tempFolder -Force -Recurse -ErrorAction Stop
    }
}

function Restart-BusinessCentralManager {
    $batchScriptPath = (($PSScriptRoot | Split-Path | Split-Path) + "\scripts\Autorun.bat")
    Start-Process -FilePath $batchScriptPath
}