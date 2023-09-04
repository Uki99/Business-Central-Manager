Set-ExecutionPolicy Unrestricted

### Update Management ###
function Update-BCManager {
    param (
        [string] $owner,
        [string] $repo,
        [string] $currentVersion
    )

    # Step 1: Send a request to get the latest release information
    $uri = "https://api.github.com/repos/$owner/$repo/releases/latest"
    $releaseInfo = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction Stop

    # Step 2: Extract the zipball URL from the release information
    $zipballUrl = $releaseInfo.zipball_url

    # Step 3: Download and extract the zipball to a temporary folder
    $tempFolder = Join-Path $env:TEMP "Business-Central-Manager-temp"
    $tempZipPath = Join-Path $env:TEMP "Business-Central-Manager-temp.zip"
    
    # Download the zipball
    Invoke-WebRequest -Uri $zipballUrl -OutFile $tempZipPath -ErrorAction Stop
    
    # Extract the zipball
    Expand-Archive -Path $tempZipPath -DestinationPath $tempFolder -Force -ErrorAction Stop

    # Step 4: Check if update is needed
    $tempSettingsJsonPath = Join-Path $tempFolder "Business-Central-Manager-$($releaseInfo.tag_name)\Business Central Manager\data\settings.json"
    $tempSettings = Get-Content $tempSettingsJsonPath -Raw | ConvertFrom-Json -ErrorAction Stop

    $tempVersion = [version] $tempSettings.Version
    $lcurrentVersion = [version] $version

    if ($tempVersion -gt $lcurrentVersion) {
        $ConfirmApplicationUpdate = [System.Windows.Forms.MessageBox]::Show("Updates for Business Central Manager were found. Do you want to update now?", "Confirm Application Update", "YesNo", "Question") | Out-Null      
        if ($ConfirmApplicationUpdate -eq "No") {
            return
        }

        Write-Host "Updating Business Central Manager application. Please wait...`n" -ForegroundColor Green

        # Step 5: Replace files in the running folder
        Copy-Item "$tempFolder\Business-Central-Manager-$($releaseInfo.tag_name)\Business Central Manager\*" -Destination ($PSScriptRoot | Split-Path) -Recurse -Force -ErrorAction Stop
        
        # Step 6: Cleanup - remove temporary files and folders
        Remove-Item -Path $tempZipPath, $tempFolder -Force -Recurse -ErrorAction Stop

        Write-Host "Successfuly updated Business Central Manager to version $tempVersion...`n" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("Application Business Central Manager successfully updated!", "Success", "OK", "Asterisk") | Out-Null
        Restart-BusinessCentralManager
    } else {
        Write-Host "Business Central Manager is up to date.`n"
        
        # Step 7: Cleanup - remove temporary files and folders
        Remove-Item -Path $tempZipPath, $tempFolder -Force -Recurse -ErrorAction Stop
    }
}

function Restart-BusinessCentralManager {
    $batchScriptPath = (($PSScriptRoot | Split-Path) + "\scripts\Autorun.bat")
    Start-Process -FilePath $batchScriptPath
    Exit
}