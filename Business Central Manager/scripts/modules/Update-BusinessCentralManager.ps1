# Management script container that has functionality for application update #

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

function Copy-BusinessCentralManagerRelease {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string] $SourceRoot,
        [Parameter(Mandatory = $true)] [string] $DestinationRoot,
        [Parameter(Mandatory = $true)] [string] $BackupRoot
    )

    $files = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse -Force -ErrorAction Stop | Sort-Object FullName)
    if ($files.Count -eq 0) { throw 'The downloaded release contains no files.' }

    $overwritten = @()
    $created = @()
    foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($SourceRoot.TrimEnd('\').Length + 1)
        $destination = Join-Path $DestinationRoot $relativePath
        if (Test-Path -LiteralPath $destination -PathType Leaf) {
            $backup = Join-Path $BackupRoot $relativePath
            $null = [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($backup))
            Copy-Item -LiteralPath $destination -Destination $backup -Force -ErrorAction Stop
            $overwritten += [pscustomobject]@{ Backup = $backup; Destination = $destination }
        } else {
            $created += $destination
        }
    }

    try {
        foreach ($file in $files) {
            $relativePath = $file.FullName.Substring($SourceRoot.TrimEnd('\').Length + 1)
            $destination = Join-Path $DestinationRoot $relativePath
            $null = [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($destination))
            Copy-Item -LiteralPath $file.FullName -Destination $destination -Force -ErrorAction Stop
        }
    } catch {
        $copyError = $_.Exception.Message
        $rollbackErrors = @()
        foreach ($entry in $overwritten) {
            try {
                Copy-Item -LiteralPath $entry.Backup -Destination $entry.Destination -Force -ErrorAction Stop
            } catch {
                $rollbackErrors += $_.Exception.Message
            }
        }
        foreach ($destination in $created) {
            try {
                if (Test-Path -LiteralPath $destination -PathType Leaf) {
                    Remove-Item -LiteralPath $destination -Force -ErrorAction Stop
                }
            } catch {
                $rollbackErrors += $_.Exception.Message
            }
        }
        if ($rollbackErrors.Count -gt 0) {
            $failure = New-Object System.InvalidOperationException ("Update failed: {0}. Rollback incomplete; backups remain at {1}: {2}" -f $copyError, $BackupRoot, ($rollbackErrors -join '; '))
            $failure.Data['PreserveBackup'] = $true
            throw $failure
        }
        throw "Update failed and previous files were restored: $copyError"
    }
}

# Used to update Business Central Application
function Update-BusinessCentralManager {
    param (
        [Parameter(Mandatory = $true)] [string] $Owner,
        [Parameter(Mandatory = $true)] [string] $Repository,
        [Parameter(Mandatory = $true)] [string] $CurrentVersion,
        [boolean] $ShowUpToDateMessage
    )

    # Step 1: Send a request to get the latest release information from GitHub
    $uri = "https://api.github.com/repos/$Owner/$Repository/releases/latest"
    $releaseInfo = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction Stop
    $installedVersion = [version] $CurrentVersion

    if ($releaseInfo.tag_name -match '^v?(\d+\.\d+\.\d+\.\d+)$' -and [version]$Matches[1] -le $installedVersion) {
        Write-Host "Business Central Manager is up to date.`n"
        if ($ShowUpToDateMessage) {
            [System.Windows.Forms.MessageBox]::Show("Business Central Manager is already up to date.`n", "Success", "OK", "Asterisk") | Out-Null
        }
        return
    }

    # Step 2: Extract the zipball URL from the release information
    $zipballUrl = $releaseInfo.zipball_url

    # Step 3: Download and extract the zipball to a private temporary folder
    $tempRoot = Join-Path $env:TEMP ("Business-Central-Manager-{0}" -f [guid]::NewGuid())
    $tempFolder = Join-Path $tempRoot 'release'
    $tempZipPath = Join-Path $tempRoot 'release.zip'
    $null = New-Item -ItemType Directory -Path $tempRoot -ErrorAction Stop
    $preserveTemp = $false

    try {
    # Download the zipball
    Invoke-WebRequest -Uri $zipballUrl -OutFile $tempZipPath -ErrorAction Stop
    
    # Extract the zipball
    Expand-Archive -Path $tempZipPath -DestinationPath $tempFolder -Force -ErrorAction Stop

    # Search for the dynamically generated folder name
    $generatedFolder = Get-ChildItem -Path $tempFolder -Directory | Where-Object { $_.Name -like "$Owner-$Repository-*" }

    # Check if the folder was found
    if ($generatedFolder) {
        # Construct the full path to the generated folder
        $fullPathToGeneratedFolder = Join-Path -Path $tempFolder -ChildPath $generatedFolder.Name
    } else {
        throw "Temp path could not be resolved while updating Business Central manager."
    }

    $tempSettings = Get-Content ($fullPathToGeneratedFolder + "\Business Central Manager\data\settings.json") -Raw | ConvertFrom-Json -ErrorAction Stop

    $tempVersion = [version] $tempSettings.settings.applicationVersion
    # Step 4: Check if update is needed
    if ($tempVersion -gt $installedVersion) {
        $ConfirmApplicationUpdate = [System.Windows.Forms.MessageBox]::Show(("Updates for Business Central Manager were found.`n`nCurrent version: {0}`nLatest version: {1}`n`nDo you want to download updates now?" -f $installedVersion, $tempVersion), "Confirm Application Update", "YesNo", "Question")      
        if ($ConfirmApplicationUpdate -eq "No") {
            return
        }

        Write-Host "Updating Business Central Manager application. Please wait...`n"

        # Step 5: Preserve user settings in the staged release
        $applicationRootLocation = ($PSScriptRoot | Split-Path| Split-Path | Split-Path)
        $oldSettings = Get-Content -LiteralPath (Join-Path $applicationRootLocation 'Business Central Manager\data\settings.json') -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in $oldSettings.settings.PSObject.Properties) {
            if ($property.Name -ne 'applicationVersion' -and $tempSettings.settings.PSObject.Properties[$property.Name]) {
                $tempSettings.settings.($property.Name) = $property.Value
            }
        }
        $stagedSettingsPath = Join-Path $fullPathToGeneratedFolder 'Business Central Manager\data\settings.json'
        $tempSettings | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $stagedSettingsPath -Encoding UTF8 -ErrorAction Stop
        Copy-BusinessCentralManagerRelease -SourceRoot $fullPathToGeneratedFolder -DestinationRoot $applicationRootLocation -BackupRoot (Join-Path $tempRoot 'backup')
        
        Restart-BusinessCentralManager
        Write-Host "Successfully updated Business Central Manager to version $tempVersion. Restarting application.`n" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show(("Successfully updated Business Central Manager to version {0}. Restarting application." -f $tempVersion), "Success", "OK", "Asterisk") | Out-Null
        Exit 200
    } else {
        Write-Host "Business Central Manager is up to date.`n"
        if ($ShowUpToDateMessage) {
            [System.Windows.Forms.MessageBox]::Show("Business Central Manager is already up to date.`n", "Success", "OK", "Asterisk") | Out-Null
        }

    }
    } catch {
        $preserveTemp = [bool]$_.Exception.Data['PreserveBackup']
        throw
    } finally {
        if (-not $preserveTemp) {
            Remove-Item -LiteralPath $tempRoot -Force -Recurse -ErrorAction SilentlyContinue
        }
    }
}

function Restart-BusinessCentralManager {
    $launcherPath = Join-Path ($PSScriptRoot | Split-Path | Split-Path) 'Start-BusinessCentralManager.cmd'
    Start-Process -FilePath $launcherPath -WindowStyle Hidden -ErrorAction Stop
}
