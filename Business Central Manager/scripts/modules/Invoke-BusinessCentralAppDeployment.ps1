param (
    [ValidateSet('Single', 'Batch')] [string] $Operation,
    [string] $AppPath,
    [string] $ServerInstance,
    [string] $Tenant,
    [string] $SyncMode,
    [string] $NavAdminToolPath,
    [bool] $UnpublishPreviousAppVersion,
    [string] $MainScriptPath
)

$ErrorActionPreference = 'Stop'
Import-Module -Name $NavAdminToolPath -ErrorAction Stop
Add-Type -AssemblyName PresentationFramework
$settings = [pscustomobject]@{
    settings = [pscustomobject]@{ unpublishLastInstalledAppDuringUpgrade = $UnpublishPreviousAppVersion }
}

$tokens = $null
$parseErrors = $null
$scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($MainScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors) { throw ($parseErrors.Message -join '; ') }
foreach ($functionName in @('Install-BusinessCentralApp', 'Update-BusinessCentralApp', 'Invoke-BusinessCentralAppDeployment')) {
    $functionAst = $scriptAst.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName }, $true)
    if (-not $functionAst) { throw "The app deployment function $functionName was not found." }
    . ([scriptblock]::Create($functionAst.Extent.Text))
}
$syncModeValue = [System.Enum]::Parse([Microsoft.Dynamics.Nav.Types.NavAppSyncMode], $SyncMode)

if ($Operation -eq 'Batch') {
    Import-Module -Name BcContainerHelper -DisableNameChecking -ErrorAction Stop
    $apps = @(Get-ChildItem -LiteralPath $AppPath -Filter '*.app' -File -ErrorAction Stop)
    if ($apps.Count -eq 0) { throw 'There are no apps for publishing in the selected folder.' }

    $versionErrors = @()
    foreach ($app in $apps) {
        $appInfo = Get-NAVAppInfo -Path $app.FullName -ErrorAction Stop
        $publishedApps = Get-NAVAppInfo -Id $appInfo.Id -ServerInstance $ServerInstance -Tenant $Tenant -TenantSpecificProperties -ErrorAction Stop
        if ($publishedApps | Where-Object { [version]$_.Version -gt [version]$appInfo.Version }) {
            $versionErrors += ("{0} ({1}) has a newer version published or installed." -f $appInfo.Name, $appInfo.Version)
        }
    }
    if ($versionErrors.Count -gt 0) {
        throw (($versionErrors -join "`n") + "`nProvide newer versions and try again.")
    }

    $orderedApps = @(Sort-AppFilesByDependencies -appFiles $apps.FullName -ErrorAction Stop)
    if ($orderedApps.Count -ne $apps.Count) { throw 'Dependency sorting did not return every app in the folder.' }
    [pscustomobject]@{ Kind = 'Count'; Total = $apps.Count }

    foreach ($app in $orderedApps) {
        $appName = [System.IO.Path]::GetFileNameWithoutExtension($app)
        $appVersion = ''
        try {
            $appInfo = Get-NAVAppInfo -Path $app -ErrorAction Stop
            $appName = $appInfo.Name
            $appVersion = $appInfo.Version
            $hasErrors = Invoke-BusinessCentralAppDeployment -AppPath $app -ServerInstance $ServerInstance -Tenant $Tenant -SyncMode $syncModeValue -SuppressGui $true
        }
        catch {
            $hasErrors = $true
            Write-Warning ("Could not process {0}: {1}" -f $appName, $_.Exception.Message)
        }
        [pscustomobject]@{ Kind = 'App'; Name = $appName; Version = $appVersion; Succeeded = (-not $hasErrors) }
    }
}
else {
    $appInfo = Get-NAVAppInfo -Path $AppPath -ErrorAction Stop
    $hasErrors = Invoke-BusinessCentralAppDeployment -AppPath $AppPath -ServerInstance $ServerInstance -Tenant $Tenant -SyncMode $syncModeValue -SuppressGui $true
    [pscustomobject]@{ Kind = 'App'; Name = $appInfo.Name; Version = $appInfo.Version; Succeeded = (-not $hasErrors) }
}