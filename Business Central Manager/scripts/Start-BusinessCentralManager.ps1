

                                                                                    ### Setup Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# Check for admin rights and ask if needed
$settingsPath = Join-Path ($PSScriptRoot | Split-Path) 'data\settings.json'
$hideConsole = $false
try {
    $hideConsole = [bool](Get-Content -LiteralPath $settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop).settings.hidePowerShellConsole
} catch { }
$windowStyle = if ($hideConsole) { 'Hidden' } else { 'Normal' }

if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process -FilePath PowerShell.exe -Verb Runas -WindowStyle $windowStyle -ArgumentList ('-NoProfile -ExecutionPolicy Unrestricted -File "{0}"' -f $PSCommandPath) -ErrorAction Stop
    Exit
}

# Set execution policy and import required assemblies
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
using System;
using System.Windows.Forms;

public sealed class BcManagerDialogOwner : IWin32Window
{
    public BcManagerDialogOwner(IntPtr handle) { Handle = handle; }
    public IntPtr Handle { get; private set; }
}
'@


# Load settings from json
try {
    $settings = Get-Content -LiteralPath $settingsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

function Invoke-BusinessCentralManagerUpdateCheck {
    param ([bool] $ShowUpToDateMessage)

    Import-Module -Force (Join-Path $PSScriptRoot 'modules\Update-BusinessCentralManager.ps1')
    Update-BusinessCentralManager -Owner 'Uki99' -Repository 'Business-Central-Manager' -CurrentVersion $settings.settings.applicationVersion -ShowUpToDateMessage $ShowUpToDateMessage
}

if ($settings.settings.checkForApplicationUpdateOnStart) {
    Write-Host "Checking for Business Central Manager updates. Please wait...`n"
    try {
        Invoke-BusinessCentralManagerUpdateCheck -ShowUpToDateMessage $false
    } catch {
        Write-Host "Error occurred during application update:`n$_" -ForegroundColor Red
    }
}

$dependencyFailures = @{}

function Resolve-NavAdminToolPath {
    $configuredPath = $settings.settings.navAdminTool
    if ($configuredPath -and (Test-Path -LiteralPath $configuredPath -PathType Leaf)) {
        return $configuredPath
    }

    $installRoot = Join-Path $env:ProgramFiles 'Microsoft Dynamics 365 Business Central'
    $candidates = @(Get-ChildItem -LiteralPath $installRoot -Directory -ErrorAction SilentlyContinue |
        ForEach-Object { Join-Path $_.FullName 'Service\NavAdminTool.ps1' } |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })

    if ($candidates.Count -eq 1) { return $candidates[0] }
    if ($candidates.Count -eq 0) {
        throw 'NavAdminTool.ps1 was not found. Specify its path in Settings.'
    }
    throw 'Multiple Business Central installations were found. Specify the NavAdminTool.ps1 path in Settings.'
}

function Import-RequiredDependency {
    param ([ValidateSet('NavAdminTool', 'BcContainerHelper')] [string] $Name)

    $source = if ($Name -eq 'NavAdminTool') { $settings.settings.navAdminTool } else { 'BcContainerHelper' }
    try {
        if ($Name -eq 'NavAdminTool') {
            $source = Resolve-NavAdminToolPath
        }
        if ($dependencyFailures.ContainsKey($Name)) {
            Import-Module -Name $source -Force -ErrorAction Stop
        } elseif (-not (Get-Module -Name $Name)) {
            Import-Module -Name $source -ErrorAction Stop
        }
        $dependencyFailures.Remove($Name) | Out-Null
        return $true
    } catch {
        $dependencyFailures[$Name] = [pscustomobject]@{
            AttemptedAt = Get-Date
            Source = $source
            Error = $_.Exception.Message
        }
        return $false
    }
}

# Import required NavAdminTool module
$null = Import-RequiredDependency -Name NavAdminTool

# Setup MainWindow.xaml for GUI
try {
    $inputXML = Get-Content -LiteralPath (Join-Path ($PSScriptRoot | Split-Path) 'forms\MainWindow.xaml') -Raw -ErrorAction Stop
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

[XML]$MainWindowXAML = $inputXML
$xamlNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
$MainWindowXAML.DocumentElement.RemoveAttribute('Class', $xamlNamespace)
$MainWindowXAML.DocumentElement.RemoveAttribute('Ignorable', 'http://schemas.openxmlformats.org/markup-compatibility/2006')

$reader = (New-Object System.Xml.XmlNodeReader $MainWindowXAML)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

# Set main window icon
$IconPath = Join-Path ($PSScriptRoot | Split-Path) 'data\icon.ico'

# Create a FileStream to read the icon file
$FileStream = [System.IO.File]::Open($IconPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)

# Create a MemoryStream to hold the icon data
$MemoryStream = [System.IO.MemoryStream]::new()

# Copy the icon data from the FileStream to the MemoryStream
$FileStream.CopyTo($MemoryStream)

# Close the FileStream
$FileStream.Close()

# Create a BitmapImage and set its source to the MemoryStream
$Icon = [System.Windows.Media.Imaging.BitmapImage]::new()
$Icon.BeginInit()
$Icon.StreamSource = $MemoryStream
$Icon.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
$Icon.EndInit()

# Set the window's icon
$Window.Icon = $Icon

# Set control variables for GUI
$namespaceManager = [System.Xml.XmlNamespaceManager]::new($MainWindowXAML.NameTable)
$namespaceManager.AddNamespace('x', $xamlNamespace)
$MainWindowXAML.SelectNodes('//*[@x:Name]', $namespaceManager) | ForEach-Object {
    try {
        $controlName = $_.GetAttribute('Name', $xamlNamespace)
        $control = $window.FindName($controlName)
        if ($null -ne $control) {
            Set-Variable -Name "var_$controlName" -Value $control -ErrorAction Stop
        }
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }
}

# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                  ### Function Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


function Select-File {
    param (
        [Parameter(Mandatory = $true)] [string] $FileFilter,
        [string] $Directory = ([environment]::GetFolderPath('Desktop')),
        [string] $Title = 'Select a file',
        [System.Windows.Forms.IWin32Window] $Owner
    )

    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    try {
        if ($Directory -and (Test-Path -LiteralPath $Directory -PathType Container)) {
            $OpenFileDialog.InitialDirectory = (Resolve-Path -LiteralPath $Directory).Path
        }
        $OpenFileDialog.RestoreDirectory = $true
        $OpenFileDialog.CheckFileExists = $true
        $OpenFileDialog.Multiselect = $false
        $OpenFileDialog.Filter = $FileFilter
        $OpenFileDialog.Title = $Title

        $result = if ($Owner) { $OpenFileDialog.ShowDialog($Owner) } else { $OpenFileDialog.ShowDialog() }
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and
            (Test-Path -LiteralPath $OpenFileDialog.FileName -PathType Leaf)) {
            return $OpenFileDialog.FileName
        }
        return $null
    } finally {
        $OpenFileDialog.Dispose()
    }
}

function Select-Folder {
    $folderBrowser = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder"
    $folderBrowser.RootFolder = 'Desktop'

    $result = $folderBrowser.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if (Test-Path $folderBrowser.SelectedPath) {
            return $folderBrowser.SelectedPath
        }
    }

    return $null
}

function Get-TopicDependencyFailure {
    param ([string] $TopicName)

    if ($TopicName -in @('AppManagementTopic', 'ServerManagementTopic') -and $dependencyFailures.ContainsKey('NavAdminTool')) {
        return 'NavAdminTool'
    }
    if ($TopicName -eq 'ContainerManagementTopic' -and $dependencyFailures.ContainsKey('BcContainerHelper')) {
        return 'BcContainerHelper'
    }
    return $null
}

function Set-TopicAvailability {
    foreach ($topic in @($var_AppManagementTopic, $var_ServerManagementTopic, $var_ContainerManagementTopic)) {
        $failureName = Get-TopicDependencyFailure -TopicName $topic.Name
        if ($failureName) {
            $topic.Opacity = 0.45
            $topic.ToolTip = "$failureName could not be loaded. Click for details."
        } else {
            $topic.Opacity = 1
            $topic.ToolTip = $null
        }
    }
}

$script:containerModuleReady = [bool](Get-Module -Name BcContainerHelper)
$script:containerTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:containerTimer.Interval = [TimeSpan]::FromMilliseconds(100)

function Receive-ContainerModuleImport {
    if (-not $script:containerInvocation -or -not $script:containerInvocation.IsCompleted) { return }

    $warning = $null
    try {
        $result = @($script:containerPowerShell.EndInvoke($script:containerInvocation))[-1]
        if (-not $result) { throw 'BCContainerHelper loading returned no result.' }
        $warning = $result.Warning
        if (-not $result.Loaded) { throw $result.Error }
        $script:containerModuleReady = $true
        $dependencyFailures.Remove('BcContainerHelper') | Out-Null
    } catch {
        $script:containerModuleReady = $false
        $dependencyFailures['BcContainerHelper'] = [pscustomobject]@{
            AttemptedAt = Get-Date
            Source = 'BcContainerHelper'
            Error = $_.Exception.Message
        }
    } finally {
        $script:containerTimer.Stop()
        $script:containerPowerShell.Dispose()
        $script:containerPowerShell = $null
        $script:containerInvocation = $null
        if (-not $script:containerModuleReady) {
            $script:containerRunspace.Dispose()
            $script:containerRunspace = $null
        }
        $var_ContainerModuleLoading.Visibility = [System.Windows.Visibility]::Collapsed
        $var_TopicsContainer.IsEnabled = $true
        $var_ContainerManagementTopic.IsEnabled = $true
        Set-TopicAvailability
    }
    if ($warning) {
        [System.Windows.Forms.MessageBox]::Show(("Could not check BCContainerHelper updates: {0}" -f $warning), 'Update check failed', 'OK', 'Warning') | Out-Null
    }
    if ($script:containerModuleReady) {
        $var_ContainerManagementTopic.IsChecked = $true
    } else {
        $failure = $dependencyFailures['BcContainerHelper']
        [System.Windows.Forms.MessageBox]::Show(("BcContainerHelper still could not be loaded from {0}.`nAttempted: {1}`nReason: {2}" -f $failure.Source, $failure.AttemptedAt, $failure.Error), 'Dependency unavailable', 'OK', 'Error') | Out-Null
    }
}

$script:containerTimer.Add_Tick({ Receive-ContainerModuleImport })

function Start-ContainerModuleImport {
    $script:activeTopic.IsChecked = $true
    $var_ContainerManagementTopic.IsEnabled = $false
    $var_TopicsContainer.IsEnabled = $false
    $var_ContainerModuleLoading.Visibility = [System.Windows.Visibility]::Visible
    try {
        $script:containerRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:containerRunspace.ApartmentState = 'STA'
        $script:containerRunspace.Open()
        $script:containerPowerShell = [PowerShell]::Create()
        $script:containerPowerShell.Runspace = $script:containerRunspace
        $null = $script:containerPowerShell.AddScript({
            param ([bool] $checkForUpdates, [string] $updaterPath, [IntPtr] $ownerHandle)

            $warning = $null
            if ($checkForUpdates) {
                try {
                    Import-Module -Name $updaterPath -ErrorAction Stop
                    Update-BcContainerHelper -Owner ([BcManagerDialogOwner]::new($ownerHandle))
                } catch {
                    $warning = $_.Exception.Message
                }
            }
            try {
                Import-Module -Name BcContainerHelper -ErrorAction Stop
                [pscustomobject]@{ Loaded = $true; Warning = $warning; Error = $null }
            } catch {
                [pscustomobject]@{ Loaded = $false; Warning = $warning; Error = $_.Exception.Message }
            }
        }.ToString()).AddArgument([bool]$settings.settings.searchForUpdateBcContainerHelper).AddArgument(
            (Join-Path $PSScriptRoot 'modules\Update-BcContainerHelper.ps1')).AddArgument(
            ([System.Windows.Interop.WindowInteropHelper]::new($window).Handle))
        $script:containerInvocation = $script:containerPowerShell.BeginInvoke()
        $script:containerTimer.Start()
    } catch {
        if ($script:containerPowerShell) { $script:containerPowerShell.Dispose(); $script:containerPowerShell = $null }
        if ($script:containerRunspace) { $script:containerRunspace.Dispose(); $script:containerRunspace = $null }
        $script:containerInvocation = $null
        $dependencyFailures['BcContainerHelper'] = [pscustomobject]@{
            AttemptedAt = Get-Date
            Source = 'BcContainerHelper'
            Error = $_.Exception.Message
        }
        $var_ContainerModuleLoading.Visibility = [System.Windows.Visibility]::Collapsed
        $var_TopicsContainer.IsEnabled = $true
        $var_ContainerManagementTopic.IsEnabled = $true
        Set-TopicAvailability
        [System.Windows.Forms.MessageBox]::Show(("Could not start BCContainerHelper loading: {0}" -f $_.Exception.Message), 'Dependency unavailable', 'OK', 'Error') | Out-Null
    }
}

function Update-ServerInstanceOptions {
    $instances = Get-NAVServerInstance -ErrorAction Stop | Where-Object { ($_.State -eq 'Running') -and ($_.Version -match $settings.settings.supportedBusinessCentralVersion) }
    foreach ($control in @($var_AppPublishingServerInstanceComboBox, $var_MultipleAppPublishingServerInstanceComboBox, $var_LicenseServerInstanceComboBox)) {
        $control.Items.Clear()
    }
    foreach ($instance in $instances) {
        $position = $instance.ServerInstance.IndexOf('$')
        $shortName = $instance.ServerInstance.Substring($position + 1)
        $null = $var_AppPublishingServerInstanceComboBox.Items.Add($shortName)
        $null = $var_MultipleAppPublishingServerInstanceComboBox.Items.Add($shortName)
        $null = $var_LicenseServerInstanceComboBox.Items.Add($shortName)
    }
}

function Resolve-TopicDependencies {
    param ([System.Windows.Controls.RadioButton] $Topic)

    while ($failureName = Get-TopicDependencyFailure -TopicName $Topic.Name) {
        $failure = $dependencyFailures[$failureName]
        $message = "{0} could not be loaded.`nSource: {1}`nLast attempt: {2}`nReason: {3}`n`nTry to load it now?" -f $failureName, $failure.Source, $failure.AttemptedAt, $failure.Error
        $answer = [System.Windows.Forms.MessageBox]::Show($message, 'Dependency unavailable', 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            return $false
        }

        if ($failureName -eq 'BcContainerHelper') {
            Start-ContainerModuleImport
            return $false
        }

        if (-not (Import-RequiredDependency -Name $failureName)) {
            $failure = $dependencyFailures[$failureName]
            [System.Windows.Forms.MessageBox]::Show(("{0} still could not be loaded from {1}.`nAttempted: {2}`nReason: {3}" -f $failureName, $failure.Source, $failure.AttemptedAt, $failure.Error), 'Dependency unavailable', 'OK', 'Error') | Out-Null
            Set-TopicAvailability
            return $false
        }
        Set-TopicAvailability
        if ($failureName -eq 'NavAdminTool') {
            Update-ServerInstanceOptions
        }
    }
    return $true
}

function Get-SyncMode {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [System.Windows.Controls.ComboBox] $SyncModeComboBox
    )

    $SelectedSyncItem = $SyncModeComboBox.SelectedItem
    $selectedSyncValue = $selectedSyncItem.Content.ToString()

    switch ($selectedSyncValue) {
        'Add (default)' {
            return [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add
        }
        'Clean' {
            return [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Clean
        }
        'Development' {
            return [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Development
        }
        'ForceSync' {
            return [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::ForceSync
        }
        'None' {
            return [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::None
        }
        default {
            # Handle unrecognized values or set a default sync mode
            return [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add
        }
    }
}

function Get-TargetTenant {
    if ([string]::IsNullOrWhiteSpace($settings.settings.tenantId)) {
        throw 'Set a tenant ID in Settings before publishing apps.'
    }
    return $settings.settings.tenantId
}

function Set-UIElement {
    param(
        [Parameter(Mandatory=$true)] $Element,
        [Parameter(Mandatory=$true)] $Property,
        [Parameter(Mandatory=$true)] $Value,
        [Parameter(Mandatory=$false)] $OverrideValue = $true
    )

    if ($OverrideValue) {
        $Element.Dispatcher.Invoke([Action]{
            $Element.$Property = $Value
        }, "Render")
    } else {
        $Element.Dispatcher.Invoke([Action]{
            $Element.$Property += $Value
        }, "Render")
    }
}

function Install-BusinessCentralApp {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [string] $AppPath,
        [Parameter(Mandatory = $true)] [string] $ServerInstance,
        [Parameter(Mandatory = $true)] [string] $Tenant,
        [Parameter(Mandatory = $false)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode] $SyncMode = [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add,
        [Parameter(Mandatory = $true)] [boolean] $SuppressGui,
        [Parameter(Mandatory = $false)] [boolean] $AlreadyPublished = $false,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.StackPanel] $ProgressBarContainer,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.ProgressBar] $ProgressBar,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.TextBlock] $ProgressInfo
    )

    if ($PSBoundParameters.ContainsKey('ProgressBarContainer') -and $PSBoundParameters.ContainsKey('ProgressBar') -and $PSBoundParameters.ContainsKey('ProgressInfo')) {
        $ShouldHandleProgressBar = $true
    } else {
        $ShouldHandleProgressBar = $false
    }
    
    $TargetAppInfo = Get-NAVAppInfo -Path $AppPath -ErrorAction Stop

    if (-not $SuppressGui) {
        $ConfirmAppInstall = [System.Windows.Forms.MessageBox]::Show(("Do you want to try and install the following app?`n`nName: {0}`nVersion: {1}" -f $TargetAppInfo.Name, $TargetAppInfo.Version), "Confirm App Install", "YesNo", "Question")      
        if ($ConfirmAppInstall -eq "No") {
            return $false
        }

        if ($SyncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
            $ConfirmSyncMode = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to install the app using sync mode = {0}?`n`nThis could be destructive." -f $SyncMode), "Confirm App Sync Mode", "YesNo", "Warning")
        
            if ($ConfirmSyncMode -eq "No") {
                return $false
            }
        }
    }


    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $true
        Set-UIElement -Element $ProgressBarContainer -Property "Visibility" -Value 0 # Visible
        $progressText = if ($AlreadyPublished) { "Synchronizing app..." } else { "Publishing app..." }
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value $progressText
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 25
    }
    if (-not $AlreadyPublished) {
        Publish-NAVApp -ServerInstance $ServerInstance -Path $AppPath -ErrorAction Stop
    }

    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Synchronizing app..."
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 50
    }
    Sync-NavApp -ServerInstance $ServerInstance -AppId $TargetAppInfo.Id -Version $TargetAppInfo.Version -Tenant $Tenant -Mode $SyncMode -Force -ErrorAction Stop
    
    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Installing app..."
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 75
    }
    try {
        Install-NAVApp -ServerInstance $ServerInstance -AppId $TargetAppInfo.Id -Version $TargetAppInfo.Version -Tenant $Tenant -Force -ErrorAction Stop
    } catch [InvalidOperationException] {
        <# This is to handle installing the apps which were uninstalled without clean mode but recognized as in need of install #>
        Start-NAVAppDataUpgrade -ServerInstance $ServerInstance -AppId $TargetAppInfo.Id -Version $TargetAppInfo.Version -Tenant $Tenant -Force -ErrorAction Stop
    }

    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Finalizing..."
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 100
    }
    if(-not $SuppressGui) {
        [System.Windows.Forms.MessageBox]::Show(("Successfully installed the extension {0} with version {1}!" -f $TargetAppInfo.Name, $TargetAppInfo.Version), "Succcess", "OK", "Asterisk") | Out-Null
    }

    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $false
        Set-UIElement -Element $ProgressBarContainer -Property "Visibility" -Value 2 # Hidden
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value ""
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 0
    }
    return $true
}

function Update-BusinessCentralApp {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [string] $AppPath,
        [Parameter(Mandatory = $true)] [string] $ServerInstance,
        [Parameter(Mandatory = $true)] [string] $Tenant,
        [Parameter(Mandatory = $false)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode] $SyncMode = [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add,
        [Parameter(Mandatory = $true)] [boolean] $SuppressGui,
        [Parameter(Mandatory = $false)] [boolean] $AlreadyPublished = $false,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.StackPanel] $ProgressBarContainer,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.ProgressBar] $ProgressBar,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.TextBlock] $ProgressInfo
    )

    if ($PSBoundParameters.ContainsKey('ProgressBarContainer') -and $PSBoundParameters.ContainsKey('ProgressBar') -and $PSBoundParameters.ContainsKey('ProgressInfo')) {
        $ShouldHandleProgressBar = $true
    } else {
        $ShouldHandleProgressBar = $false
    }

    $TargetAppInfo = Get-NAVAppInfo -Path $AppPath -ErrorAction Stop
    $OldAppInfo = Get-NAVAppInfo -ServerInstance $ServerInstance -Id $TargetAppInfo.Id -Tenant $Tenant -TenantSpecificProperties -ErrorAction Stop | Where-Object { $_.IsInstalled -eq $true }

    if (-not $SuppressGui) {
        $ConfirmAppUpdate = [System.Windows.Forms.MessageBox]::Show(("Do you want to try and update the following app?`n`nName: {0}`nNew Version: {1}`nCurrent Version: {2}" -f $TargetAppInfo.Name, $TargetAppInfo.Version, $OldAppInfo.Version), "Confirm App Update", "YesNo", "Question")       
        if ($ConfirmAppUpdate -eq "No") {
            return $false
        }

        if ($SyncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
            $ConfirmSyncMode = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to install the app using sync mode = {0}?`n`nThis could be destructive." -f $SyncMode), "Confirm App Sync Mode", "YesNo", "Warning")
        
            if ($ConfirmSyncMode -eq "No") {
                return $false
            }
        }
    }


    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $true
        Set-UIElement -Element $ProgressBarContainer -Property "Visibility" -Value 0 # Visible
        $progressText = if ($AlreadyPublished) { "Synchronizing app..." } else { "Publishing app..." }
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value $progressText
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 20
    }
    if (-not $AlreadyPublished) {
        Publish-NAVApp -ServerInstance $ServerInstance -Path $AppPath -ErrorAction Stop
    }

    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Synchronizing app..."
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 40
    }
    Sync-NavApp -ServerInstance $ServerInstance -AppId $TargetAppInfo.Id -Version $TargetAppInfo.Version -Tenant $Tenant -Mode $SyncMode -Force -ErrorAction Stop
    
    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Upgrading app..."
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 60
    }
    Start-NAVAppDataUpgrade -ServerInstance $ServerInstance -AppId $TargetAppInfo.Id -Version $TargetAppInfo.Version -Tenant $Tenant -Force -ErrorAction Stop


    if ($settings.settings.unpublishLastInstalledAppDuringUpgrade) {
        if ($ShouldHandleProgressBar) {
            Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Unpublishing old app..."
            Set-UIElement -Element $ProgressBar -Property "Value" -Value 80
        }
        try {
            $remainingTenants = @(Get-NAVAppTenant -ServerInstance $ServerInstance -Id $OldAppInfo.Id -Version $OldAppInfo.Version -IncludeFailed -ErrorAction Stop)
            if ($remainingTenants.Count -eq 0) {
                Unpublish-NAVApp -ServerInstance $ServerInstance -AppId $OldAppInfo.Id -Version $OldAppInfo.Version -ErrorAction Stop
            } else {
                Write-Warning "The previous app version is still installed on another tenant and remains published."
            }
        } catch {
            Write-Warning ("The upgrade succeeded, but the old version could not be unpublished: {0}" -f $_.Exception.Message)
        }
    }

    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value "Finalizing..."
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 100
    }
    if (-not $SuppressGui) {
        [System.Windows.Forms.MessageBox]::Show(("Successfully updated the extension {0} to the version {1}!" -f $TargetAppInfo.Name, $TargetAppInfo.Version), "Succcess", "OK", "Asterisk") | Out-Null
    }

    if ($ShouldHandleProgressBar) {
        Set-UIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $false
        Set-UIElement -Element $ProgressBarContainer -Property "Visibility" -Value 2 # Hidden
        Set-UIElement -Element $ProgressInfo -Property "Text" -Value ""
        Set-UIElement -Element $ProgressBar -Property "Value" -Value 0
    }
    return $true
}

function Invoke-BusinessCentralAppDeployment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]$AppPath,
        [Parameter(Mandatory = $true)] [string]$ServerInstance,
        [Parameter(Mandatory = $true)] [string]$Tenant,
        [Parameter(Mandatory = $true)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]$SyncMode,
        [Parameter(Mandatory = $false)] [boolean]$SuppressGui = $false,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.StackPanel]$ProgressBarContainer,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.ProgressBar]$ProgressBar,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.TextBlock]$ProgressInfo
    )

    $requestedApp = Get-NAVAppInfo -Path $AppPath -ErrorAction Stop
    $publishedApps = @(Get-NAVAppInfo -Id $requestedApp.Id -ServerInstance $ServerInstance -Tenant $Tenant -TenantSpecificProperties -ErrorAction Stop)
    $installedApps = @($publishedApps | Where-Object { $_.IsInstalled })
    $alreadyPublished = @($publishedApps | Where-Object { [version]$_.Version -eq [version]$requestedApp.Version }).Count -gt 0

    $notice = $null
    if ($publishedApps | Where-Object { [version]$_.Version -gt [version]$requestedApp.Version }) {
        $notice = "A newer version of $($requestedApp.Name) is already published or installed. No action was performed."
    } elseif ($installedApps | Where-Object { [version]$_.Version -eq [version]$requestedApp.Version }) {
        $notice = "$($requestedApp.Name) $($requestedApp.Version) is already installed. No action was performed."
    }
    if ($notice) {
        if ($SuppressGui) {
            Write-Warning $notice
        } else {
            [System.Windows.Forms.MessageBox]::Show($notice, 'App already available', 'OK', 'Information') | Out-Null
        }
        return $true
    }

    $appParameters = @{
        AppPath = $AppPath
        ServerInstance = $ServerInstance
        Tenant = $Tenant
        SyncMode = $SyncMode
        SuppressGui = $SuppressGui
        AlreadyPublished = $alreadyPublished
    }
    $showProgress = $PSBoundParameters.ContainsKey('ProgressBarContainer') -and $PSBoundParameters.ContainsKey('ProgressBar') -and $PSBoundParameters.ContainsKey('ProgressInfo')
    if ($showProgress) {
        $appParameters.ProgressBarContainer = $ProgressBarContainer
        $appParameters.ProgressBar = $ProgressBar
        $appParameters.ProgressInfo = $ProgressInfo
    }

    try {
        $applied = if ($installedApps.Count -gt 0) { Update-BusinessCentralApp @appParameters } else { Install-BusinessCentralApp @appParameters }
        return (-not $applied)
    } catch {
        if ($SuppressGui) {
            Write-Warning $_.Exception.Message
        } else {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Error', 'OK', 'Error') | Out-Null
        }
        if ($showProgress) {
            Set-UIElement -Element $ProgressBarContainer -Property 'IsEnabled' -Value $false
            Set-UIElement -Element $ProgressBarContainer -Property 'Visibility' -Value 2 # Hidden
            Set-UIElement -Element $ProgressInfo -Property 'Text' -Value ''
            Set-UIElement -Element $ProgressBar -Property 'Value' -Value 0
        }
        return $true
    }
}

function Set-SettingControlValue {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $SettingName,
        [Parameter(Mandatory = $true)] $Value
    )
    
    try {
        # Get variable in XAML that corresponds to the JSON key in settings.json
        $control = (Get-Variable -Name "var_$SettingName" -ErrorAction SilentlyContinue).Value

        if ($control -is [System.Windows.Controls.TextBox]) {
            $control.Text = $Value
        } elseif ($control -is [System.Windows.Controls.TextBlock]) {
            $control.Text = $Value
        } elseif ($control -is [System.Windows.Controls.CheckBox]) {
            $control.IsChecked = [bool]$Value
        }
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }
}

function Save-ApplicationSetting {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $SettingName,
        [Parameter(Mandatory = $true)] $PreviousValue
    )

    try {
        # Get variable in XAML that corresponds to the JSON key in settings.json
        $controlVariable = Get-Variable -Name "var_$SettingName" -ErrorAction SilentlyContinue
        if (-not $controlVariable) { return }
        $settingControl = $controlVariable.Value

        # Parse new values depending on the control type
        if ($settingControl -is [System.Windows.Controls.TextBox]) {
            $newValue = $settingControl.Text
        } elseif ($settingControl -is [System.Windows.Controls.CheckBox]) {
            $newValue = $settingControl.IsChecked
        } else {
            return
        }
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }

    # Save to Json on disk
    if ($PreviousValue -eq $newValue) {
        return
    }

    $settings.settings.$SettingName = $newValue
    Add-Type -AssemblyName System.Runtime.Serialization
    $jsonBytes = [System.Text.Encoding]::UTF8.GetBytes(($settings | ConvertTo-Json -Compress))
    $reader = [System.Runtime.Serialization.Json.JsonReaderWriterFactory]::CreateJsonReader($jsonBytes, [System.Xml.XmlDictionaryReaderQuotas]::Max)
    $stream = [System.IO.MemoryStream]::new()
    try {
        $writer = [System.Runtime.Serialization.Json.JsonReaderWriterFactory]::CreateJsonWriter($stream, [System.Text.Encoding]::UTF8, $false, $true, '    ')
        try {
            $writer.WriteNode($reader, $true)
            $writer.Flush()
        } finally {
            $writer.Dispose()
        }
        [System.IO.File]::WriteAllText($settingsPath, [System.Text.Encoding]::UTF8.GetString($stream.ToArray()), [System.Text.UTF8Encoding]::new($true))
    } finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Restart-BusinessCentralManager {
    $launcherPath = Join-Path ($PSScriptRoot | Split-Path) 'Start-BusinessCentralManager.cmd'
    Start-Process -FilePath $launcherPath -WindowStyle Hidden -ErrorAction Stop
    Exit
}



# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                      ### GUI Logic ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# ---------------------- #
### GLOBAL SETUP ###
# ---------------------- #

$script:appJob = $null
$script:appWorkerPath = Join-Path $PSScriptRoot 'modules\Invoke-BusinessCentralAppDeployment.ps1'
$script:appMainScriptPath = $PSCommandPath
$script:appTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:appTimer.Interval = [TimeSpan]::FromMilliseconds(250)

function Start-AppPublishingJob {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [ValidateSet('Single', 'Batch')] [string] $Operation,
        [Parameter(Mandatory = $true)] [string] $AppPath,
        [Parameter(Mandatory = $true)] [string] $ServerInstance,
        [Parameter(Mandatory = $true)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode] $SyncMode
    )

    if ($script:appJob) { throw 'An app operation is already running.' }
    $arguments = @($Operation, $AppPath, $ServerInstance, (Get-TargetTenant), $SyncMode.ToString(),
        (Resolve-NavAdminToolPath), [bool]$settings.settings.unpublishLastInstalledAppDuringUpgrade, $script:appMainScriptPath)
    $script:appJob = Start-Job -FilePath $script:appWorkerPath -ArgumentList $arguments -ErrorAction Stop
    $script:appOperation = $Operation
    $script:appTotal = if ($Operation -eq 'Single') { 1 } else { 0 }
    $script:appProcessed = 0
    $script:appFailed = 0
    $script:appLastResult = $null
    $script:appWarnings = @()
    $script:appError = $null

    $script:appControlStates = @(foreach ($control in @($var_AppPublishingSendAppBtn, $var_MultipleAppPublishingSendAppBtn,
                $var_AppPublishingChooseAppBtn, $var_MultipleAppPublishingChooseAppBtn,
                $var_AppPublishingServerInstanceComboBox, $var_MultipleAppPublishingServerInstanceComboBox,
                $var_AppPublishingSyncModeComboBox, $var_MultipleAppPublishingSyncModeComboBox,
                $var_SettingsSaveBtn, $var_CheckForUpdatesBtn, $var_LoadBcLicenseBtn)) {
            [pscustomobject]@{ Control = $control; LocalIsEnabled = $control.ReadLocalValue([System.Windows.UIElement]::IsEnabledProperty) }
        })
    foreach ($state in $script:appControlStates) {
        $state.Control.IsEnabled = $false
    }
    if ($Operation -eq 'Batch') {
        $var_MultipleAppPublishingAppStatusList.Items.Clear()
        $progress = $var_MultipleAppPublishingProgress
        $progressInfo = $var_MultipleAppPublishingProgressInfoTxt
        $progressBar = $var_MultipleAppPublishingProgressBar
        $progressInfo.Text = 'Checking app versions and dependencies...'
    } else {
        $progress = $var_AppPublishingProgress
        $progressInfo = $var_AppPublishingProgressInfoTxt
        $progressBar = $var_AppPublishingProgressBar
        $progressInfo.Text = 'Processing app...'
        $progressBar.IsIndeterminate = $true
    }
    $progress.IsEnabled = $true
    $progress.Visibility = [System.Windows.Visibility]::Visible
    $progressBar.Value = 0
    $script:appTimer.Start()
}

function Receive-AppPublishingJob {
    if (-not $script:appJob) { return }

    try {
        $jobWarnings = $null
        foreach ($result in @(Receive-Job -Job $script:appJob -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable jobWarnings)) {
            if ($result.Kind -eq 'Count') {
                $script:appTotal = [int]$result.Total
                $var_MultipleAppPublishingProgressInfoTxt.Text = "Processing 0 of $($script:appTotal) apps..."
            } elseif ($result.Kind -eq 'App') {
                $script:appProcessed++
                $script:appLastResult = $result
                if (-not $result.Succeeded) { $script:appFailed++ }
                if ($script:appOperation -eq 'Batch') {
                    $null = $var_MultipleAppPublishingAppStatusList.Items.Add([pscustomobject]@{
                        AppName = $result.Name
                        AppVersion = $result.Version
                        Status = if ($result.Succeeded) { '✔️' } else { '❌' }
                    })
                    $var_MultipleAppPublishingProgressBar.Value = 100 * $script:appProcessed / $script:appTotal
                    $var_MultipleAppPublishingProgressInfoTxt.Text = "Processed $($script:appProcessed) of $($script:appTotal) apps..."
                }
            }
        }
        if ($jobWarnings) {
            $script:appWarnings += @($jobWarnings | ForEach-Object { $_.Message })
        }
    } catch {
        $script:appError = $_.Exception.Message
        Stop-Job -Job $script:appJob -ErrorAction SilentlyContinue
    }

    if ($script:appJob.State -notin @('Completed', 'Failed', 'Stopped')) { return }
    if (-not $script:appError -and $script:appJob.State -ne 'Completed') {
        $reason = $script:appJob.ChildJobs[0].JobStateInfo.Reason
        $script:appError = if ($reason) { $reason.Message } else { "The app job $($script:appJob.State)." }
    }
    if (-not $script:appError -and ($script:appTotal -eq 0 -or $script:appProcessed -ne $script:appTotal)) {
        $script:appError = 'The app job ended without reporting a result for every app.'
    }

    $script:appTimer.Stop()
    if ($script:appError) {
        $message = "Processed $($script:appProcessed) of $($script:appTotal) apps before publishing stopped.`n$($script:appError)"
        $title = 'Publishing failed'
        $icon = 'Error'
    } elseif ($script:appOperation -eq 'Batch') {
        $message = "Processed $($script:appProcessed) apps; $($script:appFailed) failed or were skipped."
        $title = if ($script:appFailed) { 'Completed with issues' } else { 'Success' }
        $icon = if ($script:appFailed) { 'Warning' } else { 'Asterisk' }
    } elseif ($script:appFailed) {
        $message = "Could not complete $($script:appLastResult.Name) $($script:appLastResult.Version)."
        $title = 'Publishing skipped or failed'
        $icon = 'Warning'
    } else {
        $message = "Successfully processed $($script:appLastResult.Name) $($script:appLastResult.Version)."
        $title = 'Success'
        $icon = 'Asterisk'
    }
    if ($script:appWarnings.Count -gt 0) {
        $message += "`n`n" + (($script:appWarnings | Select-Object -First 5) -join "`n")
        if ($script:appWarnings.Count -gt 5) {
            $message += "`n...and $($script:appWarnings.Count - 5) more warnings."
        }
        if (-not $script:appError -and -not $script:appFailed) {
            $title = 'Completed with warnings'
            $icon = 'Warning'
        }
    }

    try {
        Remove-Job -Job $script:appJob -Force -ErrorAction SilentlyContinue
    } finally {
        $script:appJob = $null
        foreach ($state in $script:appControlStates) {
            if ([object]::ReferenceEquals($state.LocalIsEnabled, [System.Windows.DependencyProperty]::UnsetValue)) {
                $state.Control.ClearValue([System.Windows.UIElement]::IsEnabledProperty)
            } else {
                $state.Control.SetValue([System.Windows.UIElement]::IsEnabledProperty, $state.LocalIsEnabled)
            }
        }
        $script:appControlStates = @()
        foreach ($progress in @($var_AppPublishingProgress, $var_MultipleAppPublishingProgress)) {
            $progress.IsEnabled = $false
            $progress.Visibility = [System.Windows.Visibility]::Hidden
        }
        $var_AppPublishingProgressBar.IsIndeterminate = $false
        $var_AppPublishingProgressBar.Value = 0
        $var_MultipleAppPublishingProgressBar.Value = 0
        $var_AppPublishingProgressInfoTxt.Text = ''
        $var_MultipleAppPublishingProgressInfoTxt.Text = ''
    }
    return [pscustomobject]@{ Message = $message; Title = $title; Icon = $icon }
}

$script:appTimer.Add_Tick({
    $summary = Receive-AppPublishingJob
    if ($summary) {
        [System.Windows.Forms.MessageBox]::Show($summary.Message, $summary.Title, 'OK', $summary.Icon) | Out-Null
    }
})
$window.Add_Closing({
    param($windowSender, $closingArgs)

    if ($script:containerInvocation -and -not $script:containerInvocation.IsCompleted) {
        $confirmation = [System.Windows.Forms.MessageBox]::Show('BCContainerHelper is still loading. Closing may interrupt an update or installation. Close anyway?', 'Module loading in progress', 'YesNo', 'Warning')
        if ($confirmation -ne 'Yes') {
            $closingArgs.Cancel = $true
            return
        }
    }
    if ($script:appJob) {
        if ($script:appJob.State -notin @('Completed', 'Failed', 'Stopped')) {
            $confirmation = [System.Windows.Forms.MessageBox]::Show('App publishing is still running. Closing may interrupt it. Close anyway?', 'Publishing in progress', 'YesNo', 'Warning')
            if ($confirmation -ne 'Yes') {
                $closingArgs.Cancel = $true
                return
            }
            Stop-Job -Job $script:appJob -ErrorAction SilentlyContinue
        }
        Remove-Job -Job $script:appJob -Force -ErrorAction SilentlyContinue
        $script:appJob = $null
    }
    if ($script:containerPowerShell) {
        $script:containerTimer.Stop()
        try {
            if (-not $script:containerInvocation.IsCompleted) { $script:containerPowerShell.Stop() }
        } finally {
            $script:containerPowerShell.Dispose()
            $script:containerRunspace.Dispose()
            $script:containerPowerShell = $null
            $script:containerRunspace = $null
            $script:containerInvocation = $null
        }
    } elseif ($script:containerRunspace) {
        $script:containerRunspace.Dispose()
        $script:containerRunspace = $null
    }
    $script:appTimer.Stop()
})

<# ---- Main window loaded ---- #>

$window.Add_Loaded({
    if (-not $dependencyFailures.ContainsKey('NavAdminTool')) {
        Update-ServerInstanceOptions
    }

    # Logic for topic buttons and their tabs 
    foreach ($topic in $var_TopicsContainer.Children) {
        $topic.add_Checked({
             param($topicSender)

             if ($topicSender.Name -eq 'ContainerManagementTopic' -and
                 -not $dependencyFailures.ContainsKey('BcContainerHelper') -and
                 -not $script:containerModuleReady) {
                 Start-ContainerModuleImport
                 return
             }

             if (Get-TopicDependencyFailure -TopicName $topicSender.Name) {
                 $script:activeTopic.IsChecked = $true
                 if (Resolve-TopicDependencies -Topic $topicSender) {
                     $topicSender.IsChecked = $true
                 }
                 return
             }

             $script:activeTopic = $topicSender
             foreach ($view in $var_TopicsTabContainer.Children) {
                if ($view.Name -eq $topicSender.Tag) {
                    Set-UIElement -Element $view -Property "Visibility" -Value 0 # Visible
                    Set-UIElement -Element $view -Property "IsEnabled" -Value $true
                } else {
                    Set-UIElement -Element $view -Property "Visibility" -Value 2 # Hidden
                    Set-UIElement -Element $view -Property "IsEnabled" -Value $false
                }
             }
        })
    }

    $script:activeTopic = $var_AppManagementTopic
    Set-TopicAvailability
    if (Get-TopicDependencyFailure -TopicName $var_AppManagementTopic.Name) {
        $var_SettingsTopic.IsChecked = $true
    }

    # Load settings from settings.json to be displayed
    foreach ($key in $settings.settings.PSObject.Properties) {
        $controlName = $key.Name
        $value = $key.Value
        Set-SettingControlValue -SettingName $controlName -Value $value
    }

    # Load information in About tab
    $var_AboutTabDisplayVersionTxt.Text = $settings.settings.applicationVersion
})


# ------------------------------#
###   APP MANAGEMENT TOPIC   ###
# ------------------------------#

<# ---- App Publishing ---- #>

$var_AppPublishingChooseAppBtn.Add_Click({
   $AppPath = Select-File -FileFilter "Business Central Extension (*.app)|*.app"
   
   if ([string]::IsNullOrEmpty($AppPath)) {
        return
   }

   $var_AppPublishingAppPathTxt.Text = $AppPath

   try {
        $AppInfo = Get-NAVAppInfo -Path $var_AppPublishingAppPathTxt.Text -ErrorAction Stop
   }
   catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        return     
   }

    $var_AppPublishingAppIdTxt.Text = $AppInfo.Id
   $var_AppPublishingAppNameTxt.Text = $AppInfo.Name
   $var_AppPublishingAppVersionTxt.Text = $AppInfo.Version
   $var_AppPublishingAppPublisherTxt.Text = $AppInfo.Publisher
})

$var_AppPublishingSendAppBtn.Add_Click({
    # Test mandatory fields
    if ([string]::IsNullOrEmpty($var_AppPublishingServerInstanceComboBox.SelectedValue)) {
        [System.Windows.Forms.MessageBox]::Show("There is no server instance selected." , "Error", "OK", "Error")
        return    
    }

    if ([string]::IsNullOrEmpty($var_AppPublishingAppPathTxt.Text)) {
        [System.Windows.Forms.MessageBox]::Show("You did not select any app for publishing.", "Error", "OK", "Error")
        return    
    }

    try {
        $appPath = $var_AppPublishingAppPathTxt.Text
        $appInfo = Get-NAVAppInfo -Path $appPath -ErrorAction Stop
        $syncMode = Get-SyncMode -SyncModeComboBox $var_AppPublishingSyncModeComboBox
        $confirmation = [System.Windows.Forms.MessageBox]::Show(("Process {0} version {1} on {2}?" -f $appInfo.Name, $appInfo.Version, $var_AppPublishingServerInstanceComboBox.SelectedValue), 'Confirm app publishing', 'YesNo', 'Question')
        if ($confirmation -ne 'Yes') { return }
        if ($syncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
            $confirmation = [System.Windows.Forms.MessageBox]::Show(("Process the app using sync mode {0}?`n`nThis could be destructive." -f $syncMode), 'Confirm app sync mode', 'YesNo', 'Warning')
            if ($confirmation -ne 'Yes') { return }
        }
        Start-AppPublishingJob -Operation Single -AppPath $appPath -ServerInstance $var_AppPublishingServerInstanceComboBox.SelectedValue -SyncMode $syncMode
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Publishing could not start', 'OK', 'Error') | Out-Null
    }
})


<# ---- Multiple App Publishing ---- #>

$var_MultipleAppPublishingChooseAppBtn.Add_Click({
    $AppPath = Select-Folder
       
    if ([string]::IsNullOrEmpty($AppPath)) {
        return
    }

    $var_MultipleAppPublishingAppPathTxt.Text = $AppPath
})

$var_MultipleAppPublishingSendAppBtn.Add_Click({
    # Test mandatory fields
    if ([string]::IsNullOrEmpty($var_MultipleAppPublishingServerInstanceComboBox.SelectedValue)) {
        [System.Windows.Forms.MessageBox]::Show("There is no server instance selected." , "Error", "OK", "Error")
        return    
    }

    if ([string]::IsNullOrEmpty($var_MultipleAppPublishingAppPathTxt.Text)) {
        [System.Windows.Forms.MessageBox]::Show("You did not select any app for publishing.", "Error", "OK", "Error")
        return    
    }

    try {
        $appPath = $var_MultipleAppPublishingAppPathTxt.Text
        $syncMode = Get-SyncMode -SyncModeComboBox $var_MultipleAppPublishingSyncModeComboBox
        $confirmation = [System.Windows.Forms.MessageBox]::Show(("Process all apps in {0}?`n`nVersions must not be older than those already published or installed." -f $appPath), 'Confirm batch publishing', 'YesNo', 'Question')
        if ($confirmation -ne 'Yes') { return }
        if ($syncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
            $confirmation = [System.Windows.Forms.MessageBox]::Show(("Process all apps using sync mode {0}?`n`nThis could be destructive." -f $syncMode), 'Confirm app sync mode', 'YesNo', 'Warning')
            if ($confirmation -ne 'Yes') { return }
        }
        Start-AppPublishingJob -Operation Batch -AppPath $appPath -ServerInstance $var_MultipleAppPublishingServerInstanceComboBox.SelectedValue -SyncMode $syncMode
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Batch publishing could not start', 'OK', 'Error') | Out-Null
    }
})


# ------------------------------ #
### SERVER MANAGEMENT TOPIC ###
# ------------------------------ #

# License Management

$var_LicenseChooseBtn.Add_Click({
   $LicensePath = Select-File -FileFilter "Business Central License (*.flf; *.bclicense)|*.flf;*.bclicense"
   
   if ([string]::IsNullOrEmpty($LicensePath)) {
        return
   }

   $var_LicensePathTxt.Text = $LicensePath
})

$var_LoadBcLicenseBtn.Add_Click({
    # Test mandatory fields
    if ([string]::IsNullOrEmpty($var_LicenseServerInstanceComboBox.SelectedValue)) {
        [System.Windows.Forms.MessageBox]::Show("There is no server instance selected." , "Error", "OK", "Error")
        return    
    }

    if ([string]::IsNullOrEmpty($var_LicensePathTxt.Text)) {
        [System.Windows.Forms.MessageBox]::Show("There is no license file selected." , "Error", "OK", "Error")
        return    
    }

    $ConfirmLoadLicenseAndRestartInstance = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to load the selected license and restart BC server instance {0}? " -f $var_LicenseServerInstanceComboBox.SelectedValue), "Confirm App Sync Mode", "YesNo", "Warning")

    if ($ConfirmLoadLicenseAndRestartInstance -eq "No") {
        return
    }

    try {
        Import-NAVServerLicense -LicenseFile $var_LicensePathTxt.Text -ServerInstance $var_LicenseServerInstanceComboBox.SelectedValue -ErrorAction Stop
        Restart-NAVServerInstance -ServerInstance $var_LicenseServerInstanceComboBox.SelectedValue -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Successfully applied license file.", "Succcess", "OK", "Asterisk") | Out-Null
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        return
    }
})

$var_GetCurrentBcLicenseInfoBtn.Add_Click({
    # Test mandatory fields
    if ([string]::IsNullOrEmpty($var_LicenseServerInstanceComboBox.SelectedValue)) {
        [System.Windows.Forms.MessageBox]::Show("There is no server instance selected." , "Error", "OK", "Error")
        return    
    }

    $currentLicenseInfo = Export-NAVServerLicenseInformation -ServerInstance $var_LicenseServerInstanceComboBox.SelectedValue | Out-String
    $licenseWindow = [System.Windows.Window]::new()
    $licenseWindow.Title = 'Current Business Central license'
    $licenseWindow.Owner = $window
    $licenseWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $licenseWindow.Width = 760
    $licenseWindow.Height = 460
    $licenseWindow.FontFamily = $window.FontFamily

    $licenseText = [System.Windows.Controls.TextBox]::new()
    $licenseText.Text = $currentLicenseInfo
    $licenseText.IsReadOnly = $true
    $licenseText.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $licenseText.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $licenseWindow.Content = $licenseText
    $null = $licenseWindow.ShowDialog()
})


# ------------------------------#
###      SETTINGS TOPIC      ###
# ------------------------------#

$var_NavAdminToolBrowseBtn.Add_Click({
    try {
        $initialDirectory = Join-Path $env:ProgramFiles 'Microsoft Dynamics 365 Business Central'
        $currentPath = $var_NavAdminTool.Text.Trim().Trim('"')
        if ($currentPath) {
            $currentDirectory = Split-Path -Path $currentPath -Parent
            if ($currentDirectory -and (Test-Path -LiteralPath $currentDirectory -PathType Container)) {
                $initialDirectory = $currentDirectory
            }
        }

        $dialogOwner = [BcManagerDialogOwner]::new(([System.Windows.Interop.WindowInteropHelper]::new($window).Handle))
        $toolPath = Select-File -FileFilter 'Business Central administration tool (NavAdminTool.ps1)|NavAdminTool.ps1' -Directory $initialDirectory -Title 'Select NavAdminTool.ps1' -Owner $dialogOwner
        if ([string]::IsNullOrEmpty($toolPath)) { return }
        if ([System.IO.Path]::GetFileName($toolPath) -ine 'NavAdminTool.ps1') {
            throw 'Select the file named NavAdminTool.ps1 from your Business Central installation.'
        }

        $var_NavAdminTool.Text = $toolPath
        $null = $var_NavAdminTool.Focus()
        $var_NavAdminTool.CaretIndex = $toolPath.Length
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Could not select NavAdminTool', 'OK', 'Error') | Out-Null
    }
})

$var_SettingsSaveBtn.Add_Click({
    $navAdminToolPath = $var_NavAdminTool.Text.Trim().Trim('"')
    if ($navAdminToolPath -and
        (-not (Test-Path -LiteralPath $navAdminToolPath -PathType Leaf -ErrorAction SilentlyContinue) -or
         [System.IO.Path]::GetFileName($navAdminToolPath) -ine 'NavAdminTool.ps1')) {
        [System.Windows.Forms.MessageBox]::Show('Choose an existing NavAdminTool.ps1 file, or leave the path blank to detect the installation automatically.', 'Invalid NavAdminTool path', 'OK', 'Warning') | Out-Null
        $null = $var_NavAdminTool.Focus()
        return
    }
    $var_NavAdminTool.Text = $navAdminToolPath

    $ConfirmSaveSettings = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to update application settings and restart?", "Confirm Save and Restart", "YesNo", "Warning")
    
    if ($ConfirmSaveSettings -eq "No") {
        return
    }

    # Load settings from settings.json variable that has been changed to be saved
    foreach ($key in $settings.settings.PSObject.Properties) {
        Save-ApplicationSetting -SettingName $key.Name -PreviousValue $key.Value
    }   

    # Restart the app
    Restart-BusinessCentralManager
})


# ------------------------------ #
###        ABOUT TOPIC        ###
# ------------------------------ #

$var_CheckForUpdatesBtn.Add_Click({
    try {
        Invoke-BusinessCentralManagerUpdateCheck -ShowUpToDateMessage $true
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }
})

$var_GitHubLink.Add_Click({
    # Get the URL from the NavigateUri property of the Hyperlink control
    $url = $var_GitHubLink.NavigateUri.AbsoluteUri

    # Open the URL in the default web browser
    Start-Process $url
})

$var_ShowDocumentationBtn.Add_Click({
    # Get the path to the manual
    $manualPath = Join-Path ($PSScriptRoot | Split-Path) 'docs\Business Central Manager - Manual.pdf'

    # Open the manual with the default PDF viewer
    Start-Process $manualPath
})


$Null = $window.ShowDialog()
