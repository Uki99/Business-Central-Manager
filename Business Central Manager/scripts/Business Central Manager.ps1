

                                                                                    ### Setup Section ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# Check for admin rights and ask if needed
if(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList "-File `"$($MyInvocation.MyCommand.Path)`"  `"$($MyInvocation.MyCommand.UnboundArguments)`""
    Exit
}

# Set execution policy and import required assemblies
Set-ExecutionPolicy Unrestricted
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework


# Start Initializer.ps1
$scriptPath = ($PSScriptRoot + "\Initializer.ps1")
$Initializer = Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList "-File `"$($scriptPath)`"" -Wait -PassThru

# Code 200 is a custom exit code that indicates update has happened, the app has been restarted so the initial parent process can close
if ($Initializer.ExitCode -eq 200) {
    Exit
}

# Load settings from json
try {
    $settings = Get-Content (($PSScriptRoot | Split-Path) + "\data\settings.json") -Raw | ConvertFrom-Json -ErrorAction Stop
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

# Import required NavAdminTool module     
try {
    if (-not (Get-Module -Name NavAdminTool)) {
        Import-Module -Name $settings.settings.NavAdminTool -ErrorAction Stop
    }            
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

# Setup MainWindow.xaml for GUI
try {
    $inputXML = Get-Content (($PSScriptRoot | Split-Path) + $settings.settings.MainWindowXAMLRelativePath) -Raw -ErrorAction Stop
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$MainWindowXAML = $inputXML

$reader = (New-Object System.Xml.XmlNodeReader $MainWindowXAML)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

# Set main window icon
$IconPath = (($PSScriptRoot | Split-Path) + "\data\mainIcon.ico")

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
$MainWindowXAML.SelectNodes("//*[@Name]") | ForEach-Object {
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }
}

# Import BCContainerHelper module
if (-not $settings.settings.DelayBcContainerHelperModuleImport) {
    try {
        Import-Module -Name BcContainerHelper -ErrorAction Stop -Verbose
    }
    catch {
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
        [string]$Directory = ([environment]::GetFolderPath('Desktop'))
    )

    $OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = (Resolve-Path $Directory).Path
    $OpenFileDialog.RestoreDirectory = $true
    $OpenFileDialog.Filter = $FileFilter

    $result = $OpenFileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if (Test-Path $OpenFileDialog.FileName -PathType Leaf) {
            return $OpenFileDialog.FileName
        }
    }

    return $null
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

function Get-SyncMode {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [System.Windows.Controls.ComboBox] $SyncModeXamlCtrlComboBox
    )

    $SelectedSyncItem = $SyncModeXamlCtrlComboBox.SelectedItem
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

function UpdateUIElement {
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

function Install-App {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $ServerInstance,
        [Parameter(Mandatory = $true)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode] $SyncMode = [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add,
        [Parameter(Mandatory = $true)] [boolean] $SupressGui,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.StackPanel] $ProgressBarContainer,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.ProgressBar] $ProgressBar,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.TextBlock] $ProgressInfo
    )

    if ($PSBoundParameters.ContainsKey('ProgressBarContainer') -and $PSBoundParameters.ContainsKey('ProgressBar') -and $PSBoundParameters.ContainsKey('ProgressInfo')) {
        $ShouldHandleProgressBar = $true
    } else {
        $ShouldHandleProgressBar = $false
    }
    
    $TargetAppInfo = Get-NAVAppInfo -Path $Path -ErrorAction Stop

    if (-not $SupressGui) {
        $ConfirmAppInstall = [System.Windows.Forms.MessageBox]::Show(("Do you want to try and install the following app?`n`nName: {0}`nVersion: {1}" -f $TargetAppInfo.Name, $TargetAppInfo.Version), "Confirm App Install", "YesNo", "Question")      
        if ($ConfirmAppInstall -eq "No") {
            return
        }

        if ($SyncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
            $ConfirmSyncMode = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to install the app using sync mode = {0}?`n`nThis could be destructive." -f $SyncMode), "Confirm App Sync Mode", "YesNo", "Warning") | Out-Null
        
            if ($ConfirmSyncMode -eq "No") {
                return
            }
        }
    }


    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $true
        UpdateUIElement -Element $ProgressBarContainer -Property "Visibility" -Value 0 # Visible
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Publishing app..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 25
    }
    Publish-NAVApp -ServerInstance $ServerInstance -Path $Path -SkipVerification -ErrorAction Stop

    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Synchronizing app..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 50 
    }
    Sync-NavApp -ServerInstance $ServerInstance -Name $TargetAppInfo.Name -Version $TargetAppInfo.Version -Mode $SyncMode -Force -ErrorAction Stop    
    
    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Installing app..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 75
    }
    try {
        Install-NAVApp -ServerInstance $ServerInstance -Name $TargetAppInfo.Name -Version $TargetAppInfo.Version -Force -ErrorAction Stop
    } catch [InvalidOperationException] {
        <# This is to handle installing the apps which were uninstalled without clean mode but recognized as in need of install #>
        Start-NAVAppDataUpgrade -ServerInstance $ServerInstance -Name $TargetAppInfo.Name -Version $TargetAppInfo.Version -Force -ErrorAction Stop
    }

    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Finalizing..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 100
    }
    if(-not $SupressGui) {
        [System.Windows.Forms.MessageBox]::Show(("Successfully installed the extension {0} with version {1}!" -f $TargetAppInfo.Name, $TargetAppInfo.Version), "Succcess", "OK", "Asterisk") | Out-Null
    }

    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $false
        UpdateUIElement -Element $ProgressBarContainer -Property "Visibility" -Value 2 # Hidden
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value ""
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 0
    }
}

function Update-App {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [string] $Path,
        [Parameter(Mandatory = $true)] [string] $ServerInstance,
        [Parameter(Mandatory = $true)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode] $SyncMode = [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add,
        [Parameter(Mandatory = $true)] [boolean] $SupressGui,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.StackPanel] $ProgressBarContainer,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.ProgressBar] $ProgressBar,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.TextBlock] $ProgressInfo
    )

    if ($PSBoundParameters.ContainsKey('ProgressBarContainer') -and $PSBoundParameters.ContainsKey('ProgressBar') -and $PSBoundParameters.ContainsKey('ProgressInfo')) {
        $ShouldHandleProgressBar = $true
    } else {
        $ShouldHandleProgressBar = $false
    }

    $TargetAppInfo = Get-NAVAppInfo -Path $Path
    $OldAppInfo = Get-NAVAppInfo -ServerInstance $ServerInstance -Name $TargetAppInfo.Name -Tenant default -TenantSpecificProperties | Where-Object { $_.IsInstalled -eq $true }

    if (-not $SupressGui) {
        $ConfirmAppUpdate = [System.Windows.Forms.MessageBox]::Show(("Do you want to try and update the following app?`n`nName: {0}`nNew Version: {1}`nCurrent Version: {2}" -f $TargetAppInfo.Name, $TargetAppInfo.Version, $OldAppInfo.Version), "Confirm App Update", "YesNo", "Question")       
        if ($ConfirmAppUpdate -eq "No") {
            return
        }

        if ($SyncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
            $ConfirmSyncMode = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to install the app using sync mode = {0}?`n`nThis could be destructive." -f $SyncMode), "Confirm App Sync Mode", "YesNo", "Warning") | Out-Null
        
            if ($ConfirmSyncMode -eq "No") {
                return
            }
        }
    }


    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $true
        UpdateUIElement -Element $ProgressBarContainer -Property "Visibility" -Value 0 # Visible
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Publishing app..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 20
    }
    Publish-NAVApp -ServerInstance $ServerInstance -Path $Path -SkipVerification -ErrorAction Stop

    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Synchronizing app..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 40
    }
    Sync-NavApp -ServerInstance $ServerInstance -Name $TargetAppInfo.Name -Version $TargetAppInfo.Version -Mode $SyncMode -Force -ErrorAction Stop
    
    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Upgrading app..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 60
    }
    Start-NAVAppDataUpgrade -ServerInstance $ServerInstance -Name $TargetAppInfo.Name -Version $TargetAppInfo.Version -Force -ErrorAction Stop 


    if ($settings.settings.UnpublishLastInstalledAppDuringUpgrade)
    {
        if ($ShouldHandleProgressBar) {
            UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Unpublishing old app..."
            UpdateUIElement -Element $ProgressBar -Property "Value" -Value 80
        }
        Unpublish-NAVApp -ServerInstance $ServerInstance -Name $OldAppInfo.Name -Version $OldAppInfo.Version -ErrorAction Stop
    }

    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value "Finalizing..."
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 100
    }
    if (-not $SupressGui) {
        [System.Windows.Forms.MessageBox]::Show(("Successfully updated the extension {0} to the version {1}!" -f $TargetAppInfo.Name, $TargetAppInfo.Version), "Succcess", "OK", "Asterisk") | Out-Null
    }

    if ($ShouldHandleProgressBar) {
        UpdateUIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $false
        UpdateUIElement -Element $ProgressBarContainer -Property "Visibility" -Value 2 # Hidden        
        UpdateUIElement -Element $ProgressInfo -Property "Text" -Value ""
        UpdateUIElement -Element $ProgressBar -Property "Value" -Value 0
    }
}

function Process-App {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [string]$AppPath,
        [Parameter(Mandatory = $true)] [string]$ServerInstance,
        [Parameter(Mandatory = $true)] [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]$SyncMode,
        [Parameter(Mandatory = $false)] [boolean] $SupressGui = $false,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.StackPanel] $ProgressBarContainer,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.ProgressBar] $ProgressBar,
        [Parameter(Mandatory = $false)] [System.Windows.Controls.TextBlock] $ProgressInfo
    )

	$RequestedAppInformation = Get-NAVAppInfo -Path $AppPath
    $CurrentAppInformation = Get-NAVAppInfo -Name $RequestedAppInformation.Name -ServerInstance $ServerInstance -Tenant default -TenantSpecificProperties -ErrorAction Stop
    $CurrentInstalledAppInformation = $CurrentAppInformation | Where-Object { $_.IsInstalled -eq $true }

    $IsRequestedAppInstalled = [bool]$CurrentInstalledAppInformation.Count
    $NewerVersionExists      = [bool](($CurrentAppInformation | Where-Object { [version]$_.Version -gt [version]$RequestedAppInformation.Version }) | Measure-Object).Count
    $SameVersionExists       = [bool](($CurrentAppInformation | Where-Object { [version]$_.Version -eq [version]$RequestedAppInformation.Version }) | Measure-Object).Count
    $OlderVersionsExists     = [bool](($CurrentAppInformation | Where-Object { [version]$_.Version -lt [version]$RequestedAppInformation.Version }) | Measure-Object).Count

    if ($PSBoundParameters.ContainsKey('ProgressBarContainer') -and $PSBoundParameters.ContainsKey('ProgressBar') -and $PSBoundParameters.ContainsKey('ProgressInfo')) {
        $ShouldHandleProgressBar = $true
    } else {
        $ShouldHandleProgressBar = $false
    }

    <# Requested app is already installed under any version #>
    if ($IsRequestedAppInstalled) {
        if ($NewerVersionExists) {
            [System.Windows.Forms.MessageBox]::Show(("The requested app {0} has a newer version already published or installed.`n`nVisit page `"Extension Management`" and install it from there if it is not yet installed.`n`nNo action has been performed." -f $RequestedAppInformation.Name), "Newer App Available", "OK", "Error") | Out-Null                 
            return
        }

        if ($SameVersionExists) {
            [System.Windows.Forms.MessageBox]::Show(("The requested app {0} has the same version {1} already published or installed.`n`nVisit page `"Extension Management`" and install it from there if it is not yet installed.`n`nNo action has been performed." -f $RequestedAppInformation.Name, $RequestedAppInformation.Version), "App Already Available", "OK", "Information") | Out-Null              
            return
        }
        

        if ($ShouldHandleProgressBar) {
            try {
                Update-App -Path $AppPath -ServerInstance $ServerInstance -SyncMode $SyncMode -ProgressBarContainer $ProgressBarContainer -ProgressBar $ProgressBar -ProgressInfo $ProgressInfo -SupressGui $SupressGui
                return     
            }
            catch {
                $errorMessage = $_.ToString()
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")

                UpdateUIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $false
                UpdateUIElement -Element $ProgressBarContainer -Property "Visibility" -Value 2 # Hidden
                UpdateUIElement -Element $ProgressInfo -Property "Text" -Value ""
                UpdateUIElement -Element $ProgressBar -Property "Value" -Value 0

                return
            }
        } else {
            try {
                Update-App -Path $AppPath -ServerInstance $ServerInstance -SyncMode $SyncMode -SupressGui $SupressGui
                return     
            }
            catch {
                $errorMessage = $_.ToString()
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
                return              
            }
        }
    }   

    <# Requested app is not installed under any version #>
    if(-not $IsRequestedAppInstalled) {
        if ($NewerVersionExists) {
            [System.Windows.Forms.MessageBox]::Show(("The requested app {0} has a newer version already published.`n`nVisit page `"Extension Management`" and install it from there. No action has been performed." -f $RequestedAppInformation.Name), "Newer App Available", "OK", "Error") | Out-Null                
            return
        }

        if ($SameVersionExists) {
            [System.Windows.Forms.MessageBox]::Show(("The requested app {0} is already published with the same version {1}.`n`n`nVisit page `"Extension Management`" and install it from there. No action has been performed." -f $RequestedAppInformation.Name, $RequestedAppInformation.Version), "App Already Published", "OK", "Information") | Out-Null               
            return
        }


        if ($ShouldHandleProgressBar) {
            try {
                Install-App -Path $AppPath -ServerInstance $ServerInstance -SyncMode $SyncMode -ProgressBarContainer $ProgressBarContainer -ProgressBar $ProgressBar -ProgressInfo $ProgressInfo -SupressGui $SupressGui
            }
            catch {
                $errorMessage = $_.ToString()
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")

                UpdateUIElement -Element $ProgressBarContainer -Property "IsEnabled" -Value $false
                UpdateUIElement -Element $ProgressBarContainer -Property "Visibility" -Value 2 # Hidden
                UpdateUIElement -Element $ProgressInfo -Property "Text" -Value ""
                UpdateUIElement -Element $ProgressBar -Property "Value" -Value 0

                return
            }
        } else {
            try {
                Install-App -Path $AppPath -ServerInstance $ServerInstance -SyncMode $SyncMode -SupressGui $SupressGui
            }
            catch {
                $errorMessage = $_.ToString()
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
                return
            }
        }
    }
}

function Load-DisplaySettings {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $JsonKey,
        [Parameter(Mandatory = $true)] $Value
    )
    
    try {
        # Get variable in XAML that corresponds to the JSON key in settings.json
        $control = (Get-Variable -Name "var_$JsonKey" -ErrorAction SilentlyContinue).Value

        if ($control -is [System.Windows.Controls.TextBox]) {
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

function Update-ApplicationSettings {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] $JsonKey,
        [Parameter(Mandatory = $true)] $OldValue
    )

    try {
        # Get variable in XAML that corresponds to the JSON key in settings.json
        $ControlNewValue = (Get-Variable -Name "var_$JsonKey" -ErrorAction Stop).Value

        # Parse new values depending on the control type
        if ($ControlNewValue -is [System.Windows.Controls.TextBox]) {
            $NewValue = $controlNewValue.Text
        }  elseif ($ControlNewValue -is [System.Windows.Controls.CheckBox]) {
            $NewValue = $ControlNewValue.IsChecked          
        }
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }

    # Save to Json on disk
    if ($OldValue -eq $NewValue) {
        return
    }

    $settings.settings.$JsonKey = $NewValue
    $settings | ConvertTo-Json | Set-Content -Path (($PSScriptRoot | Split-Path) + "\data\settings.json")
}

function Restart-BusinessCentralManager {
    $batchScriptPath = (($PSScriptRoot | Split-Path) + "\scripts\Autorun.bat")
    Start-Process -FilePath $batchScriptPath
    Exit
}



# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#


                                                                                      ### GUI Logic ###
# ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# ---------------------- #
### GLOBAL SETUP ###
# ---------------------- #

<# ---- Main window loaded ---- #>

$window.Add_Loaded({
    # Load available server instances
    $Instances = Get-NAVServerInstance | Where-Object { ($_.State -eq 'Running') -and ($_.Version -match $settings.settings.SupportedBusinessCentralVersion) }

    if (($Instances | Measure-Object).Count -eq 1) {
        $position = $Instances.ServerInstance.IndexOf("$")
        $shortName = $Instances.ServerInstance.Substring($position + 1)

        $var_AppPublishingServerInstanceComboBox.Items.Add($shortName)
        $var_MultipleAppPublishingServerInstanceComboBox.Items.Add($shortName)
        $var_LicenseServerInstanceComboBox.Items.Add($shortName)
    } else {
        $Instances.ForEach({
            $position = $_.ServerInstance.IndexOf("$")
            $shortName = $_.ServerInstance.Substring($position + 1)

            $var_AppPublishingServerInstanceComboBox.Items.Add($shortName)
            $var_MultipleAppPublishingServerInstanceComboBox.Items.Add($shortName)
            $var_LicenseServerInstanceComboBox.Items.Add($shortName)
        })
    }

    # Logic for topic buttons and their tabs 
    foreach ($topic in $var_TopicsContainer.Children) {
        $topic.add_Checked({
             param(
                   $sender, 
                   $eventArgs
             )

             foreach ($tabControl in $var_TopicsTabContainer.Children) {
                if (($tabControl.Name -eq $sender.Tag) -and ($tabControl -is [System.Windows.Controls.TabControl])) {
                    UpdateUIElement -Element $tabControl -Property "Visibility" -Value 0 # Visible
                    UpdateUIElement -Element $tabControl -Property "IsEnabled" -Value $true
                } else {
                    UpdateUIElement -Element $tabControl -Property "Visibility" -Value 2 # Hidden
                    UpdateUIElement -Element $tabControl -Property "IsEnabled" -Value $false
                }
             }
        })
    }

    # Load settings from settings.json to be displayed
    foreach ($key in $settings.settings.PSObject.Properties) {
        $controlName = $key.Name
        $value = $key.Value
        Load-DisplaySettings -JsonKey $controlName -Value $value
    }

    # Load information in About tab
    $var_AboutTabDisplayVersionTxt.Text = $settings.settings.ApplicationVersion
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

   $var_AppPublishingAppIdTxt.Text = $AppInfo.AppId
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

    $SyncMode = Get-SyncMode -SyncModeXamlCtrlComboBox $var_AppPublishingSyncModeComboBox

    Process-App -AppPath $var_AppPublishingAppPathTxt.Text `
                -ServerInstance $var_AppPublishingServerInstanceComboBox.SelectedValue `
                -SyncMode $SyncMode `
                -ProgressBarContainer $var_AppPublishingProgress `
                -ProgressBar $var_AppPublishingProgressBar `
                -ProgressInfo $var_AppPublishingProgressInfoTxt
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
    $var_MultipleAppPublishingAppStatusList.Items.Clear()

    # Test mandatory fields
    if ([string]::IsNullOrEmpty($var_MultipleAppPublishingServerInstanceComboBox.SelectedValue)) {
        [System.Windows.Forms.MessageBox]::Show("There is no server instance selected." , "Error", "OK", "Error")
        return    
    }

    if ([string]::IsNullOrEmpty($var_MultipleAppPublishingAppPathTxt.Text)) {
        [System.Windows.Forms.MessageBox]::Show("You did not select any app for publishing.", "Error", "OK", "Error")
        return    
    }

	# Show confirmation message
    $ConfirmMultipleAppInstall = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to install the apps located at folder path {0}?`n`nAll apps must be greater in version if such app already exists in any form on target Business Central." -f $var_MultipleAppPublishingAppPathTxt.Text), "Confirm Multiple App Instal", "YesNo", "Question")        
    if ($ConfirmMultipleAppInstall -eq "No") {
        return
    }

    # Get Sync Mode
    $SyncMode = Get-SyncMode -SyncModeXamlCtrlComboBox $var_MultipleAppPublishingSyncModeComboBox
    if ($SyncMode -ne [Microsoft.Dynamics.Nav.Types.NavAppSyncMode]::Add) {
        $ConfirmSyncMode = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to install all apps using sync mode = {0}?`n`nThis could be destructive." -f $SyncMode), "Confirm App Sync Mode", "YesNo", "Warning") | Out-Null
        
        if ($ConfirmSyncMode -eq "No") {
            return
        }
    }

    # Initialize Progress bar
    UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "IsEnabled" -Value $true
    UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "Visibility" -Value 0 # Visible
    UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" -Value "Testing if apps are valid for install..."
    UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" -Value 0

    # Get all apps in folder
    $AppsPath = Get-ChildItem -Path $var_MultipleAppPublishingAppPathTxt.Text *.app
    $ProgressBarIncrementationValue = 100 / ($AppsPath.Count + 1)

    if ($AppsPath.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("There are no apps for publishing in specified folder.", "Error", "OK", "Error")

        # Reset Progress bar
        UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "IsEnabled" -Value $false
        UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "Visibility" -Value 2 # Hidden
        UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" -Value ""
        UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" -Value 0

        return
    }

	# Check app versions and validate
    foreach($AppPath in $AppsPath) {
        $AppInfo = Get-NAVAppInfo -Path $AppPath.FullName
        $CurrentAppInformation = Get-NAVAppInfo -Name $AppInfo.Name -ServerInstance $var_MultipleAppPublishingServerInstanceComboBox.SelectedValue -Tenant default -TenantSpecificProperties -ErrorAction Stop

        $NewerOrSameVersionExists = [bool](($CurrentAppInformation | Where-Object { [version]$_.Version -ge [version]$AppInfo.Version }) | Measure-Object).Count

        if ($NewerOrSameVersionExists) {
            $ErrorLog += ("App with name {0} and version {1} already has newer or same version published or installed.`n`n" -f $AppInfo.Name, $AppInfo.Version)
        }
    }

    if (-not [string]::IsNullOrEmpty($ErrorLog)) {
        $ErrorLog += "`nProvide newer version for these apps and try again."
        [System.Windows.Forms.MessageBox]::Show($ErrorLog, "Error", "OK", "Error")

        # Reset UI elements
        UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "IsEnabled" -Value $false
        UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "Visibility" -Value 2 # Hidden
        UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" -Value ""
        UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" -Value 0
        
        return
    }


    UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" -Value "Sorting apps by dependency order..."
    UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" -Value ($var_MultipleAppPublishingProgressBar.Value + $ProgressBarIncrementationValue)

    foreach ($App in Sort-AppFilesByDependencies -appFiles $AppsPath.FullName) {
		$AppInfo = Get-NAVAppInfo -Path $App

		# Update Progress bar
		UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" "Processing app `"$($AppInfo.Name)`""
		UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" ($var_MultipleAppPublishingProgressBar.Value + $ProgressBarIncrementationValue)

        try {
            Process-App -AppPath $App `
                        -ServerInstance $var_MultipleAppPublishingServerInstanceComboBox.SelectedValue `
                        -SyncMode $SyncMode `
                        -SupressGui $true

            $AppInstallInfo = [PSCustomObject] @{
                AppName = "$($AppInfo.Name)"
                AppVersion = "$($AppInfo.Version)"
                Status = "✔️"
            }
        } catch {
            $AppInstallInfo = [PSCustomObject] @{
                AppName = "$($AppInfo.Name)"
                AppVersion = "$($AppInfo.Version)"
                Status = "❌"
            }
        }

        $var_MultipleAppPublishingAppStatusList.Items.Add($AppInstallInfo)
    }
    	
    	 
    # Update Progress bar
	UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" -Value "Finalizing"
	UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" -Value ($var_MultipleAppPublishingProgressBar.Value + 100)

	[System.Windows.Forms.MessageBox]::Show("Successfully installed all apps!", "Succcess", "OK", "Asterisk") | Out-Null

	# Update Progress bar
	UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "IsEnabled" -Value $false
	UpdateUIElement -Element $var_MultipleAppPublishingProgress -Property "Visibility" -Value 2
	UpdateUIElement -Element $var_MultipleAppPublishingProgressInfoTxt -Property "Text" -Value ""
	UpdateUIElement -Element $var_MultipleAppPublishingProgressBar -Property "Value" -Value 0
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

    $ConfirmLoadLicenseAndRestartInstance = [System.Windows.Forms.MessageBox]::Show(("Are you sure you want to load the selected license and restart BC server instance {0}? " -f $var_LicenseServerInstanceComboBox.SelectedValue), "Confirm App Sync Mode", "YesNo", "Warning") | Out-Null

    if ($ConfirmLoadLicenseAndRestartInstance -eq "No") {
        return
    }

    try {
        Import-NAVServerLicense -LicenseFile $var_LicensePathTxt.Text -ServerInstance $var_LicenseServerInstanceComboBox.SelectedValue
        Restart-NAVServerInstance -ServerInstance $var_LicenseServerInstanceComboBox.SelectedValue
        [System.Windows.Forms.MessageBox]::Show("Successfully applied license file.", "Succcess", "OK", "Asterisk") | Out-Null
    } catch {
        $errorMessage = $_.ToString()
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
        Exit
    }
})

$var_GetCurrentBcLicenseInfoBtn.Add_Click({
    # Test mandatory fields
    if ([string]::IsNullOrEmpty($var_LicenseServerInstanceComboBox.SelectedValue)) {
        [System.Windows.Forms.MessageBox]::Show("There is no server instance selected." , "Error", "OK", "Error")
        return    
    }

    $CurrentLicenseInfo = Export-NAVServerLicenseInformation -ServerInstance $var_LicenseServerInstanceComboBox.SelectedValue | Out-String
    $Command = "`$CurrentLicenseInfo = `"$CurrentLicenseInfo`"; Write-Host `$CurrentLicenseInfo"
    $EncodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-EncodedCommand", $EncodedCommand
})


# ------------------------------#
###      SETTINGS TOPIC      ###
# ------------------------------#

$var_SettingsSaveBtn.Add_Click({
    $ConfirmSaveSettings = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to update application settings and restart?", "Confirm Save and Restart", "YesNo", "Warning")
    
    if ($ConfirmSaveSettings -eq "No") {
        return
    }

    # Load settings from settings.json variable that has been changed to be saved
    foreach ($key in $settings.settings.PSObject.Properties) {
        Update-ApplicationSettings -JsonKey $key.Name -OldValue $key.Value
    }   

    # Restart the app
    Restart-BusinessCentralManager
})


# ------------------------------ #
###        ABOUT TOPIC        ###
# ------------------------------ #

$var_CheckForUpdatesBtn.Add_Click({
    Import-Module -Force (($PSScriptRoot | Split-Path) + "\scripts\modules\BCManager-UpdateManagement.ps1")

    $owner = "Uki99"
    $repo = "Business-Central-Manager"
    
    try {
        Update-BCManagerApplication -owner $owner -repo $repo -currentVersion $settings.settings.ApplicationVersion -upToDateMessage $true
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


$Null = $window.ShowDialog()