# Load settings from json
try {
    $settings = Get-Content (($PSScriptRoot | Split-Path) + "\data\settings.json") -Raw | ConvertFrom-Json -ErrorAction Stop
}
catch {
    $errorMessage = $_.ToString()
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
    Exit
}

# Initialize console if neccesary
if ($settings.settings.HidePowerShellConsole) {
    $WindowStyleType = 1 # Hidden
} else {
    $WindowStyleType = 0 # Normal
}

# Start Business Central Manager with elevated rights
Start-Process -FilePath PowerShell.exe -WindowStyle $WindowStyleType -Verb Runas -ArgumentList "-File `"$($PSScriptRoot + "\BE-terna Business Central Manager.ps1")`"  `"$($MyInvocation.MyCommand.UnboundArguments)`""
Exit