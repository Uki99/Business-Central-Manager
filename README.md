# Business Central Manager

Business Central Manager is a PowerShell based GUI application used for Business Central managing with customizability and automatic updates. 

## Installation

Download the latest release from [releases](https://github.com/Uki99/Business-Central-Manager/releases/latest), unzip it and run `Start-BusinessCentralManager.cmd` from the `Business Central Manager` folder. Keep the folder structure intact so the launcher can find the scripts. The launcher hides its batch window after startup, though Windows may briefly flash a console when opening a CMD file. The `hidePowerShellConsole` setting still controls whether the application's own PowerShell window is shown.

## Configuration

In Settings, set the tenant ID for app publishing (the default is `default`). Use the ellipsis button beside **NavAdminTool path** to select `NavAdminTool.ps1` from the desired Business Central installation. Leave the path blank for automatic detection when exactly one installation is available. Set the version filter for your target (for example, `^28\.` for BC 28). If automatic unpublishing of previous app versions is enabled, it will only proceed when no tenant still has the old version installed.

Both publishing tabs default to Tenant scope, which publishes packages only to the selected tenant. Select Global to make packages available to all tenants on the server. Batch publishing requires BcContainerHelper; if it is missing, open Container Management to install it before retrying.
When a version is already published, its PackageId must match the selected .app file. If the package identity differs or cannot be verified, rebuild the app with a higher version before deploying.

## Usage

A detailed [manual](Business%20Central%20Manager/docs/Business%20Central%20Manager%20-%20Manual.pdf) is included in the `Business Central Manager/docs` folder.

![Business Central Manager](https://img001.prntscr.com/file/img001/mOYuhPelScqMCSQd0sQkBQ.png)

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change. Please, document all changes accordingly.

## License

[GNU General Public License v3.0](LICENSE)
