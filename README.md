# Business Central Manager

Business Central Manager is a PowerShell based GUI application used for Business Central managing with customizability and automatic updates. 

## Installation

Download the latest release from [releases](https://github.com/Uki99/Business-Central-Manager/releases/latest), unzip it and run `Launch Business Central Manager.cmd` from the extracted folder. Keep the folder structure intact so the launcher can find the scripts. The launcher hides its batch window after startup, though Windows may briefly flash a console when opening a CMD file. The `hidePowerShellConsole` setting still controls whether the application's own PowerShell window is shown.

## Configuration

In Settings, set the tenant ID for app publishing (the default is `default`). If more than one Business Central installation is found, specify the desired `NavAdminTool.ps1` path. The server list defaults to BC 22 (`^22\.`); adjust the version filter for your target. Automatic unpublishing of previous app versions is disabled by default and will only proceed when no tenant still has the old version installed.

## Usage

A detailed manual is provided at the [following](https://github.com/Uki99/Business-Central-Manager/blob/main/Business%20Central%20Manager/data/Business%20Central%20Manager%20-%20Manual.pdf) link.

![Business Central Manager](https://img001.prntscr.com/file/img001/mOYuhPelScqMCSQd0sQkBQ.png)

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change. Please, document all changes accordingly.

## License

[GNU General Public License v3.0](LICENSE)
