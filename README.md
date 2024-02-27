# external-connector-cli-m365
PowerShell script to create external connector with the CLI for Microsoft 365.

## Prerequisities

The script requires the CLI for Microsoft 365 to be installed.

```
npm install -g @pnp/cli-microsoft365
```

Check details: https://github.com/pnp/cli-microsoft365

Another prerequisite is PowerShell 7.x

```
winget install --id Microsoft.Powershell --source winget
```

## Run the script

```
pwsh <path_to_folder_>\ExternalConnector.ps1
```