# Get-ReplicaReport.ps1
This PowerShell script generates a report of all replicas in a given environment. The report is saved as an HTML file at the specified file path.

## Usage
To use this script, simply run it in a PowerShell console with the desired parameters. The available parameters are:

`ReportFilePath`: The file path where the report will be saved. Defaults to c:\temp\ReplicaReport.html.
`SkipSettingsCheck`: A switch that skips checking if VM settings match in replica(s).

**Example usage:**
`.\Get-ReplicaReport.ps1 -ReportFilePath 'C:\Users\JohnDoe\Documents\ReplicaReport.html'`

## Contributing
If you find a bug or have a suggestion for improvement, please feel free to open an issue or submit a pull request.

## License
This script is licensed under the [MIT License](https://mit-license.org/).