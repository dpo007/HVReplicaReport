# Get-HVReplicaReport.ps1
This PowerShell script generates a report of all VM replicas in a given Hyper-V environment. The report is saved as an HTML file at the specified file path, and includes HTML highlighting on hover for easy tracking.

## Requirements
This script now targets PowerShell 7 or later. Run it from a `pwsh` console, not Windows PowerShell 5.1.

## Usage
To use this script, create a `hvhosts.json` file with a list of Hyper-V hosts, and then run the script in a PowerShell console with the desired parameters.

The available parameters are:

`ReportFilePath`: The file path where the report will be saved.  (Defaults to `c:\temp\ReplicaReport.html`)

`SkipSettingsCheck`: A switch that skips checking if VM settings match in replica(s).  (Less info, but faster report generation)

`MaxReportAgeInMinutes`: Specifies the maximum age, in minutes, that the report data can be before it is considered outdated. If the report data is older than this value, a warning will be displayed in the generated HTML report.  (Default is 60 minutes)

`ThrottleLimit`: Integer parameter that controls how many Hyper-V hosts are queried in parallel when collecting replication and settings data. Must be between 1 and [int]::MaxValue. Defaults to 4. Higher values can speed up checks but may increase load on Hyper-V hosts.

The `hvhosts.json` file supports the following keys:

`hvHosts`: An array of Hyper-V host names to query.

**Example usage:**

`.\Get-HVReplicaReport.ps1 -ReportFilePath 'C:\Users\JohnDoe\Documents\ReplicaReport.html'`

## Contributing
If you find a bug or have a suggestion for improvement, please feel free to open an issue or submit a pull request.

## License
This script is licensed under the [MIT License](https://mit-license.org/).