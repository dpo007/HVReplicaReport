# Get-HVReplicaReport.ps1
This PowerShell script generates a report of all VM replicas in a given Hyper-V environment.  The report is saved as an HTML file at the specified file path, and includes fancy HTML highlighting on hover, for easy tracking.

## Usage
To use this script, create a `settings.json` file with a list of Hyper-V hosts, and then run the script in a PowerShell console with the desired parameters.

The available parameters are:

`ReportFilePath`: The file path where the report will be saved.  (Defaults to `c:\temp\ReplicaReport.html`)

`SkipSettingsCheck`: A switch that skips checking if VM settings match in replica(s).  (Less info, but faster report generation)

`MaxReportAgeInMinutes`: Specifies the maximum age, in minutes, that the report data can be before it is considered outdated. If the report data is older than this value, a warning will be displayed in the generated HTML report.  (Default is 60 minutes)

**Example usage:**

`.\Get-HVReplicaReport.ps1 -ReportFilePath 'C:\Users\JohnDoe\Documents\ReplicaReport.html'`

## Contributing
If you find a bug or have a suggestion for improvement, please feel free to open an issue or submit a pull request.

## License
This script is licensed under the [MIT License](https://mit-license.org/).