#Requires -Version 5.1
#Requires -Modules Hyper-V
param (
    [string]$ReportFilePath = 'c:\temp\ReplicaReport.html',
    [int]$MaxReportAgeInMinutes = 60,
    [switch]$SkipSettingsCheck
)

###########################
#region Function Definitions
#############################

function Test-VMReplicaSettingsMatch {
    param (
        [Parameter(Mandatory = $true)]
        [string]$VMName,
        [Parameter(Mandatory = $true)]
        [string]$PrimaryHost,
        [Parameter(Mandatory = $true)]
        [string]$ReplicaHost
    )

    # Get settings of primary VM (on PrimaryHost)
    $VM1 = Get-VM -ComputerName $PrimaryHost -Name $VMName

    # Get detailed information about the VM's memory, cpu.
    $VM1Memory = Get-VMMemory -VM $VM1
    $VM1CPU = Get-VMProcessor -VM $VM1

    # Use Get-VHD to gather information about all VHDs or VHDXs attached to the VM
    $VM1HardDrives = Get-VHD -VMId $VM1.VMId -ComputerName $PrimaryHost

    # Get a count of SCSI controllers attached to the VM
    $VM1SCSIControllers = Get-VMScsiController -VM $VM1

    # Combine the returned objects into a single object
    $VM1Settings = New-Object -TypeName PSObject -Property @{
        VMName          = $VM1.Name
        MemoryStartup   = $VM1Memory.Startup
        MemoryMinimum   = $VM1Memory.Minimum
        MemoryMaximum   = $VM1Memory.Maximum
        CPUCount        = $VM1CPU.Count
        HardDriveCount  = $VM1HardDrives.Count
        HardDriveSize   = $VM1HardDrives.Size
        SCSIControllers = $VM1SCSIControllers | Select-Object -Property Name, Drives
    }

    # Get settings of replica VM (on ReplicaHost)
    $VM2 = Get-VM -ComputerName $ReplicaHost -Name $VMName

    # Get detailed information about the VM's memory, cpu.
    $VM2Memory = Get-VMMemory -VM $VM2
    $VM2CPU = Get-VMProcessor -VM $VM2

    # Use Get-VHD to gather information about all VHDs or VHDXs attached to the VM
    $VM2HardDrives = Get-VHD -VMId $VM2.VMId -ComputerName $ReplicaHost

    # Get a count of SCSI controllers attached to the VM
    $VM2SCSIControllers = Get-VMScsiController -VM $VM2

    # Combine the returned objects into a single object
    $VM2Settings = New-Object -TypeName PSObject -Property @{
        VMName          = $VM2.Name
        MemoryStartup   = $VM2Memory.Startup
        MemoryMinimum   = $VM2Memory.Minimum
        MemoryMaximum   = $VM2Memory.Maximum
        CPUCount        = $VM2CPU.Count
        HardDriveCount  = $VM2HardDrives.Count
        HardDriveSize   = $VM2HardDrives.Size
        SCSIControllers = $VM2SCSIControllers | Select-Object -Property Name, Drives
    }

    # Convert the VM settings objects to JSON and compare them
    $VM1Json = $VM1Settings | ConvertTo-Json
    $VM2Json = $VM2Settings | ConvertTo-Json

    # Compare the settings of VM1 and VM2 returning True if the match and False if they don't.
    return !(Compare-Object $VM1Json $VM2Json)
}

##############################
#endregion Function Definitions
################################

#################
# Main Entry Point
###################

# Set default error action
$ErrorActionPreference = 'Stop'

# Load list of Hyper-V Host Servers from settings file
Write-Host ('Loading list of Hyper-V Host Servers from settings file...')
$settingsPath = Join-Path -Path $PSScriptRoot -ChildPath 'settings.json'

# If the settings file doesn't exist, abort script with error
if (!(Test-Path -Path $settingsPath)) {
    Write-Error ('Settings.json file not found at: {0}.  See ExampleSettings.json.' -f $settingsPath)
    exit
}

$jsonContent = Get-Content -Path $settingsPath -Raw
$settings = ConvertFrom-Json -InputObject $jsonContent
$hvHosts = $settings.hvHosts
Write-Host 'List of host servers loaded from settings file:'
$hvHosts | ForEach-Object { Write-Host ('- {0}' -f $_) }

# Retrieve virtual machine replication information from Hyper-V hosts
# Filter the output to only include virtual machine replications that have a RelationshipType other than 'Extended'
Write-Host ('Retrieving virtual machine replication information from {0} Hyper-V hosts...' -f $hvHosts.Count)
$repInfo = Get-VMReplication -ComputerName $hvHosts | Where-Object { $_.RelationshipType -ne 'Extended' }

Write-Host 'Generating report...'

# Sort the Replication info output by Name and Mode
# Convert the output to HTML and specify the properties to include in the table
$repInfoHTML = $repInfo | Sort-Object Name, Mode | ConvertTo-Html -Fragment -Property Name, Mode, PrimaryServer, ReplicaServer, State, Health, FrequencySec, RelationshipType -PreContent '<div id="ReplicationTable">' -PostContent '</div>'

if (!$SkipSettingsCheck) {
    Write-Host 'Performing VM settings check...'

    $namesOfVMsWithReplicas = ($repInfo | Where-Object { $_.ReplicationMode -like '*Replica' } | Select-Object -ExpandProperty Name -Unique | Sort-Object)
    $primaryVMs = $repInfo | Where-Object { $_.ReplicationMode -eq 'Primary' }
    $replicaVMs = $repInfo | Where-Object { $_.ReplicationMode -eq 'Replica' }
    $extendedReplicaVMs = $repInfo | Where-Object { $_.ReplicationMode -eq 'ExtendedReplica' }

    # Create an empty array to store the results of the comparison
    $replicaSettingsMatch = @()

    # Loop through each VM that has a replica
    foreach ($vmName in $namesOfVMsWithReplicas) {
        Write-Host ('Checking replica settings for VM: {0}' -f $vmName)

        # Get the primary and replica VM objects
        $primaryVM = $primaryVMs | Where-Object { $_.Name -eq $vmName }
        $replicaVM = $replicaVMs | Where-Object { $_.Name -eq $vmName }

        # More than one ReplicaVM returned?
        if ($replicaVM.Count -gt 1) {
            Write-Host ('* More than one replica VM returned for VM: {0}. Is there an Extended Replica initialization in progress?' -f $vmName)
            # Select only the first ReplicaVM returned
            $replicaVM = $replicaVM | Select-Object -First 1
            Write-Host ('* Comparing only the first replica VM returned (Host: {0}) for VM: {1}' -f $replicaVM.ReplicaServer, $vmName)
        }

        # Test the settings of the primary and replica VMs to see if they match
        $settingsMatch = Test-VMReplicaSettingsMatch -VMName $vmName -PrimaryHost $primaryVM.PrimaryServer -ReplicaHost $replicaVM.ReplicaServer

        # Create a custom object to store the results of the comparison
        $replicaSettingsMatch += New-Object -TypeName PSObject -Property @{
            VMName        = $vmName
            SettingsMatch = $settingsMatch
        }
    }

    # Create an empty array to store the results of the comparison
    $extendedReplicaSettingsMatch = @()

    # Loop through each VM that has an extended replica
    foreach ($vmName in $namesOfVMsWithReplicas) {
        # Get the primary and replica VM objects
        $primaryVM = $primaryVMs | Where-Object { $_.Name -eq $vmName }
        $extendedReplicaVM = $extendedReplicaVMs | Where-Object { $_.Name -eq $vmName }

        # If the extended replica is not found, skip to the next VM
        if ($null -eq $extendedReplicaVM) {
            continue
        }

        Write-Host ('Checking extended replica settings for VM: {0}' -f $vmName)

        # Test the settings of the primary and replica VMs to see if they match
        $settingsMatch = Test-VMReplicaSettingsMatch -VMName $vmName -PrimaryHost $primaryVM.PrimaryServer -ReplicaHost $extendedReplicaVM.ReplicaServer

        # Create a custom object to store the results of the comparison
        $extendedReplicaSettingsMatch += New-Object -TypeName PSObject -Property @{
            VMName        = $vmName
            SettingsMatch = $settingsMatch
        }
    }

    # Use namesOfVMsWithReplicas, replicaSettingsMatch and extendedReplicaSettingsMatch to build an array of objects containing entries for each named VM with a replica.
    # The object will contain the VM name, the replica mode, and the result of the settings comparisons.
    $replicaReport = @()
    foreach ($vmName in $namesOfVMsWithReplicas) {
        $replicaReport += New-Object -TypeName PSObject -Property @{
            Name                         = $vmName
            ReplicaSettingsMatch         = $replicaSettingsMatch | Where-Object { $_.VMName -eq $vmName } | Select-Object -ExpandProperty SettingsMatch
            ExtendedReplicaSettingsMatch = $extendedReplicaSettingsMatch | Where-Object { $_.VMName -eq $vmName } | Select-Object -ExpandProperty SettingsMatch
        }
    }

    # Convert the replicaReport array to HTML and specify the properties to include in the table
    $replicaReportHTML = $replicaReport | Sort-Object Name | ConvertTo-Html -Fragment -Property Name, ReplicaSettingsMatch, ExtendedReplicaSettingsMatch -PreContent '<div id="SettingsMatchTable">' -PostContent '</div>'

    # Combine $replicaReportHTML and $repInfoHTML into a single HTML
    $repInfoHTML = $repInfoHTML + '<br />' + $replicaReportHTML
}

# Append time and date stamp to report.
Write-Host 'Appending time and date stamp to report...'
$repInfoHTML = $repInfoHTML + ('<div id="dateStamp">Report created on {0}, at {1}</div>' -f (Get-Date).ToString('MMM dd, yyyy'), (Get-Date).ToString('h:mm:ss tt'))

# Define the HTML header.
$htmlHeader = @"
<style>
body
{
    font-family: 'Hack', monospace;
}

table
{
    border-width: 1px;
    border-style: solid;
    border-color: black;
    border-collapse: collapse;
}

th
{
    border-width: 1px;
    padding: 3px;
    border-style: solid;
    border-color: black;
    background-color: #6495ED;
}

td
{
    border-width: 1px;
    padding: 3px;
    border-style: solid;
    border-color: black;
}

tbody tr:nth-child(odd)
{
    background-color: lightgray;
    color: black;
}

div#dateStamp
{
    font-size: 12px;
    font-style: italic;
    font-weight: normal;
    margin-top: 6px;
}
</style>
"@

# Define the HTML footer that contains JS scripts.
$htmlFooter = @"
<script type="text/javascript">
    document.addEventListener('DOMContentLoaded', function() {
        // Get all <td> elements in the document
        const tds = document.getElementsByTagName('td');
        const tdsLength = tds.length;

        // Loop through each <td> element
        for (let i = 0; i < tdsLength; i++) {
            const textContent = tds[i].textContent;
            if (textContent === 'False') {
                tds[i].style.color = 'red';
            } else if (textContent === 'True') {
                tds[i].style.color = 'green';
            }
        }

        // Select the table within the element with id 'ReplicationTable'
        const table = document.querySelector('#ReplicationTable table');
        if (table) {
            // Get all <th> elements (table headers) in the table
            const headers = table.getElementsByTagName('th');
            let healthColumnIndex = -1;

            // Loop through the headers to find the 'Health' column
            Array.from(headers).forEach((header, index) => {
                if (header.textContent.trim() === 'Health') {
                    healthColumnIndex = index;
                }
            });

            // If the 'Health' column is found
            if (healthColumnIndex !== -1) {
                // Get all rows in the table
                const rows = table.getElementsByTagName('tr');
                const rowsLength = rows.length;

                // Loop through each row, starting from 1 to skip the header row
                for (let i = 1; i < rowsLength; i++) {
                    const cells = rows[i].getElementsByTagName('td');
                    const healthCell = cells[healthColumnIndex];
                    if (healthCell) {
                        const healthText = healthCell.textContent.trim();
                        healthCell.style.color = healthText === 'Normal' ? 'green' : 'red';
                    }
                }
            }
        }

        // Get all table rows
        const tableRows = document.querySelectorAll('table tr');

        // Add mouseover and mouseout event listeners to each row
        tableRows.forEach(row => {
            row.addEventListener('mouseover', () => {
                const firstCell = row.cells[0];
                const valueToMatch = firstCell.textContent;
                tableRows.forEach(otherRow => {
                    if (otherRow === row) {
                        // Highlight the row being hovered over
                        row.style.backgroundColor = '#FBF719';
                    } else if (otherRow.cells[0].textContent === valueToMatch) {
                        // Highlight rows with matching first cell content
                        otherRow.style.backgroundColor = '#E1DE16';
                    }
                });
            });

            row.addEventListener('mouseout', () => {
                // Remove background color when mouse leaves the row
                tableRows.forEach(otherRow => {
                    otherRow.style.backgroundColor = '';
                });
            });
        });

        // Get the text content of the dateStamp div
        const dateStampText = document.getElementById('dateStamp').textContent;

        // Extract the date and time part from the text
        const dateTimeString = dateStampText.match(/on (.+), at (.+)/);
        const dateString = dateTimeString[1];
        const timeString = dateTimeString[2];

        // Combine date and time into a single string
        const fullDateTimeString = ```${dateString} `${timeString}``;

        // Parse the extracted date and time into a Date object
        const reportDate = new Date(fullDateTimeString);

        // Get the current date and time
        const currentDate = new Date();

        // Calculate the difference in milliseconds
        const timeDifference = currentDate - reportDate;

        // Convert the difference to minutes
        const timeDifferenceInMinutes = timeDifference / (1000 * 60);

        // Check if the report date is more than the specified number of minutes ago
        if (timeDifferenceInMinutes > $MaxReportAgeInMinutes) {
            // Move the div element to the top of the body
            const dateStampDiv = document.getElementById('dateStamp');
            document.body.insertBefore(dateStampDiv, document.body.firstChild);

            // Style the div element
            dateStampDiv.style.fontSize = '2em';
            dateStampDiv.style.fontWeight = 'bold';
            dateStampDiv.style.color = 'red';

            // Create a new div element for the additional message
            const warningDiv = document.createElement('div');
            warningDiv.style.color = 'black';
            warningDiv.style.fontSize = '1.5em';
            warningDiv.style.fontStyle = 'italic';
            warningDiv.textContent = 'Report may be out of date, please confirm!';

            // Insert the new div element after the dateStamp div
            dateStampDiv.insertAdjacentElement('afterend', warningDiv);

            console.log('The report date is more than $MaxReportAgeInMinutes minutes ago.');
        } else {
            console.log('The report date is within the last $MaxReportAgeInMinutes minutes.');
        }
    });
</script>
"@

# Combine the HTML header, footer and the report HTML into a single HTML document.
$htmlTemplate = @"
<html>
<head>
$htmlHeader
</head>
<body>
$repInfoHTML
</body>
$htmlFooter
</html>
"@

# Write the HTML to a file.
Write-Host ('Writing the HTML report file to: {0}' -f $ReportFilePath)
$htmlTemplate | Out-File $ReportFilePath

Write-Host 'Report generation completed.'