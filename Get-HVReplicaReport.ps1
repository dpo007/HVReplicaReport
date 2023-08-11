#Requires -Version 5.1
#Requires -Modules Hyper-V
param (
    [string]$ReportFilePath = 'c:\temp\ReplicaReport.html',
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

    Write-Host ('Testing settings match for VM: {0}' -f $VMName)

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
        SCSIControllers = $VM1SCSIControllers
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
        SCSIControllers = $VM2SCSIControllers
    }

    # Compare the settings of VM1 and VM2 returning True if they match and False if they don't.
    return !(Compare-Object $VM1Settings.PSObject.Properties $VM2Settings.PSObject.Properties)
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
$repInfoHTML = $repInfo | Sort-Object Name, Mode | ConvertTo-Html -Fragment -Property Name, Mode, PrimaryServer, ReplicaServer, State, Health, FrequencySec, RelationshipType

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
    $replicaReportHTML = $replicaReport | Sort-Object Name | ConvertTo-Html -Fragment -Property Name, ReplicaSettingsMatch, ExtendedReplicaSettingsMatch

    # Combine $replicaReportHTML and $repInfoHTML into a single HTML
    $repInfoHTML = $repInfoHTML + '<br />' + $replicaReportHTML
}

# Append time and date stamp to report.
Write-Host 'Appending time and date stamp to report...'
$repInfoHTML = $repInfoHTML + ('<div id="dateStamp">Report created on {0}, at {1}</div>' -f (Get-Date).ToString('MMM dd, yyyy'), (Get-Date).ToString('h:mm:ss tt'))

# Define the HTML header.
$htmlHeader = @"
<style>
TABLE
{
    border-width: 1px;
    border-style: solid;
    border-color: black;
    border-collapse: collapse;
}

TH
{
    border-width: 1px;
    padding: 3px;
    border-style: solid;
    border-color: black;
    background-color: #6495ED;
}

TD
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
var tds = document.getElementsByTagName('td');
for (var i = 0; i < tds.length; i++) {
    if (tds[i].textContent == 'False') {
        tds[i].style.color = 'red';
    } else if (tds[i].textContent == 'True') {
        tds[i].style.color = 'green';
    }
}

const tableRows = document.querySelectorAll('table tr');

tableRows.forEach(row => {
    row.addEventListener('mouseover', () => {
        const firstCell = row.cells[0];
        const valueToMatch = firstCell.textContent;
        tableRows.forEach(otherRow => {
            if (otherRow === row) {
                row.style.backgroundColor = '#FBF719';
            } else if (otherRow.cells[0].textContent === valueToMatch) {
                otherRow.style.backgroundColor = '#E1DE16';
            }
        });
    });

    row.addEventListener('mouseout', () => {
        tableRows.forEach(otherRow => {
            otherRow.style.backgroundColor = '';
        });
    });
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