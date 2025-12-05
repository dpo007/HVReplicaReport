#Requires -Version 7.0
#Requires -Modules Hyper-V

<#
.SYNOPSIS
    Generates an HTML report of Hyper-V VM replication status and settings across multiple hosts.

.DESCRIPTION
    This script collects replication information from multiple Hyper-V hosts, compares VM settings
    between primary and replica VMs, and generates a comprehensive HTML report with visual indicators
    for replication health and settings mismatches.

.PARAMETER ReportFilePath
    Path where the HTML report will be saved. Default: c:\temp\ReplicaReport.html

.PARAMETER MaxReportAgeInMinutes
    Age threshold in minutes for report staleness warning. Default: 60

.PARAMETER SkipSettingsCheck
    Skip the VM settings comparison between primary and replica VMs.

.PARAMETER ThrottleLimit
    Controls how many Hyper-V hosts are queried in parallel. Must be between 1 and [int]::MaxValue. Default: 4.
#>

param (
    [string]$ReportFilePath = 'c:\temp\ReplicaReport.html',
    [int]$MaxReportAgeInMinutes = 60,
    [switch]$SkipSettingsCheck,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$ThrottleLimit = 4
)

#region Function Definitions

function Test-VMSettingsEqual {
    <#
    .SYNOPSIS
        Compares two pre-collected VM settings objects for equality.

    .DESCRIPTION
        Takes two settings objects (already gathered) and compares them by converting
        to JSON and checking for differences. More efficient than Test-VMReplicaSettingsMatch
        when settings are already cached.

    .PARAMETER PrimarySettings
        Settings object for the primary VM

    .PARAMETER ReplicaSettings
        Settings object for the replica VM

    .OUTPUTS
        Boolean - True if settings match, False if they differ
    #>
    param (
        [Parameter(Mandatory = $true)]
        [psobject]$PrimarySettings,

        [Parameter(Mandatory = $true)]
        [psobject]$ReplicaSettings
    )

    # Remove host-specific fields so only VM configuration is compared
    $primaryComparable = $PrimarySettings | Select-Object -Property * -ExcludeProperty HostName
    $replicaComparable = $ReplicaSettings | Select-Object -Property * -ExcludeProperty HostName

    # Convert both settings objects to JSON for comparison
    # Use a consistent depth so nested properties (e.g. SCSI controllers) are fully captured
    $primaryJson = $primaryComparable | ConvertTo-Json -Depth 10
    $replicaJson = $replicaComparable | ConvertTo-Json -Depth 10

    # Return true only when the serialized representations are identical
    return ($primaryJson -eq $replicaJson)
}

function Get-HostReplicationAndSettings {
    <#
    .SYNOPSIS
        Retrieves replication status and VM settings from a single Hyper-V host.

    .DESCRIPTION
        Connects to a Hyper-V host, retrieves all VM replication information (excluding Extended replicas),
        and collects detailed settings for each replicated VM. Designed to run in parallel across multiple hosts.

    .PARAMETER HostName
        Name of the Hyper-V host to query

    .OUTPUTS
        PSCustomObject containing HostName, ReplicationInfo array, and VmSettings array
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$HostName
    )

    # Nested helper function to build a normalized VM settings object
    function New-VMSettingsObject {
        <#
        .SYNOPSIS
            Creates a standardized settings object for a VM.

        .DESCRIPTION
            Gathers memory, CPU, storage, and SCSI controller information
            for a VM and packages it into a consistent format for comparison.
        #>
        param (
            [Parameter(Mandatory = $true)]
            [Microsoft.HyperV.PowerShell.VirtualMachine]$VM,

            [Parameter(Mandatory = $true)]
            [string]$CurrentHost
        )

        # Collect VM resource information
        $vmMemory = Get-VMMemory -VM $VM
        $vmCPU = Get-VMProcessor -VM $VM
        $vmHardDrives = Get-VHD -VMId $VM.VMId -ComputerName $CurrentHost | Sort-Object -Property Path
        $vmSCSIControllers = Get-VMScsiController -VM $VM | Sort-Object -Property Name

        # Normalize drive/controller data to avoid host-specific differences
        $normalizedHardDriveSizes = $vmHardDrives | ForEach-Object { $_.Size }
        $normalizedScsiControllers = $vmSCSIControllers | ForEach-Object {
            [pscustomobject]@{
                Name       = $_.Name
                DriveCount = ($_.Drives).Count
            }
        }

        # Return standardized settings object
        return [pscustomobject]@{
            HostName        = $CurrentHost
            VMName          = $VM.Name
            MemoryStartup   = $vmMemory.Startup
            MemoryMinimum   = $vmMemory.Minimum
            MemoryMaximum   = $vmMemory.Maximum
            CPUCount        = $vmCPU.Count
            HardDriveCount  = $vmHardDrives.Count
            HardDriveSize   = $normalizedHardDriveSizes
            SCSIControllers = $normalizedScsiControllers
        }
    }

    # Initialize result containers
    $hostReplication = @()
    $hostSettings = @()

    try {
        # Retrieve all replication relationships (excluding Extended replicas)
        $hostReplication = Get-VMReplication -ComputerName $HostName -ErrorAction Stop |
        Where-Object { $_.RelationshipType -ne 'Extended' }

        # Get unique list of VMs involved in replication
        $vmNamesOnHost = $hostReplication | Select-Object -ExpandProperty Name -Unique

        # Collect detailed settings for each replicated VM
        foreach ($vmName in $vmNamesOnHost) {
            try {
                $vm = Get-VM -ComputerName $HostName -Name $vmName -ErrorAction Stop
                $hostSettings += New-VMSettingsObject -VM $vm -CurrentHost $HostName
            }
            catch {
                Write-Warning ("[{0}] Failed to gather settings for VM '{1}': {2}" -f $HostName, $vmName, $_.Exception.Message)
            }
        }
    }
    catch {
        Write-Warning ("[{0}] Failed to retrieve replication data: {1}" -f $HostName, $_.Exception.Message)
    }

    # Return combined host data
    return [pscustomobject]@{
        HostName        = $HostName
        ReplicationInfo = $hostReplication
        VmSettings      = $hostSettings
    }
}

#endregion Function Definitions

#region Main Script Execution

# Configure error handling behavior
$ErrorActionPreference = 'Stop'

#region Load Configuration
Write-Host 'Loading list of Hyper-V Host Servers from host list file...'

# Build path to host list file
$settingsPath = Join-Path -Path $PSScriptRoot -ChildPath 'hvhosts.json'

# Validate host list file exists
if (!(Test-Path -Path $settingsPath)) {
    Write-Error ('hvhosts.json file not found at: {0}.  See hvhostsExample.json.' -f $settingsPath)
    exit
}

# Parse host list file
$jsonContent = Get-Content -Path $settingsPath -Raw
$settings = ConvertFrom-Json -InputObject $jsonContent
$hvHosts = $settings.hvHosts

# Display loaded configuration
Write-Host 'List of host servers loaded from host list file:'
$hvHosts | ForEach-Object { Write-Host ('- {0}' -f $_) }
#endregion Load Configuration

#region Collect Replication Data
Write-Host ('Retrieving virtual machine replication information from {0} Hyper-V hosts...' -f $hvHosts.Count)

# Initialize job tracking
$jobs = @()
$hostResults = @()

# Capture helper function definition so jobs can import it within their runspaces
$hostReplicationFuncDefinition = ${function:Get-HostReplicationAndSettings}.Ast.Extent.Text

# Capture time for $GenerationTimeStart
$GenerationTimeStart = Get-Date

# Start parallel jobs with throttling to avoid overwhelming system
foreach ($hvHost in $hvHosts) {
    # Wait for a job slot to become available if at throttle limit
    while ($jobs.Count -ge $ThrottleLimit) {
        $finished = Wait-Job -Job $jobs -Any
        $hostResults += Receive-Job -Job $finished
        $jobs = $jobs | Where-Object { $_.Id -ne $finished.Id }
    }

    # Start background job to query this host
    $jobs += Start-ThreadJob -ArgumentList $hvHost, $hostReplicationFuncDefinition -ScriptBlock {
        param($hostName, $functionDefinition)

        # Rehydrate Get-HostReplicationAndSettings inside this runspace
        . ([scriptblock]::Create($functionDefinition))

        Import-Module Hyper-V -ErrorAction Stop
        Get-HostReplicationAndSettings -HostName $hostName
    }
}

# Wait for remaining jobs to complete
if ($jobs.Count -gt 0) {
    Wait-Job -Job $jobs | Out-Null
    $hostResults += Receive-Job -Job $jobs
    Remove-Job -Job $jobs -Force
}

# Extract replication info and VM settings from host results
$repInfo = $hostResults | ForEach-Object { $_.ReplicationInfo } | Where-Object { $_ }
$allVmSettings = $hostResults | ForEach-Object { $_.VmSettings } | Where-Object { $_ }

# Build fast lookup hashtable: "HostName|VMName" -> Settings object
$settingsByHostAndName = @{}
foreach ($settingsRow in $allVmSettings) {
    $key = '{0}|{1}' -f $settingsRow.HostName, $settingsRow.VMName
    $settingsByHostAndName[$key] = $settingsRow
}
#endregion Collect Replication Data

#region Generate Report
Write-Host 'Generating report...'

# Generate HTML table from replication data
$repInfoHTML = $repInfo |
Sort-Object Name, Mode |
ConvertTo-Html -Fragment -Property Name, Mode, PrimaryServer, ReplicaServer, State, Health, FrequencySec, RelationshipType -PreContent '<div id="ReplicationTable">' -PostContent '</div>'
#endregion Generate Report

#region VM Settings Comparison
if (!$SkipSettingsCheck) {
    Write-Host 'Performing VM settings check...'

    # Identify VMs with replicas and categorize by replication mode
    $namesOfVMsWithReplicas = ($repInfo |
        Where-Object { $_.ReplicationMode -like '*Replica' } |
        Select-Object -ExpandProperty Name -Unique |
        Sort-Object)

    $primaryVMs = $repInfo | Where-Object { $_.ReplicationMode -eq 'Primary' }
    $replicaVMs = $repInfo | Where-Object { $_.ReplicationMode -eq 'Replica' }
    $extendedReplicaVMs = $repInfo | Where-Object { $_.ReplicationMode -eq 'ExtendedReplica' }

    # Initialize results array for replica settings comparison
    $replicaSettingsMatch = @()

    # Compare primary and replica VM settings
    foreach ($vmName in $namesOfVMsWithReplicas) {
        Write-Host ('Checking replica settings for VM: {0}' -f $vmName)

        # Locate primary and replica VM objects
        $primaryVM = $primaryVMs | Where-Object { $_.Name -eq $vmName }
        $replicaVM = $replicaVMs | Where-Object { $_.Name -eq $vmName }

        # Handle multiple replica VMs (can occur during Extended Replica initialization)
        if ($replicaVM.Count -gt 1) {
            Write-Host ('* Multiple replica VMs found for: {0}. Extended Replica initialization may be in progress.' -f $vmName)
            $replicaVM = $replicaVM | Select-Object -First 1
            Write-Host ('* Using first replica (Host: {0}) for comparison' -f $replicaVM.ReplicaServer)
        }

        # Build lookup keys for settings cache
        $primaryKey = '{0}|{1}' -f $primaryVM.PrimaryServer, $vmName
        $replicaKey = '{0}|{1}' -f $replicaVM.ReplicaServer, $vmName

        # Retrieve cached settings
        $primarySettings = $settingsByHostAndName[$primaryKey]
        $replicaSettings = $settingsByHostAndName[$replicaKey]

        # Perform comparison if both settings are available
        if ($null -eq $primarySettings -or $null -eq $replicaSettings) {
            Write-Warning ('Could not find cached settings for primary/replica pair: {0}' -f $vmName)
            $settingsMatch = $null
        }
        else {
            $settingsMatch = Test-VMSettingsEqual -PrimarySettings $primarySettings -ReplicaSettings $replicaSettings
        }

        # Store comparison result
        $replicaSettingsMatch += New-Object -TypeName PSObject -Property @{
            VMName        = $vmName
            SettingsMatch = $settingsMatch
        }
    }

    # Initialize results array for extended replica settings comparison
    $extendedReplicaSettingsMatch = @()

    # Compare primary and extended replica VM settings
    foreach ($vmName in $namesOfVMsWithReplicas) {
        # Locate primary and extended replica VM objects
        $primaryVM = $primaryVMs | Where-Object { $_.Name -eq $vmName }
        $extendedReplicaVM = $extendedReplicaVMs | Where-Object { $_.Name -eq $vmName }

        # Skip if this VM doesn't have an extended replica
        if ($null -eq $extendedReplicaVM) {
            continue
        }

        Write-Host ('Checking extended replica settings for VM: {0}' -f $vmName)

        # Build lookup keys for settings cache
        $primaryKey = '{0}|{1}' -f $primaryVM.PrimaryServer, $vmName
        $extendedKey = '{0}|{1}' -f $extendedReplicaVM.ReplicaServer, $vmName

        # Retrieve cached settings
        $primarySettings = $settingsByHostAndName[$primaryKey]
        $extendedReplicaSettings = $settingsByHostAndName[$extendedKey]

        # Perform comparison if both settings are available
        if ($null -eq $primarySettings -or $null -eq $extendedReplicaSettings) {
            Write-Warning ('Could not find cached settings for primary/extended replica pair: {0}' -f $vmName)
            $settingsMatch = $null
        }
        else {
            $settingsMatch = Test-VMSettingsEqual -PrimarySettings $primarySettings -ReplicaSettings $extendedReplicaSettings
        }

        # Store comparison result
        $extendedReplicaSettingsMatch += New-Object -TypeName PSObject -Property @{
            VMName        = $vmName
            SettingsMatch = $settingsMatch
        }
    }

    # Build consolidated report showing both replica and extended replica comparison results
    $replicaReport = @()
    foreach ($vmName in $namesOfVMsWithReplicas) {
        $replicaReport += New-Object -TypeName PSObject -Property @{
            Name                         = $vmName
            ReplicaSettingsMatch         = $replicaSettingsMatch |
            Where-Object { $_.VMName -eq $vmName } |
            Select-Object -ExpandProperty SettingsMatch
            ExtendedReplicaSettingsMatch = $extendedReplicaSettingsMatch |
            Where-Object { $_.VMName -eq $vmName } |
            Select-Object -ExpandProperty SettingsMatch
        }
    }

    # Generate HTML table from settings comparison report
    $replicaReportHTML = $replicaReport |
    Sort-Object Name |
    ConvertTo-Html -Fragment -Property Name, ReplicaSettingsMatch, ExtendedReplicaSettingsMatch -PreContent '<div id="SettingsMatchTable">' -PostContent '</div>'

    # Append settings comparison table to main replication report
    $repInfoHTML = $repInfoHTML + '<br />' + $replicaReportHTML
}
#endregion VM Settings Comparison

# Capture time for $GenerationTimeEnd
$GenerationTimeEnd = Get-Date

#region Finalize HTML Report
Write-Host 'Appending time and date stamp to report...'

# Add timestamp to report for freshness tracking, including generation duration
$generationDuration = $GenerationTimeEnd - $GenerationTimeStart
$genMinutes = $generationDuration.Minutes
$genSeconds = $generationDuration.Seconds

$minutesLabel = if ($genMinutes -eq 1) { 'minute' } else { 'minutes' }
$secondsLabel = if ($genSeconds -eq 1) { 'second' } else { 'seconds' }

$repInfoHTML = $repInfoHTML + ('<div id="dateStamp">Report created on {0}, at {1}.  (Generated in {2} {3} and {4} {5}.)</div>' -f
    (Get-Date).ToString('MMM dd, yyyy'),
    (Get-Date).ToString('h:mm:ss tt'),
    $genMinutes,
    $minutesLabel,
    $genSeconds,
    $secondsLabel
)

# Define CSS styling for the HTML report
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

# Define JavaScript for interactive report features
$htmlFooter = @"
<script type="text/javascript">
    document.addEventListener('DOMContentLoaded', function() {
        // Color-code boolean values in table cells (True = green, False = red)
        const tds = document.getElementsByTagName('td');
        const tdsLength = tds.length;

        for (let i = 0; i < tdsLength; i++) {
            const textContent = tds[i].textContent;
            if (textContent === 'False') {
                tds[i].style.color = 'red';
            } else if (textContent === 'True') {
                tds[i].style.color = 'green';
            }
        }

        // Color-code Health column values (Normal = green, others = red)
        const table = document.querySelector('#ReplicationTable table');
        if (table) {
            const headers = table.getElementsByTagName('th');
            let healthColumnIndex = -1;

            // Find the Health column index
            Array.from(headers).forEach((header, index) => {
                if (header.textContent.trim() === 'Health') {
                    healthColumnIndex = index;
                }
            });

            // Apply color coding to Health column cells
            if (healthColumnIndex !== -1) {
                const rows = table.getElementsByTagName('tr');
                const rowsLength = rows.length;

                for (let i = 1; i < rowsLength; i++) {  // Start at 1 to skip header row
                    const cells = rows[i].getElementsByTagName('td');
                    const healthCell = cells[healthColumnIndex];
                    if (healthCell) {
                        const healthText = healthCell.textContent.trim();
                        healthCell.style.color = healthText === 'Normal' ? 'green' : 'red';
                    }
                }
            }
        }

        // Highlight matching VM names across tables on hover
        const tableRows = document.querySelectorAll('table tr');

        tableRows.forEach(row => {
            row.addEventListener('mouseover', () => {
                const firstCell = row.cells[0];
                const valueToMatch = firstCell.textContent;

                tableRows.forEach(otherRow => {
                    if (otherRow === row) {
                        // Highlight the currently hovered row (bright yellow)
                        row.style.backgroundColor = '#FBF719';
                    } else if (otherRow.cells[0].textContent === valueToMatch) {
                        // Highlight matching VM names in other tables (dimmer yellow)
                        otherRow.style.backgroundColor = '#E1DE16';
                    }
                });
            });

            row.addEventListener('mouseout', () => {
                // Clear all row highlighting when mouse leaves
                tableRows.forEach(otherRow => {
                    otherRow.style.backgroundColor = '';
                });
            });
        });

        // Check report age and display warning if stale
        const dateStampText = document.getElementById('dateStamp').textContent;
        const dateTimeString = dateStampText.match(/on (.+), at (.+)/);
        const dateString = dateTimeString[1];
        const timeString = dateTimeString[2];
        const fullDateTimeString = ```${dateString} `${timeString}``;

        // Parse report timestamp and calculate age
        const reportDate = new Date(fullDateTimeString);
        const currentDate = new Date();
        const timeDifference = currentDate - reportDate;
        const timeDifferenceInMinutes = timeDifference / (1000 * 60);

        // Display warning if report exceeds configured age threshold
        if (timeDifferenceInMinutes > $MaxReportAgeInMinutes) {
            const dateStampDiv = document.getElementById('dateStamp');

            // Move timestamp to top of page
            document.body.insertBefore(dateStampDiv, document.body.firstChild);

            // Style as prominent warning
            dateStampDiv.style.fontSize = '2em';
            dateStampDiv.style.fontWeight = 'bold';
            dateStampDiv.style.color = 'red';

            // Add warning message
            const warningDiv = document.createElement('div');
            warningDiv.style.color = 'black';
            warningDiv.style.fontSize = '1.5em';
            warningDiv.style.fontStyle = 'italic';
            warningDiv.textContent = 'Report may be out of date, please confirm!';
            dateStampDiv.insertAdjacentElement('afterend', warningDiv);

            console.log('The report date is more than $MaxReportAgeInMinutes minutes ago.');
        } else {
            console.log('The report date is within the last $MaxReportAgeInMinutes minutes.');
        }
    });
</script>
"@

# Assemble complete HTML document
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

# Write the HTML report to disk
Write-Host ('Writing the HTML report file to: {0}' -f $ReportFilePath)
$htmlTemplate | Out-File $ReportFilePath

Write-Host 'Report generation completed.'
#endregion Finalize HTML Report

#endregion Main Script Execution