function Test-VMReplicaSettingsMatch {
    param (
        [string]$VMName = "lfcuvmsrvdc1",
        [string]$PrimaryHost = "sw-sarvmhost1",
        [string]$ReplicaHost = "sw-sarvmhost2"
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
        VMName         = $VM2.Name
        MemoryStartup  = $VM2Memory.Startup
        MemoryMinimum  = $VM2Memory.Minimum
        MemoryMaximum  = $VM2Memory.Maximum
        CPUCount       = $VM2CPU.Count
        HardDriveCount = $VM2HardDrives.Count
        HardDriveSize  = $VM2HardDrives.Size
        SCSIControllers = $VM2SCSIControllers
    }

    # Compare the settings of VM1 and VM2 returning True if the match and False if they don't.
    return !(Compare-Object $VM1Settings.PSObject.Properties $VM2Settings.PSObject.Properties)
}

Test-VMReplicaSettingsMatch