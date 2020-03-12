<#
.SYNOPSIS
    Compact virtual disks (vdi format) for all VirtualBox VMs registered on the system

.DESCRIPTION
    Uses CloneVDI.exe to compact virtual disk vdi files. Original vdi files are deleted and
    replaced by compacted clone. Option to empty recycle bin after each compaction if
    vdi files are on a limited-space SSD.
    Common sense warning: Be sure to have backups first!

    Notes:
        [1]: Your virtual disks will be altered. Have backups, particularly if
             you enable the EmptyRecycleBin option!
        [2]: This script has only been tested on Windows 10 x64 and Server 2019.
             It will not run on 32-bit Windows.
        [3]: Any running virtual machines will be powered off. ACPI shutdown if that works,
             forceably if not. Any running VirtualBox processes will be closed as well.
        [4]: Tested with CloneVDI versions 3.02 and 4.01. Version 4.00 does not support
             compaction from the command line.
        [5]: Compacts both VirtualBox 5.2.x and 6.1.x vdi disks.
        [6]: Regex pattern matching looks for VDI disks only.
             Other formats are skipped.
        [7]: Untested on virtual machines with multiple vdi disks.
        [8]: Cloned and compacted vdi disk has the same UUID as the original.
        [9]: Powershell WhatIf syntax supported

.PARAMETER PathToCloneVDI
    Path to CloneVDI.exe. If not specified, defaults to the directory this script is in.
    If CloneVDI.exe cannot be located, a file search dialog opens.

.PARAMETER EmptyRecycleBin
    False (default) to leave original vdi files in the recycle bin after being replaced by
    the compacted clones.

    True to empty the recycle bin after each compaction. Can be useful if compacting large
    VMs on a SSD with limited free space. All standard warnings about backups, etc. apply.

.EXAMPLE
    .\CompactAllVDI.ps1

    Compacts all virtual disks registered with VirtualBox. If CloneVDI.exe is not in the
    same directory this script is in, a dialog opens to find it.

.EXAMPLE
    .\CompactAllVDI.ps1 "C:\Some Folder\CloneVDI.exe"

    Compacts all virtual disks registered with VirtualBox using CloneVDI.exe located in C:Some Folder

.EXAMPLE
    .\CompactAllVDI.ps1 -EmptyRecycleBin $True

    Compact all virtual machines. Empty the recycle bin for the drive a vdi disk is on after each one is compacted.
    Useful if you have a SSD that does not have enough free space to handle both compacted and original copies of each disk.

.EXAMPLE
    .\CompactAllVDI.ps1 -PathToCloneVDI "C:\Some Folder\CloneVDI.exe" -EmptyRecycleBin $True -WhatIf

    Show what VM disks will be compacted.

.LINK
    CloneVDI can be downloaded from:

    https://forums.virtualbox.org/viewtopic.php?t=22422

    ------------------

    Original source signed by Dry Creek Photo

    https://www.drycreekphoto.com/

.NOTES
    Author: Ethan Hansen/Dry Creek Photo/www.drycreekphoto.com
    Last Edit: 2020-02-12
    Version 1.0  - 2020-02-10 First public release
    Version 1.1  - 2020-02-12 Verifies that ClonVDI.exe is version 3.02 or >= 4.01
    Version 1.2  - 2020-03-10 Terminates any orphaned VBoxSVC processes
                              still running after VirtualBox.exe closed.

    Released under the MIT license
    Copyright (c) 2020 Dry Creek Photo
#>

<#
Copyright (c) 2020 Dry Creek Photo

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

[CmdletBinding(SupportsShouldProcess)]

param(
    [Parameter(Mandatory=$False)]
    [String]$PathToCloneVDI = "$($PSScriptRoot)\CloneVDI.exe",

    [Parameter(Mandatory=$False)]
    [bool]$EmptyRecycleBin = $False

)




#----------------[ Functions ]------------------


Function Compact-VM {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, Position=0, ValueFromPipeline=$false)]
        [String]$vdiDisk = $(throw "-vdiDisk is required."),
 
        [Parameter(Mandatory, Position=1, ValueFromPipeline=$false)]
        [String]$vdiClone = $(throw "-vdiClone is required."),
 
        [Parameter(Mandatory, Position=2, ValueFromPipeline=$false)]
        [String]$CloneVDIexePath = $(throw "-CloneVDIexePath is required.")

   )
<#
.SYNOPSIS
    Compact a single VirtualBox vdi disk

.DESCRIPTION
    Calls CloneVDI.exe to compact a virtual disk. Should typically only be called by Compact-All for each vdi file.

.PARAMETER vdiDisk
    Virtual disk file to compact

.PARAMETER vdiClone
    Name of cloned virtual disk

.PARAMETER CloneVDIexePath
    Fully qualified path to CloneVDI.exe

.OUTPUTS
    True if CloneVDI successfully compacted the disk. False if not.

#>

    # Make sure $CloneVDIexePath actually exists
    if (!(Test-Path $CloneVDIexePath -PathType Leaf)) {
        Write-Warning ($CloneVDIexePath) not found
        return $false
    }
    
    # Check if clone exists already
    if (Test-Path $vdiClone -PathType Leaf) {
        Write-Warning ($vdiClone) exists already.
        Write-Warning Skipping $(Split-Path -Path $vdiDisk -Leaf)
        return $false
    }

    $cloneFile = Split-Path -Path $vdiClone -Leaf
    Write-Host "Compacting " -NoNewline
    Write-Host "$vdiDisk" -BackgroundColor Black -ForegroundColor Yellow -NoNewline
    Write-Host " into " -NoNewline
    Write-Host $cloneFile -BackgroundColor Black -ForegroundColor Green
    
    # Set output file, compaction, and keep-uuid
    $ArgList = $('"' + ($vdiDisk) + '" -o "' + ($vdiClone) + '" -kc')

    # Start CloneVDI. Wait for completion.
    # Using workaround for edge cases where Start-Process -Wait does
    # not return exit error code.
    if ($PSCmdlet.ShouldProcess($vdiDisk, 'Compacting Virtual Disk')) {
        $proc = Start-Process $CloneVDIexePath -ArgumentList $ArgList -PassThru
        Wait-Process -InputObject $proc

        # Check exit code. Error if non-zero
        if ($proc.ExitCode -ne 0) {
            Write-Warning "Something went wrong in cloning or CloneVDI cancelled. Figure it out yourself."
            return $false
        }
    }

    return $true
}



Function Compact-All {

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline=$false)]
        [string]$CloneVDIexePath = $(throw "-CloneVDIexePath is required."),

        [Parameter(Mandatory=$False, ValueFromPipeline=$false)]
        [bool]$EmptyRecycleBin = $false

    )
<#
.SYNOPSIS
    Compacts all VirtualBox vdi disks.

.DESCRIPTION
    Calls Compact-VM to compact each registered vdi virtual disk.
    Process: [1]: Compact a vdi disk into a clone, keeping the UUID
             [2]: Delete the original vdi disk file
             [3]: Rename the clone to the original filename.
             [4]: If $EmptyRecycleBin is $true, empty the recycle bin
                  for the drive the vdi disk is on
             [5]: Repeat for all VMs registered with VirtualBox

.PARAMETER CloneVDIexePath
    Required. Fully qualified path to CloneVDI.exe

.PARAMETER EmptyRecycleBin
    Optional. $false (default) to leave original vdi files in the recycle bin
              $true to empty the appropriate drive's recycle bin after each iteration.
              Useful if VM disk files are on a SSD with limited space.

.EXAMPLE
   Compact-All -CloneVDIexePath "C:\CloneVDI\CloneVDI.exe"

   Compacts all registered vdi disks using CloneVDI.exe located in C:\CloneVDI
   Original vdi disks are left in the recycle bin

.EXAMPLE
   Compact-All -CloneVDIexePath "C:\CloneVDI\CloneVDI.exe" -EmptyRecycleBin $true

   Same as above, but also empties recycle bin after each compaction.

#>

    # Does CloneVDI.exe exist where it is supposed to?
    if (!(Test-Path $CloneVDIexePath -PathType Leaf)) {
        Write-Warning "CloneVDI.exe not found at $CloneVDIexePath"
        return $false
    }

    # Create empty ArrayLists (can't use plain array as these are fixed size)
    $vmNamesArray  = [System.Collections.ArrayList]@()
    $vdiNamesArray  = [System.Collections.ArrayList]@()

    # Find each VM that VirtualBox knows about
    $vmList = & $Env:ProgramFiles\oracle\virtualbox\vboxmanage.exe list vms
    foreach ($vm in $vmList) {
        # Extract the name of each VM from the list
        if ($vm -match '\".*\"') {
            $vmNamesArray.Add($Matches[0]) | Out-Null
        }
    }

    # Get the virtual disk file (vdi only) associated with each VM
    # TODO: Making the assumption that there is only one vdi file per VM.
    #       Not tested with multiple disks per VM.
    foreach ($vm in $vmNamesArray) {
        # Get full info for each VM
        $mInfo = & $Env:ProgramFiles\oracle\virtualbox\vboxmanage.exe  showvminfo  $vm --machinereadable
        # Extract virtual disk (vdi) file name and path
        $mn = $mInfo | where {$_ -match '\".*\.vdi\"'}
        # Append to our list of vdi files
        $vdiNamesArray.Add(-join('"', [regex]::Match($mn, '=\"(.+\.vdi)\"').groups[1].value, '"')) | Out-Null
    }

    # Get file info for each vdi disk to process
    $numDisks = $vdiNamesArray.Count
    $currentDisk = 0
    foreach ($vdi in $vdiNamesArray) {
        $currentDisk++
        $vdi = $vdi -replace ('"', '')
        $vdiFile = Split-Path -Path $vdi -Leaf            <# File name of vdi disk #>
        $vdiPath = Split-Path -Path $vdi -Resolve         <# Path of vdi disk #>
        $vdiDrive = Split-Path -Path $vdi -Qualifier      <# Drive vdi disk is on (used if recycling after compaction) #>
        $vdiClone = $vdiPath + "\Clone of " + $vdiFile    <# Name of cloned vdi disk #>
        if (Test-Path $vdiClone -PathType Leaf) {
            Write-Warning "Skipping compaction of $vdiFile"
            Write-Warning "$vdiClone already exists. Figure out what to do with it first"
        }
        else {
            Write-Host "Processing virtual disk $currentDisk of $numDisks."
            # Compact the disk. Allow WhatIf support
            if (Compact-VM -vdiDisk $vdi -vdiClone $vdiClone -CloneVDIexePath $CloneVDIexePath) {
                # Success. Delete original file, rename clone to main disk
                # Set flag for WhatIf mode to show what would happen
                ($JustTesting = !$PSCmdlet.ShouldProcess($vdi, 'WhatIf Mode')) | Out-Null
                 if ((Test-Path $vdiClone -PathType Leaf) -or $JustTesting) {
                    Write-Host "Moving clone to $vdiFile"
                    if ($PSCmdlet.ShouldProcess($vdi, 'Deleting original VDI disk')) {
                        DeleteToRecycleBin -FileToDelete $vdi
                    }
                    if ($PSCmdlet.ShouldProcess($vdiClone, "Renaming clone to $vdiFile")) {
                        Rename-Item -Path $vdiClone -NewName $vdiFile
                    }
                    if ($EmptyRecycleBin) {
                        if ($PSCmdlet.ShouldProcess($vdiDrive, "Emptying recycle bin for $vdiDrive")) {
                            Clear-RecycleBin -DriveLetter $vdiDrive -Force
                        }
                    }
                }
                else
                {
                    Write-Warning "Clone of $vdiFile does not exist. Somebody goofed."
                }
            }
            else {
                Write-Host "Compaction of $vdiFile failed."
            }
        }
        Write-Host
    }
}



Function Find-CloneVDI {
<#
.SYNOPSIS
    Find CloneVDI.exe executable file

.DESCRIPTION
   Open a standard file selection dialog to search for CloneVDI.exe

.OUTPUTS
    Fully qualified path to CloneVDI.exe if found, $null if not

.EXAMPLE
    $CloneVDIPath = Find-CloneVDI
    if ($CloneVDIPath -ne $null) {
        <do something>
    }
    else {
        <don't do something>
    }

#>
    Write-Host
    Write-Host "Please locate CloneVDI.exe to proceed"
    Write-Host

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        Title = 'Locate CloneVDI.exe'
        InitialDirectory = Get-Location            <# Default to current directory #> `               
        Filter = 'CloneVDI|CloneVDI.exe'           <# Limit selection to CloneVDI.exe #>
    }
    $result = $FileBrowser.ShowDialog()
    if ($result -eq "OK") {
        return $FileBrowser.FileName
    }
    else {
        return $null
    }
}



Function Check-CloneVDI-Version {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline=$false)]
        [string]$CloneVDIexePath = $(throw "-CloneVDIexePath is required.")

    )
<#
.SYNOPSIS
    Checks if CloneVDI version is supported.

.DESCRIPTION
    CloneVDI version 4.00 does not support the -c (compact) command line
    parameter. Verifies that CloneVDI is not 4.00. If version 3,
    only runs with version 3.02 (other versions not verified to work)

.PARAMETER CloneVDIexePath
    Required. Fully qualified path to CloneVDI.exe


.EXAMPLE
   Check-CloneVDI-Version -CloneVDIexePath "C:\CloneVDI\CloneVDI.exe"

   Checks file version of CloneVDI.exe

.OUTPUTS
    bool: $True if CloneVDI.exe exists and is a supported version
          $False if not

.LINK
    Latest version of CloneVDI available from:

    https://forums.virtualbox.org/viewtopic.php?f=6&t=22422&sid=e50a48c8a30e7923b3e296580d1fc067
#>

    # Does CloneVDI.exe exist?
    if (!(Test-Path $CloneVDIexePath -PathType Leaf)) {
        $CloneVDIexePath = Find-CloneVDI
        if ($CloneVDIexePath -eq $null) {
            Write-Warning "CloneVDI.exe not found. Can't do anything else."
            return $false
        }
    }

    # Check what version of CloneVDI we're using. V4.00 does not support compaction from the command line.
    # Versions < 3.02 not tested
    $exeFile = Get-Item -Path $CloneVDIexePath
    $exeVersionStr = $exeFile.VersionInfo.FileVersion
    [double]$exeVersion = [convert]::ToDouble($exeVersionStr)
    $WrongExe = $false
    if ($exeVersion -lt 3.02) {
        Write-Warning "CloneVDI.exe version is $exeVersion. Versions under 3.02 not supported."
        $WrongExe = $true
    }
    else {
        if (($exeVersion -gt 3.999) -and ($exeVersion -lt 4.001)) {
            Write-Warning "CloneVDI.exe version 4.00 does not support disk compaction. Use 4.01 or above."
            $WrongExe = $true
       }
    }

    # Display link to latest and greatest CloneVDI if necessary.
    if ($WrongExe) {
        Write-Host
        Write-Warning "Download the latest version of CloneVDI from:"
        Write-Warning "https://forums.virtualbox.org/viewtopic.php?f=6&t=22422&sid=e50a48c8a30e7923b3e296580d1fc067"
        Write-Host
        return $false
    }

    # We're golden
    return $true
}


Function Shutdown-All-VMS {

    [CmdletBinding(SupportsShouldProcess)]

<#
.SYNOPSIS
    Shut down all running Virtual machines and close VirtualBox

.DESCRIPTION
    Attempts ACPI shutdown on all running VMs.
    If shutdown fails, powers off each VM.
    VirtualBox is closed if it is running.

.EXAMPLE
   PS> Shutdown-All-VMS

   Shuts down any running VirtualBox VMs, then closes VirtualBox

.EXAMPLE
   PS> Shutdown-All-VMS -WhatIf

   Displays actions that would happen if you were not WhatIffing.

#>

    # Get list of running VMs, if any
    $vmList = & $Env:ProgramFiles\oracle\virtualbox\vboxmanage.exe list runningvms
    foreach ($vm in $vmList) {
        # Extract the name of each VM from the list
        if ($vm -match '\".*\"') {
            $vmName = $Matches[0]
            $vmNameOnly = $vmName.Replace("`"","")
           if ($PSCmdlet.ShouldProcess($vmNameOnly, "Shutting down virtual machine")) {
                # Shut the sucker down
                & $Env:ProgramFiles\oracle\virtualbox\vboxmanage.exe controlvm ($vmName) acpipowerbutton
                # Wait a maximum of $maxWait seconds for the VM to shut down
                $maxWait = 60
                $poweredOff = $false
                for ($waitTime = 0; $waitTime -lt $maxWait; $waitTime++) {
                    Write-Progress -Activity "Shutting down $vmNameOnly" -Status "ACPI shutdown" -PercentComplete (100 * ($waitTime + 1)/$maxWait) -SecondsRemaining ($maxWait - $waitTime)
                    Start-Sleep -Seconds 1
                    # Get current status of VM
                    $mInfo = & $Env:ProgramFiles\oracle\virtualbox\vboxmanage.exe  showvminfo $vmName --machinereadable
                    # Check power state
                    $vmState =  $mInfo | where {$_ -match 'VMState='}
                    # Powered off?
                    if ($vmState -match 'poweroff') {
                        # Get outta here.
                        Write-Progress -Activity "$vmNameOnly shut down." -Status "ACPI success." -Completed 
                        $poweredOff = $true
                        Write-Host "$vmNameOnly shut down cleanly."
                        break;
                    }
                }

                # Kill VM if not off yet
                if (!$poweredOff) {
                    Write-Progress -Activity "Powering off $vmNameOnly" -Status "Shutdown failed." -Completed 
                    Write-Host "$vmNameOnly did not respond to shut down. Powering off... " -NoNewline
                    & $Env:ProgramFiles\oracle\virtualbox\vboxmanage.exe controlvm $vmName poweroff
                    Start-Sleep -Seconds 2
                    Write-Host "Done."
                }
                Write-Host
            }
        }
    }

    # Stop Virtualbox if it is still running
    $VBoxProc = Get-Process VirtualBox -ErrorAction SilentlyContinue
     if ([bool] ($VBoxProc)) {
        if ($PSCmdlet.ShouldProcess("VirtualBox.exe", "Shutting down VirtualBox")) {
            # Try graceful shutdown
            Write-Host "Shutting down VirtualBox... " -NoNewline
            $VBoxProc.CloseMainWindow() | Out-Null
            Start-Sleep -Seconds 5
            if ($VBoxProc.HasExited) {
                Write-Host "Done."
            }
            else {
                Write-Host "Killing VirtualBox process"
                $VBoxProc | Stop-Process -Force
                Start-Sleep -Seconds 2
            }

            # Check if VBoxSVC.exe is running. May have one or more orphaned processes.
            # If orphaned (usually the case if still running after VirtualBox terminated),
            # will not respond to CloseWindow. Skipping to Stop-Process instead
            $VBoxSVC = Get-Process -Name "VBoxSVC" -ErrorAction SilentlyContinue
            if ([bool] ($VBoxSVC -ne $null)) {
                Write-Host "Killing orphaned VBoxSVC process(s)"
                $VBoxSVC | Stop-Process -Force
                Start-Sleep -Seconds 2
            }
        }
        Write-Host
    }
}


Function DeleteToRecycleBin {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory,  Position=0, ValueFromPipeline=$false)]
        [string]$FileToDelete = $(throw "-FileToDelete is required.")

    )
<#
.SYNOPSIS
    Deletes a file to the recycle bin.

.DESCRIPTION
    Supports using the recycle bin for deleted files.
    Standard Powershell Remove-Item does not.

.PARAMETER FileToDelete
    Required. File to delete.

.EXAMPLE
   DeleteToRecycleBin "C:\Temp\Junk.txt.

   Recycles Junk.txt

.LINK
   Shamelessly adapted from

   https://stackoverflow.com/questions/502002/how-do-i-move-a-file-to-the-recycle-bin-using-powershell
#>

    # Check if the file actually exists.
    if (Test-Path $FileToDelete -PathType Leaf) {
        $fileName = Split-Path -Path $FileToDelete -Leaf    <# File name only #>
        if ($PSCmdlet.ShouldProcess($fileName, "Recycling file")) {
            $shell = new-object -comobject "Shell.Application"
            $fileItem = $shell.Namespace(0).ParseName("$FileToDelete")
            $fileItem.InvokeVerb("Delete")
        }
    }
}

#----------------[ Main Execution ]---------------

Write-Host

# Check if vboxmanage.exe is found
$vbm = $Env:ProgramFiles + "\oracle\virtualbox\vboxmanage.exe"
if (!(Test-Path $vbm -PathType Leaf)) {
    Write-Warning "vboxmanage not found in $Env:ProgramFiles\oracle\virtualbox"
    Write-Host "Cannot continue"
    exit 1
}

# Does CloneVDI.exe exist?
if (!(Test-Path $PathToCloneVDI -PathType Leaf)) {
    $PathToCloneVDI = Find-CloneVDI
    if ($PathToCloneVDI -eq $null) {
        Write-Warning "CloneVDI.exe not found. Can't do anything else."
        exit 1
    }
}

# Check if CloneVDI.exe is a supported version
if (!(Check-CloneVDI-Version -CloneVDIexePath $PathToCloneVDI)) {
    exit 1
}

# Shut down any running VMs and close Virtualbox. Major confusion occurs otherwise
Shutdown-All-VMS

# Compact all VMs
Compact-All -CloneVDIexePath $PathToCloneVDI -EmptyRecycleBin $EmptyRecycleBin

# We're done
Write-Host
Write-Host "VirtualBox disk compaction complete."
Start-Sleep -Seconds 5

exit 0

