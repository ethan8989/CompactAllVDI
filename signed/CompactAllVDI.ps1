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

# SIG # Begin signature block
# MIIctwYJKoZIhvcNAQcCoIIcqDCCHKQCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAZEFtb5BJbILuf
# FcdYJ9lrAVsQUm1jrPoaigXumYeVbKCCF70wggUnMIIED6ADAgECAhACpD0TXDDQ
# araH1r3JF61EMA0GCSqGSIb3DQEBCwUAMHYxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xNTAzBgNV
# BAMTLERpZ2lDZXJ0IFNIQTIgSGlnaCBBc3N1cmFuY2UgQ29kZSBTaWduaW5nIENB
# MB4XDTE5MDUxNzAwMDAwMFoXDTIyMDUyNTEyMDAwMFowZjELMAkGA1UEBhMCVVMx
# DzANBgNVBAgTBk9yZWdvbjESMBAGA1UEBxMJSGlsbHNib3JvMRgwFgYDVQQKEw9E
# cnkgQ3JlZWsgUGhvdG8xGDAWBgNVBAMTD0RyeSBDcmVlayBQaG90bzCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBAK2IUc6vwgmcrGbsnw2yeCSZtzHeSpfh
# ySqHeCD6k2MFfx4QchdVyW9POu6mbOwe4hZNEmDeKO2xAI9tua3lsawhyyyvUx6R
# zgPzccYxbKtgC8ERpP4I1UX+FDMHcDWX3+2BE642Q0FZ7/aT2xH8ZIxstAQwa3+W
# uZGYs4gTq5u/flbEwiTBgQtjkDhbKvsdZQ0UM7+zD3oTacH9p/HQbZKXjHDLqNM8
# 5ESmOFcfPlDnVhjdcpmn0JuLDxLaYtzH4NqzOlyvRIvXMQZnNr+J70fVyGteEiyL
# jSTQhzwx6MYTH8T0loKbb3sjO0wm7o9MGugtbuiEeklvbMk5XOWf1mkCAwEAAaOC
# Ab8wggG7MB8GA1UdIwQYMBaAFGedDyAJDMyKOuWCRnJi/PHMkOVAMB0GA1UdDgQW
# BBRXJsSmCdXMek0dGCK5pkDwUpKSBTAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwbQYDVR0fBGYwZDAwoC6gLIYqaHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL3NoYTItaGEtY3MtZzEuY3JsMDCgLqAshipodHRwOi8vY3JsNC5kaWdp
# Y2VydC5jb20vc2hhMi1oYS1jcy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1s
# AwEwKjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAI
# BgZngQwBBAEwgYgGCCsGAQUFBwEBBHwwejAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMFIGCCsGAQUFBzAChkZodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRTSEEySGlnaEFzc3VyYW5jZUNvZGVTaWduaW5nQ0Eu
# Y3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggEBAHJJERI4doRloOo8
# 85l6uHyf6eGfr/XMLuuC01PswkqzcCdeggaVeiu6fH2WUgS1/Q3OxHfZZIL4uaze
# iq49ssHX/70ICjkbE9KXpcUmlU7eXjFVmVdRkWjRvvg7bF0L+Wepw9+lYnWGpfMS
# 1biaya741mvX8puaau3SyaMSSmVcmgC0ohU/uGUUvc200InLC6lPdd8gjQ6wHF//
# Ln9ucn5mCUur7NxL5Df+dGOiJQrouYi2D4r27TXt1fqwCeExsoQ6edX4FHVSSCKw
# Z7RwsObX2rhdzBtZkuzG1667YYWCByYaEC1XDeQtYm7fvOYtGZPkiBHNFcYA3uyn
# w6TsOFQwggVPMIIEN6ADAgECAhALfhCQPDhJD/ovZ5qHoae5MA0GCSqGSIb3DQEB
# CwUAMGwxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNV
# BAsTEHd3dy5kaWdpY2VydC5jb20xKzApBgNVBAMTIkRpZ2lDZXJ0IEhpZ2ggQXNz
# dXJhbmNlIEVWIFJvb3QgQ0EwHhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAw
# WjB2MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQL
# ExB3d3cuZGlnaWNlcnQuY29tMTUwMwYDVQQDEyxEaWdpQ2VydCBTSEEyIEhpZ2gg
# QXNzdXJhbmNlIENvZGUgU2lnbmluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEP
# ADCCAQoCggEBALRKXn0HD0HexPV2Fja9cf/PP09zS5zRDf5Ky1dYXoUW3QIVVJnw
# jzwvTQJ4EGjI2DVLP8H3Z86YHK4zuS0dpApUk8SFot81sfXxPKezNPtdSMlGyWJE
# vEiZ6yhJU8M9j8AO3jWY6WJR3z1rQGHuBEHaz6dcVpbR+Uy3RISHmGnlgrkT5lW/
# yJJwkgoxb3+LMqvPa1qfYsQ+7r7tWaRTfwvxUoiKewpnJMuQzezSTTRMsOG1n5zG
# 9m8szebKU3QBn2c13jhJLc7tOUSCGXlOGrK1+7t48Elmp8/6XJZ1kosactn/UJJT
# zD7CQzIJGoYTaTz7gTIzMmR1cygmHQgwOwcCAwEAAaOCAeEwggHdMBIGA1UdEwEB
# /wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MH8GCCsGAQUFBwEBBHMwcTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNl
# cnQuY29tMEkGCCsGAQUFBzAChj1odHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20v
# RGlnaUNlcnRIaWdoQXNzdXJhbmNlRVZSb290Q0EuY3J0MIGPBgNVHR8EgYcwgYQw
# QKA+oDyGOmh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEhpZ2hBc3N1
# cmFuY2VFVlJvb3RDQS5jcmwwQKA+oDyGOmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEhpZ2hBc3N1cmFuY2VFVlJvb3RDQS5jcmwwTwYDVR0gBEgwRjA4
# BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0
# LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYEFGedDyAJDMyKOuWCRnJi/PHM
# kOVAMB8GA1UdIwQYMBaAFLE+w2kD+L9HAdSYJhoIAu9jZCvDMA0GCSqGSIb3DQEB
# CwUAA4IBAQBqDv9+E3wGpUvALoz5U2QJ4rpYkTBQ7Myf4dOoL0hGNhgp0HgoX5hW
# QA8eur2xO4dc3FvYIA3tGhZN1REkIUvxJ2mQE+sRoQHa/bVOeVl1vTgqasP2jkEr
# iqKL1yxRUdmcoMjjTrpsqEfSTtFoH4wCVzuzKWqOaiAqufIAYmS6yOkA+cyk1Lqa
# NdivLGVsFnxYId5KMND66yRdBsmdFretSkXTJeIM8ECqXE2sfs0Ggrl2RmkI2DK2
# gv7jqVg0QxuOZ2eXP2gxFjY4lT6H98fDr516dxnZ3pO1/W4r/JT5PbdMEjUsML7o
# jZ4FcJpIE/SM1ucerDjnqPOtDLd67GftMIIGajCCBVKgAwIBAgIQAwGaAjr/WLFr
# 1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQD
# ExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAwWhcNMjQx
# MDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxJTAj
# BgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBTqZ8fZFnm
# fGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWRn8YUOawk
# 6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRVfRiGBYxV
# h3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3vJ+P3mvBM
# MWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA8bLOcEaD
# 6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGjggM1MIID
# MTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggr
# BgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIBkjAoBggr
# BgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQGCCsGAQUF
# BwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUA
# cgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMA
# YwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQA
# IABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAA
# UABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkA
# bQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4A
# YwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYA
# ZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQASKxOYspkH
# 7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9MH0GA1Ud
# HwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2VydC5jb20v
# RGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEQ0EtMS5j
# cnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI//+x1Gos
# Me06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7easGAm6mlX
# IV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8OxwYtNiS7Dg
# c6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQNJsQOfxu1
# 9aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNtomHpigtt
# 7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbNMIIFtaAD
# AgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJBgNVBAYT
# AlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2Vy
# dC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0w
# NjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAf
# BgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/JM/xNRZFc
# gZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPsi3o2CAOr
# DDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ8DIhFonG
# cIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNugnM/JksU
# kK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJrGGWxwXOt
# 1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3owggN2MA4G
# A1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUHAwIGCCsG
# AQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIBxTCCAbQG
# CmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRpZ2ljZXJ0
# LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIwggFWHoIB
# UgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQAaQBmAGkA
# YwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUAcAB0AGEA
# bgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMAUAAvAEMA
# UABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEAcgB0AHkA
# IABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkAdAAgAGwA
# aQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8AcgBwAG8A
# cgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIAZQBuAGMA
# ZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsGAQUFBwEB
# BG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsG
# AQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1
# cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRo
# dHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0Eu
# Y3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSMEGDAWgBRF
# 66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+ybcoJKc4
# HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6hnKtOHis
# dV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5PsQXSDj0a
# qRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke/MV5vEwS
# V/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qquAHzunEIO
# z5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQnHcUwZ1P
# L1qVCCkQJjGCBFAwggRMAgEBMIGKMHYxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xNTAzBgNVBAMT
# LERpZ2lDZXJ0IFNIQTIgSGlnaCBBc3N1cmFuY2UgQ29kZSBTaWduaW5nIENBAhAC
# pD0TXDDQaraH1r3JF61EMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwx
# CjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGC
# NwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIC1LuDq6ElaOOTuW
# XbpdDfy+4uVdg3sP402V7LHe/o63MA0GCSqGSIb3DQEBAQUABIIBABUffE0ON5Fi
# +FWAju6n81uR1hlGb0UHfR6v8Z8LRLagKIICkws7U//e5T8+IwsDuG8uEg6WRB0A
# QHvLUrnRS26Aw5KE+ZyQi/1FeSZ8RlEHWNqqMv3/Hymhtl+YkxjTguAlCY13z3Dv
# yicebm2EyTqEhPWjlxLvsuWqSAFvaT0Wj57g4nae3rS20wTlBTvTaIj/McsyILP9
# YQNNKwErfSpsJxoMn3pUqMV+kAX3HY0VhSKl0qn8T7xOWaf1LxEf1rcZKU5X3h42
# SD0PhJMF77n1uLcaqMl8ld0a6bN8mFcTPG0nSkuC8BniHKVg9UooXuQcLToqm+h3
# v8rNHo8WhT6hggIPMIICCwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTECEAMB
# mgI6/1ixa9bV6uYX8GYwCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG
# 9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTIwMDMxMTIwMTc0OFowIwYJKoZIhvcNAQkE
# MRYEFEqMXRluShXDEMmqfF0WKd7j+7bpMA0GCSqGSIb3DQEBAQUABIIBACga8s0I
# cjHizsvzjksH1HXtFaDnuE/g19F8C1JrOdzhwvk/CfkPoOOJCg5y/NADqcMtYBBL
# Jg9laQG6Ey+Eebp1wko2B+TYy/aOGakxYO5ySLNvHwrn6dfdB1KRBgPLLEsIE8OG
# LsGvDM3gmjdD8IeSCO6tsag2Q/lh3J07Hq7LpVfSdKcqYMhKlwP8Rswk/xobNyap
# gSHJUlmXKj5Eg0ZCiA9dCHZ5uTR7+IVeG6JD0Q0EZtBT4OQ4I65rhVVNHIkbKKfT
# FjXJLaOBFiMYJtKUpwNQ/c2FqkKc9UAxAidUBKI5Kc8z7P8p/t+qhFDt+l0YbuJE
# DGiIjuP16ss3JMw=
# SIG # End signature block
