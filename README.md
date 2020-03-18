# Compact All VDI Disks

CompactAllVDI is a Windows tool to call [CloneVDI][clonevdi] on all [VirtualBox][vbox] virtual machine vdi disks registered on a system.


## Introduction

Virtual disk files in VirtualBox VMs grow with use. This is particularly true after operating system or package updates occur. Allocated file size easily grows to over 150% of actual usage within the virtual machine.

This Powershell script clones and optionally compacts existing VDI virtual disks. Original vdi files are deleted and replaced by compacted clone. CompactAllVDI automates the process, compacting the virtual disk file for each registered VM. Option to empty recycle bin after each compaction if vdi files are on a limited-space SSD.

Common sense warning: Be sure to have backups first!

### Technologies

CompactAllVDI is tested on Windows 10 and Windows Server 2019 hosts. Required additional software packages:

* [VirtualBox][vbox] - Tested on versions 6.1.2+ and 5.2.36+
* [CloneVDI][clonevdi] - Version 4.0.1+ or 3.0.2

CompactAllVDI is a GitHub [public repository][compactvdi].

### Installation

- Copy the CompactAllVDI.ps1 script from the signed folder.
-- This version is signed with a valid Authenticode certificate.
-- An unsigned version of the script is in the src folder. This is more useful for editing.


Check your [Powershell execution policy](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7)

```PowerShell
PS C:\> Get-ExecutionPolicy
```
The default execution policy for Windows clients is Restricted. This setting does not allow any scripts to run. Either set the policy to AllSigned to require all scripts be signed by a trusted publisher or to RemoteSigned to require only scripts downoloaded from the internet be signed.

Using an Execution-Policy of ByPass allows any script to run. This presents security risks and is not recommended.


```PowerShell
PS C:\> Set-ExecutionPolicy AllSigned -Scope CurrentUser

Execution Policy Change
The execution policy helps protect you from scripts that you do not trust. Changing the execution policy might
expose you to the security risks described in the about_Execution_Policies help topic at
https:/go.microsoft.com/fwlink/?LinkID=135170. Do you want to change the execution policy?
[Y] Yes  [A] Yes to All  [N] No  [L] No to All  [S] Suspend  [?] Help (default is "N"): y
PS C:\>
```

## Usage
CompactAllVDI.ps1 can be run like any other Powershell script. If no command line options are given, CompactAllVDI searches the script folder for CloneVDI.exe. If not found, a file search dialog opens.
Any running virtual machines will be powered off first (up to one minute wait for ACPI shutdows). Orphaned processes or hung VMs will be forceably terminated.

##### Parameters
All parameters are optional

**PathToCloneVDI**
Path to CloneVDI.exe. If not specified, defaults to the directory this script is in. If CloneVDI.exe cannot be located, a file search dialog opens.

**EmptyRecycleBin**
False (default) to leave original vdi files in the recycle bin after being replaced by the compacted clones.
True to empty the recycle bin after each compaction. Can be useful if compacting large VMs on a SSD with limited free space.

**WhatIf**
Powershell WhatIf syntax is supported. Use the -WhatIf flag to output what operations would be executed otherwise.

### Examples
```PowerShell
PS C:\> .\CompactAllVDI.ps1
```
Compacts all virtual disks registered with VirtualBox. If CloneVDI.exe is not in the same directory this script is in, a dialog opens to find it.

```PowerShell

PS C:\> .\CompactAllVDI.ps1 "C:\Some Folder\CloneVDI.exe"
```
Compacts all virtual disks registered with VirtualBox using CloneVDI.exe located in C:Some Folder

```PowerShell

PS C:\> .\CompactAllVDI.ps1 -EmptyRecycleBin $true
```
 Compact all virtual machines. Empty the recycle bin for the drive a vdi disk is on after each one is compacted. Useful if you have a SSD that does not have enough free space to handle both compacted and original copies of each disk. Also, starts SSD garbage collection - speeds up writes on some smaller drives.

```PowerShell

PS C:\> .\CompactAllVDI.ps1 -PathToCloneVDI "C:\Some Folder\CloneVDI.exe" -EmptyRecycleBin $True -WhatIf
```
Show what VM disks will be compacted.

### Notes
- Your virtual disks will be altered. Have backups, particularly if you enable the EmptyRecycleBin option!
- This script has only been tested on Windows 10 x64 and Server 2019. It will not run on 32-bit Windows.
- Any running virtual machines will be powered off. ACPI shutdown if that works, forceably if not. Any running VirtualBox processes will be closed as well.
- Tested with CloneVDI versions 3.02 and 4.01. Version 4.00 does not support compaction from the command line.
- Compacts both VirtualBox 5.2.x and 6.1.x vdi disks.
- Regex pattern matching looks for VDI disks only. Other formats are skipped.
- Untested on virtual machines with multiple vdi disks.
- Cloned and compacted vdi disk has the same VirtualBox UUID as the original.

### Todos

- [ ] Kill orphaned VBoxSDS processes if running as Administrator
- [x] Detail how many GB are freed up and percentage change [18-Mar-2020]
- [ ] Verify operation on virtual machines with multiple vdi disks

License
----

MIT


[![Dry Creek Photo](https://www.drycreekphoto.com/images/DryCreekLogo.gif)](https://www.drycreekphoto.com/)

[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)


   [compactvdi]: <https://github.com/ethan8989/CompactAllVDI>
   [clonevdi]: <https://forums.virtualbox.org/viewtopic.php?f=6&t=22422#p98235>
   [vbox]: <https://www.virtualbox.org/>
