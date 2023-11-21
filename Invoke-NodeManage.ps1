<#
    .SYNOPSIS
        PS script to manage a nodes updates and 3rd party apps
        Deploy with GPO!
        Run as a scheduled task!
    .DESCRIPTION
        Uses Chocolatey to install apps and manage updates, can pull from either private or community feed
        If you wish to use a private feed, use Chocolatey.server on a Windows server somewhere on your network
        Uses PSGallery module PSWindowsUpdate for Windows Update
        Reboots at a scheduled time
        Directives are in directives.json, you should keep this in a fileshare that's accessible by the
        service account that runs the scheduled task. This script saves locally to sysdrive:/opscode by default
        You can also create some automation that pulls this file from github to a fileshare
#>

# Some variables
$ProgressPreference='SilentlyContinue'

# First things first, set up logging to event log:
New-EventLog -Source NodeManageScript -LogName "Node Manage" -ErrorAction SilentlyContinue
function Logit {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information', 'Warning', 'Error')]
        [string]$Severity = 'Information',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message
    )
    Write-EventLog -Source NodeManageScript -LogName "Node Manage" -EventID 3001 -EntryType $Severity -Message "Invoke-NodeManage: $Message"
}

# Install stuff
# Chocolatey
if ( -not (Test-Path "$ENV:SystemDrive\ProgramData\Chocolatey")) {
    Logit -Severity Information -Message "Chocolatey not found - Installing."
    Invoke-WebRequest https://chocolatey.org/install.ps1 -UseBasicParsing | Invoke-Expression
}

# Install WindowsUpdate module from PSGallery
# Server 2012 doesn't have the nice TLS, so we get it the old fashioned way
# Else, use the new way
$mods = Get-Module -ListAvailable -Name PSWindowsUpdate
if (-not ($mods)) {
    if ([System.Environment]::OSVersion.Version.Major -lt '10') {

        $url = 'https://gallery.technet.microsoft.com/scriptcenter/2d191bcd-3308-4edd-9de2-88dff796b0bc/file/41459/47/PSWindowsUpdate.zip'
        $zipfile = 'C:\Windows\Temp\' + $(Split-Path -Path $url -Leaf)
        $dest = 'C:\Windows\System32\WindowsPowershell\v1.0\Modules'
    
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $url -OutFile $zipfile
    
        $ExtractShell = New-Object -ComObject Shell.Application
        $Files = $ExtractShell.Namespace($zipfile).Items()
        $ExtractShell.Namespace($dest).CopyHere($Files)
    
        Import-Module -Name "$dest\PSWindowsUpdate\PSWindowsUpdate.psm1"
    
    } else {
    
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    
        Install-Module -Name PSWindowsUpdate -Force

        Import-Module -Name PSWindowsUpdate
    
    }
} else {
    Import-Module -Name PSWindowsUpdate
    Import-Module -Name "$dest\PSWindowsUpdate\PSWindowsUpdate.psm1"
}

# Watch for .NET Previews!
function Get-NotKBs {
    param (
        $Updates
    )

    ForEach ($update in $Updates) {
        $NotKBs = @()
        if ($update.Title -like '*Preview*') {
            $NotKBs += $update.KB
        }
    } 
    
    return $NotKBs
}

function UpdateOS {
    $UpdatesAvailable = Get-WUList -UpdateType Software
    $NotKBs = Get-NotKBs -Updates $UpdatesAvailable

    # Only if there are more than 0 updates after $NotKBs
    # And boo, @installsplat doesn't work with the cmdlet :/
    if ( ($UpdatesAvailable.Count - $NotKBs.Count) -gt 0 ) {
        Logit -Severity Information -Message "Found $($UpdatesAvailable.Count). Not installing $NotKBs"
        if ($NotKBs) {
            Get-WindowsUpdate -Install -WindowsUpdate -UpdateType Software -AcceptAll -IgnoreReboot -NotKBArticleID $NotKBs
        } else {
            Get-WindowsUpdate -Install -WindowsUpdate -UpdateType Software -AcceptAll -IgnoreReboot
        }
    } else {
        Logit -Severity Information -Message "No updates available, will try again later."
    }
}

# Default directory for directives.json
function New-OpscodeDir {
    New-Item -ItemType Directory -Force -Path "$env:SYSTEMDRIVE\opscode" | Out-Null
}

# if for some reason, network issue or something, get a set of default directives
function Get-DefaultDirectives {
    $defaultdirectives = @{
        enable= "True"
        update= @{
            status= "True"
            dayofweek= "Tuesday"
            hour = 23
        }
        managedapps= @{
            status= "False"
            apps= @{}
        }
        unmanagedapps= @{
            status= "False"
            apps= @{
                googlechrome= "Latest"
                notepadplusplus= "Latest"
            }
            update= @{
                status= "False"
                dayofweek= "Tuesday"
                hour= 23
            }
        }
        reboottime= @{
            status= "False"
            dayofweek= "Sunday"
            hour= 3
        }
    }

    return $defaultdirectives
}

# Get directives.json, which lives on AD shares
# If it's newer than the current file, or no local file exists, download it
function Get-JSONDirectives {
    $computername = $env:COMPUTERNAME

    $directiveFileName = 'directives.json'
    $NetBase = "\\contoso.com\NETLOGON\Scripts"
    $localBase = "$env:SYSTEMDRIVE\opscode"
    $localBaseFile = "$env:SYSTEMDRIVE\opscode\$directiveFileName"
    $NetBaseFile = "$NetBase\$directiveFileName"

    # First run, check if directory exists
    if ( -not (Test-Path -Path $localBase )) {
        New-OpscodeDir
    }

    # Compare hashes, if not equal, copy new $directiveFileName to $localBase
    if ( Test-Path -Path $NetBase ) {
        
        $dhash = Get-FileHash "$localBaseFile" -Algorithm "SHA256" -ErrorAction SilentlyContinue
        $shash = Get-FileHash "$NetBaseFile" -Algorithm "SHA256"
        if ( $dhash.Hash -ne $shash.Hash ) {
            Copy-Item -Path "$NetBaseFile" -Destination "$localBaseFile" -Force
            Logit -Severity Information -Message "New version of $directiveFileName, moved to $localBase"
        }
    } else {
        Logit -Severity Error -Message "Error reaching $NetBase, will try again later."
    }

    # Load local directives.json file
    if ( Test-Path -Path $localBaseFile ) {
        $json = Get-Content $localBaseFile | Out-String | ConvertFrom-Json
    } else {
        # If there ain't no file and no network (somehow), then at least it's still managed
        Logit -Severity Error -Message "Error no directives file found in $localBaseFile using defaults"
        return $goals = Get-DefaultDirectives | ConvertTo-Json
    }

    # If computer is in an OU, apply that goal
    # ADSISearcher is super fast!
    # I have read that DirectorySearcher has a memory leak, but I have not seen this
    # TODO: Probably should put a condition around this so it doesn't run everytime
    $ADSISearcher = New-Object System.DirectoryServices.DirectorySearcher
    $ADSISearcher.Filter = '(&(name=' + $computername + ')(objectClass=computer))'
    $ADSISearcher.SearchScope = 'Subtree'
    $Computer = $ADSISearcher.FindAll()

    $OU = $($Computer.Properties.Item('distinguishedName')).Substring($($Computer.Properties.Item('distinguishedName')).IndexOf('OU='))

    # Get goals
    if ( $json.OUs.psobject.Properties.Name -contains $OU ) {
        $goals = $json.OUs.$OU
    } elseif ( $json.computers.$computername ) {

        # Filter for computername
        $goals = $json.computers.$computername
    }
 
    return $goals
}

function Test-PatchTuesday {
    $date = Get-Date
    ($date.DayOfWeek -eq 'Tuesday') -and ( @(8..15) -contains $date.Day)
}

# Choco tasks
# list: clist $app -r -localonly
# Simple for-each loop for managed and unmanaged

# Get a version, smol attempt to pin versions or something
# TODO: Add version check from the feed, that'd be nice
function Get-ChocoAppVersion {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$App
    )
    $query = @(
        $App,
        "-r",
        "-localonly"
    )
    $res = & clist @query

    $happ = @{}

    if ($res) {
        $tapp, $tver = $res.Split('|')
        $happ[$tapp] = $tver
    }

    return $happ
}

# Managedapps:
# These are managed, pinned, and version checked for drift
# Pulls from private feed
function Install-ManagedApps {
    param (
        $apphash
    )
    $source = 'https://some-chocoserver.contoso.com/chocolatey'

    $apphash.GetEnumerator() | ForEach-Object {
        $app = $_.Key
        $version = $_.Value

        $installstring = @(
            "install",
            "-s",
            $source,
            $app,
            "-y",
            "-version",
            $version
        )

        # Check version, if lower - upgrade, if higher - downgrade
        # Else install
        $appdeets = Get-ChocoAppVersion -App $app

        if ($appdeets) {
            # upgrade
            if ($appdeets.$app -lt $version) {
                $installstring[0] = "upgrade"
                & choco @installstring
            }
            # downgrade
            elseif ($appdeets.$app -gt $version) {
                # uninstall new version
                $uns = @(
                    "uninstall",
                    $app
                )
                & choco @uns
                & choco @installstring
            }
        } else {
            # install
            & choco @installstring
        }
        Logit -Severity Information -Message "Processing Managed App $app | $version Finished."

    }
}

# Un-ManagedApps would be any app that should be installed, but isn't 'managed'
# Pulls packages from community feed on the Internet
# Maybe think about checking if the node has internet access, while I'm at it
# But, just follows the schedule in the $apphash
function Install-UnManagedApps {
    param (
        $apphash
    )
    $autoapps = @('googlechrome', 'firefox', 'notepadplusplus')

    $apphash.GetEnumerator() | ForEach-Object {
        $app = $_.Key
        $version = $_.Value

        $appdeets = Get-ChocoAppVersion -App $app

        $installstring = @(
            "install",
            $app,
            "-y",
            "--ignore-checksums"
        )

        if ($appdeets) {
            # Could just skip app if already installed
            # Skip auto-updating apps
            if ($appdeets.keys -in $autoapps) {
                continue
            }
            # Could add version check from feed here
            # Update app, if available
            if ($apphash.updates.status -eq 'True') {
                if ( ($date.DayOfWeek -eq $apphash.update.DayOfWeek) -and ($date.Hour -eq $apphash.update.hour) ) {
                    $installstring[0] = "upgrade"
                    & choco @installstring
                }
            }
        } else {
            # Supporting pinning I guess        
            if ($version -ne 'Latest') {
                $installstring += "-version"
                $installstring += $version 
            }
            & choco @installstring
        }
    }
    Logit -Severity Information -Message "Finished Processing Un-managed apps"
}

# If you've got a lot of machines to manage, sharding updates is a good idea
# Not exactly sure this works properly
# compute shard and splay
# Convert UUID of system to MD5 hash
# Convert 8 digits of hex to integer
# Compute consistent integer somewhere between 1-100
function Get-Shard {
    $data = Get-CimInstance Win32_ComputerSystemProduct | Select-Object IdentifyingNumber
    $stringAsStream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($stringAsStream)
    $writer.write($data.IdentifyingNumber)
    $writer.Flush()
    $stringAsStream.Position = 0
    $streamHash = (Get-FileHash -Algorithm MD5 -InputStream $stringAsStream | Select-Object Hash).Hash
    $seed = $streamHash.Substring(0, 7)

    return [System.Convert]::ToInt64($seed, 16) % 100
}


# Shard updates
# Lower shard numbers are higher priority
function Get-NodeInShard {

    switch ($shard) {
        {$_ -in 0..9} { $goals.update.dayofweek -eq $date.AddDays(-1).DayOfWeek  }
        {$_ -in 10..25} { $goals.update.dayofweek -eq $date.AddDays(-2).DayOfWeek }
        {$_ -in 26..50} { $goals.update.dayofweek -eq $date.AddDays(-3).DayOfWeek }
        {$_ -in 51..75} { $goals.update.dayofweek -eq $date.AddDays(-4).DayOfWeek }
        {$_ -in 76..100} { $goals.update.dayofweek -eq $date.AddDays(-5).DayOfWeek }
        {$_ -gt 100 } { return $true } 
        Default { return $false }
    }
}

###
# Processing tasks begins below
###

$date = Get-Date
Logit -Severity Information -Message "Beginning node manage tasks at $date"

$shard = Get-Shard

# First step, load up the directives and store as goals, if any
$goals = Get-JSONDirectives

if ($goals) {
    # if goals, do go on
    Logit -Severity Information -Message "Found goals for: $env:COMPUTERNAME"
} else {
    # if no goals, exit gracefully
    Logit -Severity Information -Message "Exiting. No goals found for: $env:COMPUTERNAME"
    exit 0
}

# Process managedapps
if ($goals.managedapps.status -eq 'True') {
    if ($goals.managedapps.apps) {
        Install-ManagedApps -apphash $goals.managedapps.apps
    }
}

# Process unmanagedapps
if ($goals.unmanagedapps.status -eq 'True') {
    if ($goals.unmanagedapps.apps) {
        Install-UnManagedApps -apphash $goals.unmanagedapps.apps
    }
}

# Install updates
if ($goals.update.status -eq 'True') {
    if ((Get-NodeInShard) -and ($date.Hour -eq $goals.update.hour) ) {
        Logit -Severity Information -Message "Node is updating."
        UpdateOS
    }
} else {
    Logit -Severity Information -Message "Node is opted-out of updates."
}

# Reboot
if ($goals.reboottime.status -eq 'True') {
    if ( ($date.DayOfWeek -eq $goals.reboottime.DayOfWeek) -and ($date.Hour -eq $goals.reboottime.hour) ) {
        Logit -Severity Information -Message "It is time to reboot, Node is restarting"
        Restart-Computer -Force
    }
}