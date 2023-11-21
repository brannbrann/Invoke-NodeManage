<#
    .SYNOPSIS
        PS script to set up a scheduled task
        Deploy with GPO!
        If you use Chef or Puppet, you can probably use their resource to set this scheduled task
    .DESCRIPTION
        Stolen from an AWS SSM doc, heehee
        Since GPO can't create schtasks that run as Administrator, use the XML to define the task
        and import it into schtasks
        If you run Invoke-NodeManage.ps1 from a network, you may need to codesign the script, depending on your execution-policy
        This script assumes Invoke-NodeManage.ps1 is saved to /Users/Public/Invoke-NodeManage.ps1
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

$task = "Invoke_NodeManageScript"

[xml] $taskdef = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Date>2022-07-14T15:37:31.7934476</Date>
    <Author>CONTOSO\nm-svc-accountname</Author>
    <Description>Runs Invoke-NodeManage script every hour</Description>
    <URI>\Invoke_NodeManageScript</URI>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT1H</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>2022-07-14T15:33:57</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>C:\windows\system32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>-WindowStyle Hidden -NonInteractive -Executionpolicy RemoteSigned -File "C:\Users\Public\Invoke-NodeManage.ps1"</Arguments>
    </Exec>
  </Actions>
</Task>
"@

# Lay down the xml
$taskdef.Save("C:\Users\Public\task.xml")

# Set Location
Set-Location -Path "C:\Users\Public"

$schquery = "/query /TN $task"
$setschcmd = @"
/create /TN "$task" /XML "C:\Users\Public\task.xml"
"@

# If task exists, skip it
try {
    start-process "c:\windows\system32\schtasks.exe" -ArgumentList $schquery -RedirectStandardError .\stderr.txt -PassThru -Wait
}
catch {
    $err = $_.ErrorDetails
    Logit -Severity Error -Message "An unexpected error occurred while querying scheduled task: $err"
}

$stdErr = Get-Content -Path .\stderr.txt -ErrorAction SilentlyContinue
Remove-Item .\stderr.txt -ErrorAction SilentlyContinue

try {
    if ($null -ne $stdErr) {
        Logit -Severity Information -Message "Creating Scheduled task $task"
        start-process "c:\windows\system32\schtasks.exe" -ArgumentList $setschcmd -RedirectStandardError .\stderr.txt -PassThru -Wait
    }
    else {
        Logit -Severity Information -Message "Scheduled Task $task exists"
    }
}
catch {
    $err = Get-Content .\stderr.txt
    Logit -Severity Error -Message "There was an error: $err"
    Remove-Item .\stderr.txt -ErrorAction SilentlyContinue
}

Remove-Item .\task.xml -ErrorAction SilentlyContinue