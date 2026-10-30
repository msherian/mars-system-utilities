#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs (or removes) the scheduled tasks that keep the WSL2 bridged
    vmSwitch pointed at whichever uplink is currently connected.

.DESCRIPTION
    Run from an elevated PowerShell opened under YOUR OWN account (not a
    separate admin account), from the folder holding Update-WslSwitch.ps1.

    Creates, under Task Scheduler \WSL\:
      Select Bridge Switch   Runs Update-WslSwitch.ps1 with highest privileges
                             (needed for Get-VMSwitch), only while you are
                             logged on. Triggers:
                               - at logon (45 s delay)
                               - NetworkProfile 10000/10001 and
                                 WLAN-AutoConfig 8001/8003 events (15 s delay)
                               - every -PollMinutes as a safety net (covers
                                 cable pulls and dock removal, which do not
                                 always raise a network profile event)
      Start Docker Desktop   Non-elevated, no triggers. Called by the main task
                             so Docker Desktop is never launched elevated.

.EXAMPLE
    .\Install-WslSwitchTask.ps1 -RunNow
.EXAMPLE
    .\Install-WslSwitchTask.ps1 -SwitchPriority ethbr0,wlanbr0 -PollMinutes 2
.EXAMPLE
    .\Install-WslSwitchTask.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [string[]] $SwitchPriority = @('ethbr0', 'wlanbr0'),
    [string]   $InstallDir     = (Join-Path $env:LOCALAPPDATA 'WslSwitchSelector'),
    [ValidateRange(1, 1440)]
    [int]      $PollMinutes    = 5,
    [string]   $DockerDesktopExe = (Join-Path $env:ProgramFiles 'Docker\Docker\Docker Desktop.exe'),
    # Extra arguments passed through to Update-WslSwitch.ps1, e.g. '-IgnoreGateway -NoToast'
    [string]   $ExtraArguments = '',
    # Use powershell.exe -WindowStyle Hidden instead of conhost --headless (pre-Win11 22H2).
    [switch]   $NoHeadlessConhost,
    [switch]   $RunNow,
    [switch]   $Uninstall
)

$ErrorActionPreference = 'Stop'
$TaskPath   = '\WSL\'
$MainTask   = 'Select Bridge Switch'
$DockerTask = 'Start Docker Desktop'

if ($Uninstall) {
    foreach ($n in $MainTask, $DockerTask) {
        if (Get-ScheduledTask -TaskPath $TaskPath -TaskName $n -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskPath $TaskPath -TaskName $n -Confirm:$false
            Write-Host "Removed $TaskPath$n"
        }
    }
    Write-Host "Script and logs left in $InstallDir"
    return
}

$SwitchPriority = @($SwitchPriority | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ })
$sid  = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$esc  = { param($s) [Security.SecurityElement]::Escape($s) }

# Sanity checks
foreach ($s in $SwitchPriority) {
    $sw = Get-VMSwitch -Name $s -ErrorAction SilentlyContinue
    if (-not $sw) { Write-Warning "Switch '$s' does not exist (yet)." }
    elseif ($sw.SwitchType -ne 'External') { Write-Warning "Switch '$s' is $($sw.SwitchType), not External." }
}
if (-not (Test-Path -LiteralPath (Join-Path $env:USERPROFILE '.wslconfig'))) {
    Write-Warning "No .wslconfig in $env:USERPROFILE. Are you running elevated as a different account?"
}

# Copy the worker script
$src = Join-Path $PSScriptRoot 'Update-WslSwitch.ps1'
if (-not (Test-Path -LiteralPath $src)) { throw "Update-WslSwitch.ps1 not found next to this installer" }
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
$dst = Join-Path $InstallDir 'Update-WslSwitch.ps1'
Copy-Item -LiteralPath $src -Destination $dst -Force
Unblock-File -LiteralPath $dst

# Build the action
$psArgs = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$dst`" -SwitchPriority $($SwitchPriority -join ',') -DockerDesktopExe `"$DockerDesktopExe`" $ExtraArguments".Trim()
if ($NoHeadlessConhost) {
    $cmd = 'powershell.exe'; $args_ = $psArgs
} else {
    # conhost --headless avoids a console (or Windows Terminal) window flashing up.
    $cmd = 'conhost.exe'; $args_ = "--headless powershell.exe $psArgs"
}

$query = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-NetworkProfile/Operational">
    <Select Path="Microsoft-Windows-NetworkProfile/Operational">*[System[(EventID=10000 or EventID=10001)]]</Select>
    <Select Path="Microsoft-Windows-WLAN-AutoConfig/Operational">*[System[(EventID=8001 or EventID=8003)]]</Select>
  </Query>
</QueryList>
"@

$start = (Get-Date).Date.ToString('yyyy-MM-ddTHH:mm:ss')

$mainXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Selects the connected Hyper-V external switch for WSL2 bridged networking, updates .wslconfig and restarts WSL and Docker Desktop when it changes.</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$sid</UserId>
      <Delay>PT45S</Delay>
    </LogonTrigger>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>$(& $esc $query)</Subscription>
      <Delay>PT15S</Delay>
    </EventTrigger>
    <TimeTrigger>
      <Repetition>
        <Interval>PT$($PollMinutes)M</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>$start</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$sid</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$(& $esc $cmd)</Command>
      <Arguments>$(& $esc $args_)</Arguments>
    </Exec>
  </Actions>
</Task>
"@

$dockerXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Launches Docker Desktop non-elevated. Called by \WSL\Select Bridge Switch.</Description>
  </RegistrationInfo>
  <Triggers />
  <Principals>
    <Principal id="Author">
      <UserId>$sid</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$(& $esc $DockerDesktopExe)</Command>
    </Exec>
  </Actions>
</Task>
"@

Register-ScheduledTask -TaskPath $TaskPath -TaskName $DockerTask -Xml $dockerXml -Force | Out-Null
Write-Host "Registered $TaskPath$DockerTask"
Register-ScheduledTask -TaskPath $TaskPath -TaskName $MainTask -Xml $mainXml -Force | Out-Null
Write-Host "Registered $TaskPath$MainTask  (priority: $($SwitchPriority -join ' > '), poll every $PollMinutes min)"
Write-Host "Log: $(Join-Path $InstallDir 'WslSwitchSelector.log')"

if ($RunNow) {
    Start-ScheduledTask -TaskPath $TaskPath -TaskName $MainTask
    Write-Host 'Started. Check the log in a few seconds.'
}