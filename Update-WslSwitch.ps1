#Requires -Version 5.1
<#
.SYNOPSIS
    Picks the WSL2 bridged vmSwitch whose physical uplink is connected, updates
    .wslconfig, then restarts WSL and Docker Desktop if the selection changed.

.DESCRIPTION
    Designed to be run by the "\WSL\Select Bridge Switch" scheduled task
    (see Install-WslSwitchTask.ps1), but safe to run by hand.

    Behaviour:
      * Does nothing unless [wsl2] networkingMode=bridged. Set networkingMode to
        nat or mirrored and the script leaves the file alone (manual override).
      * Walks -SwitchPriority in order and selects the first External switch whose
        physical uplink is Up and (if the host shares the switch) whose host vNIC
        holds a non-APIPA IPv4 address and a default gateway.
      * If nothing is usable, the current config is left untouched.
      * If the selection differs from the configured vmSwitch it waits
        -SettleSeconds, re-checks (debounce), backs up .wslconfig, rewrites only
        the vmSwitch line, stops Docker Desktop, runs wsl --shutdown and starts
        Docker Desktop again (only if it was running beforehand).
      * A cooldown stops flapping links from restarting WSL repeatedly.

    Requires Hyper-V cmdlets (elevation, or membership of "Hyper-V Administrators").

.EXAMPLE
    .\Update-WslSwitch.ps1 -Verbose -WhatIf
.EXAMPLE
    .\Update-WslSwitch.ps1 -SwitchPriority ethbr0,wlanbr0 -Force
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    # Highest priority first. Comma-separated string also accepted (for -File callers).
    [string[]] $SwitchPriority = @('ethbr0', 'wlanbr0'),

    [string] $WslConfigPath = (Join-Path $env:USERPROFILE '.wslconfig'),

    # Seconds to wait and re-check before acting (debounce).
    [int] $SettleSeconds = 10,

    # Minimum seconds between two restarts.
    [int] $CooldownSeconds = 120,

    # Only require physical link Up; skip the host vNIC IPv4/gateway check.
    [switch] $IgnoreGateway,

    # Skip stopping/starting Docker Desktop.
    [switch] $SkipDocker,

    # Start Docker Desktop afterwards even if it was not running before.
    [switch] $AlwaysStartDocker,

    # Non-elevated companion task used to launch Docker Desktop in the user context.
    [string] $DockerStartTaskPath = '\WSL\',
    [string] $DockerStartTaskName = 'Start Docker Desktop',
    [string] $DockerDesktopExe = (Join-Path $env:ProgramFiles 'Docker\Docker\Docker Desktop.exe'),

    # Rewrite and restart even if the selected switch is already configured.
    [switch] $Force,

    [switch] $NoToast,

    [string] $LogDirectory = (Join-Path $env:LOCALAPPDATA 'WslSwitchSelector')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# -File passes arrays as a single literal string, so normalise.
$SwitchPriority = @($SwitchPriority | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ })

# ---------------------------------------------------------------- logging ----
if (-not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}
$LogFile   = Join-Path $LogDirectory 'WslSwitchSelector.log'
$StateFile = Join-Path $LogDirectory 'last-change.txt'

if ((Test-Path -LiteralPath $LogFile) -and ((Get-Item -LiteralPath $LogFile).Length -gt 1MB)) {
    Move-Item -LiteralPath $LogFile -Destination "$LogFile.old" -Force
}

function Write-Log {
    param([string] $Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')] [string] $Level = 'INFO')
    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $Level, $Message
    Write-Verbose $line
    if ($Level -ne 'DEBUG' -or $VerbosePreference -ne 'SilentlyContinue') {
        Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    }
}

function Show-Toast {
    param([string] $Title, [string] $Text)
    if ($NoToast) { return }
    try {
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
        $xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
            [Windows.UI.Notifications.ToastTemplateType]::ToastText02)
        $nodes = $xml.GetElementsByTagName('text')
        $null = $nodes.Item(0).AppendChild($xml.CreateTextNode($Title))
        $null = $nodes.Item(1).AppendChild($xml.CreateTextNode($Text))
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show(
            [Windows.UI.Notifications.ToastNotification]::new($xml))
    }
    catch { Write-Log "Toast failed: $($_.Exception.Message)" 'DEBUG' }
}

# ------------------------------------------------------------- .wslconfig ----
function Read-ConfigLines {
    # ReadAllLines honours a BOM if present, otherwise assumes UTF-8.
    return , [System.Collections.Generic.List[string]]([IO.File]::ReadAllLines($WslConfigPath))
}

function Write-ConfigLines {
    param([System.Collections.Generic.List[string]] $Lines)
    # UTF-8 without BOM; WSL's parser is happiest with that.
    [IO.File]::WriteAllLines($WslConfigPath, $Lines.ToArray(), (New-Object System.Text.UTF8Encoding $false))
}

function Get-IniValue {
    param([System.Collections.Generic.List[string]] $Lines, [string] $Section, [string] $Key)
    $inSection = $false
    foreach ($l in $Lines) {
        if ($l -match '^\s*\[(.+?)\]\s*$') { $inSection = ($Matches[1] -ieq $Section); continue }
        if ($inSection -and $l -match ('^\s*' + [regex]::Escape($Key) + '\s*=\s*([^#;]*)')) {
            return $Matches[1].Trim()
        }
    }
    return $null
}

function Set-IniValue {
    # Replaces the first active "Key=" line in the section, comments out any
    # further active duplicates, and inserts the key at the end of the section
    # if it is missing. Commented lines and everything else are left alone.
    param([System.Collections.Generic.List[string]] $Lines, [string] $Section, [string] $Key, [string] $Value)
    $keyRx      = '^\s*' + [regex]::Escape($Key) + '\s*='
    $inSection  = $false
    $found      = $false
    $sectionEnd = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $l = $Lines[$i]
        if ($l -match '^\s*\[(.+?)\]\s*$') {
            if ($inSection) { break }
            $inSection = ($Matches[1] -ieq $Section)
            if ($inSection) { $sectionEnd = $i }
            continue
        }
        if (-not $inSection) { continue }
        if ($l.Trim() -and $l -notmatch '^\s*[#;]') { $sectionEnd = $i }
        if ($l -match $keyRx) {
            if (-not $found) { $Lines[$i] = "$Key=$Value"; $found = $true }
            else { $Lines[$i] = "# $l  # disabled by Update-WslSwitch.ps1" }
        }
    }
    if (-not $found) {
        if ($sectionEnd -lt 0) { throw "Section [$Section] not found in $WslConfigPath" }
        $Lines.Insert($sectionEnd + 1, "$Key=$Value")
    }
}

# --------------------------------------------------------- switch probing ----
function Get-SwitchState {
    param([string] $Name)
    $r = [pscustomobject]@{
        Name = $Name; Uplink = $null; LinkUp = $false; HostIPv4 = $null; Gateway = $null
        Usable = $false; Reason = ''
    }

    $sw = Get-VMSwitch -Name $Name -ErrorAction SilentlyContinue
    if (-not $sw) { $r.Reason = 'switch not found'; return $r }
    if ($sw.SwitchType -ne 'External') { $r.Reason = "switch type is $($sw.SwitchType)"; return $r }

    $nic = Get-NetAdapter -ErrorAction SilentlyContinue |
        Where-Object { $_.InterfaceDescription -eq $sw.NetAdapterInterfaceDescription } |
        Select-Object -First 1
    if (-not $nic) { $r.Reason = "uplink '$($sw.NetAdapterInterfaceDescription)' not present"; return $r }
    $r.Uplink = $nic.Name
    if ($nic.Status -ne 'Up') { $r.Reason = "uplink '$($nic.Name)' is $($nic.Status)"; return $r }
    $r.LinkUp = $true

    if ($sw.AllowManagementOS -and -not $IgnoreGateway) {
        $vnic = Get-VMNetworkAdapter -ManagementOS -ErrorAction SilentlyContinue |
            Where-Object { $_.SwitchName -eq $Name } | Select-Object -First 1
        $hostIf = $null
        if ($vnic) { $hostIf = Get-NetAdapter -Name "vEthernet ($($vnic.Name))" -ErrorAction SilentlyContinue }
        if ($hostIf) {
            $cfg = Get-NetIPConfiguration -InterfaceIndex $hostIf.ifIndex -ErrorAction SilentlyContinue
            if ($cfg) {
                $ip = @($cfg.IPv4Address) | Where-Object { $_ -and $_.IPAddress -notlike '169.254.*' } | Select-Object -First 1
                $gw = @($cfg.IPv4DefaultGateway) | Where-Object { $_ } | Select-Object -First 1
                if ($ip) { $r.HostIPv4 = $ip.IPAddress }
                if ($gw) { $r.Gateway = $gw.NextHop }
            }
            if (-not $r.HostIPv4 -or -not $r.Gateway) {
                $r.Reason = "link up but '$($hostIf.Name)' has no IPv4 lease/gateway"
                return $r
            }
        }
    }

    $r.Usable = $true
    $r.Reason = 'ok'
    return $r
}

function Select-TargetSwitch {
    foreach ($name in $SwitchPriority) {
        $s = Get-SwitchState -Name $name
        Write-Log ("  {0,-10} usable={1,-5} uplink={2} ip={3} gw={4} ({5})" -f
            $s.Name, $s.Usable, $s.Uplink, $s.HostIPv4, $s.Gateway, $s.Reason) 'DEBUG'
        if ($s.Usable) { return $s }
    }
    return $null
}

# ------------------------------------------------------- WSL and Docker ----
$env:WSL_UTF8 = '1'   # stop wsl.exe emitting UTF-16

function Get-RunningDistros {
    $out = & wsl.exe --list --running --quiet 2>$null
    if ($LASTEXITCODE -ne 0) { return @() }
    return @($out | ForEach-Object { ($_ -replace "`0", '').Trim() } | Where-Object { $_ })
}

$DockerProcs = @('Docker Desktop', 'com.docker.backend', 'com.docker.build', 'com.docker.dev-envs')

function Test-DockerRunning {
    return [bool](Get-Process -Name 'Docker Desktop', 'com.docker.backend' -ErrorAction SilentlyContinue)
}

function Stop-DockerDesktop {
    Write-Log 'Stopping Docker Desktop'
    if (Get-Command docker.exe -ErrorAction SilentlyContinue) {
        try { & docker.exe desktop stop 2>&1 | ForEach-Object { Write-Log "  docker: $_" 'DEBUG' } } catch { }
    }
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline -and (Test-DockerRunning)) { Start-Sleep -Seconds 2 }
    $left = Get-Process -Name $DockerProcs -ErrorAction SilentlyContinue
    if ($left) {
        Write-Log "Docker did not stop cleanly; terminating $(@($left).Count) process(es)" 'WARN'
        $left | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

function Start-DockerDesktop {
    $task = Get-ScheduledTask -TaskPath $DockerStartTaskPath -TaskName $DockerStartTaskName -ErrorAction SilentlyContinue
    if ($task) {
        Write-Log "Starting Docker Desktop via task $DockerStartTaskPath$DockerStartTaskName"
        Start-ScheduledTask -InputObject $task
    }
    elseif (Test-Path -LiteralPath $DockerDesktopExe) {
        Write-Log "Starting Docker Desktop directly ($DockerDesktopExe)" 'WARN'
        Start-Process -FilePath $DockerDesktopExe
    }
    else {
        Write-Log "Docker Desktop not found at $DockerDesktopExe" 'WARN'
    }
}

function Restart-WslStack {
    param([string] $SwitchName)
    $dockerWasRunning = (-not $SkipDocker) -and (Test-DockerRunning)
    $distros = Get-RunningDistros
    Write-Log ("Running distros: {0}; Docker running: {1}" -f ($(if ($distros) { $distros -join ', ' } else { 'none' })), $dockerWasRunning)

    if ($dockerWasRunning) { Stop-DockerDesktop }

    Write-Log 'wsl --shutdown'
    & wsl.exe --shutdown
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline -and (Get-RunningDistros)) { Start-Sleep -Seconds 1 }
    if (Get-RunningDistros) { Write-Log 'WSL still reports running distros after 30s' 'WARN' }

    if (-not $SkipDocker -and ($dockerWasRunning -or $AlwaysStartDocker)) { Start-DockerDesktop }

    Show-Toast 'WSL network switched' "WSL now bridged via $SwitchName. Distros were shut down$(if ($dockerWasRunning) { '; Docker Desktop restarting' })."
}

# ------------------------------------------------------------------ main ----
$mutex = New-Object System.Threading.Mutex($false, 'Local\WslSwitchSelector')
if (-not $mutex.WaitOne(0)) { Write-Log 'Another instance is running; exiting' 'DEBUG'; return }

try {
    if (-not (Test-Path -LiteralPath $WslConfigPath)) {
        Write-Log "$WslConfigPath not found" 'ERROR'; exit 1
    }
    if (-not (Get-Command Get-VMSwitch -ErrorAction SilentlyContinue)) {
        Write-Log 'Hyper-V PowerShell module not available' 'ERROR'; exit 1
    }

    $lines   = Read-ConfigLines
    $mode    = Get-IniValue -Lines $lines -Section 'wsl2' -Key 'networkingMode'
    $current = Get-IniValue -Lines $lines -Section 'wsl2' -Key 'vmSwitch'

    if ($mode -ne 'bridged') {
        Write-Log "networkingMode is '$mode', not bridged; leaving .wslconfig alone" 'DEBUG'
        return
    }

    $target = Select-TargetSwitch
    if (-not $target) {
        Write-Log "No usable switch in [$($SwitchPriority -join ', ')]; keeping vmSwitch=$current" 'WARN'
        return
    }
    if ($target.Name -eq $current -and -not $Force) {
        Write-Log "vmSwitch=$current is already correct" 'DEBUG'
        return
    }

    if (-not $Force -and (Test-Path -LiteralPath $StateFile)) {
        $last = [datetime]::Parse((Get-Content -LiteralPath $StateFile -Raw).Trim(), [Globalization.CultureInfo]::InvariantCulture)
        $age  = ((Get-Date) - $last).TotalSeconds
        if ($age -lt $CooldownSeconds) {
            Write-Log ("Change to {0} wanted but last change was {1:N0}s ago (cooldown {2}s); deferring" -f $target.Name, $age, $CooldownSeconds) 'WARN'
            return
        }
    }

    if ($SettleSeconds -gt 0 -and -not $Force) {
        Write-Log "Change $current -> $($target.Name) detected; settling for $SettleSeconds s"
        Start-Sleep -Seconds $SettleSeconds
        $recheck = Select-TargetSwitch
        if (-not $recheck -or $recheck.Name -ne $target.Name) {
            Write-Log "Selection changed during settle ($($target.Name) -> $(if ($recheck) { $recheck.Name } else { 'none' })); will retry next run" 'WARN'
            return
        }
        $target = $recheck
    }

    if ($PSCmdlet.ShouldProcess($WslConfigPath, "Set vmSwitch=$($target.Name) and restart WSL/Docker")) {
        Copy-Item -LiteralPath $WslConfigPath -Destination "$WslConfigPath.bak" -Force
        $lines = Read-ConfigLines
        Set-IniValue -Lines $lines -Section 'wsl2' -Key 'vmSwitch' -Value $target.Name
        Write-ConfigLines -Lines $lines
        Write-Log ("vmSwitch {0} -> {1} (uplink {2}, host IP {3}, gw {4})" -f $current, $target.Name, $target.Uplink, $target.HostIPv4, $target.Gateway)

        (Get-Date).ToString('o') | Set-Content -LiteralPath $StateFile -Encoding ASCII
        Restart-WslStack -SwitchName $target.Name
        Write-Log 'Done'
    }
}
catch {
    Write-Log "Unhandled error: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" 'ERROR'
    exit 1
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}