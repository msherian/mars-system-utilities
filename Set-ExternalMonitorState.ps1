
<#
.SYNOPSIS
Sets external monitor power and input state.

.DESCRIPTION
Uses winddcutil to detect monitors whose descriptions match a configured pattern,
powers those monitors on, waits for them to become ready, and then switches
their input source to the configured target. The script writes operational log entries to
the configured log file and exits with a non-zero code if any required step fails.

.PARAMETER DdcUtilPath
Full path to winddcutil.exe. When omitted, the script uses the newest managed
install under LocalAppData and installs the latest GitHub release there on
first use.

.PARAMETER Monitors
Resolved monitor selections to target in the switch workflow. Tab completion is
populated from DetectOnly output and includes the winddcutil monitor ID. Use
quoted values when the selection contains spaces or parentheses.

.PARAMETER LogDirectory
Directory used to store the log file.

.PARAMETER LogFileName
Name of the log file created in LogDirectory.

.PARAMETER DetectRetryCount
Number of times to retry monitor detection before failing.

.PARAMETER DetectRetryDelaySeconds
Delay in seconds between monitor detection attempts.

.PARAMETER SetVcpRetryCount
Number of times to retry a failed setvcp command.

.PARAMETER SetVcpRetryDelaySeconds
Delay in seconds between setvcp retry attempts.

.PARAMETER PostPowerOnDelaySeconds
Delay in seconds after powering on monitors and before switching inputs.

.PARAMETER PowerModeCode
VCP feature code used to control monitor power state.

.PARAMETER PowerAction
Friendly name of the monitor power action to apply. The script maps this value
to the corresponding VCP code expected by winddcutil.

.PARAMETER InputSourceCode
VCP feature code used to change the monitor input source.

.PARAMETER InputSource
Friendly name of the monitor input source to select. The script maps this value
to the corresponding VCP code expected by winddcutil. Repeated identical values
are accepted, but only one distinct input source can be applied per run.

.PARAMETER DetectOnly
Runs monitor detection only and outputs matching monitor objects, including
friendly current and available input source names, instead of changing monitor state.

.PARAMETER Json
When used with DetectOnly, outputs the detected monitor objects as JSON.

.EXAMPLE
.\Set-ExternalMonitorState.ps1

Runs the script for the selected monitors.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -Monitors '[2] PHL 278B1','[5] Verbatim15 4K' -PostPowerOnDelaySeconds 20

Targets the selected monitors and waits 20 seconds after the power action before switching inputs.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -Monitors '[2] PHL 278B1 (27inch Wide LCD MONITOR ),[3] PHL 278B1 (27inch Wide LCD MONITOR )'

Accepts a single quoted comma-separated monitor selection list.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -InputSource HDMI1

Switches matching monitors to HDMI1.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -InputSource HDMI2,HDMI2

Accepts repeated identical input source values and applies HDMI2.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -PowerAction Sleep

Puts matching monitors into sleep mode before switching their input source.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -DetectOnly

Outputs matching detected monitors as PowerShell objects, including friendly
current and available input source names.

.EXAMPLE
.\Set-ExternalMonitorState.ps1 -DetectOnly -Json

Outputs matching detected monitors as JSON.

.NOTES
This script is intended to run from Task Scheduler or an interactive PowerShell session.
#>

[CmdletBinding(DefaultParameterSetName = 'Switch')]
param(
    [Parameter(ParameterSetName = 'Switch')]
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

            # Resolve winddcutil path — use the explicitly bound value, or probe known locations.
            $ddcUtilPath = if ($fakeBoundParameters.ContainsKey('DdcUtilPath')) {
                [string]$fakeBoundParameters.DdcUtilPath
            }
            else {
                @(
                    if (-not [string]::IsNullOrWhiteSpace($env:WINDDCUTIL_HOME)) {
                        Join-Path -Path $env:WINDDCUTIL_HOME -ChildPath 'bin\winddcutil.exe'
                    }
                    Join-Path -Path ${env:ProgramFiles} -ChildPath 'winddcutil\winddcutil.exe'
                    'c:\programs\winddcutil\winddcutil.exe'
                ) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
            }

            if ([string]::IsNullOrWhiteSpace($ddcUtilPath) -or -not (Test-Path -LiteralPath $ddcUtilPath -PathType Leaf)) {
                [System.Management.Automation.CompletionResult]::new("''", '(winddcutil not found)', 'ParameterValue', 'Cannot locate winddcutil.exe.')
                return
            }

            # Helper: invoke winddcutil with the given arguments; return combined stdout+stderr.
            $runWinddcutil = {
                param([string]$Exe, [string[]]$CmdArgs)
                try {
                    $si = [System.Diagnostics.ProcessStartInfo]::new()
                    $si.FileName = $Exe
                    $si.UseShellExecute = $false
                    $si.RedirectStandardOutput = $true
                    $si.RedirectStandardError = $true
                    $si.CreateNoWindow = $true
                    foreach ($a in $CmdArgs) { [void]$si.ArgumentList.Add($a) }
                    $p = [System.Diagnostics.Process]::new()
                    $p.StartInfo = $si
                    [void]$p.Start()
                    $stderrTask = $p.StandardError.ReadToEndAsync()
                    $out = $p.StandardOutput.ReadToEnd()
                    $out += $stderrTask.GetAwaiter().GetResult()
                    $p.WaitForExit()
                    return $out
                }
                catch { return '' }
            }

            # Detect monitors.
            $detectText = & $runWinddcutil $ddcUtilPath @('detect')
            $detected = [System.Collections.Generic.List[pscustomobject]]::new()
            foreach ($line in ($detectText -split "`r?`n")) {
                if ($line -match '^\s*(?<Id>\d+)\s+(?<Description>.+?)\s*$') {
                    [void]$detected.Add([pscustomobject]@{
                            Id                = $Matches.Id
                            Description       = $Matches.Description.Trim()
                            CapabilitiesModel = ''
                        })
                }
            }

            # For Generic PnP Monitors: fetch capabilities → extract model() for WMI matching.
            foreach ($det in @($detected | Where-Object { $_.Description -eq 'Generic PnP Monitor' })) {
                $capText = (& $runWinddcutil $ddcUtilPath @('capabilities', $det.Id)).Trim()
                if ($capText -match 'model\((?<Model>[^\)]+)\)') { $det.CapabilitiesModel = $Matches.Model.Trim() }
            }

            # WMI monitor details.
            $cimMonitors = [System.Collections.Generic.List[pscustomobject]]::new()
            try {
                foreach ($w in @(Get-CimInstance -ClassName WMIMonitorID -Namespace root\wmi -ErrorAction Stop)) {
                    $fn  = (-join [char[]](@($w.UserFriendlyName | Where-Object { $_ -gt 0 }))).Trim()
                    $sn  = (-join [char[]](@($w.SerialNumberID   | Where-Object { $_ -gt 0 }))).Trim()
                    $mfr = (-join [char[]](@($w.ManufacturerName  | Where-Object { $_ -gt 0 }))).Trim()
                    $inst = [string]$w.InstanceName
                    $dc  = if (($inst -split '\\').Count -ge 2) { ($inst -split '\\')[1] } else { '' }
                    [void]$cimMonitors.Add([pscustomobject]@{
                            UserFriendlyName = $fn; SerialNumber = $sn
                            Manufacturer     = $mfr; DeviceCode = $dc; Active = [bool]$w.Active
                        })
                }
            }
            catch { }

            # Score a detected monitor against a WMI entry (mirrors Add-CimMonitorDetails).
            $scoreCim = {
                param($det, $cim)
                $s   = 0
                $du  = $det.Description.ToUpperInvariant()
                $dcu = $cim.DeviceCode.ToUpperInvariant()
                $fnu = $cim.UserFriendlyName.ToUpperInvariant()
                $mfu = $cim.Manufacturer.ToUpperInvariant()
                if ($dcu -and $du.Contains($dcu)) { $s += 10 }
                if ($fnu -and $du.Contains($fnu)) { $s += 5 }
                if ($cim.Active)                  { $s += 1 }
                $dt = @([regex]::Replace($du,  '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })
                foreach ($t in @([regex]::Replace($fnu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($dt -contains $t) { $s += 1 } }
                foreach ($t in @([regex]::Replace($mfu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($dt -contains $t) { $s += 1 } }
                if (-not [string]::IsNullOrWhiteSpace($det.CapabilitiesModel)) {
                    $cmu = $det.CapabilitiesModel.ToUpperInvariant()
                    if ($fnu -and $cmu.Contains($fnu)) { $s += 8 }
                    if ($mfu -and $cmu.Contains($mfu)) { $s += 6 }
                    $cmt = @([regex]::Replace($cmu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })
                    foreach ($t in @([regex]::Replace($fnu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($cmt -contains $t) { $s += 2 } }
                    foreach ($t in @([regex]::Replace($mfu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($cmt -contains $t) { $s += 2 } }
                }
                $s
            }

            # Assign WMI entries to detected monitors (specific monitors first, generics after).
            $cimAssignments = @{}
            $availableCim = [System.Collections.Generic.List[pscustomobject]]::new($cimMonitors)
            foreach ($pass in @('specific', 'generic')) {
                foreach ($det in @($detected | Where-Object { ($_.Description -eq 'Generic PnP Monitor') -eq ($pass -eq 'generic') })) {
                    if ($cimAssignments.ContainsKey($det.Id)) { continue }
                    $best = $null; $bestScore = 0
                    foreach ($cim in @($availableCim)) {
                        $s = & $scoreCim $det $cim
                        if ($s -gt $bestScore) { $bestScore = $s; $best = $cim }
                    }
                    if ($null -ne $best -and $bestScore -gt 1) {
                        $cimAssignments[$det.Id] = $best
                        [void]$availableCim.Remove($best)
                    }
                }
            }

            # Build enriched list with resolved descriptions.
            $enriched = @(
                $detected | ForEach-Object {
                    $cim  = $cimAssignments[$_.Id]
                    $desc = $_.Description
                    if ($desc -eq 'Generic PnP Monitor') {
                        if ($null -ne $cim -and -not [string]::IsNullOrWhiteSpace($cim.UserFriendlyName)) {
                            $desc = $cim.UserFriendlyName
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace($_.CapabilitiesModel)) {
                            $desc = $_.CapabilitiesModel
                        }
                    }
                    [pscustomobject]@{
                        Id          = $_.Id
                        Description = $desc
                        SerialNumber = if ($null -ne $cim) { $cim.SerialNumber } else { '' }
                    }
                }
            )

            if ($enriched.Count -eq 0) {
                [System.Management.Automation.CompletionResult]::new("''", '(No monitors detected)', 'ParameterValue', 'winddcutil detect returned no monitors.')
                return
            }

            foreach ($m in $enriched) {
                $selectionValue = '[{0}] {1}' -f $m.Id, $m.Description
                $quotedValue    = "'{0}'" -f $selectionValue.Replace("'", "''")
                $tooltip = if (-not [string]::IsNullOrWhiteSpace($m.SerialNumber)) {
                    'Id {0}: {1} (S/N {2})' -f $m.Id, $m.Description, $m.SerialNumber
                }
                else { 'Id {0}: {1}' -f $m.Id, $m.Description }
                $matchesWord = [string]::IsNullOrWhiteSpace($wordToComplete) -or
                    $selectionValue -like "*$wordToComplete*" -or
                    $quotedValue -like "*$wordToComplete*" -or
                    $m.Description -like "*$wordToComplete*"
                if ($matchesWord) {
                    [System.Management.Automation.CompletionResult]::new($quotedValue, $selectionValue, 'ParameterValue', $tooltip)
                }
            }
        })]
    [string[]]$Monitors,
    [ValidateSet('On', 'Off', 'Sleep')]
    [string]$PowerAction = 'On',
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

            # Same map as $InputSourceValueMap in end{} — friendly name → VCP code.
            $InputSourceValueMap = @{
                'DisplayPort-1' = '0x0F'
                'DisplayPort-2' = '0x10'
                'HDMI1'         = '0x11'
                'HDMI2'         = '0x12'
            }
            # Derive VCP hex → name lookup for capabilities parsing.
            $vcpToName = @{}
            foreach ($kv in $InputSourceValueMap.GetEnumerator()) {
                $hex = ($kv.Value -replace '^0[xX]', '').ToUpperInvariant().PadLeft(2, '0')
                $vcpToName[$hex] = $kv.Key
            }

            # Default when winddcutil is unavailable or no monitors are selected.
            $availableInputSources = @($InputSourceValueMap.Keys)

            # Resolve winddcutil path — use the explicitly bound value, or probe known locations.
            $ddcUtilPath = if ($fakeBoundParameters.ContainsKey('DdcUtilPath')) {
                [string]$fakeBoundParameters.DdcUtilPath
            }
            else {
                @(
                    if (-not [string]::IsNullOrWhiteSpace($env:WINDDCUTIL_HOME)) {
                        Join-Path -Path $env:WINDDCUTIL_HOME -ChildPath 'bin\winddcutil.exe'
                    }
                    Join-Path -Path ${env:ProgramFiles} -ChildPath 'winddcutil\winddcutil.exe'
                    'c:\programs\winddcutil\winddcutil.exe'
                ) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
            }

            if (-not [string]::IsNullOrWhiteSpace($ddcUtilPath) -and (Test-Path -LiteralPath $ddcUtilPath -PathType Leaf)) {

                # Helper: invoke winddcutil with the given arguments; return combined stdout+stderr.
                $runWinddcutil = {
                    param([string]$Exe, [string[]]$CmdArgs)
                    try {
                        $si = [System.Diagnostics.ProcessStartInfo]::new()
                        $si.FileName = $Exe
                        $si.UseShellExecute = $false
                        $si.RedirectStandardOutput = $true
                        $si.RedirectStandardError = $true
                        $si.CreateNoWindow = $true
                        foreach ($a in $CmdArgs) { [void]$si.ArgumentList.Add($a) }
                        $p = [System.Diagnostics.Process]::new()
                        $p.StartInfo = $si
                        [void]$p.Start()
                        $stderrTask = $p.StandardError.ReadToEndAsync()
                        $out = $p.StandardOutput.ReadToEnd()
                        $out += $stderrTask.GetAwaiter().GetResult()
                        $p.WaitForExit()
                        return $out
                    }
                    catch { return '' }
                }

                # Detect monitors.
                $detectText = & $runWinddcutil $ddcUtilPath @('detect')
                $detected = [System.Collections.Generic.List[pscustomobject]]::new()
                foreach ($line in ($detectText -split "`r?`n")) {
                    if ($line -match '^\s*(?<Id>\d+)\s+(?<Description>.+?)\s*$') {
                        [void]$detected.Add([pscustomobject]@{
                                Id                = $Matches.Id
                                Description       = $Matches.Description.Trim()
                                CapabilitiesModel = ''
                            })
                    }
                }

                # For Generic PnP Monitors: fetch capabilities → extract model() for WMI matching.
                foreach ($det in @($detected | Where-Object { $_.Description -eq 'Generic PnP Monitor' })) {
                    $capText = (& $runWinddcutil $ddcUtilPath @('capabilities', $det.Id)).Trim()
                    if ($capText -match 'model\((?<Model>[^\)]+)\)') { $det.CapabilitiesModel = $Matches.Model.Trim() }
                }

                # WMI monitor details.
                $cimMonitors = [System.Collections.Generic.List[pscustomobject]]::new()
                try {
                    foreach ($w in @(Get-CimInstance -ClassName WMIMonitorID -Namespace root\wmi -ErrorAction Stop)) {
                        $fn  = (-join [char[]](@($w.UserFriendlyName | Where-Object { $_ -gt 0 }))).Trim()
                        $mfr = (-join [char[]](@($w.ManufacturerName  | Where-Object { $_ -gt 0 }))).Trim()
                        $inst = [string]$w.InstanceName
                        $dc  = if (($inst -split '\\').Count -ge 2) { ($inst -split '\\')[1] } else { '' }
                        [void]$cimMonitors.Add([pscustomobject]@{
                                UserFriendlyName = $fn; Manufacturer = $mfr
                                DeviceCode = $dc; Active = [bool]$w.Active
                            })
                    }
                }
                catch { }

                # Score a detected monitor against a WMI entry (mirrors Add-CimMonitorDetails).
                $scoreCim = {
                    param($det, $cim)
                    $s   = 0
                    $du  = $det.Description.ToUpperInvariant()
                    $dcu = $cim.DeviceCode.ToUpperInvariant()
                    $fnu = $cim.UserFriendlyName.ToUpperInvariant()
                    $mfu = $cim.Manufacturer.ToUpperInvariant()
                    if ($dcu -and $du.Contains($dcu)) { $s += 10 }
                    if ($fnu -and $du.Contains($fnu)) { $s += 5 }
                    if ($cim.Active)                  { $s += 1 }
                    $dt = @([regex]::Replace($du,  '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })
                    foreach ($t in @([regex]::Replace($fnu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($dt -contains $t) { $s += 1 } }
                    foreach ($t in @([regex]::Replace($mfu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($dt -contains $t) { $s += 1 } }
                    if (-not [string]::IsNullOrWhiteSpace($det.CapabilitiesModel)) {
                        $cmu = $det.CapabilitiesModel.ToUpperInvariant()
                        if ($fnu -and $cmu.Contains($fnu)) { $s += 8 }
                        if ($mfu -and $cmu.Contains($mfu)) { $s += 6 }
                        $cmt = @([regex]::Replace($cmu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })
                        foreach ($t in @([regex]::Replace($fnu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($cmt -contains $t) { $s += 2 } }
                        foreach ($t in @([regex]::Replace($mfu, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })) { if ($cmt -contains $t) { $s += 2 } }
                    }
                    $s
                }

                # Assign WMI entries to detected monitors (specific monitors first, generics after).
                $cimAssignments = @{}
                $availableCim = [System.Collections.Generic.List[pscustomobject]]::new($cimMonitors)
                foreach ($pass in @('specific', 'generic')) {
                    foreach ($det in @($detected | Where-Object { ($_.Description -eq 'Generic PnP Monitor') -eq ($pass -eq 'generic') })) {
                        if ($cimAssignments.ContainsKey($det.Id)) { continue }
                        $best = $null; $bestScore = 0
                        foreach ($cim in @($availableCim)) {
                            $s = & $scoreCim $det $cim
                            if ($s -gt $bestScore) { $bestScore = $s; $best = $cim }
                        }
                        if ($null -ne $best -and $bestScore -gt 1) {
                            $cimAssignments[$det.Id] = $best
                            [void]$availableCim.Remove($best)
                        }
                    }
                }

                # Build enriched list with resolved descriptions.
                $enriched = @(
                    $detected | ForEach-Object {
                        $cim  = $cimAssignments[$_.Id]
                        $desc = $_.Description
                        if ($desc -eq 'Generic PnP Monitor') {
                            if ($null -ne $cim -and -not [string]::IsNullOrWhiteSpace($cim.UserFriendlyName)) {
                                $desc = $cim.UserFriendlyName
                            }
                            elseif (-not [string]::IsNullOrWhiteSpace($_.CapabilitiesModel)) {
                                $desc = $_.CapabilitiesModel
                            }
                        }
                        [pscustomobject]@{ Id = $_.Id; Description = $desc }
                    }
                )

                if ($fakeBoundParameters.ContainsKey('Monitors')) {
                    # Parse selected monitor selections into (Id, Label) pairs.
                    $selectedMonitors = [System.Collections.Generic.List[pscustomobject]]::new()
                    foreach ($rawSel in @($fakeBoundParameters.Monitors)) {
                        foreach ($part in @([regex]::Split([string]$rawSel, '\s*,\s*(?=\[\d+\]\s+)'))) {
                            if ($part -match '^\[(?<Id>\d+)\]\s+(?<Label>.+?)\s*$') {
                                [void]$selectedMonitors.Add([pscustomobject]@{ Id = [string]$Matches.Id; Label = $Matches.Label.Trim() })
                            }
                        }
                    }

                    if ($selectedMonitors.Count -gt 0) {
                        # Token-score enriched descriptions against the last selected label to resolve
                        # the current Id — handles winddcutil detect Id instability across completions.
                        $lastMon     = $selectedMonitors[$selectedMonitors.Count - 1]
                        $labelUp     = $lastMon.Label.ToUpperInvariant()
                        $labelTokens = @([regex]::Replace($labelUp, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })
                        $resolvedId  = $lastMon.Id   # fallback: stored Id
                        $bestTotal   = 0
                        foreach ($e in $enriched) {
                            $edUp   = $e.Description.ToUpperInvariant()
                            $edToks = @([regex]::Replace($edUp, '[^A-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_.Length -ge 3 })
                            $ts     = @($labelTokens | Where-Object { $edToks -contains $_ }).Count
                            $ib     = if ($e.Id -eq $lastMon.Id) { 1 } else { 0 }
                            $tot    = $ts * 2 + $ib
                            if ($ts -gt 0 -and $tot -gt $bestTotal) { $bestTotal = $tot; $resolvedId = $e.Id }
                        }

                        # Get VCP 60 input sources for the resolved monitor.
                        $capText = (& $runWinddcutil $ddcUtilPath @('capabilities', $resolvedId)).Trim()
                        $availableInputSources = @()
                        if ($capText -match '60\((?<Values>[^\)]+)\)') {
                            $availableInputSources = @(
                                foreach ($v in @($Matches.Values -split '\s+')) {
                                    $nv = $v.Trim().ToUpperInvariant().PadLeft(2, '0')
                                    if ($vcpToName.ContainsKey($nv)) { $vcpToName[$nv] }
                                }
                            ) | Select-Object -Unique
                        }

                        if ($availableInputSources.Count -eq 0) {
                            foreach ($mon in $selectedMonitors) {
                                $lbl = '[{0}] {1}' -f $mon.Id, $mon.Label
                                [System.Management.Automation.CompletionResult]::new(
                                    ("'<No input source: {0}>'" -f $lbl),
                                    ("Warning: input cannot be selected for '{0}'" -f $lbl),
                                    'ParameterValue',
                                    ("'{0}' does not support input source switching via DDC/CI." -f $lbl)
                                )
                            }
                            return
                        }
                    }
                }
            }

            foreach ($name in $availableInputSources) {
                if ([string]::IsNullOrWhiteSpace($wordToComplete) -or $name -like "$wordToComplete*") {
                    [System.Management.Automation.CompletionResult]::new($name, $name, 'ParameterValue', $name)
                }
            }
        })]
    [string[]]$InputSource = @('DisplayPort-1'),
    [Parameter(ParameterSetName = 'DetectOnly')]
    [switch]$DetectOnly,
    [Parameter(ParameterSetName = 'DetectOnly')]
    [switch]$Json,
    [string]$DdcUtilPath = 'c:\programs\winddcutil\winddcutil.exe',
    [int]$PostPowerOnDelaySeconds = 5,
    [string]$LogDirectory = 'C:\Windows\Logs\Set-ExternalMonitorState',
    [string]$LogFileName = 'Set-ExternalMonitorState.log',
    [int]$DetectRetryCount = 12,
    [int]$DetectRetryDelaySeconds = 10,
    [int]$SetVcpRetryCount = 3,
    [int]$SetVcpRetryDelaySeconds = 5,
    [string]$PowerModeCode = 'D6',
    [string]$InputSourceCode = '0x60'
)

begin {
    # Stored at script scope so functions inside end{} can use it without a subprocess call.
    $script:DefaultMonitorSerialNumbers = $null
    if (-not $DetectOnly -and -not $PSBoundParameters.ContainsKey('Monitors')) {
        $script:DefaultMonitorSerialNumbers = @(
            'UKC2035000146'
            'UKC2035000138'
        )
    }
}

end {

    # Initialize cached path before Set-StrictMode so referencing it unset doesn't throw.
    $script:ResolvedDdcUtilCommandPath = $null

    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $LogPath = Join-Path -Path $LogDirectory -ChildPath $LogFileName
    $FallbackLogDirectory = Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath 'Set-ExternalMonitorState\Logs'
    $LogFallbackActivated = $false
    $MaxLogSizeBytes = 100MB
    $MaxLogAge = [TimeSpan]::FromDays(14)
    $TaskSchedulerHistoryLogName = 'Microsoft-Windows-TaskScheduler/Operational'
    $InputSourceValueMap = @{
        'DisplayPort-1'  = '0x0F'
        'DisplayPort-2'  = '0x10'
        'HDMI1'          = '0x11'
        'HDMI2'          = '0x12'
    }
    $PowerActionValueMap = @{
        On    = '1'
        Off   = '4'
        Sleep = '5'
    }
    $SelectedInputSource = @(
        @($InputSource) |
        ForEach-Object { [regex]::Split([string]$_, '\s*,\s*') } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )
    if ($SelectedInputSource.Count -ne 1) {
        throw "InputSource must contain exactly one distinct value. Received: $($SelectedInputSource -join ', ')"
    }
    $SelectedInputSource = $SelectedInputSource[0]
    $InputSourceSpecified = $PSBoundParameters.ContainsKey('InputSource')

    if (-not $InputSourceValueMap.ContainsKey($SelectedInputSource)) {
        throw "Unsupported InputSource '$SelectedInputSource'. Supported values: $($InputSourceValueMap.Keys -join ', ')"
    }

    $SelectedPowerActionValue = $PowerActionValueMap[$PowerAction]
    $SelectedInputSourceValue = $InputSourceValueMap[$SelectedInputSource]

    <#
.SYNOPSIS
Writes a formatted log entry to the console and log file.

.DESCRIPTION
Creates a timestamped log entry with a severity level, writes the entry to the
host, and appends it to the configured log file.

.PARAMETER Message
Text to write to the log.

.PARAMETER Level
Severity level for the log entry.

.EXAMPLE
Write-Log -Message 'Starting monitor switch.'

Writes an informational log entry.
#>
    function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Message,

            [ValidateSet('INFO', 'WARN', 'ERROR')]
            [string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $entry = "[$timestamp] [$Level] $Message"

        Write-Host $entry

        try {
            Add-Content -Path $LogPath -Value $entry -ErrorAction Stop
        }
        catch {
            if (-not $script:LogFallbackActivated) {
                try {
                    Set-LogLocation -Path $FallbackLogDirectory
                    Initialize-LogDirectory -Path $script:LogDirectory
                    $script:LogFallbackActivated = $true
                    Write-Host "[$timestamp] [WARN] Falling back to log path '$script:LogPath' because the configured log path was unavailable."
                    Add-Content -Path $script:LogPath -Value $entry -ErrorAction Stop
                    return
                }
                catch {
                }
            }
        }
    }

    function Set-LogLocation {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $script:LogDirectory = $Path
        $script:LogPath = Join-Path -Path $script:LogDirectory -ChildPath $script:LogFileName
    }

    function Initialize-LogDirectory {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        if (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -ItemType Directory -Force | Out-Null
        }

        $probePath = Join-Path -Path $Path -ChildPath ([System.IO.Path]::GetRandomFileName())

        try {
            Set-Content -Path $probePath -Value 'write test' -Encoding ASCII -ErrorAction Stop
            Remove-Item -Path $probePath -Force -ErrorAction Stop
        }
        catch {
            throw "The current user cannot write to log directory '$Path'. $($_.Exception.Message)"
        }
    }

    function Get-TaskSchedulerEventData {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [System.Diagnostics.Eventing.Reader.EventRecord]$EventRecord
        )

        $eventData = @{}
        [xml]$eventXml = $EventRecord.ToXml()

        foreach ($dataNode in @($eventXml.Event.EventData.Data)) {
            $dataName = [string]$dataNode.Name
            if ([string]::IsNullOrWhiteSpace($dataName)) {
                continue
            }

            $eventData[$dataName] = [string]$dataNode.InnerText
        }

        return $eventData
    }

    function ConvertTo-NullableDateTime {
        [CmdletBinding()]
        param(
            [AllowEmptyString()]
            [AllowNull()]
            [string]$Value
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return $null
        }

        try {
            return [datetime]::Parse(
                $Value,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind
            )
        }
        catch {
            return $null
        }
    }

    function ConvertTo-IntOrDefault {
        [CmdletBinding()]
        param(
            [AllowEmptyString()]
            [AllowNull()]
            [string]$Value,

            [int]$Default = 0
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return $Default
        }

        try {
            return [int]$Value
        }
        catch {
            return $Default
        }
    }

    function Test-WinddcutilSystemInstallContext {
        [CmdletBinding()]
        param()

        $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $currentIdentityName = [string]$currentIdentity.Name
        if ($currentIdentityName -ieq 'NT AUTHORITY\SYSTEM' -or $currentIdentityName -ieq 'NT SERVICE\TrustedInstaller') {
            return $true
        }

        $principal = [System.Security.Principal.WindowsPrincipal]::new($currentIdentity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    function Get-WinddcutilInstallScope {
        [CmdletBinding()]
        param()

        if (Test-WinddcutilSystemInstallContext) {
            return 'Machine'
        }

        return 'User'
    }

    function Get-WinddcutilInstallRoot {
        [CmdletBinding()]
        param(
            [string]$Scope = (Get-WinddcutilInstallScope)
        )

        if ($Scope -eq 'Machine') {
            return (Join-Path -Path ${env:ProgramFiles} -ChildPath 'winddcutil')
        }

        return (Join-Path -Path $env:LOCALAPPDATA -ChildPath 'Programs\winddcutil')
    }

    function Get-WinddcutilInstallRoots {
        [CmdletBinding()]
        param()

        if ((Get-WinddcutilInstallScope) -eq 'Machine') {
            return @(
                Get-WinddcutilInstallRoot -Scope 'Machine'
                Get-WinddcutilInstallRoot -Scope 'User'
            )
        }

        return @(
            Get-WinddcutilInstallRoot -Scope 'User'
            Get-WinddcutilInstallRoot -Scope 'Machine'
        )
    }

    function ConvertTo-NormalizedWinddcutilVersion {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Version
        )

        $trimmedVersion = $Version.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedVersion)) {
            throw 'winddcutil version is null or empty.'
        }

        return $trimmedVersion.TrimStart('v', 'V')
    }

    function Get-WinddcutilInstallPathForVersion {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Version,

            [string]$Scope = (Get-WinddcutilInstallScope)
        )

        $normalizedVersion = ConvertTo-NormalizedWinddcutilVersion -Version $Version
        return (Join-Path -Path (Get-WinddcutilInstallRoot -Scope $Scope) -ChildPath (Join-Path -Path $normalizedVersion -ChildPath 'bin\winddcutil.exe'))
    }

    function Get-WinddcutilHomeForVersion {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Version,

            [string]$Scope = (Get-WinddcutilInstallScope)
        )

        $normalizedVersion = ConvertTo-NormalizedWinddcutilVersion -Version $Version
        return (Join-Path -Path (Get-WinddcutilInstallRoot -Scope $Scope) -ChildPath $normalizedVersion)
    }

    function Get-InstalledWinddcutilPath {
        [CmdletBinding()]
        param()

        $installedVersion = @(
            foreach ($installRoot in @(Get-WinddcutilInstallRoots)) {
                if (-not (Test-Path -LiteralPath $installRoot -PathType Container)) {
                    continue
                }

                Get-ChildItem -LiteralPath $installRoot -Directory -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $candidatePath = Join-Path -Path $_.FullName -ChildPath 'bin\winddcutil.exe'
                    if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
                        return
                    }

                    $version = $null
                    if (-not [version]::TryParse($_.Name.TrimStart('v', 'V'), [ref]$version)) {
                        return
                    }

                    [pscustomobject]@{
                        Version = $version
                        Path    = $candidatePath
                    }
                }
            }
        ) |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1

        if ($null -eq $installedVersion) {
            return $null
        }

        return [string]$installedVersion.Path
    }

    function Get-LatestWinddcutilReleaseInfo {
        [CmdletBinding()]
        param()

        $releaseUri = 'https://api.github.com/repos/scottaxcell/winddcutil/releases/latest'

        try {
            $release = Invoke-RestMethod -Uri $releaseUri -Headers @{
                'Accept'     = 'application/vnd.github+json'
                'User-Agent' = 'Set-ExternalMonitorState'
            } -ErrorAction Stop
        }
        catch {
            throw "Failed to query the latest winddcutil release from '$releaseUri'. $($_.Exception.Message)"
        }

        $asset = @($release.assets | Where-Object { $_.name -eq 'winddcutil.exe' }) | Select-Object -First 1
        if ($null -eq $asset -or [string]::IsNullOrWhiteSpace([string]$asset.browser_download_url)) {
            throw 'The latest winddcutil release did not expose a winddcutil.exe asset.'
        }

        $normalizedVersion = ConvertTo-NormalizedWinddcutilVersion -Version ([string]$release.tag_name)
        $installScope = Get-WinddcutilInstallScope

        return [pscustomobject]@{
            Version      = $normalizedVersion
            DownloadUrl  = [string]$asset.browser_download_url
            InstallScope = $installScope
            InstallHome  = Get-WinddcutilHomeForVersion -Version $normalizedVersion -Scope $installScope
            InstallPath  = Get-WinddcutilInstallPathForVersion -Version $normalizedVersion -Scope $installScope
        }
    }

    function Update-WinddcutilEnvironment {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$InstallHome,

            [ValidateSet('User', 'Machine')]
            [string]$Scope
        )

        $pathEntry = '%WINDDCUTIL_HOME%\bin'
        [Environment]::SetEnvironmentVariable('WINDDCUTIL_HOME', $InstallHome, $Scope)
        $env:WINDDCUTIL_HOME = $InstallHome

        $persistedPath = [Environment]::GetEnvironmentVariable('Path', $Scope)
        $pathEntries = @(
            @([string]$persistedPath -split ';') |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        $hasPathEntry = $false
        foreach ($existingPathEntry in $pathEntries) {
            if (
                $existingPathEntry.Trim().TrimEnd('\') -ieq $pathEntry.Trim().TrimEnd('\') -or
                $existingPathEntry.Trim().TrimEnd('\') -ieq (Join-Path -Path $InstallHome -ChildPath 'bin').TrimEnd('\')
            ) {
                $hasPathEntry = $true
                break
            }
        }

        if (-not $hasPathEntry) {
            $updatedPathEntries = @($pathEntries + $pathEntry)
            [Environment]::SetEnvironmentVariable('Path', ($updatedPathEntries -join ';'), $Scope)

            $processPathEntries = @(
                @([string]$env:Path -split ';') |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )

            if ($processPathEntries -notcontains (Join-Path -Path $InstallHome -ChildPath 'bin')) {
                $env:Path = (@($processPathEntries + (Join-Path -Path $InstallHome -ChildPath 'bin')) -join ';')
            }
        }
    }

    function Install-Winddcutil {
        [CmdletBinding()]
        param()

        $releaseInfo = Get-LatestWinddcutilReleaseInfo
        $installHome = [string]$releaseInfo.InstallHome
        $installPath = [string]$releaseInfo.InstallPath
        $installScope = [string]$releaseInfo.InstallScope

        if (Test-Path -LiteralPath $installPath -PathType Leaf) {
            return (Get-Item -LiteralPath $installPath -ErrorAction Stop).FullName
        }

        $installDirectory = Split-Path -Path $installPath -Parent
        if (-not (Test-Path -LiteralPath $installDirectory -PathType Container)) {
            New-Item -Path $installDirectory -ItemType Directory -Force | Out-Null
        }

        $temporaryDownloadPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())

        try {
            Invoke-WebRequest -Uri $releaseInfo.DownloadUrl -OutFile $temporaryDownloadPath -Headers @{ 'User-Agent' = 'Set-ExternalMonitorState' } -ErrorAction Stop
            Move-Item -LiteralPath $temporaryDownloadPath -Destination $installPath -Force
        }
        catch {
            if (Test-Path -LiteralPath $temporaryDownloadPath -PathType Leaf) {
                Remove-Item -LiteralPath $temporaryDownloadPath -Force -ErrorAction SilentlyContinue
            }

            throw "Failed to install winddcutil $($releaseInfo.Version) to '$installPath'. $($_.Exception.Message)"
        }

        Update-WinddcutilEnvironment -InstallHome $installHome -Scope $installScope

        return (Get-Item -LiteralPath $installPath -ErrorAction Stop).FullName
    }

    function Resolve-ExistingWinddcutilPath {
        [CmdletBinding()]
        param(
            [AllowEmptyString()]
            [AllowNull()]
            [string]$RequestedPath,

            [switch]$WasExplicitlyProvided
        )

        if ($WasExplicitlyProvided) {
            if ([string]::IsNullOrWhiteSpace($RequestedPath)) {
                throw 'DdcUtilPath is null or empty.'
            }

            if (-not (Test-Path -LiteralPath $RequestedPath -PathType Leaf)) {
                throw "winddcutil not found at '$RequestedPath'."
            }

            return (Get-Item -LiteralPath $RequestedPath -ErrorAction Stop).FullName
        }

        if (-not [string]::IsNullOrWhiteSpace($RequestedPath) -and (Test-Path -LiteralPath $RequestedPath -PathType Leaf)) {
            return (Get-Item -LiteralPath $RequestedPath -ErrorAction Stop).FullName
        }

        # Check WINDDCUTIL_HOME set by Install-Winddcutil / Update-WinddcutilEnvironment.
        if (-not [string]::IsNullOrWhiteSpace($env:WINDDCUTIL_HOME)) {
            $homeExe = Join-Path -Path $env:WINDDCUTIL_HOME -ChildPath 'bin\winddcutil.exe'
            if (Test-Path -LiteralPath $homeExe -PathType Leaf) {
                return (Get-Item -LiteralPath $homeExe -ErrorAction Stop).FullName
            }
        }

        # Check direct (non-versioned) install in ProgramFiles.
        $programFilesExe = Join-Path -Path ${env:ProgramFiles} -ChildPath 'winddcutil\winddcutil.exe'
        if (Test-Path -LiteralPath $programFilesExe -PathType Leaf) {
            return (Get-Item -LiteralPath $programFilesExe -ErrorAction Stop).FullName
        }

        $installedPath = Get-InstalledWinddcutilPath
        if (-not [string]::IsNullOrWhiteSpace($installedPath)) {
            return $installedPath
        }

        $command = Get-Command -Name 'winddcutil.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace([string]$command.Source)) {
            return [string]$command.Source
        }

        return $null
    }

    function Resolve-WinddcutilPath {
        [CmdletBinding()]
        param(
            [AllowEmptyString()]
            [AllowNull()]
            [string]$RequestedPath,

            [switch]$WasExplicitlyProvided
        )

        $resolvedPath = Resolve-ExistingWinddcutilPath -RequestedPath $RequestedPath -WasExplicitlyProvided:$WasExplicitlyProvided
        if (-not [string]::IsNullOrWhiteSpace($resolvedPath)) {
            return $resolvedPath
        }

        return (Install-Winddcutil)
    }

    function Get-ProcessAncestry {
        [CmdletBinding()]
        param()

        $processChain = [System.Collections.Generic.List[object]]::new()
        $visitedProcessIds = [System.Collections.Generic.HashSet[int]]::new()
        $currentProcessId = [int]$PID

        while ($currentProcessId -gt 0 -and $visitedProcessIds.Add($currentProcessId)) {
            try {
                $processRecord = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $currentProcessId" -ErrorAction Stop
                $processObject = Get-Process -Id $currentProcessId -ErrorAction Stop
            }
            catch {
                break
            }

            [void]$processChain.Add([pscustomobject]@{
                    ProcessId       = [int]$processRecord.ProcessId
                    ParentProcessId = [int]$processRecord.ParentProcessId
                    Name            = [string]$processObject.ProcessName
                    CommandLine     = [string]$processRecord.CommandLine
                    StartTime       = [datetime]$processObject.StartTime
                })

            if ($processRecord.ParentProcessId -le 0 -or $processRecord.ParentProcessId -eq $processRecord.ProcessId) {
                break
            }

            $currentProcessId = [int]$processRecord.ParentProcessId
        }

        return @($processChain)
    }

    function Get-InvocationContext {
        [CmdletBinding()]
        param()

        $forwardedInvocationMode = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_INVOCATION_MODE')
        if (-not [string]::IsNullOrWhiteSpace($forwardedInvocationMode)) {
            $forwardedTaskStartTime = ConvertTo-NullableDateTime -Value ([Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_START_TIME'))
            $forwardedTaskLaunchTime = ConvertTo-NullableDateTime -Value ([Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_LAUNCH_TIME'))
            $forwardedWrapperProcessId = ConvertTo-IntOrDefault -Value ([Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_WRAPPER_PROCESS_ID'))

            return [pscustomobject]@{
                Mode                      = $forwardedInvocationMode
                Reason                    = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_INVOCATION_REASON')
                ProcessChain              = 'forwarded by wrapper'
                TaskName                  = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_NAME')
                TaskInstanceId            = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_INSTANCE_ID')
                TaskActionName            = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_ACTION_NAME')
                TaskTriggerSource         = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_TRIGGER_SOURCE')
                TaskTriggerReason         = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_TASK_TRIGGER_REASON')
                TaskLaunchTime            = $forwardedTaskLaunchTime
                TaskStartTime             = $forwardedTaskStartTime
                MatchedProcessId          = $forwardedWrapperProcessId
                MatchedProcessName        = [Environment]::GetEnvironmentVariable('SET_EXTERNAL_MONITOR_WRAPPER_PROCESS_NAME')
                MatchedProcessCommandLine = $null
            }
        }

        $processAncestry = @(Get-ProcessAncestry)
        $processChain = if ($processAncestry.Count -gt 0) {
            ($processAncestry | ForEach-Object { '{0}[{1}]' -f $_.Name, $_.ProcessId }) -join ' <- '
        }
        else {
            'unavailable'
        }

        if ($processAncestry.Count -eq 0) {
            return [pscustomobject]@{
                Mode         = 'Unknown'
                Reason       = 'Process ancestry could not be resolved.'
                ProcessChain = $processChain
            }
        }

        $processesById = @{}
        foreach ($processInfo in $processAncestry) {
            $processesById[[int]$processInfo.ProcessId] = $processInfo
        }

        try {
            $launchEvents = @(
                Get-WinEvent -FilterHashtable @{
                    LogName   = $TaskSchedulerHistoryLogName
                    Id        = 129
                    StartTime = (Get-Date).AddMinutes(-30)
                } -ErrorAction Stop
            )
        }
        catch {
            return [pscustomobject]@{
                Mode         = 'Unknown'
                Reason       = "Task Scheduler history could not be queried. $($_.Exception.Message)"
                ProcessChain = $processChain
            }
        }

        $launchCandidates = @(
            foreach ($launchEvent in $launchEvents) {
                $eventData = Get-TaskSchedulerEventData -EventRecord $launchEvent
                $launchProcessId = 0
                if (-not [int]::TryParse([string]$eventData.ProcessID, [ref]$launchProcessId)) {
                    continue
                }

                if (-not $processesById.ContainsKey($launchProcessId)) {
                    continue
                }

                $matchedProcess = $processesById[$launchProcessId]
                $timeDeltaSeconds = [math]::Abs(($launchEvent.TimeCreated - $matchedProcess.StartTime).TotalSeconds)
                if ($timeDeltaSeconds -gt 600) {
                    continue
                }

                [pscustomobject]@{
                    Event            = $launchEvent
                    EventData        = $eventData
                    MatchedProcess   = $matchedProcess
                    IsCurrentProcess = ($matchedProcess.ProcessId -eq $PID)
                    TimeDeltaSeconds = $timeDeltaSeconds
                }
            }
        )

        $selectedLaunch = @(
            $launchCandidates |
            Sort-Object -Property @{ Expression = 'IsCurrentProcess'; Descending = $true }, @{ Expression = 'TimeDeltaSeconds'; Descending = $false }, @{ Expression = { $_.Event.TimeCreated }; Descending = $true }
        ) | Select-Object -First 1

        if ($null -eq $selectedLaunch) {
            return [pscustomobject]@{
                Mode         = 'Interactive'
                Reason       = 'No Task Scheduler launch event matched the current process ancestry.'
                ProcessChain = $processChain
            }
        }

        $taskName = [string]$selectedLaunch.EventData.TaskName
        $relatedEvents = @(
            Get-WinEvent -FilterHashtable @{
                LogName   = $TaskSchedulerHistoryLogName
                Id        = 100, 200
                StartTime = $selectedLaunch.Event.TimeCreated.AddMinutes(-2)
            } -ErrorAction SilentlyContinue
        )

        $taskHistory = @(
            foreach ($relatedEvent in $relatedEvents) {
                $relatedData = Get-TaskSchedulerEventData -EventRecord $relatedEvent
                if ([string]$relatedData.TaskName -ne $taskName) {
                    continue
                }

                [pscustomobject]@{
                    Event     = $relatedEvent
                    EventData = $relatedData
                }
            }
        )

        $startEvent = @(
            $taskHistory |
            Where-Object { $_.Event.Id -eq 100 } |
            Sort-Object -Property @{ Expression = { [math]::Abs(($_.Event.TimeCreated - $selectedLaunch.Event.TimeCreated).TotalSeconds) }; Descending = $false }, @{ Expression = { $_.Event.TimeCreated }; Descending = $false }
        ) | Select-Object -First 1

        $actionStartEvent = @(
            $taskHistory |
            Where-Object { $_.Event.Id -eq 200 } |
            Sort-Object -Property @{ Expression = { [math]::Abs(($_.Event.TimeCreated - $selectedLaunch.Event.TimeCreated).TotalSeconds) }; Descending = $false }, @{ Expression = { $_.Event.TimeCreated }; Descending = $false }
        ) | Select-Object -First 1

        $taskInstanceId = if ($null -ne $actionStartEvent) {
            [string]$actionStartEvent.EventData.TaskInstanceId
        }
        elseif ($null -ne $startEvent) {
            [string]$startEvent.EventData.InstanceId
        }
        else {
            $null
        }

        return [pscustomobject]@{
            Mode                      = 'ScheduledTask'
            Reason                    = 'Matched Task Scheduler launch history to the current process ancestry.'
            ProcessChain              = $processChain
            TaskName                  = $taskName
            TaskInstanceId            = $taskInstanceId
            TaskActionName            = if ($null -ne $actionStartEvent) { [string]$actionStartEvent.EventData.ActionName } else { $null }
            TaskLaunchTime            = [datetime]$selectedLaunch.Event.TimeCreated
            TaskStartTime             = if ($null -ne $startEvent) { [datetime]$startEvent.Event.TimeCreated } else { [datetime]$selectedLaunch.Event.TimeCreated }
            MatchedProcessId          = [int]$selectedLaunch.MatchedProcess.ProcessId
            MatchedProcessName        = [string]$selectedLaunch.MatchedProcess.Name
            MatchedProcessCommandLine = [string]$selectedLaunch.MatchedProcess.CommandLine
        }
    }

    function Write-InvocationContextLog {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [pscustomobject]$InvocationContext,

            [Parameter(Mandatory)]
            [string]$Stage
        )

        if ($InvocationContext.Mode -eq 'ScheduledTask') {
            $taskStartTime = if ($null -ne $InvocationContext.TaskStartTime) {
                $InvocationContext.TaskStartTime.ToString('yyyy-MM-dd HH:mm:ss')
            }
            else {
                'unavailable'
            }

            $taskInstanceText = if ([string]::IsNullOrWhiteSpace([string]$InvocationContext.TaskInstanceId)) {
                'unavailable'
            }
            else {
                [string]$InvocationContext.TaskInstanceId
            }

            $actionText = if ([string]::IsNullOrWhiteSpace([string]$InvocationContext.TaskActionName)) {
                'unavailable'
            }
            else {
                [string]$InvocationContext.TaskActionName
            }

            $triggerText = if ([string]::IsNullOrWhiteSpace([string]$InvocationContext.TaskTriggerSource)) {
                'Unknown'
            }
            else {
                [string]$InvocationContext.TaskTriggerSource
            }

            $triggerReasonText = if ([string]::IsNullOrWhiteSpace([string]$InvocationContext.TaskTriggerReason)) {
                'Trigger classification details were unavailable.'
            }
            else {
                [string]$InvocationContext.TaskTriggerReason
            }

            Write-Log -Message "$Stage invocation context: scheduled task '$($InvocationContext.TaskName)' matched via $($InvocationContext.MatchedProcessName) [$($InvocationContext.MatchedProcessId)]. Trigger: $triggerText. Instance: $taskInstanceText. Action: $actionText. Task history start: $taskStartTime. $triggerReasonText"
            return
        }

        Write-Log -Message "$Stage invocation context: $($InvocationContext.Mode). $($InvocationContext.Reason) Process chain: $($InvocationContext.ProcessChain)"
    }

    <#
.SYNOPSIS
Invokes winddcutil and returns structured command output.

.DESCRIPTION
Calls winddcutil with one of the supported subcommands, validates the executable
path, captures stdout and stderr, throws on launch or command failures unless
IgnoreExitCode is specified, and returns a JSON-backed PowerShell object with
raw and parsed data.

.PARAMETER Detect
Runs the winddcutil detect subcommand.

.PARAMETER Capabilities
Runs the winddcutil capabilities subcommand.

.PARAMETER SetVcp
Runs the winddcutil setvcp subcommand.

.PARAMETER GetVcp
Runs the winddcutil getvcp subcommand.

.PARAMETER Display
Display identifier used by capabilities, setvcp, or getvcp.

.PARAMETER FeatureCode
VCP feature code used by setvcp or getvcp.

.PARAMETER NewValue
VCP value supplied to setvcp.

.PARAMETER IgnoreExitCode
Returns structured output even when winddcutil exits with a non-zero code.

.EXAMPLE
Invoke-Winddcutil -Detect

Returns parsed monitor information from the detect subcommand.

.EXAMPLE
Invoke-Winddcutil -SetVcp -Display 3 -FeatureCode 0x60 -NewValue 0x0F

Sets the input source for display 3 to the specified VCP value.
#>
    function Invoke-Winddcutil {
        [CmdletBinding(DefaultParameterSetName = 'detect')]
        param(
            [Parameter(Mandatory, ParameterSetName = 'detect')]
            [switch]$Detect,

            [Parameter(Mandatory, ParameterSetName = 'capabilities')]
            [switch]$Capabilities,

            [Parameter(Mandatory, ParameterSetName = 'setvcp')]
            [switch]$SetVcp,

            [Parameter(Mandatory, ParameterSetName = 'getvcp')]
            [switch]$GetVcp,

            [Parameter(Mandatory, ParameterSetName = 'capabilities')]
            [Parameter(Mandatory, ParameterSetName = 'setvcp')]
            [Parameter(Mandatory, ParameterSetName = 'getvcp')]
            [ValidatePattern('^\d+$')]
            [string]$Display,

            [Parameter(Mandatory, ParameterSetName = 'setvcp')]
            [Parameter(Mandatory, ParameterSetName = 'getvcp')]
            [string]$FeatureCode,

            [Parameter(Mandatory, ParameterSetName = 'setvcp')]
            [string]$NewValue,

            [switch]$IgnoreExitCode
        )

        if (-not $script:ResolvedDdcUtilCommandPath) {
            if ([string]::IsNullOrWhiteSpace($DdcUtilPath)) {
                throw 'DdcUtilPath is null or empty.'
            }

            if (-not (Test-Path -LiteralPath $DdcUtilPath -PathType Leaf)) {
                throw "winddcutil not found at '$DdcUtilPath'."
            }

            $script:ResolvedDdcUtilCommandPath = (Get-Item -LiteralPath $DdcUtilPath -ErrorAction Stop).FullName
        }

        $commandPath = $script:ResolvedDdcUtilCommandPath
        $commandName = switch ($PSCmdlet.ParameterSetName) {
            'detect' { 'detect' }
            'capabilities' { 'capabilities' }
            'setvcp' { 'setvcp' }
            'getvcp' { 'getvcp' }
            default { throw "Unsupported parameter set '$($PSCmdlet.ParameterSetName)'." }
        }

        [string[]]$resultArguments = switch ($PSCmdlet.ParameterSetName) {
            'detect' { @($commandName) }
            'capabilities' { @($commandName, $Display) }
            'setvcp' { @($commandName, $Display, $FeatureCode, $NewValue) }
            'getvcp' { @($commandName, $Display, $FeatureCode) }
            default { throw "Unsupported parameter set '$($PSCmdlet.ParameterSetName)'." }
        }

        try {
            $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
            $startInfo.FileName = $commandPath
            $startInfo.UseShellExecute = $false
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            $startInfo.CreateNoWindow = $true

            [void]$startInfo.ArgumentList.Add($commandName)

            switch ($PSCmdlet.ParameterSetName) {
                'capabilities' {
                    [void]$startInfo.ArgumentList.Add($Display)
                }
                'setvcp' {
                    [void]$startInfo.ArgumentList.Add($Display)
                    [void]$startInfo.ArgumentList.Add($FeatureCode)
                    [void]$startInfo.ArgumentList.Add($NewValue)
                }
                'getvcp' {
                    [void]$startInfo.ArgumentList.Add($Display)
                    [void]$startInfo.ArgumentList.Add($FeatureCode)
                }
            }

            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo = $startInfo
            [void]$process.Start()

            $standardOutput = $process.StandardOutput.ReadToEnd()
            $standardError = $process.StandardError.ReadToEnd()
            $process.WaitForExit()
            $exitCode = $process.ExitCode
        }
        catch {
            throw "Failed to start winddcutil from '$commandPath'. $($_.Exception.Message)"
        }

        $rawOutput = @(
            @($standardOutput -split "`r?`n") |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if (-not [string]::IsNullOrWhiteSpace($standardError)) {
            $rawOutput += @(
                @($standardError -split "`r?`n") |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }

        if ($null -eq $exitCode) {
            throw "winddcutil did not return an exit code for command '$commandName'. Raw output: $(($rawOutput | Out-String).Trim())"
        }

        if (-not $IgnoreExitCode -and $exitCode -ne 0) {
            $detail = if ($rawOutput) { ($rawOutput | Out-String).Trim() } else { 'No output returned.' }
            throw "winddcutil failed with exit code $exitCode. Args: $($resultArguments -join ' '). Output: $detail"
        }

        $result = [ordered]@{
            Command   = $commandName
            Arguments = $resultArguments
            ExitCode  = $exitCode
            Succeeded = ($exitCode -eq 0)
            RawOutput = $rawOutput
        }

        switch ($commandName) {
            'detect' {
                $result['Monitors'] = @(
                    $rawOutput |
                    Where-Object { $_ -match '^\s*\d+\s+' } |
                    ForEach-Object {
                        if ($_ -match '^\s*(?<Id>\d+)\s+(?<Description>.+?)\s*$') {
                            [ordered]@{
                                Id          = [int]$Matches.Id
                                Description = $Matches.Description.Trim()
                            }
                        }
                    }
                )
            }
            'capabilities' {
                $result['MonitorId'] = $Display
                $result['Capabilities'] = ($rawOutput -join [Environment]::NewLine).Trim()
            }
            'setvcp' {
                $result['MonitorId'] = $Display
                $result['Code'] = $FeatureCode
                $result['Value'] = $NewValue
            }
            'getvcp' {
                $result['MonitorId'] = $Display
                $result['Code'] = $FeatureCode
                $result['Value'] = ($rawOutput -join [Environment]::NewLine).Trim()
            }
        }

        return [pscustomobject]$result
    }

    <#
.SYNOPSIS
Gets parsed monitor objects from the winddcutil detect command.

.DESCRIPTION
Runs the detect subcommand using a direct native process invocation and returns
parsed monitor objects with Id and Description properties.

.PARAMETER IgnoreExitCode
Returns an empty set of monitors when the detect command fails instead of throwing.

.OUTPUTS
PSCustomObject[]

.EXAMPLE
Get-DetectedMonitors

Returns all detected monitors as objects.
#>
    function Get-DetectedMonitors {
        [CmdletBinding()]
        param(
            [switch]$IgnoreExitCode
        )

        if (-not $script:ResolvedDdcUtilCommandPath) {
            if (-not (Test-Path -LiteralPath $DdcUtilPath -PathType Leaf)) {
                throw "winddcutil not found at '$DdcUtilPath'."
            }

            $script:ResolvedDdcUtilCommandPath = (Get-Item -LiteralPath $DdcUtilPath -ErrorAction Stop).FullName
        }

        $commandPath = $script:ResolvedDdcUtilCommandPath

        try {
            $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
            $startInfo.FileName = $commandPath
            $startInfo.UseShellExecute = $false
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            $startInfo.CreateNoWindow = $true
            [void]$startInfo.ArgumentList.Add('detect')

            $process = [System.Diagnostics.Process]::new()
            $process.StartInfo = $startInfo
            [void]$process.Start()

            $standardOutput = $process.StandardOutput.ReadToEnd()
            $standardError = $process.StandardError.ReadToEnd()
            $process.WaitForExit()
            $exitCode = $process.ExitCode
        }
        catch {
            throw "Failed to start winddcutil from '$commandPath'. $($_.Exception.Message)"
        }

        $rawOutput = @(
            @($standardOutput -split "`r?`n") |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if (-not [string]::IsNullOrWhiteSpace($standardError)) {
            $rawOutput += @(
                @($standardError -split "`r?`n") |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }

        if (-not $IgnoreExitCode -and $exitCode -ne 0) {
            $detail = if ($rawOutput) { ($rawOutput | Out-String).Trim() } else { 'No output returned.' }
            throw "winddcutil failed with exit code $exitCode. Args: detect. Output: $detail"
        }

        return @(
            $rawOutput |
            Where-Object { $_ -match '^\s*\d+\s+' } |
            ForEach-Object {
                if ($_ -match '^\s*(?<Id>\d+)\s+(?<Description>.+?)\s*$') {
                    [pscustomobject]@{
                        Id          = [int]$Matches.Id
                        Description = $Matches.Description.Trim()
                    }
                }
            }
        )
    }

    <#
.SYNOPSIS
Gets monitor details from WMIMonitorID.

.DESCRIPTION
Queries the root\wmi namespace for WMIMonitorID instances and returns friendly
monitor properties normalized from the underlying character arrays.

.OUTPUTS
PSCustomObject[]

.EXAMPLE
Get-CimMonitorDetails

Returns monitor details reported by WMIMonitorID.
#>
    function Get-CimMonitorDetails {
        [CmdletBinding()]
        param()

        try {
            $query = Get-CimInstance -ClassName WMIMonitorID -Namespace root\wmi -ErrorAction Stop
        }
        catch {
            return @()
        }

        return @(
            @($query) |
            ForEach-Object {
                $monitor = $_
                $instanceName = [string]$monitor.InstanceName
                $instanceSegments = @($instanceName -split '\\')
                $deviceCode = if ($instanceSegments.Count -ge 2) { $instanceSegments[1] } else { '' }

                [pscustomobject]@{
                    InstanceName      = $instanceName
                    DeviceCode        = $deviceCode
                    ComputerName      = $env:COMPUTERNAME
                    Active            = $monitor.Active
                    Manufacturer      = ( -join [char[]](@($monitor.ManufacturerName | Where-Object { $_ -gt 0 }))).Trim()
                    UserFriendlyName  = ( -join [char[]](@($monitor.UserFriendlyName | Where-Object { $_ -gt 0 }))).Trim()
                    SerialNumber      = ( -join [char[]](@($monitor.SerialNumberID | Where-Object { $_ -gt 0 }))).Trim()
                    WeekOfManufacture = $monitor.WeekOfManufacture
                    YearOfManufacture = $monitor.YearOfManufacture
                }
            }
        )
    }

    function Get-MonitorMatchTokens {
        [CmdletBinding()]
        param(
            [AllowEmptyString()]
            [string[]]$Values
        )

        return @(
            foreach ($value in @($Values)) {
                if ([string]::IsNullOrWhiteSpace($value)) {
                    continue
                }

                $normalizedValue = ([regex]::Replace($value.ToUpperInvariant(), '[^A-Z0-9]+', ' ')).Trim()
                if ([string]::IsNullOrWhiteSpace($normalizedValue)) {
                    continue
                }

                foreach ($token in @($normalizedValue -split '\s+')) {
                    if ($token.Length -ge 3) {
                        $token
                    }
                }
            }
        ) | Select-Object -Unique
    }

    function Get-MonitorCapabilitiesModel {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId
        )

        try {
            $capabilitiesResult = Invoke-Winddcutil -Capabilities -Display $MonitorId -IgnoreExitCode
            $capabilitiesText = [string]$capabilitiesResult.Capabilities
        }
        catch {
            return $null
        }

        if ($capabilitiesText -match 'model\((?<Model>[^\)]+)\)') {
            return $Matches.Model.Trim()
        }

        return $null
    }

    function Get-MonitorInputSourceDescriptor {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [int]$Value
        )

        $normalizedHexValue = '{0:X2}' -f ($Value -band 0xFF)

        switch ($normalizedHexValue) {
            '0F' {
                return [pscustomobject]@{
                    Code         = '0x0F'
                    IsKnown      = $true
                    LegacyName   = 'DisplayPort-1'
                    FriendlyName = 'DisplayPort-1'
                }
            }
            '10' {
                return [pscustomobject]@{
                    Code         = '0x10'
                    IsKnown      = $true
                    LegacyName   = 'DisplayPort-2'
                    FriendlyName = 'DisplayPort-2'
                }
            }
            '11' {
                return [pscustomobject]@{
                    Code         = '0x11'
                    IsKnown      = $true
                    LegacyName   = 'HDMI-1'
                    FriendlyName = 'HDMI-1'
                }
            }
            '12' {
                return [pscustomobject]@{
                    Code         = '0x12'
                    IsKnown      = $true
                    LegacyName   = 'HDM-2'
                    FriendlyName = 'HDMI-2'
                }
            }
            '13' {
                return [pscustomobject]@{
                    Code         = '0x13'
                    IsKnown      = $true
                    LegacyName   = $null
                    FriendlyName = 'HDMI-3'
                }
            }
            '14' {
                return [pscustomobject]@{
                    Code         = '0x14'
                    IsKnown      = $true
                    LegacyName   = $null
                    FriendlyName = 'HDMI-4'
                }
            }
            '1B' {
                return [pscustomobject]@{
                    Code         = '0x1B'
                    IsKnown      = $true
                    LegacyName   = $null
                    FriendlyName = 'USB-C-1'
                }
            }
            '1C' {
                return [pscustomobject]@{
                    Code         = '0x1C'
                    IsKnown      = $true
                    LegacyName   = $null
                    FriendlyName = 'USB-C-2'
                }
            }
            default {
                return [pscustomobject]@{
                    Code         = ('0x{0}' -f $normalizedHexValue)
                    IsKnown      = $false
                    LegacyName   = $null
                    FriendlyName = ('Input-{0}' -f ('0x{0}' -f $normalizedHexValue))
                }
            }
        }
    }

    function Get-MonitorSupportedInputSourceDescriptors {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId
        )

        $capabilitiesResult = Invoke-Winddcutil -Capabilities -Display $MonitorId -IgnoreExitCode
        $capabilitiesText = [string]$capabilitiesResult.Capabilities
        if ([string]::IsNullOrWhiteSpace($capabilitiesText)) {
            return @()
        }

        if (-not ($capabilitiesText -match '60\((?<Values>[^\)]+)\)')) {
            return @()
        }

        return @(
            foreach ($value in @($Matches.Values -split '\s+')) {
                $normalizedValue = $value.Trim().ToUpperInvariant()
                if ($normalizedValue -notmatch '^[0-9A-F]{1,2}$') {
                    continue
                }

                Get-MonitorInputSourceDescriptor -Value ([Convert]::ToInt32($normalizedValue, 16))
            }
        ) | Sort-Object -Property FriendlyName -Unique
    }

    function Get-MonitorSupportedInputSources {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId
        )

        return @(
            foreach ($inputSourceDescriptor in @(Get-MonitorSupportedInputSourceDescriptors -MonitorId $MonitorId)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$inputSourceDescriptor.LegacyName)) {
                    [string]$inputSourceDescriptor.LegacyName
                }
            }
        ) | Select-Object -Unique
    }

    function Get-MonitorCurrentInputSourceDescriptor {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId
        )

        $currentInputValue = Get-MonitorVcpValue -MonitorId $MonitorId -Code $InputSourceCode
        $normalizedCurrentInputValue = Normalize-VcpValueForComparison -Code $InputSourceCode -Value $currentInputValue

        return (Get-MonitorInputSourceDescriptor -Value $normalizedCurrentInputValue)
    }

    <#
.SYNOPSIS
Adds WMIMonitorID details to detected monitor objects.

.DESCRIPTION
Merges winddcutil detect results with WMIMonitorID monitor metadata using a
one-to-one assignment so each detected monitor is paired with at most one WMI
record. Specific matches are assigned first, and any remaining WMI records are
used to resolve generic monitor descriptions.

.PARAMETER Monitors
Detected monitor objects to enrich.

.PARAMETER CimMonitorDetails
Monitor details returned by Get-CimMonitorDetails.

.OUTPUTS
PSCustomObject[]
#>
    function Add-CimMonitorDetails {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [object[]]$Monitors,

            [Parameter(Mandatory)]
            [object[]]$CimMonitorDetails
        )

        function Get-MonitorMatchScore {
            param(
                [Parameter(Mandatory)]
                [pscustomobject]$Monitor,

                [Parameter(Mandatory)]
                [pscustomobject]$CimMonitor
            )

            $description = [string]$Monitor.Description
            $descriptionTokens = @(Get-MonitorMatchTokens -Values @($description, $Monitor.CapabilitiesModel))
            $candidateTokens = @(Get-MonitorMatchTokens -Values @($CimMonitor.UserFriendlyName, $CimMonitor.Manufacturer, $CimMonitor.DeviceCode))
            $sharedTokens = @($candidateTokens | Where-Object { $descriptionTokens -contains $_ })
            $score = $sharedTokens.Count

            if (-not [string]::IsNullOrWhiteSpace($CimMonitor.DeviceCode) -and $description.ToUpperInvariant().Contains($CimMonitor.DeviceCode.ToUpperInvariant())) {
                $score += 10
            }

            if (-not [string]::IsNullOrWhiteSpace($CimMonitor.UserFriendlyName) -and $description.ToUpperInvariant().Contains($CimMonitor.UserFriendlyName.ToUpperInvariant())) {
                $score += 5
            }

            if (
                -not [string]::IsNullOrWhiteSpace([string]$Monitor.CapabilitiesModel) -and
                -not [string]::IsNullOrWhiteSpace([string]$CimMonitor.UserFriendlyName) -and
                $Monitor.CapabilitiesModel.ToUpperInvariant().Contains($CimMonitor.UserFriendlyName.ToUpperInvariant())
            ) {
                $score += 8
            }

            if (
                -not [string]::IsNullOrWhiteSpace([string]$Monitor.CapabilitiesModel) -and
                -not [string]::IsNullOrWhiteSpace([string]$CimMonitor.Manufacturer) -and
                $Monitor.CapabilitiesModel.ToUpperInvariant().Contains($CimMonitor.Manufacturer.ToUpperInvariant())
            ) {
                $score += 6
            }

            if ($CimMonitor.Active) {
                $score += 1
            }

            return $score
        }

        function New-EnrichedMonitorObject {
            param(
                [Parameter(Mandatory)]
                [pscustomobject]$Monitor,

                [pscustomobject]$CimMonitor
            )

            $resolvedDescription = [string]$Monitor.Description
            if (
                $null -ne $CimMonitor -and
                $resolvedDescription -eq 'Generic PnP Monitor' -and
                -not [string]::IsNullOrWhiteSpace([string]$CimMonitor.UserFriendlyName)
            ) {
                $resolvedDescription = [string]$CimMonitor.UserFriendlyName
            }
            elseif (
                $resolvedDescription -eq 'Generic PnP Monitor' -and
                -not [string]::IsNullOrWhiteSpace([string]$Monitor.CapabilitiesModel)
            ) {
                $resolvedDescription = [string]$Monitor.CapabilitiesModel
            }

            $result = [ordered]@{
                Id                  = $Monitor.Id
                Description         = $Monitor.Description
                ResolvedDescription = $resolvedDescription
                CapabilitiesModel   = $Monitor.CapabilitiesModel
                ComputerName        = $null
                Active              = $null
                Manufacturer        = $null
                UserFriendlyName    = $null
                SerialNumber        = $null
                WeekOfManufacture   = $null
                YearOfManufacture   = $null
            }

            if ($null -ne $CimMonitor) {
                $result.ComputerName = $CimMonitor.ComputerName
                $result.Active = $CimMonitor.Active
                $result.Manufacturer = $CimMonitor.Manufacturer
                $result.UserFriendlyName = $CimMonitor.UserFriendlyName
                $result.SerialNumber = $CimMonitor.SerialNumber
                $result.WeekOfManufacture = $CimMonitor.WeekOfManufacture
                $result.YearOfManufacture = $CimMonitor.YearOfManufacture
            }

            return [pscustomobject]$result
        }

        $availableCimMonitors = [System.Collections.Generic.List[object]]::new()
        foreach ($cimMonitor in @($CimMonitorDetails)) {
            [void]$availableCimMonitors.Add($cimMonitor)
        }

        $assignments = @{}
        $detectedMonitors = @($Monitors)

        foreach ($monitor in @($detectedMonitors | Where-Object { $_.Description -ne 'Generic PnP Monitor' })) {
            $candidate = @(
                @($availableCimMonitors) |
                ForEach-Object {
                    [pscustomobject]@{
                        CimMonitor = $_
                        Score      = Get-MonitorMatchScore -Monitor $monitor -CimMonitor $_
                    }
                } |
                Where-Object { $_.Score -gt 0 } |
                Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = { [bool]$_.CimMonitor.Active }; Descending = $true }
            ) | Select-Object -First 1

            if ($null -ne $candidate) {
                $assignments[[string]$monitor.Id] = $candidate.CimMonitor
                [void]$availableCimMonitors.Remove($candidate.CimMonitor)
            }
        }

        foreach ($monitor in @($detectedMonitors | Where-Object { -not $assignments.ContainsKey([string]$_.Id) })) {
            $selectedCimMonitor = $null
            $candidate = @(
                @($availableCimMonitors) |
                ForEach-Object {
                    [pscustomobject]@{
                        CimMonitor = $_
                        Score      = Get-MonitorMatchScore -Monitor $monitor -CimMonitor $_
                    }
                } |
                Where-Object { $_.Score -gt 0 } |
                Sort-Object -Property @{ Expression = 'Score'; Descending = $true }, @{ Expression = { [bool]$_.CimMonitor.Active }; Descending = $true }
            ) | Select-Object -First 1

            if ($null -ne $candidate) {
                $selectedCimMonitor = $candidate.CimMonitor
            }

            if ($null -ne $selectedCimMonitor) {
                $assignments[[string]$monitor.Id] = $selectedCimMonitor
                [void]$availableCimMonitors.Remove($selectedCimMonitor)
            }
        }

        return @(
            foreach ($monitor in $detectedMonitors) {
                $assignedCimMonitor = $null
                if ($assignments.ContainsKey([string]$monitor.Id)) {
                    $assignedCimMonitor = $assignments[[string]$monitor.Id]
                }

                New-EnrichedMonitorObject -Monitor $monitor -CimMonitor $assignedCimMonitor
            }
        )
    }

    <#
.SYNOPSIS
Gets detected monitors enriched with WMIMonitorID details.

.DESCRIPTION
Combines winddcutil detect output with WMIMonitorID metadata and returns the
same enriched monitor objects used by DetectOnly.

.PARAMETER IgnoreExitCode
Returns any available detect results even when winddcutil exits with a non-zero code.

.OUTPUTS
PSCustomObject[]
#>
    function Get-ResolvedDetectedMonitors {
        [CmdletBinding()]
        param(
            [switch]$IgnoreExitCode
        )

        $detectResult = @(
            Get-DetectedMonitors -IgnoreExitCode:$IgnoreExitCode |
            ForEach-Object {
                $capabilitiesModel = $null
                if ([string]$_.Description -eq 'Generic PnP Monitor') {
                    $capabilitiesModel = Get-MonitorCapabilitiesModel -MonitorId ([string]$_.Id)
                }

                [pscustomobject]@{
                    Id                = $_.Id
                    Description       = $_.Description
                    CapabilitiesModel = $capabilitiesModel
                }
            }
        )
        $cimMonitorDetails = @(Get-CimMonitorDetails)

        return @(Add-CimMonitorDetails -Monitors @($detectResult) -CimMonitorDetails $cimMonitorDetails)
    }

    function Get-DetectOnlyMonitors {
        [CmdletBinding()]
        param(
            [switch]$IgnoreExitCode
        )

        return @(
            foreach ($resolvedMonitor in @(Get-ResolvedDetectedMonitors -IgnoreExitCode:$IgnoreExitCode)) {
                $availableInputSourceDescriptors = @()
                $currentInputSourceDescriptor = $null

                try {
                    $availableInputSourceDescriptors = @(Get-MonitorSupportedInputSourceDescriptors -MonitorId ([string]$resolvedMonitor.Id))
                }
                catch {
                    $availableInputSourceDescriptors = @()
                }

                try {
                    $currentInputSourceDescriptor = Get-MonitorCurrentInputSourceDescriptor -MonitorId ([string]$resolvedMonitor.Id)
                }
                catch {
                    $currentInputSourceDescriptor = $null
                }

                [pscustomobject][ordered]@{
                    Id                        = $resolvedMonitor.Id
                    Description               = $resolvedMonitor.Description
                    ResolvedDescription       = $resolvedMonitor.ResolvedDescription
                    CapabilitiesModel         = $resolvedMonitor.CapabilitiesModel
                    ComputerName              = $resolvedMonitor.ComputerName
                    Active                    = $resolvedMonitor.Active
                    Manufacturer              = $resolvedMonitor.Manufacturer
                    UserFriendlyName          = $resolvedMonitor.UserFriendlyName
                    SerialNumber              = $resolvedMonitor.SerialNumber
                    WeekOfManufacture         = $resolvedMonitor.WeekOfManufacture
                    YearOfManufacture         = $resolvedMonitor.YearOfManufacture
                    CurrentInputSource        = if ($null -ne $currentInputSourceDescriptor) { [string]$currentInputSourceDescriptor.FriendlyName } else { $null }
                    CurrentInputSourceCode    = if ($null -ne $currentInputSourceDescriptor) { [string]$currentInputSourceDescriptor.Code } else { $null }
                    AvailableInputSources     = @($availableInputSourceDescriptors | Where-Object { $_.IsKnown } | ForEach-Object { [string]$_.FriendlyName } | Select-Object -Unique)
                    AvailableInputSourceCodes = @($availableInputSourceDescriptors | Where-Object { $_.IsKnown } | ForEach-Object { [string]$_.Code } | Select-Object -Unique)
                    AvailableInputSourceNames = @($availableInputSourceDescriptors | Where-Object { $_.IsKnown -and -not [string]::IsNullOrWhiteSpace([string]$_.LegacyName) } | ForEach-Object { [string]$_.LegacyName } | Select-Object -Unique)
                }
            }
        )
    }

    function Get-NormalizedMonitorSelections {
        [CmdletBinding()]
        param()

        $rawSelections = @($Monitors | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $normalizedSelections = [System.Collections.Generic.List[string]]::new()

        foreach ($rawSelection in $rawSelections) {
            $parts = @(
                [regex]::Split([string]$rawSelection, '\s*,\s*(?=\[\d+\]\s+)') |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )

            foreach ($part in $parts) {
                [void]$normalizedSelections.Add($part.Trim())
            }
        }

        return @($normalizedSelections | Select-Object -Unique)
    }

    <#
.SYNOPSIS
Gets selected monitor IDs detected in current winddcutil output.

.DESCRIPTION
Runs the detect subcommand and returns IDs for the selected resolved monitor
selections that are currently present in the detect output.

.OUTPUTS
System.String[]

.EXAMPLE
Get-MonitorIds

Returns selected display IDs as strings.
#>
    function Get-MonitorIds {
        [CmdletBinding()]
        param(
            [object[]]$DetectedMonitors = $null
        )

        if ($null -eq $DetectedMonitors) {
            $DetectedMonitors = @(Get-ResolvedDetectedMonitors -IgnoreExitCode)
        }

        $requestedMonitorSelections = @(Get-NormalizedMonitorSelections)

        if ($requestedMonitorSelections.Count -eq 0 -and
            $null -ne $script:DefaultMonitorSerialNumbers -and
            $script:DefaultMonitorSerialNumbers.Count -gt 0) {
            return @(
                $DetectedMonitors |
                Where-Object { $script:DefaultMonitorSerialNumbers -contains ([string]$_.SerialNumber) } |
                Sort-Object -Property @{ Expression = { [array]::IndexOf($script:DefaultMonitorSerialNumbers, [string]$_.SerialNumber) }; Descending = $false } |
                ForEach-Object { [string]$_.Id } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            )
        }

        $requestedMonitorIds = @(
            $requestedMonitorSelections |
            ForEach-Object {
                if ($_ -match '^\[(?<Id>\d+)\]\s+') {
                    [string]$Matches.Id
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        )

        return @(
            @($DetectedMonitors) |
            Where-Object {
                if ($requestedMonitorIds.Count -gt 0) {
                    return $requestedMonitorIds -contains ([string]$_.Id)
                }

                $resolvedDescription = [string]$_.ResolvedDescription
                if ([string]::IsNullOrWhiteSpace($resolvedDescription)) {
                    $resolvedDescription = [string]$_.Description
                }

                $requestedMonitorSelections -contains $resolvedDescription
            } |
            ForEach-Object { [string]$_.Id } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        )
    }

    <#
.SYNOPSIS
Retries monitor detection until selected monitors are found.

.DESCRIPTION
Calls Get-MonitorIds repeatedly using the configured retry count and delay.
Throws if one or more selected monitors are not detected after all attempts.

.OUTPUTS
System.String[]

.EXAMPLE
Get-MonitorIdsWithRetry

Returns selected display IDs after one or more detection attempts.
#>
    function Get-MonitorIdsWithRetry {
        [CmdletBinding()]
        param()

        $requestedMonitorSelections = @(Get-NormalizedMonitorSelections)
        $usingSerialDefaults = (
            $requestedMonitorSelections.Count -eq 0 -and
            $null -ne $script:DefaultMonitorSerialNumbers -and
            $script:DefaultMonitorSerialNumbers.Count -gt 0
        )

        $requestedMonitorIds = @(
            $requestedMonitorSelections |
            ForEach-Object {
                if ($_ -match '^\[(?<Id>\d+)\]\s+') {
                    [string]$Matches.Id
                }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        )

        for ($attempt = 1; $attempt -le $DetectRetryCount; $attempt++) {
            $detectedMonitors = @(Get-ResolvedDetectedMonitors -IgnoreExitCode)
            $monitorIds = @(Get-MonitorIds -DetectedMonitors $detectedMonitors)

            # Wrap in @() so that an empty result from the winning branch is preserved as
            # an empty array rather than unwrapped to $null by the pipeline, which would
            # make $missingMonitors.Count throw under Set-StrictMode -Version Latest.
            $missingMonitors = @(if ($usingSerialDefaults) {
                    $detectedSerials = @(
                        $detectedMonitors |
                        Where-Object { $monitorIds -contains ([string]$_.Id) } |
                        ForEach-Object { [string]$_.SerialNumber }
                    )
                    $script:DefaultMonitorSerialNumbers | Where-Object { $detectedSerials -notcontains $_ }
                }
                elseif ($requestedMonitorIds.Count -gt 0) {
                    $requestedMonitorIds | Where-Object { $monitorIds -notcontains $_ }
                }
                else {
                    $detectedResolvedDescriptions = @(
                        $detectedMonitors |
                        ForEach-Object {
                            if (-not [string]::IsNullOrWhiteSpace([string]$_.ResolvedDescription)) {
                                [string]$_.ResolvedDescription
                            }
                            else {
                                [string]$_.Description
                            }
                        } |
                        Select-Object -Unique
                    )
                    $requestedMonitorSelections | Where-Object { $detectedResolvedDescriptions -notcontains $_ }
                })

            if ($missingMonitors.Count -eq 0) {
                Write-Log -Message "Detected $($monitorIds.Count) requested monitor(s) on attempt ${attempt}: $($monitorIds -join ', ')"
                return $monitorIds
            }

            if ($attempt -lt $DetectRetryCount) {
                Write-Log -Message "Monitor(s) $($missingMonitors -join ', ') not yet detected on attempt ${attempt}. Retrying in $DetectRetryDelaySeconds second(s)." -Level WARN
                Start-Sleep -Seconds $DetectRetryDelaySeconds
            }
        }

        $selectionDescription = if ($usingSerialDefaults) {
            $script:DefaultMonitorSerialNumbers -join ', '
        }
        else {
            $requestedMonitorSelections -join ', '
        }

        throw "Requested monitor(s) $selectionDescription were not all detected after $DetectRetryCount attempts."
    }

    function Assert-MonitorInputSourceSupported {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string[]]$MonitorIds,

            [Parameter(Mandatory)]
            [string]$RequestedInputSource
        )

        $unsupportedMonitorIds = @(
            foreach ($monitorId in $MonitorIds) {
                $supportedInputSources = @(Get-MonitorSupportedInputSources -MonitorId $monitorId)
                if ($supportedInputSources.Count -eq 0 -or $supportedInputSources -notcontains $RequestedInputSource) {
                    $monitorId
                }
            }
        )

        if ($unsupportedMonitorIds.Count -gt 0) {
            throw "InputSource '$RequestedInputSource' is not supported by monitor(s): $($unsupportedMonitorIds -join ', ')"
        }
    }

    <#
.SYNOPSIS
Gets the current power-state value for a monitor.

.DESCRIPTION
Reads the configured power-mode VCP code from a monitor and returns the current
value as an integer when winddcutil reports it in a supported format.

.PARAMETER MonitorId
Display identifier to query.

.OUTPUTS
System.Int32
#>
    function Get-MonitorPowerStateValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId
        )

        $result = Invoke-Winddcutil -GetVcp -Display $MonitorId -FeatureCode $PowerModeCode
        $rawValue = [string]$result.Value

        if ($rawValue -match 'VCP\s+0x[0-9a-fA-F]+\s+(?<Value>0x[0-9a-fA-F]+|\d+)') {
            $parsedValue = $Matches.Value
            if ($parsedValue -match '^0x') {
                return [Convert]::ToInt32($parsedValue.Substring(2), 16)
            }

            return [int]$parsedValue
        }

        throw "Unable to parse power state from winddcutil output for monitor $MonitorId. Output: $rawValue"
    }

    function Convert-VcpValueToInt {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Value
        )

        $trimmedValue = $Value.Trim()
        if ($trimmedValue -match '^0x(?<Hex>[0-9a-fA-F]+)$') {
            return [Convert]::ToInt32($Matches.Hex, 16)
        }

        return [int]$trimmedValue
    }

    function Get-MonitorVcpValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId,

            [Parameter(Mandatory)]
            [string]$Code
        )

        $result = Invoke-Winddcutil -GetVcp -Display $MonitorId -FeatureCode $Code
        $rawValue = [string]$result.Value

        if ($rawValue -match 'VCP\s+0x[0-9a-fA-F]+\s+(?<Value>0x[0-9a-fA-F]+|\d+)') {
            return (Convert-VcpValueToInt -Value $Matches.Value)
        }

        throw "Unable to parse VCP value from winddcutil output for monitor $MonitorId and code $Code. Output: $rawValue"
    }

    function Convert-VcpValueForComparison {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Code,

            [Parameter(Mandatory)]
            [int]$Value
        )

        if ($Code -eq '0x60') {
            return ($Value -band 0xFF)
        }

        return $Value
    }

    <#
.SYNOPSIS
Sets a VCP value on a monitor with retry logic.

.DESCRIPTION
Invokes winddcutil setvcp for the specified monitor and VCP code. Retries failed
operations using the configured retry settings and throws if all attempts fail.

.PARAMETER MonitorId
Display identifier to target.

.PARAMETER Code
VCP feature code to change.

.PARAMETER Value
New VCP value to apply.

.PARAMETER ActionName
Friendly action name used in log messages.

.EXAMPLE
Set-MonitorVcpValue -MonitorId 3 -Code 0x60 -Value 0x11 -ActionName 'Switch input to HDMI1'

Attempts to switch display 3 to the requested input source.
#>
    function Set-MonitorVcpValue {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$MonitorId,

            [Parameter(Mandatory)]
            [string]$Code,

            [Parameter(Mandatory)]
            [string]$Value,

            [Parameter(Mandatory)]
            [string]$ActionName
        )

        for ($attempt = 1; $attempt -le $SetVcpRetryCount; $attempt++) {
            try {
                Invoke-Winddcutil -SetVcp -Display $MonitorId -FeatureCode $Code -NewValue $Value | Out-Null
                $expectedValue = Convert-VcpValueForComparison -Code $Code -Value (Convert-VcpValueToInt -Value $Value)
                $currentValue = Convert-VcpValueForComparison -Code $Code -Value (Get-MonitorVcpValue -MonitorId $MonitorId -Code $Code)

                if ($currentValue -ne $expectedValue) {
                    throw "Monitor $MonitorId reported VCP code $Code value $currentValue after setting $expectedValue."
                }

                Write-Log -Message "$ActionName succeeded for monitor $MonitorId on attempt ${attempt}."
                return
            }
            catch {
                if ($attempt -eq $SetVcpRetryCount) {
                    throw
                }

                Write-Log -Message "$ActionName failed for monitor $MonitorId on attempt ${attempt}. Retrying in $SetVcpRetryDelaySeconds second(s). Error: $($_.Exception.Message)" -Level WARN
                Start-Sleep -Seconds $SetVcpRetryDelaySeconds
            }
        }
    }

    <#
.SYNOPSIS
Ensures that the configured log directory exists and is writable.

.DESCRIPTION
Creates the log directory if needed, writes a temporary probe file, and removes
it again. Throws if the current user cannot create or delete files in the path.

.PARAMETER Path
Directory path to validate.

.EXAMPLE
Assert-LogDirectoryWritable -Path 'C:\Windows\Logs\Set-ExternalMonitorState'

Verifies that the log directory can be used by the current user.
#>
    function Assert-LogDirectoryWritable {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        try {
            Initialize-LogDirectory -Path $Path
        }
        catch {
            if ($Path -ne $script:FallbackLogDirectory) {
                Set-LogLocation -Path $script:FallbackLogDirectory
                Initialize-LogDirectory -Path $script:LogDirectory
                $script:LogFallbackActivated = $true
                Write-Host "[WARN] Falling back to log directory '$script:LogDirectory' because '$Path' is not writable."
                return
            }


            try {
                Write-Host $entry
            }
            catch {
            }
            throw
        }
    }

    <#
.SYNOPSIS
Clears the log file when it is too large or too old.

.DESCRIPTION
Checks the current log file size and the timestamp of the oldest log entry.
Truncates the file when it exceeds the configured size limit or when the first
timestamped entry is older than the configured retention window.

.OUTPUTS
System.String[]

.EXAMPLE
Reset-LogIfNeeded

Clears the current log file when it exceeds rotation thresholds and returns the
reasons that triggered the reset.
#>
    function Reset-LogIfNeeded {
        [CmdletBinding()]
        param()

        if (-not (Test-Path -LiteralPath $LogPath -PathType Leaf)) {
            return @()
        }

        $reasons = [System.Collections.Generic.List[string]]::new()
        $logFile = Get-Item -LiteralPath $LogPath -ErrorAction Stop

        if ($logFile.Length -gt $MaxLogSizeBytes) {
            [void]$reasons.Add("size $($logFile.Length) bytes exceeds $MaxLogSizeBytes bytes")
        }

        $firstEntry = Get-Content -LiteralPath $LogPath -TotalCount 1 -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($firstEntry)) {
            $timestampMatch = [regex]::Match($firstEntry, '^\[(?<Timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]')
            if ($timestampMatch.Success) {
                $oldestEntryTimestamp = [datetime]::ParseExact(
                    $timestampMatch.Groups['Timestamp'].Value,
                    'yyyy-MM-dd HH:mm:ss',
                    [System.Globalization.CultureInfo]::InvariantCulture
                )

                if ($oldestEntryTimestamp -lt (Get-Date).Subtract($MaxLogAge)) {
                    [void]$reasons.Add("oldest entry $($oldestEntryTimestamp.ToString('yyyy-MM-dd HH:mm:ss')) is older than $([int]$MaxLogAge.TotalDays) days")
                }
            }
        }

        if ($reasons.Count -eq 0) {
            return @()
        }

        Clear-Content -LiteralPath $LogPath -ErrorAction Stop
        return $reasons.ToArray()
    }

    try {
        $invocationContext = $null
        $ddcUtilPathWasExplicitlyProvided = $PSBoundParameters.ContainsKey('DdcUtilPath')

        $DdcUtilPath = Resolve-WinddcutilPath -RequestedPath $DdcUtilPath -WasExplicitlyProvided:$ddcUtilPathWasExplicitlyProvided

        if ($DetectOnly) {
            $matchingMonitors = @(Get-DetectOnlyMonitors -IgnoreExitCode)

            if ($Json) {
                return ($matchingMonitors | ConvertTo-Json -Depth 6)
            }
            else {
                return $matchingMonitors
            }
        }

        Assert-LogDirectoryWritable -Path $LogDirectory

        $logResetReasons = @(Reset-LogIfNeeded)
        if ($logResetReasons.Count -gt 0) {
            Write-Log -Message "Cleared log file '$LogPath' because $($logResetReasons -join '; ')."
        }

        $invocationContext = Get-InvocationContext
        Write-InvocationContextLog -InvocationContext $invocationContext -Stage 'Startup'

        $shouldSwitchInput = ($PowerAction -eq 'On' -and $InputSourceSpecified)
        if ($shouldSwitchInput) {
            Write-Log -Message "Starting external monitor state update with power action $PowerAction and input source $SelectedInputSource."
        }
        else {
            Write-Log -Message "Starting external monitor state update with power action $PowerAction."
        }

        $monitorIds = Get-MonitorIdsWithRetry
        if ($shouldSwitchInput) {
            Assert-MonitorInputSourceSupported -MonitorIds $monitorIds -RequestedInputSource $SelectedInputSource
        }

        $poweredOnAnyMonitor = $false
        foreach ($monitorId in $monitorIds) {
            if ($PowerAction -eq 'On') {
                try {
                    $currentPowerState = Get-MonitorPowerStateValue -MonitorId $monitorId
                }
                catch {
                    Write-Log -Message "Skipping power-on for monitor $monitorId because its current power state could not be determined. Error: $($_.Exception.Message)" -Level WARN
                    continue
                }

                if ($currentPowerState -ne [int]$PowerActionValueMap.Off) {
                    Write-Log -Message "Skipping power-on for monitor $monitorId because its current power state is $currentPowerState instead of $($PowerActionValueMap.Off)."
                    continue
                }
            }

            Set-MonitorVcpValue -MonitorId $monitorId -Code $PowerModeCode -Value $SelectedPowerActionValue -ActionName "Power $PowerAction"

            if ($PowerAction -eq 'On') {
                $poweredOnAnyMonitor = $true
            }
        }

        if ($shouldSwitchInput) {
            if ($poweredOnAnyMonitor) {
                Write-Log -Message "Waiting $PostPowerOnDelaySeconds second(s) after powering monitors on."
                Start-Sleep -Seconds $PostPowerOnDelaySeconds
            }
            else {
                Write-Log -Message 'Skipping post-power-on wait because no monitors required power-on.'
            }

            foreach ($monitorId in $monitorIds) {
                Set-MonitorVcpValue -MonitorId $monitorId -Code $InputSourceCode -Value $SelectedInputSourceValue -ActionName "Switch input to $SelectedInputSource"
            }
        }
        else {
            Write-Log -Message "Skipping input switch because power action $PowerAction does not require it."
        }

        Write-Log -Message 'External monitor state update completed successfully.'
        exit 0
    }
    catch {
        Write-Log -Message "Monitor state update failed. $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}