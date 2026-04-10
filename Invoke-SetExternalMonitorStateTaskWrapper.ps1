<#
.SYNOPSIS
Wraps Set-ExternalMonitorState.ps1 for Task Scheduler execution.

.DESCRIPTION
Runs Set-ExternalMonitorState.ps1 in a child PowerShell process so the wrapper
can schedule a detached helper to record the final Task Scheduler completion
history entry after the task action exits.

.EXAMPLE
.\Invoke-SetExternalMonitorStateTaskWrapper.ps1 -PowerAction On -InputSource HDMI1

Runs the monitor-state script and records Task Scheduler history details in the
shared log file when the wrapper is launched by Task Scheduler.
#>

[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')]
    [string]$ScriptPath = 'C:\programs\PowerShell\Scripts\Local\Utility\Set-ExternalMonitorState.ps1',

    [Parameter(ParameterSetName = 'Run')]
    [string]$LogDirectory = 'C:\Windows\Logs\Set-ExternalMonitorState',

    [Parameter(ParameterSetName = 'Run')]
    [string]$LogFileName = 'Set-ExternalMonitorState.log',

    [Parameter(ParameterSetName = 'Run', ValueFromRemainingArguments)]
    [object[]]$ScriptArguments,

    [Parameter(Mandatory, ParameterSetName = 'Finalize')]
    [switch]$FinalizeTaskHistory,

    [Parameter(Mandatory, ParameterSetName = 'Finalize')]
    [string]$TaskName,

    [Parameter(ParameterSetName = 'Finalize')]
    [string]$TaskInstanceId,

    [Parameter(Mandatory, ParameterSetName = 'Finalize')]
    [string]$FinalizerLogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$TaskSchedulerHistoryLogName = 'Microsoft-Windows-TaskScheduler/Operational'
$FallbackLogDirectory = Join-Path -Path ([Environment]::GetFolderPath('LocalApplicationData')) -ChildPath 'Set-ExternalMonitorState\Logs'

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

function Initialize-Directory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-EffectiveLogPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [string]$FileName
    )

    try {
        Initialize-Directory -Path $Directory
        return (Join-Path -Path $Directory -ChildPath $FileName)
    }
    catch {
        Initialize-Directory -Path $FallbackLogDirectory
        return (Join-Path -Path $FallbackLogDirectory -ChildPath $FileName)
    }
}

function Write-WrapperLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogPath,

        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] [Wrapper] $Message"
    Add-Content -Path $LogPath -Value $entry -ErrorAction Stop
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

function Get-ScheduledTaskContextForCurrentProcess {
    [CmdletBinding()]
    param()

    $processAncestry = @(Get-ProcessAncestry)
    if ($processAncestry.Count -eq 0) {
        return $null
    }

    $processesById = @{}
    foreach ($processInfo in $processAncestry) {
        $processesById[[int]$processInfo.ProcessId] = $processInfo
    }

    $startWindow = (@($processAncestry | Select-Object -ExpandProperty StartTime | Sort-Object) | Select-Object -First 1).AddMinutes(-5)

    try {
        $launchEvents = @(
            Get-WinEvent -FilterHashtable @{
                LogName   = $TaskSchedulerHistoryLogName
                Id        = 129
                StartTime = $startWindow
            } -ErrorAction Stop
        )
    }
    catch {
        return $null
    }

    $matchingLaunchEvents = @(
        foreach ($launchEvent in $launchEvents) {
            $eventData = Get-TaskSchedulerEventData -EventRecord $launchEvent
            $processId = 0
            if (-not [int]::TryParse([string]$eventData.ProcessID, [ref]$processId)) {
                continue
            }

            if (-not $processesById.ContainsKey($processId)) {
                continue
            }

            $matchedProcess = $processesById[$processId]

            [pscustomobject]@{
                Event            = $launchEvent
                EventData        = $eventData
                MatchedProcess   = $matchedProcess
                IsCurrentProcess = ($matchedProcess.ProcessId -eq $PID)
                Delta            = [math]::Abs(($launchEvent.TimeCreated - $matchedProcess.StartTime).TotalSeconds)
            }
        }
    )

    $selected = @(
        $matchingLaunchEvents |
        Sort-Object -Property @{ Expression = 'IsCurrentProcess'; Descending = $true }, @{ Expression = 'Delta'; Descending = $false }, @{ Expression = { $_.Event.TimeCreated }; Descending = $true }
    ) | Select-Object -First 1

    if ($null -eq $selected) {
        return $null
    }

    $taskHistory = @(
        Get-WinEvent -FilterHashtable @{
            LogName   = $TaskSchedulerHistoryLogName
            Id        = 100, 200
            StartTime = $selected.Event.TimeCreated.AddMinutes(-2)
        } -ErrorAction SilentlyContinue
    )

    $taskName = [string]$selected.EventData.TaskName
    $actionStartEvent = @(
        foreach ($historyEvent in $taskHistory) {
            $historyData = Get-TaskSchedulerEventData -EventRecord $historyEvent
            if ([string]$historyData.TaskName -ne $taskName) {
                continue
            }

            if ($historyEvent.Id -eq 200 -and [string]$historyData.TaskInstanceId) {
                [pscustomobject]@{ Event = $historyEvent; EventData = $historyData }
            }
        }
    ) | Sort-Object -Property @{ Expression = { $_.Event.TimeCreated }; Descending = $true } | Select-Object -First 1

    $taskStartEvent = @(
        foreach ($historyEvent in $taskHistory) {
            $historyData = Get-TaskSchedulerEventData -EventRecord $historyEvent
            if ([string]$historyData.TaskName -ne $taskName) {
                continue
            }

            if ($historyEvent.Id -eq 100) {
                [pscustomobject]@{ Event = $historyEvent; EventData = $historyData }
            }
        }
    ) | Sort-Object -Property @{ Expression = { $_.Event.TimeCreated }; Descending = $true } | Select-Object -First 1

    $resolvedTaskStartTime = if ($null -ne $taskStartEvent) {
        [datetime]$taskStartEvent.Event.TimeCreated
    }
    else {
        [datetime]$selected.Event.TimeCreated
    }

    $triggerClassification = Get-ScheduledTaskTriggerClassification -TaskName $taskName -TaskStartTime $resolvedTaskStartTime -TaskLaunchTime ([datetime]$selected.Event.TimeCreated)

    return [pscustomobject]@{
        TaskName       = $taskName
        TaskInstanceId = if ($null -ne $actionStartEvent) { [string]$actionStartEvent.EventData.TaskInstanceId } elseif ($null -ne $taskStartEvent) { [string]$taskStartEvent.EventData.InstanceId } else { $null }
        TaskActionName = if ($null -ne $actionStartEvent) { [string]$actionStartEvent.EventData.ActionName } else { $null }
        TaskStartTime  = $resolvedTaskStartTime
        TaskLaunchTime = [datetime]$selected.Event.TimeCreated
        TaskTriggerSource = [string]$triggerClassification.TriggerSource
        TaskTriggerReason = [string]$triggerClassification.Reason
    }
}

function Get-ScheduledTaskDefinitionXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TaskName
    )

    $taskPath = '\'
    $shortTaskName = $TaskName.Trim()

    if ($shortTaskName -match '^(?<TaskPath>.*\\)(?<ShortTaskName>[^\\]+)$') {
        $taskPath = [string]$Matches.TaskPath
        $shortTaskName = [string]$Matches.ShortTaskName
    }

    try {
        $taskDefinition = Export-ScheduledTask -TaskPath $taskPath -TaskName $shortTaskName -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($taskDefinition)) {
            return $null
        }

        return [xml]$taskDefinition
    }
    catch {
        return $null
    }
}

function Get-ScheduledTaskTriggerClassification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TaskName,

        [datetime]$TaskStartTime,

        [datetime]$TaskLaunchTime
    )

    $taskDefinitionXml = Get-ScheduledTaskDefinitionXml -TaskName $TaskName
    $supportsLogonTrigger = $false
    $supportsUnlockTrigger = $false

    if ($null -ne $taskDefinitionXml) {
        $namespaceManager = [System.Xml.XmlNamespaceManager]::new($taskDefinitionXml.NameTable)
        [void]$namespaceManager.AddNamespace('task', 'http://schemas.microsoft.com/windows/2004/02/mit/task')

        $supportsLogonTrigger = ($null -ne $taskDefinitionXml.SelectSingleNode('/task:Task/task:Triggers/task:LogonTrigger', $namespaceManager))
        $supportsUnlockTrigger = ($null -ne $taskDefinitionXml.SelectSingleNode('/task:Task/task:Triggers/task:SessionStateChangeTrigger[task:StateChange = "SessionUnlock"]', $namespaceManager))
    }

    $referenceTime = if ($TaskStartTime -is [datetime]) {
        $TaskStartTime
    }
    elseif ($TaskLaunchTime -is [datetime]) {
        $TaskLaunchTime
    }
    else {
        $null
    }

    $recentLogonEvent = $null
    $logonDetectionWindow = [TimeSpan]::FromSeconds(30)

    if ($supportsLogonTrigger -and $referenceTime -is [datetime]) {
        try {
            $recentLogonEvent = @(
                Get-WinEvent -FilterHashtable @{
                    LogName      = 'System'
                    ProviderName = 'Microsoft-Windows-Winlogon'
                    Id           = 7001
                    StartTime    = $referenceTime.Subtract($logonDetectionWindow)
                    EndTime      = $referenceTime.AddSeconds(5)
                } -ErrorAction Stop |
                Where-Object { $_.Message -like 'User Log-on Notification*' } |
                Sort-Object -Property TimeCreated -Descending
            ) | Select-Object -First 1
        }
        catch {
            $recentLogonEvent = $null
        }
    }

    if ($null -ne $recentLogonEvent) {
        return [pscustomobject]@{
            TriggerSource = 'Logon'
            Reason        = "Matched a Winlogon user logon notification at $($recentLogonEvent.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')) within $([int]$logonDetectionWindow.TotalSeconds) seconds of task start."
        }
    }

    if ($supportsUnlockTrigger -and -not $supportsLogonTrigger) {
        return [pscustomobject]@{
            TriggerSource = 'Unlock'
            Reason        = 'Task definition exposes only a session unlock trigger.'
        }
    }

    if ($supportsLogonTrigger -and -not $supportsUnlockTrigger) {
        return [pscustomobject]@{
            TriggerSource = 'Logon'
            Reason        = 'Task definition exposes only a logon trigger.'
        }
    }

    if ($supportsUnlockTrigger -and $supportsLogonTrigger) {
        return [pscustomobject]@{
            TriggerSource = 'Unlock'
            Reason        = "Task definition includes both logon and unlock triggers, and no Winlogon user logon notification was found within $([int]$logonDetectionWindow.TotalSeconds) seconds of task start."
        }
    }

    return [pscustomobject]@{
        TriggerSource = 'Unknown'
        Reason        = 'Task trigger type could not be resolved from the task definition or available event history.'
    }
}

function Get-ScheduledTaskCompletionEvent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TaskName,

        [string]$TaskInstanceId,

        [datetime]$StartTime = (Get-Date).AddMinutes(-10)
    )

    $historyEvents = @(
        Get-WinEvent -FilterHashtable @{
            LogName   = $TaskSchedulerHistoryLogName
            Id        = 102, 201
            StartTime = $StartTime
        } -ErrorAction SilentlyContinue
    )

    $matchingCompletionEvents = @(
        foreach ($historyEvent in $historyEvents) {
            $historyData = Get-TaskSchedulerEventData -EventRecord $historyEvent
            if ([string]$historyData.TaskName -ne $TaskName) {
                continue
            }

            $instanceMatches = $true
            if (-not [string]::IsNullOrWhiteSpace($TaskInstanceId)) {
                $instanceMatches = (
                    ([string]$historyData.TaskInstanceId -eq $TaskInstanceId) -or
                    ([string]$historyData.InstanceId -eq $TaskInstanceId)
                )
            }

            if (-not $instanceMatches) {
                continue
            }

            [pscustomobject]@{
                Event     = $historyEvent
                EventData = $historyData
            }
        }
    )

    return @(
        $matchingCompletionEvents |
        Sort-Object -Property @{ Expression = { $_.Event.TimeCreated }; Descending = $true }
    ) | Select-Object -First 1
}

function Start-TaskHistoryFinalizer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TaskName,

        [string]$TaskInstanceId,

        [Parameter(Mandatory)]
        [string]$LogPath
    )

    $hostPath = (Get-Process -Id $PID -ErrorAction Stop).Path
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $hostPath
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    [void]$startInfo.ArgumentList.Add('-NoProfile')
    [void]$startInfo.ArgumentList.Add('-ExecutionPolicy')
    [void]$startInfo.ArgumentList.Add('Bypass')
    [void]$startInfo.ArgumentList.Add('-File')
    [void]$startInfo.ArgumentList.Add($PSCommandPath)
    [void]$startInfo.ArgumentList.Add('-FinalizeTaskHistory')
    [void]$startInfo.ArgumentList.Add('-TaskName')
    [void]$startInfo.ArgumentList.Add($TaskName)

    if (-not [string]::IsNullOrWhiteSpace($TaskInstanceId)) {
        [void]$startInfo.ArgumentList.Add('-TaskInstanceId')
        [void]$startInfo.ArgumentList.Add($TaskInstanceId)
    }

    [void]$startInfo.ArgumentList.Add('-FinalizerLogPath')
    [void]$startInfo.ArgumentList.Add($LogPath)

    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    [void]$process.Start()
}

if ($FinalizeTaskHistory) {
    $deadline = (Get-Date).AddMinutes(2)
    $completionEvent = $null

    while ((Get-Date) -lt $deadline) {
        $completionEvent = Get-ScheduledTaskCompletionEvent -TaskName $TaskName -TaskInstanceId $TaskInstanceId
        if ($null -ne $completionEvent) {
            break
        }

        Start-Sleep -Milliseconds 500
    }

    if ($null -ne $completionEvent) {
        $endTime = $completionEvent.Event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
        $resultCode = if ($completionEvent.Event.Id -eq 201) {
            [string]$completionEvent.EventData.ResultCode
        }
        else {
            'unavailable'
        }

        Write-WrapperLog -LogPath $FinalizerLogPath -Message "Scheduled task history completion recorded for '$TaskName'. Instance: $TaskInstanceId. End time: $endTime. EventId: $($completionEvent.Event.Id). Result code: $resultCode."
    }
    else {
        Write-WrapperLog -LogPath $FinalizerLogPath -Message "Scheduled task history completion entry was not found within the polling window for '$TaskName'. Instance: $TaskInstanceId." -Level WARN
    }

    exit 0
}

$resolvedLogPath = Get-EffectiveLogPath -Directory $LogDirectory -FileName $LogFileName

try {
    if (-not (Test-Path -LiteralPath $ScriptPath -PathType Leaf)) {
        throw "Target script not found at '$ScriptPath'."
    }

    $taskContext = Get-ScheduledTaskContextForCurrentProcess
    $hostPath = (Get-Process -Id $PID -ErrorAction Stop).Path
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $hostPath
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $false
    $startInfo.RedirectStandardError = $false
    $startInfo.CreateNoWindow = $true
    [void]$startInfo.ArgumentList.Add('-NoProfile')
    [void]$startInfo.ArgumentList.Add('-ExecutionPolicy')
    [void]$startInfo.ArgumentList.Add('Bypass')
    [void]$startInfo.ArgumentList.Add('-File')
    [void]$startInfo.ArgumentList.Add($ScriptPath)
    [void]$startInfo.ArgumentList.Add('-LogDirectory')
    [void]$startInfo.ArgumentList.Add((Split-Path -Path $resolvedLogPath -Parent))
    [void]$startInfo.ArgumentList.Add('-LogFileName')
    [void]$startInfo.ArgumentList.Add((Split-Path -Path $resolvedLogPath -Leaf))

    $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_INVOCATION_MODE'] = if ($null -ne $taskContext) { 'ScheduledTask' } else { 'Interactive' }
    $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_INVOCATION_REASON'] = if ($null -ne $taskContext) {
        'Invocation context was forwarded by Invoke-SetExternalMonitorStateTaskWrapper.ps1.'
    }
    else {
        'Invocation context was forwarded by Invoke-SetExternalMonitorStateTaskWrapper.ps1.'
    }
    $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_WRAPPER_PROCESS_ID'] = [string]$PID
    $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_WRAPPER_PROCESS_NAME'] = [string](Get-Process -Id $PID -ErrorAction Stop).ProcessName

    if ($null -ne $taskContext) {
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_NAME'] = [string]$taskContext.TaskName
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_INSTANCE_ID'] = [string]$taskContext.TaskInstanceId
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_ACTION_NAME'] = [string]$taskContext.TaskActionName
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_START_TIME'] = $taskContext.TaskStartTime.ToString('o')
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_LAUNCH_TIME'] = $taskContext.TaskLaunchTime.ToString('o')
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_TRIGGER_SOURCE'] = [string]$taskContext.TaskTriggerSource
        $startInfo.EnvironmentVariables['SET_EXTERNAL_MONITOR_TASK_TRIGGER_REASON'] = [string]$taskContext.TaskTriggerReason
    }

    foreach ($argument in @($ScriptArguments)) {
        if ($null -eq $argument) {
            continue
        }

        [void]$startInfo.ArgumentList.Add([string]$argument)
    }

    $childProcess = [System.Diagnostics.Process]::new()
    $childProcess.StartInfo = $startInfo
    [void]$childProcess.Start()
    $childProcess.WaitForExit()
    $exitCode = $childProcess.ExitCode

    Write-WrapperLog -LogPath $resolvedLogPath -Message "Child monitor-state script exited with code $exitCode."

    if ($null -ne $taskContext) {
        Start-TaskHistoryFinalizer -TaskName $taskContext.TaskName -TaskInstanceId $taskContext.TaskInstanceId -LogPath $resolvedLogPath
    }

    exit $exitCode
}
catch {
    Write-WrapperLog -LogPath $resolvedLogPath -Message "Wrapper failed. $($_.Exception.Message)" -Level ERROR
    exit 1
}