param(
    $path,
    $outputPath = $PWD
)

$updates =  New-Object System.Collections.ArrayList
$sessions = [xml](get-content $path)
$sessions = $sessions.sessions.session
$sessions | foreach {
    $session = $_
    $tasks = $session.Tasks
    $actions = $session.Actions
    $tasks | foreach {
        $task = $_
        $packages = $task.phase.package
        $packages | foreach {
            $package = $_
            if ($package.targetState -eq 'Installed')
            {
                $update = [PSCustomObject]@{
                    name = $package.name
                    client = $session.client
                    targetState = $package.targetState
                    lastSuccessfulState = $session.lastSuccessfulState
                    status = $session.status
                    rebootRequired = $actions.phase.rebootRequired
                    Queued = $session.Queued
                    Started = $session.Started
                    Complete = $session.Complete
                    Resolved = $actions.phase.Resolved
                    Staged = $actions.phase.Staged
                    Installed = $actions.phase.Installed
                    ShutdownStart = $actions.phase.ShutdownStart
                    ShutdownFinish = $actions.phase.ShutdownFinish
                    Startup        = $actions.phase.Startup
                    StartupFinish = $actions.phase.StartupFinish
                    'session.options' = $session.options
                    'package.options' = $package.Options
                    'session.id' = $session.id
                    'package.id' = $package.id
                    currentPhase = $session.currentPhase
                    pendingFollower = $session.pendingFollower
                    retry = $session.retry
                    version = $session.version
                }

                [void]$updates.Add($update)
            }
        }
    }
}

$outputFilePath = "$outputPath\updates$((get-date).ticks).csv"
$updates = $updates | sort name -unique
$updates = $updates | where {$_.name.startswith('KB')}
$updates = $updates | sort Queued -descending
$updates | export-csv -Path $outputFilePath -NoTypeInformation
write-host "output: $outputFilePath"
invoke-item $outputFilePath
