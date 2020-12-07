<#
.SYNOPSIS
    Get-ScriptBlockLogging.ps1
.DESCRIPTION
    Exports Microsoft-Windows-PowerShell/Operational Event ID 4104 events with local system account UserId S-1-5-18 to CSV file.
    The ScriptBlock column in the CSV file will have the full contents of the script.

    To download the script into the VM from a PowerShell prompt in the VM, run: 
    (New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/craiglandis/scripts/master/Get-ScriptBlockLogging.ps1', "$PWD\Get-ScriptBlockLogging.ps1")

    Or you can view it on Github at https://github.com/craiglandis/scripts/blob/master/Get-ScriptBlockLogging.ps1, click on "Raw", then copy/paste the script into a new text file and save it as Get-ScriptBlockLogging.ps1
  
    Scripts run through RunCommand or CustomScriptExtension run under the security context of the local system account, which always has security identifier (SID) S-1-5-18
    https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/security-identifiers-in-windows

    For more information on PowerShell Script Block Logging see the following article -
    https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/wmf/whats-new/script-logging?view=powershell-5.1
.PARAMETER Path
    Optionally you can provide the full file path including CSV file name.
    If -Path is not specified, .\Get-ScriptBlockLogging.csv is created in the directory where you ran Get-ScriptBlockLogging.ps1.
.EXAMPLE
    PS C:\> Get-ScriptBlockLogging.ps1
#>

param(
    [string]$path = "$PWD\$($MyInvocation.MyCommand.Name.Replace('.ps1', '.csv'))"
)

$localSystem = 'S-1-5-18'
$filterHashtable = @{ProviderName = "Microsoft-Windows-PowerShell"; Id = 4104; UserId = $localSystem}
Write-Output "Querying Microsoft-Windows-PowerShell/Operational for Event ID 4104 events with local system UserId $localSystem"
$scriptBlockLoggingEvents = Get-WinEvent -FilterHashtable $filterHashtable -ErrorAction SilentlyContinue
if ($scriptBlockLoggingEvents)
{
    Write-Output "$($scriptBlockLoggingEvents.Count) events found for $(($scriptBlockLoggingEvents.Properties[3].Value | Sort-Object -Unique).Count) script execution(s)"
    $scriptBlockLoggingEvents | ForEach-Object {
        $scriptBlockLoggingEvent = $_
        $scriptBlockId = $scriptBlockLoggingEvent.Properties[3].Value
        $scriptEvents = $scriptBlockLoggingEvents | Where-Object {$_.Properties[3].Value -eq $scriptBlockId} | Sort-Object {$_.Properties[0].Value}
        $scriptBlock = -join ($scriptEvents | ForEach-Object {$_.Properties[2].Value})
        $scriptEvents | ForEach-Object {
            $_ | add-member -MemberType 'NoteProperty' -Name 'ScriptBlockId' -Value $scriptBlockId -Force
            $_ | add-member -MemberType 'NoteProperty' -Name 'ScriptBlock' -Value $scriptBlock -Force
        }
    }

    Write-Output "Exporting events to file: $path"
    $scriptBlockLoggingEvents | Sort-Object ScriptBlock -unique | Select-Object TimeCreated, Id, LevelDisplayName, LogName, MachineName, ProviderName, UserId, ScriptBlockId, ScriptBlock | Export-Csv -Path $path -NoTypeInformation

    if (Test-Path -Path $path)
    {
        Write-Output "Events exported to file: $path"
    }
}
else
{
    Write-Output "No events found"
}
