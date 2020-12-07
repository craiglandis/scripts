param(
  [switch]$transcript,
  [switch]$runChildScript,
  [int]$traceLevel,
  [int]$exitCode
)

if ($transcript)
{
    Start-Transcript
}

if ($traceLevel)
{
    Set-PSDebug -Trace $traceLevel
}

$scriptName = $MyInvocation.MyCommand.Name
"`$NestedPromptLevel: $NestedPromptLevel"
"'no cmdlet' output from $scriptName"
Write-Debug "Write-Debug output from $scriptName"
Write-Error "Write-Error output from $scriptName"
Write-Host "Write-Host output from $scriptName"
Write-Information "Write-Information output from $scriptName"
Write-Output "Write-Output output from $scriptName"
Write-Progress "Write-Progress output from $scriptName"
Write-Warning "Write-Warning output from $scriptName"
Write-Verbose "Write-Verbose output from $scriptName"

if ($runChildScript)
{
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/craiglandis/scripts/master/Test-CustomScriptExtensionChild.ps1'))
}

if ($exitCode)
{
    exit $exitCode
}

if ($transcript)
{
    Stop-Transcript
}
