Set-PSDebug -Trace 2
Start-Transcript
$scriptName = $MyInvocation.MyCommand.Name
"'no cmdlet' output from $scriptName"
Write-Debug "Write-Debug output from $scriptName"
Write-Error "Write-Error output from $scriptName"
Write-Host "Write-Host output from $scriptName"
Write-Information "Write-Information output from $scriptName"
Write-Output "Write-Output output from $scriptName"
Write-Progress "Write-Progress output from $scriptName"
Write-Warning "Write-Warning output from $scriptName"
Write-Verbose "Write-Verbose output from $scriptName"
Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/craiglandis/scripts/master/Test-CustomScriptExtensionChild.ps1'))
exit 1
Stop-Transcript
