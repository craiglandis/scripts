# Set-AzVMCustomScriptExtension -Location 'westus2' -ResourceGroupName 'rg1' -VMName 'vm1' -Name 'CustomScriptExtension' -FileUri 'https://raw.githubusercontent.com/craiglandis/scripts/master/Test-CustomScriptExtension.ps1' -Run 'Test-CustomScriptExtension.ps1'
Set-PSDebug -Trace 2
Start-Transcript

<#
$enableScriptBlockLoggingValue = Get-ItemPropertyValue -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -ErrorAction SilentlyContinue
if ($enableScriptBlockLoggingValue)
{
    Write-Output "EnableScriptBlockLoggingValue: $enableScriptBlockLoggingValue"
}
else
{
    Write-Output "EnableScriptBlockLoggingValue registry entry does not exist"
    $scriptBlockLoggingKey = Get-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -ErrorAction SilentlyContinue
    if ((Test-Path -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging') -eq $false)
    {
        # Create ScriptBlockLogging key if it does not exist
        New-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Force
    }
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -PropertyType 'DWord' -Value 1 -Force
    New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockInvocationLogging" -PropertyType 'DWord' -Value 1 -Force
}
#>

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
Set-PSDebug -Trace 0

<#
New-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Force
New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -PropertyType 'DWord' -Value 1 -Force
New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -PropertyType 'DWord' -Value 0 -Force
Get-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' #-ErrorAction SilentlyContinue
Get-ItemPropertyValue -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging"
#>
