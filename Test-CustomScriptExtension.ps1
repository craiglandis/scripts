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

'This is "no cmdlet" output'
Write-Debug 'This is Write-Debug output'
Write-Error 'This is Write-Error output'
Write-Host 'This is Write-Host output'
Write-Information 'This is Write-Information output'
Write-Output 'This is Write-Output output'
Write-Progress 'This is Write-Progress output'
Write-Warning 'This is Write-Warning output'
Write-Verbose 'This is Write-Verbose output'

#exit 1
Stop-Transcript

<#
New-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Force
New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -PropertyType 'DWord' -Value 1 -Force
New-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -PropertyType 'DWord' -Value 0 -Force
Get-Item -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' #-ErrorAction SilentlyContinue
Get-ItemPropertyValue -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path 'HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging' -Name "EnableScriptBlockLogging"
#>