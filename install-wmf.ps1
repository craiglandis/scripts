<#
(new-object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/craiglandis/scripts/master/install-wmf.ps1', "$env:windir\temp\install-wmf.ps1"); set-executionpolicy unrestricted -force; invoke-expression -command "$env:windir\temp\install-wmf.ps1"
$imageName = 'MicrosoftWindowsServer.WindowsServer.2008-R2-SP1-smalldisk.2.127.20180613'
new -resourceGroupName test1 -name test1 -imageName 'MicrosoftWindowsServer.WindowsServer.2008-R2-SP1-smalldisk.2.127.20180613'
https://virtual-simon.co.uk/deploying-multiple-vms-arm-templates-use-copy-copyindex/
https://github.com/Azure/azure-quickstart-templates/tree/master/201-vm-copy-index-loops
https://github.com/Azure/azure-quickstart-templates/tree/master/201-vm-copy-managed-disks/
https://docs.microsoft.com/en-us/azure/azure-resource-manager/resource-group-create-multiple
https://stackoverflow.com/questions/50593595/arm-template-how-create-multiple-vms-with-private-ips-in-one-resource-group
#>

function out-log()
{
    param(
        [string]$text,
        [string]$prefix = 'timespan'
    )

    if ($prefix -eq 'timespan' -and $startTime)
    {
        $timespan = new-timespan -Start $startTime -End (get-date)
        $timespanString = '[{0:mm}:{0:ss}]' -f $timespan
        write-host $timespanString -nonewline -ForegroundColor White
        write-host " $text"
        (($timespanString + " $text") | out-string).Trim() | out-file $logFile -Append
    }
    elseif ($prefix -eq 'both' -and $startTime)
    {
        $timestamp = get-date -format "yyyy-MM-dd hh:mm:ss"
        $timespan = new-timespan -Start $startTime -End (get-date)
        $timespanString = "$($timestamp) $('[{0:mm}:{0:ss}]' -f $timespan)"
        write-host $timespanString -nonewline -ForegroundColor White
        write-host " $text"
        (($timespanString + " $text") | out-string).Trim() | out-file $logFile -Append
    }
    else
    {
        $timestamp = get-date -format "yyyy-MM-dd hh:mm:ss"
        write-host $timestamp -nonewline -ForegroundColor Cyan
        write-host " $text"
        (($timestamp + $text) | out-string).Trim() | out-file $logFile -Append
    }
}

set-strictmode -version Latest

$startTime = get-date
$timestamp = get-date $startTime -format yyyyMMddhhmmss
$scriptPath = $MyInvocation.MyCommand.Path
$scriptPathParent = split-path -path $MyInvocation.MyCommand.Path -parent
$scriptName = (split-path -path $MyInvocation.MyCommand.Path -leaf).Split('.')[0]
$logFile = "$scriptPathParent\$($scriptName)_$($timestamp).log"
out-log "scriptPath : $scriptPath"
out-log "scriptPathParent : $scriptPathParent"
out-log $logFile
$webClient = New-Object System.Net.WebClient

if ((Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 461814)
{
    out-log "NET Framework 4.72 already installed"
}
else
{
    out-log "Installing NET Framework 4.72"
    $url = 'https://download.microsoft.com/download/6/E/4/6E48E8AB-DC00-419E-9704-06DD46E5F81D/NDP472-KB4054530-x86-x64-AllOS-ENU.exe'
    $fileName = $url -split '/' | select -last 1
    $filePath = "$env:TEMP\$fileName"
    $webClient.DownloadFile($url, $filePath)
    invoke-expression -command "$filePath /q /norestart"

    do {
        out-log "Waiting for NET Framework 4.72 install to complete"
        start-sleep 5
    } until ((Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 461814)
    out-log "NET Framework 4.72 install completed"
}

if ($host.version -ge 5.1)
{
    out-log "Already using Powershell 5.1+"
}
else
{
    out-log "Installing WMF 5.1"
    $url = 'https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7AndW2K8R2-KB3191566-x64.zip'
    $fileName = $url -split '/' | select -last 1
    $filePath = "$env:TEMP\$fileName"
    $webClient.DownloadFile($url, $filePath)
    $extractedFilePath = "$env:TEMP\$($fileName.Split('.')[0])"

    if (!(test-path $extractedFilePath))
    {
        new-item -Path $extractedFilePath -ItemType Directory -Force | out-null
    }

    $shell = new-object -com shell.application
    $zip = $shell.NameSpace($filePath)
    foreach($item in $zip.items())
    {
        $shell.Namespace($extractedFilePath).copyhere($item)
    }

    while ((Get-ChildItem -Path $extractedFilePath -Recurse -Force).Count -lt $zip.items().Count)
    {
        start-sleep 1
    }

    invoke-expression -command "$extractedFilePath\Install-WMF5.1.ps1 -AcceptEULA"

    do {
        start-sleep 5
        out-log "Installing WMF 5.1..."
    } until (get-wmiobject -Query "Select HotFixID from Win32_QuickFixEngineering where HotFixID='KB3191566'")
    out-log "Creating onstart scheduled task to run script again at startup:"
    schtasks /create /tn bootstrap /sc onstart /delay 0000:30 /rl highest /ru system /tr "powershell.exe -executionpolicy bypass -file $scriptPath" /f
    if ($?)
    {
        out-log "Restarting to complete WMF 5.1 install"
        restart-computer -Force
        exit
    }
    else
    {
        exit
    }
}

if (get-wmiobject -Query "Select HotFixID from Win32_QuickFixEngineering where HotFixID='KB3191566'")
{
    out-log "Setting LastFullPayloadTime reg value to workaround WMF 5.1 sysprep issue"
    New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\StreamProvider -Name LastFullPayloadTime -Value 0 -PropertyType DWord -Force | Out-Null
}

if (get-packageprovider | where {$_.Name -eq 'NuGet'})
{
    out-log "NuGet already installed"
}
else
{
    out-log "Installing NuGet"
    install-packageprovider -name NuGet -Force | Out-Null
}

if ((Get-PSRepository -Name PSGallery).InstallationPolicy -eq 'Trusted')
{
    out-log "PSGallery installation policy is already Trusted"
}
else
{
    out-log "Setting PSGallery installation policy to Trusted"
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
}

if (get-module -name PSWindowsUpdate -ListAvailable)
{
    out-log "PSWindowsUpdate module already installed"
}
else
{
    out-log "Installing PSWindowsUpdate module"
    install-module -name PSWindowsUpdate -Force
}

out-log "Installing Windows Updates"
get-windowsupdate -Install -AcceptAll -AutoReboot -IgnoreUserInput -RecurseCycle 5 -verbose

<#
c:\windows\system32\sysprep\sysprep.exe /generalize /oobe /shutdown
#>
