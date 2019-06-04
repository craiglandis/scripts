#
#(new-object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/craiglandis/scripts/master/install-wmf.ps1', "$env:temp\install-wmf.ps1"); set-executionpolicy unrestricted -force; invoke-expression -command "$env:temp\install-wmf.ps1"


$webClient = New-Object System.Net.WebClient

if ((Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 461814)
{
    write-host "NET Framework 4.72 already installed"
}
else
{
    write-host "Installing NET Framework 4.72"
    $url = 'https://download.microsoft.com/download/6/E/4/6E48E8AB-DC00-419E-9704-06DD46E5F81D/NDP472-KB4054530-x86-x64-AllOS-ENU.exe'
    $fileName = $url -split '/' | select -last 1
    $filePath = "$env:TEMP\$fileName"
    $webClient.DownloadFile($url, $filePath)
    invoke-expression -command "$filePath /q /norestart"

    do {
        write-host "Waiting for NET Framework 4.72 install to complete"
        start-sleep 5
    } until ((Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release -ge 461814)
    write-host "NET Framework 4.72 install completed"
}

if (get-wmiobject -Query "Select HotFixID from Win32_QuickFixEngineering where HotFixID='KB3191566'")
{
    write-host "WMF 5.1 already installed"
}
else
{
    write-host "Installing WMF 5.1"
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
    } until (get-wmiobject -Query "Select HotFixID from Win32_QuickFixEngineering where HotFixID='KB3191566'")
    write-host "restart-computer -force"
}

New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\StreamProvider -Name LastFullPayloadTime -Value 0 -PropertyType DWord -Force
install-packageprovider -name NuGet -MinimumVersion 2.8.5.201 -Force
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
install-module -name PSWindowsUpdate -AllowClobber -Force

#schtasks /create /tn install /tr c:\apps\myapp.exe /sc onstart

<#
get-windowsupdate -Install -AcceptAll -AutoReboot -IgnoreUserInput -RecurseCycle 5 -verbose
c:\windows\system32\sysprep\sysprep.exe /generalize /oobe /shutdown
#>
