param(
    [string]$resourceGroupName,
    [string]$name,
    [string]$userName,
    [string]$password
)

<# 
https://gallery.technet.microsoft.com/scriptcenter/Retrieve-all-Events-from-5db61ec8
https://blogs.technet.microsoft.com/nettracer/2013/01/18/rdp-connections-might-fail-due-to-a-problem-with-kb2621440-ms12-020/
https://p0w3rsh3ll.wordpress.com/2012/03/22/audit-rdp-connections/
https://social.technet.microsoft.com/wiki/contents/articles/37841.remote-desktop-services-rds-logon-connectivity-overview-and-troubleshooting.aspx#Capturing_a_network_trace1
https://serverfault.com/questions/882625/checking-the-encryption-level-of-remote-desktop-on-windows-server-2012
https://gallery.technet.microsoft.com/scriptcenter/Remote-Desktop-Connection-3fe225cd
#Installing PowerShell 4 or 5 on older version of Windows will not add the NetEventPacketCapture module, as that is an OS specific module, not a PowerShell specific module.
# Get-AzureRmNetworkWatcherReachabilityProvidersList
TODO: Parse for resets - "RST"
Parse for RDPBCGR
TLS Handshake Client Hello
TLS Handshake Server Hello Certificate
"Reason = Transport endpoint was not found." means no process is listening on 3389

Event logs:
Microsoft-Windows-TerminalServices-ClientActiveXCore
Microsoft-Windows-TerminalServices-RDPClient/Analytic
Microsoft-Windows-TerminalServices-RDPClient/Debug
Microsoft-Windows-TerminalServices-RDPClient/Operational

#>

function start-mstsc
{
    param(
        [string]$ipAddress,
        [string]$userName,
        [string]$password
    )
    
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $process = New-Object System.Diagnostics.Process
    $processInfo.FileName = "$($env:SystemRoot)\system32\cmdkey.exe"
    $processInfo.Arguments = "/generic:TERMSRV/$ipAddress /user:$userName /pass:$($password)"    
    $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    $process.StartInfo = $processInfo
    [void]$process.Start()
    $global:debugcmdkey = $process.StartInfo.Arguments

    $processInfo.FileName = "$($env:SystemRoot)\system32\mstsc.exe"
    $processInfo.Arguments = "/v $ipAddress"
    $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal
    $process.StartInfo = $processInfo
    [void]$process.Start()
    $global:debugmstsc = $process.StartInfo.Arguments
    return $process
}

$vm = get-azurermvm -ResourceGroupName $resourceGroupName -Name $name
$nicId = $vm.NetworkProfile.NetworkInterfaces.id
$nic = Get-AzureRmNetworkInterface | where Id -eq $nicId
$pipId = $nic.IpConfigurations[0].PublicIpAddress.id
$resourceGroupName = $vm.ResourceGroupName
$vmName = $vm.Name
$pip = (Get-AzureRmPublicIpAddress -ResourceGroupName $resourceGroupName | where Id -eq $pipId)[0]
$ipAddress = $pip.IpAddress
$localIpAddress = (get-netipaddress -AddressFamily IPv4 -PrefixOrigin Dhcp -SuffixOrigin Dhcp -AddressState Preferred -PolicyStore ActiveStore).IPAddress

"`$localIpAddress = $localIpAddress"
"`$ipAddress = $ipAddress"

$traceStartTime = get-date
$timestamp = get-date $traceStartTime.ToUniversalTime() -f 'MMddhhmmss'
$sessionName = "$vmName$timestamp"
$etlFilePath = "$pwd\$sessionName.etl"
"`$sessionName = $sessionName"
"`$etlFilePath = $etlFilePath"
$session = New-NetEventSession -Name $sessionName -LocalFilePath $etlFilePath

Add-NetEventProvider -Name 'Microsoft-Windows-TCPIP' -SessionName $sessionName
Add-NetEventProvider -Name 'Microsoft-Windows-Winsock-AFD' -SessionName $sessionName
Add-NetEventPacketCaptureProvider -SessionName $sessionName -Level 5 -TruncationLength 65535 -CaptureType 'BothPhysicalAndSwitch' -EtherType '0x0800' -IPAddresses $ipAddress,$localIpAddress # -IpProtocols 6,17
Start-NetEventSession -Name $sessionName

# Try RDP client connection, wait a little bit, then close it
$mstsc = start-mstsc -ipAddress $ipAddress -userName $userName -password $password
start-sleep -Seconds 30
$mstsc.CloseMainWindow()
$mstsc.Dispose()

Stop-NetEventSession -Name $sessionName

$log = Get-WinEvent -Path $etlFilePath -Oldest
"`$log.count = $($log.count)"
Remove-NetEventSession -Name $sessionName
$log | where ProviderName -eq 'Microsoft-Windows-TCPIP' | sort RecordId | select -first 20 | ft -a RecordId,TimeCreated,Message
$log | where ProviderName -eq 'Microsoft-Windows-Winsock-AFD' | sort RecordId | select -first 20 | ft -a RecordId,TimeCreated,Message
$global:debuglog = $log

$traceEndtime = get-date
$logName = 'Microsoft-Windows-TerminalServices-RDPClient/Operational'
$events = Get-WinEvent -FilterHashtable @{LogName = $logName;StartTime = $traceStartTime;EndTime = $traceEndTime} -ErrorAction 'SilentlyContinue'
$events | where ProviderName -eq 'Microsoft-Windows-TerminalServices-ClientActiveXCore' | sort RecordId | ft -a RecordId,TimeCreated,Id,Message
$event1027 = $events | where {$_.ProviderName -eq 'Microsoft-Windows-TerminalServices-ClientActiveXCore' -and $_.Id -eq 1027}
if($event1027){return $event1027}else{return "Event 1027 not found, RDP connection likely has failed"}
$global:debugevents = $events
