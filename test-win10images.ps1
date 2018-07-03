function write-log
{
    param(
        [string]$status,
        [string]$color = 'cyan',
        [switch]$logOnly
    )    

    $timestampString = get-date (get-date).ToUniversalTime() -Format 'yyyy-MM-ddTHH:mm:ssZ'
    $duration = new-timespan -Start $script:scriptStartTime -End (get-date).ToUniversalTime()
    $durationString = '[{0:hh}:{0:mm}:{0:ss}.{0:ff}]' -f $duration
    $timestampAndDurationString = "$timestampString $durationString"

    if ($logOnly)
    {
        "$timestampAndDurationString $status" | out-file $logFile -append 
    }
    else
    {
        write-host $timestampAndDurationString -nonewline -ForegroundColor $color
        write-host " $status"
        "$timestampAndDurationString $status" | out-file $logFile -append 
    }
}

$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

$script:scriptStartTime = (get-date).ToUniversalTime()
$startTimeString = get-date $script:scriptStartTime -f 'yyyyMMddHHmmss'
$scriptName = $MyInvocation.MyCommand.Name.Split('.')[0]
$logFile = "$($env:TEMP)\$($scriptName)_$($startTimeString).log"
$transcript = "$($env:TEMP)\$($scriptName)_$($startTimeString)_transcript.log"
$startTranscriptResult = start-transcript -Path $transcript -IncludeInvocationHeader
write-log "Log: $logFile"
write-log $startTranscriptResult

$scriptPath = split-path -parent $MyInvocation.MyCommand.Source
set-location $scriptPath
write-log $MyInvocation.Line

### RS3 6B (includes the fix) ###
write-log "RS36BA1 1-core, standard storage, Win10 RS3 w/6B package"
new.ps1 -resourceGroupName RS36BA1 -vmName RS36BA1 -vmSize Standard_A1 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.492.43

write-log "RS36BA2 2-core, standard storage, Win10 RS3 w/6B package"
new.ps1 -resourceGroupName RS36BA2 -vmName RS36BA2 -vmSize Standard_A2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.492.43

write-log "RS36BA3 4-core, standard storage, Win10 RS3 w/6B package"
new.ps1 -resourceGroupName RS36BA3 -vmName RS36BA3 -vmSize Standard_A3 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.492.43

write-log "RS36BDS1 1-core, premium storage, Win10 RS3 w/6B package"
new.ps1 -resourceGroupName RS36BDS1 -vmName RS36BDS1 -vmSize Standard_DS1_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.492.43

write-log "RS36BDS2 2-core, premium storage, Win10 RS3 w/6B package"
new.ps1 -resourceGroupName RS36BDS2 -vmName RS36BDS2 -vmSize Standard_DS2_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.492.43

write-log "RS36BDS3 4-core, premium storage, Win10 RS3 w/6B package"
new.ps1 -resourceGroupName RS36BDS3 -vmName RS36BDS3 -vmSize Standard_DS3_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.492.43

## RS3 5B (does NOT include the fix) ###
write-log "RS35BA1 1-core, standard storage, Win10 RS3 w/5B package"
new.ps1 -resourceGroupName RS35BA1 -vmName RS35BA1 -vmSize Standard_A1 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.431.41

write-log "RS35BA2 2-core, standard storage, Win10 RS3 w/5B package"
new.ps1 -resourceGroupName RS35BA2 -vmName RS35BA2 -vmSize Standard_A2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.431.41

write-log "RS35BA3 4-core, standard storage, Win10 RS3 w/5B package"
new.ps1 -resourceGroupName RS35BA3 -vmName RS35BA3 -vmSize Standard_A3 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.431.41

write-log "RS35BDS1 1-core, premium storage, Win10 RS3 w/5B package"
new.ps1 -resourceGroupName RS35BDS1 -vmName RS35BDS1 -vmSize Standard_DS1_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.431.41

write-log "RS35BDS2 2-core, premium storage, Win10 RS3 w/5B package"
new.ps1 -resourceGroupName RS35BDS2 -vmName RS35BDS2 -vmSize Standard_DS2_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.431.41

write-log "RS35BDS3 4-core, premium storage, Win10 RS3 w/5B package"
new.ps1 -resourceGroupName RS35BDS3 -vmName RS35BDS3 -vmSize Standard_DS3_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS3-Pro -version 16299.431.41

## RS4 6B (does NOT include fix) ###
write-log "RS46BA1 1-core, standard storage, Win10 RS4 w/6B package"
new.ps1 -resourceGroupName RS46BA1 -vmName RS46BA1 -vmSize Standard_A1 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.112.43

write-log "RS46BA2 2-core, standard storage, Win10 RS4 w/6B package"
new.ps1 -resourceGroupName RS46BA2 -vmName RS46BA2 -vmSize Standard_A2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.112.43

write-log "RS46BA3 4-core, standard storage, Win10 RS4 w/6B package"
new.ps1 -resourceGroupName RS46BA3 -vmName RS46BA3 -vmSize Standard_A3 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.112.43

write-log "RS46BDS1 1-core, premium storage, Win10 RS4 w/6B package"
new.ps1 -resourceGroupName RS46BDS1 -vmName RS46BDS1 -vmSize Standard_DS1_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.112.43

write-log "RS46BDS2 2-core, premium storage, Win10 RS4 w/6B package"
new.ps1 -resourceGroupName RS46BDS2 -vmName RS46BDS2 -vmSize Standard_DS2_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.112.43

write-log "RS46BDS3 4-core, premium storage, Win10 RS4 w/6B package"
new.ps1 -resourceGroupName RS46BDS3 -vmName RS46BDS3 -vmSize Standard_DS3_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.112.43

## RS4 5B (does NOT include fix) ###
write-log "RS45BA1 1-core, standard storage, Win10 RS4 w/5B package"
new.ps1 -resourceGroupName RS45BA1 -vmName RS45BA1 -vmSize Standard_A1 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.48.41

write-log "RS45BA2 2-core, standard storage, Win10 RS4 w/5B package"
new.ps1 -resourceGroupName RS45BA2 -vmName RS45BA2 -vmSize Standard_A2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.48.41

write-log "RS45BA3 4-core, standard storage, Win10 RS4 w/5B package"
new.ps1 -resourceGroupName RS45BA3 -vmName RS45BA3 -vmSize Standard_A3 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.48.41

write-log "RS45BDS1 1-core, premium storage, Win10 RS4 w/5B package"
new.ps1 -resourceGroupName RS45BDS1 -vmName RS45BDS1 -vmSize Standard_DS1_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.48.41

write-log "RS45BDS2 2-core, premium storage, Win10 RS4 w/5B package"
new.ps1 -resourceGroupName RS45BDS2 -vmName RS45BDS2 -vmSize Standard_DS2_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.48.41

write-log "RS45BDS3 4-core, premium storage, Win10 RS4 w/5B package"
new.ps1 -resourceGroupName RS45BDS3 -vmName RS45BDS3 -vmSize Standard_DS3_v2 -singleWindowsVM -publisherName MicrosoftWindowsDesktop -offer Windows-10 -skus RS4-Pro -version 17134.48.41

<#
get-azurermvmsize -location westus2 | where NumberOfCores -eq 1

Standard_A1
Standard_DS1_v2

get-azurermvmsize -location westus2 | where NumberOfCores -eq 2

Standard_A2
Standard_DS2_v2

get-azurermvmsize -location westus2 | where NumberOfCores -eq 4

Standard_A3
Standard_DS3_v2

PublisherName           Offer      Skus     Version
-------------           -----      ----     -------
MicrosoftWindowsDesktop Windows-10 RS3-Pro  16299.431.41
MicrosoftWindowsDesktop Windows-10 RS3-Pro  16299.492.43
MicrosoftWindowsDesktop Windows-10 RS3-ProN 16299.431.41
MicrosoftWindowsDesktop Windows-10 RS3-ProN 16299.492.43
MicrosoftWindowsDesktop Windows-10 RS4-Pro  17134.112.43
MicrosoftWindowsDesktop Windows-10 RS4-Pro  17134.48.41
MicrosoftWindowsDesktop Windows-10 RS4-ProN 17134.112.43
MicrosoftWindowsDesktop Windows-10 RS4-ProN 17134.48.41
#>