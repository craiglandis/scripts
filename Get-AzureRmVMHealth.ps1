param(
    $resourceGroupName,
    $name
)

$vm = get-azurermvm -ResourceGroupName $resourceGroupName -Name $name -ErrorAction Stop
$vmstatus = $vm | get-azurermvm -status -ErrorAction Stop


if ($vm.DiagnosticsProfile.Bootdiagnostics.Enabled)
{
    if ($vmstatus.bootdiagnostics.ConsoleScreenshotBlobUri)
    {
        $consoleScreenshotBlobUri = $vmstatus.bootdiagnostics.ConsoleScreenshotBlobUri
    }
    else
    {
        "ConsoleScreenshotBlobUri property not populated"
        exit
    }
}
else
{
    "Bootdiagnostics: $($vm.DiagnosticsProfile.Bootdiagnostics.Enabled)"
    exit
}

$storageAccountName = $consoleScreenshotBlobUri.split('/')[2].split('.')[0]
$storageContainer = $consoleScreenshotBlobUri.split('/')[3]
#TODO If boot diag storage account can reside in a different RG than the VM's RG, need a different way to get the RG of the boot diag storage account
$storageContext = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey (Get-AzureRmStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value
$blobs = get-azurestorageblob -Container $storageContainer -Context $storageContext
$logs = $blobs | where {$_.Name.EndsWith('.serialconsole.log')}
$logs | Get-AzureStorageBlobContent -Force | Out-Null
$logs | foreach {
    $logFileName = $_.name
    $logFilePath = "$pwd\$logFileName"
    $jsonFileName = "$($logFileName.SubString(0,$logFileName.Length-4)).json"
    $jsonFilePath = "$pwd\$jsonFileName"
    $log = get-content $logFilePath
    for($i=0; $i -le ($log.count); $i++){if($log[$i] -match 'Microsoft Azure VM Health Report - Start'){$jsonString = $log[$i+1]}}  
    $json = $jsonString | ConvertFrom-Json
    if ($json){$json | ConvertTo-Json | out-file $jsonFilePath}
    get-content $jsonFilePath
    "`nserial log:     $logFilePath"
    "vm health json: $jsonFilePath"
    "logFilePath: $logFilePath"
    "Boot diag storage account: $storageAccountName"
    "VMAgentVersion: $($vmstatus.vmagent.VmAgentVersion)"
}
