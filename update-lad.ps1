param ( 
	[Parameter(Mandatory=$true)][string]$resourceGroupName,
	[Parameter(Mandatory=$true)][string]$vmName,
    [string]$storageAccountName
)

write-host "Getting diagnostics storage account name"
$vm = get-azurermvm -resourceGroupName $resourceGroupName -Name $vmName
$location = $vm.location
$extension = $vm.extensions | where {$_.VirtualMachineExtensionType -eq 'LinuxDiagnostic'}
if(!$storageAccountName)
{
    if ($extension)
    {
        $storageAccountName = ($extension.settings.ToString() | convertfrom-json).StorageAccount
    }
    if(!$storageAccountName)
    {
        if ($vm.diagnosticsprofile.bootdiagnostics.storageuri)
        {
            $storageAccountName = $vm.diagnosticsprofile.bootdiagnostics.storageuri.Split('.')[0].Split('/')[2]
            if(!$storageAccountName)
            {                
                write-host "Please rerun with -storageAccountName <storageAccountName>"
                exit
            }
        }
    }
}
write-host "storageAccountName: $storageAccountName"

write-host "Creating storage account SAS token"
$storageAccountKey = ((Get-AzureRmStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName) | where KeyName -eq 'key1').value
$storageContext = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey
$startTime = (get-date).AddMinutes(-15).ToUniversalTime()
$expiryTime = (get-date).AddDays(7).ToUniversalTime()
$storageAccountSasToken = New-AzureStorageAccountSASToken -Service Blob,Table -ResourceType Container,Object -Permission 'wlacu' -Protocol HttpsOnly -Context $storageContext -StartTime $startTime -ExpiryTime $expiryTime
write-host "storageAccountSasToken: $storageAccountSasToken"

if ($extension)
{
    write-host "Removing extension: $($extension.VirtualMachineExtensionType) publisher: $($extension.publisher) version: $($extension.TypeHandlerVersion)"
    remove-AzureRmVMExtension -ResourceGroupName $resourceGroupName -VMName $vmName -Name $extension.name -force
}

$publisher = 'Microsoft.Azure.Diagnostics'
$extensionType = 'LinuxDiagnostic'
$extensionName = $extensionType
write-host "Getting latest version number for extension: $extensionType publisher: $publisher"
$extension = get-azurermvmextensionimage -Location $location -PublisherName $publisher -Type $extensionType
$version = $extension[-1].Version
$version = ($version.Split('.')[0] + '.' + $version.Split('.')[1])
write-host "version: $version"

$webClient = New-Object System.Net.WebClient
$url = 'https://raw.githubusercontent.com/Azure/azure-linux-extensions/master/Diagnostic/tests/lad_2_3_compatible_portal_pub_settings.json'
$fileName = $url.Split('/')[-1]
$filePath = "$pwd\$fileName"
write-host "Downloading json file: $url"
$webclient.DownloadFile($url,$filepath)

write-host "Updating StorageAccount and resourceId values in json"
$settings = get-content $filePath | convertfrom-json
$settings.StorageAccount = $storageAccountName
$settings.ladcfg.diagnosticMonitorConfiguration.metrics.resourceid = $vm.id
$settingString = $settings | ConvertTo-Json -Depth 100

write-host "Creating protectedSettingString with StorageAccount: $storageAccountName storageAccountSasToken: $storageAccountSasToken"
$protectedSettingString = ('{"storageAccountName": "' + $storageAccountName + '","storageAccountSasToken": "' + $storageAccountSasToken + '"}')

write-host "Adding extension: $extensionType publisher: $publisher version: $version to VM: $vmName in resource group: $resourceGroupName"
set-AzureRmVMExtension -ResourceGroupName $resourceGroupName -VMName $vmName -Location $location -Name $extensionName -Publisher $publisher -ExtensionType $extensionType -TypeHandlerVersion $version -SettingString $settingString -ProtectedSettingString $protectedSettingString

write-host "Getting extension status"
$vm = get-azurermvm -resourceGroupName $resourceGroupName -Name $vmName -Status
$extension = $vm.extensions | where {$_.Name -eq 'LinuxDiagnostic'}
$extension.statuses

write-host "Getting table names in diagnostics storage account $storageAccountName"
$table = get-azurestoragetable -Context $storageContext
$table.Name

#$publicSettings = (Get-AzureRmVMExtension -ResourceGroupName rh2 -VMName rh2 -Name Microsoft.Insights.VMDiagnosticsSettings).publicsettings
#$protectedSettingString = ('{"storageAccountName": "' + $storageAccountName + '","storageAccountKey": "' + $storageAccountKey + '"}')
<#
LinuxCpuVer2v0
LinuxDiskVer2v0
LinuxMemoryVer2v0
LinuxsyslogVer2v0
SchemasTable
WADMetricsPT1HP10DV2S20170530
WADMetricsPT1MP10DV2S20170530

# Export for comparison if needed later
#$vm | Export-Clixml -Depth 100 -Path ".\$(($vm.id.Replace('/','_') + '_BEFORE.xml'))"

#$vm | Export-Clixml -Depth 100 -Path ".\$(($vm.id.Replace('/','_') + '_AFTER.xml'))"

#Create a table query.
$tableQuery = New-Object Microsoft.WindowsAzure.Storage.Table.TableQuery

#Define columns to select.
$list = New-Object System.Collections.Generic.List[string]
$list.Add("Timestamp")
$list.Add("Facility")
$list.Add("Msg")

#Set query details.
$tableQuery.FilterString = "Facility eq kern"
$tableQuery.SelectColumns = $list
$tableQuery.TakeCount = 20

#Execute the query.
$tableEntities = $table.CloudTable.ExecuteQuery($tableQuery)

#Display entity properties with the table format.
$tableEntities  | Format-Table PartitionKey, RowKey, @{ Label = "Name"; Expression={$_.Properties["Name"].StringValue}}, @{ Label = "ID"; Expression={$_.Properties["ID"].Int32Value}} -AutoSize

$extension = get-azurermvmextension -ResourceGroupName $resourceGroupName -VMName $vmName -Name $name
$publicSettings = $extension.publicsettings 
$xmlCfg = ($publicSettings | convertfrom-json).xmlCfg
#>
