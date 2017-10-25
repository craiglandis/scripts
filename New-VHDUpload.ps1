param(
    $resourceGroupName,
    $storageAccountName,
    $location,
    $containerName,
    $localFilePath,
    $skuName ='Standard_LRS'
)

# Create resource group if it doesn't exit
if(!(Get-AzureRmResourceGroup -Name $resourceGroupName -Location $location -ErrorAction SilentlyContinue))
{
    New-AzureRmResourceGroup -Name $resourceGroupName -Location $location
}

# Create storage account if it doesn't exit
if(!(Get-AzureRmStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue))
{
    $storageAccount = New-AzureRmStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -SkuName $skuName -Location $location
}

$storageAccountKey = (Get-AzureRmStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value
$storageContext = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey

# Create contain if it doesn't exist
if(!(Get-AzureStorageContainer -Context $storageContext -Name $containerName -ErrorAction SilentlyContinue))
{
    New-AzureStorageContainer -Context $storageContext -Name $containerName
}

if ($storageAccount)
{
    $destination = "$($storageAccount.PrimaryEndpoints.Blob)$containerName/$(split-path $localFilePath -leaf)"
}
else 
{
    Write-Host "Make sure storage account exists: $storageAccountName"    
    exit
}

# Upload VHD
$command = "Add-AzureRmVhd -ResourceGroupName $resourceGroupName -Destination $destination -LocalFilePath $localFilePath"
write-host $command 
Invoke-Expression $command
