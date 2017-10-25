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
    $resourceGroup = New-AzureRmResourceGroup -Name $resourceGroupName -Location $location
}

# Create storage account if it doesn't exit
if ($resourceGroup)
{
    if(!(Get-AzureRmStorageAccount -Name $storageAccountName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue))
    {
        $storageAccount = New-AzureRmStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -SkuName $skuName -Location $location
    }
}
else
{
    Write-Host "Resource group not found: $resourceGroupName"    
    exit
}

if ($storageAccount)
{
    $storageAccountKey = (Get-AzureRmStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value
    $storageContext = New-AzureStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey

    # Create container if it doesn't exist
    if(!(Get-AzureStorageContainer -Context $storageContext -Name $containerName -ErrorAction SilentlyContinue))
    {
        $container = New-AzureStorageContainer -Context $storageContext -Name $containerName
    }    
}
else
{
    Write-Host "Storage account not found: $storageAccountName"    
    exit
}

if (test-path $localFilePath)
{
    $destination = "$($storageAccount.PrimaryEndpoints.Blob)$containerName/$(split-path $localFilePath -leaf)"
}
else
{
    Write-Host "File not found: $localFilePath"
    exit
}

# Upload VHD
if ($container)
{
    $command = "Add-AzureRmVhd -ResourceGroupName $resourceGroupName -Destination $destination -LocalFilePath $localFilePath"
    write-host $command
    Invoke-Expression $command
}
else
{
    Write-Host "Container not found: $containerName"
    exit
}
