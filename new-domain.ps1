param(
    $resourceGroupName = 'adrg',
    $location = 'westus2',
    $userName = 'craig',
    $password = $password,
    $domainName = 'contoso.com',
    $dnsPrefix = 'advm',
    $templateUri = 'https://raw.githubusercontent.com/Azure/azure-quickstart-templates/master/active-directory-new-domain/azuredeploy.json'
)

$resourceGroup = get-azurermresourcegroup -Name $resourceGroupName -ErrorAction SilentlyContinue
if (-not $resourceGroup)
{
    New-AzureRmResourceGroup -Name $resourceGroupName -Location $location -ErrorAction Stop
}

$templateParametersObject = @{
    adminUsername = $userName
    adminPassword = $password
    domainName = $domainName
    dnsPrefix = $dnsPrefix
}

New-AzureRmResourceGroupDeployment -ResourceGroupName $resourceGroupName -TemplateUri $templateUri -TemplateParameterObject $templateParametersObject -ErrorAction Stop
