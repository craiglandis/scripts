param(
    $vmName,
    $resourceGroupName = 'adrg',
    $location = 'westus2',
    $userName = 'craig',
    $password = $password,
    $domainJoinUserName = 'craig',
    $domainJoinUserPassword = $password,
    $dnsLabelPrefix,
    $domainFQDN = 'contoso.com',
    $ouPath = '',
    $templateUri = 'https://raw.githubusercontent.com/craiglandis/scripts/master/newmemberserver.json',
    [switch]$test,
    $timestamp = $(get-date ((get-date).ToUniversalTime()) -format yyMMddHHmmss)
)

$resourceGroup = get-azurermresourcegroup -Name $resourceGroupName -ErrorAction SilentlyContinue
if (-not $resourceGroup)
{
    New-AzureRmResourceGroup -Name $resourceGroupName -Location $location -ErrorAction Stop
}

# The params are of type securestring in the template itself, so I think convert them here results in them being converted twice, which is why the password is never working then.
#$secureString = ConvertTo-SecureString $password -AsPlainText -Force

$templateParametersObject = @{
    #vmName = $vmName
    adminUsername = $userName
    adminPassword = $password
    #adminPassword = "$secureString"
    #dnsLabelPrefix = $dnsLabelPrefix
    domainJoinUserName = $domainJoinUserName
    domainJoinUserPassword = $password
    #domainJoinUserPassword = "$secureString"
    domainFQDN = $domainFQDN
    ouPath = $ouPath
}

if ($test)
{
    Test-AzureRmResourceGroupDeployment -ResourceGroupName $resourceGroupName -TemplateUri $templateUri -TemplateParameterObject $templateParametersObject -timestamp $timestamp -ErrorAction Stop
}
else
{
    New-AzureRmResourceGroupDeployment -ResourceGroupName $resourceGroupName -TemplateUri $templateUri -TemplateParameterObject $templateParametersObject -timestamp $timestamp -ErrorAction Stop
}