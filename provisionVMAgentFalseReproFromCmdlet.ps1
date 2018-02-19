$location = 'westcentralus'
$provisionVMAgent = $false
$userName = 'craig'
$vmSize = 'Standard_A1'
$publisherName = 'MicrosoftWindowsServer'
$offer = 'WindowsServer'
$skus = '2016-Datacenter'
$version = 'latest'
$password = $password
$vmName = "vm$(get-date -format "MMddhhmmss")"
$resourceGroupName = $vmName

$secureString = ConvertTo-SecureString -String $password -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($userName, $secureString) 
$resourceGroup = New-AzureRmResourceGroup -Name $resourceGroupName -Location $location
$subnetConfig = New-AzureRmVirtualNetworkSubnetConfig -Name "$($vmName)-subnet" -AddressPrefix '192.168.1.0/24'
$vnet = New-AzureRmVirtualNetwork -ResourceGroupName $resourceGroupName -Location $location -Name "$($vmName)-vnet" -AddressPrefix '192.168.0.0/16' -Subnet $subnetConfig
$pip = New-AzureRmPublicIpAddress -ResourceGroupName $resourceGroupName -Location $location -Name "$($vmName)-ip" -AllocationMethod Static -IdleTimeoutInMinutes 4
$ipAddress = $pip.ipAddress
$nsgRule = New-AzureRmNetworkSecurityRuleConfig -Name 'allow-rdp' -Protocol Tcp -Direction Inbound -Priority 1000 -SourceAddressPrefix * -SourcePortRange * -DestinationAddressPrefix * -DestinationPortRange 3389 -Access Allow
$nsg = New-AzureRmNetworkSecurityGroup -ResourceGroupName $resourceGroupName -Location $location -Name "$($vmName)-nsg" -SecurityRules $nsgRule
$nic = New-AzureRmNetworkInterface -ResourceGroupName $resourceGroupName -Location $location -Name "$($vmName)-nic" -SubnetId $vnet.Subnets[0].Id -PublicIpAddressId $pip.Id -NetworkSecurityGroupId $nsg.Id
$vm = New-AzureRmVMConfig -VMName $vmName -VMSize $vmSize
$vm = Set-AzureRmVMOperatingSystem -VM $vm -Windows -ComputerName $vmName -Credential $cred -ProvisionVMAgent:$provisionVMAgent
$vm = Set-AzureRmVMSourceImage -VM $vm -PublisherName $publisherName -Offer $offer -Skus $skus -Version $version
$vm = Add-AzureRmVMNetworkInterface -VM $vm -Id $nic.Id

$DebugPreference = 'Continue'
$file1 = ".\New-AzureRmVm_$($resourceGroupName)_$($vmName)_$(get-date -f yyyyMMddhhmmss).txt"
$vm = New-AzureRmVM -ResourceGroupName $resourceGroupName -Location $location -VM $vm *>&1 | tee-object -FilePath $file1
start-sleep -Seconds 10
invoke-item $file1