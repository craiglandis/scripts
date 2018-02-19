$template = @'
{
    "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#",
    "contentVersion": "1.0.0.0",
    "parameters": {
      "addressPrefix": {"defaultValue": "10.0.0.0/16", "type": "string"},
      "adminUsername": {"type": "string"},
      "adminPassword": {"type": "securestring"},
      "diagnosticsStorageAccountName": {"type": "string"},
      "diagnosticsStorageAccountType": {"defaultValue": "Standard_LRS", "type": "string"},
      "dnsLabelPrefix": {"type": "string"},
      "enableAutomaticUpdates": {"defaultValue": true, "type": "bool"},    
      "nicName": {"type": "string"},
      "nsgName": {"type": "string"},
      "offer": {"type": "string"},
      "pipName": {"type": "string"},
      "provisionVMAgent": {"defaultValue": false, "type": "bool"},
      "publisher": {"type": "string"},
      "sku": {"type": "string"},
      "storageAccountName": {"type": "string"},
      "storageAccountType": {"defaultValue": "Standard_LRS", "type": "string"},
      "subnetName": {"defaultValue": "default", "type": "string"},
      "subnetPrefix": {"defaultValue": "10.0.0.0/24", "type": "string"},
      "timeZone": {"defaultValue": "GMT Standard Time", "type": "string"},
      "version": {"type": "string"},
      "vmName": {"type": "string"},
      "vmSize": {"defaultValue": "Standard_A1", "type": "string"},
      "vnetName": {"type": "string"}
    },
    "variables": {
      "subnetRef": "[resourceId('Microsoft.Network/virtualNetworks/subnets', parameters('vnetName'), parameters('subnetName'))]"
    },
    "resources": [
      {
        "type": "Microsoft.Storage/storageAccounts",
        "name": "[parameters('storageAccountName')]",
        "apiVersion": "2017-06-01",
        "location": "[resourceGroup().location]",
        "sku": {
          "name": "[parameters('storageAccountType')]"
        },
        "kind": "Storage",
        "properties": {}
      },
      {
        "type": "Microsoft.Storage/storageAccounts",
        "name": "[parameters('diagnosticsStorageAccountName')]",
        "apiVersion": "2017-06-01",
        "location": "[resourceGroup().location]",
        "sku": {
          "name": "[parameters('diagnosticsStorageAccountType')]"
        },
        "kind": "Storage",
        "properties": {}
      },
      {
        "apiVersion": "2017-06-01",
        "type": "Microsoft.Network/publicIPAddresses",
        "name": "[parameters('pipName')]",
        "location": "[resourceGroup().location]",
        "properties": {
          "publicIPAllocationMethod": "Dynamic",
          "dnsSettings": {
            "domainNameLabel": "[parameters('dnsLabelPrefix')]"
          }
        }
      },
      {
        "apiVersion": "2017-06-01",
        "type": "Microsoft.Network/virtualNetworks",
        "name": "[parameters('vnetName')]",
        "location": "[resourceGroup().location]",
        "properties": {
          "addressSpace": {
            "addressPrefixes": [
              "[parameters('addressPrefix')]"
            ]
          },
          "subnets": [
            {
              "name": "[parameters('subnetName')]",
              "properties": {
                "addressPrefix": "[parameters('subnetPrefix')]"
              }
            }
          ]
        }
      },
      {
        "apiVersion": "2017-06-01",
        "type": "Microsoft.Network/networkInterfaces",
        "name": "[parameters('nicName')]",
        "location": "[resourceGroup().location]",
        "dependsOn": [
          "[resourceId('Microsoft.Network/publicIPAddresses/', parameters('pipName'))]",
          "[resourceId('Microsoft.Network/virtualNetworks/', parameters('vnetName'))]"
        ],
        "properties": {
          "ipConfigurations": [
            {
              "name": "ipconfig1",
              "properties": {
                "privateIPAllocationMethod": "Dynamic",
                "publicIPAddress": {
                  "id": "[resourceId('Microsoft.Network/publicIPAddresses',parameters('pipName'))]"
                },
                "subnet": {
                  "id": "[variables('subnetRef')]"
                }
              }
            }
          ],
          "networkSecurityGroup": {
            "id": "[resourceId('Microsoft.Network/networkSecurityGroups', parameters('nsgName'))]"
          }
        }
      },
      {
        "name": "[parameters('nsgName')]",
        "type": "Microsoft.Network/networkSecurityGroups",
        "apiVersion": "2016-09-01",
        "location": "[resourceGroup().location]",
        "properties": {
          "securityRules": [
            {
              "name": "default-allow-rdp",
              "properties": {
                "priority": 1000,
                "sourceAddressPrefix": "*",
                "protocol": "TCP",
                "destinationPortRange": "3389",
                "access": "Allow",
                "direction": "Inbound",
                "sourcePortRange": "*",
                "destinationAddressPrefix": "*"
              }
            }
          ]
        }
      },
      {
        "apiVersion": "2017-03-30",
        "type": "Microsoft.Compute/virtualMachines",
        "name": "[parameters('vmName')]",
        "location": "[resourceGroup().location]",
        "dependsOn": [
          "[resourceId('Microsoft.Storage/storageAccounts/', parameters('storageAccountName'))]",
          "[resourceId('Microsoft.Network/networkInterfaces/', parameters('nicName'))]"
        ],
        "properties": {
          "hardwareProfile": {
              "vmSize": "[parameters('vmSize')]"
          },
          "osProfile": {
            "computerName": "[parameters('vmName')]",
            "adminUsername": "[parameters('adminUsername')]",
            "adminPassword": "[parameters('adminPassword')]",
            "windowsConfiguration": {
              "provisionVMAgent": "[parameters('provisionVMAgent')]",
              "enableAutomaticUpdates": "[parameters('enableAutomaticUpdates')]",
              "timeZone": "[parameters('timeZone')]"
            }
          },
          "storageProfile": {
            "imageReference": {
              "publisher": "[parameters('publisher')]",
              "offer": "[parameters('offer')]",
              "sku": "[parameters('sku')]",
              "version": "[parameters('version')]"
            },
            "osDisk": {
              "createOption": "FromImage"
            },
            "dataDisks": []
          },
          "networkProfile": {
            "networkInterfaces": [
              {
                "id": "[resourceId('Microsoft.Network/networkInterfaces',parameters('nicName'))]"
              }
            ]
          },
          "diagnosticsProfile": {
            "bootDiagnostics": {
              "enabled": true,
              "storageUri": "[reference(resourceId('Microsoft.Storage/storageAccounts/', parameters('diagnosticsStorageAccountName'))).primaryEndpoints.blob]"
            }
          }
        }
      }
    ],
    "outputs": {
      "hostname": {
        "type": "string",
        "value": "[reference(parameters('pipName')).dnsSettings.fqdn]"
      }
    }
  }
'@

$templatefile = ".\provisionVMAgentFalseReproFromTemplate.json"
$template | out-file $templatefile -Force

$location = 'westcentralus'
$vmName = "vm$(get-date -format "MMddhhmmss")"
$resourceGroupName = $vmName
$resourceGroup = New-AzureRmResourceGroup -Name $resourceGroupName -Location $location

$params = [ordered]@{
    publisher = 'MicrosoftWindowsServer'
    offer = 'WindowsServer'
    sku = '2016-Datacenter'
    version = 'latest'
    addressPrefix = "10.0.0.0/16"
    adminUsername = 'craig'
    adminPassword = $password
    dnsLabelPrefix = $vmName.ToLower()
    diagnosticsStorageAccountName = "$($vmName.ToLower())diagsa1"
    enableAutomaticUpdates = $true
    nicName = "$($vmName)-nic"
    nsgName = "$($vmName)-nsg"
    pipName = "$($vmName)-ip"
    provisionVMAgent = $false
    storageAccountName = "$($vmName.ToLower())sa1"
    storageAccountType = 'Standard_LRS'
    subnetName = "$($vmName)-subnet"
    subnetPrefix = "10.0.0.0/24"
    timeZone = "GMT Standard Time"    
    vmName = $vmName
    vmSize = 'Standard_A1'
    vnetName = "$($vmName)-vnet"
}

New-AzureRmResourceGroupDeployment -Name $resourceGroupName -ResourceGroupName $resourceGroupName -TemplateParameterObject $params -TemplateFile $templateFile