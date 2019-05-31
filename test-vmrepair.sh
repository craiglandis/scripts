#!/bin/bash

location="westus2"
timestamp="$(date -d today +%m%d%H%M%S)"
vmname="vm$timestamp"
image="UbuntuLTS"
az group create --name $vmname --location $location
az vm create --name $vmname --resource-group $vmname --image $image --generate-ssh-keys

az extension show --name vm-repair --output none
if [ $? -eq 1 ]
then
    az extension add --name vm-repair
fi

az vm repair create --name $vmname --resource-group $vmname --repair-password "$(openssl rand -base64 12)" --verbose
az vm repair restore --name $vmname --resource-group $vmname --yes --verbose
