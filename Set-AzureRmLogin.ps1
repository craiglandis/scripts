param(
    $name = $env:COMPUTERNAME
)

# All of these AAD cmdlets require you be logged in manually first
Connect-AzureRmAccount

$oldapp = Get-AzureRmADApplication -ObjectId (Get-AzureRmADApplication -IdentifierUri "https://$name").ObjectId.Guid

if ($oldapp)
{
    $choice = read-host -Prompt "Application exists with DisplayName $name, do you want to remove it? [Y/N]"
    if ($choice.ToUpper() -eq 'Y')
    {
        Remove-AzureRmADApplication -ObjectId (Get-AzureRmADApplication -IdentifierUri "https://$name").ObjectId.Guid -Force
    }
    else 
    {
        exit
    }
}
else 
{
    "No app found with DisplayName $name."    
}

$cert = New-SelfSignedCertificate -CertStoreLocation 'cert:\CurrentUser\My' -Subject "CN=$name" -KeySpec KeyExchange
$certValue = [System.Convert]::ToBase64String($cert.GetRawCertData())
$azureAdApplication = New-AzureRmADApplication -DisplayName $name -HomePage "https://$name" -IdentifierUris "https://$name" -certValue $certValue -EndDate $cert.NotAfter -StartDate $cert.NotBefore
New-AzureRmADServicePrincipal -ApplicationId $azureAdApplication.ApplicationId

for ($i = 0; $i -ile 10; $i++){
    $servicPrincipal = Get-AzureRmADServicePrincipal -ServicePrincipalName 3206f586-8a12-4ab1-b985-dad469c78edc -ErrorAction SilentlyContinue
    if ($servicPrincipal){break}
    start-sleep -Seconds 5    
}

if (!$servicPrincipal){"Service principal not found."}

New-AzureRmRoleAssignment -RoleDefinitionName Owner -ServicePrincipalName $azureAdApplication.ApplicationId.Guid

$tenantId = (Get-AzureRmSubscription)[0].TenantId
$applicationId = (Get-AzureRmADApplication -IdentifierUri "https://$name").ApplicationId.Guid
$thumbprint = (Get-ChildItem -Path cert:\CurrentUser\My | where Subject -eq "CN=$name" | sort NotBefore -desc)[0].Thumbprint

$command = "Connect-AzureRmAccount -ApplicationId $applicationId -CertificateThumbprint $thumbprint -ServicePrincipal -TenantId $tenantId"

$choice = read-host -Prompt "Add to profile? [Y/N]`n`n$command"

if ($choice.ToUpper() -eq 'Y')
{
    if (test-path $profile.CurrentUserAllHosts)
    {
        $currentProfile = get-content $profile.CurrentUserAllHosts
        $currentProfile | foreach {if($_ -notmatch 'Connect-AzureRmAccount'){$newProfile += "$_`r`n"}}
        rename-item $profile.CurrentUserAllHosts "$($profile.CurrentUserAllHosts).$((get-date).ticks)"
    }
    
    $newProfile += $command
    $newProfile | out-file -FilePath $profile.CurrentUserAllHosts -Force    
}