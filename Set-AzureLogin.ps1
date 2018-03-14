Add-AzureAccount

$now = get-date
Get-AzurePublishSettingsfile

do
{
    start-sleep -Seconds 10
}
while (!((dir "$env:userprofile\downloads\*.publishsettings" | where LastWriteTime -gt $now)))

Import-AzurePublishSettingsFile -PublishSettingsFile (dir "$env:userprofile\downloads\*.publishsettings" | sort LastWriteTime -desc)[0].FullName