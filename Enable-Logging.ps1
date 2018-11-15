$logName = 'Microsoft-Windows-DNS-Client/Operational'
$log = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration $logName
$log.IsEnabled=$true
$log.SaveChanges()