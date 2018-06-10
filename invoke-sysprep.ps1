$path = "c:\scripts"
$filePath = "$scriptPath\invoke-sysprep.ps1"
new-item -Path $path -ItemType Directory
'start-sleep -seconds 300' | out-file -filepath $filePath -append
'c:\windows\system32\sysprep\sysprep.exe /generalize /oobe /shutdown /mode:vm' | out-file -filepath $filePath -append
$taskName = 'sysprep'
$action = New-ScheduledTaskAction -execute 'powershell.exe' -argument "-executionpolicy bypass -file $filePath"
register-scheduledtask -action $action -taskname $taskName -user system -force
start-scheduledtask -TaskName $taskName
