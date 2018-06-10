$taskName = 'sysprep'
$action = New-ScheduledTaskAction -execute 'c:\windows\system32\sysprep\sysprep.exe' -argument '/generalize /oobe /shutdown /mode:vm'
register-scheduledtask -action $action -taskname $taskName -user system -force
start-sleep -seconds 300
start-scheduledtask -TaskName $taskName
