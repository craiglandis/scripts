$taskName = 'sysprep'
$action = New-ScheduledTaskAction -execute 'c:\windows\system32\sysprep\sysprep.exe' -argument '/generalize /oobe /shutdown /mode:vm'
register-scheduledtask -action $action -taskname $taskName -force
start-scheduledtask -TaskName $taskName
