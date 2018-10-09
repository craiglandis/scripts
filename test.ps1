$newItem = New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name DoNotOpenServerManagerAtLogon -PropertyType DWORD -Value 1 –force

$nestedGuestVmName = 'ProblemVM'
$batchFile = "$env:allusersprofile\Microsoft\Windows\Start Menu\Programs\StartUp\RunHyperVManagerAndVMConnect.cmd"
"start $env:windir\System32\mmc.exe $env:windir\System32\virtmgmt.msc`n" | out-file -FilePath $batchFile -Force -Encoding Default
"start $env:windir\System32\vmconnect.exe localhost $nestedGuestVmName`n" | out-file -FilePath $batchFile -Append -Encoding Default
get-content $batchFile
