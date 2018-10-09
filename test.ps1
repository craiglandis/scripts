$newItem = New-ItemProperty -Path HKCU:\Software\Microsoft\ServerManager -Name DoNotOpenServerManagerAtLogon -PropertyType DWORD -Value 1 –force
$nestedGuestVmName = 'ProblemVM'
$batchFileContents = @"
start $env:windir\System32\mmc.exe $env:windir\System32\virtmgmt.msc
start $env:windir\System32\vmconnect.exe localhost $nestedGuestVmName
"@
$batchFile = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\RunHyperVManagerAndVMConnect.cmd'
$batchFileContents | out-file -FilePath $batchFile -Force -Encoding Default
get-content $batchFile
