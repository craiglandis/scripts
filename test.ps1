$batchFile = "$env:allusersprofile\Microsoft\Windows\Start Menu\Programs\StartUp\RunHyperVManagerAndVMConnect.cmd"
"start $env:windir\System32\mmc.exe $env:windir\System32\virtmgmt.msc" | out-file -FilePath $batchFile -Append -Encoding Default
"start $env:windir\System32\vmconnect.exe localhost $nestedGuestVmName" | out-file -FilePath $batchFile -Append -Encoding Default
