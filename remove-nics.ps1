If (!(Test-Path "$Env:SystemDrive\Temp"))
{
      New-Item "$Env:SystemDrive\Temp" -type directory
}
 
$LogFile = "$Env:SystemDrive\Temp\FixAzureVM-RemoveNICs.log"
 
If (!(Test-Path $LogFile))
{
      " ####################" | Out-File $LogFile -Append
      $Stamp = Get-Date
      " Start time: $Stamp" | Out-File $LogFile -Append
      " ####################" | Out-File $LogFile -Append
 
      If (Test-Path "$Env:WinDir\System32\DevCon.exe")
      {
            #list NICs
            " Listing NICs on this system..." | Out-File $LogFile -Append
            [array] $AllNICs = Invoke-Expression "$Env:WinDir\System32\DevCon.exe findall 'VMBUS\{F8615163-DF3E-46C5-913F-F2D2F965ED0E}*'"
            $AllNICs | Out-File $LogFile -Append
            #remove the last element
            $AllNICs = $AllNICs[0..($($AllNICs.Count)-2)]
            $AllNICs | %{$NICProps = New-Object -TypeName PSObject;Add-Member -InputObject $NICProps -MemberType NoteProperty -Name "Device" -Value $_.Split(":")[0].Trim();Add-Member -InputObject $NICProps -MemberType NoteProperty -Name "Name" -Value $_.Split(":")[1].Trim();[array] $AllNICsParsed += $NICProps}
            [array] $AllNICsSorted = $AllNICsParsed | Sort-Object -Property Name
            ForEach ($NIC in $AllNICsSorted)
            {
                  $NICName = $NIC.Name
                  $NICDevice = $NIC.Device
                  $NICDevice = "@" + $NICDevice
                  " --------------------" | Out-File $LogFile -Append
                  " Removing $NICName" | Out-File $LogFile -Append
                  " Device: $NICDevice" | Out-File $LogFile -Append
                  $RemoveResult = Invoke-Expression "$Env:WinDir\System32\DevCon.exe remove '$NICDevice'"
                  $RemoveResult | Out-File $LogFile -Append
                  " --------------------" | Out-File $LogFile -Append
            }
      }
      Else
      {
            #devcon.exe not found
            " DevCon.exe was not found in System32. Exiting..." | Out-File $LogFile -Append
      }
      " ####################" | Out-File $LogFile -Append
      $Stamp = Get-Date
      " End time: $Stamp" | Out-File $LogFile -Append
      " ####################" | Out-File $LogFile -Append
}
Else
{
      " ####################" | Out-File $LogFile -Append
      $Stamp = Get-Date
      " Start time: $Stamp" | Out-File $LogFile -Append
      " ####################" | Out-File $LogFile -Append
      " This script has already been executed. Doing nothing. Please remove this from running as a startup script." | Out-File $LogFile -Append
      " ####################" | Out-File $LogFile -Append
      $Stamp = Get-Date
      " End time: $Stamp" | Out-File $LogFile -Append
      " ####################" | Out-File $LogFile -Append
}