$process = Start-Process -FilePath "$env:SystemRoot\System32\cmd.exe" -ArgumentList "/c echo foo" -PassThru -Verb RunAs -Wait
