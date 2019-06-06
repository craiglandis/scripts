import-module pswindowsupdate
get-windowsupdate -Install -AcceptAll -AutoReboot -IgnoreUserInput -RecurseCycle 5 -verbose
