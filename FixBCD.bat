set Default=xxx
set DisplayFirst=xxx
REM Search each entry in bootmgr for default and display order.
for /f “tokens=1,2” %%i in (‘bcdedit /store <blahblah> /enum {bootmgr}’) do :FindDefault %%i %%j

REM See if default is set – just end
if NOT “%Default%” == “xxx” goto :eof

REM Not set – set it to first entry in displayorder
bcdedit /store <blahblah> /default %DisplayFirst%

REM Go on to use scripts that expect {default} to exist.
goto :eof

:FindDefault
if /I “%1” == “default” (
    set Default=%2
   goto :eof
)
if /I “%1” == “displayorder” (
    set DisplayFirst=%2
    goto :eof
)
goto :eof
