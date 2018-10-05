param([string]$regFilePath)

$regloadlib = @"
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using System;
using System.Runtime.InteropServices;

namespace RegLoadLib
{

    public static class PrivateRegistry
    {
        [Flags]
        public enum RegSAM
        {
            AllAccess = 0x000f003f,
            Read = 0x20019,
            Write = 0x20006,
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int RegLoadAppKey(String hiveFile, out int hKey, RegSAM samDesired, int options, int reserved);
    }
}
"@

$RegLoad = Add-Type $regloadlib -PassThru

$regHandle = 0
$result = [RegLoadLib.PrivateRegistry]::RegLoadAppKey($regFilePath, [ref]$regHandle, [RegLoadLib.PrivateRegistry+RegSAM]::Read, 0, 0)

if ($result -eq 0)
{
    $safeRegistryHandle = [Microsoft.Win32.SafeHandles.SafeRegistryHandle]::new([IntPtr]::new($regHandle), $true)
    $appKey = [Microsoft.Win32.RegistryKey]::FromHandle($safeRegistryHandle)
    $appKey.GetSubKeyNames()
}
else
{
    Write-Error ('Unable to load the file {0} - Error code: {1}' -f $regFilePath, $result)
}
