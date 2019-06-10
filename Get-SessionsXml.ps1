param(
    $path = "$env:windir\servicing\sessions\sessions.xml"
)

$win32_OperatingSystem = Get-WmiObject -class Win32_OperatingSystem
$version = $win32_OperatingSystem.Version
$caption = $win32_OperatingSystem.Caption
if ($caption -match 'Server')
{
    $caption = 'Windows Server'
}
elseif ($caption -match 'Windows 10')
{
    $caption = 'WIN10'
}
if ($version.StartsWith('10'))
{
    $releaseId = (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name ReleaseID).ReleaseId
}
switch -regex ($version) {
#	 '7601' {if($osVersion -match 'windows server'){$os = 'R208'}else{$os =  'WIN7'}}
#	 '9200' {if($osVersion -match 'windows server'){$os = 'WS12'}else{$os =  'WIN8'}}
#	 '9600' {if($osVersion -match 'windows server'){$os = 'R212'}else{$os = 'WIN81'}}
    '10240' {$codeName = 'TH1'}
    '10586' {$codeName = 'TH2'}
    '14393' {$codeName = 'RS1'}
    '15063' {$codeName = 'RS2'}
    '16299' {$codeName = 'RS3'}
    '17134' {$codeName = 'RS4'}
    '17763' {$codeName = 'RS5'}
    '18362' {$codeName = '19H1'}
}
$osVersion = "$caption $codeName $releaseId $version"
$psversion = "$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"
"`n$osVersion PS $psversion"

$installedPackages =  New-Object System.Collections.ArrayList

$sessions = [xml](get-content $path)
$sessions = $sessions.sessions.session

$sessions | foreach {
    $session = $_
    if ($session.lastSuccessfulState -eq 'Complete' -and $session.status -eq '0x0')
    {
        $tasks = $session.Tasks
        $actions = $session.Actions
        $tasks | foreach {
            $task = $_
            $packages = $task.phase.package
            $packages | foreach {
                $package = $_
                if ($package.targetState -eq 'Installed')
                {
                    $installedPackage = [PSCustomObject]@{
                        Name = $package.name
                        Installed = $actions.phase.Installed
                    }
                    [void]$installedPackages.Add($installedPackage)
                }
            }
        }
    }
}

$installedPackages | sort Installed -Descending

<#

= $sessions.sessions.session | where {$_.lastSuccessfulState -eq 'Complete' -and $_.status -eq '0x0'}

$sessions = $sessions | where {$_.tasks.phase.package.name}

$sessions = $sessions | where {$_.tasks.phase.package.name.Startswith('KB')}

# Of those successfully completed sessions, get the sessions where package name starts with "KB" and targetState is "Installed"
$sessions | foreach {
    if ($_.tasks.phase.package.name.Startswith('KB') -and $_.tasks.phase.package.targetState -eq 'Installed')
    {
        $name = $_.tasks.phase.package.name
        $installed = $_.Actions.Phase.Installed
        write-host $name $installed
    }
}

#>
