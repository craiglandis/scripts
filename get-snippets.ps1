<#
.SYNOPSIS
    Outputs a flat list showing both the PowerShell snippets included with the PowerShell extension for VSCode which it shows a Type of Builtin, as well as any user-created snippets which it shows as Type of User.

.DESCRIPTION

    PowerShell snippets include with the PowerShell extension for VSCode reside in a PowerShell.json file:
    
    C:\Users\<username>\.vscode\extensions\ms-vscode.powershell-<version>\snippets\PowerShell.json

    You can also see them in the vscode-powershell GitHub repro:

    https://github.com/PowerShell/vscode-powershell/blob/master/snippets/PowerShell.json

    Users can add their own to this PowerShell.json file (file doesn't exist by default):

    C:\Users\<username>\AppData\Roaming\Code\User\Snippets\PowerShell.json

    The Settings Sync extension is one way to keep snippets synced between your machines using a GitHub gist.

    https://marketplace.visualstudio.com/items?itemName=Shan.code-settings-sync

#>

$builtinSnippetsFilePath = (get-childitem "$env:userprofile\.vscode\extensions\ms-vscode.powershell-*\snippets\PowerShell.json").FullName
$builtinSnippets = convertfrom-json (get-content $snippetsFilePath -raw)

$userSnippetsFilePath = (get-childitem "$env:appdata\Code\User\Snippets\PowerShell.json").FullName
$userSnippets = convertfrom-json (get-content $userSnippetsFilePath -raw)

$snippets = @()

if ($builtinSnippets)
{
    $builtinSnippets.psobject.properties | foreach {

        $snippet = [ordered]@{        
            'Prefix' = $_.value.prefix
            'Name' = $_.Name
            'Type' = 'Builtin'
            'Description' = $_.value.Description
        }
        
        $snippets += New-Object -TypeName PSObject -Property $snippet    
    
    }
}

if ($userSnippets)
{
    $userSnippets.psobject.properties | foreach {

        $snippet = [ordered]@{
            'Prefix' = $_.value.prefix
            'Name' = $_.Name
            'Type' = 'User'
            'Description' = $_.value.Description
        }
        
        $snippets += New-Object -TypeName PSObject -Property $snippet    

    }
}

$snippets | sort Prefix | format-table Prefix,Name,Type,Description -AutoSize