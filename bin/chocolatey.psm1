

function Test-ChocolateyIsInstalled {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.PSSession] $Session
    )
    Invoke-Command -Session $session -ScriptBlock {$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")} -ErrorVariable ReturnError -OutVariable result |Out-Null
    Invoke-Command -Session $Session -ScriptBlock {choco -v} -ErrorVariable ReturnError -OutVariable result |Out-Null
    If ($null -eq $ReturnError -or $ReturnError.Count -ne 0) {
        $log4n.Warn("Chocolatey is not installed on \\$($session.ComputerName).")
        return $false
    }
    else{
        $log4n.info("Chocolatey version $result is installed on \\$($session.ComputerName).")
        return $true
    }
}

function Install-Chocolatey {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.PSSession] $Session
    )

    $log4n.Info("Installing Chocolatey on \\$($session.ComputerName).")
    Invoke-Command -Session $session -ScriptBlock {Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))} -ErrorVariable ReturnError -OutVariable result |Out-Null
    Invoke-Command -Session $session -ScriptBlock {$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine")} -ErrorVariable ReturnError -OutVariable result |Out-Null
    Invoke-Command -Session $Session -ScriptBlock {choco -v} -ErrorVariable ReturnError -OutVariable result |Out-Null
    if ($null -eq $ReturnError -or $ReturnError.Count -ne 0) {
        $log4n.Error("Failed to install Chocolatey on \\$($session.ComputerName)!");
        throw $ReturnError
    }
    $log4n.info("Chocolatey version $result is now installed on \\$($session.ComputerName).")
    return $true
}

if ($null -eq (Get-Module -name support)){
    Import-Module $PSScriptRoot\support.psm1
}
