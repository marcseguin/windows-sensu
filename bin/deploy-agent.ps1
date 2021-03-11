
function Install-SensuAgent {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.PSSession] $Session
    )
    $log4n.Info("Installing Sensu-Agent on \\$($session.ComputerName) using Chocolatey.")
    invoke-command -Session $session -ScriptBlock {choco install sensu-agent -y|out-null;$lastexitcode} -ErrorVariable ReturnError -OutVariable result
    if ($result -eq 0){
        Test-SensuAgentInstallation -session $session
    }
}

function Test-SensuAgentIsInstalled {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.PSSession] $Session
    )
    invoke-command -Session $session -ScriptBlock {(Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Sensu Agent" })} -ErrorVariable ReturnError -OutVariable result |Out-Null
    if (!$result) {
        $log4n.warn("The Sensu-Agent is not installed on \\$($session.ComputerName)!")
        return $false
    }
    elseif ($result.count -ne 1) {
        $log4n.Fatal("Ups - Looks like Sensu-Agent is installed multiple times on \\$($session.ComputerName)!")
        throw
    }
    else {
        $log4n.info("The $($result.displayname) version $($result.displayversion) is installed on \\$($session.ComputerName).")
        return $true
    }
}

function Test-SensuAgentServiceIsInstalled {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.PSSession] $Session
    )
    Invoke-Command -Session $session -ScriptBlock {Get-Service -name SensuAgent -ErrorAction SilentlyContinue } -ErrorVariable ReturnError -OutVariable result
    if ($result.count -ne 0){
        $log4n.info("The $($result.displayname) service on \\$($session.computername) is installed and its status is $($result.status.tostring().tolower()).")
        return $true
    }
    else {
        $log4n.warn("The Sensu Agent as a service is not installed on \\$($session.computername)!")
        return $false
    }
}

function Install-SensuAgentService {
    param(
        [Parameter(Mandatory)][System.Management.Automation.Runspaces.PSSession] $Session
    )
    $log4n.info("Installing Sensu Agent as a Windows service on \\$($Session.ComputerName).")
    invoke-command -Session $session -ScriptBlock {(Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -eq "Sensu Agent" })} -ErrorVariable ReturnError -OutVariable result |Out-Null
    if (!$result) {
        $log4n.fatal("Ups - Sensu-Agent application cannot be found!")
        throw
    }
    elseif ($result.count -ne 1) {
        $log4n.fatal("Ups - Looks like multiple Sensu Agents are installed, 'don't know which to use further!")
        throw
    }
    $SensuAgentExe=$result.InstallLocation+"bin\sensu-agent.exe"
    if ((invoke-command -Session $session -ScriptBlock {Test-Path -Path $using:SensuAgentExe} -ErrorVariable ReturnError -OutVariable result) -eq $false)
    {
        $log4n.fatal("Ups - $SensuAgentExe on \\$($session.ComputerName) cannot be found!")
        throw
    }
    $log4n.info("Copying the configuration files to \\$($session.ComputerName).")
    try{
        copy-item -path "$PSScriptRoot\..\config-agent\agent.yml" -Destination "C:\ProgramData\Sensu\config\agent.yml" -ToSession $Session
        copy-item -path "$PSScriptRoot\..\config-agent\srv-mon01.ks.netoffice-kassel.de-cert.pem" -Destination "C:\ProgramData\Sensu\config\srv-mon01.ks.netoffice-kassel.de-cert.pem" -ToSession $Session
    }
    catch {
        $log4n.error("Something went wrong when copying the configuration files to \\$($session.ComputerName). We still continue with the service installation but you need to check the configuration files locally on the remote computer.")
    }
    $log4n.info("Installing the Sensu Agent as a Windows service \\$($session.ComputerName) and start it.")
    invoke-command -session $session -scriptblock {& 'C:\Program Files\Sensu\sensu-agent\bin\sensu-agent.exe' service install 2>&1 |Out-Null;$lastexitcode} -ErrorVariable ReturnError -OutVariable result
    if ($result -eq 0){
        $log4n.Info("Sensu Agent is successfully installed on \\$($session.ComputerName).")
    else
        $log4n.Error("Failed to install Sensu Agent on \\$($session.ComputerName). You have to check on the remote computer.")
    }
    Write-Host $result
}



#region main
#region load external modules and init
if ($null -ne (Get-Module -name support)){
    Remove-Module -Name support
}
Import-Module $PSScriptRoot\support.psm1
if ($null -ne (Get-Module -name chocolatey)){
    Remove-Module -Name chocolatey
}
Import-Module $PSScriptRoot\chocolatey.psm1
Start-Logging
$error.Clear()
#endregion 

$Computername="srv-mst01"
$credential=Read-Credentials -filename "$env:userprofile\.netoffice.pscredential"
try{
    $Session = New-PSSession -ComputerName $Computername -Credential $credential
}
catch{
    $log4n.Fatal("Can't connect to computer $computername!")
    throw
}

# What about Chocolatey?
if ((Test-ChocolateyIsInstalled -session $Session) -eq $false){
    Install-Chocolatey -session $Session
}

# OK, now let's install & configure the Sensu-Agent using Chocolatey
if ((Test-SensuAgentIsInstalled -session $session) -eq $false){
    Install-SensuAgent -session $session
}

if ((Test-SensuAgentServiceIsInstalled -session $session) -eq $false)
{
    Install-SensuAgentService -session $session
}

#region cleaning up
Remove-PSSession -Session $session
Stop-Logging
if ($null -ne (Get-Module -name chocolatey)){
    Remove-Module -Name chocolatey
}
if ($null -ne (Get-Module -name support)){
    Remove-Module -Name support
}
#endregion
#endregion