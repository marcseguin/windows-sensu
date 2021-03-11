
function Start-Logging
{
<#
.Synopsis
    Initializes the logging facility log4n using the configuration file .log4net.xml
.Description
    Initializes the logging facility log4n
.Example
    Read-vSphere-Credentials
.Notes
    AUTHOR: Marc Seguin
.Link
#>
    try {
        $_PSScriptRoot=$PSScriptRoot # _PSScriptRoot is introduced as $PSScriptRoot cannot be debugged properly (https://github.com/PowerShell/vscode-powershell/issues/633)
        [void][Reflection.Assembly]::LoadFile("$PSScriptRoot\log4net.dll");
        [log4net.LogManager]::ResetConfiguration();
        [log4net.Config.XmlConfigurator]::Configure(("$PSScriptRoot\.log4net.xml"))
        $Global:Log4n=[log4net.LogManager]::GetLogger("root")
        $Log4n.Debug("***** Starting Logging *****")

        $Log4n.Debug("Module Name: $($MyInvocation.ScriptName)")
        $Log4n.Debug("Module Path: $PSScriptRoot")
    }
    catch {
        throw "Cannot initialize logging faciliy."
    }
}

function Stop-Logging
{
<#
.Synopsis
    Stops the logging facility log4n
.Description
    Stops the logging facility log4n
.Example
    Stop-Logging
.Notes
    AUTHOR: Marc Seguin
.Link
#>
    try {
        $Log4n.Debug("***** Stopping Logging *****")
        Remove-Variable -name "Log4n" -Scope "Global"
        [log4net.LogManager]::ResetConfiguration();
    }
    catch {
        throw "Cannot initialize logging faciliy."
    }
}


function Read-Credentials
{
<#
.Synopsis
    Reads credentials from a file
.Description
    Reads credentials from <filename> (if exists) or asks and stores them in <filname> in an encrypted format (if the file does not exist)
.Example
    Read-Credentials
.Notes
    AUTHOR: Marc Seguin
.Link
#>
Param(
    [Parameter(Mandatory=$true)] [string] $filename) 

    try{
    if (test-path -Path  $filename -PathType Leaf) {
        $log4n.info("Loading Credential file '$filename'")
        $credential=Import-Clixml $filename
    }
    else {
        $log4n.info("Quering Credentials...")
        Write-Host "Die Anmeldeinformationsdatei " -NoNewline -ForegroundColor Yellow  
        Write-Host $filename  -NoNewline 
        Write-Host " existiert nicht. Geben Sie die Anmeldeinformationen manuell ein." -ForegroundColor Yellow  

        $credential=Get-Credential
        $decision = $Host.UI.PromptForChoice("Anmeldeinformationsdatei '$filename' speichern", "Sollen die Zugangsdaten fuer spaetere Zugriffe gespeichert werden (Das Kennwort wird verschluesselt gespeichert)?", @('&Ja', '&Nein'), 0)
        if ($decision -eq 0){
        $log4n.info("Saving Credential file $filename")
        Export-Clixml -Path $filename -InputObject $credential
        }
        else {
        $log4n.info("Credential file '$filename' won't be saved.")
        }
    }
    return $credential
    }
    catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw
    }
}
