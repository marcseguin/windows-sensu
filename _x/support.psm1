 #region Script Diagnostic Functions
function Get-ScriptLineNumber { return $MyInvocation.ScriptLineNumber }
function Get-ScriptName { return $MyInvocation.ScriptName }

new-item alias:__LINE__ -value Get-ScriptLineNumber
new-item alias:__FILE__ -value Get-ScriptName   

#endregion

function Read-Credentials{
<#
.Synopsis
  Reads credentials from a file
.Description
  Reads credentials from %userprofile%\<filename> (if exists) or asks and stores them in an encrypted format
.Example
  Read-vSphere-Credentials
.Notes
  AUTHOR: Marc Seguin {NET.Office GmbH for ATIS systems GmbH}
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

function Add-PlatformRole{
<#
.Synopsis
  Creates a group
.Description
  Creates a group if it does not exist
.Example
  Add-PlattformRole -GroupName -GroupDescription -OU
.Notes
  AUTHOR: Marc Seguin {NET.Office GmbH for ATIS systems GmbH}
.Link
#>
Param(
  [Parameter(Mandatory=$true)] [string] $GroupName,
  [Parameter(Mandatory=$true)] [string] $GroupDescription)
  
  try {
    if (($null -eq $global:ADcredential) -or ($global:ADcredential.GetType().fullname -ne "System.Management.Automation.PSCredential")){
      $global:ADcredential=Read-Credentials -filename "$env:userprofile\.ActiveDirectory.pscredential"
    }
    $Group=Get-ADGroup -Server $global:KlariosConfig.System.DomainController -Credential $global:ADcredential -LDAPFilter "(SAMAccountName=$GroupName)"
    if ($null -eq $group){
      $group=New-ADGroup -Server $global:KlariosConfig.System.DomainController -Credential $global:ADcredential `
      -Name $GroupName `
      -GroupScope Global `
      -DisplayName $GroupName `
      -Path $global:KlariosConfig.System.roleOU `
      -Description $GroupDescription
      if ($null -eq $group) {
        Throw "Cannot create group $GroupName"
      }
    }
    else {
      $GroupOU=$Group.DistinguishedName.substring($group.name.length+4).tolower()
      if ($GroupOU -eq $roleOU.ToLower()) {
        Write-Host "Warning: Group $GroupName already exists." -ForegroundColor Yellow
      }
      else {
        Write-Host "Warning: Group $GroupName already exists but in wrong OU $GroupOU." -ForegroundColor Yellow
      }
    }
    return $Group
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw
  }
}  
function Set-ESXAdvancedSettings{
<#
.Synopsis
  This function sets a set of advanced settings of a VMHost to predefined values.
.Description
  This function sets a set of advanced settings of a VMHost to predefined values.
.Example
  Set-ESXAdvancedSettings -VMHost host.vsphere.local
.Notes
  AUTHOR: Marc Seguin {NET.Office GmbH for ATIS systems GmbH}
.Link
#>
Param(
  [Parameter(Mandatory=$true)] [string] $VMHostname)

  try {
    if (!$global:DefaultVIServer.IsConnected) {
      $log4n.Error("I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!")
      Throw "I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!"
    } 

    $VMHost=Get-VMHost -Name $VMHostname
    if ($null -eq $VMHost) {
      $log4n.Error("Cannot find ESXi-Host with hostname '$VMHostname'!")
      Throw "Cannot find ESXi-Host with hostname '$VMHostname'!"
    }
    $klariosconfig.vSphere.Hosts.AdvancedSettings|ForEach-Object {
      $item=$_
      Write-Host "Verarbeite vSphere.Host.AdvancedSettings '$($item.Name)'"
      if (Invoke-Expression $item.Filter) {
        Write-Host "Filter '$($item.Filter)' ist WAHR - Einstellungen werden verarbeitet."
        $item.Item|ForEach-Object {
          $AdvancedSetting=Get-AdvancedSetting -Entity $VMHost -Name $_.name
          if ($AdvancedSetting.Value -ne $_.Value){
            Write-Host "Aktualisiere '$($_.Name)'."
            $AdvancedSetting|Set-AdvancedSetting -Value $_.value -Confirm:$false
          }
        }
      }
      else {
        Write-Host "Filter '$($item.Filter)' ist FALSCH - Einstellungen werden nicht verarbeitet."
      }
    }
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw
  }
}

function Set-ESXScratchLocation{
<#
.Synopsis
  This function sets the scratch location of a VMhost to a specific datastore.
.Description
  This function sets the scratch location of a VMhost to a specific datastore.
.Example
  Set-VMHostScratchLocation -VMHostname "host.vsphere.local" -DataStoreName "MyDataStore"
.Notes
  AUTHOR: Marc Seguin {NET.Office GmbH for ATIS systems GmbH}
.Link
#>
Param(
  [Parameter(Mandatory=$true)] [string] $VMHostname,
  [Parameter(Mandatory=$true)] [string] $DataStoreName)

  try{
    $VMHostname=$VMHostname.ToLower()
    # some checks before executing anything...
    if (!$global:DefaultVIServer.IsConnected) {
      $log4n.Error("I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!")
      Throw "I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!"
    } 

    $VMHost=Get-VMHost -Name $VMHostname
    if ($null -eq $VMHost) {
      Throw "Cannot find VM-Host with hostname '$VMHostname'!"
    }
    
    #prepares datastore
    $DataStore=Get-Datastore -Name $DataStoreName
    if ($null -eq $DataStore) {
      Throw "Cannot access DataStore '$DataStoreName'!"
    }        
    New-PSDrive -Location $datastore -Name DS -PSProvider VimDatastore -Root "/"
    if ((Test-Path "DS:/Scratch/$VMHostname") -eq $false) {
      $log4n.Info("Creating directory [$DataStoreName]/Scratch/$vmhostname'")
      New-Item -Path "DS:/Scratch/$VMHostname" -ItemType Directory
    }
    Remove-PSDrive -Name DS -Confirm:$false

    #set ScratchConfig
    $ScratchPath="$($datastore.name)/Scratch/$($VMHostname)"
    Get-AdvancedSetting -Entity $VMHost -Name "ScratchConfig.ConfiguredScratchLocation" | Set-AdvancedSetting -Value "/vmfs/volumes/$ScratchPath" -Confirm:$false

    #Check conditions for rebooting the ESXi-Host
    if ($vmhost.ConnectionState -ne "Maintenance") {
      $VMtotal=($vmhost |get-vm).count
      $VMrunning=($vmhost |get-vm |Where-Object {$_.PowerState -eq "PoweredOn"}).count
      Write-Host "Der ESXi-Host '$vmhostname' befindet sich nicht im Maintenance Modus!" -ForegroundColor Yellow
      Write-Host "Es sind $VMtotal VMs registiert, davon sind $VMrunning VMs eingeschaltet."
      $reboot= ($Host.UI.PromptForChoice("Reboot erzwingen?", "Soll der Server trotzdem neu gestartet werden?", @('&Ja', '&Nein'), 1) -eq 0)
      if ($reboot) {
        $log4n.Info("Der ESXi-Host '$vmhostname' befindet sich nicht im Maintenance Modus! Der Benutzer hat trotzdem entschieden den Host neu zu starten.")
      }
    }
    else {
      $reboot=$true
    }

    #reboot the ESXi-Host
    if ($reboot -eq $true){
      $null=Restart-VMHost -VMHost $VMHost -Force -RunAsync -Confirm:$false
      $log4n.Info("Der ESXi-Host wird jetzt neu gestartet.")
      Write-Host "Der ESXi-Host wird jetzt neu gestartet. Nach dem Neustart müssen Sie ggf. den Wartungsmodus manuell beenden." -ForegroundColor Yellow
    }
    else {
      log4n.Info("Der ESXi-Host wird nicht neu gestartet. Bitte fuehren Sie die Aktion manuell durch.")
      Write-Warning "Der ESXi-Host wird nicht neu gestartet. Bitte fuehren Sie die Aktion manuell durch." -ForegroundColor Yellow
    }
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw
  }
}

function Import-vSphereRoles{
<#   
.Synopsis   
  Imports vSphere roles from .role-files.
.Description   
  This script imports custom created roles.
.Example   
  Import-vSphereRoles -ImportPath c:\temp
  Imports Roles to vSphere.
.Notes  
  AUTHOR: Marc Seguin {NET.Office GmbH}
#>
  
Param(   
    [Parameter(Mandatory=$true)] [AllowEmptyString ()] [string] $ImportPath)

  try {
    if (!$global:DefaultVIServer.IsConnected){
      $log4n.Error("I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!")
      Throw "I'm currently not connected to a vCenter! You must successfully call 'Connect-VIServer' first!"
    } 

    $ExistingRoles=Get-VIRole | Select-Object -ExpandProperty Name
    $RoleFiles=Get-Item -path "$ImportPath\*.role"
    foreach ($rolefile in $roleFiles){
      $NewRoleName = $rolefile.BaseName
      $RolesPrivileges = Get-Content -Path $rolefile.FullName
      if ($newRoleName -notin $ExistingRoles){
        $NewRole=New-Virole -Name $NewRoleName
        Write-Host "Created Role $NewRoleName" -BackgroundColor DarkGreen  
        foreach ($Privilege in $RolesPrivileges){
          if (-not($null -eq $privilege -or $privilage -eq "")){
            Write-Host "Setting Permissions '$Privilege' on Role '$NewRoleName'" -ForegroundColor Yellow
            Set-VIRole -Role $NewRoleName -AddPrivilege (Get-VIPrivilege -ID $privilege) | Out-Null
          }
        }
      }
      else {
        Write-Warning "Role '$NewRoleName' already exists - import skipped"  
      }
    }
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw
  }
}
function Export-vSphereRoles {
<#
.Synopsis
  Exports vsphere roles to text file extension roles.
.Description
  This script exports only the custom created roles by users
.Example
  Export-vSphereRoles -Path c:\temp
  Exports Roles to the folder.
.Notes
  NAME: Export-vSphereRoles
  AUTHOR: Marc Seguin {NET.Office GmbH}
  KEYWORDS: Export vSphere Roles

.Link
#>
Param(
  [Parameter(Mandatory=$true)] [AllowEmptyString ()] [string] $Path,
  [switch]$All = $false)

  try {
    $DefaultRoles = "com.vmware.Content.Admin","vSphere Client Solution User","InventoryService.Tagging.TaggingAdmin","AutoUpdateUser","VirtualMachineConsoleUser","Virtual machine power user (sample)","NoCryptoAdmin","NoAccess", "Anonymous", "View", "ReadOnly", "Admin", "VirtualMachinePowerUser", "VirtualMachineUser", "ResourcePoolAdministrator", "VMwareConsolidatedBackupUser", "DatastoreConsumer", "NetworkConsumer"
    if (!$global:DefaultVIServer.IsConnected) {
      Throw "I'm currently not connected to a vCenter! You must successfully call 'Connect-VIServer' first!"
    } 

    $AllVIRoles = Get-VIRole
    foreach ($role in $AllVIRoles){
      if ($role.name -notin $DefaultRoles -or $all){
        $completePath = Join-Path -Path $Path -ChildPath "$role.role"
        $log4n.Info("Exporting Role '$role' to $completePath")
        $priv=Get-VIPrivilege -Role $Role 
        $priv | select-object -ExpandProperty Id | Out-File -FilePath $completePath
        Remove-Variable priv
      }
    }
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw 
  }
}

function Add-ESXEmergencyUser{
<#
.Synopsis
  This function sets the scratch location of a VMhost to a specific datastore.
.Description
  This function sets the scratch location of a VMhost to a specific datastore.
.Example
  ESXEmergencyUser -VMHostname "host.vsphere.local" 
.Notes
  AUTHOR: Marc Seguin {NET.Office GmbH for ATIS systems GmbH}
.Link
#>
Param(
  [Parameter(Mandatory=$true)] [string] $VMHostname)

  try {
    if (!$global:DefaultVIServer.IsConnected) {
      Throw "I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!"
    } 
    $ESXcli=Get-ESXcli -VMHost $VMHostname -V2
    if ($null -eq $ESXcli) {
      Throw "I can't connect to VMHost $VMHostname. Please check hostname and server availibility!"
    }

    $EmergencyUser=Read-Credentials -Filename "$env:userprofile\.EmergencyUser.pscredential"
    if ($null -eq ($esxcli.system.account.list.invoke()|where-object {$_.userid -eq $EmergencyUser.username})) {
      $accountArgs = $ESXcli.system.account.add.CreateArgs()
      $accountArgs.id=$EmergencyUser.username
      $accountArgs.description = "Notfallkonto"
      $accountArgs.password=$EmergencyUser.GetNetworkCredential().Password
      $accountArgs.passwordconfirmation=$EmergencyUser.GetNetworkCredential().Password
      $esxcli.system.account.add.Invoke($accountArgs)
    }

    $permissionArgs = $esxcli.system.permission.set.CreateArgs()
    $permissionArgs.id = $EmergencyUser.username
    $permissionArgs.group = $false
    $permissionArgs.role = "Admin"
    if ($esxcli.system.permission.set.Invoke($permissionArgs) -eq $false) {
      Throw "I can't assign role 'admin' to user!"
      return $false
    }
    else {
      $log4n.Info("Das Notfallkonto '$($EmergencyUser.username)' ist angelegt und der Rolle 'Adminstrators' zugeordnet.")
      return $true
    }
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw 
  }
}


function Set-ESXStorageMultiPathingPolicy{
<#
.Description
  The function sets the default pathing policy for new/existing LUNs based on the configuration file. The host must be rebooted before the settings are getting activated!
.Example
  ESXStorageMultiPathingPolicy -VMHostname "host.vsphere.local" 
.Notes
  AUTHOR: Marc Seguin {NET.Office GmbH for ATIS systems GmbH}
.Link
#>
Param(
  [Parameter(Mandatory=$true)] [string] $VMHostname,
  [Parameter(Mandatory=$false)] [bool] $Reboot =$false)

  try{
    $ErrorActionPreference='Stop'
    if (!$global:DefaultVIServer.IsConnected){
      Throw "I'm currently not connected to a vCenter. You must successfully call 'Connect-VIServer' first!"
    } 
    $ESXcli=Get-ESXcli -VMHost $VMHostname -V2
    if ($null -eq $ESXcli) {
      Throw "I can't connect to VMHost $VMHostname. Please check hostname as well as host availibility!"
    }

    $KlariosConfig.vSphere.Hosts.DefaultPathingPolicies|ForEach-Object {
      $result=$esxcli.storage.nmp.satp.set.invoke(@{boot=$false;defaultpsp=$_.PSP;satp=$_.SATP})
      $log4n.info("Host $VMHostname : $result")
    }
  }
  catch {
    $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
    throw 
  }
}

#region Script MAIN

try {
  $MyScriptRoot=$PSScriptRoot # MyScriptRoot is introduced as $PSScriptRoot cannot be debugged properly (https://github.com/PowerShell/vscode-powershell/issues/633)
  [void][Reflection.Assembly]::LoadFile("$MyScriptRoot\log4net.dll");
  [log4net.LogManager]::ResetConfiguration();
  [log4net.Config.XmlConfigurator]::Configure(("$MyScriptRoot\.log4net.configuration.xml"))
  $Global:Log4n=[log4net.LogManager]::GetLogger("root")
  $Log4n.Debug("Module Name: $($MyInvocation.MyCommand)")
  $Log4n.Debug("Module Path: $MyScriptRoot")
}
catch {
  throw "Cannot initialize logging faciliy."
}

try{
  $Script:KlariosConfigFile="$MyScriptRoot\.klarios.configuration.json"
  if (test-path -Path  $KlariosConfigFile -PathType Leaf) {
    $global:KlariosConfig=Get-Content $KlariosConfigFile|ConvertFrom-Json
    if ($null -ne $global:DefaultVIServer){
      Disconnect-VIServer -Server * -Force -Confirm:$false
    }
    $Log4n.Info("Connecting to vCenter $($KlariosConfig.global.vcenter)...")
    Connect-VIServer -Server $KlariosConfig.global.vcenter -Credential (Read-Credentials -filename "$env:userprofile\.vsphere.pscredential")
    if ($global:defaultviserver.isconnected){
      $Log4n.Info("Verbindung hergestellt: $($global:defaultviserver.user)@$($global:defaultviserver.name):$($global:defaultviserver.port)")
    }
  }
  else { 
    throw "Die Konfigurationsdatei '$KlariosConfigFile' konnte nicht gefunden werden." }
}
catch {
  $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
  throw
}
#endregion