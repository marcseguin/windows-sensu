#trap {
#  $Log4n.Error("$_`r`n$($_.InvocationInfo.PositionMessage)")
#  exit  
#}


if ($null -ne (get-module|where-object {$_.name -eq "support"})){
  remove-module support
}
Import-Module ./support.psm1

#Add-ESXEmergencyUser -VMHostname "srv-esxi-0101.platform.klarios.net" 
#Set-ESXStorageMultiPathingPolicy -VMHostname "srv-esxi-0251xx.platform.klarios.net"
#Set-ESXScratchLocation -VMHostName "srv-esxi-02xx.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-01"
Set-ESXScratchLocation -VMHostName "srv-esxi-0251.platform.klarios.net" -DatastoreName "TEMP"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0102.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-01"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0111.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-01"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0112.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-01"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0151.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-01"

#Set-VMHostScratchLocation -VMHostName "srv-esxi-0201.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-02"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0202.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-02"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0211.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-02"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0212.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-02"
#Set-VMHostScratchLocation -VMHostName "srv-esxi-0251.platform.klarios.net" -DatastoreName "BUS-VSPHERE-LOGS-02"

  
  #Export-vSphereRoles -Path "c:\temp" -All
  #.viDisconnect-vCenter

