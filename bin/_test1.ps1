$computername="localhost"
$IcingaSession="Incinga-Core"

$credential=Import-Clixml "$env:userprofile\.netoffice.pscredential"
$session=Get-PSSession -Credential $credential -Name $IcingaSession -ComputerName $computername | Connect-PSSession
if (!$session)
{
    $session=New-PSSession -ComputerName $computername -Name $IcingaSession -Credential $credential
    Invoke-Command -Session $session -ScriptBlock {Use-Icinga}|Out-Null
}
Invoke-Command -Session $session -ScriptBlock {Invoke-IcingaCheckUsedPartitionSpace}
Disconnect-PSSession $session|Out-Null


