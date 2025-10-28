<#
Creates a WMI filter that targets the PDC Emulator (DomainRole = 5),
creates a GPO and assigns the WMI filter to that GPO.
Imports settings for PDC to time source and recommended settings to prevent time jumping more than 1 hour

Requirements:
 - Run this on the PDC
 - Run as a Domain Admin
 - ActiveDirectory + GroupPolicy modules
#>

#path to NTP GPO settings to import
$gpoNTPPath = "C:\ADBackups\PDCNTP"

#Parameters
$FilterName  = "PDC Role Filter"
$Description = "Targets DC holding the PDC Emulator role (DomainRole = 5)"
$Query       = "Select * from Win32_ComputerSystem where DomainRole = 5"
$GPOName     = "NTP Settings for PDC"

#Resolve domain info
$domain      = Get-ADDomain
$domainDNS   = $domain.DNSRoot
$namingContext = $domain.DistinguishedName
$dc          = (Get-ADDomainController -Discover).HostName[0]

#Build msWMI-Parm2 value
$filterString = "1;3;10;{0};WQL;root\CIMv2;{1};" -f $Query.Length, $Query

#Create identifiers
$wmiGuid = ([guid]::NewGuid()).ToString("B").ToUpper()
$timestamp = (Get-Date).ToUniversalTime().ToString("yyyyMMddHHmmss.ffffff-000")

#Prepare attributes
$attributes = @{
    "showInAdvancedViewOnly" = "TRUE"
    "msWMI-Name"             = $FilterName
    "msWMI-Parm1"            = $Description
    "msWMI-Parm2"            = $filterString
    "msWMI-Author"           = "$($env:USERNAME)@$($env:USERDNSDOMAIN)"
    "msWMI-ID"               = $wmiGuid
    "instanceType"           = 4
    "msWMI-ChangeDate"       = $timestamp
    "msWMI-CreationDate"     = $timestamp
}

#Create the WMI filter AD object
$wmiPath = "CN=SOM,CN=WMIPolicy,CN=System,$namingContext"
Write-Host "Creating WMI filter '$FilterName'..."
New-ADObject -Name $wmiGuid -Type "msWMI-Som" -Path $wmiPath -OtherAttributes $attributes -ErrorAction Stop
Write-Host "WMI filter created with ID $wmiGuid"

#Create a GPO
Write-Host "Creating GPO '$GPOName'..."
$gpo = New-GPO -Name $GPOName -Comment "Applies only to PDC Emulator"


#Find GPO AD object
$gpoObj = Get-ADObject -LDAPFilter "(&(objectClass=groupPolicyContainer)(cn={$($gpo.Id)}))" -ErrorAction Stop

#Link WMI filter to the GPO
$filterStringForGPO = "[{0};{{{1}}};0]" -f $domainDNS, ($wmiGuid.Trim("{}").ToUpper())
Write-Host "Linking WMI filter to GPO..."
Set-ADObject -Identity $gpoObj.DistinguishedName -Replace @{ gPCWQLFilter = $filterStringForGPO } -ErrorAction Stop

Write-Host "Done. WMI filter '$FilterName' created and linked to GPO '$GPOName'." -ForegroundColor Green

$gpoID = (get-gpo -Name $GPOName).id 
Import-GPO -Path $gpoNTPPath -TargetGuid 7ecb42df-2e71-44c6-90f1-bb3d95fefa7a -BackupId A5214940-95CC-4E93-837D-5D64CA58935C


