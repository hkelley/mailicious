
# Load local Exchange cmdlets
if($null -eq (Get-Command Search-OnPremMailbox -ErrorAction SilentlyContinue))
{

    $rootdse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
    $scp = Get-ADObject -LDAPFilter '(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))' -SearchBase $rootdse.configurationNamingContext.Value -ResultSetSize 1

    $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $scp.Name) -Authentication Kerberos
    Import-PSSession $s -AllowClobber -Prefix OnPrem | Out-Null
    Write-Host ("Exchange cmdlets imported from {0} with OnPrem prefix" -f $scp.Name)
    $s = $null
}


# Load O365 cmdlets
if($null -eq (Get-Command Get-O365MessageTrace -ErrorAction SilentlyContinue))
{
    Import-Module "$PSScriptRoot\Microsoft.Exchange.Management.ExoPowershellModule.dll"
    Import-PSSession ( New-ExoPSSession   -UserPrincipalName (whoami.exe /upn) ) -AllowClobber -Prefix O365 | Out-Null
    Write-Host "Office 365 cmdlets imported with O365 prefix"
}
