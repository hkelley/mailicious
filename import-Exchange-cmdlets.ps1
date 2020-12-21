param 
(
      [Parameter(Mandatory = $false)] $upn = (whoami.exe /upn)
)

# Load local Exchange cmdlets
if($null -eq (Get-Command Search-OnPremMailbox -ErrorAction SilentlyContinue))
{
    $rootdse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE")
    foreach($scp in (Get-ADObject -LDAPFilter '(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))' -SearchBase $rootdse.configurationNamingContext.Value ))
    {
        if(Test-Connection -ComputerName $scp.Name -Count 2 -ErrorAction SilentlyContinue )
        {
            $s = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri ("http://{0}/PowerShell/" -f $scp.Name) -Authentication Kerberos

            Import-PSSession $s -AllowClobber -Prefix OnPrem | Out-Null
            Write-Host ("Exchange cmdlets imported from {0} with OnPrem prefix" -f $scp.Name)
            $s = $null

            break
        }
        else
        {
            Write-Warning ("Could not reach Exchange host {0} described in {1}" -f $scp.Name,$scp.DistinguishedName)
        }
    }
}


# Load O365 cmdlets 
if($null -eq (Get-Command Get-O365MessageTrace -ErrorAction SilentlyContinue))
{
    $exoModulePath = $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"}|select -First 1)
    Write-Host "Loading Exchange Online cmdlets with assembly: $exoModulePath"
    Import-Module $exoModulePath
    Import-PSSession ( New-ExoPSSession -UserPrincipalName $upn ) -AllowClobber -Prefix O365 | Out-Null
    Write-Host "Office 365 cmdlets imported with O365 prefix"
}


# Load Security and Compliance cmdlets  - requires latest Exchange Online Remote PowerShell Module to support Connect-IPPSSession 
if($null -eq (Get-Command Get-ComplianceSearch -ErrorAction SilentlyContinue))
{
#    # Import-Module "$PSScriptRoot\Microsoft.Exchange.Management.ExoPowershellModule.dll"
#    $exoModulePath = $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|?{$_ -notmatch "_none_"}|select -First 1)
#    Write-Host "Loading Exchange Online cmdlets with assembly: $exoModulePath"
#    Import-Module $exoModulePath
    Import-PSSession ( New-ExoPSSession -ConnectionUri “https://ps.compliance.protection.outlook.com/PowerShell-LiveId”  -UserPrincipalName $upn ) -AllowClobber | Out-Null
    Write-Host "Security and Compliance cmdlets imported with no prefix"
}
