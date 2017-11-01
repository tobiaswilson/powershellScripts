# connects to office 365
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

<# would it be better to ask for username instead because that can be used for AD and stick @domain on end? Could ask for both?#>

$username = Read-Host -Prompt 'Input username' # input username
$email = $username + "@toriglobal.com" # obviously this is specific to Tori

<# old code 
Get-Mailbox -identity $email | set-mailbox -type "Shared" where-object {$_.type -eq 'Regular'} # set mailbox type to be shared
# need testing loop here to make sure the mailbox is changed to a shared one.
#>

<# do until loop to keep setting mailbox to be shared until it's registered as shared #>

do{ Get-Mailbox -identity $email -type "Shared" Where-Object {$_.type -eq 'Regular'} until ($_.type -eq 'Shared')

Connect-MsolService
$MSOLSKU = (Get-MSOLUser -UserPrincipalName $email).Licenses[0].AccountSkuID # gets the correct sku of the license to remove
Set-MsolUserLicense -UserPrincipalName $email -RemoveLicenses $MOLSKU # removes license

Get-Mailbox -Identity $email | Format-List RecipitentTypeDetails # prints out current details

<# this works on the STG01 server, i.e the AD disable #>

import-module activedirectory # connects to AD -> how to connect to specific AD? or just put on server?
Disable-ADAccount -Identity $username # disables AD account

Remove-PSSession $Session # disconnects from office 365