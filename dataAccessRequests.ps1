$mb = Get-Mailbox -Server severName -ResultSize Unlimited -RecipientTypeDetails "usermailbox"
$mb | % {$_ | New-MailboxExportRequest -Name $_.alias -ContentFilter {(Received -gt 'date') -and (All -eq 'example@example.com') -or (All -like "firstNameExample*") -or (All -like "surnameExample") } -FilePath "\\server\fileshare\$($_.alias).pst"}
