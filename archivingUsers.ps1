

#Get UserName
param(
    [Parameter(Mandatory = $true,
                    Position = 0)]
    [String]
    $AccountToDisable
    )
#Load ActiveDirectory Module
If (!(Get-module ActiveDirectory )) 
{
    write-host "Loading Active Directory modules" -foregroundcolor "green"
    Import-Module ActiveDirectory
}

#Check if username exists
$User = $(try {get-aduser -identity $AccountToDisable} catch {$null})
If ($User -eq $Null)
{
    cls
    write-host "!!! Username" $AccountToDisable "Does not exist!!! " -foregroundcolor "red"
}
else
{
    #Set date as variable
    $date = Get-Date -format "yyyyMMdd-HHmm"

    #Load assembly to show message box
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null

    #Prompt for confirmation of user account removal
    if([System.Windows.Forms.MessageBox]::Show("Disable " + $AccountToDisable + " account, archive mailbox and remove from all groups?", "Question",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes")
    {
        cls
        write-host "Processing the account" $AccountToDisable "for archive and disabling" -foregroundcolor "green"


        #Check that the command is running from the Exchange Management Shell
        if (!(Get-Command get-exchangeserver -errorAction SilentlyContinue))
        {
            Write-Host "!!! Run this script from the Exchange Management Shell !!!" -foregroundcolor "Red"
            Break
        }
        connect-exchangeserver -auto


        #Creates Archive Path

        $Exportpath = "\\ARCHIVESERVER\Archive\" + $AccountToDisable
        if(!(Test-Path -Path $Exportpath))
        {
            new-item -path $exportpath -type directory | out-null
            write-host "* Created Archive folder" $exportpath -foregroundcolor "green"
        }
        else
        {
            write-host "!!! Archive path already exists !!!" -foregroundcolor "red"
        }
        #Remove Activesync Access

        IF ((Get-CASMailbox $AccountToDisable | where-object {$_.ActiveSyncEnabled -eq $true})) 
        {
            Set-CASMailbox -Identity $AccountToDisable -ActiveSyncEnabled $false
            write-host "* Disabled Activesync for" $AccountToDisable -foregroundcolor "green"
        }
        else
        {
            write-host "* Activesync already disabled for" $AccountToDisable -foregroundcolor "green"
        }

        #Load the other parts of Exchange console so that the mailbox export works
        if(!(Get-PSSnapin | Where-Object {$_.name -eq "Microsoft.Exchange.Management.PowerShell.E2010"})) 
        {
            write-host "Loading Exchange 2010 modules" -foregroundcolor "green"
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 | out-null
            EXCHANGE VARIABLEs ########################################################
            $global:exbin = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "bin\"
            $global:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath
            $global:exscripts = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "scripts\"
            #LOAD CONNECTION FUNCTIONS #################################################
            . $global:exbin"CommonConnectFunctions.ps1" | out-null
            . $global:exbin"ConnectFunctions.ps1" | out-null
            $FormatEnumerationLimit = 16
        }

        #Mailbox Archive

        $mailboxtarget = "\\ARCHIVESERVER\Archive\" + $AccountToDisable + "\" + $AccountToDisable + "-Mailbox-" + $date + ".csv"
        IF ((get-mailbox -identity $AccountToDisable -ErrorAction SilentlyContinue))
        {
            #Exports mailbox permissions
            Get-MailboxPermission -identity $AccountToDisable | where-object {$_.Deny -eq $False -and $_.IsInherited -eq $False} | select-object User,{$_.Accessrights},InheritanceType |export-csv -path $mailboxtarget -notypeinformation

            #Archive Mailbox
            $MailboxExport = "\\ARCHIVESERVER\archive\$AccountToDisable\" + $AccountToDisable + "-" + $date + ".pst"
            $Batchname = $AccountToDisable + "-" + $date
            New-MailboxExportRequest -mailbox $AccountToDisable -filepath $MailboxExport -BatchName $Batchname -ErrorAction Stop

            #Wait for mailbox export to complete
            while ((Get-MailboxExportRequest -BatchName $BatchName | Where {$_.Status -eq "Queued" -or $_.Status -eq "InProgress"}))
            {
                write-host "Waiting for Mailbox export to complete, waiting 60 seconds" -foregroundcolor "green"
                Get-MailboxExportRequest -BatchName $BatchName | Get-MailboxExportRequestStatistics | select-object Batchname,Status,PercentComplete
                sleep 60
            }
            #Checks to make sure that the export doesnt fail
            while ((Get-MailboxExportRequest -BatchName $BatchName | Where {$_.Status -eq "Failed"}))
            {
                write-host "!!! Mailbox export failed !!!" -foregroundcolor "Red"
                Break
            }


            #Disable Mailbox
            Disable-Mailbox -Identity $AccountToDisable -Confirm:$false
        }
        else
        {
            write-host "!!! Mailbox for" $AccountToDisable "not found !!!" -foregroundcolor "red"
            "!!! Mailbox Not Found !!!" | out-file $mailboxtarget
        }


        #Move to "Disabled Users" OU

        Get-ADUser $AccountToDisable| Move-ADObject -TargetPath 'OU=xStaff,DC=domain,DC=local'
        write-host "*" $AccountToDisable "moved to xStaff OU" -foregroundcolor "green"

        #Change Description to "Disabled YYYY.MM.DD - CURRENT USER"
        $terminatedby = $env:username
        $termDate = get-date -uformat "%Y.%m.%d"
        $termUserDesc = "Disabled " + $termDate + " - " + $terminatedby
        set-ADUser $AccountToDisable -Description $termUserDesc 
        write-host "*" $AccountToDisable "description set to" $termUserDesc -foregroundcolor "green"

        #removes from all distribution groups
        $grouptarget = "\\ARCHIVESERVER\Archive\" + $AccountToDisable + "\" + $AccountToDisable + "-Groups-" + $date + ".csv"
        $dlists =(GET-ADUSER -Identity $AccountToDisable -Properties MemberOf | Select-Object MemberOf).MemberOf
        $dlistcount = $dlists.count
        if (($dlistcount -eq "0"))
        {
            write-host "!!! No Group Memberships found for" $AccountToDisable "!!!" -foregroundcolor "red"
            "!!! No Group Memberships found !!!" | out-file $grouptarget
        }
        else
        {
            #Exports Group Memberships to CSV
            $grouptarget = "\\ARCHIVESERVER\Archive\" + $AccountToDisable + "\" + $AccountToDisable + "-Groups-" + $date + ".csv"
            Get-ADPrincipalGroupMembership $AccountToDisable | select name | Export-Csv -path $grouptarget -notypeinformation 
            write-host "* Group Memberships archived to" $grouptarget -foregroundcolor "green"
            $dlistremove =(Get-ADUser $AccountToDisable -Properties memberof | select -expand memberof)
            foreach($item in $dlistremove){Remove-ADGroupMember $AccountToDisable -Identity $item -Confirm:$False}
            write-host "* Group Memberships Removed" -foregroundcolor "green"
        }

        #moves home drive to archive
        IF(test-path \\FILESERVER\shares\Users\$AccountToDisable)
        {
            . robocopy \\FILESERVER\shares\Users\$AccountToDisable \\ARCHIVESERVER\Archive\$AccountToDisable /MOVE | out-null
            write-host "* User Drive archived to \\ARCHIVESERVER\Archive\$AccountToDisable" -foregroundcolor "green"
        }
        else
        {
            write-host "!!! No User Drive to Archive !!!" -foregroundcolor "red"
        }

        #disable user
        $Disabled = Get-Aduser $AccountToDisable

        if ($Disabled.enabled -eq $true)
        {
            Disable-ADAccount -Identity $AccountToDisable
            write-host "*** " $AccountToDisable "account has been disabled ***" -foregroundcolor "green"
        }


    }
    else
    {
        write-host "!!! No Changes Made !!!" -foregroundcolor "red"
    }
}