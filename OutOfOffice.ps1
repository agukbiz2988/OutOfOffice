Install-Module –Name ExchangeOnlineManagement
Disconnect-ExchangeOnline
Connect-ExchangeOnline -ShowBanner:$false

#***CSV Location***
$User = Import-Csv C:\outofoffice-main\importautoreply.csv

#***American Format***
#**** MONTH / DAY / YEAR ****
$StartDate = "03/30/2025 12:00:00"
$EndDate = "05/04/2025 12:00:00"

#***Messages***
$Message = "out of office internal"

foreach($user in $User){

    Write-Host "`nAdding auto-reply message to -$user"

    try {
        #Code to Use for Scheduled Auto reply message
        Set-MailboxAutoreplyConfiguration -Identity $User.user -AutoReplyState Scheduled –StartTime $StartDate -EndTime $EndDate -Internalmessage $Message -ExternalMessage $Message

        #Code to use if you DO NOT want a scehduled auto reply message
        #Set-MailboxAutoreplyConfiguration -Identity $User.user -AutoReplyState Enabled -Internalmessage $Message -ExternalMessage $Message

        Write-Host "Auto reply message added to account $User.user `n" -ForegroundColor Green 
    }
    catch {
        Write-Error -Message "Error occurred while adding auto-reply message to $user"
    }

}

