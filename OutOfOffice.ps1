Install-Module –Name ExchangeOnlineManagement
Connect-ExchangeOnline

#***CSV Location***
$User = Import-Csv C:\importautoreply.csv
#***American Format***
$StartDate = "12/24/2024 12:00:00"
$EndDate = "01/01/2025 12:00:00"
#***Messages***
$internalMessage = "out of office internal"
$externalMessage = "out of office external"

foreach($user in $User){

    Write-Host "Adding auto-reply message to -$user"

    try {
        Set-MailboxAutoreplyConfiguration -Identity $User.user -AutoReplyState Scheduled –StartTime $StartDate -EndTime $EndDate -Internalmessage $internalMessage -ExternalMessage $externalMessage
    }
    catch {
        Write-Error -Message "Error occurred while adding auto-reply message to $user"
    }

}