# ------------------------------------------------------------------------
# NAME: Exchange-Online-Extract.ps1
# AUTHOR: Cody Eding
# DATE: 5/8/2016
#
# COMMENTS: Extracts messages matching a sender email address and subject 
# from all mailboxes in Exchange Online and moves the messages to a holding
# mailbox for viewing/deletion.
#
# ------------------------------------------------------------------------

# Edit these variables to suit your environment ##############

$From = "spamextract@domain.com"
$To = "admin@domain.com"
$SmtpServer = "mail.domain.com"
$HoldingMailbox = "holdingmailbox@domain.com"
$ErrorActionPreference = "Stop"

##############################################################

Clear-Host

Write-Host ""
Write-Host "Office 365 Mail Extract" -ForegroundColor "Green"

Start-Sleep 1

Write-Host ""
Write-Host "Please enter an Office 365 credential with proper permissions to complete this action."
Write-Host ""

Try {

    # Get Office 365 Credentials
    $O365Credentials = Get-Credential -Credential $Null

}
Catch {

    Write-Host "Unable to gather credentials." -Foreground "Yellow"
    Write-Host ""
    exit 1

}

Try {

    # Login to Outlook PowerShell
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $O365Credentials -Authentication Basic -AllowRedirection -WarningAction:SilentlyContinue

    # Import the new PS Session
    Import-PSSession $Session -AllowClobber -DisableNameChecking | Out-Null

}
Catch {

    Write-Host "Unable to connect to Exchange Online. Please ensure your credentials are correct." -ForegroundColor "Yellow"
    Write-Host ""
    exit 1

}


$Sender = Read-Host 'Enter sender email address (not case sensitive)'
$Subject = Read-Host 'Enter subject search phrase (not case sensitive)'
Write-Host ""


If ( !$Sender -or !$Subject) {
    Write-Host 'Sender and/or subject can not be blank. Exiting with error status.' -ForegroundColor "Yellow"
    Write-Host ""
    exit 1
}

Write-Host "Search for email from: $Sender" -ForegroundColor "Green"
Write-Host "Search for subject: $Subject" -ForegroundColor "Green"
Write-Host ""
$Confirm = Read-Host 'Is this correct? (y/n)'
Write-Host ""

If ( $Confirm -eq 'y' ) {

    $Timestamp = (Get-Date).ToString("MM-dd-yyyy-hhmmss")
    $Recipients = ( Get-MessageTrace -SenderAddress $Sender -PageSize 5000 | Where-Object { $PSItem.Status -eq "Delivered" } ).RecipientAddress
    $Mailboxes = @()

    $Recipients | ForEach-Object {

        $Mailboxes += Get-Mailbox $PSItem

    }

    $MailboxCount = $Mailboxes.Count

    $Query = [ScriptBlock]::create('from:"' + $Sender + '" AND ' + 'subject:"' + $Subject + '"')

    $Reconfirm = Read-Host "Ready to extract mail from $MailboxCount mailbox(es), continue? (y/n)"
    Write-Host ""
	
    If ( $Reconfirm -eq 'y' ) {
	
        $StartMessageSubject = "Message Extract $Timestamp Started"
        $StartMessageBody = "Message extract started by $(whoami) on host $env:COMPUTERNAME.`n`nSearch sender: $Sender`nSearch subject: $Subject`n`nSearching $MailboxCount mailboxes.`n`nMessages will be moved to extraction mailbox $HoldingMailbox."
        Send-MailMessage -From $From -To $To -Subject $StartMessageSubject -Body $StartMessageBody -SmtpServer $SmtpServer
		
	    $Mailboxes | Search-Mailbox -SearchQuery $Query -TargetMailbox $HoldingMailbox -TargetFolder "$Timestamp" -LogLevel Full -DeleteContent -Confirm:$False -Force -WarningAction:SilentlyContinue | Out-Null

        $EndMessageSubject = "Message Extract $Timestamp Completed"
        $EndMessageBody = "Message extract completed.`n`nPlease open the $Timestamp folder in the $HoldingMailbox mailbox to view results."
        Send-MailMessage -From $From -To $To -Subject $EndMessageSubject -Body $EndMessageBody -SmtpServer $SmtpServer

        Write-Host "Message extract complete." -ForegroundColor "Green"
        Write-Host ""
		
    }
    Elseif ( $Reconfirm -eq 'n' ) {
        Write-Host 'Exiting due to user input.'
        Write-Host ""
        exit 0

        Remove-PSSession $Session
    } 
    Else {
        Write-Host 'Unexpected input. Exiting with error status.' -ForegroundColor "Yellow"
        Write-Host ""
        exit 1
    }
} 
Elseif ( $Confirm -eq 'n' ) {
    Write-Host 'Exiting due to user input.'
    Write-Host ""
    exit 0

    Remove-PSSession $Session
} 
Else {
    Write-Host 'Unexpected input. Exiting with error status.' -ForegroundColor "Yellow"
    Write-Host ""
    exit 1
}

Remove-PSSession $Session
exit 0