<#
=============================================================================================
Name = Cengiz YILMAZ
Date = 1.03.2023
www.cengizyilmaz.net
www.cozumpark.com/author/cengizyilmaz
============================================================================================
#>

# Prompt for DN of the OU
$ouDN = Read-Host "Please enter the distinguished name (DN) of the OU"

# Get all mailboxes in the OU
$mailboxes = Get-Mailbox -OrganizationalUnit $ouDN -ResultSize Unlimited | Get-MailboxStatistics | Select DisplayName, TotalItemSize, Database

$newDbs = @()
$dbSize = 0
$dbCount = 1

foreach ($mailbox in $mailboxes) {
    $mailboxSize = $mailbox.TotalItemSize.Value.ToGB()
    if (($dbSize + $mailboxSize) -le 250) {
        # add to current DB
        $newDbs += New-Object PSObject -Property @{
            'Mailbox' = $mailbox.DisplayName
            'DB' = "DB$dbCount"
            'SizeInGB' = $mailboxSize
        }
        $dbSize += $mailboxSize
    } else {
        # create a new DB
        $dbCount++
        $dbSize = $mailboxSize
        $newDbs += New-Object PSObject -Property @{
            'Mailbox' = $mailbox.DisplayName
            'DB' = "DB$dbCount"
            'SizeInGB' = $mailboxSize
        }
    }
}

# Export to CSV format with UTF8 encoding
$newDbs | Export-Csv -Path "new_dbs.csv" -NoTypeInformation -Encoding UTF8
