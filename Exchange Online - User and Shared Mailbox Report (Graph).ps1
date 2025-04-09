# -------------------------------------
# Author: Mark Eggink
# Version 1.0
# Comment: User and Shared Mailbox Report (Graph)
# -------------------------------------


# Install modules if not present
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}

# Connect to Microsoft Graph and Exchange Online
Connect-MgGraph -Scopes "User.Read.All","MailboxSettings.Read","Reports.Read.All"
Connect-ExchangeOnline -UserPrincipalName (Get-MgContext).Account

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited

# Split into user and shared mailboxes
$userMailboxes = $mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox'}
$sharedMailboxes = $mailboxes | Where-Object {$_.RecipientTypeDetails -eq 'SharedMailbox'}

# Function to calculate total size in MB
function Get-MailboxTotalSizeMB {
    param (
        [array]$MailboxList
    )

    $totalSize = 0
    foreach ($mb in $MailboxList) {
        $stats = Get-MailboxStatistics -Identity $mb.UserPrincipalName
        if ($stats.TotalItemSize -match "(\d+(,\d+)?(\.\d+)?)\sMB") {
            $sizeMB = [double]($matches[1] -replace ",", "")
        } elseif ($stats.TotalItemSize -match "(\d+(,\d+)?(\.\d+)?)\sGB") {
            $sizeMB = [double]($matches[1] -replace ",", "") * 1024
        } else {
            $sizeMB = 0
        }
        $totalSize += $sizeMB
    }
    return [Math]::Round($totalSize, 2)
}

# Calculate info
$userCount = $userMailboxes.Count
$userTotalSizeMB = Get-MailboxTotalSizeMB -MailboxList $userMailboxes

$sharedCount = $sharedMailboxes.Count
$sharedTotalSizeMB = Get-MailboxTotalSizeMB -MailboxList $sharedMailboxes

# Create report object
$report = @()
$report += [PSCustomObject]@{
    MailboxType = "User"
    Count       = $userCount
    TotalSizeMB = $userTotalSizeMB
}
$report += [PSCustomObject]@{
    MailboxType = "Shared"
    Count       = $sharedCount
    TotalSizeMB = $sharedTotalSizeMB
}

# Export to CSV
$report | Export-Csv -Path "MailboxReport.csv" -NoTypeInformation -Encoding UTF8

Write-Host "Report saved to MailboxReport.csv" -ForegroundColor Green
