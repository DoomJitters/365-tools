# Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get all mail-enabled users and export UPN & Aliases
$users = Get-Mailbox -ResultSize Unlimited

$result = foreach ($user in $users) {
    [PSCustomObject]@{
        DisplayName = $user.DisplayName
        UPN         = $user.UserPrincipalName
        PrimarySMTP = $user.PrimarySmtpAddress
        Aliases     = ($user.EmailAddresses | Where-Object { $_ -like "smtp:*" -and $_ -ne $user.PrimarySmtpAddress } | ForEach-Object { $_.ToString().Replace("smtp:", "") }) -join ", "
    }
}

# Output to screen
$result | Format-Table -AutoSize

# Export to CSV
$result | Export-Csv -Path ".\M365_Users_and_Aliases.csv" -NoTypeInformation

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
