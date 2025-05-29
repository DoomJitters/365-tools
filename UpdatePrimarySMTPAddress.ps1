# Load Exchange Online module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser
}
Import-Module ExchangeOnlineManagement

# Prompt for admin creds and connect
Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter your admin UPN")

# Ask for CSV and new domain - CSV has one column header "EmailAddress"
$csvPath = Read-Host "Enter path to CSV file (e.g., C:\emails.csv)"
$newDomain = Read-Host "Enter the new domain (e.g., newdomain.com)"

# Ask if WhatIf mode for preview of changes
$whatIf = Read-Host "Run in WhatIf (preview only) mode? (Y/N)"
$confirmChanges = $false
if ($whatIf -eq 'Y') {
    $confirmChanges = $false
    Write-Host "`nRunning in WHAT-IF mode. No changes will be made until confirmed." -ForegroundColor Yellow
} else {
    $confirmChanges = $true
}

# Read CSV
try {
    $users = Import-Csv -Path $csvPath
} catch {
    Write-Error "Failed to read CSV file. Please check the path."
    exit
}

$changesToApply = @()

foreach ($user in $users) {
    $oldEmail = $user.EmailAddress
    if ($oldEmail -match "^([^@]+)@(.+)$") {
        $localPart = $matches[1]
        $newEmail = "$localPart@$newDomain"

        try {
            # Get mailbox
            $mailbox = Get-Mailbox -Identity $oldEmail -ErrorAction Stop

            # Get current email addresses
            $currentAddresses = $mailbox.EmailAddresses

            # Prepare new addresses
            $proposedAddresses = $currentAddresses | Where-Object { $_ -notlike "SMTP:*" }

            if ($currentAddresses -notcontains "smtp:$oldEmail") {
                $proposedAddresses += "smtp:$oldEmail"
            }

            $proposedAddresses += "SMTP:$newEmail"

            if (-not $confirmChanges) {
                Write-Host "`n--- WHAT-IF: $oldEmail ---" -ForegroundColor Cyan
                Write-Host "Current primary: $($mailbox.PrimarySmtpAddress)"
                Write-Host "New primary:     $newEmail"
                Write-Host "Adding alias:    $oldEmail"
            } else {
                $changesToApply += [PSCustomObject]@{
                    Identity = $mailbox.Identity
                    NewPrimary = "SMTP:$newEmail"
                    OldAlias = "smtp:$oldEmail"
                    NewAddresses = $proposedAddresses
                }
            }
        }
        catch {
            Write-Warning "Failed to process $oldEmail: $_"
        }
    } else {
        Write-Warning "Invalid email format: $oldEmail"
    }
}

if (-not $confirmChanges) {
    $proceed = Read-Host "`nProceed with making these changes? (Y/N)"
    if ($proceed -ne 'Y') {
        Write-Host "Aborting changes." -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false
        exit
    } else {
        Write-Host "`nApplying changes..." -ForegroundColor Green
        $confirmChanges = $true
    }
}

# Apply changes
foreach ($change in $changesToApply) {
    try {
        Set-Mailbox -Identity $change.Identity -EmailAddresses $change.NewAddresses
        Write-Host "Updated: $($change.Identity) -> $($change.NewPrimary) (alias: $($change.OldAlias))" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to apply change to $($change.Identity): $_"
    }
}

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false
