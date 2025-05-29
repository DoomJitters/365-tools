# Enable Windows Forms for file dialog
Add-Type -AssemblyName System.Windows.Forms

# Ensure Microsoft Graph modules are installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.AccessAsUser.All"

# Open File Explorer to pick the CSV file
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Title = "Select CSV File"
$openFileDialog.Filter = "CSV files (*.csv)|*.csv"
$openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')

$csvPath = if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $openFileDialog.FileName
} else {
    Write-Host "No file selected. Exiting..." -ForegroundColor Red
    exit
}

# Prompt for new domain
$newDomain = Read-Host "Enter the new domain (e.g., newdomain.com)"

# Ask if the script should run in WhatIf mode
$whatIf = Read-Host "Run in WhatIf mode? (Y/N)"
$confirmChanges = $whatIf -ne 'Y'

if (-not $confirmChanges) {
    Write-Host "Running in WHAT-IF mode. No changes will be made." -ForegroundColor Yellow
}

# Read CSV
try {
    $users = Import-Csv -Path $csvPath
} catch {
    Write-Error "Failed to read CSV file. Please check the path and format."
    exit
}

foreach ($user in $users) {
    $oldEmail = $user.EmailAddress
    if ($oldEmail -match "^([^@]+)@(.+)$") {
        $localPart = $matches[1]
        $newUPN = "$localPart@$newDomain"

        try {
            # Get user by UPN
            $aadUser = Get-MgUser -Filter "userPrincipalName eq '$oldEmail'" -ErrorAction Stop

            # Prepare alias addition (email addresses are in ProxyAddresses)
            $proxyAddresses = @()
            if ($aadUser.ProxyAddresses) {
                $proxyAddresses = $aadUser.ProxyAddresses
            }

            # Add old email as alias if not already present
            $aliasToAdd = "smtp:$oldEmail"
            if ($proxyAddresses -notcontains $aliasToAdd) {
                $proxyAddresses += $aliasToAdd
            }

            if (-not $confirmChanges) {
                Write-Host "--- WHAT-IF: $oldEmail ---" -ForegroundColor Cyan
                Write-Host "Current UPN:   $($aadUser.UserPrincipalName)"
                Write-Host "New UPN:       $newUPN"
                Write-Host "Add alias:     $aliasToAdd"
            } else {
                # Step 1: Update UPN
                Update-MgUser -UserId $aadUser.Id -UserPrincipalName $newUPN

                # Step 2: Add alias (proxyAddresses must include current primary SMTP)
                Update-MgUser -UserId $aadUser.Id -ProxyAddresses $proxyAddresses

                Write-Host "Updated UPN: $oldEmail → $newUPN and added alias $oldEmail" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Failed to update $oldEmail: $_"
        }
    } else {
        Write-Warning "Invalid email format in CSV: $oldEmail"
    }
}
