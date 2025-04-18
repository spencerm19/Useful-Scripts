# Ensure required module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
}

# Import the module
Import-Module ExchangeOnlineManagement

# Get the script's directory
$scriptPath = $PSScriptRoot
if (-not $scriptPath) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Check for existing session
$existingSession = Get-PSSession | Where-Object {
    $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"
}

if (-not $existingSession) {
    
    # Connect to Exchange Online with the provided UPN
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowProgress $true
}
else {
    Write-Host "Using existing Exchange Online session..." -ForegroundColor Green
}

# Prompt for Group Identity
$groupName = Read-Host -Prompt "Enter the group email address"

# Get the group
$group = Get-UnifiedGroup -Identity $groupName -ErrorAction Stop

# Get group members
Write-Host "Retrieving members of group: $($group.DisplayName)" -ForegroundColor Yellow
$members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members

# Output result
if ($members.Count -eq 0) {
    Write-Host "No members found in group $($group.DisplayName)." -ForegroundColor Red
} else {
    $members | Select-Object Name, PrimarySmtpAddress, RecipientType | Format-Table -AutoSize
    
    # Get current date in MMDDYYYY format
    $currentDate = Get-Date -Format "MMddyyyy"
    
    # Export to CSV in the same directory as the script with date in filename
    $exportPath = Join-Path $scriptPath "O365GroupMembers-$currentDate.csv"
    
    # Export members to CSV
    $members | Select-Object Name, PrimarySmtpAddress, RecipientType |
        Export-Csv -Path $exportPath -NoTypeInformation

    # Add total count row
    "" | Add-Content -Path $exportPath # Add blank line
    "Total Members,$($members.Count)," | Add-Content -Path $exportPath
    
    Write-Host "Exported to: $exportPath" -ForegroundColor Green
    Write-Host "Total members: $($members.Count)" -ForegroundColor Yellow
}

# Disconnect session
Disconnect-ExchangeOnline -Confirm:$false