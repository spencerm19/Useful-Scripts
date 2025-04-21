# Script: Get-O365GroupMembers.ps1
# Description: Retrieves and exports members of a Microsoft 365 Group to CSV
#
# Note: This script requires PowerShell execution policy to be set to RemoteSigned or less restrictive.
# To set the execution policy, run the following command as administrator:
#     Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
# Alternatively, you can run the script using:
#     powershell -ExecutionPolicy Bypass -File .\Get-O365GroupMembers.ps1

# Function to ensure module is installed and imported
function Install-RequiredModule {
    param (
        [string]$ModuleName
    )
    
    try {
        # Check if module is installed
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            Write-Host "Installing $ModuleName module..." -ForegroundColor Yellow
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        
        # Import the module
        Write-Host "Importing $ModuleName module..." -ForegroundColor Cyan
        Import-Module $ModuleName -Force -ErrorAction Stop
        Write-Host "$ModuleName module imported successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Error installing/importing $ModuleName module: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please try running the following commands manually:" -ForegroundColor Yellow
        Write-Host "1. Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber" -ForegroundColor Yellow
        Write-Host "2. Import-Module $ModuleName -Force" -ForegroundColor Yellow
        exit 1
    }
}

# Get the script's directory
$scriptPath = $PSScriptRoot
if (-not $scriptPath) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Install and import required module
Install-RequiredModule -ModuleName "ExchangeOnlineManagement"

# Check for existing session
$existingSession = Get-PSSession | Where-Object {
    $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"
}

if (-not $existingSession) {
    try {
        # Connect to Exchange Online with the provided UPN
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowProgress $true -ErrorAction Stop
    }
    catch {
        Write-Host "Error connecting to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please ensure you have the correct permissions and try again." -ForegroundColor Yellow
        exit 1
    }
}
else {
    Write-Host "Using existing Exchange Online session..." -ForegroundColor Green
}

try {
    # Prompt for Group Identity
    $groupName = Read-Host -Prompt "Enter the group email address"

    # Get the group
    Write-Host "Retrieving group information..." -ForegroundColor Cyan
    $group = Get-UnifiedGroup -Identity $groupName -ErrorAction Stop

    # Get group members
    Write-Host "Retrieving members of group: $($group.DisplayName)" -ForegroundColor Yellow
    $members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -ErrorAction Stop

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
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure you have the correct permissions and the group exists." -ForegroundColor Yellow
}
finally {
    # Disconnect session
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Cyan
    }
    catch {
        # Ignore disconnection errors
    }
}