# Useful Scripts

This is a collection of useful scripts to quickly accomplish a variety of internal IT related tasks.

## Get-VivaCommunityMembers.ps1

### Description
This PowerShell script retrieves and exports the members of a specified Microsoft 365 Group or Viva Engage Community. It connects to Exchange Online, fetches member information, displays it in the console, and exports the data to a CSV file.

### Prerequisites
- PowerShell 5.1 or later
- Exchange Online Management module (automatically installed if missing)
- Microsoft 365 account with appropriate permissions to:
  - Access Exchange Online
  - Read group membership information
  - Install PowerShell modules (if ExchangeOnlineManagement module is not present)

### Features
- Automatic installation of required Exchange Online Management module
- Reuses existing Exchange Online sessions if available
- Interactive prompt for group email address
- Exports member information to CSV including:
  - Member names
  - Primary SMTP addresses
  - Recipient types
- Includes total member count in both console output and CSV
- Automatic session cleanup after execution

### Usage
1. Open PowerShell
2. Navigate to the script directory
3. Run the script:
   ```powershell
   .\Get-VivaCommunityMembers.ps1
   ```
4. When prompted, enter the email address of the group/community
5. If not already connected, authenticate to Exchange Online when prompted

### Output
- **Console Output**:
  - Connection status
  - Member information in a formatted table
  - Total member count
  - Export file location

- **CSV File** (`O365GroupMembers.csv`):
  - Name
  - PrimarySmtpAddress
  - RecipientType
  - Total member count (appended at the end)

### Notes
- The script automatically disconnects from Exchange Online upon completion
- The CSV file is created in the same directory as the script
- Existing Exchange Online sessions are preserved and reused
- Error handling is implemented for group retrieval

### Error Handling
- Validates Exchange Online connection
- Verifies group existence
- Handles empty group scenarios
- Provides colored console output for status visibility
