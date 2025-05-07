# Organization Hierarchy Generator

A Python script that generates a JSON file containing your organization's reporting structure using Microsoft Graph API. The script efficiently handles large organizations by implementing pagination and local filtering of service accounts.

## Features

- Fetches all enabled users from Microsoft Graph API
- Builds complete organizational hierarchy tree
- Filters out service accounts and non-standard users
- Automatically identifies organization root (typically CEO/highest level executive)
- Supports starting from a specific manager's email
- Generates optimized JSON output of the reporting structure
- Includes logging for troubleshooting

## Requirements

- Python 3.8 or newer
- Internet connection
- Microsoft Graph API credentials (Azure AD App registration)
- Required Python packages (automatically installed if missing):
  - msal
  - requests
  - python-dotenv

## Setup

1. Clone or download this repository
2. Copy `.env.template` to `.env` and fill in your Microsoft Graph API credentials:
   ```
   # Microsoft Graph API Authentication (required for get_org_hierarchy.py)
   TENANT_ID=your_tenant_id_here
   CLIENT_ID=your_client_id_here
   CLIENT_SECRET=your_client_secret_here
   ```
3. Ensure your Azure AD application has the necessary Microsoft Graph API permissions:
   - User.Read.All
   - Directory.Read.All

## Usage

1. Open a terminal or command prompt
2. Navigate to the project directory
3. Run the script:

   ```bash
   # Generate full org hierarchy from the top
   python get_org_hierarchy_solo.py

   # Start from a specific manager
   python get_org_hierarchy_solo.py "manager@yourdomain.com"
   ```

## Output Files

- `org_hierarchy_optimized_[timestamp].json`: Organization structure in JSON format
- `get_org_hierarchy.log`: Detailed execution logs including any errors or warnings

## Filtering

The script automatically filters out:
- Disabled accounts
- External users
- Service accounts
- Administrative accounts
- Accounts without email addresses or job titles
- Specific patterns like "#EXT#", "@*.onmicrosoft.com", etc.

## Troubleshooting

1. Check the `get_org_hierarchy.log` file for detailed error messages and execution flow
2. Verify your Azure AD credentials in the `.env` file
3. Ensure your Azure AD application has the required permissions
4. Check your internet connection

## Security Note

⚠️ The `.env` file contains sensitive API credentials - ensure it is not shared or committed to version control. 