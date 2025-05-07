# Organization Hierarchy Generator

Generates a JSON file containing your organization's reporting structure using Microsoft Graph API.

## ⚠️ Important
This script contains API credentials - do not share it with unauthorized users.

## Requirements
- Windows 10/11
- Python 3.8 or newer (from Microsoft Store or python.org)
- Internet connection

## Quick Start (Windows)

1. Save `get_org_hierarchy_solo.py` to a folder (e.g., `C:\Scripts`)
2. Open PowerShell or Command Prompt:
   - Press `Win + X` and select "Windows PowerShell" or "Command Prompt"
   - Or type "PowerShell" in Windows Search
3. Navigate to your script folder:
   ```cmd
   cd C:\path\to\your\script\folder
   ```
4. Run the script:
   ```cmd
   # Generate full org hierarchy
   python get_org_hierarchy_solo.py

   # Or start from specific manager
   python get_org_hierarchy_solo.py "manager@yourdomain.com"
   ```

## Output
- `org_hierarchy_[timestamp].json`: Organization structure
- `get_org_hierarchy.log`: Execution logs and errors

## Need Help?
Check the log file for error messages or contact your system administrator. 