"""
Script for retrieving and building the complete organizational hierarchy from Microsoft Graph API.

This script is responsible for:
1. Fetching all enabled users from Microsoft Graph API
2. Building the complete organizational hierarchy tree
3. Identifying and filtering out service accounts and non-standard users
4. Determining the root of the organization (typically CEO/highest level executive) if no starting email is provided
5. Generating a structured JSON output of the entire reporting structure from the specified starting point

The output JSON file (org_hierarchy_optimized_*.json) is used by process_distribution_lists.py
to determine the membership of distribution lists.

Dependencies:
- Microsoft Graph API access (tenant ID, client ID, and secret in .env)
- Required Python modules (will be installed if missing): msal, requests, python-dotenv
"""

import os
import sys
import json
import logging
import subprocess
from datetime import datetime
from pathlib import Path
from collections import defaultdict

def check_and_install_modules():
    """Check for required modules and install them if missing."""
    required_modules = ['msal', 'requests', 'python-dotenv']
    missing_modules = []
    
    for module in required_modules:
        try:
            __import__(module.replace('-', '_'))
        except ImportError:
            missing_modules.append(module)
    
    if missing_modules:
        print(f"Installing missing modules: {', '.join(missing_modules)}")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing_modules)
            print("Successfully installed required modules.")
        except subprocess.CalledProcessError as e:
            print(f"Error installing modules: {str(e)}")
            sys.exit(1)

# Install required modules before importing them
check_and_install_modules()

# Now import the modules that might have been installed
import requests
import msal
from dotenv import load_dotenv

# Get the directory where the script is located
SCRIPT_DIR = Path(__file__).parent.absolute()

# Configure logging to use script directory
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-8s | %(name)s | %(funcName)s:%(lineno)d | %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(SCRIPT_DIR / 'get_org_hierarchy.log')
    ]
)
logger = logging.getLogger(__name__)

# Load environment variables from .env file in script directory
load_dotenv(SCRIPT_DIR / '.env')

# Get credentials from environment variables
CLIENT_ID = os.environ.get('CLIENT_ID')
TENANT_ID = os.environ.get('TENANT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')

if not all([CLIENT_ID, TENANT_ID, CLIENT_SECRET]):
    raise ValueError("Missing required environment variables. Please check your .env file.")

def get_graph_token():
    """Get Microsoft Graph API access token."""
    try:
        app = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET
        )
        
        result = app.acquire_token_silent(["https://graph.microsoft.com/.default"], account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" in result:
            return result["access_token"]
        else:
            logger.error(f"Error getting token: {result.get('error_description', 'Unknown error')}")
            return None
    except Exception as e:
        logger.error(f"Exception in get_graph_token: {str(e)}")
        return None

def get_all_org_users(headers):
    """Fetches all relevant users and their manager IDs efficiently using pagination."""
    all_users = []
    
    # Simpler filter for API call - more filtering will happen locally
    filter_query = "accountEnabled eq true and userType eq 'Member'"
    
    select_query = "id,displayName,userPrincipalName,mail,jobTitle,department"
    expand_query = "manager($select=id)"
    
    # Use Beta endpoint (or v1.0 might work with this simpler filter)
    users_url = f"https://graph.microsoft.com/beta/users?$filter={filter_query}&$select={select_query}&$expand={expand_query}&$top=999"
    
    # ConsistencyLevel and Count might not be needed for this simpler filter, but keeping them doesn't hurt
    headers['ConsistencyLevel'] = 'eventual' 
    users_url += '&$count=true'

    page_num = 1
    while users_url:
        try:
            logger.info(f"Fetching page {page_num} of users...")
            response = requests.get(users_url, headers=headers)
            response.raise_for_status()  # Raise an exception for bad status codes (4xx or 5xx)
            data = response.json()
            
            current_page_users = data.get("value", [])
            all_users.extend(current_page_users)
            logger.info(f"Fetched {len(current_page_users)} users on this page. Total fetched: {len(all_users)}")
            
            # Get the next page link
            users_url = data.get("@odata.nextLink")
            page_num += 1
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Error fetching users page {page_num}: {str(e)}")
            if response:
                 logger.error(f"Response Status: {response.status_code}, Body: {response.text[:500]}") # Log part of the response
            return None # Indicate failure
        except Exception as e:
             logger.error(f"Unexpected error processing users page {page_num}: {str(e)}")
             return None

    logger.info(f"Successfully fetched a total of {len(all_users)} users.")
    return all_users

def build_local_hierarchy(user_id, users_by_id, reports_by_manager):
    """Recursively builds the hierarchy using pre-fetched local data."""
    if user_id not in users_by_id:
        logger.warning(f"User ID {user_id} found in reports_by_manager but not in main user list. Skipping.")
        return None
        
    user_data = users_by_id[user_id]
    
    # Find direct reports for the current user
    direct_report_ids = reports_by_manager.get(user_id, [])
    
    # Build direct reports first to determine if any are managers
    direct_reports = []
    has_manager_reports = False
    
    for report_id in direct_report_ids:
        report_node = build_local_hierarchy(report_id, users_by_id, reports_by_manager)
        if report_node:
            direct_reports.append(report_node)
            # Check if this direct report has their own reports
            if report_node.get("directReports") and len(report_node["directReports"]) > 0:
                has_manager_reports = True
    
    hierarchy_node = {
        "id": user_data.get("id"),
        "displayName": user_data.get("displayName"),
        "userPrincipalName": user_data.get("userPrincipalName"),
        "mail": user_data.get("mail"),
        "jobTitle": user_data.get("jobTitle"),
        "department": user_data.get("department"),
        "directReports": direct_reports,
        "hasManagerReports": has_manager_reports,  # Flag indicating if this manager has other managers reporting to them
        "needsStandardList": has_manager_reports   # Flag indicating if this manager should have a Standard Distribution List
    }
            
    return hierarchy_node

def find_user_by_email_in_list(email, users_list):
    """Find a user by their email in the filtered users list."""
    if not email:
        return None
        
    email = email.lower()
    for user in users_list:
        if user.get('mail', '').lower() == email or user.get('userPrincipalName', '').lower() == email:
            return user
    return None

def main(start_email=None):
    """
    Main function to fetch all users and build hierarchy locally.
    
    Args:
        start_email (str, optional): Email address of the manager to start the hierarchy from.
                                   If None, builds hierarchy from the organization root.
    """
    # Get access token
    access_token = get_graph_token()
    if not access_token:
        logger.error("Failed to get access token")
        return

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # 1. Fetch all enabled Member users efficiently
    logger.info("Starting to fetch all enabled Member users...")
    all_users_raw = get_all_org_users(headers)
    
    if all_users_raw is None:
        logger.error("Failed to fetch user data. Aborting.")
        return
    if not all_users_raw:
        logger.warning("No users found matching the basic filter criteria.")
        return
        
    # --- Start Local Filtering ---
    logger.info(f"Fetched {len(all_users_raw)} raw users. Applying local filters...")
    all_users = []
    excluded_patterns = [
        "#EXT#",
        "@avetta1.onmicrosoft.com",
        "admin@",
        "noreply",
        "accounts@", 
        "adm.", 
        "abuse@"
    ]
    
    for user in all_users_raw:
        # Basic checks
        if not user.get('mail') or not user.get('jobTitle'):
            continue
            
        # Filter based on excluded patterns in UPN
        upn = user.get("userPrincipalName", "").lower()
        display_name = user.get("displayName", "").lower()
        if any(pattern in upn for pattern in excluded_patterns) or display_name.startswith('adm'):
            continue
            
        # If all checks pass, add to the final list
        all_users.append(user)
        
    logger.info(f"Finished local filtering. {len(all_users)} users remaining.")
    # --- End Local Filtering ---
        
    # 2. Process the *filtered* list into dictionaries for quick lookup
    users_by_id = {user['id']: user for user in all_users}
    reports_by_manager = defaultdict(list)
    root_candidates = []

    logger.info("Processing filtered user data to build lookup tables...")
    for user in all_users:
        manager_info = user.get('manager')
        if manager_info and manager_info.get('id') in users_by_id:
            reports_by_manager[manager_info['id']].append(user['id'])
        elif manager_info and manager_info.get('id') not in users_by_id:
             logger.warning(f"User {user.get('displayName')} ({user.get('id')}) has manager {manager_info.get('id')} who was filtered out. Considering this user as potential root.")
             root_candidates.append(user)
        else:
             root_candidates.append(user)
            
    logger.info(f"Built lookup tables. Found {len(root_candidates)} potential root users.")

    # 3. Select the starting point - either specified email or best root user
    root_user = None
    if start_email:
        root_user = find_user_by_email_in_list(start_email, all_users)
        if not root_user:
            logger.error(f"Could not find user with email {start_email} in the organization. Aborting.")
            return
        logger.info(f"Starting hierarchy from specified user: {root_user.get('displayName')} (Email: {root_user.get('mail')})")
    else:
        # Use existing root selection logic
        if not root_candidates:
            logger.error("No root users identified. Cannot build hierarchy. Check filters or data.")
            return
            
        # Prioritize roots based on job title
        executive_titles_highest = ['ceo', 'chief executive', 'president']
        executive_titles_clevel = ['cto', 'cfo', 'cio', 'coo', 'chief '] # Trailing space handles 'chief architect', 'chief people', etc.
        executive_titles_svp = ['senior vice president', 'svp']
        executive_titles_vp = ['vice president', 'vp']

        root_candidates.sort(key=lambda user: (
            not any(title in user.get('jobTitle', '').lower() for title in executive_titles_highest),
            not any(title in user.get('jobTitle', '').lower() for title in executive_titles_clevel),
            not any(title in user.get('jobTitle', '').lower() for title in executive_titles_svp),
            not any(title in user.get('jobTitle', '').lower() for title in executive_titles_vp),
            user.get('displayName', '')
        ))

        root_user = root_candidates[0]
        if len(root_candidates) > 1:
             logger.warning(f"Multiple root candidates found. Sorted by title and selected: {root_user.get('displayName')}")
             for i, candidate in enumerate(root_candidates[1:11], 1):
                 logger.info(f"  Other root candidate {i}: {candidate.get('displayName')} ({candidate.get('jobTitle')})")
             if len(root_candidates) > 11:
                    logger.info(f"  ... and {len(root_candidates) - 11} more potential root candidates.")
                 
        logger.info(f"Selected root user: {root_user.get('displayName')} (ID: {root_user.get('id')}) - Job: {root_user.get('jobTitle')}")

    # 4. Build the hierarchy locally
    logger.info("Building hierarchy from local data...")
    org_hierarchy = build_local_hierarchy(root_user['id'], users_by_id, reports_by_manager)
    
    # 5. Save the hierarchy and generate summary
    if org_hierarchy:
        # Use script directory for output
        output_dir = SCRIPT_DIR
        
        # Add email to filename if specified
        filename_prefix = "org_hierarchy"
        if start_email:
            email_part = start_email.split('@')[0]  # Use part before @ for filename
            filename_prefix = f"org_hierarchy_{email_part}"
            
        output_file = output_dir / f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        # Count managers and distribution list requirements
        def count_managers_in_hierarchy(node):
            total_managers = 0
            standard_list_managers = 0
            if node.get("directReports"):
                total_managers += 1
                if node.get("needsStandardList"):
                    standard_list_managers += 1
            for report in node.get("directReports", []):
                t, s = count_managers_in_hierarchy(report)
                total_managers += t
                standard_list_managers += s
            return total_managers, standard_list_managers
        
        total_managers, standard_list_managers = count_managers_in_hierarchy(org_hierarchy)
        logger.info(f"Organization Summary:")
        logger.info(f"Total managers (requiring Dynamic Lists): {total_managers}")
        logger.info(f"Managers requiring Standard Lists: {standard_list_managers}")
        logger.info(f"Bottom-tier managers (Dynamic List only): {total_managers - standard_list_managers}")
        
        try:
            with open(output_file, 'w') as f:
                json.dump(org_hierarchy, f, indent=2)
            logger.info(f"Optimized organizational hierarchy saved to {output_file}")
        except Exception as e:
             logger.error(f"Error saving hierarchy to file {output_file}: {str(e)}")
    else:
        logger.error("Failed to build organizational hierarchy from local data.")

if __name__ == "__main__":
    import sys
    start_email = sys.argv[1] if len(sys.argv) > 1 else None
    main(start_email) 