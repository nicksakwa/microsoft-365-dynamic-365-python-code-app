import msal
import requests
import json
import os

# --- Configuration (REPLACE THESE WITH YOUR VALUES) ---
# You can use a dedicated file for secrets, but for simplicity, we'll keep them here.
CLIENT_ID = "YOUR_APPLICATION_CLIENT_ID_HERE"
TENANT_ID = "YOUR_DIRECTORY_TENANT_ID_HERE"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"] # Used for acquiring the token
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0/me/messages'

def acquire_access_token():
    """Acquires an access token for Microsoft Graph using Device Code Flow."""
    
    # Initialize the MSAL Public Client Application
    app = msal.PublicClientApplication(
        CLIENT_ID, 
        authority=AUTHORITY
    )

    # Initiate the Device Code Flow
    flow = app.initiate_device_flow(scopes=SCOPE)
    if "user_code" not in flow:
        raise ValueError("Failed to initiate device flow.")
    
    print("--- AUTHENTICATION REQUIRED ---")
    print(f"1. Open your web browser to: {flow['verification_uri']}")
    print(f"2. Enter the code: {flow['user_code']}")
    print("------------------------------")

    # Wait for the user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        return result['access_token']
    else:
        print(f"Error during authentication: {result.get('error_description')}")
        return None

def get_emails(token):
    """Fetches the latest emails using the acquired access token."""
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json'
    }
    
    # Request the top 5 messages, selecting only the subject, sender, and body preview
    params = {
        '$top': 5,
        '$select': 'subject,sender,bodyPreview'
    }

    print("\nAttempting to fetch emails from Microsoft Graph...")
    
    try:
        response = requests.get(GRAPH_API_ENDPOINT, headers=headers, params=params)
        response.raise_for_status() # Raises an HTTPError for bad responses (4xx or 5xx)
        
        data = response.json()
        
        print(f"\nSuccessfully retrieved {len(data.get('value', []))} emails:\n")
        
        for email in data.get('value', []):
            print(f"  Subject: {email.get('subject')}")
            print(f"  From: {email.get('sender', {}).get('emailAddress', {}).get('address')}")
            print(f"  Preview: {email.get('bodyPreview')[:70]}...")
            print("-" * 50)
            
    except requests.exceptions.HTTPError as e:
        print(f"\nHTTP Error fetching emails: {e}")
        print("Check if you granted admin consent for the Mail.Read permission.")
    except Exception as e:
        print(f"\nAn error occurred: {e}")


if __name__ == '__main__':
    # 1. Acquire Token
    access_token = acquire_access_token()

    if access_token:
        # 2. Get Emails
        get_emails(access_token)
    else:
        print("Failed to acquire access token. Cannot proceed.")