import msal
import requests
import json
import os

CLIENT_ID = 'your_client_id'
TENANT_ID = 'your_tenant_id'
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPES = ['https://graph.microsoft.com/.default']
GRAPH_API_ENDPOINT='https://graph.microsoft.com/v1.0/me/messages'

def get_access_token():
    app=msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY
    )
    flow = app.initiate_device_flow(scopes=SCOPES)
    if 'user_code' not in flow:
        raise ValueError('Failed to create device flow')
    print("---AUTHENTICATION REQUIRED---")
    print(f"1. Open your browser to: {flow['verification_url']}")
    print(f"2. Enter the code: {flow['user_code']}")
    print("-----------------------------")
    result = app.acquire_token_by_device_flow(flow)
    if 'access_token' in result:
        return result['access_token']
    else:
        print(f"Error during authentication: {result.get('error_description')}")
        return None
def get_user_emails(access_toke):
    headers={
        'Authorization': f'bearer {token}',
        'Content-Type': 'application/json'
    }
    params ={
        '$top': 5,
        '$select': 'Subject, sender, bodyPreview'
    }
    print("\n Attemptin to fetch emails from Microsoft Graph...")
    try:
        response= requests.get(GRAPH_API_ENDPOINT, headers=headers, params=params)
        response.raise_for_status()
        data=response.json()
        print(f"\nSuccessfully retrieved {len(data.get('value',[]))} emails:\n")
        for email in data.get('value', []):
            print(f" Subject: {email.get('subject')}")
            print(f" From: {email.get('sender', {}).get('emailAddress', {}).get('address')}")
            print(f" Preview: {email.get('bodyPreview')[:70]}...")
            print("-" * 50)
        except requests.exceptions.HTTPError as e:
            print(f"\nHTTP Error fetching emails: {e}")
            print(f"Check if you granted admin consent for email. read permissions.")
        except Exceptions as e:
            print(f"\nAn error occured: {e}")

if __name__ == '__main__':
    access_token = get_access_token()
    if access_token:
        get_emails(access_token)
    else:
        print("Failed to require access token. Cannot proceed.")