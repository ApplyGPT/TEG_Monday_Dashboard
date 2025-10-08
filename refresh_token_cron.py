#!/usr/bin/env python3
"""
Daily QuickBooks Token Refresh Script
Run this script daily (e.g., via Windows Task Scheduler or CRON) to keep refresh token valid
"""

import requests
import toml
import os
from datetime import datetime

def refresh_quickbooks_token():
    """Refresh QuickBooks token and update secrets.toml"""
    
    # Load credentials from secrets.toml
    base_dir = os.path.dirname(os.path.abspath(__file__))
    secrets_path = os.path.join(base_dir, '.streamlit', 'secrets.toml')
    
    try:
        with open(secrets_path, 'r') as f:
            config = toml.load(f)
        
        quickbooks_config = config.get('quickbooks', {})
        client_id = quickbooks_config.get('client_id')
        client_secret = quickbooks_config.get('client_secret')
        refresh_token = quickbooks_config.get('refresh_token')
        
        if not all([client_id, client_secret, refresh_token]):
            print(f"[ERROR] Missing QuickBooks credentials in {secrets_path}")
            return False
        
        # Authenticate with QuickBooks
        auth_url = "https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer"
        
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json"
        }
        
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token
        }
        
        auth = requests.auth.HTTPBasicAuth(client_id, client_secret)
        response = requests.post(auth_url, data=data, headers=headers, auth=auth)
        
        if response.status_code != 200:
            print(f"[ERROR] Authentication failed: {response.status_code}")
            print(f"Response: {response.text}")
            return False
        
        # Get tokens
        auth_response = response.json()
        new_access_token = auth_response.get("access_token")
        new_refresh_token = auth_response.get("refresh_token")
        
        if not new_access_token:
            print("[ERROR] No access token received")
            return False
        
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✅ QuickBooks authentication successful")
        
        # Update refresh token if a new one is provided
        if new_refresh_token and new_refresh_token != refresh_token:
            try:
                # Update the config
                config['quickbooks']['refresh_token'] = new_refresh_token
                
                # Write back to secrets.toml
                with open(secrets_path, 'w') as f:
                    toml.dump(config, f)
                
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 🔄 Refresh token updated and saved")
                print(f"   Old token: {refresh_token[:30]}...")
                print(f"   New token: {new_refresh_token[:30]}...")
            except Exception as e:
                print(f"[ERROR] Failed to update secrets.toml: {str(e)}")
                return False
        else:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ℹ️  Refresh token unchanged (still valid)")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Unexpected error: {str(e)}")
        return False

if __name__ == "__main__":
    print("=" * 80)
    print("QuickBooks Daily Token Refresh")
    print("=" * 80)
    
    success = refresh_quickbooks_token()
    
    if success:
        print("\n✅ Token refresh completed successfully!")
        print("=" * 80)
        exit(0)
    else:
        print("\n❌ Token refresh failed!")
        print("=" * 80)
        print("\nIf the refresh token expired (100 days), you need to re-authorize:")
        print("1. Visit: https://developer.intuit.com/v2/OAuth2Playground/RedirectUrl")
        print("2. Get a new authorization code")
        print("3. Run: python get_new_token.py")
        print("=" * 80)
        exit(1)

