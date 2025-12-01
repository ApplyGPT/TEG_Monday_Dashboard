"""
QuickBooks Token Refresh Script

This script refreshes QuickBooks access tokens and updates secrets.toml.
Designed to run as a cron job every 15 minutes to keep tokens fresh.

Usage:
    python refresh_qb_tokens.py [--log-file LOG_FILE]

Exit codes:
    0 - Success
    1 - Error (check logs)

Logging:
    If --log-file is not specified, logs go to stderr (visible in cron emails).
    If --log-file is specified, logs are appended to that file.
"""

import argparse
import os
import sys
import toml
import requests
from pathlib import Path
from datetime import datetime
from typing import Optional


class Logger:
    """Simple logger that can write to file or stderr"""
    def __init__(self, log_file: Optional[Path] = None):
        self.log_file = log_file
        self.log_handle = None
        if log_file:
            try:
                self.log_handle = open(log_file, 'a', encoding='utf-8')
            except Exception as e:
                # Fall back to stderr if file can't be opened
                print(f"Warning: Could not open log file {log_file}: {e}", file=sys.stderr)
                self.log_file = None
    
    def log(self, message: str, level: str = "INFO"):
        """Log a message"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"[{timestamp}] [{level}] {message}\n"
        
        if self.log_handle:
            self.log_handle.write(log_message)
            self.log_handle.flush()
        else:
            print(log_message.strip(), file=sys.stderr)
    
    def close(self):
        """Close the log file"""
        if self.log_handle:
            self.log_handle.close()
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


def update_secrets_toml(access_token: str, refresh_token: str, logger: Optional[Logger] = None) -> bool:
    """
    Update secrets.toml file with new tokens, preserving other settings
    
    Args:
        access_token: New access token
        refresh_token: New refresh token
        logger: Optional logger instance
        
    Returns:
        True if successful, False otherwise
    """
    # Find secrets.toml file relative to script location
    script_dir = Path(__file__).parent.absolute()
    secrets_path = script_dir / '.streamlit' / 'secrets.toml'
    
    if not secrets_path.exists():
        error_msg = f"secrets.toml file not found at {secrets_path.absolute()}"
        if logger:
            logger.log(error_msg, "ERROR")
        else:
            print(f"ERROR: {error_msg}", file=sys.stderr)
        return False
    
    try:
        # Read existing secrets.toml content
        with secrets_path.open('r', encoding='utf-8') as f:
            config = toml.load(f)
        
        # Ensure quickbooks section exists
        if 'quickbooks' not in config:
            config['quickbooks'] = {}
        
        # Update tokens
        config['quickbooks']['access_token'] = access_token
        config['quickbooks']['refresh_token'] = refresh_token
        
        # Write back to file
        with secrets_path.open('w', encoding='utf-8') as f:
            toml.dump(config, f)
        
        if logger:
            logger.log(f"Updated secrets.toml with new tokens")
        
        return True
        
    except Exception as e:
        error_msg = f"Failed to update secrets.toml file: {e}"
        if logger:
            logger.log(error_msg, "ERROR")
        else:
            print(f"ERROR: {error_msg}", file=sys.stderr)
        return False


def refresh_tokens(logger: Optional[Logger] = None) -> bool:
    """
    Refresh QuickBooks tokens and update secrets.toml
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Change to script directory
        original_cwd = os.getcwd()
        script_dir = Path(__file__).parent.absolute()
        
        if logger:
            logger.log(f"Script directory: {script_dir}")
            logger.log(f"Original working directory: {original_cwd}")
        
        os.chdir(script_dir)
        
        try:
            # Load configuration from secrets.toml
            secrets_path = script_dir / '.streamlit' / 'secrets.toml'
            
            if not secrets_path.exists():
                error_msg = f"secrets.toml file not found at {secrets_path.absolute()}"
                if logger:
                    logger.log(error_msg, "ERROR")
                else:
                    print(f"ERROR: {error_msg}", file=sys.stderr)
                return False
            
            with secrets_path.open('r', encoding='utf-8') as f:
                config = toml.load(f)
            
            quickbooks_config = config.get('quickbooks', {})
            client_id = quickbooks_config.get('client_id', '').strip()
            client_secret = quickbooks_config.get('client_secret', '').strip()
            refresh_token = quickbooks_config.get('refresh_token', '').strip()
            
            # Debug: Log credentials (masked for security)
            if logger:
                logger.log(f"Loaded client_id: {client_id[:10]}...{client_id[-5:] if len(client_id) > 15 else '***'}")
                logger.log(f"Loaded client_secret: {'*' * min(len(client_secret), 20)}...")
                logger.log(f"Loaded refresh_token: {refresh_token[:10]}...{refresh_token[-5:] if len(refresh_token) > 15 else '***'}")
            
            if not all([client_id, client_secret, refresh_token]):
                error_msg = f"Missing required QuickBooks credentials in secrets.toml. Need: client_id, client_secret, refresh_token"
                if logger:
                    logger.log(error_msg, "ERROR")
                    logger.log(f"client_id present: {bool(client_id)}, client_secret present: {bool(client_secret)}, refresh_token present: {bool(refresh_token)}")
                else:
                    print(f"ERROR: {error_msg}", file=sys.stderr)
                return False
            
            # Refresh the access token
            if logger:
                logger.log("Refreshing QuickBooks tokens...")
            else:
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Refreshing QuickBooks tokens...")
            
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
            
            if logger:
                logger.log(f"Making request to: {auth_url}")
                logger.log(f"Using client_id length: {len(client_id)}, client_secret length: {len(client_secret)}")
            
            response = requests.post(auth_url, data=data, headers=headers, auth=auth, timeout=30)
            
            if response.status_code != 200:
                error_msg = f"Authentication failed: {response.status_code} - {response.text}"
                if logger:
                    logger.log(error_msg, "ERROR")
                    logger.log(f"Request URL: {auth_url}")
                    logger.log(f"Client ID (first 10 chars): {client_id[:10] if client_id else 'None'}")
                else:
                    print(f"ERROR: {error_msg}", file=sys.stderr)
                return False
            
            # Get tokens from response
            auth_response = response.json()
            new_access_token = auth_response.get("access_token")
            new_refresh_token = auth_response.get("refresh_token")
            expires_in = auth_response.get("expires_in", 3600)
            
            if not new_access_token:
                error_msg = "Token refresh succeeded but no access token returned."
                if logger:
                    logger.log(error_msg, "ERROR")
                else:
                    print(f"ERROR: {error_msg}", file=sys.stderr)
                return False
            
            # Use new refresh token if provided, otherwise keep the old one
            if not new_refresh_token:
                new_refresh_token = refresh_token
            
            # Update secrets.toml with new tokens
            if update_secrets_toml(new_access_token, new_refresh_token, logger):
                success_msg = "Tokens refreshed and secrets.toml updated successfully"
                expires_msg = f"Access token expires in: {expires_in // 3600} hours ({expires_in} seconds)"
                if logger:
                    logger.log(success_msg)
                    logger.log(expires_msg)
                else:
                    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✓ {success_msg}")
                    print(f"  {expires_msg}")
                return True
            else:
                error_msg = "Token refresh succeeded but secrets.toml update failed."
                if logger:
                    logger.log(error_msg, "ERROR")
                else:
                    print(f"ERROR: {error_msg}", file=sys.stderr)
                return False
                
        finally:
            # Restore original working directory
            os.chdir(original_cwd)
            
    except Exception as e:
        error_msg = f"Unexpected error during token refresh: {e}"
        if logger:
            logger.log(error_msg, "ERROR")
            import traceback
            logger.log(traceback.format_exc(), "ERROR")
        else:
            print(f"ERROR: {error_msg}", file=sys.stderr)
            import traceback
            traceback.print_exc(file=sys.stderr)
        return False


def main():
    """Main entry point for cron job"""
    parser = argparse.ArgumentParser(
        description="Refresh QuickBooks access tokens and update secrets.toml"
    )
    parser.add_argument(
        "--log-file",
        type=str,
        help="Path to log file (logs will be appended). If not specified, logs go to stderr.",
    )
    
    args = parser.parse_args()
    
    log_file = Path(args.log_file) if args.log_file else None
    
    with Logger(log_file) as logger:
        success = refresh_tokens(logger)
    
    if success:
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == '__main__':
    main()
