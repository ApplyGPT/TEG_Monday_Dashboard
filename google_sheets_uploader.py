"""
Utility helpers to upload generated workbooks to Google Sheets.
"""

from __future__ import annotations

import io
import os
import socket
from typing import Optional

import streamlit as st

try:
    from google.oauth2.service_account import Credentials as SACredentials
    from google_auth_oauthlib.flow import InstalledAppFlow, Flow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseUpload
except Exception:
    SACredentials = None
    InstalledAppFlow = None
    Flow = None
    Request = None
    build = None
    HttpError = None
    MediaIoBaseUpload = None

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/drive.file"]


class GoogleSheetsUploadError(Exception):
    """Raised when the workbook cannot be uploaded to Google Sheets."""


def load_oauth_credentials():
    """Load user OAuth credentials from session state or initiate OAuth flow."""
    if not Flow or not Request:
        return None
    
    try:
        oauth = st.secrets.get('google_oauth')
        if not oauth:
            return None
        
        scopes = GOOGLE_SCOPES
        
        # Check if we have stored credentials in session state
        if 'wb_google_creds' in st.session_state and st.session_state.wb_google_creds:
            try:
                # Try to refresh the credentials
                creds = st.session_state.wb_google_creds
                if creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                    st.session_state.wb_google_creds = creds
                return creds
            except Exception:
                # If refresh fails, clear the stored credentials
                st.session_state.wb_google_creds = None
        
        # Determine redirect URI based on actual environment, not just secrets flag
        # Check if we're explicitly on a production domain
        hostname = socket.getfqdn().lower()
        is_production_domain = (
            'blanklabelshop.com' in hostname or
            'streamlit.app' in hostname or
            os.environ.get('STREAMLIT_SERVER_HEADLESS') == 'true'
        )
        
        # For localhost, use web-based flow with localhost redirect URI
        # For deployed, use production redirect URI
        # Priority: explicit production domain > localhost (default)
        if is_production_domain:
            # Explicitly on production domain - use production redirect URI
            redirect_uri = oauth.get("workbook_redirect_uri") or "https://blanklabelshop.com/ads-dashboard/workbook_creator"
        else:
            # Not on production domain - assume localhost
            redirect_uri = oauth.get("workbook_redirect_uri") or "http://localhost:8501/workbook_creator"
        
        # Use web-based flow for both localhost and deployed environments
        # Ensure redirect_uri matches exactly what's registered in Google Cloud Console
        flow = Flow.from_client_config(
            {
                "web": {
                    "client_id": oauth["client_id"],
                    "client_secret": oauth["client_secret"],
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": [redirect_uri]
                }
            },
            scopes=scopes,
        )
        flow.redirect_uri = redirect_uri
        
        # Check for OAuth callback with authorization code
        query_params = st.query_params
        if "code" in query_params:
            # Check if we've already processed this code (prevent reuse on page reload)
            code_value = query_params.get("code")
            processed_code = st.session_state.get('wb_processed_code')
            
            if processed_code == code_value:
                # We've already processed this code, return existing credentials
                if 'wb_google_creds' in st.session_state and st.session_state.wb_google_creds:
                    return st.session_state.wb_google_creds
                # If no credentials, clear the processed code and try again
                st.session_state.wb_processed_code = None
            
            try:
                # Verify state parameter for security
                if "state" in query_params and st.session_state.get('wb_oauth_state'):
                    if query_params["state"] != st.session_state.get("wb_oauth_state"):
                        st.error("❌ Invalid OAuth state. Please try again.")
                        # Clear query params to prevent retry
                        st.query_params.clear()
                        return None
                
                # Exchange code for credentials
                # The redirect_uri is already set on the flow object
                flow.fetch_token(code=code_value)
                creds = flow.credentials
                
                # Store credentials in session state
                st.session_state.wb_google_creds = creds
                st.session_state.wb_processed_code = code_value
                
                # Clear OAuth state and query parameters immediately
                if 'wb_oauth_state' in st.session_state:
                    del st.session_state['wb_oauth_state']
                # Clear query params to prevent reuse
                st.query_params.clear()
                
                st.success("✅ Google Drive authentication successful!")
                return creds
                
            except Exception as token_error:
                error_msg = str(token_error)
                # Clear query params on error to prevent retry with same code
                st.query_params.clear()
                
                # Provide more helpful error message
                if "invalid_grant" in error_msg.lower():
                    st.error("❌ Authorization code expired or already used. Please try connecting again.")
                    # Clear any stored state
                    if 'wb_oauth_state' in st.session_state:
                        del st.session_state['wb_oauth_state']
                    if 'wb_processed_code' in st.session_state:
                        del st.session_state['wb_processed_code']
                else:
                    st.error(f"❌ Failed to exchange authorization code: {error_msg}")
                return None
        
        # Generate authorization URL
        try:
            auth_url, state = flow.authorization_url(
                access_type="offline",
                prompt="consent"
            )
            
            # Store the flow state for verification
            st.session_state['wb_oauth_state'] = state
            
            # Show authorization link
            st.markdown(f"[🔗 Google Sheets access]({auth_url})")
            return None
            
        except Exception as e:
            st.error(f"❌ OAuth authentication failed: {str(e)}")
            return None
        
        return None
        
    except Exception as e:
        st.warning(f"OAuth authentication failed: {str(e)}")
        return None


def _get_credentials():
    """Get credentials - prefer OAuth user credentials, fallback to service account."""
    # Try OAuth first (user credentials with real Drive storage)
    oauth_creds = load_oauth_credentials()
    if oauth_creds:
        return oauth_creds
    
    # Fallback to service account (but this has 0 bytes quota)
    if not SACredentials:
        raise GoogleSheetsUploadError(
            "Google API libraries not available. Please install google-auth-oauthlib and google-api-python-client."
        )
    
    info = st.secrets.get("google_service_account")
    if not info:
        raise GoogleSheetsUploadError(
            "Google OAuth credentials not available. Please connect your Google Drive account using the button above."
        )
    return SACredentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)


def _get_drive_targets() -> tuple[Optional[str], Optional[str]]:
    """
    Return (shared_drive_id, parent_folder_id) from secrets, if provided.
    
    If only parent_folder_id is set, it's treated as a regular folder (not a Shared Drive).
    If both are set and different, parent_folder_id is a folder within the Shared Drive.
    """
    cfg = st.secrets.get("google_drive", {}) or {}
    shared_drive_id = cfg.get("shared_drive_id") or cfg.get("drive_id")
    parent_folder_id = cfg.get("parent_folder_id") or cfg.get("folder_id")
    
    # If shared_drive_id equals parent_folder_id, it's likely a regular folder, not a Shared Drive
    if shared_drive_id == parent_folder_id and parent_folder_id:
        shared_drive_id = None
    
    return shared_drive_id, parent_folder_id


def _format_bytes(value: Optional[str | int]) -> str:
    """Return human readable file size."""
    try:
        num = int(value)
        for unit in ["bytes", "KB", "MB", "GB", "TB"]:
            if num < 1024:
                return f"{num:.2f} {unit}"
            num /= 1024
    except Exception:
        pass
    return str(value)


def _log_drive_quota(drive_service, shared_drive_id: Optional[str] = None) -> dict:
    """Fetch and log current Drive quota usage for debugging."""
    try:
        about_req = drive_service.about().get(
            fields="storageQuota(limit, usage, usageInDrive, usageInDriveTrash)"
        )
        about = about_req.execute()
        quota = about.get("storageQuota", {})
        limit = quota.get("limit")
        usage = quota.get("usage")
        usage_drive = quota.get("usageInDrive")
        usage_trash = quota.get("usageInDriveTrash")
        readable = {
            "limit": _format_bytes(limit),
            "usage": _format_bytes(usage),
            "usage_drive": _format_bytes(usage_drive),
            "usage_trash": _format_bytes(usage_trash),
        }
        msg = (
            "Google Drive quota — "
            f"limit: {readable['limit']}, "
            f"usage: {readable['usage']} "
            f"(drive: {readable['usage_drive']}, trash: {readable['usage_trash']})"
        )
        if shared_drive_id:
            msg += f" [shared_drive_id={shared_drive_id}]"
        st.info(msg)
        return readable
    except Exception as exc:  # pragma: no cover - best effort logging
        st.warning(f"Unable to fetch Drive quota info: {exc}")
        return {}


def cleanup_old_workbooks(
    drive_service,
    shared_drive_id: Optional[str],
    parent_folder_id: Optional[str],
    max_files: int = 25,
) -> tuple[int, int]:
    """Remove older Google Sheets created by this tool to avoid quota issues."""
    try:
        query = (
            "mimeType='application/vnd.google-apps.spreadsheet' "
            "and name contains 'Development Package'"
        )
        if parent_folder_id:
            query += f" and '{parent_folder_id}' in parents"
        list_kwargs = {
            "q": query,
            "orderBy": "createdTime desc",
            "fields": "files(id,name,createdTime)",
            "supportsAllDrives": True,
            "includeItemsFromAllDrives": True,
        }
        # Only use Shared Drive API parameters if we have a shared_drive_id
        # AND it's different from parent_folder_id (meaning it's actually a Shared Drive)
        if shared_drive_id and shared_drive_id != parent_folder_id:
            list_kwargs.update(
                {
                    "corpora": "drive",
                    "driveId": shared_drive_id,
                }
            )
        results = (
            drive_service.files()
            .list(**list_kwargs)
            .execute()
        )
        files = results.get("files", [])
        if len(files) <= max_files:
            return 0, len(files)

        deleted = 0
        for file_meta in files[max_files:]:
            try:
                drive_service.files().delete(
                    fileId=file_meta["id"], supportsAllDrives=bool(shared_drive_id)
                ).execute()
                deleted += 1
            except Exception as exc:  # pragma: no cover - best effort cleanup
                st.warning(f"Could not delete old Google Sheet {file_meta.get('name')}: {exc}")
        if deleted:
            st.info(f"Deleted {deleted} old Google Sheets to free space.")
        return deleted, len(files)
    except Exception as exc:  # pragma: no cover - best effort cleanup
        st.warning(f"Cleanup of old Google Sheets failed: {exc}")
        return 0, 0


def upload_workbook_to_google_sheet(
    workbook_bytes: bytes, sheet_name: str
) -> str:
    """
    Upload the XLSX workbook bytes to Google Drive as a Google Sheet.

    Returns the web URL for the newly created Google Sheet.
    """
    if not workbook_bytes:
        raise GoogleSheetsUploadError("Workbook data is empty; nothing to upload.")

    credentials = _get_credentials()
    shared_drive_id, parent_folder_id = _get_drive_targets()
    drive_service = build("drive", "v3", credentials=credentials)

    # Attempt to delete older spreadsheets so the service account does not exceed quota.
    cleanup_old_workbooks(drive_service, shared_drive_id, parent_folder_id)

    file_metadata = {
        "name": sheet_name or "Workbook Copy",
        "mimeType": "application/vnd.google-apps.spreadsheet",
    }
    if parent_folder_id:
        file_metadata["parents"] = [parent_folder_id]
    elif shared_drive_id:
        file_metadata["parents"] = [shared_drive_id]

    media = MediaIoBaseUpload(
        io.BytesIO(workbook_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False,
    )

    try:
        create_kwargs = {
            "body": file_metadata,
            "media_body": media,
            "fields": "id, webViewLink",
        }
        if shared_drive_id or parent_folder_id:
            create_kwargs["supportsAllDrives"] = True
        created_file = drive_service.files().create(**create_kwargs).execute()
    except HttpError as exc:  # pragma: no cover - network errors
        error_text = str(exc)
        if getattr(exc, "resp", None) and getattr(exc.resp, "status", None) == 403:
            if "storageQuotaExceeded" in error_text:
                raise GoogleSheetsUploadError(
                    "Google Drive storage quota has been exceeded for the service "
                    "account. Please delete older files or empty the Drive trash, "
                    "then try again."
                ) from exc
        raise GoogleSheetsUploadError(f"Google Drive upload failed: {exc}") from exc
    except Exception as exc:  # pragma: no cover - network errors
        raise GoogleSheetsUploadError(f"Google Drive upload failed: {exc}") from exc

    file_id = created_file.get("id")
    web_view = created_file.get("webViewLink")
    if not file_id:
        raise GoogleSheetsUploadError("Upload succeeded but Google Drive did not return a file ID.")

    return web_view or f"https://docs.google.com/spreadsheets/d/{file_id}"

