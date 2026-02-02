"""
Utility helpers to upload generated workbooks to Google Sheets.
"""

from __future__ import annotations

import io
from typing import Optional, Tuple

import streamlit as st

try:
    from google.oauth2.service_account import Credentials as SACredentials
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseUpload
except Exception:
    SACredentials = None
    build = None
    HttpError = None
    MediaIoBaseUpload = None

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
]


class GoogleSheetsUploadError(Exception):
    """Raised when the workbook cannot be uploaded to Google Sheets."""


def _get_credentials():
    """Load ONLY service account credentials from Streamlit secrets."""
    if not SACredentials:
        raise GoogleSheetsUploadError(
            "Google API libraries not available. Please install google-auth and google-api-python-client."
        )

    info = st.secrets.get("google_service_account")
    if not info:
        raise GoogleSheetsUploadError(
            "Google Cloud service account credentials missing in secrets. "
            "Add `google_service_account` to your Streamlit secrets."
        )

    return SACredentials.from_service_account_info(
        info,
        scopes=GOOGLE_SCOPES
    )


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


def _is_storage_quota_error(exc: HttpError) -> bool:
    """Return True if the Drive API error is a storage quota issue."""
    if getattr(exc, "resp", None) is None:
        return False
    if getattr(exc.resp, "status", None) != 403:
        return False
    error_text = str(exc)
    return "storageQuotaExceeded" in error_text


def upload_workbook_to_google_sheet(
    workbook_bytes: bytes, sheet_name: str
) -> Tuple[str, bool]:
    """
    Upload the XLSX workbook bytes to Google Drive.

    Returns (web_url, converted_to_google_sheet).
    We first upload the file as XLSX so it stores under the shared folder owner's quota.
    Then we attempt to convert it to a native Google Sheet; on quota failures we keep the XLSX.
    """
    if not workbook_bytes:
        raise GoogleSheetsUploadError("Workbook data is empty; nothing to upload.")

    credentials = _get_credentials()
    shared_drive_id, parent_folder_id = _get_drive_targets()
    drive_service = build("drive", "v3", credentials=credentials)

    # If we have a parent_folder_id but no shared_drive_id, check if the folder is in a Shared Drive
    if parent_folder_id and not shared_drive_id:
        try:
            folder_info = drive_service.files().get(
                fileId=parent_folder_id,
                fields="id, name, driveId",
                supportsAllDrives=True
            ).execute()
            folder_drive_id = folder_info.get("driveId")
            if folder_drive_id:
                # This folder is in a Shared Drive, use the drive ID
                shared_drive_id = folder_drive_id
        except Exception:
            # If we can't get folder info, continue without Shared Drive support
            pass

    file_metadata = {
        "name": sheet_name or "Workbook Copy",
        # Keep the file as XLSX initially so it uses the shared folder owner's quota.
        "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
        # Use supportsAllDrives for Shared Drives
        if shared_drive_id:
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

    converted = False
    parents = file_metadata.get("parents")
    try:
        copy_body = {
            "name": sheet_name or "Workbook Copy",
            "mimeType": "application/vnd.google-apps.spreadsheet",
        }
        if parents:
            copy_body["parents"] = parents
        copy_kwargs = {
            "fileId": file_id,
            "body": copy_body,
            "fields": "id, webViewLink",
        }
        # Use supportsAllDrives for Shared Drives
        if shared_drive_id:
            copy_kwargs["supportsAllDrives"] = True
        converted_file = drive_service.files().copy(**copy_kwargs).execute()
        new_id = converted_file.get("id")
        new_link = converted_file.get("webViewLink")
        if new_id:
            # Delete the original XLSX to avoid clutter.
            try:
                delete_kwargs = {"fileId": file_id}
                if shared_drive_id:
                    delete_kwargs["supportsAllDrives"] = True
                drive_service.files().delete(**delete_kwargs).execute()
            except Exception:
                pass
            file_id = new_id
            web_view = new_link or web_view
            converted = True
    except HttpError as exc:
        if _is_storage_quota_error(exc):
            st.warning(
                "Google Sheets conversion failed because the service account has zero Drive quota. "
                "The XLSX workbook is still uploaded to the shared folder."
            )
        else:
            st.warning(f"Google Sheets conversion failed: {exc}")

    final_url = web_view or f"https://drive.google.com/file/d/{file_id}/view"
    return final_url, converted

