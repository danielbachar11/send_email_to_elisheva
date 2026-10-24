"""Configuration loader and validator for Outlook-to-OneDrive automation."""

import os
from dataclasses import dataclass
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from project root or current working directory
load_dotenv()


@dataclass
class Config:
    # Azure MSAL Configuration
    azure_client_id: str = os.getenv("AZURE_CLIENT_ID", "").strip()
    azure_authority: str = os.getenv("AZURE_AUTHORITY", "https://login.microsoftonline.com/consumers").strip()
    token_cache_file: str = os.getenv("TOKEN_CACHE_FILE", "token_cache.bin").strip()

    # Google Gemini Configuration
    gemini_api_key: str = os.getenv("GEMINI_API_KEY", "").strip()
    gemini_model: str = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite").strip()

    # Target OneDrive Excel Configuration
    onedrive_file_id: str = os.getenv("ONEDRIVE_FILE_ID", "5E0328FC-6FF0-4F3D-914A-83CEE6BFD4B5").strip()
    excel_worksheet_name: str = os.getenv("EXCEL_WORKSHEET_NAME", "Sheet1").strip()

    # Email Filter and Processing
    target_sender_email: str = os.getenv("TARGET_SENDER_EMAIL", "elishevaliber@gmail.com").strip().lower()
    target_recipient_email: str = os.getenv("TARGET_RECIPIENT_EMAIL", "").strip().lower()
    mark_as_read: bool = os.getenv("MARK_AS_READ", "true").strip().lower() in ("true", "1", "yes")
    db_path: str = os.getenv("DB_PATH", "processed_emails.db").strip()

    # Scheduler Settings
    poll_interval_hours: float = float(os.getenv("POLL_INTERVAL_HOURS", "1"))

    # Required MSAL Delegated Scopes for Microsoft Graph API:
    # Mail.Read / Mail.ReadWrite for personal mailbox access
    # Files.ReadWrite for Excel workbook manipulation in personal OneDrive
    # offline_access for persistent refresh token retention
    @property
    def msal_scopes(self) -> list[str]:
        scopes = [
            "https://graph.microsoft.com/Files.ReadWrite",
            
        ]
        if self.mark_as_read:
            scopes.append("https://graph.microsoft.com/Mail.ReadWrite")
        else:
            scopes.append("https://graph.microsoft.com/Mail.Read")
        return scopes

    def validate(self, check_azure: bool = True, check_gemini: bool = True) -> list[str]:
        """Validate critical configuration settings and return list of missing keys."""
        errors = []

        if check_azure:
            if not self.azure_client_id or self.azure_client_id == "your_azure_client_id_here":
                errors.append("AZURE_CLIENT_ID is missing or set to placeholder in .env")

        if check_gemini:
            if not self.gemini_api_key or self.gemini_api_key == "your_gemini_api_key_here":
                errors.append("GEMINI_API_KEY is missing or set to placeholder in .env")

        if not self.onedrive_file_id:
            errors.append("ONEDRIVE_FILE_ID is missing in .env")

        if not self.excel_worksheet_name:
            errors.append("EXCEL_WORKSHEET_NAME is missing in .env")

        if not self.target_sender_email:
            errors.append("TARGET_SENDER_EMAIL is missing in .env")

        if not self.target_recipient_email:
            errors.append("TARGET_RECIPIENT_EMAIL is missing in .env")

        return errors


# Global config singleton
config = Config()
