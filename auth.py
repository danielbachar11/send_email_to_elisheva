"""MSAL authentication manager for personal Microsoft accounts with token cache persistence."""

import os
import logging
from pathlib import Path
from typing import Optional, List
import msal

logger = logging.getLogger("outlook_sync.auth")


class MSALAuth:
    def __init__(
        self,
        client_id: str,
        authority: str = "https://login.microsoftonline.com/consumers",
        scopes: Optional[List[str]] = None,
        cache_file: str = "token_cache.bin",
    ):
        if not client_id or client_id == "your_azure_client_id_here":
            raise ValueError(
                "Invalid AZURE_CLIENT_ID. Please register an app in Microsoft Azure "
                "and set AZURE_CLIENT_ID in your .env file."
            )
        self.client_id = client_id
        self.authority = authority
        self.scopes = scopes or [
            "https://graph.microsoft.com/Mail.ReadWrite",
            "https://graph.microsoft.com/Files.ReadWrite",
            "offline_access",
        ]
        self.cache_file = Path(cache_file)
        self.cache = msal.SerializableTokenCache()
        self._load_cache()

        # Initialize PublicClientApplication with the persistent token cache
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            token_cache=self.cache,
        )

    def _load_cache(self) -> None:
        """Load serialized token cache from disk if available."""
        if self.cache_file.exists():
            try:
                with open(self.cache_file, "r", encoding="utf-8") as f:
                    self.cache.deserialize(f.read())
                logger.debug("Loaded MSAL token cache from %s", self.cache_file)
            except Exception as e:
                logger.warning("Failed to deserialize token cache from %s: %s", self.cache_file, e)

    def _save_cache(self) -> None:
        """Save serialized token cache back to disk if modified."""
        if self.cache.has_state_changed:
            try:
                with open(self.cache_file, "w", encoding="utf-8") as f:
                    f.write(self.cache.serialize())
                logger.debug("Saved updated MSAL token cache to %s", self.cache_file)
            except Exception as e:
                logger.error("Failed to persist token cache to %s: %s", self.cache_file, e)

    def get_access_token(self, interactive: bool = True) -> str:
        """
        Obtain a valid access token for Microsoft Graph.
        Attempts silent acquisition using cached refresh token first.
        If no cached token exists and interactive=True, initiates Device Code flow.
        """
        accounts = self.app.get_accounts()
        result = None

        if accounts:
            logger.info("Found cached account '%s'. Attempting silent token acquisition...", accounts[0].get("username"))
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                logger.info("Successfully acquired token silently using cached credentials.")
                return result["access_token"]
            else:
                logger.warning("Silent token acquisition failed. Refresh token may have expired or changed.")

        if not interactive:
            raise RuntimeError(
                "No valid cached token found and interactive mode is disabled. "
                "Please run 'python main.py --auth-only' in a terminal to authenticate first."
            )

        # Interactive Device Code Flow
        logger.info("Initiating Microsoft Device Code Flow for authentication...")
        flow = self.app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            error_desc = flow.get("error_description", flow.get("error", "Unknown error initiating device flow"))
            raise RuntimeError(f"Failed to initiate Device Code Flow: {error_desc}")

        user_code = flow["user_code"]
        verification_uri = flow.get("verification_uri", "https://microsoft.com/devicelogin")
        message = flow.get("message", f"Go to {verification_uri} and enter code: {user_code}")

        print("\n" + "=" * 76)
        print("  MICROSOFT ACCOUNT AUTHENTICATION REQUIRED")
        print("=" * 76)
        print(f"\n{message}\n")
        print("  1. Open a browser and visit: " + verification_uri)
        print(f"  2. Enter the one-time code:  {user_code}")
        print("  3. Sign in with your personal Outlook/Microsoft account and grant permissions.")
        print("=" * 76 + "\n")

        # Block until user completes authentication in browser
        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            self._save_cache()
            account_username = result.get("id_token_claims", {}).get("preferred_username", "Unknown")
            logger.info("Successfully authenticated as '%s' and saved token cache.", account_username)
            print(f"\nAuthentication successful for {account_username}! Token cached at {self.cache_file}.\n")
            return result["access_token"]
        else:
            error = result.get("error")
            desc = result.get("error_description")
            raise RuntimeError(f"Authentication failed: {error} - {desc}")
