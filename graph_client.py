"""Microsoft Graph API client for Outlook email polling and OneDrive Excel logging."""

import time
import logging
from typing import Optional, List, Dict, Any, Tuple
import requests
from auth import MSALAuth

logger = logging.getLogger("outlook_sync.graph")

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class MicrosoftGraphClient:
    def __init__(self, auth: MSALAuth):
        self.auth = auth
        self.session = requests.Session()

    def _get_headers(self) -> Dict[str, str]:
        """Retrieve a valid bearer access token and return request headers."""
        token = self.auth.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def _request_with_retry(
        self,
        method: str,
        url: str,
        max_retries: int = 4,
        initial_backoff: float = 2.0,
        **kwargs,
    ) -> requests.Response:
        """
        Execute an HTTP request with exponential backoff and 429 / 5xx retry handling.
        Respects the 'Retry-After' header if provided by Microsoft Graph.
        """
        backoff = initial_backoff
        for attempt in range(1, max_retries + 1):
            headers = self._get_headers()
            if "headers" in kwargs:
                headers.update(kwargs.pop("headers"))

            try:
                logger.debug("Graph API %s %s (attempt %d/%d)", method, url, attempt, max_retries)
                response = self.session.request(method, url, headers=headers, **kwargs)

                # Success codes
                if response.status_code in (200, 201, 204):
                    return response

                # Throttling / Rate limiting
                if response.status_code == 429:
                    retry_after = response.headers.get("Retry-After")
                    wait_time = float(retry_after) if retry_after else backoff
                    logger.warning(
                        "Graph API rate limited (429). Retrying after %.1f seconds...",
                        wait_time,
                    )
                    time.sleep(wait_time)
                    backoff *= 2
                    continue

                # Server error codes that warrant retry
                if response.status_code in (500, 502, 503, 504):
                    logger.warning(
                        "Graph API server error %d. Retrying in %.1f seconds...",
                        response.status_code,
                        backoff,
                    )
                    time.sleep(backoff)
                    backoff *= 2
                    continue

                # For client errors (400, 401, 403, 404), raise immediately
                response.raise_for_status()
                return response

            except requests.exceptions.RequestException as e:
                logger.warning(
                    "Network error during Graph API %s %s: %s (attempt %d/%d)",
                    method,
                    url,
                    e,
                    attempt,
                    max_retries,
                )
                if attempt == max_retries:
                    raise
                time.sleep(backoff)
                backoff *= 2

        raise RuntimeError(f"Exhausted {max_retries} retries calling Graph API {method} {url}")

    def fetch_unread_emails_from_sender(self, target_sender: str, top: int = 50) -> List[Dict[str, Any]]:
        """
        Fetch unread emails from the inbox and filter for the specified sender.
        Uses a resilient query strategy compatible with personal Outlook accounts.
        """
        target_sender = target_sender.strip().lower()
        url = f"{GRAPH_BASE_URL}/me/mailFolders/inbox/messages"

        # Universal filter: query unread messages ordered by received time
        params = {
            "$filter": "isRead eq false",
            "$select": "id,conversationId,subject,body,from,receivedDateTime,sentDateTime,isRead",
            "$orderby": "receivedDateTime desc",
            "$top": str(top),
        }

        logger.info("Fetching unread emails from Outlook inbox...")
        try:
            resp = self._request_with_retry("GET", url, params=params)
            data = resp.json()
            messages = data.get("value", [])
        except requests.exceptions.HTTPError as e:
            logger.error("Failed to query inbox messages: %s", e)
            raise

        matching_emails = []
        for msg in messages:
            from_obj = msg.get("from", {})
            sender_address = (
                from_obj.get("emailAddress", {}).get("address", "") if from_obj else ""
            ).strip().lower()

            if sender_address == target_sender:
                matching_emails.append(msg)

        logger.info(
            "Found %d total unread inbox messages; %d match target sender '%s'",
            len(messages),
            len(matching_emails),
            target_sender,
        )
        return matching_emails

    def fetch_sent_emails_to_recipient(
        self, target_recipient: str, top: int = 50
    ) -> List[Dict[str, Any]]:
        """Fetch sent messages whose To recipients include the configured address."""
        target_recipient = target_recipient.strip().lower()
        url = f"{GRAPH_BASE_URL}/me/mailFolders/sentitems/messages"
        params = {
            "$select": "id,conversationId,subject,body,toRecipients,sentDateTime",
            "$orderby": "sentDateTime desc",
            "$top": str(top),
        }

        logger.info("Fetching sent emails addressed to '%s'...", target_recipient)
        resp = self._request_with_retry("GET", url, params=params)
        messages = resp.json().get("value", [])

        matching_emails = []
        for msg in messages:
            recipients = msg.get("toRecipients") or []
            recipient_addresses = {
                (recipient.get("emailAddress", {}).get("address") or "").strip().lower()
                for recipient in recipients
            }
            if target_recipient in recipient_addresses:
                matching_emails.append(msg)

        logger.info(
            "Found %d sent messages; %d match target recipient '%s'",
            len(messages),
            len(matching_emails),
            target_recipient,
        )
        return matching_emails

    def mark_email_as_read(self, message_id: str) -> bool:
        """Mark an email message as read in Outlook."""
        url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
        try:
            self._request_with_retry("PATCH", url, json={"isRead": True})
            logger.debug("Successfully marked email %s as read.", message_id)
            return True
        except Exception as e:
            logger.warning("Failed to mark email %s as read: %s", message_id, e)
            return False

    def append_row_to_excel(
        self,
        file_id: str,
        worksheet_name: str,
        row_values: List[str],
    ) -> bool:
        """
        Append a new row of values to an Excel workbook on OneDrive.
        Dual-mode support:
        1. If an Excel Table exists in the worksheet, uses the Table Row API.
        2. Otherwise, calculates the usedRange and appends to the next available row,
           automatically adding header row if the worksheet is empty.
        
        row_values schema: [Date, Subject, Email Content, Summary]
        """
        base_sheet_url = f"{GRAPH_BASE_URL}/me/drive/items/{file_id}/workbook/worksheets/{worksheet_name}"

        # -------------------------------------------------------------
        # Mode 1: Check if an Excel Table exists in this worksheet
        # -------------------------------------------------------------
        tables_url = f"{base_sheet_url}/tables"
        try:
            tables_resp = self._request_with_retry("GET", tables_url)
            tables = tables_resp.json().get("value", [])
            if tables:
                table_id = tables[0].get("id") or tables[0].get("name")
                logger.info(
                    "Detected Excel Table '%s' in %s. Appending row via Table API...",
                    table_id,
                    worksheet_name,
                )
                add_row_url = f"{base_sheet_url}/tables/{table_id}/rows/add"
                self._request_with_retry("POST", add_row_url, json={"values": [row_values]})
                logger.info("Successfully appended row to Table '%s'", table_id)
                return True
        except Exception as e:
            logger.debug("Table API not applicable or no tables found (%s). Falling back to usedRange mode.", e)

        # -------------------------------------------------------------
        # Mode 2: usedRange Mode (works on raw worksheet cells)
        # -------------------------------------------------------------
        used_range_url = f"{base_sheet_url}/usedRange"
        header_row = ["Date", "Subject", "Email Content", "Summary"]

        try:
            used_range_resp = self._request_with_retry("GET", used_range_url)
            used_range = used_range_resp.json()
            row_count = used_range.get("rowCount", 0)
            existing_values = used_range.get("values", [])
        except requests.exceptions.HTTPError as e:
            # If sheet is completely empty, usedRange may return 404 ItemNotFound
            if e.response is not None and e.response.status_code == 404:
                row_count = 0
                existing_values = []
            else:
                raise

        # If sheet is completely blank, initialize with header row first
        if row_count == 0 or not existing_values:
            logger.info("Worksheet '%s' appears empty. Initializing headers in A1:D1...", worksheet_name)
            header_range_url = f"{base_sheet_url}/range(address='A1:D1')"
            self._request_with_retry("PATCH", header_range_url, json={"values": [header_row]})
            next_row = 2
        else:
            # Check if row 1 already has headers
            first_row = existing_values[0] if existing_values else []
            has_headers = any("date" in str(cell).lower() or "subject" in str(cell).lower() for cell in first_row)

            if not has_headers and row_count == 1 and all(not str(c).strip() for c in first_row):
                # Sheet only has 1 empty row
                logger.info("Writing header row to A1:D1...")
                header_range_url = f"{base_sheet_url}/range(address='A1:D1')"
                self._request_with_retry("PATCH", header_range_url, json={"values": [header_row]})
                next_row = 2
            else:
                next_row = row_count + 1

        target_address = f"A{next_row}:D{next_row}"
        logger.info("Appending new row to worksheet '%s' at address %s...", worksheet_name, target_address)

        write_range_url = f"{base_sheet_url}/range(address='{target_address}')"
        self._request_with_retry("PATCH", write_range_url, json={"values": [row_values]})
        logger.info("Successfully appended row at %s", target_address)
        return True
