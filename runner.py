"""Scheduled orchestrator and runner for Outlook-to-OneDrive synchronization."""

import time
import logging
import signal
import sys
from typing import Dict, Any

from config import Config
from db import Database
from auth import MSALAuth
from graph_client import MicrosoftGraphClient
from email_cleaner import parse_and_format_date, clean_subject, clean_email_body
from summarizer import EmailSummarizer

logger = logging.getLogger("outlook_sync.runner")


class SyncRunner:
    def __init__(self, config: Config):
        self.config = config
        self.db = Database(db_path=config.db_path)
        self.auth = MSALAuth(
            client_id=config.azure_client_id,
            authority=config.azure_authority,
            scopes=config.msal_scopes,
            cache_file=config.token_cache_file,
        )
        self.graph_client = MicrosoftGraphClient(auth=self.auth)
        self.summarizer = EmailSummarizer(
            api_key=config.gemini_api_key,
            model_name=config.gemini_model,
        )
        self._stop_requested = False

    def request_stop(self, signum=None, frame=None):
        """Signal handler for graceful shutdown."""
        logger.info("Shutdown requested. Finishing current cycle and stopping...")
        self._stop_requested = True

    def run_sync_cycle(self) -> Dict[str, int]:
        """
        Execute a single synchronization pass:
        1. Fetch unread emails from target sender and sent emails to target recipient.
        2. Deduplicate using SQLite.
        3. Extract fields (Date, Subject, Content).
        4. Summarize via Gemini (gemini-3.5-flash-lite).
        5. Append row to Excel on OneDrive.
        6. Record in DB and optionally mark as read.
        """
        stats = {"found": 0, "processed": 0, "skipped_already_processed": 0, "failed": 0}

        logger.info(
            "Starting sync cycle for sender '%s' targeting file '%s' (%s)...",
            self.config.target_sender_email,
            self.config.onedrive_file_id,
            self.config.excel_worksheet_name,
        )

        try:
            incoming_emails = self.graph_client.fetch_unread_emails_from_sender(
                target_sender=self.config.target_sender_email
            )
            sent_emails = self.graph_client.fetch_sent_emails_to_recipient(
                target_recipient=self.config.target_recipient_email
            )
        except Exception as e:
            logger.error("Failed to fetch emails from Outlook: %s", e, exc_info=True)
            stats["failed"] += 1
            return stats

        emails = incoming_emails + sent_emails
        sent_message_ids = {email.get("id") for email in sent_emails}
        stats["found"] = len(emails)
        if not emails:
            logger.info("No matching unread emails found.")
            return stats

        for email in emails:
            if self._stop_requested:
                logger.info("Sync interrupted by shutdown request.")
                break

            msg_id = email.get("id")
            raw_subject = email.get("subject", "")
            sent_date_raw = email.get("sentDateTime") or email.get("receivedDateTime")
            sender_obj = email.get("from", {}).get("emailAddress", {})
            sender_email = sender_obj.get("address", self.config.target_sender_email)
            is_sent_email = msg_id in sent_message_ids
            if is_sent_email:
                sender_email = "me"

            # Step 1: Idempotency check via SQLite
            if self.db.is_processed(msg_id):
                logger.info("Email '%s' (ID: %s) already processed in database. Skipping.", raw_subject, msg_id[:16])
                stats["skipped_already_processed"] += 1
                continue

            try:
                # Step 2: Data Extraction & Cleaning
                formatted_date = parse_and_format_date(sent_date_raw)
                subject = clean_subject(raw_subject)
                email_content = clean_email_body(email.get("body"))

                logger.info("Processing email '%s' from %s (%s)...", subject, sender_email, formatted_date)

                # Step 3: Summarization via Gemini SDK (gemini-3.5-flash-lite)
                summary = self.summarizer.summarize(
                    email_content=email_content,
                    subject=subject,
                )

                # Step 4: Append row to OneDrive Excel workbook
                # Row Schema: Date | Subject | Email Content | Summary
                row_values = [
                    formatted_date,
                    subject,
                    email_content,
                    summary,
                ]

                appended = self.graph_client.append_row_to_excel(
                    file_id=self.config.onedrive_file_id,
                    worksheet_name=self.config.excel_worksheet_name,
                    row_values=row_values,
                )

                if appended:
                    # Step 5: Record in SQLite database
                    self.db.record_processed(
                        message_id=msg_id,
                        subject=subject,
                        sender=sender_email,
                        received_date=formatted_date,
                        summary=summary,
                        excel_row_appended=True,
                    )

                    # Step 6: Mark email as Read in Outlook if enabled
                    if self.config.mark_as_read and not is_sent_email:
                        self.graph_client.mark_email_as_read(msg_id)

                    stats["processed"] += 1
                    logger.info("Successfully processed and logged email: '%s'", subject)
                else:
                    logger.error("Failed to append row to Excel for email '%s'", subject)
                    stats["failed"] += 1

            except Exception as e:
                logger.error("Error processing email ID '%s': %s", msg_id, e, exc_info=True)
                stats["failed"] += 1

        logger.info(
            "Sync cycle finished: %d found, %d processed, %d skipped, %d failed.",
            stats["found"],
            stats["processed"],
            stats["skipped_already_processed"],
            stats["failed"],
        )
        return stats

    def run_daemon(self):
        """Run the synchronization continuously at the configured interval."""
        interval_seconds = max(60, int(self.config.poll_interval_hours * 3600))
        logger.info(
            "Starting scheduled daemon runner. Polling every %.2f hours (%d seconds)...",
            self.config.poll_interval_hours,
            interval_seconds,
        )

        signal.signal(signal.SIGINT, self.request_stop)
        signal.signal(signal.SIGTERM, self.request_stop)

        while not self._stop_requested:
            start_time = time.time()
            try:
                self.run_sync_cycle()
            except Exception as e:
                logger.error("Unexpected error in sync cycle: %s", e, exc_info=True)

            # Sleep in 1-second chunks to respond quickly to shutdown signals
            elapsed = time.time() - start_time
            sleep_needed = max(0, interval_seconds - elapsed)
            logger.info("Next sync scheduled in %.1f minutes.", sleep_needed / 60)

            sleep_end = time.time() + sleep_needed
            while time.time() < sleep_end and not self._stop_requested:
                time.sleep(1)

        logger.info("Daemon runner stopped.")
