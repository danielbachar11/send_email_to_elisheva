"""Unit and component tests for Outlook-to-OneDrive automation modules."""

import os
import unittest
import tempfile
import sqlite3
from unittest.mock import MagicMock
from pathlib import Path

from db import Database
from email_cleaner import parse_and_format_date, clean_subject, clean_email_body
from config import Config
from graph_client import MicrosoftGraphClient


class TestDatabaseDeduplication(unittest.TestCase):
    def setUp(self):
        # Create temporary database
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_emails.db")
        self.db = Database(db_path=self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_database_idempotency(self):
        msg_id = "AAMkAGI2T...AAA="
        subject = "Project Update"
        sender = "elishevaliber@gmail.com"
        date_str = "2026-09-30 12:00:00"
        summary = "Project update regarding Q3 milestones."

        # Should initially be False
        self.assertFalse(self.db.is_processed(msg_id))
        self.assertEqual(self.db.get_processed_count(), 0)

        # Record email
        self.db.record_processed(
            message_id=msg_id,
            subject=subject,
            sender=sender,
            received_date=date_str,
            summary=summary,
            excel_row_appended=True,
        )

        # Now is_processed should be True
        self.assertTrue(self.db.is_processed(msg_id))
        self.assertEqual(self.db.get_processed_count(), 1)

        # Recording the exact same ID again should not duplicate
        self.db.record_processed(
            message_id=msg_id,
            subject=subject,
            sender=sender,
            received_date=date_str,
            summary=summary,
            excel_row_appended=True,
        )
        self.assertEqual(self.db.get_processed_count(), 1)

        # Retrieve recent records
        recent = self.db.get_recent_processed(limit=5)
        self.assertEqual(len(recent), 1)
        self.assertEqual(recent[0]["message_id"], msg_id)
        self.assertEqual(recent[0]["subject"], subject)
        self.assertEqual(recent[0]["summary"], summary)


class TestEmailCleaner(unittest.TestCase):
    def test_date_formatting(self):
        # Standard ISO 8601 with Z
        iso_z = "2026-09-30T14:35:20Z"
        formatted = parse_and_format_date(iso_z)
        self.assertEqual(formatted, "2026-09-30 14:35:20")

        # Standard ISO 8601 with UTC offset
        iso_offset = "2026-09-30T10:15:00+02:00"
        formatted_offset = parse_and_format_date(iso_offset)
        self.assertEqual(formatted_offset, "2026-09-30 10:15:00")

        # None or empty fallback
        fallback = parse_and_format_date(None)
        self.assertRegex(fallback, r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")

    def test_subject_cleaning(self):
        self.assertEqual(clean_subject("   Important   Notice  \n\t"), "Important Notice")
        self.assertEqual(clean_subject(None), "(No Subject)")
        self.assertEqual(clean_subject(""), "(No Subject)")

    def test_html_body_cleaning(self):
        html_input = {
            "contentType": "html",
            "content": """
            <html>
                <head><style>p { color: red; }</style><script>alert('x');</script></head>
                <body>
                    <p>Dear Daniel,</p>
                    <p>Here is the weekly update &amp; summary for <b>Project Orion</b>.&nbsp;Everything is on track.</p>
                    <br/>
                    <div>Best regards,<br/>Elisheva</div>
                </body>
            </html>
            """,
        }
        text = clean_email_body(html_input)
        self.assertNotIn("alert", text)
        self.assertNotIn("color: red", text)
        self.assertIn("Dear Daniel,", text)
        self.assertIn("weekly update & summary for Project Orion", text)
        self.assertIn("Everything is on track.", text)
        self.assertIn("Elisheva", text)

    def test_plain_text_body(self):
        text_input = {
            "contentType": "text",
            "content": "Hello,\n\nPlease see the attachment.\n\nThanks!",
        }
        text = clean_email_body(text_input)
        self.assertEqual(text, "Hello,\n\nPlease see the attachment.\n\nThanks!")


class TestConfigurationValidation(unittest.TestCase):
    def test_validation_detects_placeholders(self):
        cfg = Config(
            azure_client_id="your_azure_client_id_here",
            gemini_api_key="your_gemini_api_key_here",
        )
        errors = cfg.validate()
        self.assertIn("AZURE_CLIENT_ID is missing or set to placeholder in .env", errors)
        self.assertIn("GEMINI_API_KEY is missing or set to placeholder in .env", errors)

    def test_validation_passes_with_keys(self):
        cfg = Config(
            azure_client_id="abc-123",
            gemini_api_key="AIzaSyTestKey",
            onedrive_file_id="5E0328FC-6FF0-4F3D-914A-83CEE6BFD4B5",
            excel_worksheet_name="Sheet1",
            target_sender_email="elishevaliber@gmail.com",
            target_recipient_email="recipient@example.com",
        )
        errors = cfg.validate()
        self.assertEqual(len(errors), 0)


class TestSentEmailFiltering(unittest.TestCase):
    def test_fetch_sent_emails_filters_to_recipient(self):
        client = MicrosoftGraphClient.__new__(MicrosoftGraphClient)
        response = MagicMock()
        response.json.return_value = {
            "value": [
                {
                    "id": "sent-1",
                    "toRecipients": [
                        {"emailAddress": {"address": "target@example.com"}}
                    ],
                },
                {
                    "id": "sent-2",
                    "toRecipients": [
                        {"emailAddress": {"address": "other@example.com"}}
                    ],
                },
            ]
        }
        client._request_with_retry = MagicMock(return_value=response)

        result = client.fetch_sent_emails_to_recipient(" TARGET@example.com ")

        self.assertEqual([message["id"] for message in result], ["sent-1"])
        client._request_with_retry.assert_called_once()
        request = client._request_with_retry.call_args
        self.assertEqual(request.args[0], "GET")
        self.assertIn("/me/mailFolders/sentitems/messages", request.args[1])


if __name__ == "__main__":
    unittest.main()
