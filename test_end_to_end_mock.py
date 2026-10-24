"""End-to-end integration test with mocked external APIs."""

import os
import unittest
import tempfile
from unittest.mock import MagicMock, patch

from config import Config
from runner import SyncRunner


class TestSyncCycleMocked(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = os.path.join(self.temp_dir.name, "test_sync.db")
        self.cache_file = os.path.join(self.temp_dir.name, "token_cache.bin")

        self.config = Config(
            azure_client_id="mock-client-id",
            azure_authority="https://login.microsoftonline.com/consumers",
            token_cache_file=self.cache_file,
            gemini_api_key="mock-gemini-key",
            gemini_model="gemini-3.5-flash-lite",
            onedrive_file_id="5E0328FC-6FF0-4F3D-914A-83CEE6BFD4B5",
            excel_worksheet_name="Sheet1",
            target_sender_email="elishevaliber@gmail.com",
            target_recipient_email="recipient@example.com",
            mark_as_read=True,
            db_path=self.db_path,
        )

    def tearDown(self):
        self.temp_dir.cleanup()

    @patch("graph_client.requests.Session")
    @patch("auth.msal.PublicClientApplication")
    @patch("summarizer.genai.Client")
    def test_full_sync_pipeline(self, mock_genai_client_class, mock_msal_app_class, mock_session_class):
        # Setup mock GenAI client
        mock_genai_client = MagicMock()
        mock_response = MagicMock()
        mock_response.text = "This is a concise summary of the incoming proposal."
        mock_genai_client.models.generate_content.return_value = mock_response
        mock_genai_client_class.return_value = mock_genai_client

        # Setup mock MSAL
        mock_msal_app = MagicMock()
        mock_msal_app.get_accounts.return_value = [{"username": "testuser@outlook.com"}]
        mock_msal_app.acquire_token_silent.return_value = {"access_token": "mock-token-xyz"}
        mock_msal_app_class.return_value = mock_msal_app

        runner = SyncRunner(config=self.config)

        # Mock Graph Client methods
        sample_email = {
            "id": "MSG_123456789",
            "subject": "Q3 Budget Review",
            "sentDateTime": "2026-09-30T10:15:00Z",
            "from": {
                "emailAddress": {
                    "name": "Elisheva Liber",
                    "address": "elishevaliber@gmail.com",
                }
            },
            "body": {
                "contentType": "html",
                "content": "<p>Hi Daniel,</p><p>Please find attached the Q3 budget review figures for our review meeting.</p>",
            },
        }

        # Mock fetch unread emails
        runner.graph_client.fetch_unread_emails_from_sender = MagicMock(return_value=[sample_email])
        runner.graph_client.fetch_sent_emails_to_recipient = MagicMock(return_value=[])
        runner.graph_client.append_row_to_excel = MagicMock(return_value=True)
        runner.graph_client.mark_email_as_read = MagicMock(return_value=True)

        # Execute first cycle
        stats = runner.run_sync_cycle()
        self.assertEqual(stats["found"], 1)
        self.assertEqual(stats["processed"], 1)
        self.assertEqual(stats["skipped_already_processed"], 0)
        self.assertEqual(stats["failed"], 0)

        # Verify Excel append call
        runner.graph_client.append_row_to_excel.assert_called_once()
        call_args = runner.graph_client.append_row_to_excel.call_args[1]
        self.assertEqual(call_args["file_id"], "5E0328FC-6FF0-4F3D-914A-83CEE6BFD4B5")
        self.assertEqual(call_args["worksheet_name"], "Sheet1")
        row = call_args["row_values"]
        self.assertEqual(row[0], "2026-09-30 10:15:00")  # Formatted Date
        self.assertEqual(row[1], "Q3 Budget Review")      # Clean Subject
        self.assertIn("Q3 budget review figures", row[2]) # Clean Content
        self.assertEqual(row[3], "This is a concise summary of the incoming proposal.")  # Summary

        # Verify marked as read
        runner.graph_client.mark_email_as_read.assert_called_once_with("MSG_123456789")

        # Verify recorded in DB
        self.assertTrue(runner.db.is_processed("MSG_123456789"))

        # Execute second cycle with the same email returned
        stats_second = runner.run_sync_cycle()
        self.assertEqual(stats_second["found"], 1)
        self.assertEqual(stats_second["processed"], 0)
        self.assertEqual(stats_second["skipped_already_processed"], 1)
        self.assertEqual(stats_second["failed"], 0)


if __name__ == "__main__":
    unittest.main()
