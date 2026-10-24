"""SQLite database operations for email deduplication and idempotency tracking."""

import sqlite3
import logging
from typing import Optional, List, Dict, Any
from datetime import datetime

from contextlib import contextmanager

logger = logging.getLogger("outlook_sync.db")


class Database:
    def __init__(self, db_path: str = "processed_emails.db"):
        self.db_path = db_path
        self.init_db()

    @contextmanager
    def get_connection(self):
        """Context manager yielding SQLite connection and guaranteeing closure."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()

    def init_db(self) -> None:
        """Initialize the database schema if it doesn't already exist."""
        with self.get_connection() as conn:
            cursor = conn.cursor()

            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS processed_emails (
                    message_id TEXT PRIMARY KEY,
                    subject TEXT,
                    sender TEXT,
                    received_date TEXT,
                    summary TEXT,
                    processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    excel_row_appended INTEGER DEFAULT 1
                )
                """
            )
            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_processed_emails_sender 
                ON processed_emails (sender)
                """
            )
            cursor.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_processed_emails_date 
                ON processed_emails (received_date)
                """
            )
            conn.commit()
        logger.debug("Database initialized at %s", self.db_path)

    def is_processed(self, message_id: str) -> bool:
        """Check if an email message ID has already been successfully processed."""
        if not message_id:
            return False
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT 1 FROM processed_emails WHERE message_id = ? LIMIT 1",
                (message_id,),
            )
            return cursor.fetchone() is not None

    def record_processed(
        self,
        message_id: str,
        subject: str,
        sender: str,
        received_date: str,
        summary: str,
        excel_row_appended: bool = True,
    ) -> None:
        """Record an email as processed in the database."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT OR REPLACE INTO processed_emails 
                (message_id, subject, sender, received_date, summary, processed_at, excel_row_appended)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    message_id,
                    subject,
                    sender,
                    received_date,
                    summary,
                    datetime.now().isoformat(),
                    1 if excel_row_appended else 0,
                ),
            )
            conn.commit()
        logger.debug("Recorded email %s in processed database", message_id)

    def get_processed_count(self) -> int:
        """Return the total number of processed emails in database."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM processed_emails")
            return cursor.fetchone()[0]

    def get_recent_processed(self, limit: int = 10) -> List[Dict[str, Any]]:
        """Retrieve recent processed emails for auditing/status reporting."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT message_id, subject, sender, received_date, summary, processed_at, excel_row_appended
                FROM processed_emails
                ORDER BY processed_at DESC
                LIMIT ?
                """,
                (limit,),
            )
            return [dict(row) for row in cursor.fetchall()]
