"""Command Line Interface (CLI) entrypoint for Outlook-to-OneDrive automation."""

import argparse
import sys
import logging
from pathlib import Path

from config import config
from runner import SyncRunner
from db import Database
from auth import MSALAuth


def setup_logging(verbose: bool = False, log_file: str = "sync.log") -> None:
    """Configure console and file logging."""
    log_level = logging.DEBUG if verbose else logging.INFO
    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)-7s] [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # File handler
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)  # Keep full debug logs in the file
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)


def handle_auth_only():
    """Perform one-time interactive Microsoft authentication to seed the token cache."""
    print("Initiating Microsoft account authentication...")
    errors = config.validate(check_azure=True, check_gemini=False)
    if errors:
        print("\nConfiguration Errors:")
        for err in errors:
            print(f"  - {err}")
        print("\nPlease update your .env file with your AZURE_CLIENT_ID first.\n")
        sys.exit(1)

    auth = MSALAuth(
        client_id=config.azure_client_id,
        authority=config.azure_authority,
        scopes=config.msal_scopes,
        cache_file=config.token_cache_file,
    )
    token = auth.get_access_token(interactive=True)
    if token:
        print("Success! Access token obtained and refresh token saved to cache.")
        print(f"Token cache stored at: {Path(config.token_cache_file).resolve()}")
        print("You can now run scheduled tasks headlessly without further prompts.\n")


def handle_status():
    """Print the current synchronization and database status."""
    db = Database(db_path=config.db_path)
    count = db.get_processed_count()
    recent = db.get_recent_processed(limit=5)

    print("\n" + "=" * 60)
    print("  OUTLOOK-TO-ONEDRIVE AUTOMATION STATUS")
    print("=" * 60)
    print(f"Target Sender:       {config.target_sender_email}")
    print(f"Target OneDrive ID:  {config.onedrive_file_id}")
    print(f"Worksheet Name:      {config.excel_worksheet_name}")
    print(f"Gemini Model:        {config.gemini_model}")
    print(f"Poll Interval:       {config.poll_interval_hours} hour(s)")
    print(f"Database Path:       {Path(config.db_path).resolve()}")
    print(f"Token Cache:         {Path(config.token_cache_file).resolve()} (exists: {Path(config.token_cache_file).exists()})")
    print(f"Total Processed:     {count} email(s)")
    print("-" * 60)

    if recent:
        print("Recent Processed Emails:")
        for i, item in enumerate(recent, 1):
            print(f" {i}. [{item['received_date']}] {item['subject']}")
            print(f"    Summary: {item['summary']}")
            print(f"    Excel Appended: {'Yes' if item['excel_row_appended'] else 'No'}")
    else:
        print("No emails processed yet.")
    print("=" * 60 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Monitor Outlook inbox and log incoming emails to Excel on OneDrive."
    )
    parser.add_argument(
        "--run-once",
        action="store_true",
        help="Execute a single synchronization pass and exit (ideal for cron jobs).",
    )
    parser.add_argument(
        "--daemon",
        action="store_true",
        help="Run continuously in the background at the configured interval.",
    )
    parser.add_argument(
        "--auth-only",
        action="store_true",
        help="Perform initial interactive Microsoft device code authentication.",
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Display SQLite tracking statistics and recent logged emails.",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Enable detailed debug logging in console.",
    )

    args = parser.parse_args()

    setup_logging(verbose=args.verbose)

    if args.status:
        handle_status()
        return

    if args.auth_only:
        handle_auth_only()
        return

    # Validate configuration for full runs
    validation_errors = config.validate(check_azure=True, check_gemini=True)
    if validation_errors:
        print("\n" + "!" * 60)
        print("CONFIGURATION ERROR: Please configure your .env file:")
        for err in validation_errors:
            print(f"  • {err}")
        print("!" * 60)
        print("Refer to .env.example or README.md for instructions.\n")
        sys.exit(1)

    runner = SyncRunner(config=config)

    if args.daemon:
        runner.run_daemon()
    else:
        # Default behavior: run-once
        stats = runner.run_sync_cycle()
        print(f"\nExecution summary: {stats}\n")


if __name__ == "__main__":
    main()
