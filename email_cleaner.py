"""Utility functions for parsing, cleaning, and formatting email attributes."""

import re
import html
from datetime import datetime, timezone
from typing import Optional, Dict, Any
from bs4 import BeautifulSoup


def parse_and_format_date(iso_date_str: Optional[str]) -> str:
    """
    Parse an ISO 8601 date string from Microsoft Graph (e.g. '2026-09-30T09:25:00Z')
    and return it formatted as 'YYYY-MM-DD HH:mm:ss'.
    """
    if not iso_date_str:
        return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")

    # Clean trailing 'Z' if present for standard fromisoformat parsing
    clean_str = iso_date_str.strip()
    if clean_str.endswith("Z"):
        clean_str = clean_str[:-1] + "+00:00"

    try:
        dt = datetime.fromisoformat(clean_str)
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        # Fallback if non-standard format is returned
        try:
            # Try basic split if ISO format has microseconds or unusual offset
            clean_str = clean_str.split(".")[0]
            dt = datetime.fromisoformat(clean_str)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def clean_subject(raw_subject: Optional[str]) -> str:
    """Clean and normalize email subject string."""
    if not raw_subject or not raw_subject.strip():
        return "(No Subject)"
    return " ".join(raw_subject.strip().split())


def clean_email_body(body_dict: Optional[Dict[str, Any]]) -> str:
    """
    Extract plain text from an Outlook email body object,
    handling HTML tags, entities, and line breaks cleanly.
    """
    if not body_dict:
        return ""

    content = body_dict.get("content", "") or ""
    content_type = (body_dict.get("contentType", "text") or "text").lower()

    if not content.strip():
        return ""

    if content_type == "html":
        # Parse HTML with BeautifulSoup
        soup = BeautifulSoup(content, "html.parser")

        # Remove irrelevant elements like scripts, styles, metadata
        for tag in soup(["script", "style", "head", "meta", "noscript", "svg"]):
            tag.decompose()

        # Add newline after block-level tags so text doesn't run together
        for block_tag in soup.find_all(["p", "div", "br", "tr", "li", "h1", "h2", "h3", "h4", "h5", "h6"]):
            block_tag.append("\n")

        text = soup.get_text()
    else:
        text = content

    # Unescape HTML entities (e.g. &amp;, &nbsp;, &lt;)
    text = html.unescape(text)

    # Replace non-breaking spaces with standard space
    text = text.replace("\xa0", " ")

    # Normalize line breaks and compress excessive blank lines
    text = re.sub(r"\r\n|\r", "\n", text)
    # Replace sequences of 3+ newlines with 2 newlines
    text = re.sub(r"\n{3,}", "\n\n", text)
    # Clean whitespace per line
    lines = [line.strip() for line in text.split("\n")]
    cleaned_text = "\n".join(lines).strip()

    return cleaned_text
