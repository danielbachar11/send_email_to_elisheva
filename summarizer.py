"""Google Gemini email summarizer module using Google GenAI SDK."""

import time
import logging
from typing import Optional
from google import genai
from google.genai.errors import APIError

logger = logging.getLogger("outlook_sync.summarizer")


class EmailSummarizer:
    def __init__(self, api_key: str, model_name: str = "gemini-3.5-flash-lite"):
        if not api_key:
            raise ValueError("GEMINI_API_KEY must be provided to initialize EmailSummarizer.")
        self.api_key = api_key
        self.model_name = model_name
        self.client = genai.Client(api_key=self.api_key)

    def summarize(
        self,
        email_content: str,
        subject: str = "",
        max_retries: int = 3,
        initial_backoff: float = 2.0,
    ) -> str:
        """
        Generate a concise 1-2 sentence summary of the email content using Google GenAI SDK.
        Includes retry logic for API rate limits and transient connection issues.
        """
        clean_content = (email_content or "").strip()
        if not clean_content:
            return "Empty email content."

        # If email is extremely short (e.g., under 15 words), a full summary call might just echo it
        if len(clean_content.split()) <= 10 and not subject:
            return clean_content

        prompt = (
            "You are an executive assistant processing emails. "
            "Generate a clear, factual, and concise 1 to 2 sentence summary of the email below. "
            "Output ONLY the summary sentences without preamble, quotes, bullet points, or markdown formatting.\n\n"
            f"Subject: {subject}\n"
            f"Content:\n{clean_content}"
        )

        backoff = initial_backoff
        for attempt in range(1, max_retries + 1):
            try:
                logger.debug(
                    "Calling Gemini model '%s' (attempt %d/%d)...",
                    self.model_name,
                    attempt,
                    max_retries,
                )
                response = self.client.models.generate_content(
                    model=self.model_name,
                    contents=prompt,
                )

                if response and response.text:
                    summary = response.text.strip()
                    # Clean any accidental quotes
                    if summary.startswith('"') and summary.endswith('"') and len(summary) > 2:
                        summary = summary[1:-1].strip()
                    logger.info("Successfully generated summary via Gemini (%d chars)", len(summary))
                    return summary
                else:
                    logger.warning("Empty response received from Gemini for email subject '%s'", subject)
                    return "(No summary generated)"

            except APIError as e:
                logger.warning(
                    "Gemini API error on attempt %d/%d: %s. Retrying in %.1fs...",
                    attempt,
                    max_retries,
                    e,
                    backoff,
                )
                if attempt == max_retries:
                    logger.error("Exhausted retries calling Gemini API: %s", e)
                    raise
                time.sleep(backoff)
                backoff *= 2
            except Exception as e:
                logger.warning(
                    "Unexpected error calling Gemini API on attempt %d/%d: %s. Retrying in %.1fs...",
                    attempt,
                    max_retries,
                    e,
                    backoff,
                )
                if attempt == max_retries:
                    logger.error("Exhausted retries calling Gemini API: %s", e)
                    raise
                time.sleep(backoff)
                backoff *= 2

        return "(Summary generation failed)"
