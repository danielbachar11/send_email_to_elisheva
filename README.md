# Outlook to OneDrive Excel Automation

A robust, scheduled automation in Python that monitors a personal Microsoft Outlook account, filters incoming emails from a specific sender and sent emails addressed to a specific recipient, generates concise 1–2 sentence summaries using Google Gemini (`gemini-3.5-flash-lite`), and appends the structured data directly into an Excel workbook hosted on OneDrive.

---

## Architecture Overview

```
                                      +--------------------------+
                                      | Personal Outlook Mailbox |
                                      +-------------+------------+
                                                    |
                                                    | Polls incoming mail from
                                                    | TARGET_SENDER_EMAIL and
                                                    | sent mail to
                                                    | TARGET_RECIPIENT_EMAIL
                                                    v
+--------------------------+          +-------------+------------+
|  Local SQLite Database   |<-------->|       SyncRunner         |
|  (Deduplication Store)   |  Checks  | (Orchestrator & Filter)  |
+--------------------------+  ID      +-------------+------------+
                                                    |
                                                    | Extracted Content & Subject
                                                    v
                                      +-------------+------------+
                                      |   Google GenAI SDK       |
                                      |  (gemini-3.5-flash-lite) |
                                      +-------------+------------+
                                                    |
                                                    | 1-2 Sentence Summary
                                                    v
                                      +-------------+------------+
                                      | Microsoft Graph Excel    |
                                      | (Personal OneDrive Item) |
                                      +--------------------------+
```

### Key Technical Highlights:
1. **Personal Microsoft Account Authentication (MSAL Device Code Flow)**:
   - Personal consumer accounts (`@outlook.com`, `@hotmail.com`, etc.) do not support app-only client credentials.
   - Uses MSAL with **Device Code Flow** and a persistent token cache (`token_cache.bin`).
   - You only authenticate **once** in your browser. Refresh tokens are automatically managed and renewed silently for all future scheduled executions.
2. **Robust Deduplication & Idempotency**:
   - Every processed email is tracked in a local SQLite database (`processed_emails.db`) by its unique Outlook `message_id`.
   - Prevents duplicate rows even if emails remain unread or network interruptions occur.
3. **Data Extraction & AI Summarization**:
   - Extracts and normalizes:
     - `Date`: Formatted as `YYYY-MM-DD HH:mm:ss`.
     - `Subject`: Cleaned string.
     - `Email Content`: HTML sanitized and converted to clean text via BeautifulSoup.
     - `Summary`: 1–2 sentence executive summary generated via Google GenAI (`gemini-3.5-flash-lite`).
4. **Excel Target Schema (`Sheet1`)**:
   - Target File ID: `5E0328FC-6FF0-4F3D-914A-83CEE6BFD4B5`
   - Target Worksheet: `Sheet1`
   - Column layout: `Date` | `Subject` | `Email Content` | `Summary`
   - Dual-mode Excel appender: Works seamlessly whether you use an Excel formatted Table or standard worksheet cells (auto-creates headers if blank).
5. **Production Reliability**:
   - Exponential backoff with jitter and retry logic for network transients and API throttling (handles HTTP 429 `Retry-After` headers and 5xx errors).

---

## Directory Structure

```
outlook-onedrive-sync/
├── .env.example              # Configuration template
├── requirements.txt          # Python dependencies
├── config.py                 # Configuration loader and validator
├── auth.py                   # MSAL Authenticator with persistent cache
├── db.py                     # SQLite idempotency tracking database
├── email_cleaner.py          # HTML/text sanitization and date parser
├── summarizer.py             # Google GenAI summarizer (gemini-3.5-flash-lite)
├── graph_client.py           # Microsoft Graph API client for Outlook & Excel
├── runner.py                 # Orchestrator and polling loop
├── main.py                   # CLI entrypoint (--run-once, --daemon, --auth-only, --status)
├── test_components.py        # Unit tests
├── test_end_to_end_mock.py   # Integration mock tests
└── systemd/                  # Background service template
    └── outlook-sync.service
```

---

## Step 1: Azure App Registration Setup

Personal Microsoft accounts require a registered application in Microsoft Entra (Azure Portal) with delegated permissions. Follow these steps:

1. **Sign In to Azure Portal**:
   - Go to [Microsoft Azure Portal](https://portal.azure.com/) or [Microsoft Entra Admin Center](https://entra.microsoft.com/).
   - Sign in with your personal Microsoft account.

2. **Register a New Application**:
   - In the search bar, search for **App registrations** and select it.
   - Click **+ New registration**.
   - Fill in:
     - **Name**: `Outlook OneDrive Excel Sync` (or any friendly name).
     - **Supported account types**: Select **"Personal Microsoft accounts only"** (or *"Accounts in any organizational directory and personal Microsoft accounts"*).
     - **Redirect URI**: Select platform **Public client/native (mobile & desktop)** and enter:
       `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Click **Register**.

3. **Enable Public Client Flows**:
   - In the left menu of your new app, click **Authentication**.
   - Scroll down to **Advanced settings** -> **Allow public client flows**.
   - Set **"Enable the following mobile and desktop flows"** to **Yes**.
   - Click **Save** at the top.

4. **Configure API Permissions**:
   - In the left menu, click **API permissions**.
   - Click **+ Add a permission** -> Select **Microsoft Graph** -> **Delegated permissions**.
   - Check the following permissions:
     - `Mail.Read` (or `Mail.ReadWrite` to allow marking processed emails as read)
     - `Files.ReadWrite` (to read and append rows to your OneDrive Excel workbook)
     - `offline_access` (to maintain a persistent refresh token for background execution)
   - Click **Add permissions**.

5. **Copy the Client ID**:
   - Click **Overview** in the left menu.
   - Copy the **Application (client) ID** (a GUID format like `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`).

---

## Step 2: Google Gemini API Key Setup

1. Visit [Google AI Studio](https://aistudio.google.com/).
2. Sign in with your Google account.
3. Click **Get API key** in the left navigation panel.
4. Click **Create API key** (create in a new project or select an existing Google Cloud project).
5. Copy your generated API key (e.g. `AIzaSy...`).

---

## Step 3: Installation & Configuration

1. **Activate Virtual Environment & Install Dependencies**:
   ```bash
   cd /home/daniel/.gemini/antigravity/scratch/outlook-onedrive-sync
   source venv/bin/activate
   pip install -r requirements.txt
   ```

2. **Create and Edit `.env`**:
   Copy the example template:
   ```bash
   cp .env.example .env
   ```
   Open `.env` in your editor and fill in your keys:
   ```env
   AZURE_CLIENT_ID=your_azure_client_id_from_step_1
   AZURE_AUTHORITY=https://login.microsoftonline.com/consumers
   TOKEN_CACHE_FILE=token_cache.bin

   GEMINI_API_KEY=your_gemini_api_key_from_step_2
   GEMINI_MODEL=gemini-3.5-flash-lite

   ONEDRIVE_FILE_ID=5E0328FC-6FF0-4F3D-914A-83CEE6BFD4B5
   EXCEL_WORKSHEET_NAME=Sheet1

   TARGET_SENDER_EMAIL=elishevaliber@gmail.com
   TARGET_RECIPIENT_EMAIL=recipient@example.com
   MARK_AS_READ=true
   DB_PATH=processed_emails.db

   POLL_INTERVAL_HOURS=1
   ```

---

## Step 4: One-Time Authentication (Device Code Flow)

To grant your script permission to access your personal Outlook and OneDrive without exposing a web server:

```bash
python main.py --auth-only
```

You will see an output like:
```
============================================================================
  MICROSOFT ACCOUNT AUTHENTICATION REQUIRED
============================================================================

  1. Open a browser and visit: https://microsoft.com/devicelogin
  2. Enter the one-time code:  ABCD-EFGH
  3. Sign in with your personal Outlook/Microsoft account and grant permissions.
============================================================================
```

- Follow the prompt on your phone, laptop, or browser.
- Once completed, your credentials and refresh token will be stored in `token_cache.bin`.
- **You will not need to log in again.** Silent refresh runs automatically on schedule.

---

## Step 5: Running the Automation

### 1. Manual Single Run (Test / Cron)
Runs one pass, processes any matching unread emails, and exits:
```bash
python main.py --run-once
```

### 2. Continuous Background Daemon
Monitors in a continuous loop every `POLL_INTERVAL_HOURS`:
```bash
python main.py --daemon
```

### 3. Check Database & Sync Status
View how many emails have been logged and see recent entries:
```bash
python main.py --status
```

---

## Scheduling Options

### Option A: Linux Cron Job (Recommended for headless runners)
To run every 1 hour via cron:
1. Open the crontab editor:
   ```bash
   crontab -e
   ```
2. Add the following line (pointing to the project's venv python):
   ```cron
   0 * * * * cd /home/daniel/.gemini/antigravity/scratch/outlook-onedrive-sync && ./venv/bin/python main.py --run-once >> sync.log 2>&1
   ```
   *(To run every 3 hours, replace `0 * * * *` with `0 */3 * * *`)*.

### Option B: Systemd Service (For continuous daemon)
1. Copy the provided unit file:
   ```bash
   sudo cp systemd/outlook-sync.service /etc/systemd/system/
   ```
2. Reload and enable:
   ```bash
   sudo systemctl daemon-reload
   sudo systemctl enable --now outlook-sync.service
   ```
3. Check status & logs:
   ```bash
   sudo systemctl status outlook-sync.service
   journalctl -u outlook-sync.service -f
   ```

---

## Testing & Verification

Run the built-in test suite to verify database deduplication, HTML cleaning, and mocked end-to-end sync:
```bash
./venv/bin/python test_components.py
./venv/bin/python test_end_to_end_mock.py
```
Both suites run locally without requiring live external credentials.
# send_email_to_elisheva
