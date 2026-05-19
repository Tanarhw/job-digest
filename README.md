# Job Digest

A personalized daily job digest that emails curated listings every morning, ranked against a resume using Claude AI. Learns from plain-English reply feedback and adapts as the recipient updates their resume.

## Features

- Pulls jobs from **LinkedIn** and **Remotive** (remote roles)
- **Resume-based ranking** — Claude scores each listing against the candidate's resume and filters out weak matches
- **Salary, work arrangement, and job type** shown as badges on each card
- **Reply to refine** — respond in plain English ("skip contract roles", "focus on nonprofits") and tomorrow's digest adjusts automatically
- **Attach an updated resume** in a reply and it becomes the new ranking baseline
- Deduplication — only new listings are sent each day
- Scheduled via macOS launchd (7am daily)

## Setup

### 1. Install dependencies

```bash
pip3 install requests beautifulsoup4 anthropic python-docx pypdf \
             google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

### 2. Configure `job_digest.py`

Edit the config block at the top:

```python
RECIPIENT_NAME    = "Your Name"          # used in the email greeting
BASE_SEARCH_TERMS = ["marketing manager", "communications manager"]
LOCATION          = "San Francisco, CA"
RECIPIENT_EMAIL   = "recipient@gmail.com"
SENDER_EMAIL      = "sender@gmail.com"
```

### 3. Google Cloud credentials

- Enable the Gmail API in [Google Cloud Console](https://console.cloud.google.com)
- Create an OAuth 2.0 **Desktop app** credential
- Download the JSON and save it as `credentials.json` in this directory
- Add yourself as a test user under OAuth consent screen → Test users

### 4. Authorize Gmail (one-time)

```bash
python3 authenticate.py
```

Opens a browser for a Google login and saves `token.json`.

### 5. Add a resume

Save the candidate's resume as `resume.txt` in this directory. The script uses it to rank job matches. The recipient can also email an updated PDF or DOCX resume as a reply to any digest and it will be used from the next run onward.

### 6. Set environment variables

```bash
export ANTHROPIC_API_KEY='your-key-here'
```

### 7. Schedule with launchd (macOS)

Copy `com.jobdigest.plist` to `~/Library/LaunchAgents/`, fill in `ANTHROPIC_API_KEY`, then:

```bash
launchctl load ~/Library/LaunchAgents/com.jobdigest.plist
```

### Test run

```bash
ANTHROPIC_API_KEY='your-key' python3 job_digest.py
```

## File reference

| File | Purpose |
|---|---|
| `job_digest.py` | Main script |
| `authenticate.py` | One-time Gmail OAuth setup |
| `credentials.json` | Google OAuth credentials (**not committed**) |
| `token.json` | OAuth access token (**not committed**) |
| `resume.txt` | Candidate resume for ranking (**not committed**) |
| `seen_jobs.json` | Deduplication state (**not committed**) |
| `preferences.json` | Accumulated search preferences from replies (**not committed**) |
