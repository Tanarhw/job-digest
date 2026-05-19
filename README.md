# Job Digest

Daily email digest of new marketing job postings, personalized using Claude AI.

## What it does

- Searches **Indeed**, **LinkedIn**, and **Idealist** for Marketing Manager and Communications Manager roles in San Francisco
- Ranks results against a resume using Claude to surface the best matches
- Sends a clean HTML digest email every morning at 7am
- Learns from replies — respond in plain English to refine the search, or attach an updated resume to use for future rankings

## Setup

### 1. Install dependencies

```bash
pip3 install feedparser requests beautifulsoup4 anthropic python-docx pypdf \
             google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

### 2. Google Cloud credentials

- Enable the Gmail API in Google Cloud Console
- Create an OAuth 2.0 Desktop App credential
- Download the JSON and save it as `credentials.json` in this directory

### 3. Authorize Gmail

```bash
python3 authenticate.py
```

This opens a browser for a one-time login and saves `token.json`.

### 4. Add your resume

Save a plain text version of the candidate's resume as `resume.txt` in this directory.

### 5. Set environment variables

```bash
export ANTHROPIC_API_KEY='your-key-here'
```

### 6. Schedule with launchd (macOS)

Copy `com.jobdigest.plist` to `~/Library/LaunchAgents/`, fill in your `ANTHROPIC_API_KEY`, then:

```bash
launchctl load ~/Library/LaunchAgents/com.jobdigest.plist
```

## Files

| File | Purpose |
|---|---|
| `job_digest.py` | Main script |
| `authenticate.py` | One-time OAuth setup |
| `credentials.json` | Google OAuth credentials (**not committed**) |
| `token.json` | OAuth access token (**not committed**) |
| `resume.txt` | Candidate resume for ranking (**not committed**) |
| `seen_jobs.json` | Deduplication state (**not committed**) |
| `preferences.json` | Search preferences from replies (**not committed**) |
