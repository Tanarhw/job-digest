#!/usr/bin/env python3
"""
Daily job digest — sends new marketing job postings to meganhalpin98@gmail.com
Supports reply-based preference refinement via Claude API.
"""

import io
import os
import json
import base64
import hashlib
import time
import requests
import anthropic
import docx
from pypdf import PdfReader
from bs4 import BeautifulSoup
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ── Config ────────────────────────────────────────────────────────────────────
RECIPIENT_NAME    = "Megan"
BASE_SEARCH_TERMS = ["marketing manager", "communications manager"]
LOCATION          = "San Francisco, CA"
RECIPIENT_EMAIL   = "meganhalpin98@gmail.com"
SENDER_EMAIL      = "tanarmath@gmail.com"
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
BASE_DIR          = os.path.dirname(__file__)
SEEN_JOBS_FILE    = os.path.join(BASE_DIR, "seen_jobs.json")
PREFERENCES_FILE  = os.path.join(BASE_DIR, "preferences.json")
CREDENTIALS_FILE  = os.path.join(BASE_DIR, "credentials.json")
TOKEN_FILE        = os.path.join(BASE_DIR, "token.json")
RESUME_FILE       = os.path.join(BASE_DIR, "resume.txt")
SEARCH_CONFIG_FILE = os.path.join(BASE_DIR, "search_config.json")
RADIUS_MILES      = 25
GMAIL_SCOPES      = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
]
# ──────────────────────────────────────────────────────────────────────────────


# ── Seen jobs ─────────────────────────────────────────────────────────────────

def load_seen() -> set:
    if os.path.exists(SEEN_JOBS_FILE):
        with open(SEEN_JOBS_FILE) as f:
            return set(json.load(f))
    return set()


def save_seen(seen: set):
    with open(SEEN_JOBS_FILE, "w") as f:
        json.dump(list(seen), f)


# ── Preferences ───────────────────────────────────────────────────────────────

DEFAULT_PREFS = {
    "extra_search_terms": [],
    "exclude_keywords":   [],
    "notes":              "",
}


def load_preferences() -> dict:
    if os.path.exists(PREFERENCES_FILE):
        with open(PREFERENCES_FILE) as f:
            return json.load(f)
    return DEFAULT_PREFS.copy()


def save_preferences(prefs: dict):
    with open(PREFERENCES_FILE, "w") as f:
        json.dump(prefs, f, indent=2)


# ── Resume-derived search terms ───────────────────────────────────────────────

def derive_search_terms() -> list[str]:
    """Ask Claude to pick the best job title search terms for this resume. Cached in search_config.json."""
    if not os.path.exists(RESUME_FILE):
        return BASE_SEARCH_TERMS

    with open(RESUME_FILE) as f:
        resume = f.read()

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    prompt = f"""Based on this resume, suggest the 4-6 best job title search terms to use when hunting for roles on LinkedIn and job boards.

Consider: current seniority, transferable skills, realistic next steps (lateral moves and one step up), and any niche strengths.

Return only a JSON array of strings — job title search terms, no explanation.
Example: ["Senior Marketing Manager", "Director of Marketing", "Head of Content"]

RESUME:
{resume}"""

    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-opus-4-7",
                max_tokens=400,
                messages=[{"role": "user", "content": prompt}],
            )
            break
        except Exception as ex:
            print(f"[Claude] search terms attempt {attempt+1} failed: {ex}")
            if attempt < 2:
                time.sleep(10)
            else:
                return BASE_SEARCH_TERMS

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1].lstrip("json").strip() if len(parts) > 1 else raw

    terms = json.loads(raw)
    with open(SEARCH_CONFIG_FILE, "w") as f:
        json.dump({"search_terms": terms}, f, indent=2)
    print(f"[Search terms] Derived from resume: {terms}")
    return terms


def load_search_terms() -> list[str]:
    if os.path.exists(SEARCH_CONFIG_FILE):
        with open(SEARCH_CONFIG_FILE) as f:
            return json.load(f).get("search_terms", BASE_SEARCH_TERMS)
    return BASE_SEARCH_TERMS


# ── Gmail API client ──────────────────────────────────────────────────────────

def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, GMAIL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


# ── Reply inbox check ─────────────────────────────────────────────────────────

def _all_parts(payload: dict) -> list[dict]:
    """Recursively flatten all MIME parts from a message payload."""
    result = []
    if "parts" in payload:
        for part in payload["parts"]:
            result.extend(_all_parts(part))
    else:
        result.append(payload)
    return result


def _extract_resume_text(data: bytes, filename: str) -> str:
    if filename.lower().endswith(".docx"):
        doc = docx.Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)
    elif filename.lower().endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return ""


def check_reply_emails(service) -> tuple[list[str], bool]:
    """Check for unread replies from Megan. Returns (reply_texts, resume_was_updated)."""
    try:
        query = 'is:unread from:meganhalpin98@gmail.com subject:"Re: Job digest"'
        result = service.users().messages().list(userId="me", q=query).execute()
        messages = result.get("messages", [])

        reply_texts = []
        resume_updated = False

        for m in messages:
            msg = service.users().messages().get(
                userId="me", id=m["id"], format="full"
            ).execute()

            service.users().messages().modify(
                userId="me", id=m["id"], body={"removeLabelIds": ["UNREAD"]}
            ).execute()

            parts = _all_parts(msg.get("payload", {}))

            # Extract plain text body
            body = ""
            for part in parts:
                if part.get("mimeType") == "text/plain" and not part.get("filename"):
                    data = part.get("body", {}).get("data", "")
                    if data:
                        body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
                        break
            # Fallback for non-multipart
            if not body:
                data = msg.get("payload", {}).get("body", {}).get("data", "")
                if data:
                    body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
            if body.strip():
                reply_texts.append(body.strip())

            # Check for resume attachment
            for part in parts:
                filename = part.get("filename", "")
                if filename.lower().endswith((".pdf", ".docx")):
                    att_id = part.get("body", {}).get("attachmentId")
                    if att_id:
                        att = service.users().messages().attachments().get(
                            userId="me", messageId=m["id"], id=att_id
                        ).execute()
                        raw = base64.urlsafe_b64decode(att["data"])
                        text = _extract_resume_text(raw, filename)
                        if text.strip():
                            with open(RESUME_FILE, "w") as f:
                                f.write(text)
                            resume_updated = True
                            print(f"[Resume] Updated from attachment: {filename}")

        return reply_texts, resume_updated
    except Exception as ex:
        print(f"[Gmail API] error checking replies: {ex}")
        return [], False


# ── Claude preference parser ───────────────────────────────────────────────────

def parse_preferences_from_reply(reply_text: str, current_prefs: dict) -> dict:
    """Use Claude to merge a free-form reply into structured search preferences."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    prompt = f"""You help refine a job search digest based on feedback emails.

Current search preferences:
- Extra search terms: {current_prefs.get('extra_search_terms', [])}
- Exclude keywords: {current_prefs.get('exclude_keywords', [])}
- Notes: {current_prefs.get('notes', '')}

Megan replied to her job digest email with:
---
{reply_text}
---

Produce updated preferences as JSON with exactly these keys:
- extra_search_terms: list of additional job title/type strings to search for
- exclude_keywords: list of words or phrases — jobs containing these should be filtered out
- notes: 1-2 sentence plain-English summary of her current refined preferences

Merge intelligently with current preferences. Add new items; remove items only if she explicitly says to drop something. Return only valid JSON, no markdown, no explanation."""

    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-opus-4-7",
                max_tokens=600,
                messages=[{"role": "user", "content": prompt}],
            )
            break
        except Exception as ex:
            print(f"[Claude] prefs attempt {attempt+1} failed: {ex}")
            if attempt < 2:
                time.sleep(10)
            else:
                return current_prefs

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1].lstrip("json").strip() if len(parts) > 1 else raw
    try:
        return json.loads(raw)
    except Exception as ex:
        print(f"[Claude] failed to parse preferences JSON: {ex}\nRaw: {raw}")
        return current_prefs


# ── Job fetchers ──────────────────────────────────────────────────────────────

def job_id(title: str, company: str, url: str) -> str:
    return hashlib.md5(f"{title.lower()}{company.lower()}{url}".encode()).hexdigest()


def fetch_remotive(term: str) -> list[dict]:
    url = f"https://remotive.com/api/remote-jobs?category=marketing&search={requests.utils.quote(term)}&limit=20"
    try:
        resp = requests.get(url, timeout=15)
        jobs = []
        for j in resp.json().get("jobs", []):
            jobs.append({
                "title":    j.get("title", ""),
                "company":  j.get("company_name", "Unknown"),
                "url":      j.get("url", ""),
                "location": j.get("candidate_required_location") or "Remote",
                "source":   "Remotive",
                "summary":  BeautifulSoup(j.get("description", ""), "html.parser").get_text()[:200].strip(),
                "_salary":  j.get("salary", ""),
                "_job_type": j.get("job_type", "").replace("-", " ").title(),
                "_work_arrangement": "Remote",
            })
        return jobs
    except Exception as ex:
        print(f"[Remotive] error for '{term}': {ex}")
        return []


def fetch_linkedin(term: str, location: str) -> list[dict]:
    keywords = term.replace(" ", "%20")
    loc      = location.replace(" ", "%20").replace(",", "%2C")
    url      = (
        "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
        f"?keywords={keywords}&location={loc}&start=0"
    )
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        soup = BeautifulSoup(resp.text, "html.parser")
        jobs = []
        for card in soup.find_all("li"):
            title_el   = card.find("h3", class_="base-search-card__title")
            company_el = card.find("h4", class_="base-search-card__subtitle")
            loc_el     = card.find("span", class_="job-search-card__location")
            link_el    = card.find("a", class_="base-card__full-link")
            if not (title_el and link_el):
                continue
            jobs.append({
                "title":    title_el.get_text(strip=True),
                "company":  company_el.get_text(strip=True) if company_el else "Unknown",
                "url":      link_el.get("href", "").split("?")[0],
                "location": loc_el.get_text(strip=True) if loc_el else location,
                "source":   "LinkedIn",
                "summary":  "",
            })
        return jobs
    except Exception as ex:
        print(f"[LinkedIn] error for '{term}': {ex}")
        return []


def fetch_idealist(term: str) -> list[dict]:
    q   = term.replace(" ", "+")
    url = f"https://www.idealist.org/en/jobs?q={q}&loc=San+Francisco%2C+CA%2C+US"
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        resp   = requests.get(url, headers=headers, timeout=15)
        soup   = BeautifulSoup(resp.text, "html.parser")
        script = soup.find("script", id="__NEXT_DATA__")
        if not script:
            print("[Idealist] __NEXT_DATA__ not found — site structure may have changed")
            return []

        data = json.loads(script.string)
        hits = (
            data.get("props", {})
                .get("pageProps", {})
                .get("initialData", {})
                .get("hits", [])
        )

        jobs = []
        for item in hits:
            slug = item.get("slug", "")
            jobs.append({
                "title":    item.get("name", ""),
                "company":  item.get("org", {}).get("name", "Unknown"),
                "url":      f"https://www.idealist.org/en/job/{slug}",
                "location": item.get("location", "San Francisco, CA"),
                "source":   "Idealist",
                "summary":  (item.get("description") or "")[:200].strip(),
            })
        return jobs
    except Exception as ex:
        print(f"[Idealist] error for '{term}': {ex}")
        return []


def gather_all_jobs(search_terms: list[str]) -> list[dict]:
    all_jobs = []
    for term in search_terms:
        print(f"Fetching: {term}")
        all_jobs += fetch_linkedin(term, LOCATION)
        all_jobs += fetch_remotive(term)
        all_jobs += fetch_idealist(term)
        time.sleep(1)
    return all_jobs


# ── Resume-based ranking ──────────────────────────────────────────────────────

def rank_jobs_by_fit(jobs: list[dict]) -> list[dict]:
    """Use Claude to score and filter jobs against Megan's resume. Returns sorted list with match notes."""
    if not ANTHROPIC_API_KEY or not os.path.exists(RESUME_FILE):
        return jobs

    with open(RESUME_FILE) as f:
        resume = f.read()

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    job_list = "\n".join(
        f"[{i}] {j['title']} at {j['company']} ({j['source']})\n    {j['summary']}"
        for i, j in enumerate(jobs)
    )

    prompt = f"""You are a career coach helping filter and enrich job listings for a candidate based on their resume.

RESUME:
{resume}

JOB LISTINGS (indexed 0 to {len(jobs)-1}):
{job_list}

For each job, return a JSON object with:
- index: the job's index number
- score: 1-10 fit score (seniority, skills match, sector relevance)
- reason: max 8 words on why it fits her background
- work_arrangement: "Remote", "Hybrid", "On-site", or "Flexible" — infer from title/description
- job_type: "Full-time", "Contract", "Part-time", or "" if unclear

Return a JSON array sorted best-to-worst, including only jobs with score >= 6.
Example: [{{"index": 0, "score": 9, "reason": "Senior digital marketing, nonprofit sector match", "work_arrangement": "Hybrid", "job_type": "Full-time"}}]

Return only valid JSON, no explanation."""

    for attempt in range(3):
        try:
            response = client.messages.create(
                model="claude-opus-4-7",
                max_tokens=1500,
                messages=[{"role": "user", "content": prompt}],
            )
            break
        except Exception as ex:
            print(f"[Claude] ranking attempt {attempt+1} failed: {ex}")
            if attempt < 2:
                time.sleep(10)
            else:
                return jobs

    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        parts = raw.split("```")
        raw = parts[1].lstrip("json").strip() if len(parts) > 1 else raw

    try:
        ranked = json.loads(raw)
        result = []
        for item in ranked:
            job = dict(jobs[item["index"]])
            job["_match_reason"] = item.get("reason", "")
            job["_match_score"]  = item.get("score", 0)
            # Only overwrite arrangement/type if not already set by the source API
            if not job.get("_work_arrangement"):
                job["_work_arrangement"] = item.get("work_arrangement", "")
            if not job.get("_job_type"):
                job["_job_type"] = item.get("job_type", "")
            result.append(job)
        print(f"[Claude] Ranked {len(jobs)} jobs → {len(result)} good matches")
        return result
    except Exception as ex:
        print(f"[Claude] ranking failed: {ex}")
        return jobs


# ── Filtering ─────────────────────────────────────────────────────────────────

def filter_new(jobs: list[dict], seen: set, exclude_keywords: list[str]) -> list[dict]:
    new_jobs, seen_this_run = [], set()
    exclude_lower = [kw.lower() for kw in exclude_keywords]
    for job in jobs:
        jid = job_id(job["title"], job["company"], job["url"])
        if jid in seen or jid in seen_this_run:
            continue
        if exclude_lower:
            job_text = f"{job['title']} {job['company']} {job['summary']}".lower()
            if any(kw in job_text for kw in exclude_lower):
                continue
        job["_id"] = jid
        new_jobs.append(job)
        seen_this_run.add(jid)
    return new_jobs


# ── Email builder ─────────────────────────────────────────────────────────────

SOURCE_COLORS = {
    "LinkedIn": "#0A66C2",
    "Remotive": "#16A34A",
    "Idealist": "#EA580C",
}

SOURCE_BG = {
    "LinkedIn": "#EFF6FF",
    "Remotive": "#F0FDF4",
    "Idealist": "#FFF7ED",
}


def build_email_html(jobs: list[dict], prefs: dict) -> str:
    today = datetime.now().strftime("%B %d, %Y")
    count = len(jobs)
    noun  = "job" if count == 1 else "jobs"

    by_source: dict[str, list] = {}
    for job in jobs:
        by_source.setdefault(job["source"], []).append(job)

    def pill(text: str, bg: str, fg: str) -> str:
        return (f'<span style="display:inline-block;padding:3px 10px;background:{bg};color:{fg};'
                f'font-size:11px;font-weight:600;border-radius:99px;margin-right:6px;">{text}</span>')

    def job_card(j: dict, color: str) -> str:
        summary_html = (
            f'<p style="margin:10px 0 0;color:#6B7280;font-size:13px;line-height:1.6;">{j["summary"]}</p>'
            if j["summary"] else ""
        )

        badges = ""
        if j.get("_work_arrangement"):
            badges += pill(j["_work_arrangement"], "#EFF6FF", "#1D4ED8")
        if j.get("_job_type"):
            badges += pill(j["_job_type"], "#F3F4F6", "#374151")
        if j.get("_salary"):
            badges += pill(j["_salary"], "#FEF9C3", "#854D0E")
        badges_html = f'<div style="margin-top:10px;">{badges}</div>' if badges else ""

        match_badge = ""
        if j.get("_match_reason"):
            match_badge = pill(j["_match_reason"], "#DCFCE7", "#15803D")

        return f"""
        <div style="background:#FFFFFF;border:1px solid #E5E7EB;border-left:4px solid {color};
                    border-radius:8px;padding:18px 20px;margin-bottom:12px;">
          <a href="{j['url']}" style="color:#111827;text-decoration:none;font-size:15px;
                                      font-weight:700;line-height:1.3;">{j['title']}</a>
          <p style="margin:5px 0 0;color:#6B7280;font-size:13px;">
            <span style="font-weight:600;color:#374151;">{j['company']}</span>
            &nbsp;&middot;&nbsp;{j['location']}
          </p>
          {badges_html}
          {summary_html}
          <div style="margin-top:12px;">
            {match_badge}
          </div>
          <div style="margin-top:12px;">
            <a href="{j['url']}" style="display:inline-block;padding:7px 16px;background:{color};
               color:#fff;font-size:12px;font-weight:600;text-decoration:none;border-radius:6px;">
              View role &rarr;
            </a>
          </div>
        </div>"""

    def source_section(source: str, listings: list) -> str:
        color = SOURCE_COLORS.get(source, "#555")
        bg    = SOURCE_BG.get(source, "#F9FAFB")
        cards = "".join(job_card(j, color) for j in listings)
        badge = f'<span style="background:{color};color:#fff;font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;vertical-align:middle;">{len(listings)}</span>'
        return f"""
        <div style="margin-bottom:32px;">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;
                      padding:10px 14px;background:{bg};border-radius:8px;">
            <span style="font-size:13px;font-weight:700;color:{color};letter-spacing:.04em;
                         text-transform:uppercase;">{source}</span>
            &nbsp;{badge}
          </div>
          {cards}
        </div>"""

    sections = "".join(source_section(src, lst) for src, lst in by_source.items())

    prefs_html = ""
    has_prefs = prefs.get("extra_search_terms") or prefs.get("exclude_keywords") or prefs.get("notes")
    if has_prefs:
        extra = ", ".join(prefs["extra_search_terms"]) if prefs["extra_search_terms"] else "—"
        excl  = ", ".join(prefs["exclude_keywords"])   if prefs["exclude_keywords"]   else "—"
        notes = prefs.get("notes", "")
        notes_row = f'<p style="margin:6px 0 0;font-size:12px;color:#6B7280;font-style:italic;">{notes}</p>' if notes else ""
        prefs_html = f"""
        <div style="background:#F0F9FF;border:1px solid #BAE6FD;border-radius:8px;
                    padding:14px 18px;margin-bottom:28px;">
          <p style="margin:0 0 8px;font-size:11px;font-weight:700;color:#0369A1;
                    letter-spacing:.08em;text-transform:uppercase;">Active search preferences</p>
          <p style="margin:0;font-size:12px;color:#374151;">
            <span style="color:#6B7280;">Also searching:</span> {extra}
          </p>
          <p style="margin:4px 0 0;font-size:12px;color:#374151;">
            <span style="color:#6B7280;">Filtering out:</span> {excl}
          </p>
          {notes_row}
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#F3F4F6;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,sans-serif;">
  <div style="max-width:620px;margin:0 auto;padding:32px 16px 48px;">

    <!-- Header -->
    <div style="background:linear-gradient(135deg,#1E3A8A 0%,#2563EB 100%);
                border-radius:12px 12px 0 0;padding:32px 36px 28px;">
      <p style="margin:0 0 4px;font-size:12px;font-weight:600;color:#93C5FD;
                letter-spacing:.1em;text-transform:uppercase;">Daily Job Digest</p>
      <h1 style="margin:0 0 6px;font-size:26px;font-weight:700;color:#FFFFFF;line-height:1.2;">
        Good morning, {RECIPIENT_NAME}!
      </h1>
      <p style="margin:0;font-size:14px;color:#BFDBFE;">
        {today} &mdash; <strong style="color:#fff;">{count} new {noun}</strong> matched to your profile
      </p>
    </div>

    <!-- Body -->
    <div style="background:#FFFFFF;border-radius:0 0 12px 12px;padding:32px 36px;
                border:1px solid #E5E7EB;border-top:none;">
      {prefs_html}
      {sections}

      <!-- Footer -->
      <div style="border-top:1px solid #F3F4F6;margin-top:8px;padding-top:24px;">
        <p style="margin:0 0 6px;font-size:13px;color:#374151;text-align:center;font-weight:600;">
          Want different results?
        </p>
        <p style="margin:0;font-size:13px;color:#6B7280;text-align:center;line-height:1.6;">
          Reply in plain English &mdash; e.g. &ldquo;focus more on tech&rdquo; or
          &ldquo;skip healthcare&rdquo; &mdash; and tomorrow&rsquo;s digest will adjust automatically.<br>
          You can also <strong style="color:#374151;">attach an updated resume</strong> and I&rsquo;ll use it for future rankings.
        </p>
        <p style="margin:16px 0 0;font-size:11px;color:#9CA3AF;text-align:center;">
          Searching Marketing Manager &amp; Communications Manager &bull; San Francisco + Remote
        </p>
      </div>
    </div>

  </div>
</body>
</html>"""


# ── Email send ────────────────────────────────────────────────────────────────

def send_email(service, subject: str, html: str, to: str = RECIPIENT_EMAIL):
    msg            = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = to
    msg.attach(MIMEText(html, "html"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(
        userId="me", body={"raw": raw}
    ).execute()


# ── Main ──────────────────────────────────────────────────────────────────────

def wait_for_network(timeout: int = 120):
    """Wait until network is stable — checks both Google and Anthropic endpoints."""
    import socket
    targets = [("api.anthropic.com", 443), ("www.googleapis.com", 443)]
    deadline = time.time() + timeout
    while time.time() < deadline:
        if all(_can_connect(h, p) for h, p in targets):
            time.sleep(15)  # extra buffer for connections to fully stabilize
            return
        time.sleep(5)
    print("[Network] Timed out waiting for network — proceeding anyway.")


def _can_connect(host: str, port: int) -> bool:
    import socket
    try:
        socket.create_connection((host, port), timeout=5).close()
        return True
    except OSError:
        return False


def main():
    wait_for_network()
    service = get_gmail_service()

    # 1. Check for reply emails, resume attachments, and update preferences
    prefs = load_preferences()
    if ANTHROPIC_API_KEY:
        replies, resume_updated = check_reply_emails(service)
        if resume_updated:
            print("[Resume] Resume updated — re-deriving search terms.")
            derive_search_terms()
        for reply_text in replies:
            print(f"[Prefs] Parsing reply: {reply_text[:80]}...")
            prefs = parse_preferences_from_reply(reply_text, prefs)
            save_preferences(prefs)
            print(f"[Prefs] Updated: {prefs}")
    else:
        print("[Prefs] ANTHROPIC_API_KEY not set — skipping reply parsing.")

    # 2. Build full search term list from resume-derived terms + any reply additions
    if ANTHROPIC_API_KEY and not os.path.exists(SEARCH_CONFIG_FILE) and os.path.exists(RESUME_FILE):
        derive_search_terms()
    base_terms = load_search_terms()
    all_terms  = base_terms + prefs.get("extra_search_terms", [])

    # 3. Gather, deduplicate, filter
    seen     = load_seen()
    all_jobs = gather_all_jobs(all_terms)
    new_jobs = filter_new(all_jobs, seen, prefs.get("exclude_keywords", []))

    if not new_jobs:
        print(f"[{datetime.now():%Y-%m-%d %H:%M}] No new jobs, skipping email.")
        return

    # 4. Rank by resume fit
    new_jobs = rank_jobs_by_fit(new_jobs)

    if not new_jobs:
        print(f"[{datetime.now():%Y-%m-%d %H:%M}] No strong matches after ranking, skipping email.")
        return

    # 5. Build and send
    html    = build_email_html(new_jobs, prefs)
    subject = f"Job digest: {len(new_jobs)} new posting{'s' if len(new_jobs) != 1 else ''} — {datetime.now():%b %d}"
    send_email(service, subject, html)

    seen.update(j["_id"] for j in new_jobs)
    save_seen(seen)
    print(f"[{datetime.now():%Y-%m-%d %H:%M}] Sent {len(new_jobs)} new jobs to {RECIPIENT_EMAIL}")


if __name__ == "__main__":
    main()
