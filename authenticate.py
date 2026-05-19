#!/usr/bin/env python3
"""
Run this once to authorize the app and generate token.json.
A browser window will open — log in as tanarmath@gmail.com and click Allow.
"""

from google_auth_oauthlib.flow import InstalledAppFlow
import os

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
]

BASE = os.path.dirname(__file__)
flow = InstalledAppFlow.from_client_secrets_file(
    os.path.join(BASE, "credentials.json"), SCOPES
)
creds = flow.run_local_server(port=0)

with open(os.path.join(BASE, "token.json"), "w") as f:
    f.write(creds.to_json())

print("Done! token.json saved. You can now run job_digest.py.")
