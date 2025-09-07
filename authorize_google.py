import os
import sys
import pathlib
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Google Drive nur-lesend
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

BASE = pathlib.Path(__file__).parent.resolve()
CREDENTIALS = BASE / "credentials.json"
TOKEN = BASE / "token.json"

def main():
    if not CREDENTIALS.exists():
        print("❌ Fehler: credentials.json fehlt im aktuellen Ordner.")
        print("   Bitte aus der Google Cloud Console herunterladen und hier ablegen.")
        sys.exit(1)

    creds = None
    if TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    print(f"✅ Token gespeichert unter: {TOKEN}")

if __name__ == "__main__":
    main()
