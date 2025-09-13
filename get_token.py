import sys
import pathlib
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Google Drive nur-lesend
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

BASE = pathlib.Path(__file__).parent.resolve()
TOKEN = BASE / "token.json"

def find_credentials_file():
    """Suche eine Datei client_secret*.json im aktuellen Ordner."""
    for file in BASE.glob("client_secret*.json"):
        return file
    return None

def main():
    CREDENTIALS = find_credentials_file()
    if not CREDENTIALS:
        print("❌ Fehler: Keine client_secret*.json im aktuellen Ordner gefunden.")
        print("   Bitte die Datei aus der Google Cloud Console hier ablegen.")
        sys.exit(1)

    creds = None
    if TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)

    # Wenn Token fehlt oder ungültig -> löschen und neu anfordern
    if not creds or not creds.valid:
        if TOKEN.exists():
            TOKEN.unlink()  # alte token.json löschen
            print("⚠️ Alte token.json gelöscht, neuer Flow gestartet…")

        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), SCOPES)
            try:
                creds = flow.run_local_server(port=0)
            except Exception:
                print("⚠️ Lokaler Server fehlgeschlagen – nutze run_console()…")
                creds = flow.run_console()

        with open(TOKEN, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    print(f"✅ Token gespeichert unter: {TOKEN}")

if __name__ == "__main__":
    main()
