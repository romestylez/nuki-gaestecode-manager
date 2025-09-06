import os, io, json, time, datetime as dt, argparse, logging, smtplib, re
from zoneinfo import ZoneInfo
from email.message import EmailMessage

import requests
import pandas as pd
from dotenv import load_dotenv

# --- Google Drive (OAuth) ---
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request as GRequest

# ========================= .env laden =========================
load_dotenv("/app/.env")

TZ = ZoneInfo(os.getenv("TZ", "Europe/Amsterdam"))

# Minutengenaue Zeiten aus .env (Format HH:MM)
ci = os.getenv("CHECKIN_TIME", "15:00")
co = os.getenv("CHECKOUT_TIME", "11:00")
CHECKIN_HOUR, CHECKIN_MIN = map(int, ci.split(":"))
CHECKOUT_HOUR, CHECKOUT_MIN = map(int, co.split(":"))

# Spaltennamen (konfigurierbar)
COL_ARRIVAL = os.getenv("COL_ARRIVAL", "Aankomstdatum")
COL_DEPARTURE = os.getenv("COL_DEPARTURE", "Vertrekdatum")

# Log-Retention
LOG_RETENTION_DAYS = int(os.getenv("LOG_RETENTION_DAYS", "30"))

# Mail-Einstellungen (Report wird IMMER versendet)
SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_STARTTLS = os.getenv("SMTP_STARTTLS", "true").strip().lower() in ("1","true","yes","on")
MAIL_FROM = os.getenv("MAIL_FROM", SMTP_USER or "noreply@example.com")
MAIL_TO = os.getenv("MAIL_TO", "")
MAIL_SUBJECT_PREFIX = os.getenv("MAIL_SUBJECT_PREFIX", "Nuki Scheduler Report")

NUKI_TOKEN = os.environ["NUKI_ACCESS_TOKEN"]
HEADERS = {"Authorization": f"Bearer {NUKI_TOKEN}", "Content-Type": "application/json"}

# ========================= Apartments dynamisch laden =========================
def load_apartments_from_env():
    """
    Liest Apartments dynamisch aus Umgebungsvariablen.

    Erwartete Struktur:
      - APTS=ID1,ID2,ID3
      - APT_<ID>_DRIVE_FILE_ID=...  # Google Drive File-ID (XLSX)
      - APT_<ID>_SMARTLOCK_ID=...   # Nuki Smartlock-ID (Zahl)
      - APT_<ID>_PIN=...            # Gäste-PIN (Zahl)
      - [optional] APT_<ID>_NAME=...# Anzeigename fürs Logging/Mail
    """
    apts = {}
    apt_ids = [x.strip() for x in os.getenv("APTS", "").split(",") if x.strip()]
    if not apt_ids:
        # Fallback: automatisch alle APT_<ID>_SMARTLOCK_ID Variablen finden
        pattern = re.compile(r"^APT_(?P<id>[A-Za-z0-9\-]+)_SMARTLOCK_ID$")
        for k in os.environ.keys():
            m = pattern.match(k)
            if m:
                apt_ids.append(m.group("id"))
        apt_ids = sorted(set(apt_ids))

    for aid in apt_ids:
        prefix = f"APT_{aid}_"
        name = os.getenv(prefix + "NAME", f"Apartment {aid}")
        file_id = os.getenv(prefix + "DRIVE_FILE_ID")
        smartlock_id = os.getenv(prefix + "SMARTLOCK_ID")
        pin = os.getenv(prefix + "PIN")

        missing = [k for k in ["DRIVE_FILE_ID", "SMARTLOCK_ID", "PIN"] if not os.getenv(prefix + k)]
        if missing:
            logging.error(f"[ERR] Konfiguration unvollständig für {aid}: fehlt {', '.join(missing)} – wird übersprungen")
            continue

        try:
            apts[str(aid)] = {
                "name": name,
                "file_id": file_id,
                "smartlock_id": int(smartlock_id),
                "pin": int(pin),
            }
        except ValueError:
            logging.error(f"[ERR] Ungültige Zahl in APT_{aid}_SMARTLOCK_ID oder APT_{aid}_PIN – wird übersprungen")
            continue

    if not apts:
        raise RuntimeError("Keine gültigen Apartments aus .env geladen. Bitte APTS und APT_<ID>_* Variablen setzen.")
    return apts

APT = load_apartments_from_env()

# ========================= Logging (deutsche Zeit & TZ) =========================
class TZFormatter(logging.Formatter):
    def __init__(self, fmt=None, datefmt=None, tz=None):
        super().__init__(fmt=fmt, datefmt=datefmt)
        self.tz = tz

    def formatTime(self, record, datefmt=None):
        dt_obj = dt.datetime.fromtimestamp(record.created, tz=self.tz)
        if datefmt:
            return dt_obj.strftime(datefmt)
        return dt_obj.strftime("%d.%m.%Y %H:%M:%S")

def setup_logging():
    logdir = "/app/log"
    os.makedirs(logdir, exist_ok=True)
    today_str = dt.datetime.now(TZ).strftime("%Y-%m-%d")
    logfile = f"{logdir}/nuki-{today_str}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    for h in list(logger.handlers):
        logger.removeHandler(h)

    formatter = TZFormatter("%(asctime)s %(levelname)s: %(message)s",
                            "%d.%m.%Y %H:%M:%S",
                            tz=TZ)

    fh = logging.FileHandler(logfile, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.ERROR)
    logging.getLogger("googleapiclient.discovery").setLevel(logging.ERROR)

    return logfile

def cleanup_logs(logdir="/app/log", retention_days=30):
    now = dt.datetime.now(TZ)
    try:
        for name in os.listdir(logdir):
            path = os.path.join(logdir, name)
            if not os.path.isfile(path):
                continue
            mtime = dt.datetime.fromtimestamp(os.path.getmtime(path), TZ)
            if (now - mtime).days > retention_days:
                os.remove(path)
    except Exception:
        pass

def send_report_mail(subject, body):
    if not SMTP_HOST or not MAIL_TO:
        return
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO
    msg.set_content(body)

    if SMTP_PORT == 465:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            if SMTP_USER:
                s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            if SMTP_STARTTLS:
                s.starttls()
            if SMTP_USER:
                s.login(SMTP_USER, SMTP_PASSWORD)
            s.send_message(msg)

logfile_path = setup_logging()
log = logging.getLogger(__name__)

# ========================= Drive Helper =========================
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
TOKEN_PATH = "/secrets/token.json"

def get_drive_service():
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, DRIVE_SCOPES)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(GRequest())
        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def download_xlsx_from_drive(file_id: str) -> bytes:
    drive = get_drive_service()
    request = drive.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    return fh.getvalue()

# ========================= Buchungen =========================
def _norm(s) -> str:
    return str(s).strip().casefold()

def load_bookings(file_id: str) -> pd.DataFrame:
    # XLSX komplett ohne Header einlesen, um Headerzeile zu detektieren
    raw = download_xlsx_from_drive(file_id)
    df0 = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=None)

    # Gesuchte Spaltennamen normalisieren
    want_arr = _norm(COL_ARRIVAL)
    want_dep = _norm(COL_DEPARTURE)

    # Headerzeile finden: eine Zeile, die beide gewünschten Spaltennamen enthält
    header_row = None
    max_scan = min(50, len(df0))
    for i in range(max_scan):
        row_vals = [_norm(v) for v in list(df0.iloc[i].values)]
        if want_arr in row_vals and want_dep in row_vals:
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            f"Kopfzeile nicht gefunden. Erwartete Spaltennamen: "
            f"'{COL_ARRIVAL}' und '{COL_DEPARTURE}'."
        )

    # Datensatz ab erkannter Headerzeile mit echten Spaltennamen erneut einlesen
    df = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=header_row)

    # Spalten robust auf 'arrival'/'departure' mappen (case/whitespace-insensitiv)
    colmap = {}
    for c in df.columns:
        cname = _norm(c)
        if cname == want_arr:
            colmap[c] = "arrival"
        elif cname == want_dep:
            colmap[c] = "departure"

    if "arrival" not in colmap.values() or "departure" not in colmap.values():
        raise ValueError(
            f"Spalten konnten nicht zugeordnet werden. Gefundene Spalten: {list(df.columns)}. "
            f"Erwartet: '{COL_ARRIVAL}' und '{COL_DEPARTURE}'."
        )

    df = df.rename(columns=colmap)

    # Aufräumen & Datumsparsing
    df = df.dropna(how="all")
    df["arrival"] = pd.to_datetime(df["arrival"], errors="coerce", dayfirst=True).dt.date
    df["departure"] = pd.to_datetime(df["departure"], errors="coerce", dayfirst=True).dt.date
    df = df.dropna(subset=["arrival", "departure"])
    df = df[df["departure"] > df["arrival"]]
    return df

def next_stay_interval(df: pd.DataFrame, today: dt.date):
    candidates = []
    for _, row in df.iterrows():
        a, d = row["arrival"], row["departure"]
        if a <= today < d:
            candidates.append((a, d, 0))
        elif a >= today:
            candidates.append((a, d, (a - today).days))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[2], x[0]))
    a, d, _ = candidates[0]

    start_local = dt.datetime(a.year, a.month, a.day, CHECKIN_HOUR, CHECKIN_MIN, tzinfo=TZ)
    end_local   = dt.datetime(d.year, d.month, d.day, CHECKOUT_HOUR, CHECKOUT_MIN, tzinfo=TZ)
    if end_local <= start_local:
        end_local += dt.timedelta(days=1)

    start_utc = start_local.astimezone(dt.timezone.utc).replace(tzinfo=None)
    end_utc   = end_local.astimezone(dt.timezone.utc).replace(tzinfo=None)
    return start_utc, end_utc

# ========================= Nuki Web API =========================
BASE = "https://api.nuki.io"

def get_auth_list_for(smartlock_id):
    url = f"{BASE}/smartlock/{smartlock_id}/auth"
    r = requests.get(url, headers=HEADERS, timeout=20)
    if r.status_code == 204 or not (r.content or b"").strip():
        return []
    r.raise_for_status()
    try:
        return r.json()
    except ValueError:
        return []

def find_gaeste_auth(smartlock_id):
    for a in get_auth_list_for(smartlock_id):
        try:
            if a.get("type") != 13:
                continue
            name = str(a.get("name", "")).strip()
            if name.casefold() == "gäste".casefold():
                return a
        except Exception:
            continue
    return None

def ensure_gaeste_auth(smartlock_id, pin):
    a = find_gaeste_auth(smartlock_id)
    if a:
        return a
    payload = {
        "name": "Gäste",
        "type": 13,
        "code": pin,
        "smartlockIds": [smartlock_id],
        "allowedWeekDays": 127
    }
    url = f"{BASE}/smartlock/auth"
    r = requests.put(url, headers=HEADERS, data=json.dumps(payload), timeout=20)
    if r.status_code == 409:
        a = find_gaeste_auth(smartlock_id)
        if a:
            return a
        r.raise_for_status()
    if r.status_code in (200, 201):
        try:
            return r.json()
        except ValueError:
            pass
    return find_gaeste_auth(smartlock_id)

def update_auth_timewindow(smartlock_id, auth_id, start_utc, end_utc):
    payload = {
        "allowedFromDate": start_utc.isoformat(timespec="milliseconds") + "Z",
        "allowedUntilDate": end_utc.isoformat(timespec="milliseconds") + "Z",
        "allowedWeekDays": 127
    }
    url = f"{BASE}/smartlock/{smartlock_id}/auth/{auth_id}"
    r = requests.post(url, headers=HEADERS, data=json.dumps(payload), timeout=20)
    r.raise_for_status()

def clear_auth_timewindow(smartlock_id, auth_id):
    payload = {"allowedFromDate": None, "allowedUntilDate": None}
    url = f"{BASE}/smartlock/{smartlock_id}/auth/{auth_id}"
    r = requests.post(url, headers=HEADERS, data=json.dumps(payload), timeout=20)
    r.raise_for_status()

def parse_nuki_iso_utc_naive(s: str | None) -> dt.datetime | None:
    if not s:
        return None
    s = s.strip()
    if not s:
        return None
    if s.endswith("Z"):
        s = s[:-1]
    try:
        return dt.datetime.fromisoformat(s)
    except ValueError:
        return None

def times_equal(a: dt.datetime | None, b: dt.datetime | None, tol_sec: int = 60) -> bool:
    if a is None and b is None:
        return True
    if (a is None) != (b is None):
        return False
    return abs((a - b).total_seconds()) <= tol_sec

# ========================= Hauptlauf =========================
def run_once() -> tuple[bool, str]:
    had_error = False
    summary_lines = []
    today = dt.datetime.now(TZ).date()

    for ap_id, cfg in APT.items():
        try:
            df = load_bookings(cfg["file_id"])
            iv = next_stay_interval(df, today)

            auth = ensure_gaeste_auth(cfg["smartlock_id"], cfg["pin"])
            auth_id = auth.get("id") or auth.get("authId") or auth.get("authID")
            if not auth_id:
                raise RuntimeError(f"Keine authId für 'Gäste' ({cfg.get('name','App '+ap_id)} / ID {ap_id}): {auth}")

            if iv:
                desired_from_utc, desired_until_utc = iv
                current_auth = find_gaeste_auth(cfg["smartlock_id"]) or {}
                cur_from = parse_nuki_iso_utc_naive(current_auth.get("allowedFromDate"))
                cur_until = parse_nuki_iso_utc_naive(current_auth.get("allowedUntilDate"))

                start_local = desired_from_utc.replace(tzinfo=dt.timezone.utc).astimezone(TZ)
                end_local   = desired_until_utc.replace(tzinfo=dt.timezone.utc).astimezone(TZ)
                start_str = start_local.strftime("%d.%m.%Y %H:%M")
                end_str   = end_local.strftime("%d.%m.%Y %H:%M")

                apt_name = cfg.get("name", f"App {ap_id}")
                if times_equal(cur_from, desired_from_utc) and times_equal(cur_until, desired_until_utc):
                    msg = f"[OK] {apt_name}: Code 'Gäste' bereits korrekt: {start_str} bis {end_str}"
                    logging.info(msg)
                    summary_lines.append(msg)
                else:
                    update_auth_timewindow(cfg["smartlock_id"], auth_id, desired_from_utc, desired_until_utc)
                    msg = f"[OK] {apt_name}: Code 'Gäste' gültig von {start_str} bis {end_str} gesetzt"
                    logging.info(msg)
                    summary_lines.append(msg)
            else:
                current_auth = find_gaeste_auth(cfg["smartlock_id"]) or {}
                cur_from = parse_nuki_iso_utc_naive(current_auth.get("allowedFromDate"))
                cur_until = parse_nuki_iso_utc_naive(current_auth.get("allowedUntilDate"))
                apt_name = cfg.get("name", f"App {ap_id}")
                if cur_from is None and cur_until is None:
                    msg = f"[OK] {apt_name}: Kein Aufenthalt – Code 'Gäste' war bereits deaktiviert"
                    logging.info(msg)
                    summary_lines.append(msg)
                else:
                    clear_auth_timewindow(cfg["smartlock_id"], auth_id)
                    msg = f"[OK] {apt_name}: Kein Aufenthalt – Code 'Gäste' deaktiviert"
                    logging.info(msg)
                    summary_lines.append(msg)
        except Exception as e:
            had_error = True
            apt_name = cfg.get("name", f"App {ap_id}")
            msg = f"[ERR] {apt_name}: {e}"
            logging.error(msg)
            summary_lines.append(msg)

    # Genau eine Leerzeile zwischen Zeilen im Mail-Body
    return had_error, "\n\n".join(summary_lines)

def main(loop_mode: bool):
    global logfile_path
    logfile_path = setup_logging()
    cleanup_logs("/app/log", LOG_RETENTION_DAYS)

    if loop_mode:
        while True:
            had_error, summary = run_once()
            try:
                date_str = dt.datetime.now(TZ).strftime("%d.%m.%Y")
                suffix = "FEHLER" if had_error else "OK"
                subject = f"{suffix} - {MAIL_SUBJECT_PREFIX} - {date_str}"
                send_report_mail(subject, summary)
            except Exception as e:
                logging.error(f"[ERR] Mailversand fehlgeschlagen: {e}")

            now = dt.datetime.now(TZ)
            tomorrow = (now + dt.timedelta(days=1)).date()
            wake = dt.datetime(tomorrow.year, tomorrow.month, tomorrow.day, 5, 0, tzinfo=TZ)
            sleep_s = (wake - dt.datetime.now(TZ)).total_seconds()
            if sleep_s < 60:
                sleep_s = 3600
            time.sleep(sleep_s)
    else:
        had_error, summary = run_once()
        try:
            date_str = dt.datetime.now(TZ).strftime("%d.%m.%Y")
            suffix = "FEHLER" if had_error else "OK"
            subject = f"{suffix} - {MAIL_SUBJECT_PREFIX} - {date_str}"
            send_report_mail(subject, summary)
        except Exception as e:
            logging.error(f"[ERR] Mailversand fehlgeschlagen: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Nuki Scheduler")
    parser.add_argument("--once", action="store_true", help="einmal ausführen und beenden (für Cron/Task Scheduler)")
    args = parser.parse_args()

    loop_mode = not args.once
    main(loop_mode)
