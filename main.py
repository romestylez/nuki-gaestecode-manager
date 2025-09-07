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

# Zeiten (Format HH:MM)
ci = os.getenv("CHECKIN_TIME", "15:00")
co = os.getenv("CHECKOUT_TIME", "11:00")
CHECKIN_HOUR, CHECKIN_MIN = map(int, ci.split(":"))
CHECKOUT_HOUR, CHECKOUT_MIN = map(int, co.split(":"))

# Tägliche Laufzeit (Standard 05:00)
rt = os.getenv("RUN_TIME", "05:00")
RUN_HOUR, RUN_MIN = map(int, rt.split(":"))

# Spaltennamen in der Excel-Tabelle (konfigurierbar)
COL_ARRIVAL = os.getenv("COL_ARRIVAL", "Aankomstdatum").strip()
COL_DEPARTURE = os.getenv("COL_DEPARTURE", "Vertrekdatum").strip()

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
    Struktur:
      - APTS=ID1,ID2,ID3
      - APT_<ID>_DRIVE_FILE_ID=...
      - APT_<ID>_SMARTLOCK_ID=...
      - [optional] APT_<ID>_PIN=...              # nur nötig, wenn Code neu angelegt werden soll
      - [optional] APT_<ID>_NAME=...             # Anzeigename (Mail/Log)
      - [optional] APT_<ID>_AUTH_NAME=...        # Name des Gästecodes (Default: 'Gäste')
    """
    apts = {}
    apt_ids = [x.strip() for x in os.getenv("APTS", "").split(",") if x.strip()]
    if not apt_ids:
        pattern = re.compile(r"^APT_(?P<id>[A-Za-z0-9\-]+)_SMARTLOCK_ID$")
        for k in os.environ.keys():
            m = pattern.match(k)
            if m:
                apt_ids.append(m.group("id"))
        apt_ids = sorted(set(apt_ids))

    if not apt_ids:
        raise RuntimeError("Keine Apartment-IDs gefunden. Setze APTS oder APT_<ID>_* Variablen.")

    for aid in apt_ids:
        prefix = f"APT_{aid}_"
        name = os.getenv(prefix + "NAME", f"Apartment {aid}")
        auth_name = os.getenv(prefix + "AUTH_NAME", "Gäste").strip()
        file_id = os.getenv(prefix + "DRIVE_FILE_ID")
        smartlock_id = os.getenv(prefix + "SMARTLOCK_ID")
        pin_raw = os.getenv(prefix + "PIN")  # optional

        missing = [k for k in ["DRIVE_FILE_ID", "SMARTLOCK_ID"] if not os.getenv(prefix + k)]
        if missing:
            logging.error(f"[ERR] Konfiguration unvollständig für {aid}: fehlt {', '.join(missing)} – wird übersprungen")
            continue

        try:
            pin_val = int(pin_raw) if pin_raw not in (None, "", "None") else None
            apts[str(aid)] = {
                "name": name,
                "auth_name": auth_name,
                "file_id": file_id,
                "smartlock_id": int(smartlock_id),
                "pin": pin_val,  # kann None sein
            }
        except ValueError:
            logging.error(f"[ERR] Ungültige Zahl in APT_{aid}_SMARTLOCK_ID oder APT_{aid}_PIN – wird übersprungen")
            continue

    if not apts:
        raise RuntimeError("Keine gültigen Apartments aus .env geladen.")
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
                            "%d.%m.%Y %H:%M:%S", tz=TZ)

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

setup_logging()
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
def load_bookings(file_id: str) -> pd.DataFrame:
    raw = download_xlsx_from_drive(file_id)
    df0 = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=None)

    header_row = None
    for i in range(min(30, len(df0))):
        val = str(df0.iat[i, 0]).strip().lower()
        if val.startswith("naam") or val.startswith("name") or "gasten" in val or "gäste" in val:
            header_row = i
            break
    if header_row is None:
        header_row = 0

    df = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=header_row)

    # Spaltennamen vereinheitlichen anhand .env
    rename_map = {}
    for c in df.columns:
        if str(c).strip() == COL_ARRIVAL:
            rename_map[c] = "arrival"
        if str(c).strip() == COL_DEPARTURE:
            rename_map[c] = "departure"
    df = df.rename(columns=rename_map)

    if "arrival" not in df.columns or "departure" not in df.columns:
        raise ValueError(f"Spalten nicht gefunden. Erwartet: '{COL_ARRIVAL}' und '{COL_DEPARTURE}'.")

    df = df.dropna(how="all")
    df["arrival"] = pd.to_datetime(df["arrival"], errors="coerce", dayfirst=True).dt.date
    df["departure"] = pd.to_datetime(df["departure"], errors="coerce", dayfirst=True).dt.date
    df = df.dropna(subset=["arrival", "departure"])
    df = df[df["departure"] > df["arrival"]]
    return df

def is_turnover_today(df: pd.DataFrame, today: dt.date) -> bool:
    has_departure = any((row["departure"] == today) for _, row in df.iterrows())
    has_arrival = any((row["arrival"] == today) for _, row in df.iterrows())
    return has_departure and has_arrival

def next_stay_interval_today(df: pd.DataFrame, today: dt.date):
    """
    Liefert ein Zeitfenster, wenn:
      - heute innerhalb eines Aufenthalts liegt (a <= today < d), ODER
      - heute der Anreisetag ist (a == today).
    Sonst: None.
    """
    for _, row in df.iterrows():
        a, d = row["arrival"], row["departure"]
        if a <= today < d or a == today:
            start_local = dt.datetime(a.year, a.month, a.day, CHECKIN_HOUR, CHECKIN_MIN, tzinfo=TZ)
            end_local   = dt.datetime(d.year, d.month, d.day, CHECKOUT_HOUR, CHECKOUT_MIN, tzinfo=TZ)
            if end_local <= start_local:
                end_local += dt.timedelta(days=1)
            start_utc = start_local.astimezone(dt.timezone.utc).replace(tzinfo=None)
            end_utc   = end_local.astimezone(dt.timezone.utc).replace(tzinfo=None)
            return start_utc, end_utc
    return None

def next_future_interval(df: pd.DataFrame, today: dt.date):
    """Nächstes zukünftiges Intervall (a > today), sonst None."""
    fut = df[df["arrival"] > today].sort_values("arrival").head(1)
    if fut.empty:
        return None
    a = fut.iloc[0]["arrival"]; d = fut.iloc[0]["departure"]
    start_local = dt.datetime(a.year, a.month, a.day, CHECKIN_HOUR, CHECKIN_MIN, tzinfo=TZ)
    end_local   = dt.datetime(d.year, d.month, d.day, CHECKOUT_HOUR, CHECKOUT_MIN, tzinfo=TZ)
    if end_local <= start_local:
        end_local += dt.timedelta(days=1)
    return (start_local.astimezone(dt.timezone.utc).replace(tzinfo=None),
            end_local.astimezone(dt.timezone.utc).replace(tzinfo=None))

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

def find_auth_by_name(smartlock_id, auth_name: str):
    target = auth_name.strip().casefold()
    for a in get_auth_list_for(smartlock_id):
        try:
            if a.get("type") != 13:
                continue
            name = str(a.get("name", "")).strip().casefold()
            if name == target:
                return a
        except Exception:
            continue
    return None

def ensure_auth(smartlock_id, auth_name: str, pin: int | None):
    """
    Sucht den Auth-Eintrag nach Name.
    - Falls vorhanden → zurückgeben.
    - Falls nicht vorhanden:
        - Wenn PIN vorhanden → anlegen.
        - Sonst → Fehler.
    """
    a = find_auth_by_name(smartlock_id, auth_name)
    if a:
        return a

    if pin is None:
        raise RuntimeError(f"Auth '{auth_name}' nicht gefunden und kein PIN hinterlegt, um ihn anzulegen.")

    payload = {
        "name": auth_name,
        "type": 13,
        "code": pin,
        "smartlockIds": [smartlock_id],
        "allowedWeekDays": 127
    }
    url = f"{BASE}/smartlock/auth"
    r = requests.put(url, headers=HEADERS, data=json.dumps(payload), timeout=20)
    if r.status_code == 409:
        a = find_auth_by_name(smartlock_id, auth_name)
        if a:
            return a
        r.raise_for_status()
    if r.status_code in (200, 201):
        try:
            return r.json()
        except ValueError:
            pass
    return find_auth_by_name(smartlock_id, auth_name)

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
        apt_name = cfg.get("name", f"App {ap_id}")
        auth_name = cfg.get("auth_name", "Gäste")
        try:
            df = load_bookings(cfg["file_id"])

            # Turnover-Hinweis (nur für den Report/Log)
            if is_turnover_today(df, today):
                info = f"[INFO] {apt_name}: Turnover Day erkannt (Abreise & Anreise heute)"
                logging.info(info)
                summary_lines.append(info)

            iv_today = next_stay_interval_today(df, today)

            auth = ensure_auth(cfg["smartlock_id"], auth_name, cfg.get("pin"))
            auth_id = auth.get("id") or auth.get("authId") or auth.get("authID")
            if not auth_id:
                raise RuntimeError(f"Keine authId für '{auth_name}' ({apt_name} / ID {ap_id}): {auth}")

            if iv_today:
                desired_from_utc, desired_until_utc = iv_today
                current_auth = find_auth_by_name(cfg["smartlock_id"], auth_name) or {}
                cur_from = parse_nuki_iso_utc_naive(current_auth.get("allowedFromDate"))
                cur_until = parse_nuki_iso_utc_naive(current_auth.get("allowedUntilDate"))

                start_local = desired_from_utc.replace(tzinfo=dt.timezone.utc).astimezone(TZ)
                end_local   = desired_until_utc.replace(tzinfo=dt.timezone.utc).astimezone(TZ)
                start_str = start_local.strftime("%d.%m.%Y %H:%M")
                end_str   = end_local.strftime("%d.%m.%Y %H:%M")

                if times_equal(cur_from, desired_from_utc) and times_equal(cur_until, desired_until_utc):
                    msg = f"[OK] {apt_name}: Code '{auth_name}' bereits korrekt: {start_str} bis {end_str}"
                    logging.info(msg)
                    summary_lines.append(msg)
                else:
                    update_auth_timewindow(cfg["smartlock_id"], auth_id, desired_from_utc, desired_until_utc)
                    msg = f"[OK] {apt_name}: Code '{auth_name}' gültig von {start_str} bis {end_str} gesetzt"
                    logging.info(msg)
                    summary_lines.append(msg)
            else:
                # Heute kein Aufenthalt → Code deaktivieren ...
                current_auth = find_auth_by_name(cfg["smartlock_id"], auth_name) or {}
                cur_from = parse_nuki_iso_utc_naive(current_auth.get("allowedFromDate"))
                cur_until = parse_nuki_iso_utc_naive(current_auth.get("allowedUntilDate"))

                if cur_from is not None or cur_until is not None:
                    clear_auth_timewindow(cfg["smartlock_id"], auth_id)
                    msg = f"[OK] {apt_name}: Kein Aufenthalt heute – Code '{auth_name}' deaktiviert"
                    logging.info(msg)
                    summary_lines.append(msg)
                else:
                    msg = f"[OK] {apt_name}: Kein Aufenthalt heute – Code '{auth_name}' war bereits deaktiviert"
                    logging.info(msg)
                    summary_lines.append(msg)

                # ... und nächsten Aufenthalt sicher vorplanen, falls vorhanden
                nf = next_future_interval(df, today)
                if nf:
                    desired_from_utc, desired_until_utc = nf
                    update_auth_timewindow(cfg["smartlock_id"], auth_id, desired_from_utc, desired_until_utc)

                    start_local = desired_from_utc.replace(tzinfo=dt.timezone.utc).astimezone(TZ)
                    end_local   = desired_until_utc.replace(tzinfo=dt.timezone.utc).astimezone(TZ)
                    start_str = start_local.strftime("%d.%m.%Y %H:%M")
                    end_str   = end_local.strftime("%d.%m.%Y %H:%M")

                    msg2 = f"[OK] {apt_name}: Nächster Aufenthalt vorgeplant – {start_str} bis {end_str}"
                    logging.info(msg2)
                    summary_lines.append(msg2)

        except Exception as e:
            had_error = True
            msg = f"[ERR] {apt_name}: {e}"
            logging.error(msg)
            summary_lines.append(msg)

    # Genau eine Leerzeile zwischen den Einträgen im Mail-Body
    return had_error, "\n\n".join(summary_lines)

def sleep_until_next_run():
    now = dt.datetime.now(TZ)
    tomorrow = (now + dt.timedelta(days=1)).date()
    wake = dt.datetime(tomorrow.year, tomorrow.month, tomorrow.day, RUN_HOUR, RUN_MIN, tzinfo=TZ)
    sleep_s = (wake - dt.datetime.now(TZ)).total_seconds()
    if sleep_s < 60:
        sleep_s = 3600
    time.sleep(sleep_s)

def main(loop_mode: bool):
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
            sleep_until_next_run()
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
