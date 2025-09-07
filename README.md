# 🔐 Nuki Gästecode Manager

Automatisches Setzen von Gäste-PINs für Nuki-Smartlocks basierend auf Buchungsdaten aus Google Drive (Excel-Dateien).

## 🕒 Funktionsweise

Automatisiert die **Zeitfenster** eines festen Gäste-Codes („**Gäste**“) auf deinen Nuki-Keypads anhand einer **Belegungs-Tabelle (XLSX)** auf **Google Drive**:

- Am **Anreisetag** wird der Code ab `CHECKIN_TIME` aktiviert  
- Der Code bleibt bis `CHECKOUT_TIME` am **Abreisetag** gültig  
- Wenn **kein Aufenthalt** anliegt, wird der Code deaktiviert  
- Läuft dauerhaft im Container und prüft die aktuellen Buchungslisten **täglich zur Uhrzeit aus `RUN_TIME` (Standard: 05:00)**  
- Ein Lauf **nach der `CHECKOUT_TIME`** ist sinnvoll, weil bei sogenannten *Turnover-Tagen* (Abreise + Anreise am gleichen Tag) sichergestellt wird, dass der alte Code zuverlässig deaktiviert ist, bevor der neue Gast aktiviert wird (z. B. `RUN_TIME=10:30`).  
- Schickt **täglich eine E-Mail** mit dem Ergebnis (OK/Fehler)  
- Unterstützt **beliebig viele Apartments** per `.env` – keine Codeänderungen nötig  
- Alle Aktionen werden zusätzlich in `/app/log` protokolliert (automatische Bereinigung nach `LOG_RETENTION_DAYS`)  

So ist jederzeit sichergestellt, dass der Gäste-PIN **immer korrekt gesetzt** ist und deine Gäste zuverlässig Zugang haben.

## 🚀 Features
- Liest Belegungslisten (Excel, Google Drive).
- Setzt automatisch Gäste-Codes in Nuki Smartlocks.
- Zeitfenster basiert auf `CHECKIN_TIME` und `CHECKOUT_TIME`.
- Mehrere Apartments über `.env` konfigurierbar.
- Tägliche Reports per Mail.
- Automatisches Log-Rotation.

---

## ⚙️ Voraussetzungen

- Python ≥ 3.10 oder Docker
- Google Drive API Zugriff (`token.json` muss einmalig generiert werden)
- Nuki Web API Access Token
- Zugriffsdaten für Mailserver (SMTP), falls tägliche Reports gewünscht sind.

---

## 🔑 Benötigte IDs und Tokens

### 1. Google Drive – `DRIVE_FILE_ID`
- Öffne die Belegungsliste in Google Drive.
- Rechtsklick auf die Datei → **Link abrufen** → **Link in die Zwischenablage kopieren**.
- Im Link findest du die **Drive File ID**.  
  Beispiel:  
  `https://drive.google.com/open?id=1zU8useBfWatmG8rTzYnTde9zRg1hxRzT&usp=drive_fs`  
  → Die ID ist `1zU8useBfWatmG8rTzYnTde9zRg1hxRzT`

### 2. Nuki Smartlock ID
- Melde dich bei [Nuki Web](https://web.nuki.io) an.
- Wähle dein Smart Lock.
- In der URL steht die ID, z. B. `https://web.nuki.io/de/#/smartlock/12345678901`  
  → `12345678901` ist die **SMARTLOCK_ID**.

### 3. Nuki Access Token
- Gehe zu [Nuki Web API](https://web.nuki.io/de/#/admin/web-api).
- Erstelle ein neues API-Token mit vollen Berechtigungen.
- Kopiere den Token in `.env` → `NUKI_ACCESS_TOKEN`.

### 4. Google Drive Token (`/secrets/token.json`)
Damit das Skript auf deine Belegungstabellen zugreifen darf, musst du einmalig ein OAuth-Token erzeugen:

1. Erstelle in der [Google Cloud Console](https://console.cloud.google.com/) ein **OAuth 2.0 Client ID** (Desktop-App).
2. Lade die Datei **credentials.json** herunter.
3. Lege sie in dein Projektverzeichnis neben `authorize_google.py`.
4. Führe einmalig aus:
   ```bash
   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
   python authorize_google.py
   ```
5. Beim ersten Start öffnet sich ein Browserfenster → melde dich mit dem Google-Account an, der Zugriff auf die Belegungstabellen hat.
6. Nach erfolgreicher Freigabe wird automatisch eine Datei **token.json** erzeugt.
7. **Nur diese `token.json` kommt in den Container** → nach `/secrets/token.json`.  
   Die `credentials.json` bleibt lokal bei dir und wird im Container nicht benötigt.

> ⚠️ Ohne diesen Schritt kann das Skript nicht auf Google Drive zugreifen.

---

## 📦 Installation

### Variante A: Docker (empfohlen)

```bash
docker run -d   --name nuki-scheduler   -v /pfad/zur/.env:/app/.env   -v /pfad/zu/secrets:/secrets   -v /pfad/zu/log:/app/log   python:3.11-slim   sh -c "pip install --no-cache-dir -r /app/requirements.txt && python -u /app/main.py"
```

### Variante B: Lokal (Python direkt)

```bash
git clone https://github.com/deinuser/nuki-scheduler.git
cd nuki-scheduler
pip install -r requirements.txt
python main.py
```

---

## 📝 .env Konfiguration

Siehe [`.env.example`](.env.example).

Wichtige Variablen:
- `APTS` → Liste der Apartments (IDs frei wählbar z.B. APTS=1,2).
- `APT_<ID>_DRIVE_FILE_ID` → Google Drive File-ID (z.B. APT_1_DRIVE_FILE_ID=)
- `APT_<ID>_SMARTLOCK_ID` → Nuki Smartlock-ID.
- `APT_<ID>_PIN` → Gäste-PIN.
- `COL_ARRIVAL` / `COL_DEPARTURE` → Spaltenüberschriften aus der XLSX (z. B. `Anreise` / `Abreise` oder `Aankomstdatum` / `Vertrekdatum`).
- `AUTH_NAME` → Der Name des Gäste-Codes im Nuki (Standard: `Gäste`).
- `RUN_TIME` → Zeitpunkt des täglichen Laufs (Standard: `05:00`, z.B. `10:30` für Turnover-Optimierung).

---

## 📧 E-Mail Report

Das Skript sendet nach jedem Lauf einen Report mit:
- Übersicht der Apartments
- Status der gesetzten Gäste-Codes
- Fehler, falls vorhanden

Beispiel Betreff:
```
OK - Nuki Scheduler Report - 06.09.2025
```

---

## 🔄 Logs

- Logs werden unter `/app/log` abgelegt.
- Pro Tag eine Datei: `nuki-YYYY-MM-DD.log`.
- Alte Logs werden nach `LOG_RETENTION_DAYS` gelöscht.

---

## 👨‍💻 Entwicklung

### Installation von Abhängigkeiten (lokal)
```bash
pip install -r requirements.txt
```

### Manuell ausführen
```bash
python main.py --once
```

---

## 📜 Lizenz

MIT License – frei verwendbar.
