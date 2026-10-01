from __future__ import annotations

import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.environ.get("WRS_DATA_DIR", os.path.join(BASE_DIR, "data"))
os.makedirs(DATA_DIR, exist_ok=True)

SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
ADDRESS_FILE = os.path.join(DATA_DIR, "address_book.json")
TRACKER_FILE = os.path.join(DATA_DIR, "streak_tracker.json")
CORRECTIONS_FILE = os.path.join(DATA_DIR, "corrections.json")
RECON_FILE = os.path.join(DATA_DIR, "reconstruction.json")
EVENTS_FILE = os.path.join(DATA_DIR, "events_log.json")
HISTORY_FILE = os.path.join(DATA_DIR, "history.json")
DVORNIK_FILE = os.path.join(DATA_DIR, "dvornik_schedule.json")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

APP_TITLE = "WRS Mobile"
APP_VER = "1.1.0"
