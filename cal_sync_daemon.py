import os
import time
import sqlite3
import random
import datetime as dt
import re
import json
from dataclasses import dataclass
from typing import Optional, Dict, Any, List, Tuple, Set

import pythoncom
import win32com.client  # pywin32
from dateutil import tz
from dateutil.parser import isoparse
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from datetime import timezone

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_env_file(path: str) -> None:
    if not path or not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as handle:
            for raw in handle:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("export "):
                    line = line[len("export "):].strip()
                if "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                if not key:
                    continue
                if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
                    value = value[1:-1]
                os.environ.setdefault(key, value)
    except Exception as e:
        print(f"[config] failed to read env file '{path}': {e}")

ENV_FILE_PATH = os.environ.get("SYNC_ENV_PATH", os.path.join(SCRIPT_DIR, "sync.env"))
_load_env_file(ENV_FILE_PATH)

# ----------------------------
# Config
# ----------------------------
SCOPES = ["https://www.googleapis.com/auth/calendar"]

def _resolve_timezone() -> tz.tzfile:
    tz_name = os.environ.get("SYNC_TIMEZONE", "").strip()
    if tz_name:
        tzinfo = _timezone_from_name(tz_name)
        if tzinfo:
            return tzinfo
        print(f"[config] invalid SYNC_TIMEZONE '{tz_name}', falling back to local timezone")
    return tz.tzlocal()

def _timezone_from_name(tz_name: str) -> Optional[dt.tzinfo]:
    if not tz_name:
        return None
    try:
        return ZoneInfo(tz_name)
    except ZoneInfoNotFoundError:
        pass
    return tz.gettz(tz_name)

POLL_SECONDS = int(os.environ.get("SYNC_POLL_SECONDS", "60"))
LOOKBACK_DAYS = int(os.environ.get("SYNC_LOOKBACK_DAYS", "365"))
LOOKAHEAD_DAYS = int(os.environ.get("SYNC_LOOKAHEAD_DAYS", "365"))

GOOGLE_CALENDAR_ID = os.environ.get("GOOGLE_CALENDAR_ID", "primary")

DB_PATH = os.environ.get("SYNC_DB_PATH", "sync_state.sqlite3")

LOCAL_TZ = _resolve_timezone()
_SYNC_TZ_EXPLICIT = bool(os.environ.get("SYNC_TIMEZONE", "").strip())

CREDS_PATH = os.path.join(SCRIPT_DIR, "credentials.json")
TOKEN_PATH = os.path.join(SCRIPT_DIR, "token.json")

# Persistent tag names (survive DB deletion)
OUTLOOK_PROP_GCAL_ID = "GCAL_EVENT_ID"
GOOGLE_EXT_OUTLOOK_ID = "outlook_entry_id"
GOOGLE_EXT_SYNC_TAG = "sync_tag"
SYNC_TAG_VALUE = os.environ.get("SYNC_TAG_VALUE", "outlook-google-sync")

# Category/color sync (Outlook Categories <-> Google colorId)
GOOGLE_EXT_OUTLOOK_CATEGORIES = "outlook_categories"   # stored in extendedProperties.private
OUTLOOK_PROP_GCAL_COLOR_ID = "GCAL_COLOR_ID"           # optional Outlook userprop
DEFAULT_GCAL_COLOR_ID = os.environ.get("DEFAULT_GCAL_COLOR_ID", "").strip() or None

# JSON dict: {"Work Appointments":"5","Personal":"2"}
_OUTLOOK_CATEGORY_COLOR_MAP_RAW = os.environ.get("OUTLOOK_CATEGORY_COLOR_MAP", "").strip()
# JSON dict: {"5":"Work Appointments","2":"Personal"}
_GCAL_COLOR_OUTLOOK_CATEGORY_MAP_RAW = os.environ.get("GCAL_COLOR_OUTLOOK_CATEGORY_MAP", "").strip()

def _load_json_map(raw: str) -> Dict[str, str]:
    if not raw:
        return {}
    try:
        obj = json.loads(raw)
        if isinstance(obj, dict):
            # force strings
            return {str(k): str(v) for k, v in obj.items()}
    except Exception as e:
        print(f"[config] bad JSON map env var: {e}")
    return {}

OUTLOOK_CATEGORY_COLOR_MAP: Dict[str, str] = _load_json_map(_OUTLOOK_CATEGORY_COLOR_MAP_RAW)
GCAL_COLOR_OUTLOOK_CATEGORY_MAP: Dict[str, str] = _load_json_map(_GCAL_COLOR_OUTLOOK_CATEGORY_MAP_RAW)

# Dedupe options
DEDUPE_ON_START = os.environ.get("DEDUPE_ON_START", "0").strip() == "1"
DEDUPE_DRY_RUN = os.environ.get("DEDUPE_DRY_RUN", "0").strip() == "1"
DEDUPE_SLEEP_BETWEEN_DELETES_S = float(os.environ.get("DEDUPE_SLEEP_BETWEEN_DELETES_S", "0.2"))
DEDUPE_MAX_RESULTS = int(os.environ.get("DEDUPE_MAX_RESULTS", "2500"))

# Optional: print htmlLink on create/patch
GOOGLE_UPSERT_DEBUG = os.environ.get("GOOGLE_UPSERT_DEBUG", "0").strip() == "1"


# ----------------------------
# Data model
# ----------------------------
@dataclass
class EventRecord:
    side: str  # "outlook" or "google"
    uid: str   # google: eventId, outlook: EntryID
    title: str
    start: dt.datetime
    end: dt.datetime
    all_day: bool
    location: str
    description: str
    last_modified: dt.datetime
    deleted: bool = False

    # Persistent linking hints
    outlook_entry_id_hint: Optional[str] = None  # from google extendedProperties, if present
    gcal_event_id_hint: Optional[str] = None     # from outlook user property, if present

    # Category/color syncing
    categories: List[str] = None                 # Outlook categories (list)
    google_color_id: Optional[str] = None        # Google event.colorId (string)


# ----------------------------
# SQLite helpers (WAL + retry)
# ----------------------------
def _db_connect() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH, timeout=30)
    cur = con.cursor()
    cur.execute("PRAGMA journal_mode=WAL;")
    cur.execute("PRAGMA synchronous=NORMAL;")
    cur.execute("PRAGMA busy_timeout=30000;")
    con.commit()
    return con

def _db_retry(fn, *args, **kwargs):
    backoff = 0.05
    for attempt in range(10):
        try:
            return fn(*args, **kwargs)
        except sqlite3.OperationalError as e:
            msg = str(e).lower()
            if "database is locked" in msg or "database is busy" in msg:
                time.sleep(backoff + random.uniform(0, backoff))
                backoff = min(1.5, backoff * 2)
                continue
            raise
    raise sqlite3.OperationalError("database is locked (retries exhausted)")

def db_init():
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS mapping (
            outlook_id TEXT PRIMARY KEY,
            google_id  TEXT UNIQUE,
            last_sync_outlook_mod TEXT,
            last_sync_google_mod  TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS kv (
            k TEXT PRIMARY KEY,
            v TEXT
        )
        """)
        con.commit()
        con.close()
    _db_retry(_impl)

def kv_get(key: str) -> Optional[str]:
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute("SELECT v FROM kv WHERE k=?", (key,))
        row = cur.fetchone()
        con.close()
        return row[0] if row else None
    return _db_retry(_impl)

def kv_set(key: str, value: str):
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute(
            "INSERT INTO kv(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v",
            (key, value),
        )
        con.commit()
        con.close()
    _db_retry(_impl)

def map_get_by_outlook(outlook_id: str) -> Optional[Tuple[str, str, str]]:
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute(
            "SELECT google_id, last_sync_outlook_mod, last_sync_google_mod FROM mapping WHERE outlook_id=?",
            (outlook_id,),
        )
        row = cur.fetchone()
        con.close()
        return row if row else None
    return _db_retry(_impl)

def map_get_by_google(google_id: str) -> Optional[Tuple[str, str, str]]:
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute(
            "SELECT outlook_id, last_sync_outlook_mod, last_sync_google_mod FROM mapping WHERE google_id=?",
            (google_id,),
        )
        row = cur.fetchone()
        con.close()
        return row if row else None
    return _db_retry(_impl)

def parse_sync_mod(raw: Optional[str]) -> Optional[dt.datetime]:
    if not raw:
        return None
    try:
        return isoparse(raw).astimezone(LOCAL_TZ)
    except Exception:
        return None

def map_upsert(outlook_id: str, google_id: str, outlook_mod: Optional[dt.datetime], google_mod: Optional[dt.datetime]):
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        try:
            cur.execute("""
            INSERT INTO mapping(outlook_id, google_id, last_sync_outlook_mod, last_sync_google_mod)
            VALUES(?,?,?,?)
            ON CONFLICT(outlook_id) DO UPDATE SET
              google_id=excluded.google_id,
              last_sync_outlook_mod=excluded.last_sync_outlook_mod,
              last_sync_google_mod=excluded.last_sync_google_mod
            """, (
                outlook_id,
                google_id,
                outlook_mod.isoformat() if outlook_mod else None,
                google_mod.isoformat() if google_mod else None
            ))
            con.commit()
        except sqlite3.IntegrityError as e:
            msg = str(e).lower()
            if "unique constraint failed: mapping.google_id" in msg:
                cur.execute("DELETE FROM mapping WHERE google_id=? AND outlook_id<>?", (google_id, outlook_id))
                cur.execute("""
                INSERT INTO mapping(outlook_id, google_id, last_sync_outlook_mod, last_sync_google_mod)
                VALUES(?,?,?,?)
                ON CONFLICT(outlook_id) DO UPDATE SET
                  google_id=excluded.google_id,
                  last_sync_outlook_mod=excluded.last_sync_outlook_mod,
                  last_sync_google_mod=excluded.last_sync_google_mod
                """, (
                    outlook_id,
                    google_id,
                    outlook_mod.isoformat() if outlook_mod else None,
                    google_mod.isoformat() if google_mod else None
                ))
                con.commit()
            else:
                raise
        finally:
            con.close()
    _db_retry(_impl)

def map_delete_by_outlook(outlook_id: str):
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute("DELETE FROM mapping WHERE outlook_id=?", (outlook_id,))
        con.commit()
        con.close()
    _db_retry(_impl)

def map_delete_by_google(google_id: str):
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute("DELETE FROM mapping WHERE google_id=?", (google_id,))
        con.commit()
        con.close()
    _db_retry(_impl)

def db_get_all_google_ids_in_mapping() -> Set[str]:
    def _impl():
        con = _db_connect()
        cur = con.cursor()
        cur.execute("SELECT google_id FROM mapping WHERE google_id IS NOT NULL")
        rows = cur.fetchall()
        con.close()
        return set(r[0] for r in rows if r and r[0])
    return _db_retry(_impl)


# ----------------------------
# Normalization / safe time comparisons (NO timestamp())
# ----------------------------
_ws_re = re.compile(r"\s+")

def normalize_text(s: str) -> str:
    s = (s or "").replace("\r\n", "\n").strip()
    s = _ws_re.sub(" ", s)  # collapse whitespace
    return s

def dt_round_sec(d: dt.datetime) -> dt.datetime:
    if d.tzinfo is None:
        d = d.replace(tzinfo=LOCAL_TZ)
    return d.astimezone(LOCAL_TZ).replace(microsecond=0)

def dt_equal_sec(a: dt.datetime, b: dt.datetime) -> bool:
    return dt_round_sec(a) == dt_round_sec(b)

def parse_outlook_categories(raw: str) -> List[str]:
    # Outlook Categories is a comma-separated string in most setups
    if not raw:
        return []
    parts = [p.strip() for p in raw.split(",")]
    parts = [p for p in parts if p]
    # de-dupe while preserving order
    seen = set()
    out = []
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out

def normalize_categories(cats: List[str]) -> str:
    if not cats:
        return ""
    # Stable compare: lowercase + sorted
    return ",".join(sorted([c.strip().lower() for c in cats if c and c.strip()]))

def pick_primary_category(cats: List[str]) -> Optional[str]:
    if not cats:
        return None
    # Policy: first category wins
    return cats[0]

def categories_to_google_color_id(cats: List[str]) -> Optional[str]:
    primary = pick_primary_category(cats)
    if not primary:
        return DEFAULT_GCAL_COLOR_ID
    # exact match first
    if primary in OUTLOOK_CATEGORY_COLOR_MAP:
        return OUTLOOK_CATEGORY_COLOR_MAP[primary]
    # case-insensitive match
    for k, v in OUTLOOK_CATEGORY_COLOR_MAP.items():
        if k.strip().lower() == primary.strip().lower():
            return v
    return DEFAULT_GCAL_COLOR_ID

def google_color_id_to_outlook_category(color_id: Optional[str]) -> Optional[str]:
    if not color_id:
        return None
    if color_id in GCAL_COLOR_OUTLOOK_CATEGORY_MAP:
        return GCAL_COLOR_OUTLOOK_CATEGORY_MAP[color_id]
    return None

def rec_equivalent(a: EventRecord, b: EventRecord) -> bool:
    # If both sides provide categories/colors, include them to avoid patch loops
    cats_a = normalize_categories(a.categories or [])
    cats_b = normalize_categories(b.categories or [])
    color_a = (a.google_color_id or "").strip()
    color_b = (b.google_color_id or "").strip()

    return (
        normalize_text(a.title) == normalize_text(b.title)
        and a.all_day == b.all_day
        and normalize_text(a.location) == normalize_text(b.location)
        and normalize_text(a.description) == normalize_text(b.description)
        and dt_equal_sec(a.start, b.start)
        and dt_equal_sec(a.end, b.end)
        and cats_a == cats_b
        and color_a == color_b
    )

def event_key(rec: EventRecord) -> str:
    # stable dedupe key without timestamp()
    return "|".join([
        normalize_text(rec.title),
        "AD" if rec.all_day else "TD",
        dt_round_sec(rec.start).isoformat(),
        dt_round_sec(rec.end).isoformat(),
        normalize_text(rec.location),
        normalize_text(rec.description),
    ])

def should_ignore(rec: EventRecord) -> bool:
    now = dt.datetime.now(tz=LOCAL_TZ)
    if rec.end < now - dt.timedelta(days=LOOKBACK_DAYS):
        return True
    if rec.start > now + dt.timedelta(days=LOOKAHEAD_DAYS):
        return True
    return False


# ----------------------------
# Google helpers
# ----------------------------
def google_service():
    global LOCAL_TZ
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_PATH):
                raise RuntimeError(f"Missing credentials.json at: {CREDS_PATH}")
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    svc = build("calendar", "v3", credentials=creds, cache_discovery=False)

    # ---- DEBUG: what is 'primary' actually?
    try:
        cal = svc.calendars().get(calendarId="primary").execute()
        print(f"[google] primary calendar summary='{cal.get('summary')}' id='{cal.get('id')}' tz='{cal.get('timeZone')}'")
    except Exception as e:
        print(f"[google] primary calendar lookup failed: {e}")

    # ---- DEBUG: target calendar resolution
    try:
        cal2 = svc.calendars().get(calendarId=GOOGLE_CALENDAR_ID).execute()
        print(f"[google] target calendarId='{GOOGLE_CALENDAR_ID}' resolves to summary='{cal2.get('summary')}' id='{cal2.get('id')}' tz='{cal2.get('timeZone')}'")
        if not _SYNC_TZ_EXPLICIT:
            tz_name = (cal2.get("timeZone") or "").strip()
            tzinfo = _timezone_from_name(tz_name)
            if tzinfo:
                LOCAL_TZ = tzinfo
            elif tz_name:
                print(f"[config] invalid calendar timeZone '{tz_name}', keeping local timezone")
    except Exception as e:
        print(f"[google] target calendar lookup failed: {e}")

    return svc

def is_rate_limited(err: Exception) -> bool:
    if not isinstance(err, HttpError):
        return False
    s = str(err)
    return ("rateLimitExceeded" in s) or ("Rate Limit Exceeded" in s) or ("userRateLimitExceeded" in s)

def is_event_type_restriction(err: Exception) -> bool:
    if not isinstance(err, HttpError):
        return False
    s = str(err)
    return "eventTypeRestriction" in s or "event type must not have private extended properties" in s

def google_list_changes(svc) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    sync_token = (kv_get("google_sync_token") or "").strip()

    time_min = (dt.datetime.now(tz=LOCAL_TZ) - dt.timedelta(days=LOOKBACK_DAYS)).isoformat()
    time_max = (dt.datetime.now(tz=LOCAL_TZ) + dt.timedelta(days=LOOKAHEAD_DAYS)).isoformat()

    params = {
        "calendarId": GOOGLE_CALENDAR_ID,
        "singleEvents": True,
        "showDeleted": True,
        "maxResults": 2500,
    }

    items: List[Dict[str, Any]] = []
    next_sync_token = None

    try:
        page_token = None
        while True:
            if sync_token:
                req = svc.events().list(**params, syncToken=sync_token, pageToken=page_token)
            else:
                req = svc.events().list(**params, timeMin=time_min, timeMax=time_max, orderBy="updated", pageToken=page_token)

            resp = req.execute()
            items.extend(resp.get("items", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                next_sync_token = resp.get("nextSyncToken")
                break

    except HttpError as e:
        if "Sync token is no longer valid" in str(e) or "status 410" in str(e):
            kv_set("google_sync_token", "")
            return google_list_changes(svc)
        raise

    return items, next_sync_token

def google_to_record(ev: Dict[str, Any]) -> EventRecord:
    status = ev.get("status")
    deleted = (status == "cancelled") or ev.get("deleted", False)

    summary = ev.get("summary", "") or ""
    location = ev.get("location", "") or ""
    description = ev.get("description", "") or ""

    updated = isoparse(ev["updated"]).astimezone(LOCAL_TZ) if ev.get("updated") else dt.datetime.now(tz=LOCAL_TZ)

    all_day = False
    if "dateTime" in ev.get("start", {}):
        start = isoparse(ev["start"]["dateTime"]).astimezone(LOCAL_TZ)
        end = isoparse(ev["end"]["dateTime"]).astimezone(LOCAL_TZ)
    else:
        all_day = True
        start = isoparse(ev["start"]["date"]).replace(tzinfo=LOCAL_TZ)
        end = isoparse(ev["end"]["date"]).replace(tzinfo=LOCAL_TZ)

    ext = (ev.get("extendedProperties", {}) or {}).get("private", {}) or {}
    outlook_hint = ext.get(GOOGLE_EXT_OUTLOOK_ID)

    raw_cats = ext.get(GOOGLE_EXT_OUTLOOK_CATEGORIES, "") or ""
    cats = parse_outlook_categories(raw_cats)

    color_id = (ev.get("colorId") or "").strip() or None

    return EventRecord(
        side="google",
        uid=ev["id"],
        title=summary,
        start=start,
        end=end,
        all_day=all_day,
        location=location,
        description=description,
        last_modified=updated,
        deleted=deleted,
        outlook_entry_id_hint=outlook_hint,
        categories=cats,
        google_color_id=color_id,
    )

def google_upsert_event(
    svc,
    rec: EventRecord,
    google_id: Optional[str] = None,
    outlook_entry_id: Optional[str] = None
) -> Tuple[str, Optional[dt.datetime]]:
    # Determine google color from Outlook categories (if this record is coming from Outlook)
    # If rec already has google_color_id explicitly set, keep it.
    color_id = rec.google_color_id
    if not color_id:
        color_id = categories_to_google_color_id(rec.categories or [])

    body: Dict[str, Any] = {
        "summary": rec.title,
        "location": rec.location or None,
        "description": rec.description or None,
        "extendedProperties": {
            "private": {
                GOOGLE_EXT_SYNC_TAG: SYNC_TAG_VALUE,
            }
        }
    }

    if outlook_entry_id:
        body["extendedProperties"]["private"][GOOGLE_EXT_OUTLOOK_ID] = outlook_entry_id

    # store Outlook categories string in Google so we can round-trip it
    if rec.categories:
        body["extendedProperties"]["private"][GOOGLE_EXT_OUTLOOK_CATEGORIES] = ", ".join(rec.categories)

    if color_id:
        body["colorId"] = str(color_id)

    if rec.all_day:
        body["start"] = {"date": rec.start.date().isoformat()}
        body["end"] = {"date": rec.end.date().isoformat()}
    else:
        body["start"] = {"dateTime": rec.start.astimezone(LOCAL_TZ).isoformat()}
        body["end"] = {"dateTime": rec.end.astimezone(LOCAL_TZ).isoformat()}

    body = {k: v for k, v in body.items() if v is not None}

    if google_id:
        updated = svc.events().patch(calendarId=GOOGLE_CALENDAR_ID, eventId=google_id, body=body).execute()
        if GOOGLE_UPSERT_DEBUG:
            print(f"[google] PATCHED id={updated.get('id')} link={updated.get('htmlLink')}")
        updated_time = None
        if updated.get("updated"):
            updated_time = isoparse(updated["updated"]).astimezone(LOCAL_TZ)
        return updated["id"], updated_time
    else:
        created = svc.events().insert(calendarId=GOOGLE_CALENDAR_ID, body=body).execute()
        if GOOGLE_UPSERT_DEBUG:
            print(f"[google] CREATED id={created.get('id')} link={created.get('htmlLink')}")
        created_time = None
        if created.get("updated"):
            created_time = isoparse(created["updated"]).astimezone(LOCAL_TZ)
        return created["id"], created_time

def google_delete_event(svc, google_id: str):
    svc.events().delete(calendarId=GOOGLE_CALENDAR_ID, eventId=google_id).execute()

def google_find_by_outlook_entry_id(svc, outlook_entry_id: str, rec: EventRecord) -> Optional[str]:
    """
    Strong match: if we previously stamped the Outlook EntryID into Google extendedProperties.
    Search in a tight time window to keep it fast.
    """
    time_min = (rec.start - dt.timedelta(days=3)).astimezone(LOCAL_TZ).isoformat()
    time_max = (rec.end + dt.timedelta(days=3)).astimezone(LOCAL_TZ).isoformat()
    try:
        resp = svc.events().list(
            calendarId=GOOGLE_CALENDAR_ID,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            showDeleted=False,
            maxResults=50,
            privateExtendedProperty=f"{GOOGLE_EXT_OUTLOOK_ID}={outlook_entry_id}",
        ).execute()
        for ev in resp.get("items", []):
            return ev.get("id")
    except HttpError:
        raise
    except Exception:
        pass
    return None

def google_find_match_fuzzy(svc, rec: EventRecord) -> Optional[str]:
    """
    Fallback match: title+time window and then field compare.
    """
    title = normalize_text(rec.title)
    if not title:
        return None

    time_min = (rec.start - dt.timedelta(minutes=2)).astimezone(LOCAL_TZ).isoformat()
    time_max = (rec.end + dt.timedelta(minutes=2)).astimezone(LOCAL_TZ).isoformat()

    resp = svc.events().list(
        calendarId=GOOGLE_CALENDAR_ID,
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        q=title,
        maxResults=10
    ).execute()

    for ev in resp.get("items", []):
        grec = google_to_record(ev)
        if not grec.deleted and rec_equivalent(rec, grec):
            return ev["id"]
    return None


# ----------------------------
# Outlook helpers (COM)
# ----------------------------
def outlook_namespace():
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    return ns

def outlook_calendar_folder(ns):
    # 9 = olFolderCalendar
    return ns.GetDefaultFolder(9)

def dt_to_outlook_filter(d: dt.datetime) -> str:
    local = d.astimezone(LOCAL_TZ)
    return local.strftime("%m/%d/%Y %I:%M %p")

def outlook_list_changed_items(ns, since: dt.datetime) -> List[Any]:
    cal = outlook_calendar_folder(ns)
    items = cal.Items
    items.IncludeRecurrences = True
    items.Sort("[LastModificationTime]")

    flt = f"[LastModificationTime] >= '{dt_to_outlook_filter(since)}'"
    restricted = items.Restrict(flt)

    out = []
    try:
        for i in range(1, restricted.Count + 1):
            out.append(restricted.Item(i))
    except Exception:
        for it in restricted:
            out.append(it)
    return out

def outlook_list_deleted_items(ns, since: dt.datetime) -> List[Tuple[str, Optional[str]]]:
    # 3 = olFolderDeletedItems
    deleted_folder = ns.GetDefaultFolder(3)
    items = deleted_folder.Items
    items.IncludeRecurrences = True
    items.Sort("[LastModificationTime]")

    flt = f"[LastModificationTime] >= '{dt_to_outlook_filter(since)}'"
    restricted = items.Restrict(flt)

    out: List[Tuple[str, Optional[str]]] = []
    try:
        for i in range(1, restricted.Count + 1):
            it = restricted.Item(i)
            if outlook_item_is_appointment(it):
                entry_id = getattr(it, "EntryID", None)
                if entry_id:
                    gcal_hint = outlook_get_userprop_str(it, OUTLOOK_PROP_GCAL_ID)
                    out.append((entry_id, gcal_hint))
    except Exception:
        for it in restricted:
            if outlook_item_is_appointment(it):
                entry_id = getattr(it, "EntryID", None)
                if entry_id:
                    gcal_hint = outlook_get_userprop_str(it, OUTLOOK_PROP_GCAL_ID)
                    out.append((entry_id, gcal_hint))
    return out

def outlook_item_is_appointment(item) -> bool:
    # 26 = olAppointmentItem
    return getattr(item, "Class", None) == 26

def outlook_get_userprop_str(item, name: str) -> Optional[str]:
    try:
        ups = item.UserProperties
        if not ups:
            return None
        p = ups.Find(name)
        if p:
            val = p.Value
            return str(val) if val is not None else None
    except Exception:
        return None
    return None

def outlook_set_userprop_str(item, name: str, value: str):
    try:
        ups = item.UserProperties
        if not ups:
            return
        p = ups.Find(name)
        if not p:
            # 1 = olText
            p = ups.Add(name, 1, False)
        p.Value = value
    except Exception:
        return

def outlook_item_to_record(item) -> EventRecord:
    entry_id = item.EntryID
    subject = (getattr(item, "Subject", "") or "")
    location = (getattr(item, "Location", "") or "")
    body = (getattr(item, "Body", "") or "")

    start_utc = item.StartUTC
    end_utc = item.EndUTC
    start = start_utc.replace(tzinfo=timezone.utc).astimezone(LOCAL_TZ)
    end = end_utc.replace(tzinfo=timezone.utc).astimezone(LOCAL_TZ)

    all_day = bool(getattr(item, "AllDayEvent", False))

    last_mod = item.LastModificationTime
    if last_mod.tzinfo is None:
        last_mod = last_mod.replace(tzinfo=LOCAL_TZ)
    else:
        last_mod = last_mod.astimezone(LOCAL_TZ)

    gcal_hint = outlook_get_userprop_str(item, OUTLOOK_PROP_GCAL_ID)

    cats_raw = ""
    try:
        cats_raw = getattr(item, "Categories", "") or ""
    except Exception:
        cats_raw = ""
    cats = parse_outlook_categories(cats_raw)

    color_hint = outlook_get_userprop_str(item, OUTLOOK_PROP_GCAL_COLOR_ID)
    color_hint = (color_hint or "").strip() or None

    return EventRecord(
        side="outlook",
        uid=entry_id,
        title=subject,
        start=start,
        end=end,
        all_day=all_day,
        location=location,
        description=body,
        last_modified=last_mod,
        deleted=False,
        gcal_event_id_hint=gcal_hint,
        categories=cats,
        google_color_id=color_hint,
    )

def outlook_create_or_update(ns, rec: EventRecord, outlook_entry_id: Optional[str] = None, google_id_to_stamp: Optional[str] = None) -> str:
    """
    Safe writer: if the target item can't be updated (COM weirdness),
    create a fresh AppointmentItem rather than crashing the whole run.
    """
    app = ns.Application
    item = None

    if outlook_entry_id:
        try:
            item = ns.GetItemFromID(outlook_entry_id)
        except Exception:
            item = None

    if not item:
        item = app.CreateItem(1)  # 1 = olAppointmentItem

    def write_and_save(apt) -> str:
        if getattr(apt, "Class", None) != 26:
            raise TypeError(f"Not an AppointmentItem (Class={getattr(apt,'Class',None)})")

        apt.Subject = rec.title
        apt.Location = rec.location
        apt.Body = rec.description

        apt.AllDayEvent = rec.all_day
        apt.StartUTC = rec.start.astimezone(timezone.utc).replace(tzinfo=None)
        apt.EndUTC = rec.end.astimezone(timezone.utc).replace(tzinfo=None)

        # Category round-trip: if record has categories, set Outlook Categories string
        if rec.categories:
            try:
                apt.Categories = ", ".join(rec.categories)
            except Exception:
                pass
        else:
            # Optional reverse mapping: Google colorId -> Outlook category
            c = google_color_id_to_outlook_category(rec.google_color_id)
            if c:
                try:
                    apt.Categories = c
                except Exception:
                    pass

        if google_id_to_stamp:
            outlook_set_userprop_str(apt, OUTLOOK_PROP_GCAL_ID, google_id_to_stamp)

        if rec.google_color_id:
            outlook_set_userprop_str(apt, OUTLOOK_PROP_GCAL_COLOR_ID, rec.google_color_id)

        apt.Save()
        return apt.EntryID

    try:
        return write_and_save(item)
    except Exception as e:
        print(f"[outlook] cannot write target EntryID={outlook_entry_id} err={e} -> creating new item")
        fresh = app.CreateItem(1)
        return write_and_save(fresh)

def outlook_delete(ns, outlook_entry_id: str):
    try:
        item = ns.GetItemFromID(outlook_entry_id)
        item.Delete()
    except Exception:
        return


# ----------------------------
# Google dedupe (optional)
# ----------------------------
def google_list_window(svc) -> List[Dict[str, Any]]:
    time_min = (dt.datetime.now(tz=LOCAL_TZ) - dt.timedelta(days=LOOKBACK_DAYS)).isoformat()
    time_max = (dt.datetime.now(tz=LOCAL_TZ) + dt.timedelta(days=LOOKAHEAD_DAYS)).isoformat()

    items: List[Dict[str, Any]] = []
    page_token = None

    while True:
        resp = svc.events().list(
            calendarId=GOOGLE_CALENDAR_ID,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            showDeleted=False,
            maxResults=DEDUPE_MAX_RESULTS,
            pageToken=page_token
        ).execute()

        items.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    return items

def google_dedupe_window(svc):
    print(f"[dedupe] starting (dry_run={DEDUPE_DRY_RUN}) within window -{LOOKBACK_DAYS}d .. +{LOOKAHEAD_DAYS}d")

    mapped_google_ids = db_get_all_google_ids_in_mapping()

    raw = google_list_window(svc)
    recs: List[EventRecord] = []
    for ev in raw:
        try:
            r = google_to_record(ev)
            if not r.deleted and not should_ignore(r):
                recs.append(r)
        except Exception as e:
            print(f"[dedupe] skip bad google event id={ev.get('id')} err={e}")

    groups: Dict[str, List[EventRecord]] = {}
    for r in recs:
        k = event_key(r)
        groups.setdefault(k, []).append(r)

    dup_groups = {k: v for k, v in groups.items() if len(v) > 1}
    print(f"[dedupe] scanned {len(recs)} events, found {len(dup_groups)} duplicate groups")

    delete_count = 0

    for k, lst in dup_groups.items():
        mapped = [r for r in lst if r.uid in mapped_google_ids]
        if mapped:
            keeper = sorted(mapped, key=lambda x: x.last_modified, reverse=True)[0]
        else:
            keeper = sorted(lst, key=lambda x: x.last_modified, reverse=True)[0]

        to_delete = [r for r in lst if r.uid != keeper.uid]

        print(f"[dedupe] keep {keeper.uid} ({keeper.title!r} {dt_round_sec(keeper.start)}..{dt_round_sec(keeper.end)}) delete {len(to_delete)} others")

        for r in to_delete:
            delete_count += 1
            if DEDUPE_DRY_RUN:
                print(f"[dedupe] DRY RUN delete google eventId={r.uid}")
                continue
            try:
                google_delete_event(svc, r.uid)
                if r.uid in mapped_google_ids:
                    map_delete_by_google(r.uid)
                time.sleep(DEDUPE_SLEEP_BETWEEN_DELETES_S)
            except HttpError as e:
                print(f"[dedupe] delete failed eventId={r.uid} err={e}")
                if is_rate_limited(e):
                    raise
            except Exception as e:
                print(f"[dedupe] delete failed eventId={r.uid} err={e}")

    print(f"[dedupe] done: deleted {delete_count} events (dry_run={DEDUPE_DRY_RUN})")


# ----------------------------
# One sync pass
# ----------------------------
def sync_once() -> Tuple[bool, Optional[str]]:
    db_init()

    pythoncom.CoInitialize()
    ns = outlook_namespace()
    gsvc = google_service()

    sync_start = dt.datetime.now(tz=LOCAL_TZ)

    last_outlook = kv_get("outlook_last_poll")
    if last_outlook:
        try:
            since = isoparse(last_outlook).astimezone(LOCAL_TZ)
        except Exception:
            since = sync_start - dt.timedelta(days=LOOKBACK_DAYS)
    else:
        since = sync_start - dt.timedelta(days=LOOKBACK_DAYS)

    google_token_safe_to_advance = True
    next_sync_token: Optional[str] = None
    outlook_updated_from_google: Set[str] = set()

    try:
        # --- Outlook changes
        outlook_changed: List[EventRecord] = []
        deleted_outlook_items: List[Tuple[str, Optional[str]]] = []
        try:
            for it in outlook_list_changed_items(ns, since):
                try:
                    if outlook_item_is_appointment(it):
                        rec = outlook_item_to_record(it)
                        if not should_ignore(rec):
                            outlook_changed.append(rec)
                except Exception as e:
                    print(f"[outlook] bad item EntryID={getattr(it,'EntryID',None)} Subject={getattr(it,'Subject',None)} err={e}")
            deleted_outlook_items = outlook_list_deleted_items(ns, since)
        except Exception as e:
            print(f"[outlook] error reading changes: {e}")

        # --- Google changes
        google_items, next_sync_token = google_list_changes(gsvc)
        google_changed: List[EventRecord] = []
        for ev in google_items:
            try:
                rec = google_to_record(ev)
                if not should_ignore(rec):
                    google_changed.append(rec)
            except Exception as e:
                print(f"[google] bad event id={ev.get('id')} err={e}")

        # --- Google -> Outlook
        for grec in google_changed:
            try:
                mapped = map_get_by_google(grec.uid)
                outlook_id = None
                last_outlook_mod = None
                last_google_mod = None

                if grec.outlook_entry_id_hint:
                    outlook_id = grec.outlook_entry_id_hint

                if mapped:
                    outlook_id_from_map, last_outlook_mod_raw, last_google_mod_raw = mapped
                    last_outlook_mod = parse_sync_mod(last_outlook_mod_raw)
                    last_google_mod = parse_sync_mod(last_google_mod_raw)
                    if not outlook_id:
                        outlook_id = outlook_id_from_map

                if last_google_mod and last_google_mod >= grec.last_modified:
                    continue

                if grec.deleted:
                    if outlook_id:
                        print(f"[sync] Google deleted -> Outlook delete: {grec.uid} -> {outlook_id}")
                        outlook_delete(ns, outlook_id)
                    map_delete_by_google(grec.uid)
                    continue

                if outlook_id:
                    try:
                        oitem = ns.GetItemFromID(outlook_id)
                        orec = outlook_item_to_record(oitem)
                    except Exception:
                        orec = None

                    if orec and orec.last_modified > grec.last_modified:
                        continue

                    if orec and rec_equivalent(orec, grec):
                        map_upsert(outlook_id, grec.uid, orec.last_modified, grec.last_modified)
                        continue

                    print(f"[sync] Google update -> Outlook patch: {grec.uid} -> {outlook_id}")
                    new_outlook_id = outlook_create_or_update(ns, grec, outlook_id, google_id_to_stamp=grec.uid)
                    # protect against UNIQUE constraint by ensuring outlook_id is the actual saved EntryID
                    updated_mod = None
                    try:
                        updated_item = ns.GetItemFromID(new_outlook_id)
                        updated_mod = outlook_item_to_record(updated_item).last_modified
                    except Exception:
                        updated_mod = None
                    map_upsert(new_outlook_id, grec.uid, updated_mod, grec.last_modified)
                    if new_outlook_id:
                        outlook_updated_from_google.add(new_outlook_id)

                else:
                    print(f"[sync] Google new -> Outlook create: {grec.uid}")
                    outlook_id_new = outlook_create_or_update(ns, grec, None, google_id_to_stamp=grec.uid)
                    created_mod = None
                    try:
                        created_item = ns.GetItemFromID(outlook_id_new)
                        created_mod = outlook_item_to_record(created_item).last_modified
                    except Exception:
                        created_mod = None
                    map_upsert(outlook_id_new, grec.uid, created_mod, grec.last_modified)
                    if outlook_id_new:
                        outlook_updated_from_google.add(outlook_id_new)

            except Exception as e:
                print(f"[sync] Google->Outlook failed googleId={grec.uid} err={e}")

        # --- Outlook deletions -> Google
        if deleted_outlook_items:
            for outlook_id, gcal_hint in deleted_outlook_items:
                google_id = gcal_hint
                if not google_id:
                    mapped = map_get_by_outlook(outlook_id)
                    if mapped:
                        google_id, _, _ = mapped
                if not google_id:
                    continue
                try:
                    print(f"[sync] Outlook deleted -> Google delete: {outlook_id} -> {google_id}")
                    google_delete_event(gsvc, google_id)
                except HttpError as e:
                    print(f"[sync] Outlook delete -> Google delete failed EntryID={outlook_id} err={e}")
                    if is_rate_limited(e):
                        google_token_safe_to_advance = False
                        raise
                except Exception as e:
                    print(f"[sync] Outlook delete -> Google delete failed EntryID={outlook_id} err={e}")
                else:
                    map_delete_by_outlook(outlook_id)
                    map_delete_by_google(google_id)

        # --- Outlook -> Google
        for orec in outlook_changed:
            try:
                if orec.uid in outlook_updated_from_google:
                    continue
                google_id = orec.gcal_event_id_hint

                mapped = map_get_by_outlook(orec.uid)
                last_outlook_mod = None
                last_google_mod = None
                if mapped:
                    google_id_from_map, last_outlook_mod_raw, last_google_mod_raw = mapped
                    last_outlook_mod = parse_sync_mod(last_outlook_mod_raw)
                    last_google_mod = parse_sync_mod(last_google_mod_raw)
                    if not google_id:
                        google_id = google_id_from_map

                if last_outlook_mod and last_outlook_mod >= orec.last_modified:
                    continue

                if not google_id:
                    try:
                        google_id = google_find_by_outlook_entry_id(gsvc, orec.uid, orec)
                    except HttpError as e:
                        print(f"[google] lookup by extendedProperties failed EntryID={orec.uid} err={e}")
                        if is_rate_limited(e):
                            google_token_safe_to_advance = False
                            raise

                if not google_id:
                    try:
                        google_id = google_find_match_fuzzy(gsvc, orec)
                    except HttpError as e:
                        print(f"[google] fuzzy match failed EntryID={orec.uid} err={e}")
                        if is_rate_limited(e):
                            google_token_safe_to_advance = False
                            raise

                if google_id:
                    print(f"[sync] Outlook update -> Google patch: {orec.uid} -> {google_id}")
                    try:
                        new_gid, google_updated = google_upsert_event(
                            gsvc,
                            orec,
                            google_id=google_id,
                            outlook_entry_id=orec.uid
                        )
                        # Stamp both ID and color into Outlook for persistence
                        updated_outlook_mod = orec.last_modified
                        try:
                            it = ns.GetItemFromID(orec.uid)
                            outlook_set_userprop_str(it, OUTLOOK_PROP_GCAL_ID, new_gid)
                            color_to_stamp = categories_to_google_color_id(orec.categories or [])
                            if color_to_stamp:
                                outlook_set_userprop_str(it, OUTLOOK_PROP_GCAL_COLOR_ID, str(color_to_stamp))
                            it.Save()
                            updated_outlook_mod = outlook_item_to_record(it).last_modified
                        except Exception:
                            pass
                        map_upsert(
                            orec.uid,
                            new_gid,
                            updated_outlook_mod,
                            google_updated or dt.datetime.now(tz=LOCAL_TZ)
                        )
                    except HttpError as e:
                        print(f"[sync] Outlook->Google patch failed EntryID={orec.uid} err={e}")
                        if is_event_type_restriction(e):
                            map_upsert(orec.uid, google_id, orec.last_modified, dt.datetime.now(tz=LOCAL_TZ))
                            continue
                        if is_rate_limited(e):
                            google_token_safe_to_advance = False
                            raise
                else:
                    print(f"[sync] Outlook new -> Google create: {orec.uid}")
                    try:
                        gid, google_updated = google_upsert_event(
                            gsvc,
                            orec,
                            google_id=None,
                            outlook_entry_id=orec.uid
                        )
                        updated_outlook_mod = orec.last_modified
                        try:
                            it = ns.GetItemFromID(orec.uid)
                            outlook_set_userprop_str(it, OUTLOOK_PROP_GCAL_ID, gid)
                            color_to_stamp = categories_to_google_color_id(orec.categories or [])
                            if color_to_stamp:
                                outlook_set_userprop_str(it, OUTLOOK_PROP_GCAL_COLOR_ID, str(color_to_stamp))
                            it.Save()
                            updated_outlook_mod = outlook_item_to_record(it).last_modified
                        except Exception:
                            pass
                        map_upsert(
                            orec.uid,
                            gid,
                            updated_outlook_mod,
                            google_updated or dt.datetime.now(tz=LOCAL_TZ)
                        )
                    except HttpError as e:
                        print(f"[sync] Outlook->Google create failed EntryID={orec.uid} err={e}")
                        if is_event_type_restriction(e):
                            continue
                        if is_rate_limited(e):
                            google_token_safe_to_advance = False
                            raise

            except HttpError:
                raise
            except Exception as e:
                print(f"[sync] Outlook->Google failed EntryID={orec.uid} err={e}")

        return google_token_safe_to_advance, next_sync_token

    finally:
        kv_set("outlook_last_poll", sync_start.isoformat())
        pythoncom.CoUninitialize()


# ----------------------------
# Main loop
# ----------------------------
def main():
    print("Outlook <-> Google Calendar sync daemon starting.")
    print(f"- Poll: {POLL_SECONDS}s | Window: -{LOOKBACK_DAYS}d .. +{LOOKAHEAD_DAYS}d | Google calendar: {GOOGLE_CALENDAR_ID}")
    print(f"- Persistent tag: {SYNC_TAG_VALUE} | Dedupe on start: {DEDUPE_ON_START} (dry_run={DEDUPE_DRY_RUN})")
    if GOOGLE_UPSERT_DEBUG:
        print("- Google upsert debug: ON (prints htmlLink on create/patch)")
    if OUTLOOK_CATEGORY_COLOR_MAP:
        print(f"- Category->Color mappings loaded: {len(OUTLOOK_CATEGORY_COLOR_MAP)}")
    if DEFAULT_GCAL_COLOR_ID:
        print(f"- Default Google colorId: {DEFAULT_GCAL_COLOR_ID}")

    db_init()
    gsvc = google_service()

    if DEDUPE_ON_START:
        backoff = 1
        while True:
            try:
                google_dedupe_window(gsvc)
                break
            except Exception as e:
                print(f"[dedupe] error: {e}")
                if is_rate_limited(e):
                    sleep_s = min(900, backoff) + random.uniform(0, 1)
                    print(f"[backoff] Google rate limit during dedupe. Sleeping {sleep_s:.1f}s")
                    time.sleep(sleep_s)
                    backoff = min(900, backoff * 2)
                    continue
                break

    backoff = 1

    while True:
        try:
            google_token_safe_to_advance, next_sync_token = sync_once()

            if next_sync_token and google_token_safe_to_advance:
                kv_set("google_sync_token", next_sync_token)

            backoff = 1
            time.sleep(POLL_SECONDS)

        except KeyboardInterrupt:
            print("Stopping.")
            return

        except Exception as e:
            print(f"[fatal] sync loop error: {e}")

            if is_rate_limited(e):
                sleep_s = min(900, backoff) + random.uniform(0, 1)
                print(f"[backoff] Google rate limit. Sleeping {sleep_s:.1f}s")
                time.sleep(sleep_s)
                backoff = min(900, backoff * 2)
                continue

            time.sleep(min(60, POLL_SECONDS))


if __name__ == "__main__":
    main()
