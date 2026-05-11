#!/usr/bin/env python3
"""
Chấm Công — Headless Attendance Processor
Designed for GitHub Actions / Task Scheduler.
Pulls yesterday's attendance from MISA Timesheet Detail, exports xlsx, emails it.

Required .env keys:
    MISA_CLIENT_ID
    MISA_CLIENT_SECRET
    MISA_TIMESHEET_ID       ← non-technical user edits this when timesheet rolls over

    EMAIL_SMTP_HOST         e.g. smtp.gmail.com
    EMAIL_SMTP_PORT         e.g. 587
    EMAIL_USERNAME
    EMAIL_PASSWORD
    EMAIL_FROM
    EMAIL_TO                comma-separated list of recipients
    EMAIL_CC                (optional) comma-separated
"""

import re
import os
import sys
import hmac
import uuid
import json
import hashlib
import base64
import smtplib
import urllib.request
import urllib.error
from datetime import datetime, date, time, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── PATHS ─────────────────────────────────────────────────────────────────────
def _script_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

SCRIPT_DIR = _script_dir()
TEMP_DIR   = os.path.join(SCRIPT_DIR, "temp")
os.makedirs(TEMP_DIR, exist_ok=True)

# ── CONFIG (hardcoded defaults) ───────────────────────────────────────────────
WORK_HOURS          = 8
SAT_WORK_HOURS      = 4
SAT_EARLY_THRESHOLD = time(12, 30)
GRACE_MINUTES       = 15
LATE_THRESHOLD      = time(8, 30)
EARLY_THRESHOLD     = time(17, 30)
ABSENCE_MINUTES     = 240
LONGITUDE           = 106.69385465934707
LATITUDE            = 10.73676344055389

MISA_BASE_URL = "https://amisapp.misa.vn/APIS/TimesheetOpenAPI/api/Open"

MISA_LEAVE_KEY_NAMES = {
    1: "PTO", 2: "Unpaid Leave", 3: "Marriage Leave",
    4: "Child Marriage Leave", 5: "Bereavement Leave",
    6: "Social Insurance Leave", 7: "Maternity Leave", 8: "Compensatory Leave",
}

EXCLUDED_NAME_PATTERNS = [
    r"khachvip\d*",
    r"tapvu",
    r"^juli$",
]

STATUS_COLOURS = {
    "normal":         "C6EFCE", "late":           "FFE699",
    "early":          "FFE699", "late/early":      "FFE699",
    "absence":        "FFC7CE", "business trip":   "DDEBF7",
    "pto":            "E2EFDA", "unpaid leave":    "FCE4D6",
    "remote":         "EAF0FB", "approved late":   "FFF2CC",
    "approved early": "FFF2CC", "leave":           "F2DCDB",
}
STATUS_FONT_COLOURS = {
    "normal":         "276221", "late":           "9C5700",
    "early":          "9C5700", "late/early":      "9C5700",
    "absence":        "9C0006", "business trip":   "1F4E79",
    "pto":            "375623", "unpaid leave":    "843C0C",
    "remote":         "17375E", "approved late":   "7F6000",
    "approved early": "7F6000", "leave":           "632523",
}

STATUS_VI = {
    "normal": "Bình thường", "late": "Muộn", "early": "Về sớm",
    "late/early": "Muộn/Về sớm", "absence": "Vắng mặt",
    "business trip": "Công tác", "pto": "Nghỉ phép",
    "unpaid leave": "Nghỉ không lương", "remote": "Làm việc tại nhà",
    "approved late": "Đã phê duyệt muộn", "approved early": "Đã phê duyệt về sớm",
    "leave": "Nghỉ", "marriage leave": "Nghỉ kết hôn",
    "child marriage leave": "Nghỉ kết hôn con", "bereavement leave": "Nghỉ tang",
    "social insurance leave": "Nghỉ bảo hiểm", "maternity leave": "Nghỉ thai sản",
    "compensatory leave": "Nghỉ bù",
}

# ── ENV LOADER ────────────────────────────────────────────────────────────────
_ENV_CACHE: dict[str, str] | None = None

def _load_env_all() -> dict[str, str]:
    global _ENV_CACHE
    if _ENV_CACHE is not None:
        return _ENV_CACHE
    candidates = [
        os.path.join(SCRIPT_DIR, ".env"),
        os.path.join(os.getcwd(), ".env"),
    ]
    env: dict[str, str] = {}
    for path in candidates:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith("#") or "=" not in line:
                            continue
                        k, _, v = line.partition("=")
                        env[k.strip()] = v.strip().strip('"').strip("'")
                _ENV_CACHE = env
                return env
            except Exception:
                pass
    _ENV_CACHE = env
    return env

def _env(key: str) -> str:
    val = _load_env_all().get(key, "")
    if not val:
        raise ValueError(f"Missing .env key: {key}")
    return val

def _env_optional(key: str, default: str = "") -> str:
    return _load_env_all().get(key, default)

# ── SCHEDULE CHECKS ───────────────────────────────────────────────────────────
def _load_date_list(filename: str) -> set[date]:
    """Load a newline-separated list of YYYY-MM-DD dates from a file next to the script."""
    result = set()
    for path in (os.path.join(SCRIPT_DIR, filename), os.path.join(os.getcwd(), filename)):
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith("#"):
                            continue
                        try:
                            result.add(datetime.strptime(line, "%Y-%m-%d").date())
                        except ValueError:
                            pass
                return result
            except Exception:
                pass
    return result

def should_run(target_date: date) -> tuple[bool, str]:
    """
    Determine whether to run for target_date.
    Returns (should_run: bool, reason: str).
    Rules:
      - Sunday → always skip
      - In holidays.txt → skip
      - Saturday → only run if date is in saturdays.txt
      - Monday–Friday → run
    """
    weekday = target_date.weekday()  # 0=Mon … 6=Sun

    if weekday == 6:
        return False, f"{target_date} is Sunday — skipping"

    holidays = _load_date_list("holidays.txt")
    if target_date in holidays:
        return False, f"{target_date} is a holiday — skipping"

    if weekday == 5:
        saturdays = _load_date_list("saturdays.txt")
        if target_date not in saturdays:
            return False, f"{target_date} is Saturday but not in saturdays.txt — skipping"
        return True, f"{target_date} is an exception Saturday — running"

    return True, f"{target_date} is a workday — running"

# ── HELPERS ───────────────────────────────────────────────────────────────────
def is_excluded(name: str) -> bool:
    name_lower = str(name).strip().lower()
    return any(re.fullmatch(p, name_lower) for p in EXCLUDED_NAME_PATTERNS)

def parse_hms(t_str: str) -> time | None:
    if not t_str or t_str == "N/A":
        return None
    m = re.match(r'(\d{1,2}):(\d{2}):(\d{2})', t_str.strip())
    if not m:
        return None
    h, mn = int(m.group(1)), int(m.group(2))
    if h > 23:
        return None
    return time(h, mn, int(m.group(3)))

def time_to_minutes(t: time) -> int:
    return t.hour * 60 + t.minute

def _iso_to_date(iso_str: str) -> date | None:
    if not iso_str:
        return None
    m = re.match(r'(\d{4}-\d{2}-\d{2})', str(iso_str))
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d").date()
    except ValueError:
        return None

def fmt_hm(t_str: str) -> str:
    if not t_str or t_str == "N/A":
        return "N/A"
    m = re.match(r'(\d{1,2}:\d{2})', t_str.strip())
    return m.group(1) if m else "N/A"

def normalize_id(raw: str) -> str:
    s = str(raw).strip().lstrip("'")
    try:
        return str(int(s))
    except ValueError:
        return s

def style_cell(cell, bold=False, h_align="center", font_colour="000000",
               fill_colour=None, border=None, size=10):
    cell.font = Font(name="Arial", bold=bold, color=font_colour, size=size)
    cell.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=False)
    if fill_colour:
        cell.fill = PatternFill("solid", fgColor=fill_colour)
    if border:
        cell.border = border

# ── STATUS COMPUTATION ────────────────────────────────────────────────────────
def compute_status(check_in_str: str, check_out_str: str, target_date: date | None = None):
    """Returns (status, late_str, early_min, late_min)."""
    ci = parse_hms(check_in_str)
    co = parse_hms(check_out_str)

    if ci is None and co is None:
        return "Absence", "0 min", 0, 0

    late_min = early_min = 0

    early_threshold = EARLY_THRESHOLD
    if target_date and target_date.weekday() == 5:
        early_threshold = SAT_EARLY_THRESHOLD

    if ci is not None and ci > LATE_THRESHOLD:
        late_min = time_to_minutes(ci) - time_to_minutes(LATE_THRESHOLD)
    if co is not None and co < early_threshold:
        early_min = time_to_minutes(early_threshold) - time_to_minutes(co)

    if 0 < late_min  <= GRACE_MINUTES: late_min  = 0
    if 0 < early_min <= GRACE_MINUTES: early_min = 0

    if   late_min >= ABSENCE_MINUTES or early_min >= ABSENCE_MINUTES: status = "Absence"
    elif late_min > 0 and early_min > 0:                               status = "Late/Early"
    elif late_min > 0:                                                  status = "Late"
    elif early_min > 0:                                                 status = "Early"
    else:                                                               status = "Normal"

    return status, (f"{late_min} min" if late_min > 0 else "0 min"), early_min, late_min

# ── MISA API ──────────────────────────────────────────────────────────────────
def _misa_headers() -> dict:
    tid    = str(uuid.uuid4())
    secret = _env("MISA_CLIENT_SECRET")
    token  = base64.b64encode(
        hmac.new(secret.encode(), tid.encode(), hashlib.sha256).digest()
    ).decode()
    return {
        "Content-Type":    "application/json",
        "x-clientid":      _env("MISA_CLIENT_ID"),
        "x-transactionid": tid,
        "x-token":         token,
    }

def _misa_post(endpoint: str, payload: dict | list) -> dict:
    url = f"{MISA_BASE_URL}/{endpoint}"
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode(), headers=_misa_headers(), method="POST")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode()
    except urllib.error.HTTPError as e:
        raise ValueError(f"MISA HTTP {e.code}: {e.read().decode()[:200]}")
    except urllib.error.URLError as e:
        raise ValueError(f"MISA connection error: {e.reason}")
    result = json.loads(body)
    if not result.get("Success", False):
        msg = result.get("UserMessage") or result.get("SystemMessage") or "Unknown"
        raise ValueError(f"MISA error (SubCode {result.get('SubCode',0)}): {msg}")
    return result

def _paginate_misa(endpoint: str, payload: dict,
                   page_size: int = 100, max_pages: int = 50) -> list:
    records = []
    for page in range(1, max_pages + 1):
        payload["PageSize"]  = page_size
        payload["PageIndex"] = page
        result    = _misa_post(endpoint, payload)
        data      = result.get("Data", {})
        page_data = data.get("PageData", [])
        total     = data.get("Total", 0)
        records.extend(page_data)
        if len(records) >= total or not page_data:
            break
    return records

def _misa_get(endpoint: str) -> dict:
    url = f"{MISA_BASE_URL}/{endpoint}"
    req = urllib.request.Request(url, headers=_misa_headers(), method="GET")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = r.read().decode()
    except urllib.error.HTTPError as e:
        raise ValueError(f"MISA HTTP {e.code}: {e.read().decode()[:200]}")
    except urllib.error.URLError as e:
        raise ValueError(f"MISA connection error: {e.reason}")
    result = json.loads(body)
    if not result.get("Success", False):
        msg = result.get("UserMessage") or result.get("SystemMessage") or "Unknown"
        raise ValueError(f"MISA error (SubCode {result.get('SubCode',0)}): {msg}")
    return result

# ── MISA EMPLOYEE STORE ───────────────────────────────────────────────────────
_MISA_EMP_STORE:  dict[str, dict] = {}
_TDS_TO_SHIFT:    dict[str, str]  = {}
_MISA_NAME_CACHE: dict[str, str]  = {}
# CHECK_TO_EMP: shift_code → tds_code (used for anchoring report rows)
CHECK_TO_EMP: dict[str, str] = {}

def load_misa_employees() -> tuple[bool, str]:
    global _MISA_EMP_STORE, _TDS_TO_SHIFT, _MISA_NAME_CACHE, CHECK_TO_EMP
    try:
        result    = _misa_get("get-data-employee")
        employees = result.get("Data", [])
        if not isinstance(employees, list):
            return False, "Unexpected response format from get-data-employee"

        store: dict[str, dict] = {}
        tds_to_shift: dict[str, str] = {}
        name_cache: dict[str, str]   = {}
        check_to_emp: dict[str, str] = {}

        for emp in employees:
            tds        = str(emp.get("EmployeeCode", "")).strip()
            full_name  = str(emp.get("FullName", "")).strip()
            shift_code = str(emp.get("ShiftCode", "")).strip()
            status_id  = int(emp.get("EmployeeStatusID", 0) or 0)

            if not tds or not shift_code:
                continue

            shift_norm = normalize_id(shift_code)
            store[shift_norm] = {"tds": tds, "full_name": full_name, "status_id": status_id}
            tds_to_shift[tds]       = shift_norm
            name_cache[shift_norm]  = full_name
            check_to_emp[shift_norm] = tds

        _MISA_EMP_STORE  = store
        _TDS_TO_SHIFT    = tds_to_shift
        _MISA_NAME_CACHE = name_cache
        CHECK_TO_EMP     = check_to_emp
        return True, f"{len(store)} employees loaded from MISA"

    except Exception as e:
        return False, str(e)

def _to_check_id(tds_code: str) -> str:
    tds    = str(tds_code).strip()
    mapped = _TDS_TO_SHIFT.get(tds, "")
    if mapped:
        return normalize_id(mapped)
    return normalize_id(tds)

# ── MISA TIMESHEET DETAIL ─────────────────────────────────────────────────────
def fetch_timesheet_detail(timesheet_id: int) -> dict:
    payload = {
        "PageSize":    -1,
        "PageIndex":   1,
        "Filter":      None,
        "CustomFilter": None,
        "QuickSearch": {},
        "CustomParam": {"TimeSheetID": timesheet_id},
    }
    return _misa_post("get-data-timesheet-detail", payload)

def process_timesheet_detail(page_data: list, target_date: date) -> tuple[dict, dict]:
    """
    Parse timesheet detail records for target_date.
    Returns (punch_records, form_records).
    """
    day_key = f"Day{target_date.day}"

    punch_records: dict = {}
    form_records:  dict = {}

    FULL_DAY_FORMS = {
        "PTO", "Unpaid Leave", "Leave", "Business Trip",
        "Marriage Leave", "Child Marriage Leave", "Bereavement Leave",
        "Social Insurance Leave", "Maternity Leave", "Compensatory Leave",
    }

    def _hms_to_min(raw) -> int | None:
        if not raw:
            return None
        m = re.search(r'(\d{1,2}):(\d{2})(?::\d{2})?', str(raw).strip())
        if not m:
            return None
        try:
            return int(m.group(1)) * 60 + int(m.group(2))
        except Exception:
            return None

    for emp in page_data:
        tds       = str(emp.get("EmployeeCode", "")).strip()
        full_name = str(emp.get("FullName", "")).strip()
        cid       = _to_check_id(tds)
        if not cid:
            continue

        display_name = _MISA_NAME_CACHE.get(cid) or full_name

        data_daily_raw = emp.get("DataDaily", "")
        if not data_daily_raw:
            punch_records[cid] = {"full_name": display_name, "check_in": "N/A", "check_out": "N/A"}
            continue

        try:
            data_daily = json.loads(data_daily_raw)
        except (json.JSONDecodeError, TypeError):
            punch_records[cid] = {"full_name": display_name, "check_in": "N/A", "check_out": "N/A"}
            continue

        entries = data_daily.get(day_key, [])
        if not entries:
            punch_records[cid] = {"full_name": display_name, "check_in": "N/A", "check_out": "N/A"}
            continue

        all_starts: list[str] = []
        all_ends:   list[str] = []
        actual_hours: float   = 0.0
        shift_code            = ""
        day_type              = 1
        shift_work_min: int | None = None

        for entry in entries:
            shift_code    = str(entry.get("WorkingShiftCode", "") or "").strip()
            day_type      = int(entry.get("DayType", 1) or 1)
            actual_hours += float(entry.get("ActualHours", 0) or 0)

            s_raw = str(entry.get("CheckStartTime", "") or "")
            e_raw = str(entry.get("CheckEndTime",   "") or "")
            m_s   = re.search(r'T(\d{2}:\d{2}:\d{2})', s_raw)
            m_e   = re.search(r'T(\d{2}:\d{2}:\d{2})', e_raw)
            if m_s: all_starts.append(m_s.group(1))
            if m_e: all_ends.append(m_e.group(1))

            if shift_work_min is None:
                st  = _hms_to_min(entry.get("StartTime"))
                et  = _hms_to_min(entry.get("EndTime"))
                bst = _hms_to_min(entry.get("StartBreakTime"))
                bet = _hms_to_min(entry.get("EndBreakTime"))
                if st is not None and et is not None:
                    work = et - st
                    if bst is not None and bet is not None and bet > bst:
                        work -= (bet - bst)
                    if work > 0:
                        shift_work_min = work

            # Parse embedded leave info from ListAttendance
            list_att_raw = str(entry.get("ListAttendance", "") or "")
            if list_att_raw and list_att_raw != "[]":
                try:
                    att_list = json.loads(list_att_raw)
                    for att in att_list:
                        att_name_en = str(att.get("AttendanceTypeName_EN", "") or "").strip()
                        att_name_vn = str(att.get("AttendanceTypeName",    "") or "").strip()
                        leave_day   = float(att.get("LeaveDay", 0) or 0)
                        is_business = bool(att.get("IsBussiness", False))
                        dk          = att.get("DictionaryKey")

                        if is_business:
                            label = "Business Trip"
                        elif dk == 1:
                            label = "PTO"
                        elif dk == 2:
                            label = "Unpaid Leave"
                        elif att_name_en:
                            label = att_name_en
                        elif att_name_vn:
                            label = att_name_vn
                        else:
                            label = "Leave"

                        if leave_day > 0 and cid not in form_records:
                            form_records[cid] = {
                                "tds_code":    tds,
                                "full_name":   display_name,
                                "status":      label,
                                "waive_late":  False,
                                "waive_early": False,
                            }
                except (json.JSONDecodeError, TypeError, ValueError):
                    pass

        check_in  = min(all_starts) if all_starts else "N/A"
        check_out = max(all_ends)   if all_ends   else "N/A"
        attended_min = int(round(actual_hours * 60))

        punch_records[cid] = {
            "full_name":      display_name,
            "check_in":       check_in,
            "check_out":      check_out,
            "attended_min":   attended_min,
            "shift_work_min": shift_work_min,
            "shift_code":     shift_code,
            "day_type":       day_type,
        }

    # ── External MISA forms (Attendance, MissionAllowance, LateInEarlyOut) ────
    date_str = target_date.strftime("%Y-%m-%d")

    def _fetch(sub: str) -> list:
        try:
            return _paginate_misa("get-data-application", {
                "Filter": None, "CustomFilter": None, "QuickSearch": {},
                "CustomParam": {"SubSystemCode": sub,
                                "FromDate": date_str, "ToDate": date_str}
            })
        except Exception:
            return []

    def _register_form(tds: str, full_name: str, status_label: str,
                       waive_late=False, waive_early=False):
        cid = _to_check_id(tds)
        if not cid:
            return
        existing = form_records.get(cid)
        if existing is None:
            form_records[cid] = {
                "tds_code":    tds,
                "full_name":   full_name,
                "status":      status_label,
                "waive_late":  waive_late,
                "waive_early": waive_early,
            }
        else:
            if waive_late:  existing["waive_late"]  = True
            if waive_early: existing["waive_early"] = True
            if existing["status"] is None:
                existing["status"] = status_label

    for form in _fetch("Attendance"):
        if form.get("Status") != 2: continue
        tds    = str(form.get("EmployeeCode", "")).strip()
        name   = str(form.get("FullName", "")).strip()
        from_d = _iso_to_date(form.get("FromDate", ""))
        to_d   = _iso_to_date(form.get("ToDate", ""))
        if from_d and to_d and not (from_d <= target_date <= to_d):
            continue
        dk = form.get("DictionaryKey")
        if dk == 1:   label = "PTO"
        elif dk == 2: label = "Unpaid Leave"
        else:         label = str(form.get("AttendanceTypeName") or MISA_LEAVE_KEY_NAMES.get(dk, "Leave"))
        _register_form(tds, name, label)

    for form in _fetch("MissionAllowance"):
        if form.get("Status") != 2: continue
        tds    = str(form.get("EmployeeCode", "")).strip()
        name   = str(form.get("FullName", "")).strip()
        from_d = _iso_to_date(form.get("FromDate", ""))
        to_d   = _iso_to_date(form.get("ToDate", ""))
        if from_d and to_d and not (from_d <= target_date <= to_d):
            continue
        cid = _to_check_id(tds)
        if cid:
            form_records[cid] = {"tds_code": tds, "full_name": name,
                                  "status": "Business Trip",
                                  "waive_late": False, "waive_early": False}

    for form in _fetch("LateInEarlyOut"):
        if form.get("Status") != 2: continue
        tds    = str(form.get("EmployeeCode", "")).strip()
        name   = str(form.get("FullName", "")).strip()
        from_d = _iso_to_date(form.get("FromDate", ""))
        to_d   = _iso_to_date(form.get("ToDate", ""))
        if from_d and to_d and not (from_d <= target_date <= to_d):
            continue
        wl  = int(form.get("CheckInLateStartTime",    0) or 0) > 0
        we  = int(form.get("CheckOutEarlyEndTimeint", 0) or 0) > 0
        cid = _to_check_id(tds)
        if not cid: continue
        existing = form_records.get(cid)
        if existing:
            if wl: existing["waive_late"]  = True
            if we: existing["waive_early"] = True
        else:
            lbl = "Approved Late/Early" if (wl and we) else ("Approved Late" if wl else "Approved Early")
            form_records[cid] = {"tds_code": tds, "full_name": name,
                                  "status": lbl, "waive_late": wl, "waive_early": we}

    return punch_records, form_records

# ── MERGE ─────────────────────────────────────────────────────────────────────
def merge_from_timesheet(timesheet_punch: dict, timesheet_forms: dict,
                          target_date: date) -> list:
    """
    Produce merged records from timesheet detail as the sole punch source.
    Rows are anchored on all employees returned by get-data-employee (CHECK_TO_EMP).
    """
    is_saturday = target_date.weekday() == 5

    FULL_DAY_FORMS = {
        "Business Trip", "PTO", "Unpaid Leave", "Leave",
        "Marriage Leave", "Child Marriage Leave", "Bereavement Leave",
        "Social Insurance Leave", "Maternity Leave", "Compensatory Leave",
    }

    anchor_ids = sorted(CHECK_TO_EMP.keys(),
                        key=lambda nid: CHECK_TO_EMP.get(nid, "ZZZ"))

    merged = []
    for nid in anchor_ids:
        mp   = timesheet_punch.get(nid, {})
        form = timesheet_forms.get(nid, {})

        name = (_MISA_NAME_CACHE.get(nid) or
                mp.get("full_name") or form.get("full_name") or f"ID {nid}")

        shift_work_min   = mp.get("shift_work_min", None)
        default_work_min = SAT_WORK_HOURS * 60 if is_saturday else WORK_HOURS * 60
        work_hours_min   = shift_work_min if shift_work_min is not None else default_work_min

        final_ci = mp.get("check_in",  "N/A")
        final_co = mp.get("check_out", "N/A")
        precomputed_min = mp.get("attended_min", None)

        status, late_str, early_min, late_min = compute_status(final_ci, final_co, target_date)

        form_status = form.get("status")
        if form.get("waive_late"):  late_min = 0; late_str = "0 min"
        if form.get("waive_early"): early_min = 0
        if late_min == 0 and early_min == 0 and status in ("Late", "Early", "Late/Early"):
            status = "Normal"

        if form_status and form_status in FULL_DAY_FORMS:
            status       = form_status
            attended_min = -1
        else:
            if form_status:
                status = form_status
            if precomputed_min is not None:
                attended_min = precomputed_min
            else:
                ci_t, co_t = parse_hms(final_ci), parse_hms(final_co)
                attended_min = max(0, time_to_minutes(co_t) - time_to_minutes(ci_t)) if ci_t and co_t else 0

        merged.append({
            "name":             name,
            "check_id":         nid,
            "tds_code":         CHECK_TO_EMP.get(nid, ""),
            "status":           status,
            "check_in":         final_ci,
            "check_out":        final_co,
            "late":             late_str,
            "early_leave":      f"{early_min} min" if early_min > 0 else "0 min",
            "attended_min":     attended_min,
            "work_hours_min":   work_hours_min,
        })
    return merged

# ── EXCEL EXPORT ──────────────────────────────────────────────────────────────
def export_report(records: list, report_date: date, output_path: str, language: str = "VI"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Chấm Công"

    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    default_work_hours = SAT_WORK_HOURS if report_date.weekday() == 5 else WORK_HOURS

    HEADERS_EN = {"Name": "Name", "Attendance": "Attendance", "Check-In": "Check-In",
                  "Check-Out": "Check-Out", "Late": "Late",
                  "Early Leave": "Early Leave", "Hours Worked": "Hours Worked"}
    HEADERS_VI = {"Name": "Họ tên", "Attendance": "Tình trạng", "Check-In": "Giờ vào",
                  "Check-Out": "Giờ ra", "Late": "Muộn",
                  "Early Leave": "Về sớm", "Hours Worked": "Giờ làm"}
    HEADERS_MAP  = HEADERS_VI if language == "VI" else HEADERS_EN
    shift_col_hdr = "Ca" if language == "VI" else "Shift"

    HEADERS = [
        HEADERS_MAP["Name"], HEADERS_MAP["Attendance"],
        HEADERS_MAP["Check-In"], HEADERS_MAP["Check-Out"],
        HEADERS_MAP["Late"], HEADERS_MAP["Early Leave"],
        HEADERS_MAP["Hours Worked"], shift_col_hdr,
    ]

    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        style_cell(cell, bold=True, font_colour="FFFFFF",
                   fill_colour="1F3864", border=border, size=10)
    ws.row_dimensions[1].height = 20

    for r_idx, rec in enumerate(records, 2):
        sk   = rec["status"].lower()
        fill = STATUS_COLOURS.get(sk, "FFFFFF")
        fc   = STATUS_FONT_COLOURS.get(sk, "000000")
        row_work_min = rec.get("work_hours_min", default_work_hours * 60)
        row_work_h   = round(row_work_min / 60, 2) if row_work_min % 60 else int(row_work_min // 60)
        hrs          = row_work_h if rec["attended_min"] == -1 else round(rec["attended_min"] / 60, 2)
        shift_label  = f"/ {row_work_h}h"
        display_status = STATUS_VI.get(sk, rec["status"]) if language == "VI" else rec["status"]

        vals = [
            rec["name"], display_status,
            fmt_hm(rec["check_in"]), fmt_hm(rec["check_out"]),
            rec["late"], rec["early_leave"], hrs, shift_label,
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            if c_idx == 2:
                style_cell(cell, h_align="center", font_colour=fc,
                           fill_colour=fill, border=border, bold=True)
            else:
                style_cell(cell, h_align="left" if c_idx == 1 else "center",
                           fill_colour="F5F7FA" if r_idx % 2 == 0 else "FFFFFF",
                           border=border)
        ws.row_dimensions[r_idx].height = 18

    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 40)

    ws.freeze_panes = "A2"
    wb.save(output_path)
    return output_path, len(records)

# ── EMAIL ─────────────────────────────────────────────────────────────────────
def send_email(attachment_path: str, report_date: date, employee_count: int):
    smtp_host = _env("EMAIL_SMTP_HOST")
    smtp_port = int(_env("EMAIL_SMTP_PORT"))
    username  = _env("EMAIL_USERNAME")
    password  = _env("EMAIL_PASSWORD")
    from_addr = _env("EMAIL_FROM")
    to_raw    = _env("EMAIL_TO")
    cc_raw    = _env_optional("EMAIL_CC", "")

    to_list = [a.strip() for a in to_raw.split(",") if a.strip()]
    cc_list = [a.strip() for a in cc_raw.split(",") if a.strip()]

    date_str = report_date.strftime("%d/%m/%Y")
    weekday_vi = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]
    day_name = weekday_vi[report_date.weekday()]

    subject = f"Bảng chấm công {day_name} {date_str}"
    body    = (
        f"Kính gửi,\n\n"
        f"Đính kèm bảng chấm công ngày {day_name} {date_str} "
        f"({employee_count} nhân viên).\n\n"
        f"Trân trọng."
    )

    msg = MIMEMultipart()
    msg["From"]    = from_addr
    msg["To"]      = ", ".join(to_list)
    if cc_list:
        msg["Cc"]  = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))

    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition",
                    f'attachment; filename="{os.path.basename(attachment_path)}"')
    msg.attach(part)

    all_recipients = to_list + cc_list
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.login(username, password)
        server.sendmail(from_addr, all_recipients, msg.as_string())

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    yesterday = date.today() - timedelta(days=1)
    log_path  = os.path.join(TEMP_DIR, f"run_{yesterday.isoformat()}.log")
    lines     = [f"Run at {datetime.now().isoformat()} for {yesterday}"]

    # ── Schedule check ────────────────────────────────────────────────────────
    run, reason = should_run(yesterday)
    lines.append(reason)
    if not run:
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print("\n".join(lines))
        return

    try:
        # ── Load employees ────────────────────────────────────────────────────
        ok, msg = load_misa_employees()
        lines.append(f"Employees: {msg}")
        if not ok:
            raise RuntimeError(f"Failed to load employee list: {msg}")

        if not CHECK_TO_EMP:
            raise RuntimeError("No employees found in MISA — cannot build report.")

        # ── Fetch timesheet detail ────────────────────────────────────────────
        ts_id = int(_env("MISA_TIMESHEET_ID"))
        lines.append(f"Fetching timesheet detail (ID={ts_id})…")
        resp      = fetch_timesheet_detail(ts_id)
        page_data = resp.get("Data", {}).get("PageData", [])
        if not page_data:
            raise RuntimeError("Timesheet detail returned no records.")
        lines.append(f"Timesheet records: {len(page_data)}")

        # ── Process + merge ───────────────────────────────────────────────────
        ts_punch, ts_forms = process_timesheet_detail(page_data, yesterday)
        merged = merge_from_timesheet(ts_punch, ts_forms, yesterday)
        if not merged:
            raise RuntimeError("No employee records after merge.")
        lines.append(f"Merged: {len(merged)} employees")

        # ── Export xlsx ───────────────────────────────────────────────────────
        filename = f"cham_cong_{yesterday.strftime('%d.%m.%Y')}.xlsx"
        out_path = os.path.join(SCRIPT_DIR, filename)
        export_report(merged, yesterday, out_path, language="VI")
        lines.append(f"Report saved: {filename}")

        # ── Email ─────────────────────────────────────────────────────────────
        send_email(out_path, yesterday, len(merged))
        lines.append(f"Email sent to {_env('EMAIL_TO')}")

    except Exception as ex:
        lines.append(f"FATAL: {ex}")
        print("\n".join(lines))
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        sys.exit(1)

    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("\n".join(lines))


if __name__ == "__main__":
    main()
