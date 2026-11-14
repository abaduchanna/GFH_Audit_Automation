"""
GFH Telecom LLC Inventory Audit — Audit Automation Edition (PySide6/Qt build)
==============================================================================

Same tool as the Tkinter build (GFH_Audit_Automation.py): the GFH Inventory
Status / Variance Audit workflow, the B2B Soft + Timesheet portal scrapers,
the Audit Scheduler and the WhatsApp Desktop automation - a PySide6 UI layer
on the same business logic. See the Tkinter build's module docstring for the
full feature description.

Developed by www.3SVerse.com
"""
from __future__ import annotations

import sys
import os
import csv
import datetime as dt
import hashlib
import json
import re
import shutil
import sqlite3
import struct
import subprocess
import tempfile
import threading
import time
import traceback
import zipfile
import zlib
from datetime import date
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET
import base64

# ── Required package availability ──────────────────────────────────────────
# The Tkinter build auto-installed missing packages behind a branded Tk
# splash (_auto_install_packages there). This Qt build follows the
# GFH_Inventory_Aging_Processor_Qt pattern: packages are probed at import
# time and any shortfall is reported by _check_required_packages() in
# main(), AFTER the QApplication exists so a QMessageBox can be shown.
_MISSING: List[str] = []

try:
    import openpyxl
except Exception:
    _MISSING.append("openpyxl")

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk, ImageGrab
except Exception:
    Image = None
    ImageDraw = None
    ImageFont = None
    ImageTk = None
    ImageGrab = None
    _MISSING.append("pillow")

try:
    # Presence checks only - no import. pyautogui/pyperclip/pytesseract/
    # win32api/pygetwindow are imported lazily inside the automation code
    # itself, same as the Tkinter build.
    import importlib.util as _ilu
    if _ilu.find_spec("pyautogui") is None:
        _MISSING.append("pyautogui")
    if _ilu.find_spec("pyperclip") is None:
        _MISSING.append("pyperclip")
    if _ilu.find_spec("pytesseract") is None:
        _MISSING.append("pytesseract")
    if _ilu.find_spec("win32api") is None:
        _MISSING.append("pywin32")
    if _ilu.find_spec("pygetwindow") is None:
        _MISSING.append("pygetwindow")
except Exception:
    pass

# --- 3SVerse runtime tamper guard (t197): frozen EXE main debugger = silent exit ---
import runtime_guard as _rguard3sv

_rguard3sv.enforce()

from theme_manager_qt import ThemeManagerQt


def _check_required_packages(app) -> None:
    """Call once a QApplication exists. Mirrors the Tkinter build's
    auto-install list; the Qt build shows a dialog instead of a Tk splash
    when something is missing."""
    from PySide6.QtWidgets import QMessageBox
    if _MISSING:
        QMessageBox.critical(
            None, "GFH Telecom LLC Inventory Audit - missing packages",
            "Required packages not installed:\n\n"
            + "\n".join(f"  - {p}" for p in _MISSING)
            + "\n\nInstall and reopen:\n\npip install " + " ".join(_MISSING),
        )
        sys.exit(1)


APP_NAME = "GFH Telecom LLC Inventory Audit"
def _safe_read_asset(filename: str, default: str = "") -> str:
    try:
        if getattr(sys, "frozen", False):
            base = getattr(sys, "_MEIPASS", ".")
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return open(os.path.join(base, "assets", filename), "r").read().strip()
    except Exception:
        return default

EMBEDDED_LOGO_B64 = _safe_read_asset("embedded_logo_b64.txt")
EMBEDDED_ICON_B64 = _safe_read_asset("embedded_icon_b64.txt")

if getattr(sys, "frozen", False):
    PACKAGE_DIR = Path(sys.executable).resolve().parent
else:
    PACKAGE_DIR = Path(__file__).resolve().parent
PORTABLE_APP_DIR = PACKAGE_DIR / "GFH_Inventory_Audit_Data"
LEGACY_APP_DIR = Path.home() / "GFH_Inventory_Variance_GUI"

def _choose_app_dir() -> Path:
    """Always place data folder next to the EXE (or script), regardless of location."""
    candidate = PORTABLE_APP_DIR
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        test_file = candidate / ".write_test"
        test_file.write_text("ok", encoding="utf-8")
        try:
            test_file.unlink()
        except Exception:
            pass
        return candidate
    except Exception as e:
        # Tkinter build showed a tk messagebox here; the Qt build uses a
        # QMessageBox instead (a QApplication is created on the spot if none
        # exists yet, because this runs at module import time, before main()).
        try:
            from PySide6.QtWidgets import QApplication as _QApp, QMessageBox as _QMsgBox
            _qapp = _QApp.instance() or _QApp(sys.argv)
            _QMsgBox.critical(
                None,
                "Folder not writable",
                f"Cannot write to the data folder:\n"
                f"  {candidate}\n\n"
                f"Error: {e}\n\n"
                f"Please ensure the folder is writable.",
            )
        except Exception:
            print("Folder not writable:", e)
        raise SystemExit(1)

APP_DIR = _choose_app_dir()
DB_PATH = APP_DIR / "inventory_variance_log.sqlite3"
IMAGE_DIR = APP_DIR / "whatsapp_images"
EXPORT_DIR = APP_DIR / "exports"
STORE_CONFIG_PATH = APP_DIR / "store_master_list.csv"
PACKAGED_STORE_LIST_PATH = PACKAGE_DIR / "store_master_list.csv"


def _resource_path(name: str) -> Path:
    """Resolve a bundled asset for a PyInstaller onefile EXE.

    onefile builds extract data files to sys._MEIPASS at runtime; fall back
    to the directory next to the app/script for normal runs."""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            bundled = Path(meipass) / name
            if bundled.exists():
                return bundled
    return PACKAGE_DIR / name


HEADER_LOGO_PATH = _resource_path("GFH_Telecom_Logo.png")
STATUS_LOGO_PATH = HEADER_LOGO_PATH
APP_ICON_PATH = PACKAGE_DIR / "gfh_telecom_llc_icon.ico"
APP_ICON_PNG_PATH = PACKAGE_DIR / "gfh_telecom_llc_logo.png"

# ---------------------------------------------------------------------------
# Embedded icon assets — written to PACKAGE_DIR at startup so no external
# file is needed. Swap the base64 strings to change the window icon.
# ---------------------------------------------------------------------------
APP_ICON_ICO_BASE64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app_icon_ico_base64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "app_icon_ico_base64.txt"), "r").read().strip()

APP_ICON_PNG_BASE64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app_icon_png_base64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "app_icon_png_base64.txt"), "r").read().strip()


def ensure_app_icon_files() -> None:
    """Decode and write embedded ICO + PNG icons to PACKAGE_DIR once."""
    try:
        ico_bytes = base64.b64decode(APP_ICON_ICO_BASE64.strip().replace("\n", ""))
        if not APP_ICON_PATH.exists() or APP_ICON_PATH.stat().st_size != len(ico_bytes):
            APP_ICON_PATH.write_bytes(ico_bytes)
    except Exception:
        pass
    try:
        png_bytes = base64.b64decode(APP_ICON_PNG_BASE64.strip().replace("\n", ""))
        if not APP_ICON_PNG_PATH.exists() or APP_ICON_PNG_PATH.stat().st_size != len(png_bytes):
            APP_ICON_PNG_PATH.write_bytes(png_bytes)
    except Exception:
        pass



SEARCH_SHORTCUT = "ctrl+f"
SEND_ONLY_UNSENT_BY_DEFAULT = True
AUTO_SEND_CHUNK_SIZE = 22

EXCEL_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

SEND_MODE_LABELS = {
    "district": "District",
    "store": "Store",
    "rep": "Sales Rep",
}



DEFAULT_STORE_MASTER = []

def default_store_master_records() -> List[Dict[str, str]]:
    return [{"District": district, "Store": store} for district, store in DEFAULT_STORE_MASTER]


def migrate_legacy_app_data_if_needed() -> None:
    """No-op: legacy per-user folder migration removed.
    All instances now share the single portable DB next to the script."""
    pass


def save_store_master_records(records: List[Dict[str, str]], path: Path = STORE_CONFIG_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    normalized_rows: Dict[str, Dict[str, str]] = {}
    for rec in records:
        district = normalize_district(rec.get("District", ""))
        store = display_store(rec.get("Store", ""))
        norm = normalize_store(store)
        if not norm or not district or district == "Unknown":
            continue
        normalized_rows[norm] = {"District": district, "Store": store}
    ordered = sorted(normalized_rows.values(), key=lambda r: (normalize_district(r["District"]), r["Store"].lower()))
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["District", "Store"])
        writer.writeheader()
        writer.writerows(ordered)


def ensure_store_master_file() -> None:
    if STORE_CONFIG_PATH.exists():
        return
    STORE_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    if PACKAGED_STORE_LIST_PATH.exists():
        shutil.copy2(PACKAGED_STORE_LIST_PATH, STORE_CONFIG_PATH)
    else:
        save_store_master_records(default_store_master_records(), STORE_CONFIG_PATH)


def load_store_master_records(path: Path = STORE_CONFIG_PATH) -> List[Dict[str, str]]:
    ensure_store_master_file()
    rows: List[Dict[str, str]] = []
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for rec in reader:
                district = normalize_district(rec.get("District", ""))
                store = display_store(rec.get("Store", ""))
                if district and district != "Unknown" and store:
                    rows.append({"District": district, "Store": store})
    except Exception:
        rows = default_store_master_records()
        save_store_master_records(rows, path)
    if not rows:
        rows = default_store_master_records()
        save_store_master_records(rows, path)
    return rows


@dataclass
class VarianceRow:
    key: str
    district: str
    store: str
    product: str
    imei: str
    status: str
    created_by: str = ""
    rep_name: str = ""
    created_date: str = ""
    document_status: str = ""
    source_file: str = ""
    cleared: bool = False
    sent_count: int = 0
    last_sent_at: str = ""
    cleared_at: str = ""
    notes: str = ""


@dataclass
class InventoryStatusRow:
    key: str
    district: str
    store: str
    status: str
    rep_name: str = ""
    source_file: str = ""

def now_text() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def normalize_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", safe_text(text).lower())


def normalize_store(text: str) -> str:
    value = safe_text(text)
    value = re.sub(r"\s+", " ", value).strip()
    value = value.replace("–", "-").replace("—", "-")
    return value.lower()


def display_store(text: str) -> str:
    value = safe_text(text)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalize_phone(text: str) -> str:
    value = safe_text(text)
    if not value:
        return ""
    value = value.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    value = value.replace(".", "")
    if value.startswith("00"):
        value = "+" + value[2:]
    return value


def whatsapp_mention(phone: str) -> str:
    phone = normalize_phone(phone)
    if not phone:
        return ""
    if phone.startswith("@"):
        return phone
    return "@" + phone


def person_name_key(text: str) -> str:
    value = safe_text(text).lower()
    value = value.replace(",", " ")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    parts = [p for p in value.split() if p]
    return " ".join(sorted(parts))


def device_rule_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", safe_text(text).lower())


def device_matches_rule(product: str, imei: str, rule_text: str, match_type: str) -> bool:
    rule_clean = device_rule_key(rule_text)
    if not rule_clean:
        return False

    product_clean = device_rule_key(product)
    imei_clean = device_rule_key(imei)
    match_type_clean = normalize_header(match_type)

    if match_type_clean in {"productexact", "exactproduct"}:
        return product_clean == rule_clean
    if match_type_clean in {"imeiexact", "serialexact", "esnexact"}:
        return imei_clean == rule_clean
    if match_type_clean in {"anycontains", "containsany"}:
        return rule_clean in product_clean or rule_clean in imei_clean
    return rule_clean in product_clean


def normalize_district(text: str) -> str:
    value = safe_text(text).replace("–", "-").replace("—", "-")
    value = re.sub(r"\s+", " ", value).strip()
    if not value:
        return "Unknown"
    lower = value.lower()
    if lower.startswith("arizona"):
        return "Arizona"
    aliases = {
        "atlanta": "Atlanta",
        "colorado west": "Colorado West",
        "houston": "Houston",
        "colorado east": "Colorado East",
        "tennessee": "Tennessee",
        "louisiana": "Louisiana",
    }
    return aliases.get(lower, value)


def group_name_for_district(district: str, db=None) -> str:
    normalized = normalize_district(district)
    if db is not None:
        try:
            saved = db.find_whatsapp_group(normalized)
            if saved:
                return saved
        except Exception:
            pass
    return f"GFH TELECOM {normalized.upper()}"


def is_sim_product(product: str) -> bool:
    text = safe_text(product).lower()
    compact = normalize_header(product)
    if not text and not compact:
        return False
    if re.search(r"(^|[^a-z0-9])e?[-\s]?sim(s)?([^a-z0-9]|$)", text):
        return True
    if "simcard" in compact or compact in {"sim", "sims", "esim", "esims"}:
        return True
    if compact.startswith(("sim", "esim")) and any(token in compact for token in ("card", "kit", "pack", "starter")):
        return True
    return False


def excel_serial_to_datetime(value) -> Optional[dt.datetime]:
    text = safe_text(value)
    if not text:
        return None
    try:
        serial = float(text)
    except ValueError:
        # Try common date strings if present.
        for fmt in ("%m/%d/%Y", "%m/%d/%Y %I:%M %p", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        return None
    base = dt.datetime(1899, 12, 30)
    return base + dt.timedelta(days=serial)


def excel_serial_to_date_text(value) -> str:
    parsed = excel_serial_to_datetime(value)
    if parsed is None:
        return safe_text(value)
    return parsed.strftime("%m/%d/%Y %I:%M %p")


def numeric_excel_date(value) -> float:
    text = safe_text(value)
    try:
        return float(text)
    except ValueError:
        parsed = excel_serial_to_datetime(text)
        if parsed is None:
            return -1.0
        base = dt.datetime(1899, 12, 30)
        return (parsed - base).total_seconds() / 86400.0


def variance_key(store: str, imei: str, product: str, status: str, created_by: str = "", created_date: str = "") -> str:
    raw = "|".join([
        normalize_store(store),
        safe_text(imei).lower(),
        safe_text(product).lower(),
        safe_text(status).lower(),
        safe_text(created_by).lower(),
        safe_text(created_date).lower(),
    ])
    return hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()



# ── win32api keyboard fallback (for Python 3.14 where pynput breaks) ────────
_WA_VK_MAP = {
    "ctrl": 0x11, "control": 0x11, "alt": 0x12, "menu": 0x12,
    "shift": 0x10, "win": 0x5B, "enter": 0x0D, "return": 0x0D,
    "tab": 0x09, "esc": 0x1B, "escape": 0x1B, "backspace": 0x08,
    "delete": 0x2E, "del": 0x2E, "home": 0x24, "end": 0x23,
    "pageup": 0x21, "pagedown": 0x22, "up": 0x26, "down": 0x28,
    "left": 0x25, "right": 0x27, "space": 0x20, "a": 0x41, "c": 0x43,
    "f": 0x46, "v": 0x56, "x": 0x58,
}

def _wa_win32_key_down(vk_code):
    import ctypes
    ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0)

def _wa_win32_key_up(vk_code):
    import ctypes
    ctypes.windll.user32.keybd_event(vk_code, 0, 0x0002, 0)

def _wa_win32_hotkey(*keys):
    codes = []
    for k in keys:
        k_lower = k.lower()
        if k_lower in _WA_VK_MAP:
            codes.append(_WA_VK_MAP[k_lower])
        elif len(k) == 1:
            codes.append(ord(k.upper()))
        else:
            return
    for code in codes:
        _wa_win32_key_down(code)
    for code in reversed(codes):
        _wa_win32_key_up(code)

def _wa_win32_press(key):
    k_lower = key.lower()
    if k_lower in _WA_VK_MAP:
        code = _WA_VK_MAP[k_lower]
    elif len(key) == 1:
        code = ord(key.upper())
    else:
        return
    _wa_win32_key_down(code)
    _wa_win32_key_up(code)

def _wa_win32_write(text, interval=0.0):
    import ctypes
    for char in text:
        vk = ctypes.windll.user32.VkKeyScanW(ord(char))
        if vk == -1:
            continue
        code = vk & 0xFF
        shift = (vk >> 8) & 1
        if shift:
            _wa_win32_key_down(_WA_VK_MAP["shift"])
        _wa_win32_key_down(code)
        _wa_win32_key_up(code)
        if shift:
            _wa_win32_key_up(_WA_VK_MAP["shift"])
        if interval > 0:
            import time as _t
            _t.sleep(interval)

_wa_hotkey = _wa_win32_hotkey
_wa_press = _wa_win32_press
_wa_write = _wa_win32_write


class RobustXlsxReader:
    """Small XLSX reader with a fallback for exports missing a central ZIP directory."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.parts = self._read_parts()
        self.shared_strings = self._read_shared_strings()

    def _read_parts(self) -> Dict[str, bytes]:
        parts: Dict[str, bytes] = {}
        try:
            with zipfile.ZipFile(self.path) as zf:
                for name in zf.namelist():
                    parts[name] = zf.read(name)
            return parts
        except Exception:
            data = self.path.read_bytes()
            pos = 0
            while pos + 30 <= len(data) and data[pos:pos + 4] == b"PK\x03\x04":
                try:
                    (
                        _sig,
                        _ver,
                        _flag,
                        method,
                        _mtime,
                        _mdate,
                        _crc,
                        csize,
                        _usize,
                        nlen,
                        xlen,
                    ) = struct.unpack_from("<IHHHHHIIIHH", data, pos)
                    name_start = pos + 30
                    name_end = name_start + nlen
                    name = data[name_start:name_end].decode("utf-8", errors="replace")
                    start = name_end + xlen
                    comp = data[start:min(start + csize, len(data))]
                    if method == 8:
                        try:
                            content = zlib.decompress(comp, -15)
                        except Exception:
                            content = b""
                    elif method == 0:
                        content = comp
                    else:
                        content = b""
                    if content:
                        parts[name] = content
                    pos = start + csize
                except Exception:
                    break
            if not parts:
                raise RuntimeError(f"Could not read Excel file: {self.path}")
            return parts

    @staticmethod
    def _clean_xml(raw: bytes) -> bytes:
        return raw.lstrip(b"\xef\xbb\xbf")

    def _read_shared_strings(self) -> List[str]:
        raw = self.parts.get("xl/sharedStrings.xml")
        if not raw:
            return []
        root = ET.fromstring(self._clean_xml(raw))
        strings: List[str] = []
        for si in root.findall(EXCEL_NS + "si"):
            strings.append("".join((t.text or "") for t in si.iter(EXCEL_NS + "t")))
        return strings

    def _workbook_rels(self) -> Dict[str, str]:
        raw = self.parts.get("xl/_rels/workbook.xml.rels")
        if not raw:
            return {}
        root = ET.fromstring(self._clean_xml(raw))
        rels: Dict[str, str] = {}
        for rel in root.findall(REL_NS + "Relationship"):
            rid = rel.attrib.get("Id", "")
            target = rel.attrib.get("Target", "")
            if not target:
                continue
            if not target.startswith("/"):
                target = "xl/" + target
            else:
                target = target.lstrip("/")
            rels[rid] = target
        return rels

    def sheet_paths(self) -> List[Tuple[str, str]]:
        raw = self.parts.get("xl/workbook.xml")
        if not raw:
            if "xl/worksheets/sheet1.xml" in self.parts:
                return [("Sheet1", "xl/worksheets/sheet1.xml")]
            return []
        root = ET.fromstring(self._clean_xml(raw))
        rels = self._workbook_rels()
        sheets: List[Tuple[str, str]] = []
        for sheet in root.findall(".//" + EXCEL_NS + "sheet"):
            name = sheet.attrib.get("name", "Sheet")
            rid = sheet.attrib.get(R_NS + "id", "")
            path = rels.get(rid)
            if path and path in self.parts:
                sheets.append((name, path))
        if not sheets and "xl/worksheets/sheet1.xml" in self.parts:
            sheets.append(("Sheet1", "xl/worksheets/sheet1.xml"))
        return sheets

    @staticmethod
    def _cell_to_col(ref: str) -> Optional[int]:
        match = re.match(r"([A-Z]+)(\d+)", ref or "")
        if not match:
            return None
        col = 0
        for ch in match.group(1):
            col = col * 26 + ord(ch) - 64
        return col

    def _cell_value(self, cell: ET.Element) -> str:
        cell_type = cell.attrib.get("t")
        value_node = cell.find(EXCEL_NS + "v")
        if cell_type == "s":
            if value_node is None or value_node.text is None:
                return ""
            try:
                idx = int(value_node.text)
            except ValueError:
                return ""
            if 0 <= idx < len(self.shared_strings):
                return self.shared_strings[idx]
            return ""
        if cell_type == "inlineStr":
            inline = cell.find(EXCEL_NS + "is")
            if inline is None:
                return ""
            return "".join((t.text or "") for t in inline.iter(EXCEL_NS + "t"))
        if value_node is None or value_node.text is None:
            return ""
        return value_node.text

    def read_sheet(self, preferred_name: Optional[str] = None) -> List[Dict[str, str]]:
        sheets = self.sheet_paths()
        if not sheets:
            raise RuntimeError(f"No worksheet found in {self.path.name}")
        selected_name, selected_path = sheets[0]
        if preferred_name:
            want = preferred_name.lower().strip()
            for name, path in sheets:
                if want in name.lower().strip():
                    selected_name, selected_path = name, path
                    break
        raw = self.parts[selected_path]
        root = ET.fromstring(self._clean_xml(raw))
        rows: List[Tuple[int, Dict[int, str]]] = []
        for row in root.findall(".//" + EXCEL_NS + "row"):
            row_num = int(row.attrib.get("r", "0") or 0)
            values: Dict[int, str] = {}
            for cell in row.findall(EXCEL_NS + "c"):
                col = self._cell_to_col(cell.attrib.get("r", ""))
                if not col:
                    continue
                values[col] = self._cell_value(cell)
            if values:
                rows.append((row_num, values))
        if not rows:
            return []

        header_row_idx, header_values = rows[0]
        max_col = max(header_values)
        headers = [safe_text(header_values.get(i, "")) for i in range(1, max_col + 1)]
        seen: Dict[str, int] = {}
        clean_headers = []
        for i, header in enumerate(headers, start=1):
            name = header if header else f"Column{i}"
            base = name
            if base in seen:
                seen[base] += 1
                name = f"{base}_{seen[base]}"
            else:
                seen[base] = 1
            clean_headers.append(name)

        records: List[Dict[str, str]] = []
        for row_num, value_map in rows[1:]:
            if row_num <= header_row_idx:
                continue
            record: Dict[str, str] = {}
            max_data_col = max(max_col, max(value_map.keys()) if value_map else max_col)
            for i in range(1, max_data_col + 1):
                header = clean_headers[i - 1] if i <= len(clean_headers) else f"Column{i}"
                record[header] = safe_text(value_map.get(i, ""))
            if any(safe_text(v) for v in record.values()):
                records.append(record)
        return records


def find_column(record: Dict[str, str], candidates: Iterable[str]) -> Optional[str]:
    normalized = {normalize_header(k): k for k in record.keys()}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized:
            return normalized[key]
    for candidate in candidates:
        key = normalize_header(candidate)
        for norm, original in normalized.items():
            if key and (key in norm or norm in key):
                return original
    return None


def read_xlsx_records(path: str | Path) -> List[Dict[str, str]]:
    reader = RobustXlsxReader(path)
    return reader.read_sheet()


def build_ts_store_to_employee(time_sheet_records: List[Dict[str, str]]) -> Dict[str, str]:
    """Map normalized store → employee name from Timesheet rows.

    Keeps rows without a Clock In too — the employee may still be at the
    store with the count pending, and the pending-store rep lookup is
    built from these rows. When a store has several rows, the latest
    Clock In wins (falling back to row order when Clock In is missing).
    """
    ts_store_to_employee: Dict[str, str] = {}
    if not time_sheet_records:
        return ts_store_to_employee
    sample_ts    = time_sheet_records[0]
    ts_store_col = find_column(sample_ts, ["Store"])
    ts_emp_col   = find_column(sample_ts, ["Employee", "Employee Name",
                                           "Salesperson", "Sales Person", "Rep Name"])
    ts_clock_col = find_column(sample_ts, ["Clock In", "Clock-In", "Clock In "])
    latest_ts: Dict[str, float] = {}
    for idx, rec in enumerate(time_sheet_records):
        emp_raw = safe_text(rec.get(ts_emp_col, "")) if ts_emp_col else ""
        # Skip TOTAL/subtotal rows
        if not emp_raw or "TOTAL" in emp_raw.upper() or "\u2014 " in emp_raw or "-- " in emp_raw:
            continue
        store_raw = rec.get(ts_store_col, "") if ts_store_col else ""
        norm = normalize_store(store_raw)
        if not norm:
            continue
        clock_in_val = safe_text(rec.get(ts_clock_col, "")) if ts_clock_col else ""
        score = numeric_excel_date(clock_in_val) if clock_in_val.strip() else float(idx)
        if score >= latest_ts.get(norm, -1.0):
            ts_store_to_employee[norm] = emp_raw.strip()
            latest_ts[norm] = score
    return ts_store_to_employee


def store_name_tokens(norm_store: str) -> Tuple[str, ...]:
    """Alphanumeric tokens of an already-normalized store name, deduped."""
    return tuple(dict.fromkeys(t for t in re.findall(r"[a-z0-9]+", norm_store or "") if t))


def match_store_employee(norm_store: str, ts_store_to_employee: Dict[str, str]) -> str:
    """Resolve the timesheet employee for a store.

    Exact normalized-store lookup first, then a tolerant token-overlap
    match. The timesheet export and the count file often spell the same
    store slightly differently ("Kings Highway #1204" vs "Kings Highway
    Store 1204"), which left pending stores without an employee name even
    when the timesheet had a clock-in for them. A token match is accepted
    when one name's tokens are a subset of the other's — but never when a
    distinguishing store number conflicts ("Store 1204" vs "Store 1205"
    do not match). Ties go to the candidate with the most token overlap,
    then the shortest name.
    """
    if not norm_store or not ts_store_to_employee:
        return ""
    exact = ts_store_to_employee.get(norm_store)
    if exact:
        return exact
    target_tokens = set(store_name_tokens(norm_store))
    if not target_tokens:
        return ""
    best_name = ""
    best_overlap = 0
    best_len = -1
    for ts_norm, employee in ts_store_to_employee.items():
        if not ts_norm or not employee:
            continue
        ts_tokens = set(store_name_tokens(ts_norm))
        if not ts_tokens:
            continue
        if target_tokens <= ts_tokens or ts_tokens <= target_tokens:
            overlap = len(target_tokens & ts_tokens)
            if overlap > best_overlap or (overlap == best_overlap and best_len != -1 and len(ts_norm) < best_len):
                best_name = employee
                best_overlap = overlap
                best_len = len(ts_norm)
            elif best_len == -1:
                best_name = employee
                best_overlap = overlap
                best_len = len(ts_norm)
    return best_name


def build_store_maps(
    inventory_records: List[Dict[str, str]],
    time_sheet_records: List[Dict[str, str]],
) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, str]]:
    """Build store lookups using Count Details as the source of truth.

    - Store names and Districts come from Inventory_Count_Result_Details.
    - Employee (rep) names come from the Timesheet by matching on Store,
      tolerating slight store-name differences between the two exports.
    - Timesheet is ONLY used to resolve who worked at each store.
    """
    district_by_store: Dict[str, str] = {}
    display_by_store:  Dict[str, str] = {}
    rep_by_store:      Dict[str, str] = {}   # norm_store → full Employee name

    # ── Step 1: Build store → employee map from Timesheet ─────────────────
    ts_store_to_employee = build_ts_store_to_employee(time_sheet_records)

    # ── Step 2: Build district/display maps from Count Details ────────────
    if not inventory_records:
        return district_by_store, display_by_store, rep_by_store

    sample_inv   = inventory_records[0]
    inv_store_col    = find_column(sample_inv, ["Store"])
    inv_district_col = find_column(sample_inv, ["District"])

    if not inv_store_col:
        return district_by_store, display_by_store, rep_by_store

    for rec in inventory_records:
        store_raw = rec.get(inv_store_col, "")
        norm = normalize_store(store_raw)
        if not norm:
            continue

        # Store display name from count file
        display_by_store[norm] = display_store(store_raw)

        # District from count file
        if inv_district_col:
            district = normalize_district(rec.get(inv_district_col, ""))
            if district and district != "Unknown":
                district_by_store[norm] = district

        # Employee name from timesheet (matched by store, tolerant of
        # slight name differences between the two exports)
        if norm not in rep_by_store:
            employee = ts_store_to_employee.get(norm) or match_store_employee(norm, ts_store_to_employee)
            if employee:
                rep_by_store[norm] = employee

    return district_by_store, display_by_store, rep_by_store


def build_rep_by_store_from_created_by(
    inventory_records: List[Dict[str, str]],
) -> Dict[str, str]:
    """Latest "Created By" value per store from the Inventory_Count_Result_Details rows.

    Replaces the old Employee_Time_Sheet-based rep lookup. Employee/rep name for a store is
    simply whoever most recently counted that store, per the Created Date/Time column.
    """
    rep_by_store: Dict[str, str] = {}
    latest_score_by_store: Dict[str, float] = {}
    if not inventory_records:
        return rep_by_store

    sample = inventory_records[0]
    store_col = find_column(sample, ["Store"])
    created_by_col = find_column(sample, ["Created By", "Count By", "User Login"])
    created_date_col = find_column(sample, ["Created Date", "Created Date/Time", "Date Time", "Date"])
    if not store_col or not created_by_col:
        return rep_by_store

    for index, rec in enumerate(inventory_records):
        norm = normalize_store(rec.get(store_col, ""))
        if not norm:
            continue
        created_by = safe_text(rec.get(created_by_col, ""))
        if not created_by:
            continue
        score = numeric_excel_date(rec.get(created_date_col, "")) if created_date_col else float(index)
        if score < 0:
            score = float(index) / 1000000.0
        if score >= latest_score_by_store.get(norm, -1.0):
            rep_by_store[norm] = created_by
            latest_score_by_store[norm] = score

    return rep_by_store


def filter_latest_inventory_records(inventory_records: List[Dict[str, str]]) -> Tuple[List[Dict[str, str]], Dict[str, int]]:
    """Keep only the latest Inventory_Count_Result_Details rows per Store + Created By.

    GFH exports can include more than one count from the same Created By user for the same store.
    Older rows should not drive the current variance UI because a later recount can convert earlier
    Deficit rows into Matched rows. The Created Date value includes date and time as an Excel serial.
    """
    metrics = {
        "raw_inventory_rows": len(inventory_records),
        "latest_inventory_rows": len(inventory_records),
        "stale_inventory_rows": 0,
        "latest_created_by_groups": 0,
    }
    if not inventory_records:
        return inventory_records, metrics

    sample = inventory_records[0]
    store_col = find_column(sample, ["Store"])
    created_by_col = find_column(sample, ["Created By", "Count By", "User Login"])
    created_date_col = find_column(sample, ["Created Date", "Created Date/Time", "Date Time", "Date"])

    if not store_col or not created_by_col or not created_date_col:
        return inventory_records, metrics

    latest_by_group: Dict[Tuple[str, str], float] = {}
    scores: List[Tuple[Tuple[str, str], float]] = []

    for index, rec in enumerate(inventory_records):
        store_key = normalize_store(rec.get(store_col, "")) or "__unknown_store__"
        created_by_key = safe_text(rec.get(created_by_col, "")).lower().strip() or "__unknown_created_by__"
        group_key = (store_key, created_by_key)
        score = numeric_excel_date(rec.get(created_date_col, ""))
        if score < 0:
            score = float(index) / 1000000.0
        scores.append((group_key, score))
        if group_key not in latest_by_group or score > latest_by_group[group_key]:
            latest_by_group[group_key] = score

    filtered: List[Dict[str, str]] = []
    for rec, (group_key, score) in zip(inventory_records, scores):
        latest_score = latest_by_group.get(group_key, score)
        if abs(score - latest_score) <= 0.0000001:
            filtered.append(rec)

    metrics["latest_inventory_rows"] = len(filtered)
    metrics["stale_inventory_rows"] = len(inventory_records) - len(filtered)
    metrics["latest_created_by_groups"] = len(latest_by_group)
    return filtered, metrics


def extract_variances(
    inventory_records: List[Dict[str, str]],
    time_sheet_records: List[Dict[str, str]],
    master_store_records: Optional[List[Dict[str, str]]] = None,
    source_file: str = "",
) -> Tuple[List[VarianceRow], Dict[str, int]]:
    if not inventory_records:
        return [], {"completed": 0, "pending": 0, "stores_total": 0, "skipped_sims": 0, "raw_inventory_rows": 0, "latest_inventory_rows": 0, "stale_inventory_rows": 0, "latest_created_by_groups": 0}

    # Employee (rep) names come from timesheet matched by store.
    # Store names and Districts come from count details (source of truth).
    _, _, rep_by_store = build_store_maps(inventory_records, time_sheet_records)

    # Build district/display maps from count details directly
    district_by_store: Dict[str, str] = {}
    display_by_store: Dict[str, str] = {}
    sample0 = inventory_records[0]
    _store_col0    = find_column(sample0, ["Store"])
    _district_col0 = find_column(sample0, ["District"])
    for rec in inventory_records:
        store_raw = rec.get(_store_col0, "") if _store_col0 else ""
        norm = normalize_store(store_raw)
        if not norm:
            continue
        display_by_store[norm] = display_store(store_raw)
        if _district_col0:
            d = normalize_district(rec.get(_district_col0, ""))
            if d and d != "Unknown":
                district_by_store[norm] = d

    # Fill gaps from master store list
    for rec in master_store_records or []:
        district = normalize_district(rec.get("District", ""))
        store = display_store(rec.get("Store", ""))
        norm = normalize_store(store)
        if not norm:
            continue
        if district and district != "Unknown" and norm not in district_by_store:
            district_by_store[norm] = district
        if store and norm not in display_by_store:
            display_by_store[norm] = store
    sample = inventory_records[0]
    store_col = find_column(sample, ["Store"])
    product_col = find_column(sample, ["Product Description", "Product"])
    imei_col = find_column(sample, ["Serial #", "Serial", "IMEI", "ESN"])
    status_col = find_column(sample, ["Status"])
    created_by_col = find_column(sample, ["Created By", "Count By", "User Login"])
    created_date_col = find_column(sample, ["Created Date", "Date"])
    document_status_col = find_column(sample, ["Document Status"])

    missing = [
        name for name, col in [
            ("Store", store_col),
            ("Product Description", product_col),
            ("Serial # / IMEI", imei_col),
            ("Status", status_col),
        ] if not col
    ]
    if missing:
        raise RuntimeError("Missing required inventory column(s): " + ", ".join(missing))

    inventory_records, latest_metrics = filter_latest_inventory_records(inventory_records)

    # Deduplicate by (Store, IMEI) to prevent duplicate IMEI rows
    dedup_map: Dict[Tuple[str, str], Tuple[Dict[str, str], float]] = {}
    for rec in inventory_records:
        store_raw = rec.get(store_col, "")
        imei = safe_text(rec.get(imei_col, ""))
        if not imei:
            continue
        norm_store = normalize_store(store_raw)
        key = (norm_store, imei.lower())
        date_score = numeric_excel_date(rec.get(created_date_col, "")) if created_date_col else 0
        if key not in dedup_map or date_score > dedup_map[key][1]:
            dedup_map[key] = (rec, date_score)
    inventory_records = [rec for rec, _ in dedup_map.values()]
    latest_metrics["dedup_inventory_rows"] = len(inventory_records)

    completed_store_norms = set()
    for rec in inventory_records:
        norm = normalize_store(rec.get(store_col, ""))
        if norm:
            completed_store_norms.add(norm)

    all_store_norms = set(display_by_store.keys()) | completed_store_norms
    pending_store_norms = all_store_norms - completed_store_norms

    variance_rows: List[VarianceRow] = []
    skipped_sims = 0

    for rec in inventory_records:
        status = safe_text(rec.get(status_col, ""))
        if not status:
            continue
        if normalize_header(status) in {"matched", "match", "ok", "balanced"}:
            continue

        store = display_store(rec.get(store_col, ""))
        product = safe_text(rec.get(product_col, ""))
        imei = safe_text(rec.get(imei_col, ""))

        if not imei:
            continue
        if product and normalize_header(product) in {"accessorycommission", "accessoriescommission"}:
            continue
        if is_sim_product(product):
            skipped_sims += 1
            continue

        norm_store = normalize_store(store)
        district = normalize_district(district_by_store.get(norm_store, "Unknown"))
        created_by = safe_text(rec.get(created_by_col, "")) if created_by_col else ""
        # Rep/employee name now comes straight from the "Created By" column of the
        # inventory count Excel — no more separate Employee_Time_Sheet lookup.
        rep_name = created_by
        created_date = excel_serial_to_date_text(rec.get(created_date_col, "")) if created_date_col else ""
        document_status = safe_text(rec.get(document_status_col, "")) if document_status_col else ""
        key = variance_key(store, imei, product, status, created_by, created_date)

        variance_rows.append(
            VarianceRow(
                key=key,
                district=district,
                store=store,
                product=product,
                imei=imei,
                status=status,
                created_by=created_by,
                rep_name=rep_name,
                created_date=created_date,
                document_status=document_status,
                source_file=source_file,
            )
        )

    summary = {
        "completed": len(completed_store_norms),
        "pending": len(pending_store_norms),
        "stores_total": len(all_store_norms),
        "skipped_sims": skipped_sims,
        **latest_metrics,
    }
    return variance_rows, summary




def build_inventory_status_rows(
    inventory_records: List[Dict[str, str]],
    time_sheet_records: List[Dict[str, str]],
    master_store_records: Optional[List[Dict[str, str]]] = None,
    source_file: str = "",
) -> Tuple[List[InventoryStatusRow], Dict[str, int]]:
    # build_store_maps gives us employee (rep) names matched by store.
    # It does NOT drive which stores appear — count details is the source of truth.
    _, _, rep_by_store = build_store_maps(inventory_records, time_sheet_records)
    # Timesheet store → employee map used as a fallback for pending stores
    # that never appear in the count file (master-list-only stores): those
    # stores are missing from rep_by_store because build_store_maps only
    # transfers employees for count-file stores, yet the timesheet clearly
    # shows who clocked in there.
    ts_store_to_employee = build_ts_store_to_employee(time_sheet_records)
    # Timesheet district fallback — third layer after the count file and the
    # Store List (master) so a store known only to the timesheet still gets
    # its district instead of "Unknown".
    ts_district_by_store: Dict[str, str] = {}
    if time_sheet_records:
        ts_sample = time_sheet_records[0]
        ts_store_col = find_column(ts_sample, ["Store"])
        ts_district_col = find_column(ts_sample, ["District"])
        if ts_store_col and ts_district_col:
            for rec in time_sheet_records:
                norm = normalize_store(rec.get(ts_store_col, ""))
                if not norm:
                    continue
                d = normalize_district(rec.get(ts_district_col, ""))
                if d and d != "Unknown":
                    ts_district_by_store.setdefault(norm, d)

    # Build store display names AND districts directly from count details so
    # every store in the count file appears regardless of timesheet coverage.
    count_display_by_store: Dict[str, str] = {}
    count_district_by_store: Dict[str, str] = {}
    completed_store_norms: set[str] = set()
    inv_display_by_store: Dict[str, str] = {}

    latest_metrics = {
        "raw_inventory_rows": len(inventory_records),
        "latest_inventory_rows": len(inventory_records),
        "stale_inventory_rows": 0,
        "latest_created_by_groups": 0,
    }

    if inventory_records:
        sample = inventory_records[0]
        store_col    = find_column(sample, ["Store"])
        district_col = find_column(sample, ["District"])

        # First pass: collect ALL stores+districts from the count file
        for rec in inventory_records:
            store_raw = rec.get(store_col, "") if store_col else ""
            norm = normalize_store(store_raw)
            if not norm:
                continue
            count_display_by_store[norm] = display_store(store_raw)
            if district_col:
                d = normalize_district(rec.get(district_col, ""))
                if d and d != "Unknown":
                    count_district_by_store[norm] = d

        # Second pass: find completed stores (those that have records in latest batch)
        if store_col:
            filtered_records, latest_metrics = filter_latest_inventory_records(inventory_records)
            for rec in filtered_records:
                store_raw = rec.get(store_col, "")
                norm = normalize_store(store_raw)
                if norm:
                    inv_display_by_store[norm] = display_store(store_raw)
                    completed_store_norms.add(norm)

    # Fallback: also include master store list for any stores not in count file
    master_display_by_store: Dict[str, str] = {}
    master_district_by_store: Dict[str, str] = {}
    for rec in master_store_records or []:
        store = display_store(rec.get("Store", ""))
        norm = normalize_store(store)
        if not norm:
            continue
        master_display_by_store[norm] = store
        d = normalize_district(rec.get("District", ""))
        if d and d != "Unknown":
            master_district_by_store[norm] = d

    # All stores = everything from count file (primary) + master list (fallback)
    all_store_norms = sorted(
        set(count_display_by_store.keys())
        | set(master_display_by_store.keys())
    )

    rows: List[InventoryStatusRow] = []
    for norm in all_store_norms:
        # Display name: count file first, then master list
        store = (count_display_by_store.get(norm)
                 or master_display_by_store.get(norm)
                 or norm.title())
        # District: count file first, then master list
        district = normalize_district(
            count_district_by_store.get(norm)
            or master_district_by_store.get(norm)
            or ts_district_by_store.get(norm)
            or "Unknown"
        )
        rep_name = safe_text(rep_by_store.get(norm, ""))
        if not rep_name:
            # Pending stores that never appear in the count file still get
            # their timesheet employee here (exact match, then tolerant
            # token-overlap match on the store name).
            rep_name = safe_text(match_store_employee(norm, ts_store_to_employee))
        status = "Completed" if norm in completed_store_norms else "Pending"
        key_raw = f"{norm}|{district}|{status}|status"
        key = hashlib.sha1(key_raw.encode("utf-8", errors="ignore")).hexdigest()
        rows.append(InventoryStatusRow(
            key=key, district=district, store=store,
            status=status, rep_name=rep_name, source_file=source_file
        ))

    summary = {
        "completed": len(completed_store_norms),
        "pending": len([r for r in rows if r.status == "Pending"]),
        "stores_total": len(rows),
        **latest_metrics,
    }
    return rows, summary


class VarianceDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_wal_mode()
        self._backup_on_startup()
        self._init_db()

    def _ensure_wal_mode(self) -> None:
        try:
            con = sqlite3.connect(self.db_path, timeout=10, isolation_level=None)
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA synchronous=NORMAL")
            con.close()
        except Exception:
            pass

    def _backup_on_startup(self) -> None:
        backup_path = self.db_path.with_suffix(".bak.sqlite3")
        try:
            if self.db_path.exists():
                con = sqlite3.connect(self.db_path, timeout=10)
                result = con.execute("PRAGMA integrity_check").fetchone()
                if result and result[0] == "ok":
                    import shutil
                    shutil.copy2(self.db_path, backup_path)
                else:
                    if backup_path.exists():
                        import shutil
                        shutil.copy2(backup_path, self.db_path)
                con.close()
        except Exception:
            pass

    def connect(self):
        con = sqlite3.connect(self.db_path, timeout=30,
                              isolation_level=None,
                              check_same_thread=False)
        try:
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA synchronous=NORMAL")
            con.execute("PRAGMA busy_timeout=30000")
            con.execute("PRAGMA cache_size=-8000")
            con.execute("PRAGMA temp_store=MEMORY")
            con.execute("PRAGMA mmap_size=134217728")
        except Exception:
            pass
        con.isolation_level = ""
        return con

    @staticmethod
    def _table_columns(con: sqlite3.Connection, table: str) -> set[str]:
        try:
            return {row[1] for row in con.execute(f"PRAGMA table_info({table})").fetchall()}
        except Exception:
            return set()

    def _init_db(self) -> None:
        with self.connect() as con:
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS variances (
                    key TEXT PRIMARY KEY,
                    district TEXT,
                    store TEXT,
                    product TEXT,
                    imei TEXT,
                    status TEXT,
                    created_by TEXT,
                    rep_name TEXT,
                    created_date TEXT,
                    document_status TEXT,
                    source_file TEXT,
                    first_seen_at TEXT,
                    last_seen_at TEXT,
                    cleared INTEGER DEFAULT 0,
                    cleared_at TEXT,
                    sent_count INTEGER DEFAULT 0,
                    last_sent_at TEXT,
                    notes TEXT
                )
                """
            )
            variance_columns = self._table_columns(con, "variances")
            for column, definition in {
                "rep_name": "TEXT DEFAULT ''",
                "created_by": "TEXT DEFAULT ''",
                "created_date": "TEXT DEFAULT ''",
                "document_status": "TEXT DEFAULT ''",
                "source_file": "TEXT DEFAULT ''",
                "notes": "TEXT DEFAULT ''",
            }.items():
                if column not in variance_columns:
                    con.execute(f"ALTER TABLE variances ADD COLUMN {column} {definition}")

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS send_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    variance_key TEXT,
                    district TEXT,
                    group_name TEXT,
                    batch_title TEXT,
                    image_path TEXT,
                    sent_at TEXT,
                    mode TEXT,
                    error TEXT
                )
                """
            )
            send_columns = self._table_columns(con, "send_log")
            if "batch_title" not in send_columns:
                con.execute("ALTER TABLE send_log ADD COLUMN batch_title TEXT DEFAULT ''")

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS inventory_status_state (
                    store_key TEXT PRIMARY KEY,
                    district TEXT,
                    store TEXT,
                    last_status TEXT,
                    previous_status TEXT,
                    status_changed_at TEXT,
                    last_sent_at TEXT,
                    last_loaded_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS whatsapp_groups (
                    district TEXT PRIMARY KEY,
                    group_name TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS district_managers (
                    district TEXT PRIMARY KEY,
                    dm_name TEXT,
                    phone TEXT,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS store_accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    district TEXT NOT NULL,
                    store TEXT NOT NULL,
                    store_key TEXT UNIQUE NOT NULL,
                    account_id TEXT,
                    username TEXT,
                    password TEXT,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS sales_reps (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rep_name TEXT NOT NULL,
                    rep_key TEXT UNIQUE NOT NULL,
                    phone TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS device_exclusions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rule_text TEXT NOT NULL,
                    rule_key TEXT UNIQUE NOT NULL,
                    match_type TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )
            exclusion_columns = self._table_columns(con, "device_exclusions")
            if "product" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN product TEXT DEFAULT ''")
            if "imei" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN imei TEXT DEFAULT ''")
            if "comments" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN comments TEXT DEFAULT ''")
            if "district" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN district TEXT DEFAULT ''")

            # Created By → Employee mapping table (Timesheet Edition)
            con.execute("""
                CREATE TABLE IF NOT EXISTS created_by_mappings (
                    created_by TEXT PRIMARY KEY,
                    employee_name TEXT DEFAULT '',
                    phone TEXT DEFAULT ''
                )
            """)

            # Portal credentials (B2B Soft + Timesheet) — base64-obfuscated passwords
            con.execute("""
                CREATE TABLE IF NOT EXISTS portal_credentials (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL DEFAULT ''
                )
            """)

    def get_created_by_mappings(self) -> List[Dict[str, str]]:
        """Return all created_by → employee mappings."""
        with self.connect() as con:
            rows = con.execute(
                "SELECT created_by, employee_name, phone FROM created_by_mappings ORDER BY created_by"
            ).fetchall()
        return [{"created_by": r[0], "employee_name": r[1], "phone": r[2]} for r in rows]

    def upsert_created_by_mapping(self, created_by: str, employee_name: str, phone: str) -> None:
        with self.connect() as con:
            con.execute(
                "INSERT INTO created_by_mappings(created_by, employee_name, phone) VALUES(?,?,?) "
                "ON CONFLICT(created_by) DO UPDATE SET employee_name=excluded.employee_name, phone=excluded.phone",
                (created_by.strip(), employee_name.strip(), phone.strip()),
            )

    def delete_created_by_mapping(self, created_by: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM created_by_mappings WHERE created_by=?", (created_by,))

    # ── Portal credential helpers ────────────────────────────────────────────
    @staticmethod
    def _obf_cred(v: str) -> str:
        return base64.b64encode(v.encode()).decode() if v else ""

    @staticmethod
    def _deobf_cred(v: str) -> str:
        if not v:
            return ""
        try:
            return base64.b64decode(v.encode()).decode()
        except Exception:
            return v  # legacy plain-text tolerance

    def save_portal_credentials(self, data: dict) -> None:
        """Save portal credentials to the SQLite DB (passwords base64-obfuscated)."""
        rows = []
        for section, fields in data.items():
            for field, value in fields.items():
                key = f"{section}.{field}"
                stored = self._obf_cred(value) if field == "password" else (value or "")
                rows.append((key, stored))
        with self.connect() as con:
            for key, val in rows:
                con.execute(
                    "INSERT INTO portal_credentials(key, value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (key, val),
                )

    def load_portal_credentials(self) -> dict:
        """Load portal credentials from SQLite, returning a nested dict."""
        with self.connect() as con:
            rows = con.execute("SELECT key, value FROM portal_credentials").fetchall()
        result: dict = {}
        for key, val in rows:
            if "." not in key:
                continue
            section, field = key.split(".", 1)
            result.setdefault(section, {})[field] = (
                self._deobf_cred(val) if field == "password" else val
            )
        return result

    def resolve_employee_for_created_by(self, created_by: str) -> Dict[str, str]:
        """Return {employee_name, phone} for a given Created By value, or empty strings."""
        with self.connect() as con:
            row = con.execute(
                "SELECT employee_name, phone FROM created_by_mappings WHERE created_by=?",
                (created_by.strip(),),
            ).fetchone()
        if row:
            return {"employee_name": row[0], "phone": row[1]}
        return {"employee_name": "", "phone": ""}

    def upsert_rows(self, rows: List[VarianceRow]) -> None:
        with self.connect() as con:
            for row in rows:
                existing = con.execute(
                    "SELECT cleared, cleared_at, sent_count, last_sent_at, notes, first_seen_at FROM variances WHERE key=?",
                    (row.key,),
                ).fetchone()
                if existing:
                    con.execute(
                        """
                        UPDATE variances
                        SET district=?, store=?, product=?, imei=?, status=?, created_by=?, rep_name=?, created_date=?,
                            document_status=?, source_file=?, last_seen_at=?
                        WHERE key=?
                        """,
                        (
                            row.district, row.store, row.product, row.imei, row.status,
                            row.created_by, row.rep_name, row.created_date, row.document_status,
                            row.source_file, now_text(), row.key,
                        ),
                    )
                else:
                    con.execute(
                        """
                        INSERT INTO variances (
                            key, district, store, product, imei, status, created_by, rep_name, created_date,
                            document_status, source_file, first_seen_at, last_seen_at, cleared,
                            cleared_at, sent_count, last_sent_at, notes
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, '', 0, '', '')
                        """,
                        (
                            row.key, row.district, row.store, row.product, row.imei, row.status,
                            row.created_by, row.rep_name, row.created_date, row.document_status, row.source_file,
                            now_text(), now_text(),
                        ),
                    )

    def rows(self, include_cleared: bool = False) -> List[VarianceRow]:
        query = """
            SELECT key, district, store, product, imei, status, created_by, rep_name, created_date,
                   document_status, source_file, cleared, sent_count, last_sent_at, cleared_at, notes
            FROM variances
        """
        if not include_cleared:
            query += " WHERE cleared=0"
        query += " ORDER BY district, store, rep_name, status, product"
        with self.connect() as con:
            data = con.execute(query).fetchall()
        rows: List[VarianceRow] = []
        for item in data:
            rows.append(
                VarianceRow(
                    key=item[0], district=item[1] or "", store=item[2] or "", product=item[3] or "",
                    imei=item[4] or "", status=item[5] or "", created_by=item[6] or "",
                    rep_name=item[7] or "", created_date=item[8] or "", document_status=item[9] or "",
                    source_file=item[10] or "", cleared=bool(item[11]), sent_count=int(item[12] or 0),
                    last_sent_at=item[13] or "", cleared_at=item[14] or "", notes=item[15] or "",
                )
            )
        return rows

    def get_rows_by_keys(self, keys: Iterable[str]) -> List[VarianceRow]:
        keys = list(keys)
        if not keys:
            return []
        placeholders = ",".join("?" for _ in keys)
        with self.connect() as con:
            data = con.execute(
                f"""
                SELECT key, district, store, product, imei, status, created_by, rep_name, created_date,
                       document_status, source_file, cleared, sent_count, last_sent_at, cleared_at, notes
                FROM variances WHERE key IN ({placeholders})
                ORDER BY district, store, rep_name, status, product
                """,
                keys,
            ).fetchall()
        return [
            VarianceRow(
                key=i[0], district=i[1] or "", store=i[2] or "", product=i[3] or "", imei=i[4] or "",
                status=i[5] or "", created_by=i[6] or "", rep_name=i[7] or "", created_date=i[8] or "",
                document_status=i[9] or "", source_file=i[10] or "", cleared=bool(i[11]),
                sent_count=int(i[12] or 0), last_sent_at=i[13] or "", cleared_at=i[14] or "", notes=i[15] or "",
            )
            for i in data
        ]

    def set_cleared(self, key: str, cleared: bool) -> None:
        with self.connect() as con:
            con.execute(
                "UPDATE variances SET cleared=?, cleared_at=? WHERE key=?",
                (1 if cleared else 0, now_text() if cleared else "", key),
            )

    def mark_sent(self, rows: Iterable[VarianceRow], group_name: str, batch_title: str, image_path: str, mode: str, error: str = "") -> None:
        sent_at = now_text()
        with self.connect() as con:
            for row in rows:
                con.execute(
                    """
                    INSERT INTO send_log (variance_key, district, group_name, batch_title, image_path, sent_at, mode, error)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (row.key, row.district, group_name, batch_title, image_path, sent_at, mode, error),
                )
                if not error:
                    con.execute(
                        "UPDATE variances SET sent_count=COALESCE(sent_count,0)+1, last_sent_at=? WHERE key=?",
                        (sent_at, row.key),
                    )


    def upsert_inventory_status_rows(self, rows: List[InventoryStatusRow]) -> None:
        loaded_at = now_text()
        with self.connect() as con:
            for row in rows:
                store_key = normalize_store(row.store)
                if not store_key:
                    continue
                existing = con.execute(
                    "SELECT last_status, previous_status, status_changed_at, last_sent_at FROM inventory_status_state WHERE store_key=?",
                    (store_key,),
                ).fetchone()

                last_status = safe_text(row.status)
                if existing:
                    old_status = safe_text(existing[0])
                    old_previous = safe_text(existing[1])
                    old_changed_at = safe_text(existing[2])
                    old_sent_at = safe_text(existing[3])
                    if old_status and normalize_header(old_status) != normalize_header(last_status):
                        previous_status = old_status
                        changed_at = loaded_at
                        # New status means new status still needs its own send state.
                        sent_at = ""
                    else:
                        previous_status = old_previous
                        changed_at = old_changed_at
                        sent_at = old_sent_at
                    con.execute(
                        """
                        UPDATE inventory_status_state
                        SET district=?, store=?, last_status=?, previous_status=?, status_changed_at=?, last_sent_at=?, last_loaded_at=?
                        WHERE store_key=?
                        """,
                        (row.district, row.store, last_status, previous_status, changed_at, sent_at, loaded_at, store_key),
                    )
                else:
                    con.execute(
                        """
                        INSERT INTO inventory_status_state
                        (store_key, district, store, last_status, previous_status, status_changed_at, last_sent_at, last_loaded_at)
                        VALUES (?, ?, ?, ?, '', '', '', ?)
                        """,
                        (store_key, row.district, row.store, last_status, loaded_at),
                    )

    def mark_status_sent(self, rows: Iterable[InventoryStatusRow]) -> None:
        sent_at = now_text()
        with self.connect() as con:
            for row in rows:
                store_key = normalize_store(row.store)
                if not store_key:
                    continue
                con.execute(
                    """
                    INSERT INTO inventory_status_state
                    (store_key, district, store, last_status, previous_status, status_changed_at, last_sent_at, last_loaded_at)
                    VALUES (?, ?, ?, ?, '', '', ?, ?)
                    ON CONFLICT(store_key) DO UPDATE SET
                        district=excluded.district,
                        store=excluded.store,
                        last_status=excluded.last_status,
                        last_sent_at=excluded.last_sent_at,
                        last_loaded_at=excluded.last_loaded_at
                    """,
                    (store_key, row.district, row.store, row.status, sent_at, sent_at),
                )

    def inventory_status_state_map(self) -> Dict[str, Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                """
                SELECT store_key, district, store, last_status, previous_status, status_changed_at, last_sent_at, last_loaded_at
                FROM inventory_status_state
                """
            ).fetchall()
        return {
            r[0]: {
                "District": r[1] or "",
                "Store": r[2] or "",
                "Last Status": r[3] or "",
                "Previous Status": r[4] or "",
                "Status Changed At": r[5] or "",
                "Last Sent At": r[6] or "",
                "Last Loaded At": r[7] or "",
            }
            for r in rows
        }

    def save_whatsapp_group(self, district: str, group_name: str) -> None:
        district = normalize_district(district)
        group_name = safe_text(group_name).strip()
        if not district or district == "Unknown":
            raise ValueError("District is required.")
        if not group_name:
            raise ValueError("WhatsApp group name is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO whatsapp_groups (district, group_name, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(district) DO UPDATE SET
                    group_name=excluded.group_name,
                    updated_at=excluded.updated_at
                """,
                (district, group_name, now_text(), now_text()),
            )

    def delete_whatsapp_group(self, district: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM whatsapp_groups WHERE district=?", (normalize_district(district),))

    def whatsapp_groups(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                "SELECT district, group_name FROM whatsapp_groups ORDER BY district"
            ).fetchall()
        return [{"District": r[0] or "", "Group Name": r[1] or ""} for r in rows]

    def find_whatsapp_group(self, district: str) -> str:
        with self.connect() as con:
            row = con.execute(
                "SELECT group_name FROM whatsapp_groups WHERE district=?",
                (normalize_district(district),),
            ).fetchone()
        return safe_text(row[0]) if row else ""

    def get_setting(self, key: str, default: str = "") -> str:
        """Read a key from the app_settings table (created on first use)."""
        try:
            with self.connect() as con:
                con.execute(
                    "CREATE TABLE IF NOT EXISTS app_settings "
                    "(key TEXT PRIMARY KEY, value TEXT DEFAULT '')"
                )
                row = con.execute(
                    "SELECT value FROM app_settings WHERE key=?", (key,)
                ).fetchone()
            return safe_text(row[0]) if row else default
        except Exception:
            return default

    def save_setting(self, key: str, value: str) -> None:
        """Persist a key/value pair in the app_settings table."""
        try:
            with self.connect() as con:
                con.execute(
                    "CREATE TABLE IF NOT EXISTS app_settings "
                    "(key TEXT PRIMARY KEY, value TEXT DEFAULT '')"
                )
                con.execute(
                    "INSERT INTO app_settings(key,value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (key, value),
                )
        except Exception:
            pass

    def save_district_manager(self, district: str, dm_name: str, phone: str) -> None:
        district = normalize_district(district)
        dm_name = safe_text(dm_name)
        phone = normalize_phone(phone)
        if not district or district == "Unknown":
            raise ValueError("District is required.")
        if not phone:
            raise ValueError("DM phone number is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO district_managers (district, dm_name, phone, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(district) DO UPDATE SET
                    dm_name=excluded.dm_name,
                    phone=excluded.phone,
                    updated_at=excluded.updated_at
                """,
                (district, dm_name, phone, now_text(), now_text()),
            )

    def delete_district_manager(self, district: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM district_managers WHERE district=?", (normalize_district(district),))

    def district_managers(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                "SELECT district, dm_name, phone FROM district_managers ORDER BY district"
            ).fetchall()
        return [{"District": r[0] or "", "DM Name": r[1] or "", "Phone": r[2] or ""} for r in rows]

    def all_known_districts(self) -> List[str]:
        """Return sorted list of unique districts from all district-having tables."""
        districts = set()
        with self.connect() as con:
            for table in ("store_accounts", "whatsapp_groups", "district_managers", "device_exclusions", "variance_log"):
                try:
                    rows = con.execute(f"SELECT DISTINCT district FROM {table} WHERE district IS NOT NULL AND district != ''").fetchall()
                    for r in rows:
                        d = safe_text(r[0]).strip()
                        if d:
                            districts.add(normalize_district(d))
                except Exception:
                    pass
        return sorted(districts)

    def find_district_manager_phone(self, district: str) -> str:
        with self.connect() as con:
            row = con.execute(
                "SELECT phone FROM district_managers WHERE district=?",
                (normalize_district(district),),
            ).fetchone()
        return normalize_phone(row[0]) if row else ""

    def save_store_account(self, district: str, store: str, account_id: str, username: str, password: str) -> None:
        district = normalize_district(district)
        store = display_store(store)
        store_key = normalize_store(store)
        if not store_key:
            raise ValueError("Store name is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO store_accounts (district, store, store_key, account_id, username, password, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(store_key) DO UPDATE SET
                    district=excluded.district,
                    store=excluded.store,
                    account_id=excluded.account_id,
                    username=excluded.username,
                    password=excluded.password,
                    updated_at=excluded.updated_at
                """,
                (district, store, store_key, safe_text(account_id), safe_text(username), safe_text(password), now_text(), now_text()),
            )

    def delete_store_account(self, store_key: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM store_accounts WHERE store_key=?", (store_key,))

    def store_accounts(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                """
                SELECT district, store, store_key, account_id, username, password
                FROM store_accounts
                ORDER BY district, store
                """
            ).fetchall()
        return [
            {
                "District": row[0] or "",
                "Store": row[1] or "",
                "StoreKey": row[2] or "",
                "Account ID": row[3] or "",
                "Username": row[4] or "",
                "Password": row[5] or "",
            }
            for row in rows
        ]

    def replace_store_accounts(self, records: List[Dict[str, str]]) -> int:
        cleaned: Dict[str, Dict[str, str]] = {}
        for rec in records:
            district = normalize_district(rec.get("District", ""))
            store = display_store(rec.get("Store", ""))
            store_key = normalize_store(store)
            if not store_key or not district or district == "Unknown":
                continue
            cleaned[store_key] = {"District": district, "Store": store, "StoreKey": store_key}

        now = now_text()
        last_error = None
        for attempt in range(5):
            try:
                with self.connect() as con:
                    con.execute("BEGIN IMMEDIATE")
                    con.execute("DELETE FROM store_accounts")
                    for rec in sorted(cleaned.values(), key=lambda r: (r["District"], r["Store"].lower())):
                        con.execute(
                            """
                            INSERT INTO store_accounts (district, store, store_key, account_id, username, password, created_at, updated_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (rec["District"], rec["Store"], rec["StoreKey"], "", "", "", now, now),
                        )
                    con.commit()
                return len(cleaned)
            except sqlite3.OperationalError as exc:
                last_error = exc
                if "locked" not in str(exc).lower():
                    raise
                time.sleep(0.8 + attempt * 0.6)
        raise RuntimeError("Database is locked. Close any other running copy of this GUI, then import again.") from last_error

    def store_master_records(self) -> List[Dict[str, str]]:
        return [{"District": row["District"], "Store": row["Store"]} for row in self.store_accounts()]

    def save_sales_rep(self, rep_name: str, phone: str) -> None:
        rep_name = safe_text(rep_name)
        phone = normalize_phone(phone)
        rep_key = person_name_key(rep_name)
        if not rep_key:
            raise ValueError("Sales rep name is required.")
        if not phone:
            raise ValueError("Phone number is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO sales_reps (rep_name, rep_key, phone, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(rep_key) DO UPDATE SET
                    rep_name=excluded.rep_name,
                    phone=excluded.phone,
                    updated_at=excluded.updated_at
                """,
                (rep_name, rep_key, phone, now_text(), now_text()),
            )

    def delete_sales_rep(self, rep_key: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM sales_reps WHERE rep_key=?", (rep_key,))

    def sales_reps(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                "SELECT rep_name, rep_key, phone FROM sales_reps ORDER BY rep_name"
            ).fetchall()
        return [{"Rep Name": r[0] or "", "RepKey": r[1] or "", "Phone": r[2] or ""} for r in rows]

    def sales_rep_phone_map(self) -> Dict[str, str]:
        return {row["RepKey"]: row["Phone"] for row in self.sales_reps()}

    def find_sales_rep_phone(self, rep_name: str) -> str:
        key = person_name_key(rep_name)
        if not key:
            return ""
        reps = self.sales_reps()
        exact = {row["RepKey"]: row["Phone"] for row in reps}
        if key in exact:
            return exact[key]
        # Drop single-character tokens (initials like "E" from "Menatallah.E",
        # "M" from "Abdullah.M"). person_name_key() splits on punctuation, so
        # dotted usernames become e.g. {"e", "menatallah"} — and the single
        # letter becomes a required word in the subset check below, which
        # breaks matching against employees whose names don't contain that
        # letter as a standalone word. Filtering both sides by len > 1
        # restores the subset match: {"menatallah"} <= {"elsafty", "menatallah"}.
        key_parts = {p for p in set(key.split()) if len(p) > 1}
        for row in reps:
            rep_parts = {p for p in set((row["RepKey"] or "").split()) if len(p) > 1}
            if key_parts and rep_parts and (key_parts <= rep_parts or rep_parts <= key_parts):
                return row["Phone"]
        return ""

    def resolve_phone_for_rep(self, rep_name: str, created_by: str = "") -> str:
        """Timesheet-aware phone resolver.

        Resolution order:
        1. created_by_mappings by created_by username (most direct — user set this)
        2. created_by_mappings by employee_name (fuzzy/normalized match)
        3. sales_reps table by rep_name (fuzzy/normalized match)
        4. created_by as a loose sales_reps key (last resort)

        Name matching uses person_name_key() normalization: lowercases,
        strips punctuation, sorts word tokens.  So "John Smith" matches
        "smith, john" or "John  Smith" (double space) or "JOHN SMITH".
        """
        # 1. Created By mapping by username
        if created_by:
            cb = created_by.strip()
            mapping = self.resolve_employee_for_created_by(cb)
            if mapping.get("phone"):
                return mapping["phone"]

        # 2. Created By mapping by employee_name (fuzzy match)
        if rep_name:
            rn_key = person_name_key(rep_name)
            if rn_key:
                for m in self.get_created_by_mappings():
                    emp_name = m.get("employee_name", "").strip()
                    if emp_name and person_name_key(emp_name) == rn_key and m.get("phone"):
                        return m["phone"]

        # 3. Standard sales_reps lookup (fuzzy match)
        phone = self.find_sales_rep_phone(rep_name)
        if phone:
            return phone

        # 4. Try created_by as a name in sales_reps
        if created_by:
            phone = self.find_sales_rep_phone(created_by)
            if phone:
                return phone

        return ""

    def update_device_exclusion_comment(self, rule_key: str, comments: str) -> None:
        with self.connect() as con:
            con.execute(
                "UPDATE device_exclusions SET comments=?, updated_at=? WHERE rule_key=?",
                (safe_text(comments), now_text(), rule_key),
            )

    def save_device_exclusion(self, product: str = "", imei: str = "", comments: str = "", district: str = "") -> None:
        product = safe_text(product)
        imei = safe_text(imei)
        comments = safe_text(comments)
        district = normalize_district(district) if safe_text(district) else ""

        if not product and not imei:
            raise ValueError("Product or IMEI is required.")

        # IMEI exclusions stay global across all districts.
        if imei:
            district = ""
            rule_text = imei
            match_type = "IMEI Exact"
        else:
            if not district or district == "Unknown":
                raise ValueError("District is required for Product exclusions.")
            rule_text = product
            match_type = "Product Contains"

        # Include district in product exclusion key, but keep IMEI exclusions global.
        rule_key = device_rule_key(f"{district}|{product}|{imei}" if not imei else f"|{imei}")
        if not rule_key:
            raise ValueError("Product or IMEI is required.")

        with self.connect() as con:
            # Prevent duplicates by ensuring old keys or un-districted versions are cleaned up
            if imei:
                con.execute("DELETE FROM device_exclusions WHERE imei=?", (imei,))
            elif product and district:
                con.execute("DELETE FROM device_exclusions WHERE product=? AND district=?", (product, district))

            con.execute(
                """
                INSERT INTO device_exclusions (rule_text, rule_key, match_type, product, imei, comments, district, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(rule_key) DO UPDATE SET
                    rule_text=excluded.rule_text,
                    match_type=excluded.match_type,
                    product=excluded.product,
                    imei=excluded.imei,
                    comments=excluded.comments,
                    district=excluded.district,
                    updated_at=excluded.updated_at
                """,
                (rule_text, rule_key, match_type, product, imei, comments, district, now_text(), now_text()),
            )

    def delete_device_exclusion(self, rule_key: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM device_exclusions WHERE rule_key=?", (rule_key,))

    def delete_all_device_exclusions(self) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM device_exclusions")

    def device_exclusions(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                """
                SELECT rule_text, rule_key, match_type, product, imei, comments, district
                FROM device_exclusions
                ORDER BY COALESCE(district, ''), COALESCE(product, ''), COALESCE(imei, '')
                """
            ).fetchall()

        out: List[Dict[str, str]] = []
        for r in rows:
            rule_text = r[0] or ""
            rule_key = r[1] or ""
            match_type = r[2] or ""
            product = r[3] or ""
            imei = r[4] or ""
            comments = r[5] or ""
            district = normalize_district(r[6] or "") if safe_text(r[6] or "") else ""

            # Backfill display values for exclusions created in old versions.
            if not product and not imei and rule_text:
                if normalize_header(match_type) in {"imeiexact", "serialexact", "esnexact"}:
                    imei = rule_text
                    district = ""
                else:
                    product = rule_text

            if imei:
                district = ""

            out.append({
                "District": district,
                "Product": product,
                "IMEI": imei,
                "Comments": comments,
                "Rule Text": rule_text,
                "RuleKey": rule_key,
                "Match Type": match_type,
            })
        return out

    def is_device_excluded(self, district: str, product: str, imei: str) -> bool:
        row_district = normalize_district(district)
        product_clean = device_rule_key(product)
        imei_clean = device_rule_key(imei)

        for rule in self.device_exclusions():
            rule_product = device_rule_key(rule.get("Product", ""))
            rule_imei = device_rule_key(rule.get("IMEI", ""))
            rule_district = normalize_district(rule.get("District", "")) if safe_text(rule.get("District", "")) else ""

            # IMEI exclusions stay global. District is ignored for IMEI.
            if rule_imei and imei_clean and rule_imei == imei_clean:
                return True

            # Product exclusions are district-specific only.
            if rule_product and product_clean and rule_product in product_clean:
                if rule_district and normalize_district(rule_district) == row_district:
                    return True

        return False

    def exclusion_reason(self, district: str, product: str, imei: str) -> str:
        row_district = normalize_district(district)
        product_clean = device_rule_key(product)
        imei_clean = device_rule_key(imei)

        for rule in self.device_exclusions():
            rule_product = device_rule_key(rule.get("Product", ""))
            rule_imei = device_rule_key(rule.get("IMEI", ""))
            rule_district = normalize_district(rule.get("District", "")) if safe_text(rule.get("District", "")) else ""

            if rule_imei and imei_clean and rule_imei == imei_clean:
                return f"IMEI: {rule.get('IMEI', '')}"

            if rule_product and product_clean and rule_product in product_clean:
                if rule_district and normalize_district(rule_district) == row_district:
                    return f"{rule_district} Product: {rule.get('Product', '')}"

        return ""

    def export_xlsx(self, output_path: Path, keys: Optional[Iterable[str]] = None) -> None:
        keys = list(keys) if keys is not None else None
        where_sql = ""
        params: List[str] = []
        if keys is not None:
            if not keys:
                headers = [
                    "District", "Store", "Product", "IMEI", "Status", "Rep Name", "Count By", "Created Date",
                    "Document Status", "Clearance Status", "Cleared At", "Sent Count", "Last Sent At",
                    "First Seen At", "Last Seen At", "Source File", "Notes"
                ]
                output_path.parent.mkdir(parents=True, exist_ok=True)
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(headers)
                wb.save(output_path)
                return
            placeholders = ",".join("?" for _ in keys)
            where_sql = f"WHERE key IN ({placeholders})"
            params = keys

        query = f"""
                SELECT district, store, product, imei, status, rep_name, created_by, created_date,
                       document_status,
                       CASE WHEN cleared=1 THEN 'Cleared' ELSE 'Not Cleared' END,
                       cleared_at, sent_count, last_sent_at, first_seen_at, last_seen_at, source_file, notes
                FROM variances
                {where_sql}
                ORDER BY district, store, rep_name, status, product
                """
        with self.connect() as con:
            rows = con.execute(query, params).fetchall()

        headers = [
            "District", "Store", "Product", "IMEI", "Status", "Rep Name", "Count By", "Created Date",
            "Document Status", "Clearance Status", "Cleared At", "Sent Count", "Last Sent At",
            "First Seen At", "Last Seen At", "Source File", "Notes"
        ]

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)

        imei_col_idx = headers.index("IMEI") + 1

        for r_idx, row in enumerate(rows, start=2):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'

        wb.save(output_path)

class ImageRenderer:
    def __init__(self):
        IMAGE_DIR.mkdir(parents=True, exist_ok=True)
        if Image is None:
            raise RuntimeError("Pillow is required. Install with: py -m pip install pillow")

    @staticmethod
    def _font(size: int, bold: bool = False):
        candidates = []
        if bold:
            candidates.extend(["arialbd.ttf", "Arial Bold.ttf", "segoeuib.ttf"])
        candidates.extend(["arial.ttf", "Segoe UI.ttf", "DejaVuSans.ttf"])
        for candidate in candidates:
            try:
                return ImageFont.truetype(candidate, size)
            except Exception:
                continue
        return ImageFont.load_default()

    @staticmethod
    def _wrap(draw, text: str, font, max_width: int) -> List[str]:
        text = safe_text(text)
        if not text:
            return [""]
        words = text.split()
        lines: List[str] = []
        current = ""
        for word in words:
            probe = word if not current else current + " " + word
            bbox = draw.textbbox((0, 0), probe, font=font)
            if bbox[2] - bbox[0] <= max_width:
                current = probe
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines or [""]

    def render_rows(self, batch_title: str, rows: List[VarianceRow], mode: str = "pending") -> Path:
        if not rows:
            raise ValueError("No rows to render")

        width = 1560
        margin = 32
        title_font = self._font(34, True)
        sub_font = self._font(18, False)
        header_font = self._font(20, True)
        cell_font = self._font(18, False)
        small_font = self._font(16, False)
        bold_small_font = self._font(16, True)
        row_height_base = 52

        tmp = Image.new("RGB", (width, 400), "white")
        draw = ImageDraw.Draw(tmp)
        col_widths = [165, 220, 480, 205, 140, 250]
        product_width = col_widths[2] - 20
        rep_width = col_widths[5] - 20
        row_heights = []
        for row in rows:
            product_lines = self._wrap(draw, str(row.product or ""), cell_font, product_width)
            rep_lines = self._wrap(draw, str(row.rep_name or ""), small_font, rep_width)
            row_heights.append(max(row_height_base, 26 * max(len(product_lines), len(rep_lines)) + 22))

        dark = (18, 20, 43)
        gray = (95, 95, 102)
        red = (233, 27, 47)
        light = (246, 247, 249)
        border = (215, 218, 223)
        navy = (9, 13, 38)
        white = (255, 255, 255)

        logo_img = None
        logo_w = 0
        logo_h_used = 0
        if STATUS_LOGO_PATH.exists():
            try:
                logo_img = Image.open(STATUS_LOGO_PATH).convert("RGBA")
                scale = min(150 / logo_img.width, 56 / logo_img.height)
                size = (max(1, int(logo_img.width * scale)), max(1, int(logo_img.height * scale)))
                logo_img = logo_img.resize(size)
                logo_w, logo_h_used = size
            except Exception:
                logo_img = None
                logo_w = 0
                logo_h_used = 0

        # Standard GFH brand bar (matches gfh_inventory_aging_processor's
        # wa_render_table_image and every other GFH/VDS app's window header):
        # navy bar, logo + red divider on the left, title centered in white.
        # Previously this was a white background with the logo/divider just
        # floating on it and black text - not the navy bar every other app uses.
        brand_bar_h = 90
        table_top = brand_bar_h + 18
        table_height = 54 + sum(row_heights)
        height = table_top + table_height + 58

        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        clean_mode = safe_text(mode).replace("manual_", "").replace("pending_", "")
        mode_label = SEND_MODE_LABELS.get(clean_mode, clean_mode.replace("_", " ").title())

        draw.rectangle([0, 0, width, brand_bar_h], fill=navy)
        logo_right_edge = margin
        if logo_img is not None:
            img.paste(logo_img, (margin, (brand_bar_h - logo_h_used) // 2), logo_img)
            logo_right_edge = margin + logo_w
        divider_x = logo_right_edge + 14
        draw.rectangle([divider_x, 14, divider_x + 3, brand_bar_h - 14], fill=red)

        title_text = "GFH Inventory Variance"
        sub_text = f"{mode_label}: {batch_title}   Rows: {len(rows)}   Generated: {now_text()}"
        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_w = title_bbox[2] - title_bbox[0]
        sub_bbox = draw.textbbox((0, 0), sub_text, font=sub_font)
        sub_w = sub_bbox[2] - sub_bbox[0]
        draw.text(((width - title_w) // 2, 16), title_text, fill=white, font=title_font)
        draw.text(((width - sub_w) // 2, 16 + 38), sub_text, fill=white, font=sub_font)

        y = table_top
        x = margin
        headers = ["District", "Store", "Product", "IMEI", "Status", "Rep Name"]
        draw.rectangle((x, y, width - margin, y + 54), fill=red)
        cx = x
        for idx, header in enumerate(headers):
            draw.text((cx + 10, y + 14), header, fill="white", font=header_font)
            cx += col_widths[idx]
        y += 54

        for idx, row in enumerate(rows):
            fill = light if idx % 2 == 0 else (255, 255, 255)
            rh = row_heights[idx]
            draw.rectangle((x, y, width - margin, y + rh), fill=fill, outline=border)
            values = [row.district, row.store, row.product, row.imei, row.status, row.rep_name]
            cx = x
            for cidx, value in enumerate(values):
                max_w = col_widths[cidx] - 20
                font = bold_small_font if cidx in {0, 1, 5} else (small_font if cidx == 4 else cell_font)
                lines = self._wrap(draw, str(value or ""), font, max_w)
                yy = y + 11
                for line in lines[:4]:
                    draw.text((cx + 10, yy), line, fill=dark, font=font)
                    yy += 25
                cx += col_widths[cidx]
            y += rh

        footer = "Please provide resolution image or valid variance explanation. Cleared variances will not be auto-sent again."
        draw.text((margin, height - 36), footer, fill=gray, font=small_font)

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", batch_title)[:60] or "Variance"
        safe_mode = re.sub(r"[^A-Za-z0-9_-]+", "_", safe_text(mode).replace("manual_", "").replace("pending_", ""))[:30] or "mode"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = IMAGE_DIR / f"GFH_Variance_{safe_mode}_{safe_title}_{stamp}.png"
        img.save(path)
        return path


# ── ROBUST WHATSAPP DESKTOP LAYER ───────────────────────────────────────────
# Root causes of "WhatsApp opens but nothing is typed":
#  1. _wa_find_hwnd matched ANY window with "WhatsApp" in the title (an Edge
#     "WhatsApp Web" tab, our own window...), so keystrokes went elsewhere.
#  2. SetForegroundWindow is silently refused by Windows (foreground lock) and
#     the result was never checked, so typing started in the wrong window.
#  3. keybd_event without scan codes / per-char VkKeyScan typing is ignored by
#     WhatsApp's WebView2 UI; text now goes through the clipboard (Ctrl+V) with
#     SendInput-Unicode as the fallback.
# Every step now re-verifies WhatsApp is the foreground window and ABORTS
# instead of typing into whatever else happens to be in front.
import ctypes as _wa_ct
from ctypes import wintypes as _wa_wt

_WA_BROWSER_PROCS = ("msedge", "chrome", "firefox", "brave", "opera",
                     "iexplore", "vivaldi")
_WA_SELF_PROCS = ("vidapay", "python")
_WA_EXT_VKS = (0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x2D, 0x2E)


def _wa_user32():
    u = _wa_ct.windll.user32
    u.GetForegroundWindow.restype = _wa_wt.HWND
    u.GetAncestor.restype = _wa_wt.HWND
    u.GetAncestor.argtypes = [_wa_wt.HWND, _wa_wt.UINT]
    return u


def _wa_proc_name(pid):
    """Lower-case exe name of a process id ('' when unknown)."""
    try:
        k32 = _wa_ct.windll.kernel32
        k32.OpenProcess.restype = _wa_ct.c_void_p
        k32.OpenProcess.argtypes = [_wa_wt.DWORD, _wa_wt.BOOL, _wa_wt.DWORD]
        k32.QueryFullProcessImageNameW.argtypes = [
            _wa_ct.c_void_p, _wa_wt.DWORD, _wa_wt.LPWSTR,
            _wa_ct.POINTER(_wa_wt.DWORD)]
        k32.CloseHandle.argtypes = [_wa_ct.c_void_p]
        h = k32.OpenProcess(0x1000, False, pid)   # QUERY_LIMITED_INFORMATION
        if not h:
            return ""
        try:
            buf = _wa_ct.create_unicode_buffer(520)
            n = _wa_wt.DWORD(520)
            if k32.QueryFullProcessImageNameW(h, 0, buf, _wa_ct.byref(n)):
                return os.path.basename(buf.value).lower()
        finally:
            k32.CloseHandle(h)
    except Exception:
        pass
    return ""


def _wa_window_pid(hwnd):
    pid = _wa_wt.DWORD(0)
    _wa_user32().GetWindowThreadProcessId(hwnd, _wa_ct.byref(pid))
    return pid.value


def _wa_find_hwnd():
    """HWND of the real WhatsApp Desktop window (never a browser tab or this
    app's own window), or None. Minimised windows are included."""
    if sys.platform != "win32":
        return None
    try:
        u = _wa_user32()
        own = os.getpid()
        found = []
        enum_proc = _wa_ct.WINFUNCTYPE(_wa_wt.BOOL, _wa_wt.HWND, _wa_wt.LPARAM)

        def _cb(hwnd, _lp):
            try:
                if not u.IsWindowVisible(hwnd):
                    return True
                n = u.GetWindowTextLengthW(hwnd)
                if n <= 0:
                    return True
                buf = _wa_ct.create_unicode_buffer(n + 1)
                u.GetWindowTextW(hwnd, buf, n + 1)
                title = buf.value
                if "whatsapp" not in title.lower():
                    return True
                pid = _wa_window_pid(hwnd)
                if pid == own:
                    return True
                pn = _wa_proc_name(pid)
                if any(b in pn for b in _WA_BROWSER_PROCS):
                    return True
                if any(pn.startswith(s) for s in _WA_SELF_PROCS):
                    return True
                if "whatsapp" in pn:
                    score = 0
                elif pn.startswith("applicationframehost"):
                    score = 1
                elif title.strip().lower() == "whatsapp":
                    score = 2
                else:
                    return True
                found.append((score, hwnd))
            except Exception:
                pass
            return True

        u.EnumWindows(enum_proc(_cb), 0)
        found.sort(key=lambda t: t[0])
        return found[0][1] if found else None
    except Exception:
        return None


def _wa_is_foreground(hwnd):
    """True when WhatsApp (or its child/core window) owns the foreground."""
    try:
        u = _wa_user32()
        fg = u.GetForegroundWindow()
        if not fg:
            return False
        if fg == hwnd or u.GetAncestor(fg, 2) == hwnd:   # GA_ROOT
            return True
        return "whatsapp" in _wa_proc_name(_wa_window_pid(fg))
    except Exception:
        return False


def _wa_focus_hwnd(hwnd, log=None):
    """Bring WhatsApp to the foreground and VERIFY it. Returns True/False.
    Escalates: ALT-tap unlock + AttachThreadInput -> SwitchToThisWindow ->
    minimise/restore. Only un-minimises when actually minimised, so a snapped
    or maximised window keeps its layout."""
    try:
        u = _wa_user32()
        k32 = _wa_ct.windll.kernel32
        try:   # allow this process to take the foreground
            u.SystemParametersInfoW(0x2001, 0, 0, 0x0002)
        except Exception:
            pass
        if u.IsIconic(hwnd):
            u.ShowWindow(hwnd, 9)          # SW_RESTORE (from minimised only)
            time.sleep(0.5)
        for attempt in range(1, 5):
            _wa_win32_key_down(0x12)       # ALT tap unlocks SetForegroundWindow
            _wa_win32_key_up(0x12)
            fg = u.GetForegroundWindow()
            fg_thread = u.GetWindowThreadProcessId(fg, None) if fg else 0
            our_thread = k32.GetCurrentThreadId()
            wa_thread = u.GetWindowThreadProcessId(hwnd, None)
            pairs = [(fg_thread, our_thread), (fg_thread, wa_thread)]
            for a, b in pairs:
                if a and b and a != b:
                    u.AttachThreadInput(a, b, True)
            u.BringWindowToTop(hwnd)
            u.ShowWindow(hwnd, 5)          # SW_SHOW (keeps maximised state)
            u.SetForegroundWindow(hwnd)
            for a, b in pairs:
                if a and b and a != b:
                    u.AttachThreadInput(a, b, False)
            if attempt >= 2:
                u.SwitchToThisWindow(hwnd, True)
            if attempt >= 3:
                u.ShowWindow(hwnd, 6)      # SW_MINIMIZE
                time.sleep(0.3)
                u.ShowWindow(hwnd, 9)      # restores previous state
            time.sleep(0.4)
            if _wa_is_foreground(hwnd):
                return True
    except Exception as exc:
        if log:
            log(f"WhatsApp focus error: {exc}")
    return False


def _wa_ensure_focused(log=None):
    log = log or _wa_log
    """Find (launching if needed) WhatsApp Desktop, focus it, verify it.
    Returns hwnd on success, None on failure."""
    hwnd = _wa_find_hwnd()
    if hwnd is None:
        log("WhatsApp Desktop not open — launching...")
        _wa_launch()
        for _ in range(40):                # wait up to 20 s for the window
            time.sleep(0.5)
            hwnd = _wa_find_hwnd()
            if hwnd:
                time.sleep(2.0)            # let the chat list finish loading
                break
    if hwnd is None:
        log("WhatsApp Desktop window not found — is it installed and signed in?")
        return None
    for _ in range(3):
        if _wa_focus_hwnd(hwnd, log):
            time.sleep(0.4)
            log("WhatsApp Desktop focused.")
            return hwnd
        hwnd = _wa_find_hwnd() or hwnd
        time.sleep(1.0)
    hint = ""
    try:
        if _wa_ct.windll.shell32.IsUserAnAdmin():
            hint = (" This app is running as Administrator; Windows blocks it "
                    "from controlling a normal-user WhatsApp - run both at "
                    "the same level.")
    except Exception:
        pass
    log("Could not bring WhatsApp Desktop to the front." + hint)
    return None


def _wa_guard(hwnd, log=None):
    log = log or _wa_log
    """Re-verify (and re-take) focus right before typing. Never type blind."""
    if _wa_is_foreground(hwnd):
        return True
    log("WhatsApp lost focus — re-focusing...")
    if _wa_focus_hwnd(hwnd, log):
        time.sleep(0.3)
        return True
    log("WhatsApp is not in front — aborting so nothing is typed elsewhere.")
    return False


# ---- keyboard: scan-code aware key events + Unicode SendInput text ---------
class _WA_KEYBDINPUT(_wa_ct.Structure):
    _fields_ = [("wVk", _wa_wt.WORD), ("wScan", _wa_wt.WORD),
                ("dwFlags", _wa_wt.DWORD), ("time", _wa_wt.DWORD),
                ("dwExtraInfo", _wa_ct.c_size_t)]


class _WA_MOUSEINPUT(_wa_ct.Structure):
    _fields_ = [("dx", _wa_ct.c_long), ("dy", _wa_ct.c_long),
                ("mouseData", _wa_wt.DWORD), ("dwFlags", _wa_wt.DWORD),
                ("time", _wa_wt.DWORD), ("dwExtraInfo", _wa_ct.c_size_t)]


class _WA_INPUT(_wa_ct.Structure):
    class _U(_wa_ct.Union):
        _fields_ = [("ki", _WA_KEYBDINPUT), ("mi", _WA_MOUSEINPUT)]
    _anonymous_ = ("u",)
    _fields_ = [("type", _wa_wt.DWORD), ("u", _U)]


def _wa_send_key(vk, scan, flags):
    inp = _WA_INPUT(type=1)
    inp.ki = _WA_KEYBDINPUT(vk, scan, flags, 0, 0)
    _wa_ct.windll.user32.SendInput(1, _wa_ct.byref(inp), _wa_ct.sizeof(inp))


def _wa_win32_key_down(vk_code):
    scan = _wa_ct.windll.user32.MapVirtualKeyW(vk_code, 0)
    _wa_send_key(vk_code, scan, 0x1 if vk_code in _WA_EXT_VKS else 0)


def _wa_win32_key_up(vk_code):
    scan = _wa_ct.windll.user32.MapVirtualKeyW(vk_code, 0)
    _wa_send_key(vk_code, scan, 0x2 | (0x1 if vk_code in _WA_EXT_VKS else 0))


def _wa_win32_hotkey(*keys):
    codes = []
    for k in keys:
        kl = k.lower()
        if kl in _WA_VK_MAP:
            codes.append(_WA_VK_MAP[kl])
        elif len(k) == 1:
            codes.append(ord(k.upper()))
        else:
            return
    for c in codes:
        _wa_win32_key_down(c)
        time.sleep(0.04)
    time.sleep(0.04)
    for c in reversed(codes):
        _wa_win32_key_up(c)
        time.sleep(0.04)


def _wa_win32_press(key):
    kl = key.lower()
    if kl in _WA_VK_MAP:
        code = _WA_VK_MAP[kl]
    elif len(key) == 1:
        code = ord(key.upper())
    else:
        return
    _wa_win32_key_down(code)
    time.sleep(0.04)
    _wa_win32_key_up(code)
    time.sleep(0.04)


def _wa_win32_write(text, interval=0.0):
    """Type text as Unicode characters (layout independent, any language)."""
    for ch in str(text):
        for up in (0, 2):
            _wa_send_key(0, ord(ch), 0x4 | up)      # KEYEVENTF_UNICODE
        time.sleep(max(0.01, interval))


def _wa_set_clipboard_text(text):
    """Put plain text on the clipboard via ctypes (no pyperclip needed)."""
    u = _wa_ct.windll.user32
    k = _wa_ct.windll.kernel32
    k.GlobalAlloc.argtypes = [_wa_wt.UINT, _wa_ct.c_size_t]
    k.GlobalAlloc.restype = _wa_ct.c_void_p
    k.GlobalLock.argtypes = [_wa_ct.c_void_p]
    k.GlobalLock.restype = _wa_ct.c_void_p
    k.GlobalUnlock.argtypes = [_wa_ct.c_void_p]
    u.OpenClipboard.argtypes = [_wa_ct.c_void_p]
    u.SetClipboardData.argtypes = [_wa_wt.UINT, _wa_ct.c_void_p]
    u.SetClipboardData.restype = _wa_ct.c_void_p
    data = (str(text) + "\x00").encode("utf-16-le")
    for _ in range(10):                      # clipboard may be briefly busy
        if u.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        return False
    try:
        u.EmptyClipboard()
        h = k.GlobalAlloc(0x0002, len(data))
        if not h:
            return False
        p = k.GlobalLock(h)
        _wa_ct.memmove(p, data, len(data))
        k.GlobalUnlock(h)
        return bool(u.SetClipboardData(13, h))   # CF_UNICODETEXT
    finally:
        u.CloseClipboard()



_WA_LOGGER = None


def _wa_log(msg):
    try:
        (_WA_LOGGER or print)(msg)
    except Exception:
        pass


def _wa_launch():
    """Open WhatsApp Desktop (several methods)."""
    import subprocess as _sp
    for cmd in (["explorer.exe", "whatsapp:"],
                "start whatsapp:",
                r'explorer.exe "shell:AppsFolder\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"'):
        try:
            _sp.Popen(cmd, shell=isinstance(cmd, str))
            return
        except Exception:
            continue


def _wa_active_guard():
    """Never type blind: WhatsApp must be the foreground window right now."""
    hwnd = _wa_find_hwnd()
    if hwnd is None:
        raise RuntimeError("WhatsApp Desktop window not found - nothing was typed.")
    if not _wa_guard(hwnd):
        raise RuntimeError("WhatsApp Desktop is not in front - nothing was typed.")


def _wa_paste_text(text):
    """Enter text into the focused field: clipboard + Ctrl+V, falling back to
    Unicode typing (guarded)."""
    _wa_active_guard()
    try:
        if _wa_set_clipboard_text(text):
            time.sleep(0.2)
            _wa_win32_hotkey("ctrl", "v")
            time.sleep(0.3)
            return
    except Exception:
        pass
    _wa_win32_write(text, 0.02)


def _wa_hotkey(*keys):
    _wa_active_guard()
    _wa_win32_hotkey(*keys)


def _wa_press(key):
    _wa_active_guard()
    _wa_win32_press(key)


def _wa_write(text, interval=0.0):
    _wa_active_guard()
    _wa_win32_write(text, interval)
# ── END ROBUST WHATSAPP DESKTOP LAYER ───────────────────────────────────────



class WhatsAppSender:
    def __init__(self, status_callback=None, mode="desktop"):
        self.status_callback = status_callback or (lambda text: None)
        # "desktop" uses pyautogui to drive WhatsApp Desktop app.
        # "web" opens web.whatsapp.com in the default browser.
        self.mode = mode if mode in ("desktop", "web") else "desktop"

    def log(self, text: str) -> None:
        self.status_callback(text)

    def _find_whatsapp_hwnd(self) -> int:
        """hwnd of the REAL WhatsApp Desktop window (not a browser tab), or 0."""
        return _wa_find_hwnd() or 0

    def _save_whatsapp_rect(self) -> tuple:
        """Save WhatsApp window rect before focus (snap layout preservation)."""
        try:
            import ctypes
            from ctypes import wintypes
            hwnd = self._find_whatsapp_hwnd()
            if not hwnd:
                return (0, None)
            rect = wintypes.RECT()
            ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
            return (hwnd, (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top))
        except Exception:
            return (0, None)

    def _restore_whatsapp_rect(self, hwnd: int, saved_rect: tuple) -> None:
        """Restore WhatsApp window to saved rect so snap layout is not disrupted."""
        if not hwnd or not saved_rect:
            return
        try:
            import ctypes
            x, y, w, h = saved_rect
            SWP_NOZORDER = 0x0004
            SWP_NOACTIVATE = 0x0010
            ctypes.windll.user32.SetWindowPos(hwnd, 0, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE)
        except Exception:
            pass

    def _force_focus_whatsapp(self) -> bool:
        """Find WhatsApp Desktop (launching it if needed), bring it to the front
        and VERIFY it is the foreground window. False => nothing may be typed."""
        global _WA_LOGGER
        _WA_LOGGER = self.log
        return _wa_ensure_focused(self.log) is not None

    @staticmethod
    def _import_pyautogui():
        try:
            import pyautogui
            pyautogui.FAILSAFE = True
            return pyautogui
        except Exception as exc:
            raise RuntimeError("pyautogui is required. Install with: py -m pip install pyautogui") from exc

    @staticmethod
    def _open_whatsapp() -> None:
        if sys.platform.startswith("win"):
            try:
                subprocess.Popen("start whatsapp:", shell=True)
                time.sleep(3)
                return
            except Exception:
                pass
        time.sleep(1)

    @staticmethod
    def _activate_whatsapp_window() -> bool:
        try:
            import pygetwindow as gw
            windows = [w for w in gw.getAllWindows() if "whatsapp" in (w.title or "").lower()]
            if not windows:
                return False
            win = windows[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            time.sleep(1)
            return True
        except Exception:
            return False

    @staticmethod
    def _copy_image_to_clipboard(image_path: Path) -> None:
        """Copy image to clipboard — same as GFH_Inventory_Audit.py (proven working)."""
        if Image is None:
            raise RuntimeError("Pillow is required to copy images to clipboard.")
        if not sys.platform.startswith("win"):
            raise RuntimeError("Automatic image clipboard paste is implemented for Windows only.")
        try:
            import win32clipboard
            import win32con
        except Exception as exc:
            raise RuntimeError("pywin32 is required. Install with: py -m pip install pywin32") from exc
        image = Image.open(image_path).convert("RGB")
        output = BytesIO()
        image.save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_DIB, data)
        finally:
            win32clipboard.CloseClipboard()

    @staticmethod
    def _type_group_search(pyautogui, group_name: str) -> None:
        shortcut = SEARCH_SHORTCUT.lower().strip()
        if "+" in shortcut:
            keys = [k.strip() for k in shortcut.split("+") if k.strip()]
        else:
            keys = [k.strip() for k in shortcut.split() if k.strip()]
        if not keys:
            keys = ["ctrl", "f"]
        _wa_hotkey(*keys)
        time.sleep(0.7)
        _wa_hotkey("ctrl", "a")
        time.sleep(0.2)
        _wa_paste_text(group_name)
        time.sleep(1.2)
        _wa_press("enter")
        time.sleep(1.5)

    @staticmethod
    def _paste_text(pyautogui, text: str) -> None:
        text = safe_text(text)
        if not text:
            return
        _wa_paste_text(text)

    def send_image(self, group_name: str, image_path: Path, text_message: str = "") -> None:
        if self.mode == "web":
            self._send_image_web(group_name, image_path, text_message)
            return
        pyautogui = self._import_pyautogui()
        self.log(f"Opening WhatsApp Desktop for {group_name}...")
        self._open_whatsapp()
        wa_hwnd, wa_rect = self._save_whatsapp_rect()
        if not self._force_focus_whatsapp():
            raise RuntimeError("Could not bring WhatsApp Desktop to the front - nothing was typed.")
        self.log(f"Searching group: {group_name}")
        self._type_group_search(pyautogui, group_name)

        self.log("Copying image to clipboard...")
        self._copy_image_to_clipboard(image_path)

        # Paste the image — this opens WhatsApp's image-preview/send dialog
        _wa_hotkey("ctrl", "v")
        # Wait for the preview dialog to fully load before doing anything else
        time.sleep(3.0)

        caption = safe_text(text_message)
        if caption:
            self.log("Typing caption into image preview field...")
            # WhatsApp Desktop focuses the caption field automatically when the
            # image preview dialog opens. Type directly into it.
            self._paste_text(pyautogui, caption)
            time.sleep(0.8)

        # Send the image (Enter confirms the image-preview dialog)
        _wa_press("enter")
        time.sleep(2.0)
        self._restore_whatsapp_rect(wa_hwnd, wa_rect)
        self.log(f"Sent image to {group_name}")

    def _send_image_web(self, group_name: str, image_path: Path, text_message: str = "") -> None:
        """Send image via WhatsApp Web by opening it in the default browser."""
        self.log(f"WhatsApp Web: attaching image for {group_name} (manual step required — "
                 "open web.whatsapp.com, find the group, and attach the image manually).")
        self.log(f"Image path: {image_path}")

    def send_text(self, group_name: str, text_message: str) -> None:
        if self.mode == "web":
            self._send_text_web(group_name, text_message)
            return
        message = safe_text(text_message)
        if not message:
            return
        pyautogui = self._import_pyautogui()
        self.log(f"Opening WhatsApp Desktop for {group_name}...")
        self._open_whatsapp()
        wa_hwnd, wa_rect = self._save_whatsapp_rect()
        if not self._force_focus_whatsapp():
            raise RuntimeError("Could not bring WhatsApp Desktop to the front - nothing was typed.")
        self.log(f"Searching group: {group_name}")
        self._type_group_search(pyautogui, group_name)
        self.log("Sending WhatsApp text message...")
        self._paste_text(pyautogui, message)
        time.sleep(0.7)
        _wa_press("enter")
        time.sleep(1.0)
        self._restore_whatsapp_rect(wa_hwnd, wa_rect)
        self.log(f"Sent text to {group_name}")

    def _send_text_web(self, group_name: str, text_message: str) -> None:
        """Send text via WhatsApp Web in Edge at port 9227. Keeps tab open after send."""
        global _wa_fallback_opened
        message = safe_text(text_message)
        if not message:
            return
        from selenium.webdriver.common.by import By
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        try:
            driver = _edge_debug_driver()
            _find_or_open_tab(driver, _WA_URL)
            wait = WebDriverWait(driver, 30)
            # Find search box (WA Web uses contenteditable divs; try multiple selectors)
            search_box = None
            for by, sel in [
                (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='3']"),
                (By.XPATH, "//div[@contenteditable='true'][@title='Search input textbox']"),
                (By.CSS_SELECTOR, "div[data-testid='chat-list-search']"),
            ]:
                try:
                    search_box = wait.until(EC.element_to_be_clickable((by, sel)))
                    break
                except Exception:
                    continue
            if search_box is None:
                self.log(f"WhatsApp Web: could not find search box for {group_name!r}")
                return
            search_box.click()
            search_box.send_keys(Keys.CONTROL + "a")
            search_box.send_keys(Keys.DELETE)
            search_box.send_keys(group_name)
            time.sleep(2)
            # Click first result whose title contains group_name
            clicked = False
            for by, sel in [
                (By.XPATH, f"//span[contains(@title, '{group_name}')]"),
                (By.CSS_SELECTOR, "div[data-testid='cell-frame-container']"),
            ]:
                try:
                    result = wait.until(EC.element_to_be_clickable((by, sel)))
                    result.click()
                    clicked = True
                    break
                except Exception:
                    continue
            if not clicked:
                self.log(f"WhatsApp Web: group {group_name!r} not found in results")
                return
            time.sleep(1)
            # Find message input box
            msg_box = None
            for by, sel in [
                (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']"),
                (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='6']"),
                (By.XPATH, "//div[@contenteditable='true'][@title='Type a message']"),
            ]:
                try:
                    msg_box = wait.until(EC.element_to_be_clickable((by, sel)))
                    break
                except Exception:
                    continue
            if msg_box is None:
                self.log(f"WhatsApp Web: could not find message input for {group_name!r}")
                return
            msg_box.click()
            import pyperclip
            pyperclip.copy(message)
            msg_box.send_keys(Keys.CONTROL + "v")
            time.sleep(0.5)
            msg_box.send_keys(Keys.ENTER)
            time.sleep(0.5)
            self.log(f"✓ WhatsApp Web: sent to {group_name!r}")
            # Tab stays open — do not close driver or switch away
        except Exception as e:
            self.log(f"WhatsApp Web send error ({group_name!r}): {e}")
            if not _wa_fallback_opened:
                import webbrowser
                webbrowser.open(_WA_URL)
                _wa_fallback_opened = True






def show_startup_error(exc: BaseException) -> None:
    error_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        (APP_DIR / "startup_error.log").write_text(error_text, encoding="utf-8")
    except Exception:
        try:
            Path(tempfile.gettempdir(), "gfh_inventory_audit_startup_error.log").write_text(error_text, encoding="utf-8")
        except Exception:
            pass
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        _app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "GFH Inventory Audit Error", str(exc))
    except Exception:
        print("GFH Inventory Audit Error:", exc)
        print(error_text)
GFH_SQUARE_ICON_B64 = _safe_read_asset("gfh_square_icon_b64.txt")


# PortalCredentialStore — credentials are now stored directly in the SQLite DB
# via VarianceDatabase.save_portal_credentials / load_portal_credentials.
# This stub is kept only so any legacy references don't break at import time.


# ─────────────────────────────────────────────────────────────────────────────
# Human / Cloudflare / reCAPTCHA verification helpers
# (ported from VidaPay workflow — used by B2BSoftScraper)
# ─────────────────────────────────────────────────────────────────────────────
def _b2b_is_human_verification_page(driver) -> bool:
    """Return True if Cloudflare or reCAPTCHA challenge is currently shown."""
    try:
        return bool(driver.execute_script("""
            const t = (document.body.innerText || '').toLowerCase();
            if (t.includes('verify you are human') || t.includes('verify human')) return true;
            if (t.includes('performing security verification'))                   return true;
            if (t.includes('just a moment') &&
                document.title.toLowerCase().includes('just a moment'))           return true;
            if (document.querySelector('#BbLB6'))                                 return true;
            if (document.querySelector('#challenge-stage'))                       return true;
            if (document.querySelector('.cf-turnstile'))                          return true;
            if (document.querySelector('[name="cf-turnstile-response"]'))         return true;
            if (document.querySelector('input[id*="cf-chl-widget"]'))             return true;
            if (document.querySelector('iframe[src*="challenges.cloudflare.com"]')) return true;
            if (document.querySelector('iframe[src*="recaptcha/api2/anchor"]') ||
                document.querySelector('iframe[src*="recaptcha/enterprise/anchor"]')) return true;
            return false;
        """))
    except Exception:
        return False


# >>> MOUSE-FREE v2 PATCH START >>>
# =============================================================
# MOUSE-FREE CHALLENGE CLICK v2 + HUMAN TYPING  (injected patch)
#
# The previous trusted click was dispatched on the PAGE DevTools session.
# Cloudflare's Turnstile iframe is a cross-origin, out-of-process frame
# (OOPIF) and page-session input events are not reliably routed into it -
# the click dispatched "successfully" while the checkbox never registered.
#
# This patch attaches directly to the CHALLENGE FRAME's own DevTools target
# and dispatches the humanized click inside the frame's own coordinate
# space, where it is always trusted and always lands.
#
# Result: no OS mouse movement, no window focus, no foreground needed.
# You keep working on the PC while the bot runs.
# =============================================================
ALLOW_REAL_MOUSE_FALLBACK = False   # keep False: your physical mouse is NEVER touched


def _human_send_keys(element, text):
    """Per-character typing with randomized inter-key delays.
    Whole-string send_keys is a paste-burst fingerprint (every keystroke
    lands in one instant event burst)."""
    import random as _rnd
    try:
        try:
            element.clear()
        except Exception:
            pass
        time.sleep(_rnd.uniform(0.2, 0.5))
        for _ch in str(text):
            element.send_keys(_ch)
            _pause = _rnd.uniform(0.05, 0.16)
            if _rnd.random() < 0.12:
                _pause = _rnd.uniform(0.2, 0.45)
            time.sleep(_pause)
        return True
    except Exception:
        try:
            element.send_keys(text)
        except Exception:
            pass
        return False


def _mpf_turnstile_token_ok(driver):
    """True when a Turnstile response token exists in the main DOM."""
    try:
        return bool(driver.execute_script(
            "const i = document.querySelector('[name=\"cf-turnstile-response\"]');"
            "return !!(i && i.value);"))
    except Exception:
        return False


def _mpf_http_json(url, timeout=5):
    import json as _json
    from urllib.request import urlopen as _urlopen
    with _urlopen(url, timeout=timeout) as _resp:
        return _json.loads(_resp.read().decode("utf-8", "ignore"))


def _mpf_cdp_call(ws, method, params=None, session_id=None, timeout=8,
                  _state={"id": 0}):
    """One CDP command over the raw browser websocket. Returns the response
    message (dict) or None on timeout/error."""
    import json as _json
    import time as _time
    _state["id"] = _state.get("id", 0) + 1
    mid = _state["id"]
    payload = {"id": mid, "method": method, "params": params or {}}
    if session_id:
        payload["sessionId"] = session_id
    try:
        ws.settimeout(timeout)
        ws.send(_json.dumps(payload))
        deadline = _time.time() + timeout
        while _time.time() < deadline:
            raw = ws.recv()
            if not raw:
                continue
            msg = _json.loads(raw)
            if msg.get("id") == mid:
                return msg
        return None
    except Exception:
        return None


def _mpf_browser_endpoint(driver, log=print):
    """Locate the browser-level DevTools HTTP endpoint for THIS driver."""
    # 1. Selenium capability - the exact browser behind this driver
    try:
        da = ((driver.capabilities or {}).get("goog:chromeOptions")
              or {}).get("debuggerAddress") or ""
        host, _, port = da.partition(":")
        if port and port != "0":
            return f"http://127.0.0.1:{int(port)}"
    except Exception:
        pass
    # 2. DevToolsActivePort files (Edge / Chrome / automation profile)
    import os as _os
    candidates = []
    la = _os.environ.get("LOCALAPPDATA")
    if la:
        candidates += [
            _os.path.join(la, "Microsoft", "Edge", "User Data",
                          "DevToolsActivePort"),
            _os.path.join(la, "Google", "Chrome", "User Data",
                          "DevToolsActivePort"),
        ]
    try:
        candidates.append(_os.path.join(AUTOMATION_PROFILE_DIR,
                                        "DevToolsActivePort"))
    except Exception:
        pass
    for path in candidates:
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                port = int((fh.readline() or "").strip())
            if 0 < port < 65536:
                return f"http://127.0.0.1:{port}"
        except Exception:
            continue
    # 3. this app's fixed debug port constant, if defined
    try:
        port = int(REMOTE_DEBUGGING_PORT)
        if 0 < port < 65536:
            return f"http://127.0.0.1:{port}"
    except Exception:
        pass
    return None


def _cdp_frame_checkbox_click(driver, log=print,
                              frame_markers=("challenges.cloudflare.com",),
                              verify=None, attempts=3):
    """THE MOUSE-FREE SOLVER.
    Attach to the cross-origin challenge iframe's own DevTools target and
    click its checkbox with humanized trusted input in the FRAME's own
    coordinate space. verify() is polled after every attempt; returns True
    only when the challenge actually cleared."""
    import random as _rnd
    import time as _time

    # websocket-client (auto-install once)
    try:
        import websocket as _ws_mod  # noqa
    except Exception:
        try:
            import subprocess as _sub
            import sys as _sys
            _sub.run([_sys.executable, "-m", "pip", "install",
                      "websocket-client", "--quiet",
                      "--disable-pip-version-check"],
                     capture_output=True, timeout=120)
            import websocket as _ws_mod  # noqa
        except Exception:
            log("  [mouse-free] websocket-client unavailable - frame click skipped.")
            return False

    endpoint = _mpf_browser_endpoint(driver, log=log)
    if not endpoint:
        log("  [mouse-free] no DevTools endpoint - frame click skipped.")
        return False

    ws_url = None
    try:
        info = _mpf_http_json(endpoint + "/json/version")
        ws_url = (info or {}).get("webSocketDebuggerUrl")
    except Exception:
        ws_url = None
    if not ws_url:
        log("  [mouse-free] DevTools websocket not reachable - frame click skipped.")
        return False

    try:
        ws = _ws_mod.create_connection(ws_url, timeout=8, suppress_origin=True)
    except Exception as exc:
        log(f"  [mouse-free] websocket connect failed: {exc}")
        return False

    def _checkbox_pos(session_id):
        """Real checkbox rect INSIDE the frame (frame coordinate space)."""
        js = (
            "(() => { const el = document.querySelector("
            "\"input[type='checkbox'], [role='checkbox']\");"
            " const r = el ? el.getBoundingClientRect() : null;"
            " return r ? {x: r.x + Math.min(30, r.width / 2),"
            " y: r.y + r.height / 2, w: window.innerWidth,"
            " h: window.innerHeight} : null; })()"
        )
        ev = _mpf_cdp_call(ws, "Runtime.evaluate",
                           {"expression": js, "returnByValue": True},
                           session_id=session_id, timeout=6)
        try:
            return (((ev or {}).get("result") or {}).get("result") or {}).get("value")
        except Exception:
            return None

    def _human_click_at(session_id, tx, ty, vw, vh):
        """Curved humanized approach + press/release inside the frame."""
        def _fire(etype, x, y, button="none", buttons=0, count=0):
            _mpf_cdp_call(ws, "Input.dispatchMouseEvent",
                          {"type": etype, "x": float(x), "y": float(y),
                           "button": button, "buttons": buttons,
                           "clickCount": count},
                          session_id=session_id, timeout=5)

        tx = min(max(float(tx), 3.0), max(4.0, vw - 3.0))
        ty = min(max(float(ty), 3.0), max(4.0, vh - 3.0))
        sx = _rnd.uniform(0.15, 0.6) * vw
        sy = _rnd.uniform(0.2, 0.8) * vh
        if abs(sx - tx) < 40 and abs(sy - ty) < 20:
            sx = (sx + vw * 0.5) % max(4.0, vw)
            sy = (sy + vh * 0.5) % max(4.0, vh)
        cx = (sx + tx) / 2.0 + _rnd.uniform(-24, 24)
        cy = (sy + ty) / 2.0 + _rnd.uniform(-14, 14)
        steps = _rnd.randint(10, 18)

        def _bez(t):
            u = 1.0 - t
            return (u * u * sx + 2 * u * t * cx + t * t * tx,
                    u * u * sy + 2 * u * t * cy + t * t * ty)

        _fire("mouseMoved", sx, sy)
        _time.sleep(_rnd.uniform(0.05, 0.14))
        for i in range(1, steps + 1):
            t = 1 - (1 - i / steps) ** 2
            x, y = _bez(t)
            if i == steps:
                x, y = tx + _rnd.uniform(-1.2, 1.2), ty + _rnd.uniform(-1.0, 1.0)
            _fire("mouseMoved", x, y)
            _time.sleep(_rnd.uniform(0.008, 0.026))
        _time.sleep(_rnd.uniform(0.2, 0.55))  # hover before pressing
        px = min(max(tx + _rnd.uniform(-0.8, 0.8), 2.0), max(3.0, vw - 2.0))
        py = min(max(ty + _rnd.uniform(-0.8, 0.8), 2.0), max(3.0, vh - 2.0))
        _fire("mousePressed", px, py, button="left", buttons=1, count=1)
        _time.sleep(_rnd.uniform(0.05, 0.13))
        _fire("mouseReleased", px, py, button="left", buttons=0, count=1)
        return True

    ok = False
    try:
        # the challenge frame often renders a beat after the page
        targets = []
        for _try in range(4):
            res = _mpf_cdp_call(ws, "Target.getTargets", {}, timeout=6)
            infos = ((res or {}).get("result") or {}).get("targetInfos") or []
            targets = [t for t in infos
                       if t.get("type") == "iframe"
                       and any(m in (t.get("url") or "") for m in frame_markers)]
            if targets:
                break
            _time.sleep(1.5)
        if not targets:
            log("  [mouse-free] no cross-origin challenge frame found in this tab.")
            return False

        for tgt in targets[:2]:
            att = _mpf_cdp_call(ws, "Target.attachToTarget",
                                {"targetId": tgt["targetId"], "flatten": True},
                                timeout=6)
            sid = ((att or {}).get("result") or {}).get("sessionId")
            if not sid:
                continue
            _time.sleep(_rnd.uniform(0.4, 0.9))
            for attempt in range(1, attempts + 1):
                pos = _checkbox_pos(sid) or {"x": 28.0, "y": 32.0,
                                             "w": 300, "h": 65}
                _human_click_at(sid, pos["x"], pos["y"],
                                pos.get("w", 300), pos.get("h", 65))
                deadline = _time.time() + 8.0
                while _time.time() < deadline:
                    _time.sleep(0.5)
                    if verify is not None:
                        try:
                            if verify():
                                log("  [mouse-free] challenge CLEARED via "
                                    f"frame-target click (attempt {attempt}) - "
                                    "OS mouse never touched.")
                                ok = True
                                break
                        except Exception:
                            pass
                if ok:
                    break
                log(f"  [mouse-free] frame click attempt {attempt}/{attempts} "
                    "dispatched - challenge not confirmed yet.")
                _time.sleep(_rnd.uniform(1.0, 2.0))
            try:
                _mpf_cdp_call(ws, "Target.detachFromTarget",
                              {"sessionId": sid}, timeout=4)
            except Exception:
                pass
            if ok:
                break
    except Exception as exc:
        log(f"  [mouse-free] frame-target click error: {exc}")
        ok = False
    finally:
        try:
            ws.close()
        except Exception:
            pass
    return ok
# <<< MOUSE-FREE v2 PATCH END <<<


def _human_cdp_click(driver, vp_x, vp_y, log=print, label="target"):
    """Humanized trusted left-click at VIEWPORT coords via CDP Input events.

    Why: the old single move->press->release burst at fixed 60/90ms
    intervals is exactly the rhythm bot-fingerprinting flags, even with
    every event isTrusted=true. A real pointer travels a curved path with
    variable speed, jitters +-2px, hovers briefly, then presses for
    60-150ms. This reproduces all of that while staying screen-free
    (no OS cursor, no focus steal, any monitor, window just not minimized)."""
    import random as _rnd
    try:
        vp_w = int(driver.execute_script("return window.innerWidth") or 1280)
        vp_h = int(driver.execute_script("return window.innerHeight") or 800)
    except Exception:
        vp_w, vp_h = 1280, 800

    def _fire(etype, x, y, button="none", buttons=0, count=0):
        driver.execute_cdp_cmd("Input.dispatchMouseEvent", {
            "type": etype, "x": float(x), "y": float(y),
            "button": button, "buttons": buttons, "clickCount": count,
        })

    try:
        # Clamp into the viewport - clicks outside it are silent no-ops.
        tx = min(max(float(vp_x), 2.0), vp_w - 2.0)
        ty = min(max(float(vp_y), 2.0), vp_h - 2.0)

        # Start the pointer somewhere else, like a real session would.
        sx = _rnd.uniform(0.15, 0.85) * vp_w
        sy = _rnd.uniform(0.15, 0.85) * vp_h
        if abs(sx - tx) < 140 and abs(sy - ty) < 140:
            sx = (sx + vp_w * 0.5) % vp_w
            sy = (sy + vp_h * 0.5) % vp_h

        # Quadratic Bezier control point -> natural curved approach.
        cx = (sx + tx) / 2.0 + _rnd.uniform(-90, 90)
        cy = (sy + ty) / 2.0 + _rnd.uniform(-70, 70)

        def _bez(t):
            u = 1.0 - t
            return (u * u * sx + 2 * u * t * cx + t * t * tx,
                    u * u * sy + 2 * u * t * cy + t * t * ty)

        _fire("mouseMoved", sx, sy)
        time.sleep(_rnd.uniform(0.04, 0.11))

        steps = _rnd.randint(22, 38)
        prev = (sx, sy)
        for i in range(1, steps + 1):
            t = 1 - (1 - i / steps) ** 2  # ease-out: fast start, slow finish
            x, y = _bez(t)
            if i == steps:  # land on target with a +-1.5px human jitter
                x, y = tx + _rnd.uniform(-1.5, 1.5), ty + _rnd.uniform(-1.5, 1.5)
            if abs(x - prev[0]) < 0.5 and abs(y - prev[1]) < 0.5:
                continue
            _fire("mouseMoved", x, y)
            time.sleep(_rnd.uniform(0.006, 0.022))
            prev = (x, y)

        time.sleep(_rnd.uniform(0.25, 0.7))  # hover before pressing

        px = min(max(tx + _rnd.uniform(-1.0, 1.0), 1.0), vp_w - 1.0)
        py = min(max(ty + _rnd.uniform(-1.0, 1.0), 1.0), vp_h - 1.0)
        _fire("mousePressed", px, py, button="left", buttons=1, count=1)
        time.sleep(_rnd.uniform(0.06, 0.15))  # real press duration
        _fire("mouseReleased", px, py, button="left", buttons=0, count=1)
        return True
    except Exception as exc:
        log(f"CDP trusted click failed ({label}): {exc}")
        return False


# >>> HUMAN-ELEMENT-CLICK v2 PATCH START >>>
def _human_click_element(driver, element, log=print, label="button"):
    """Humanized trusted click on a Selenium element (MAIN document):
    scroll into view, aim at the element centre, dispatch the humanized
    CDP click (same engine as the challenge checkbox - curved approach,
    hover, real press timing, isTrusted=true). Falls back to the
    element's own .click() when the element lives inside a cross-origin
    frame (Selenium routes those correctly) or the CDP path fails."""
    try:
        in_top = driver.execute_script("return window.top === window.self;")
    except Exception:
        in_top = True
    if in_top:
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", element)
            time.sleep(0.3)
            r = driver.execute_script(
                "const r = arguments[0].getBoundingClientRect();"
                "return {x: r.x + r.width / 2, y: r.y + r.height / 2};",
                element)
            if r and r.get("x") is not None and _cdp_trusted_click(
                    driver, r["x"], r["y"], log=log):
                return True
        except Exception as exc:
            log(f"  humanized element click fallback ({label}): {exc}")
    try:
        element.click()
        return True
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception:
            return False
# <<< HUMAN-ELEMENT-CLICK v2 PATCH END <<<


def _cdp_trusted_click(driver, vp_x, vp_y, log=print):
    """One trusted left-click at VIEWPORT coords via CDP Input events.

    Screen-free: no OS cursor movement, no focus change - works while the
    Edge window sits on any monitor or behind other windows (it just must
    not be minimized). The events are isTrusted=true, which is what Google
    reCAPTCHA now requires (a JS .click() is isTrusted=false and ignored).

    Idea ref: chrome-devtools-mcp issue #1826."""
    return _human_cdp_click(driver, vp_x, vp_y, log=log,
                            label="trusted click")


def _cdp_click_registered(driver, log=print):
    """True when the challenge reacted to the last click. Signals, checked
    for up to ~5s (challenge frames render late):
      1. reCAPTCHA anchor checkbox aria-checked == "true" (solved outright)
      2. a visible reCAPTCHA challenge frame (bframe) - click registered,
         Google opened the image/audio challenge (caller keeps solving)
      3. a Turnstile response input carrying a token (Turnstile solved)
    """
    deadline = time.time() + 5.0
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            bf = driver.find_element(
                By.CSS_SELECTOR,
                "iframe[src*='recaptcha/api2/bframe'],"
                "iframe[src*='recaptcha/enterprise/bframe'],"
                "iframe[title*='recaptcha challenge'],"
                "iframe[title*='challenge expires']")
            if bf.is_displayed():
                log("  challenge frame opened - click registered.")
                return True
        except Exception:
            pass
        try:
            if driver.execute_script(
                    "const i = document.querySelector("
                    "'[name=\"cf-turnstile-response\"]');"
                    "return !!(i && i.value);"):
                return True
        except Exception:
            pass
        try:
            anchor = driver.find_element(
                By.CSS_SELECTOR,
                "iframe[src*='recaptcha/api2/anchor'],"
                "iframe[src*='recaptcha/enterprise/anchor'],"
                "iframe[title*='recaptcha'],"
                "iframe[title*='not a robot']")
            driver.switch_to.default_content()
            driver.switch_to.frame(anchor)
            checked = driver.execute_script(
                "const el = document.getElementById('recaptcha-anchor');"
                "return el ? el.getAttribute('aria-checked') : null;")
            if checked == "true":
                driver.switch_to.default_content()
                return True
        except Exception:
            pass
        finally:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
    return False


def _cdp_click_iframe_checkbox(driver, iframe_el, offset_x=28, log=print):
    """Trusted CDP click on the checkbox inside a cross-origin challenge
    iframe (reCAPTCHA anchor / Turnstile). The click point is computed
    from the iframe's viewport rect + offset, so it never needs the
    screen, the mouse, or JS. Returns False when it could not fire so
    the caller can fall back to the legacy clickers."""
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", iframe_el)
        time.sleep(0.4)
        rect = driver.execute_script(
            "const r = arguments[0].getBoundingClientRect();"
            "return {left: r.left, top: r.top, width: r.width,"
            "        height: r.height};", iframe_el)
        if not rect or not rect.get("width"):
            log("Challenge iframe has zero size - not rendered yet.")
            return False
        x = rect["left"] + offset_x
        y = rect["top"] + rect["height"] / 2.0
        log(f"CDP trusted click at viewport ({x:.0f},{y:.0f}) "
            f"(iframe {rect['width']:.0f}x{rect['height']:.0f}).")

        # HUMANIZED + VERIFIED: dispatch up to 3 humanized clicks and after
        # each one CHECK that the challenge actually reacted (checkbox
        # aria-checked flipped, challenge frame opened, or a Turnstile
        # token appeared). The old single blind click could dispatch
        # "successfully" and still be ignored by Google.
        import random as _rnd
        for attempt in range(3):
            if attempt:
                x = rect["left"] + offset_x + _rnd.uniform(-3, 3)
                y = rect["top"] + rect["height"] / 2.0 + _rnd.uniform(-2, 2)
                time.sleep(_rnd.uniform(0.8, 1.6))
            if not _cdp_trusted_click(driver, x, y, log=log):
                continue
            if _cdp_click_registered(driver, log=log):
                return True
            log(f"  trusted click {attempt + 1}/3 dispatched but the "
                f"checkbox did not register - retrying humanized.")
        return False
    except Exception as exc:
        log(f"CDP iframe checkbox click error: {exc}")
        return False


def _discover_devtools_port(log=print):
    """Zero-config CDP port discovery (chrome-devtools-mcp#1826).
    Chrome/Edge write DevToolsActivePort (line 1 = port) whenever remote
    debugging is on - including the chrome://inspect remote-debugging
    toggle on the user's normal browser. Returns the first port that
    answers, or None."""
    candidates = []
    la = os.environ.get("LOCALAPPDATA")
    if la:
        candidates += [
            os.path.join(la, "Microsoft", "Edge", "User Data",
                         "DevToolsActivePort"),
            os.path.join(la, "Google", "Chrome", "User Data",
                         "DevToolsActivePort"),
        ]
    try:
        candidates.append(os.path.join(GFH_AUTOMATION_PROFILE_DIR,
                                       "DevToolsActivePort"))
    except Exception:
        pass
    for path in candidates:
        try:
            with open(path, "r", encoding="utf-8",
                      errors="ignore") as fh:
                port = int((fh.readline() or "").strip())
        except Exception:
            continue
        if not 0 < port < 65536:
            continue
        try:
            with socket.create_connection(("127.0.0.1", port),
                                          timeout=1):
                log(f"CDP port {port} discovered via {path}")
                return port
        except Exception:
            continue
    return None


def _b2b_cdp_click_turnstile(driver, log=print) -> bool:
    """
    Solve the Cloudflare Turnstile widget using a CDP trusted click.

    Why this replaced the old pyautogui approach:
      - pyautogui moves the REAL OS mouse and needs Edge in the FOREGROUND
        (SetForegroundWindow), so it hijacks the user's screen. That breaks
        the "run automation on one screen, work on another" workflow.
      - CDP Input.dispatchMouseEvent synthesizes trusted (isTrusted=true)
        browser-level input at VIEWPORT coordinates. From the page's point
        of view it is indistinguishable from a real click, but it never
        touches the OS mouse, never steals focus, and works regardless of
        which monitor Edge sits on or what DPI scaling is in play.

    Widget location (unchanged, proven trick): Cloudflare renders the
    Turnstile iframe inside a CLOSED shadow root, so querySelector can never
    see the iframe itself. But Cloudflare ALWAYS injects a hidden
    [name="cf-turnstile-response"] input into the MAIN DOM; the parent chain
    of that input hosts the ~300x65 widget box, and getBoundingClientRect()
    on it yields the exact viewport coordinates for the CDP click.

    Success signal: the response token filling in the MAIN DOM:
        document.querySelector('[name="cf-turnstile-response"]').value
    Non-empty token = Turnstile accepted the solve. Checked BEFORE clicking
    (managed / non-interactive challenges pass with no click at all) and
    polled after every click attempt.
    """

    def _token():
        try:
            return driver.execute_script("""
                const el = document.querySelector('[name="cf-turnstile-response"]');
                return el && el.value ? el.value : '';
            """) or ""
        except Exception:
            return ""

    def _cleared():
        try:
            still = driver.execute_script("""
                const t = (document.body.innerText || '').toLowerCase();
                if (t.includes('verify you are human') || t.includes('verify human')) return true;
                if (t.includes('performing security verification')) return true;
                if (t.includes('just a moment')) return true;
                return false;
            """)
            return not still
        except Exception:
            return False

    def _cdp_click(x, y):
        """Humanized trusted click at viewport (x, y) via CDP input
        events (Bezier approach + jitter + real press timing)."""
        return _human_cdp_click(driver, x, y, log=log, label="turnstile")

    def _actions_click(el, x_off, w):
        """Fallback trusted click via Selenium ActionChains (also CDP-level,
        also completely screen-free)."""
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(driver).move_to_element_with_offset(
                el, int(x_off - w / 2), 0).click().perform()
            return True
        except Exception as e:
            log(f"  ActionChains click error: {e}")
            return False

    # -- 0. Already solved? (managed / non-interactive auto-pass) -------------
    tok = _token()
    if tok:
        log("  Turnstile token already present - no click needed.")
        return True

    time.sleep(2.0)  # let the widget finish rendering

    # -- 0.5 MOUSE-FREE frame-target click: attach to the cross-origin
    #        Turnstile iframe's own DevTools target and click INSIDE it
    #        (page-session input does not reliably reach OOPIF frames).
    #        No OS mouse, no focus, works while the PC is being used.
    try:
        if _cdp_frame_checkbox_click(
                driver, log=log,
                frame_markers=("challenges.cloudflare.com",),
                verify=lambda: bool(_token())):
            return True
    except Exception:
        pass

    # -- 1. Locate the widget box by geometry (closed-shadow safe) ------------
    rect = None
    for tick in range(10):
        try:
            rect = driver.execute_script("""
                const cfInput = document.querySelector(
                    '[name="cf-turnstile-response"], input[id*="cf-chl-widget"]');
                if (cfInput) {
                    let el = cfInput.parentElement;
                    for (let i = 0; i < 6 && el; i++) {
                        const r = el.getBoundingClientRect();
                        if (r.width >= 250 && r.width <= 380 &&
                            r.height >= 45 && r.height <= 110 &&
                            (r.top > 0 || r.left > 0)) {
                            return {left:r.left, top:r.top,
                                    width:r.width, height:r.height,
                                    source:'cf-input-parent'};
                        }
                        el = el.parentElement;
                    }
                }
                const all = document.querySelectorAll('div');
                for (const el of all) {
                    const r = el.getBoundingClientRect();
                    if (r.width >= 280 && r.width <= 320 && r.height >= 55 && r.height <= 75) {
                        const txt = (el.innerText || '').trim().toLowerCase();
                        if (txt.includes('verify you are human') || txt.includes('verify')) {
                            return {left:r.left, top:r.top, width:r.width, height:r.height};
                        }
                    }
                }
                return null;
            """)
            if rect:
                log(f"  Turnstile widget located: "
                    f"{int(rect['width'])}x{int(rect['height'])}px.")
                break
            # Bail out early when this is clearly NOT a Turnstile page, so the
            # reCAPTCHA branch is not starved.
            if tick >= 4:
                has_cf = driver.execute_script("""
                    return !!(document.querySelector('[name="cf-turnstile-response"]') ||
                              document.querySelector('input[id*="cf-chl-widget"]') ||
                              document.querySelector('.cf-turnstile'));
                """)
                if not has_cf:
                    log("  No Turnstile widget/response input found - not a Turnstile page.")
                    return False
        except Exception:
            pass
        time.sleep(1)

    if not rect:
        log("  Could not locate the Turnstile widget box by geometry scan.")
        return False

    # -- 2. Scroll the widget into view and re-read its rect ------------------
    try:
        driver.execute_script("""
            const cfInput = document.querySelector('[name="cf-turnstile-response"]');
            let el = null;
            if (cfInput) {
                el = cfInput.parentElement;
                for (let i = 0; i < 6 && el; i++) {
                    const r = el.getBoundingClientRect();
                    if (r.width >= 250 && r.width <= 380 && r.height >= 45 && r.height <= 110) break;
                    el = el.parentElement;
                }
            }
            if (!el) {
                const all = document.querySelectorAll('div');
                for (const c of all) {
                    const r = c.getBoundingClientRect();
                    const txt = (c.innerText || '').trim().toLowerCase();
                    if (r.width >= 280 && r.width <= 320 && r.height >= 55 && r.height <= 75 &&
                        txt.includes('verify')) { el = c; break; }
                }
            }
            if (el) el.scrollIntoView({block:'center', inline:'center'});
        """)
        time.sleep(0.8)
        rect2 = driver.execute_script("""
            const cfInput = document.querySelector('[name="cf-turnstile-response"]');
            if (cfInput) {
                let el = cfInput.parentElement;
                for (let i = 0; i < 6 && el; i++) {
                    const r = el.getBoundingClientRect();
                    if (r.width >= 250 && r.width <= 380 && r.height >= 45 && r.height <= 110 &&
                        r.top > 0 && r.left > 0) {
                        return {left:r.left, top:r.top, width:r.width, height:r.height};
                    }
                    el = el.parentElement;
                }
            }
            return null;
        """)
        if rect2:
            rect = rect2
    except Exception:
        pass

    time.sleep(1.0)  # let Cloudflare's JS listeners fully bind

    # -- 3. Click attempts across the checkbox zone ---------------------------
    w = rect["width"]
    attempts = [(x_off, "checkbox") for x_off in (22, 16, 30, 12, 40)]
    attempts.append((w / 2, "centre"))

    for x_off, label in attempts:
        if _token():
            log("  Turnstile token present - solved.")
            return True
        cx = rect["left"] + x_off
        cy = rect["top"] + rect["height"] / 2
        log(f"  CDP trusted click at viewport ({int(cx)},{int(cy)}) [{label}]...")
        clicked = _cdp_click(cx, cy)
        if not clicked:
            # Fallback: ActionChains on the wrapper element (still screen-free)
            try:
                el = driver.execute_script("""
                    const cfInput = document.querySelector('[name="cf-turnstile-response"]');
                    if (cfInput) {
                        let p = cfInput.parentElement;
                        for (let i = 0; i < 6 && p; i++) {
                            const r = p.getBoundingClientRect();
                            if (r.width >= 250 && r.width <= 380 && r.height >= 45 && r.height <= 110)
                                return p;
                            p = p.parentElement;
                        }
                    }
                    return null;
                """)
                if el:
                    clicked = _actions_click(el, x_off, rect["width"])
            except Exception:
                pass
        # Poll for the token - the definitive success signal (up to 8s)
        for _w in range(16):
            time.sleep(0.5)
            if _token():
                log("  Turnstile token captured - verified CLEAR.")
                return True
            if _w == 5 and _cleared():
                log("  Challenge text gone - treated as cleared.")
                return True

    log("  CDP Turnstile click did not yield a token - falling through to human wait.")
    return False

def _b2b_try_solve_recaptcha(driver, log=print) -> bool:
    """Solve reCAPTCHA v2 via Google STT audio challenge (with Whisper fallback)."""
    import shutil, subprocess, tempfile
    try:
        import speech_recognition as sr
    except ImportError:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "SpeechRecognition",
                            "--quiet", "--disable-pip-version-check"], capture_output=True, timeout=90)
            import speech_recognition as sr
        except Exception:
            log("  SpeechRecognition not available — cannot solve reCAPTCHA audio.")
            return False

    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        log("  ffmpeg not found — cannot convert audio for reCAPTCHA.")
        return False

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    try:
        driver.switch_to.default_content()

        # Find anchor iframe
        anchor = None
        for sel in ["iframe[src*='recaptcha/api2/anchor']",
                    "iframe[src*='recaptcha/enterprise/anchor']",
                    "iframe[title*='reCAPTCHA']"]:
            try:
                anchor = driver.find_element(By.CSS_SELECTOR, sel)
                break
            except Exception:
                pass
        if not anchor:
            log("  reCAPTCHA anchor iframe not found.")
            return False

        _cb_trusted = False
        try:
            # TRUSTED click on the checkbox, straight through the anchor
            # iframe (screen-free). Google ignores JS .click() events
            # (isTrusted=false), so this is tried BEFORE the legacy
            # element-click path below.
            _cb_trusted = _cdp_click_iframe_checkbox(driver, anchor,
                                                     offset_x=28, log=log)
            if _cb_trusted:
                log("  reCAPTCHA checkbox clicked (CDP trusted - screen-free).")
                time.sleep(2)
        except Exception as _cdp_exc:
            log(f"  CDP trusted click unavailable ({_cdp_exc}) - "
                f"falling back to the element click.")
        if not _cb_trusted:
            driver.switch_to.frame(anchor)
            try:
                cb = driver.find_element(By.ID, "recaptcha-anchor")
                if driver.execute_script(
                        "return arguments[0].getAttribute('aria-checked');",
                        cb) != "true":
                    cb.click()
                    log("  reCAPTCHA checkbox clicked.")
                    time.sleep(2)
            except Exception:
                pass
        driver.switch_to.default_content()
        time.sleep(1.5)

        # Find bframe
        bframe = None
        for sel in ["iframe[src*='recaptcha/api2/bframe']",
                    "iframe[src*='recaptcha/enterprise/bframe']",
                    "iframe[title*='recaptcha challenge']"]:
            try:
                bframe = driver.find_element(By.CSS_SELECTOR, sel)
                break
            except Exception:
                pass
        if not bframe:
            log("  reCAPTCHA bframe not found.")
            return False

        driver.switch_to.frame(bframe)
        try:
            WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.ID, "recaptcha-audio-button"))
            ).click()
            log("  Audio button clicked.")
            time.sleep(2)
        except Exception as e:
            log(f"  Audio button error: {e}")
            driver.switch_to.default_content()
            return False

        # Audio challenge loop (up to 3 cycles)
        BFRAME_SELS = ["iframe[src*='recaptcha/api2/bframe']",
                       "iframe[src*='recaptcha/enterprise/bframe']",
                       "iframe[title*='recaptcha challenge']"]

        def _find_bframe():
            driver.switch_to.default_content()
            for sel in BFRAME_SELS:
                try:
                    return driver.find_element(By.CSS_SELECTOR, sel)
                except Exception:
                    pass
            return None

        submitted = False
        for cycle in range(3):
            if cycle > 0:
                _bf = _find_bframe()
                if not _bf:
                    break
                driver.switch_to.frame(_bf)
                try:
                    WebDriverWait(driver, 6).until(
                        EC.element_to_be_clickable((By.ID, "recaptcha-reload-button"))
                    ).click()
                    log(f"  Audio reloaded (cycle {cycle+1}).")
                    time.sleep(2.5)
                except Exception:
                    driver.switch_to.default_content()
                    break
                driver.switch_to.default_content()

            _bf2 = _find_bframe()
            if not _bf2:
                break
            driver.switch_to.frame(_bf2)
            time.sleep(1)
            mp3_url = None
            try:
                mp3_url = driver.find_element(
                    By.CSS_SELECTOR, "a.rc-audiochallenge-tdownload-link"
                ).get_attribute("href")
            except Exception:
                pass
            if not mp3_url:
                try:
                    mp3_url = driver.find_element(
                        By.CSS_SELECTOR, "audio#audio-source, audio source"
                    ).get_attribute("src")
                except Exception:
                    pass
            driver.switch_to.default_content()
            if not mp3_url:
                continue

            transcript = None
            with tempfile.TemporaryDirectory() as tmp:
                import os as _os
                mp3 = _os.path.join(tmp, "c.mp3")
                wav = _os.path.join(tmp, "c.wav")
                try:
                    import urllib.request as _ur
                    _ur.urlretrieve(mp3_url, mp3)
                    subprocess.run([ffmpeg, "-y", "-i", mp3, "-ar", "16000", "-ac", "1", wav],
                                   capture_output=True, timeout=30)
                    rec = sr.Recognizer()
                    with sr.AudioFile(wav) as _src:
                        _audio = rec.record(_src)
                    transcript = rec.recognize_google(_audio)
                    log(f"  STT result: '{transcript}'")
                except Exception as e:
                    log(f"  STT failed (cycle {cycle+1}): {e}")

            if not transcript:
                continue

            _bf3 = _find_bframe()
            if not _bf3:
                break
            driver.switch_to.frame(_bf3)
            try:
                inp = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.ID, "audio-response"))
                )
                inp.clear()
                inp.send_keys(transcript.lower().strip())
                time.sleep(0.4)
                driver.find_element(By.ID, "recaptcha-verify-button").click()
                log(f"  reCAPTCHA answer submitted (cycle {cycle+1}).")
                time.sleep(2.5)
                submitted = True
            except Exception as e:
                log(f"  Answer submit error: {e}")
            driver.switch_to.default_content()
            if submitted:
                break

        driver.switch_to.default_content()
        return submitted

    except Exception as e:
        log(f"  reCAPTCHA solver error: {e}")
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return False


def _b2b_try_auto_click_human_verification(driver, log=print) -> bool:
    """Detect verification type (Cloudflare vs reCAPTCHA) and dispatch."""
    from selenium.webdriver.common.by import By
    is_cf = False
    try:
        is_cf = bool(driver.execute_script("""
            if (document.querySelector('[name="cf-turnstile-response"]'))           return true;
            if (document.querySelector('input[id*="cf-chl-widget"]'))               return true;
            if (document.querySelector('.cf-turnstile'))                            return true;
            if (document.querySelector('iframe[src*="challenges.cloudflare.com"]')) return true;
            if (document.querySelector('#challenge-stage'))                         return true;
            const t = (document.body.innerText || '').toLowerCase();
            if (t.includes('performing security verification'))                     return true;
            if (t.includes('just a moment') &&
                document.title.toLowerCase().includes('just a moment'))             return true;
            return false;
        """))
    except Exception:
        pass

    if is_cf:
        log("Cloudflare Turnstile detected — CDP trusted click (screen-free).")
        return _b2b_cdp_click_turnstile(driver, log=log)

    # reCAPTCHA
    for sel in ["iframe[src*='recaptcha/api2/anchor']",
                "iframe[src*='recaptcha/enterprise/anchor']",
                "iframe[title*='reCAPTCHA']"]:
        try:
            driver.find_element(By.CSS_SELECTOR, sel)
            log("reCAPTCHA detected — running audio solver.")
            return _b2b_try_solve_recaptcha(driver, log=log)
        except Exception:
            pass

    return False


def _b2b_wait_for_human_verification_clear(driver, stop_event=None, timeout: int = 120, log=print) -> bool:
    """Wait for any Cloudflare/reCAPTCHA challenge to pass, auto-solving if possible."""
    if not _b2b_is_human_verification_page(driver):
        return True
    log("Human verification detected on B2B portal.")

    # Check for managed challenge (auto-resolves itself)
    try:
        is_managed = bool(driver.execute_script("""
            const t = (document.body.innerText || '').toLowerCase();
            return t.includes('performing security verification') ||
                   (t.includes('just a moment') && document.title.toLowerCase().includes('just a moment'));
        """))
    except Exception:
        is_managed = False

    if is_managed:
        log("  Managed challenge — waiting up to 15s for auto-resolve…")
        for _ in range(15):
            time.sleep(1)
            if not _b2b_is_human_verification_page(driver):
                log("  Managed challenge resolved. Continuing.")
                return True
        log("  Managed challenge still present — trying click strategies.")

    for cycle in range(3):
        if stop_event is not None and stop_event.is_set():
            return False
        log(f"  Verification cycle {cycle+1}/3…")
        if _b2b_try_auto_click_human_verification(driver, log=log):
            for _ in range(10):
                time.sleep(1)
                if not _b2b_is_human_verification_page(driver):
                    log("  Verification cleared.")
                    return True
            log(f"  Cycle {cycle+1}: still showing after click.")
        else:
            log(f"  Cycle {cycle+1}: no click delivered.")
        if cycle < 2:
            time.sleep(5)
            if not _b2b_is_human_verification_page(driver):
                return True
    log("Human verification did not clear automatically.")
    # -- Grace period: beep + wait for a manual solve instead of failing fast --
    # The automation may run on a second screen while the user works elsewhere;
    # give them a chance to solve the challenge by hand before giving up.
    try:
        import winsound
        for _b in range(3):
            winsound.Beep(1200, 300)
            time.sleep(0.25)
        log("  Beeped 3x - waiting up to 120s for a manual solve...")
    except Exception:
        log("  Waiting up to 120s for a manual solve...")
    _cleared_manually = False
    for _g in range(60):  # 60 x 2s = 120s
        if stop_event is not None and stop_event.is_set():
            return False
        time.sleep(2)
        if not _b2b_is_human_verification_page(driver):
            _cleared_manually = True
            break
        if _g in (7, 22, 37, 52):
            log("  Still waiting for manual solve...")
    if _cleared_manually:
        log("  Verification cleared manually. Continuing.")
        return True
    log("  No manual solve within 120s. Stopping this store.")
    return False


# ── B2B post-login page-state machine (mirrors VidaPay's finish_setup_steps) ──

_B2B_STATE_LOGIN          = "LOGIN"
_B2B_STATE_VERIFY         = "HUMAN_VERIFY"
_B2B_STATE_NEW_SIGN       = "NEW_SIGN_IN"
_B2B_STATE_TWO_FA         = "TWO_FACTOR"
_B2B_STATE_TRUST_DEVICE   = "TRUST_DEVICE"
_B2B_STATE_SETUP_NEXT     = "SETUP_NEXT"
_B2B_STATE_SECURITY_UPG   = "SECURITY_UPGRADE"
_B2B_STATE_READY_TO_GO    = "READY_TO_GO"
_B2B_STATE_PORTAL         = "PORTAL"
_B2B_STATE_UNKNOWN        = "UNKNOWN"


def _b2b_get_body(driver) -> str:
    from selenium.webdriver.common.by import By  # local import — By is not module-level
    try:
        return driver.find_element(By.TAG_NAME, "body").text.lower()
    except Exception:
        return ""


def _b2b_has_heading(driver, phrase: str) -> bool:
    """JS-based heading check — works before body.text populates (SPA render lag).
    Checks visible h1/h2/h3/h4 elements, same approach as VidaPay has_h3_heading."""
    phrase = phrase.lower().strip()
    try:
        texts = driver.execute_script(
            """
            return Array.from(document.querySelectorAll('h1,h2,h3,h4,legend')).map(function(el) {
                var style = window.getComputedStyle(el);
                var rect = el.getBoundingClientRect();
                var visible = style.display !== 'none' && style.visibility !== 'hidden'
                              && rect.width > 0 && rect.height > 0;
                return visible ? (el.innerText || el.textContent || '').trim().toLowerCase() : '';
            }).filter(Boolean);
            """
        )
        return any(phrase in str(t) for t in (texts or []))
    except Exception:
        return phrase in _b2b_get_body(driver)


def _b2b_get_page_state(driver) -> str:
    """Classify current B2B page.

    Detection order (most-specific first):
      1. URL keyword — immediate, works before SPA body text renders.
      2. JS heading check — visible h1-h4, faster than body.text on SPA.
      3. body.text fallback — plain text, slower but broadest coverage.

    Mirrors VidaPay get_page_state URL+heading strategy exactly.
    """
    # Cloudflare / reCAPTCHA — check first, URL might be whitelisted
    if _b2b_is_human_verification_page(driver):
        return _B2B_STATE_VERIFY

    url = ""
    try:
        url = driver.current_url.lower()
    except Exception:
        pass

    # ── URL-based detection (instant, no SPA render needed) ──────────────
    # B2BSoft uses same URL slug patterns as VidaPay portal.
    if "twofactornewdevicesignin" in url or "newdevicesignin" in url:
        return _B2B_STATE_NEW_SIGN
    if "twofactorcheck" in url or "twofactor/check" in url:
        return _B2B_STATE_TWO_FA
    if "twofactorupdatename" in url or "trustdevice" in url:
        return _B2B_STATE_TRUST_DEVICE
    if "twofactorready" in url or "readytogo" in url:
        return _B2B_STATE_READY_TO_GO
    if "secureupgradeoptions" in url or "securityupgrade" in url:
        return _B2B_STATE_SECURITY_UPG
    if "twofactorsetup" in url:
        return _B2B_STATE_SETUP_NEXT

    # ── Heading-based detection (works during SPA partial render) ─────────
    if _b2b_has_heading(driver, "new sign in"):
        return _B2B_STATE_NEW_SIGN
    if (_b2b_has_heading(driver, "2-factor authentication") or
            _b2b_has_heading(driver, "two-factor authentication") or
            _b2b_has_heading(driver, "verify your identity")):
        return _B2B_STATE_TWO_FA
    if _b2b_has_heading(driver, "trust this device") or _b2b_has_heading(driver, "trust device"):
        return _B2B_STATE_TRUST_DEVICE
    if _b2b_has_heading(driver, "ready to go") or _b2b_has_heading(driver, "you're all set"):
        return _B2B_STATE_READY_TO_GO
    if (_b2b_has_heading(driver, "upgrade security") or
            _b2b_has_heading(driver, "important: upgrade security")):
        return _B2B_STATE_SECURITY_UPG

    # Setup-Next pages sit between Trust Device and Ready To Go and carry a
    # #setupNextBtn button (VidaPay is_twofactor_setup_next_page): any
    # twofactor.* URL with a Next/setupNextBtn control needs an explicit click,
    # otherwise the flow stalls until timeout.
    if "twofactor" in url and _b2b_has_any_setup_next_button(driver):
        return _B2B_STATE_SETUP_NEXT

    # ── body.text fallback (slowest — SPA must have fully rendered) ───────
    body = _b2b_get_body(driver)

    if "new sign in" in body or "new sign-in" in body or "don't recognize this device" in body or "unrecognized device" in body:
        return _B2B_STATE_NEW_SIGN
    if any(k in body for k in ("2-factor", "two-factor", "authentication code",
                                "verify your identity", "enter the code", "otp")):
        return _B2B_STATE_TWO_FA
    if "trust this device" in body or "trust device" in body or "remember this device" in body:
        return _B2B_STATE_TRUST_DEVICE
    if "ready to go" in body or "you're all set" in body or "you are all set" in body:
        return _B2B_STATE_READY_TO_GO

    # ── Alerts page — portal landing showing the ALERTS (N) panel with
    # Later-dismiss buttons. The extractor treats ALERTS as a terminal
    # state equal to PORTAL; login/2FA pages never carry Later buttons. ──
    try:
        if "b2bsoft.com" in url and _b2b_find_later_buttons(driver) \
                and "alert" in _b2b_get_body(driver):
            return _B2B_STATE_PORTAL
    except Exception:
        pass

    # ── Portal check — on b2bsoft domain, no login fields, no 2FA text ───
    if "wsreports.b2bsoft.com" in url:
        try:
            has_login = any(
                driver.find_elements(By.ID, fid)
                for fid in ("companyId", "AccountId", "Username", "btnSubmit", "btnClick")
            )
        except Exception:
            has_login = False
        if not has_login:
            return _B2B_STATE_PORTAL

    # ── Still on login form ───────────────────────────────────────────────
    try:
        if driver.find_elements(By.ID, "companyId") or driver.find_elements(By.ID, "Username"):
            return _B2B_STATE_LOGIN
    except Exception:
        pass

    return _B2B_STATE_UNKNOWN


def _b2b_click_any_next(driver, log=print) -> bool:
    """Click any visible Next/Continue/Submit button — used for New Sign In and 2FA steps."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    for xpath in [
        "//button[contains(normalize-space(),'Next')]",
        "//button[contains(normalize-space(),'Continue')]",
        "//input[@type='submit']",
        "//button[@type='submit']",
    ]:
        try:
            btn = WebDriverWait(driver, 8).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            _human_click_element(driver, btn, log=log, label="login-next")
            log(f"Clicked button: {btn.text.strip() or xpath}")
            return True
        except Exception:
            continue
    return False


def _b2b_click_button(driver, label, text_contains=None, onclick_contains=None,
                      button_id=None, timeout: int = 25, log=print) -> bool:
    """Find a usable button by id/text/onclick and click it via JS.

    Faithful port of VidaPay click_matching_button_js: un-disables the control,
    scrolls it into view, clicks natively and falls back to a MouseEvent —
    needed because B2BSoft renders setup controls hidden/disabled in the DOM
    before they become interactive.
    """
    end_time = time.time() + timeout
    last_log = 0.0
    while time.time() < end_time:
        try:
            clicked = driver.execute_script(
                """
                const textContains = (arguments[0] || '').toLowerCase();
                const onclickContains = (arguments[1] || '').toLowerCase();
                const buttonId = arguments[2] || '';

                const elements = Array.from(document.querySelectorAll('button, input[type=button], input[type=submit], a'));

                function textOf(el) {
                    return (el.innerText || el.value || el.textContent || '').trim().toLowerCase();
                }

                function isUsable(el) {
                    if (!el) return false;
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return (
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.opacity !== '0' &&
                        rect.width > 0 &&
                        rect.height > 0 &&
                        el.getClientRects().length > 0 &&
                        !el.disabled &&
                        el.getAttribute('aria-disabled') !== 'true'
                    );
                }

                const btn = elements.find(el => {
                    const id = el.id || '';
                    const onclick = (el.getAttribute('onclick') || '').toLowerCase();
                    const text = textOf(el);
                    if (buttonId && id !== buttonId) return false;
                    if (onclickContains && !onclick.includes(onclickContains)) return false;
                    if (textContains && !text.includes(textContains)) return false;
                    return isUsable(el);
                });

                if (!btn) return false;

                btn.removeAttribute('disabled');
                btn.disabled = false;
                btn.removeAttribute('aria-disabled');
                btn.scrollIntoView({block: 'center'});

                try { btn.focus(); } catch (err) {}

                try {
                    btn.click();
                    return true;
                } catch (err) {
                    const event = new MouseEvent('click', {bubbles: true, cancelable: true, view: window});
                    btn.dispatchEvent(event);
                    return true;
                }
                """,
                text_contains,
                onclick_contains,
                button_id,
            )

            if clicked:
                log(f"[B2B] Clicked: {label}")
                time.sleep(1.2)
                return True
        except Exception:
            pass

        now = time.time()
        if now - last_log >= 5:
            log(f"[B2B] Waiting for {label}...")
            last_log = now

        time.sleep(0.25)

    log(f"[B2B] {label} was not found or not clickable within {timeout} seconds.")
    return False


def _b2b_has_any_setup_next_button(driver) -> bool:
    """True when a #setupNextBtn / visible Next control exists (VidaPay
    has_any_setup_next_button). Checks DOM presence first — B2BSoft renders
    the button before Selenium reports it visible."""
    from selenium.webdriver.common.by import By
    try:
        if driver.find_elements(By.ID, "setupNextBtn"):
            return True
        for btn in driver.find_elements(By.XPATH, "//button[normalize-space()='Next']"):
            try:
                if btn.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass
    try:
        return bool(driver.execute_script(
            """
            function visible(el) {
                if (!el) return false;
                const style = window.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                return (
                    style.display !== 'none' &&
                    style.visibility !== 'hidden' &&
                    style.opacity !== '0' &&
                    rect.width > 0 &&
                    rect.height > 0 &&
                    el.getClientRects().length > 0 &&
                    !el.disabled &&
                    el.getAttribute('aria-disabled') !== 'true'
                );
            }

            const byId = document.querySelector('button#setupNextBtn');
            if (visible(byId)) return true;

            return Array.from(document.querySelectorAll('button, input[type=button], input[type=submit]')).some(el => {
                const text = (el.innerText || el.value || '').trim().toLowerCase();
                return visible(el) && (text === 'next' || text.includes('next'));
            });
            """
        ))
    except Exception:
        return False


def _b2b_select_trust_radio(driver, stop_event=None, timeout: int = 180, log=print) -> bool:
    """Select the Trust Device radio (#trustRadio) as soon as it exists in the DOM.

    Port of VidaPay select_trust_radio_quickly: the Next button stays disabled
    until the radio is checked, and B2BSoft places the radio in the HTML before
    it renders visually — so force-check it via JS instead of waiting for
    visibility. Keeps waiting through 2FA approval and Cloudflare interludes.
    """
    deadline = time.time() + timeout
    last_log = 0.0
    while time.time() < deadline:
        if stop_event is not None and stop_event.is_set():
            return False
        try:
            selected = driver.execute_script(
                """
                const radio = document.querySelector('#trustRadio');
                if (!radio) return false;

                radio.removeAttribute('disabled');
                radio.disabled = false;
                radio.removeAttribute('aria-disabled');
                radio.checked = true;

                try { radio.scrollIntoView({block: 'center'}); } catch (err) {}
                try { radio.focus(); } catch (err) {}
                try { radio.click(); } catch (err) {}

                radio.dispatchEvent(new Event('input', { bubbles: true }));
                radio.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
                """
            )
            if selected:
                log("[B2B] Trust Device radio selected from DOM location.")
                time.sleep(0.2)
                return True
        except Exception:
            pass

        state = _b2b_get_page_state(driver)

        if state == _B2B_STATE_VERIFY:
            if not _b2b_wait_for_human_verification_clear(driver, stop_event=stop_event, log=log):
                return False
            continue

        if state in (_B2B_STATE_PORTAL, _B2B_STATE_READY_TO_GO,
                     _B2B_STATE_SECURITY_UPG, _B2B_STATE_SETUP_NEXT):
            log("[B2B] Already past Trust Device radio — continuing setup flow.")
            return True

        now = time.time()
        if now - last_log >= 8:
            if state == _B2B_STATE_TWO_FA:
                log("[B2B] Still on 2FA page — waiting for approval and Trust Device controls...")
            else:
                log(f"[B2B] Waiting for Trust Device radio… (page: {state})")
            last_log = now

        time.sleep(0.5)

    log("[B2B] Trust Device radio did not appear before timeout.")
    return False


def _b2b_wait_for_states(driver, stop_event=None, wanted=(), timeout: int = 90,
                         log=print, prefix="State wait"):
    """Poll until one of `wanted` states is reached (port of VidaPay wait_for_state)."""
    deadline = time.time() + timeout
    last_state = None
    while time.time() < deadline:
        if stop_event is not None and stop_event.is_set():
            return None
        state = _b2b_get_page_state(driver)
        if state != last_state:
            log(f"[B2B] {prefix}: {state}")
            last_state = state
        if state == _B2B_STATE_VERIFY:
            if not _b2b_wait_for_human_verification_clear(driver, stop_event=stop_event, log=log):
                return None
            last_state = None
            continue
        if state in wanted:
            return state
        time.sleep(0.5)
    log(f"[B2B] Timed out while waiting for: {', '.join(wanted)}")
    return None


def _b2b_find_later_buttons(driver):
    """All visible 'Later' alert-dismissal buttons (port of VidaPay finder)."""
    from selenium.webdriver.common.by import By
    xpaths = [
        "//button[normalize-space()='Later']",
        "//button[contains(normalize-space(),'Later')]",
        "//a[normalize-space()='Later']",
        "//a[contains(normalize-space(),'Later')]",
        "//*[@role='button' and normalize-space()='Later']",
        "//*[@role='button' and contains(normalize-space(),'Later')]",
        "//*[self::button or self::a or @role='button'][contains(normalize-space(), 'Later')]",
    ]
    seen, out = set(), []
    for xp in xpaths:
        try:
            for btn in driver.find_elements(By.XPATH, xp):
                try:
                    if btn.is_displayed() and btn.id not in seen:
                        seen.add(btn.id)
                        out.append(btn)
                except Exception:
                    continue
        except Exception:
            continue
    return out


def _b2b_clear_later_alerts(driver, log=print, max_clicks: int = 12) -> int:
    """Dismiss the portal's 'Later' popup alerts (port of VidaPay
    clear_all_later_alerts) so report navigation isn't blocked."""
    clicked, idle_rounds = 0, 0
    while clicked < max_clicks:
        buttons = _b2b_find_later_buttons(driver)
        if not buttons:
            idle_rounds += 1
            if idle_rounds >= 4:
                break
            # Alerts pop up one at a time — give the next one time to render
            # before concluding the panel is clear (user: there may be 1–4
            # Later buttons, and ALL of them must be cleared).
            time.sleep(1.5)
            continue
        idle_rounds = 0
        btn = buttons[0]
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.3)
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
            clicked += 1
            log(f"[B2B] Dismissed portal alert (Later) #{clicked}.")
            time.sleep(1.0)
        except Exception:
            break
    if clicked:
        log(f"[B2B] Later alerts cleared — clicked {clicked} button(s).")
    else:
        log("[B2B] No Later alerts found.")
    return clicked


def _b2b_complete_setup_next_flow(driver, stop_event=None, log=print,
                                  max_next_clicks: int = 6) -> bool:
    """Walk every remaining 2FA setup page until the portal loads.

    Port of VidaPay complete_remaining_setup_next_flow:
      Trust Device (radio variant, or the "Trust This Device" variant with
      only #setupNextBtn) → Setup Next pages → Ready To Go
      (vidapayAutomaticSignIn Continue) → portal + ALERTS panel (all Later
      buttons dismissed by the caller). The Security Upgrade page's Next must
      NOT be clicked — it throws error=io without a selected option — so we
      wait for it to advance on its own.
    """
    steps_done = 0
    while steps_done < max_next_clicks:
        if stop_event is not None and stop_event.is_set():
            return False

        state = _b2b_get_page_state(driver)
        log(f"[B2B] Setup flow page: {state}")

        if state == _B2B_STATE_VERIFY:
            if not _b2b_wait_for_human_verification_clear(driver, stop_event=stop_event, log=log):
                return False
            continue

        if state == _B2B_STATE_PORTAL:
            log("[B2B] Setup flow completed — portal loaded.")
            return True

        if state == _B2B_STATE_SECURITY_UPG:
            log("[B2B] Security Upgrade page — NOT clicking Next (errors without a selected option); waiting for it to advance…")
            wait_state = _b2b_wait_for_states(
                driver, stop_event=stop_event,
                wanted=(_B2B_STATE_READY_TO_GO, _B2B_STATE_PORTAL),
                timeout=20, log=log, prefix="Security upgrade wait")
            if wait_state == _B2B_STATE_READY_TO_GO:
                continue
            if wait_state == _B2B_STATE_PORTAL:
                return True
            log("[B2B] Security upgrade page did not advance safely — manual review needed.")
            return False

        if state == _B2B_STATE_TRUST_DEVICE:
            # Two page variants share this state:
            #   - radio variant: #trustRadio + #setupNextBtn — Next stays
            #     disabled until the radio is checked;
            #   - "Trust This Device" variant (user page):
            #     <h3>Trust This Device</h3> +
            #     <button id="setupNextBtn">Next</button> with NO radio —
            #     waiting 180s for one would stall the flow.
            # So: short non-fatal radio attempt, then always click Next.
            if not _b2b_select_trust_radio(driver, stop_event=stop_event,
                                           log=log, timeout=6):
                log("[B2B] No Trust Device radio appeared — clicking Next directly.")
            # User page button: #setupNextBtn with text Next (no onclick) —
            # id-first, onclick submitthisform as fallback.
            if not _b2b_click_button(driver, label="Trust Device Next",
                                     text_contains="next",
                                     button_id="setupNextBtn",
                                     timeout=30, log=log):
                if not _b2b_click_button(driver, label="Trust Device Next fallback",
                                         text_contains="next",
                                         onclick_contains="submitthisform",
                                         timeout=30, log=log):
                    return False
            steps_done += 1
            time.sleep(1)
            continue

        if state == _B2B_STATE_SETUP_NEXT:
            if not _b2b_click_button(driver, label="Setup Next",
                                     text_contains="next", button_id="setupNextBtn",
                                     timeout=30, log=log):
                state_after = _b2b_get_page_state(driver)
                if state_after == _B2B_STATE_PORTAL:
                    log("[B2B] Portal loaded before Setup Next click — continuing.")
                    return True
                if state_after == _B2B_STATE_READY_TO_GO:
                    continue
                return False
            steps_done += 1
            time.sleep(1)
            continue

        if state == _B2B_STATE_READY_TO_GO:
            # User page: <button onclick="vidapayAutomaticSignIn()">Continue</button>.
            # VidaPay click_ready_to_go_continue: onclick-first, then
            # text-only, then the old #setupNextBtn Next.
            clicked = _b2b_click_button(
                driver, label="Ready To Go Continue",
                text_contains="continue", onclick_contains="vidapayautomaticsignin",
                timeout=18, log=log)
            if not clicked:
                clicked = _b2b_click_button(
                    driver, label="Ready To Go Continue fallback",
                    text_contains="continue", timeout=18, log=log)
            if not clicked:
                # Older B2BSoft variants reuse #setupNextBtn on Ready To Go.
                clicked = _b2b_click_button(driver, label="Ready To Go final Next fallback",
                                            text_contains="next", button_id="setupNextBtn",
                                            timeout=8, log=log)
            if not clicked:
                return False
            wait_state = _b2b_wait_for_states(
                driver, stop_event=stop_event, wanted=(_B2B_STATE_PORTAL,),
                timeout=90, log=log, prefix="After Ready To Go Continue")
            return wait_state == _B2B_STATE_PORTAL

        # NEW_SIGN / TWO_FA / UNKNOWN / LOGIN — wait for a setup or terminal state
        wait_state = _b2b_wait_for_states(
            driver, stop_event=stop_event,
            wanted=(_B2B_STATE_TRUST_DEVICE, _B2B_STATE_SETUP_NEXT,
                    _B2B_STATE_READY_TO_GO, _B2B_STATE_SECURITY_UPG,
                    _B2B_STATE_PORTAL),
            timeout=45, log=log, prefix="next setup step")
        if wait_state is None:
            log("[B2B] No more setup pages found and portal did not load.")
            return False
        if wait_state == _B2B_STATE_PORTAL:
            return True
        steps_done += 1

    log(f"[B2B] Setup flow hit the safety limit after {max_next_clicks} setup action(s).")
    return _b2b_get_page_state(driver) == _B2B_STATE_PORTAL


def _b2b_finish_login_flow(driver, stop_event=None, log=print, timeout: int = 300) -> bool:
    """State machine for post-login B2B flow — ports VidaPay's
    handle_ibm_verify_and_setup → finish_setup_steps →
    complete_remaining_setup_next_flow chain.

    Handles (everything AFTER the access code / credentials steps):
      - New Sign In (<h3>New Sign In</h3>) → click Next (onclick goToTwoFactorCheck)
      - 2FA page → wait for the Trust Device radio (#trustRadio), select it,
        click its Next (#setupNextBtn / onclick submitthisform)
      - Trust Device → #setupNextBtn walk → Ready To Go → portal
      - Security Upgrade page → wait it out (never click its Next)
      - Cloudflare re-verification → wait to clear
      - Portal/alert popups → dismiss 'Later' alerts → done
    """
    deadline = time.time() + timeout
    last_log = 0.0
    last_state = None
    steps = 0

    while time.time() < deadline:
        if stop_event is not None and stop_event.is_set():
            return False

        state = _b2b_get_page_state(driver)

        if state != last_state:
            log(f"[B2B] Page state: {state}")
            last_state = state
            steps = 0

        if state == _B2B_STATE_PORTAL:
            log("✓ B2B portal reached — login complete.")
            _b2b_clear_later_alerts(driver, log=log)
            return True

        if state == _B2B_STATE_VERIFY:
            if not _b2b_wait_for_human_verification_clear(driver, stop_event=stop_event, log=log):
                return False
            last_state = None
            continue

        if state == _B2B_STATE_NEW_SIGN:
            # Faithful port of VidaPay click_new_sign_in_next. The real page
            # shows <h3>New Sign In</h3> with
            # <button onclick="goToTwoFactorCheck()">Next</button>. Selenium
            # clickability waits miss it (SPA renders the button
            # non-interactive first), so match text+onclick and click via the
            # force-enabling JS path. Retry every pass — a silently failed
            # click must not strand the flow on this page.
            clicked = _b2b_click_button(
                driver, label="New Sign In Next",
                text_contains="next", onclick_contains="gototwofactorcheck",
                timeout=30, log=log)
            if not clicked:
                clicked = _b2b_click_button(
                    driver, label="New Sign In Next fallback",
                    text_contains="next", timeout=10, log=log)
            if not clicked:
                _b2b_click_any_next(driver, log=log)
            steps += 1
            time.sleep(1)
            continue

        if state == _B2B_STATE_TWO_FA:
            # 2FA page: once the code is sent/approved, the Trust Device
            # radio (#trustRadio) appears on this page. Port of VidaPay
            # select_trust_radio_quickly — select it the moment it exists in
            # the DOM, then submit via its Next (#setupNextBtn / onclick
            # submitthisform). User flow: read New Sign In → click Next →
            # continue with the 2FA radio.
            log("[B2B] 2FA page — waiting for the Trust Device radio and selecting it automatically…")
            if not _b2b_select_trust_radio(driver, stop_event=stop_event, log=log):
                return False
            if _b2b_get_page_state(driver) == _B2B_STATE_TWO_FA:
                # Radio was checked on the 2FA page itself — click its Next.
                # id-first (user pages show #setupNextBtn without onclick),
                # onclick submitthisform as fallback.
                if not _b2b_click_button(
                        driver, label="2FA radio Next",
                        text_contains="next", button_id="setupNextBtn",
                        timeout=30, log=log):
                    _b2b_click_button(
                        driver, label="2FA radio Next fallback",
                        text_contains="next", onclick_contains="submitthisform",
                        timeout=10, log=log)
            last_state = None
            continue

        if state in (_B2B_STATE_TRUST_DEVICE, _B2B_STATE_SETUP_NEXT,
                     _B2B_STATE_READY_TO_GO, _B2B_STATE_SECURITY_UPG):
            # Trust radio → #setupNextBtn walk → Ready To Go → portal is one
            # continuous walk; delegate to the ported VidaPay flow.
            ok = _b2b_complete_setup_next_flow(driver, stop_event=stop_event, log=log)
            if ok:
                _b2b_clear_later_alerts(driver, log=log)
            return ok

        if state == _B2B_STATE_LOGIN:
            log("B2B returned to login page — credentials may be wrong.")
            return False

        # UNKNOWN — SPA may still be rendering; wait before re-evaluating
        time.sleep(2)

    log("B2B login flow timed out waiting for portal.")
    return False


# Legacy stubs kept for any external callers
def _b2b_is_new_sign_in_page(driver) -> bool:
    return _b2b_get_page_state(driver) == _B2B_STATE_NEW_SIGN

def _b2b_click_new_sign_in_next(driver, log=print) -> bool:
    return _b2b_click_any_next(driver, log=log)

def _b2b_wait_for_new_sign_in_clear(driver, stop_event=None, timeout: int = 300, log=print) -> bool:
    return _b2b_finish_login_flow(driver, stop_event=stop_event, log=log, timeout=timeout)


# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# Edge remote-debug helpers (port 9227 = user profile with --remote-debugging-port)
# ─────────────────────────────────────────────────────────────────────────────
EDGE_DEBUG_PORT = 9227
_WA_URL = "https://web.whatsapp.com"
_B2B_URL = "https://wsreports.b2bsoft.com/#"
_GFH_APP_URL = "https://gfh-telecom-app.web.app/timesheet"
_wa_fallback_opened: bool = False  # prevent webbrowser.open firing multiple times when Edge not running

# ── Tesseract / Ghostscript detection (ported from VidaPay Transfer Bot) ─────
_TESSERACT_CANDIDATES = [
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Tesseract-OCR", "tesseract.exe"),
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    r"C:\Tesseract-OCR\tesseract.exe",
]
_GHOSTSCRIPT_EXES = ("gswin64c", "gswin32c", "gswin64", "gswin32", "gs")


# ── OCR dependency detection (same pattern as VidaPay Transfer Bot) ─────────
# pytesseract is imported at MODULE LEVEL so PyInstaller bundles it into the
# frozen EXE (a lazy import inside a function can be missed when the package
# is absent from the build environment). Detection uses the module-level flag.
try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except Exception:
    pytesseract = None
    PYTESSERACT_AVAILABLE = False

# Tesseract silent installer fallback (UB-Mannheim build, same as Transfer Bot)
_OCR_TESSERACT_URL = (
    "https://digi.bib.uni-mannheim.de/tesseract/"
    "tesseract-ocr-w64-setup-5.3.3.20231005.exe"
)


def _is_tesseract_installed() -> bool:
    if shutil.which("tesseract"):
        return True
    return any(os.path.isfile(p) for p in _TESSERACT_CANDIDATES)


def _locate_tesseract() -> str:
    for p in _TESSERACT_CANDIDATES:
        if os.path.isfile(p):
            return p
    found = shutil.which("tesseract")
    return found if found else _TESSERACT_CANDIDATES[0]


if PYTESSERACT_AVAILABLE:
    # Point pytesseract at the best known tesseract.exe right away
    # (Transfer Bot does this at module level too).
    pytesseract.pytesseract.tesseract_cmd = _locate_tesseract()


def _tool_on_path(name: str) -> bool:
    return shutil.which(name) is not None


def _run_cmd_quiet(cmd, timeout: float = 300):
    """Run a command and return (ok, output). Never raises."""
    try:
        import subprocess as _sp
        proc = _sp.run(cmd, capture_output=True, text=True, timeout=timeout)
        output = ((proc.stdout or "") + (proc.stderr or "")).strip()
        return proc.returncode == 0, output
    except Exception as exc:
        return False, str(exc)


def _download_file(url: str, dest: str, log=print, timeout: float = 240) -> bool:
    """Download url to dest (Transfer Bot helper). Returns True on success."""
    try:
        import urllib.request as _urlreq
        log(f"Downloading {os.path.basename(url)} ...")
        req = _urlreq.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with _urlreq.urlopen(req, timeout=timeout) as resp, open(dest, "wb") as out:
            shutil.copyfileobj(resp, out)
        return os.path.isfile(dest) and os.path.getsize(dest) > 0
    except Exception as exc:
        log(f"Download failed: {exc}")
        return False


def _pip_cmd():
    """Return a pip command usable in this Python, even inside a frozen
    PyInstaller exe where sys.executable points at the exe itself."""
    if getattr(sys, "frozen", False):
        for cand in ("python", "python3", "py"):
            found = shutil.which(cand)
            if found:
                return [found, "-m", "pip"]
        return None
    return [sys.executable, "-m", "pip"]


def _refresh_tesseract_path() -> None:
    """Re-point pytesseract at the best tesseract.exe after an install."""
    if PYTESSERACT_AVAILABLE:
        pytesseract.pytesseract.tesseract_cmd = _locate_tesseract()


def _install_tesseract_binary(log=print) -> bool:
    """Install the Tesseract OCR binary automatically (VidaPay Transfer Bot approach):
    winget first, then the official silent installer as fallback."""
    if _is_tesseract_installed():
        return True
    if _tool_on_path("winget"):
        log("Installing Tesseract OCR via winget…")
        _run_cmd_quiet(
            ["winget", "install", "--id", "UB-Mannheim.TesseractOCR", "-e",
             "--accept-source-agreements", "--accept-package-agreements", "--silent"],
            timeout=600,
        )
        if _is_tesseract_installed():
            log("Tesseract installed via winget.")
            return True
    installer = os.path.join(APP_DIR, "tesseract-setup.exe")
    if _download_file(_OCR_TESSERACT_URL, installer, log):
        log("Running Tesseract silent installer (this can take a minute)…")
        _run_cmd_quiet([installer, "/S"], timeout=900)
        try:
            os.remove(installer)
        except Exception:
            pass
        if _is_tesseract_installed():
            log("Tesseract installed from official installer.")
            return True
    return False


def _ghostscript_installed() -> bool:
    for name in _GHOSTSCRIPT_EXES:
        if shutil.which(name):
            return True
    import sys as _sys
    if _sys.platform.startswith("win"):
        for base in (r"C:\Program Files\gs", r"C:\Program Files (x86)\gs"):
            if not os.path.isdir(base):
                continue
            try:
                for ver_dir in os.listdir(base):
                    bin_dir = os.path.join(base, ver_dir, "bin")
                    if any(os.path.isfile(os.path.join(bin_dir, n + ".exe")) for n in _GHOSTSCRIPT_EXES):
                        return True
            except OSError:
                continue
    return False

# Dedicated Edge profile for GFH automation (same pattern as VidaPay transfer bot).
# Edge is launched with --remote-debugging-port=9227 against this profile so
# WhatsApp, B2B, and GFH app sessions persist across restarts.
GFH_AUTOMATION_PROFILE_DIR = r"C:\GFH_Edge_Automation_Profile"


def _get_edge_exe() -> Optional[str]:
    import shutil as _shutil
    for candidate in (
        _shutil.which("msedge"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ):
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


def _is_edge_port_open(port: int = EDGE_DEBUG_PORT, timeout: float = 1.0) -> bool:
    import socket as _socket
    try:
        with _socket.create_connection(("127.0.0.1", port), timeout=timeout):
            return True
    except Exception:
        return False


def _ensure_edge_open(port: int = EDGE_DEBUG_PORT, log=None) -> bool:
    """Launch Edge at debug port using GFH automation profile if not already running."""
    if _is_edge_port_open(port):
        return True
    edge_exe = _get_edge_exe()
    if not edge_exe:
        if log:
            log("⚠ Microsoft Edge not found — install Edge and retry.")
        return False
    import subprocess as _sp
    os.makedirs(GFH_AUTOMATION_PROFILE_DIR, exist_ok=True)
    _sp.Popen([
        edge_exe,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={GFH_AUTOMATION_PROFILE_DIR}",
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "about:blank",
    ])
    for _ in range(20):  # wait up to 10 s
        time.sleep(0.5)
        if _is_edge_port_open(port):
            return True
    if log:
        log("⚠ Edge launched but port 9227 not ready — wait and retry.")
    return False


def _edge_debug_driver(port: int = EDGE_DEBUG_PORT):
    """Return a Selenium driver attached to the already-running Edge profile at port."""
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options as EdgeOptions
    try:
        if not _is_edge_port_open(port):
            # requested port closed - zero-config discovery from
            # DevToolsActivePort files (chrome://inspect remote debugging;
            # idea: chrome-devtools-mcp#1826)
            _disc = _discover_devtools_port()
            if _disc:
                port = _disc
    except Exception:
        pass
    opts = EdgeOptions()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    return webdriver.Edge(options=opts)


# Tabs that Edge starts with and that are safe to REUSE for a real page
# (the automation Edge is launched with a about:blank start tab — navigating
# it beats spawning a second tab and leaving the empty window in front).
_BLANK_TAB_URLS = ("about:blank", "edge://newtab", "data:,")


def _find_or_open_tab(driver, url: str) -> None:
    """Switch to existing tab on the same origin as url; open a new tab if absent.

    A leftover blank/new-tab tab (Edge starts on about:blank) is NAVIGATED to
    the target URL instead of spawning an extra tab, so the automation window
    shows the real page rather than an empty about:blank tab in front."""
    from urllib.parse import urlparse
    parsed = urlparse(url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    blank_handle = None
    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            cur = driver.current_url
        except Exception:
            continue
        if cur.startswith(origin):
            return
        if blank_handle is None and any(cur == b or cur.startswith(b) for b in _BLANK_TAB_URLS):
            blank_handle = handle
    if blank_handle is not None:
        try:
            driver.switch_to.window(blank_handle)
            driver.get(url)
            return
        except Exception:
            pass
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])


def _humanize_list(items: list) -> str:
    """['B2B', 'GFH app', 'WhatsApp Web'] -> 'B2B, GFH app, and WhatsApp Web'."""
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + ", and " + items[-1]


def open_monitoring_tabs(port: int = EDGE_DEBUG_PORT, include_whatsapp: bool = True) -> list:
    """Open B2B, GFH app (and optionally WhatsApp Web) tabs in Edge at debug port.

    Returns the list of tab names that are now open. WhatsApp Web is only
    opened when include_whatsapp is True (i.e. WhatsApp Web mode is selected);
    with WhatsApp Desktop mode only B2B and GFH app tabs are opened.
    """
    opened: list = []
    try:
        driver = _edge_debug_driver(port)
        for url, name in ((_B2B_URL, "B2B"), (_GFH_APP_URL, "GFH app")):
            _find_or_open_tab(driver, url)
            opened.append(name)
        if include_whatsapp:
            _find_or_open_tab(driver, _WA_URL)
            opened.append("WhatsApp Web")
    except Exception:
        pass
    return opened


# B2B Soft Scraper (wsreports.b2bsoft.com)
# ─────────────────────────────────────────────────────────────────────────────
class B2BSoftScraper:
    """Selenium scraper for the B2B Soft inventory portal.

    Login flow:
        Step 1 → Enter the Company ID / Access Code (#companyId) → click #btnSubmit
                 (the SSO page labels #companyId "Access Code" — the Company
                 ID doubles as the Access Code)
        Step 2 → Enter Account ID (#AccountId)
        Step 3 → Enter Username (#Username) + Password (#Password) → click #btnClick

    After login, navigates to Inventory Count Result Details report and
    downloads the XLSX file.
    """
    PORTAL_URL = "https://wsreports.b2bsoft.com/#"
    REPORT_SELECTOR = "[uniq='9-212318']"

    def __init__(self, company_id: str, account_id: str,
                 username: str, password: str,
                 download_dir: Path, log_fn=None, access_code: str = ""):
        self.company_id = company_id
        # Optional separate Access Code. The SSO login page labels its
        # #companyId field "Access Code", but the Company ID doubles as the
        # Access Code — when this is empty, the Company ID is typed there.
        self.access_code = (access_code or "").strip()
        self.account_id = account_id
        self.username = username
        self.password = password
        self.download_dir = Path(download_dir)
        _base_log = log_fn or (lambda m: None)

        def _log_no_double_prefix(m):
            # log_fn already prepends "[B2B] " — strip hardcoded copies from
            # messages so the scheduler log never shows "[B2B] [B2B] ...".
            m = str(m)
            _base_log(m[len("[B2B] "):] if m.startswith("[B2B] ") else m)

        self.log = _log_no_double_prefix
        self.driver = None

    def _make_driver(self):
        self._using_debug_port = False
        # Try user's running Edge profile first (port 9227)
        try:
            self.driver = _edge_debug_driver()
            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": str(self.download_dir),
            })
            _find_or_open_tab(self.driver, _B2B_URL)
            self._using_debug_port = True
            return
        except Exception:
            pass
        # Fallback: new Edge or Chrome instance
        try:
            from selenium import webdriver
            from selenium.webdriver.edge.options import Options as EdgeOptions
            opts = EdgeOptions()
            prefs = {
                "download.default_directory": str(self.download_dir),
                "download.prompt_for_download": False,
                "plugins.always_open_pdf_externally": True,
            }
            opts.add_experimental_option("prefs", prefs)
            self.driver = webdriver.Edge(options=opts)
        except Exception:
            try:
                from selenium import webdriver
                from selenium.webdriver.chrome.options import Options as ChromeOptions
                opts = ChromeOptions()
                prefs = {
                    "download.default_directory": str(self.download_dir),
                    "download.prompt_for_download": False,
                }
                opts.add_experimental_option("prefs", prefs)
                self.driver = webdriver.Chrome(options=opts)
            except Exception as exc:
                raise RuntimeError(f"Could not start Edge or Chrome WebDriver: {exc}")

    def login(self, stop_event=None) -> bool:
        """Full B2B login with Cloudflare Turnstile + reCAPTCHA + 2FA support.

        Flow:
            Open portal → [verification?] → Step 1 Company ID → [verification?]
            → Step 2 Account ID → Step 3 Username + Password → [verification?]
            → confirm authenticated
        """
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        if not all([self.company_id, self.username, self.password]):
            raise RuntimeError("B2B credentials incomplete — fill Portal Credentials tab (Account ID is optional).")
        if self.driver is None:
            self._make_driver()
        self.download_dir.mkdir(parents=True, exist_ok=True)
        # Switch to B2B tab before navigating — prevents overwriting another scraper's tab.
        _find_or_open_tab(self.driver, self.PORTAL_URL)
        self.log(f"Opening {self.PORTAL_URL}")
        self.driver.get(self.PORTAL_URL)
        time.sleep(2)

        # Clear any initial verification challenge
        if _b2b_is_human_verification_page(self.driver):
            if not _b2b_wait_for_human_verification_clear(self.driver, stop_event=stop_event, log=self.log):
                raise RuntimeError("Cloudflare/reCAPTCHA challenge on landing page was not resolved.")

        wait = WebDriverWait(self.driver, 30)
        drv = self.driver

        def _robust_type(element, text):
            """JS-clear + event-dispatch before typing — same pattern as VidaPay extractor."""
            drv.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
            drv.execute_script("arguments[0].click();", element)
            try:
                drv.execute_script(
                    "arguments[0].value='';"
                    "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                    "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                    element,
                )
            except Exception:
                pass
            try:
                element.clear()
            except Exception:
                pass
            try:
                existing = element.get_attribute("value") or ""
                if existing:
                    from selenium.webdriver.common.keys import Keys as _K
                    element.send_keys(_K.BACK_SPACE * (len(existing) + 2))
            except Exception:
                pass
            # HUMAN TYPING: char-by-char with randomized inter-key delays
            # (45-160ms, occasional 200-420ms hesitations) instead of an
            # instant paste-burst of the whole string.
            try:
                import random as _rnd
                time.sleep(_rnd.uniform(0.25, 0.7))
                for _ch in str(text):
                    element.send_keys(_ch)
                    _pause = _rnd.uniform(0.045, 0.16)
                    if _rnd.random() < 0.12:
                        _pause = _rnd.uniform(0.2, 0.42)
                    time.sleep(_pause)
            except Exception:
                element.send_keys(text)

        def _js_click(element):
            try:
                element.click()
            except Exception:
                drv.execute_script("arguments[0].click();", element)

        def _visible_fields_snapshot():
            """List visible input/button identifiers on the page (diagnostics)."""
            try:
                return drv.execute_script("""
                    const out = [];
                    document.querySelectorAll('input, button').forEach(el => {
                        const r = el.getBoundingClientRect();
                        if (r.width > 0 && r.height > 0) {
                            out.push((el.id ? '#' + el.id : '')
                                + (el.name ? '[name=' + el.name + ']' : '')
                                + (el.type && el.type !== 'text' ? '[' + el.type + ']' : ''));
                        }
                    });
                    return out.slice(0, 20);
                """)
            except Exception:
                return []

        def _wait_for_field_or_verify(field_id: str, label: str, timeout: int = 60):
            """Wait for a login-form field to appear, clearing any Cloudflare that blocks it."""
            deadline = time.time() + timeout
            last_status_log = 0.0
            while time.time() < deadline:
                if stop_event is not None and stop_event.is_set():
                    raise RuntimeError("Cancelled by stop event.")
                if _b2b_is_human_verification_page(drv):
                    self.log(f"[B2B] Cloudflare detected before {label} — waiting to clear…")
                    if not _b2b_wait_for_human_verification_clear(drv, stop_event=stop_event, log=self.log):
                        raise RuntimeError(f"Cloudflare not resolved before {label}.")
                els = drv.find_elements(By.ID, field_id)
                if els:
                    try:
                        if els[0].is_displayed():
                            return els[0]
                    except Exception:
                        pass
                now = time.time()
                if now - last_status_log >= 15:
                    elapsed = int(now - (deadline - timeout))
                    try:
                        url = drv.current_url
                    except Exception:
                        url = "unknown"
                    self.log(f"[B2B] Waiting for {label} ({elapsed}s elapsed, on {url})")
                    _fields = _visible_fields_snapshot()
                    if _fields:
                        self.log(f"[B2B] Visible fields on page: {', '.join(_fields)}")
                    last_status_log = now
                time.sleep(0.5)
            raise RuntimeError(f"B2B field #{field_id} ({label}) not visible after {timeout}s.")

        # ── Step 1: Company ID ─────────────────────────────────────────────
        # The #companyId input is dual-purpose: on the wsreports landing page
        # it asks for the Company ID, while the SSO login page labels the
        # same field "Access Code" (placeholder). The Company ID doubles as
        # the Access Code. Pick the value from the field's placeholder,
        # falling back to the current URL when the placeholder is missing.
        def _stage_value(el):
            try:
                ph = (el.get_attribute("placeholder") or "").strip().lower()
            except Exception:
                ph = ""
            want_access = "access code" in ph
            if not ph:
                try:
                    want_access = "sso.b2bsoft.com" in (drv.current_url or "")
                except Exception:
                    want_access = False
            if want_access:
                # The Company ID doubles as the Access Code when no separate
                # one is configured.
                return ((self.access_code or "").strip()
                        or (self.company_id or "").strip(), "Access Code")
            return (self.company_id or "").strip(), "Company ID"

        def _type_and_verify(el, text):
            """Type text, then verify the DOM value actually took — if the
            field ignored send_keys (readonly/disabled), force it via JS so
            the form POST still carries the value. Returns the final value."""
            _robust_type(el, text)
            try:
                got = (el.get_attribute("value") or "").strip()
            except Exception:
                got = ""
            if got != text:
                try:
                    drv.execute_script(
                        "arguments[0].value = arguments[1];"
                        "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                        el, text)
                    got = (el.get_attribute("value") or "").strip()
                except Exception:
                    pass
            return got

        try:
            f = wait.until(EC.visibility_of_element_located((By.ID, "companyId")))
            val, label = _stage_value(f)
            got = _type_and_verify(f, val)
            if got != val:
                self.log(f"⚠ {label} field did not accept typing (value now '{got}') — attempting submit anyway.")
            self.log(f"{label} entered: {got}")
            btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
            _js_click(btn)
            self.log(f"{label} submitted.")
            # Wait for the companyId field to go hidden (page transition started)
            # before polling for AccountId — prevents burning 150s on the wrong page.
            # Landing on sso.b2bsoft.com counts as transitioned — its Step-2
            # handling below deals with the SSO page's own company-ID stage.
            _trans_deadline = time.time() + 20
            while time.time() < _trans_deadline:
                try:
                    try:
                        if "sso.b2bsoft.com" in (drv.current_url or ""):
                            break
                    except Exception:
                        pass
                    els = drv.find_elements(By.ID, "companyId")
                    if not els or not els[0].is_displayed():
                        break
                except Exception:
                    break
                time.sleep(0.5)
        except Exception as e:
            raise RuntimeError(f"B2B Step 1 (Company ID) failed: {e}")

        # ── Step 2: Account ID — adaptive wait (fixes the SSO stall) ────────
        # The SSO login page (sso.b2bsoft.com/account/login) may show an
        # Account ID field — or may only show Username + Password (company ID
        # carried by the SSO session). Poll for all candidates at once, type
        # whatever actually appears, and log the page's visible fields every
        # 15s so any future stall is explained directly in the log.
        try:
            deadline = time.time() + 150
            account_el = None
            username_seen = False
            last_diag = 0.0
            sso_submits = 0
            while time.time() < deadline:
                if stop_event is not None and stop_event.is_set():
                    raise RuntimeError("Cancelled by stop event.")
                if _b2b_is_human_verification_page(drv):
                    self.log("Cloudflare detected before Account ID — waiting to clear…")
                    if not _b2b_wait_for_human_verification_clear(drv, stop_event=stop_event, log=self.log):
                        raise RuntimeError("Cloudflare not resolved before Account ID.")
                # Account ID field present and visible? → use it.
                acc = drv.find_elements(By.ID, "AccountId")
                if acc:
                    try:
                        if acc[0].is_displayed():
                            account_el = acc[0]
                            break
                    except Exception:
                        pass
                # Username already visible without AccountId? → SSO page has
                # no Account ID step; skip straight to Username + Password.
                usr = drv.find_elements(By.ID, "Username")
                if usr:
                    try:
                        if usr[0].is_displayed():
                            username_seen = True
                            self.log("No Account ID field on this page — going straight to Username.")
                            break
                    except Exception:
                        pass
                # SSO company-ID stage (user-supplied HTML): the page shows
                # ONLY #companyId (placeholder "Access Code", required="True")
                # + #btnSubmit ("Continue") — the Company ID doubles as the
                # Access Code. Every attempt LOGS the typed value (silent
                # typing hid a visually-empty field), re-forces the DOM value
                # right before submit (required + empty field = the browser
                # silently blocks the submit), unlocks a read-only field via
                # #btnCompanyIdEdit when one exists, and after a failed
                # advance reads the page's validation and error text so a
                # server rejection is named in the log.
                comp = drv.find_elements(By.ID, "companyId")
                sso_stuck = False
                if comp:
                    try:
                        if comp[0].is_displayed():
                            if sso_submits >= 4:
                                sso_stuck = True
                            else:
                                # Give the page a moment to finish loading —
                                # attempt 1 used to fire within the same second
                                # as the redirect, possibly before the form's
                                # JS handlers had attached.
                                try:
                                    for _ in range(10):
                                        if drv.execute_script("return document.readyState;") == "complete":
                                            break
                                        time.sleep(0.5)
                                except Exception:
                                    pass
                                val, label = _stage_value(comp[0])
                                # Instrument the page once per render: record
                                # submit events + fetch/XHR so a failed attempt
                                # shows whether the form actually posted, and
                                # log the form's action/method/button binding.
                                try:
                                    _sso_info = drv.execute_script("""
                                        if (window.__b2bNetInstalled) return {already: true};
                                        window.__b2bNetInstalled = true;
                                        window.__b2bNet = [];
                                        window.addEventListener('submit', function(e) {
                                            window.__b2bNet.push('submit-event -> ' + ((e.target && e.target.action) || 'no-action'));
                                        }, true);
                                        var _fs = HTMLFormElement.prototype.submit;
                                        HTMLFormElement.prototype.submit = function() {
                                            window.__b2bNet.push('form.submit() -> ' + (this.action || 'no-action'));
                                            return _fs.apply(this, arguments);
                                        };
                                        if (window.fetch) {
                                            var _f = window.fetch;
                                            window.fetch = function() {
                                                try { window.__b2bNet.push('fetch -> ' + (arguments[0] && (arguments[0].url || arguments[0]))); } catch (e) {}
                                                return _f.apply(this, arguments);
                                            };
                                        }
                                        var _xo = XMLHttpRequest.prototype.open;
                                        XMLHttpRequest.prototype.open = function(m, u) {
                                            try { window.__b2bNet.push('xhr -> ' + m + ' ' + u); } catch (e) {}
                                            return _xo.apply(this, arguments);
                                        };
                                        var el = document.querySelector('#companyId');
                                        var btn = document.querySelector('#btnSubmit');
                                        var form = el ? el.form : null;
                                        return {
                                            installed: true,
                                            formAction: form ? String(form.action || '') : null,
                                            formMethod: form ? String(form.method || '') : null,
                                            btnInForm: !!(btn && form && btn.form === form),
                                            btnDisabled: !!(btn && (btn.disabled || btn.getAttribute('aria-disabled') === 'true')),
                                            inputs: Array.prototype.map.call(
                                                document.querySelectorAll('input'),
                                                function(i) {
                                                    return (i.id ? '#' + i.id : (i.name ? '[name=' + i.name + ']' : 'input'))
                                                        + (i.type && i.type !== 'text' ? '[' + i.type + ']' : '')
                                                        + (i.required ? '[required]' : '');
                                                }).slice(0, 12)
                                        };
                                    """) or {}
                                    if _sso_info.get("installed"):
                                        self.log(
                                            "SSO page form: action=" + str(_sso_info.get("formAction"))
                                            + ", method=" + str(_sso_info.get("formMethod"))
                                            + ", #btnSubmit inside form: " + ("yes" if _sso_info.get("btnInForm") else "NO")
                                            + (", #btnSubmit DISABLED" if _sso_info.get("btnDisabled") else "")
                                            + ", inputs: " + ", ".join(_sso_info.get("inputs") or []))
                                except Exception:
                                    _sso_info = {}
                                # Read-only/disabled field? The Edit button
                                # exists to unlock it — click it first.
                                try:
                                    ro = drv.execute_script(
                                        "const el = document.querySelector('#companyId');"
                                        "return !!(el && (el.readOnly || el.disabled));")
                                except Exception:
                                    ro = False
                                if ro:
                                    edit_btns = drv.find_elements(By.ID, "btnCompanyIdEdit")
                                    if edit_btns:
                                        _js_click(edit_btns[0])
                                        self.log("SSO Access Code field is read-only — clicked #btnCompanyIdEdit to unlock it.")
                                        time.sleep(0.8)
                                try:
                                    cur_val = (comp[0].get_attribute("value") or "").strip()
                                except Exception:
                                    cur_val = ""
                                if cur_val != val:
                                    got = _type_and_verify(comp[0], val)
                                    if got != val:
                                        self.log(f"⚠ SSO {label} field did not accept typing (value now '{got}') — attempting submit anyway.")
                                    else:
                                        self.log(f"SSO {label} typed: '{got}'.")
                                else:
                                    self.log(f"SSO {label} field already holds '{cur_val}'.")
                                    # Value was pre-filled by the server — the
                                    # page's JS may still be waiting for
                                    # input/change events before it accepts the
                                    # form. Fire them (harmless when unneeded).
                                    try:
                                        drv.execute_script(
                                            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                                            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                            comp[0])
                                    except Exception:
                                        pass
                                # A previous failed submit can leave the Continue
                                # button disabled (double-submit guard) — a
                                # disabled button swallows every click.
                                try:
                                    if drv.execute_script(
                                            "var b = document.querySelector('#btnSubmit');"
                                            "if (b && (b.disabled || b.getAttribute('aria-disabled') === 'true')) {"
                                            "  b.disabled = false; b.removeAttribute('aria-disabled');"
                                            "  return true;"
                                            "} return false;"):
                                        self.log("SSO #btnSubmit was disabled — force-enabled it before submitting.")
                                except Exception:
                                    pass
                                # The real page marks the field required="True" —
                                # an EMPTY value makes the browser silently block
                                # every submit with "Please fill out this field".
                                # Re-check the DOM value at the last moment and
                                # force it in via JS so the POST carries the code.
                                try:
                                    pre_val = (comp[0].get_attribute("value") or "").strip()
                                except Exception:
                                    pre_val = val
                                if pre_val != val:
                                    try:
                                        drv.execute_script(
                                            "arguments[0].value = arguments[1];"
                                            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                                            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                                            comp[0], val)
                                        self.log(f"SSO {label} value re-forced via JS before submit ('{val}').")
                                    except Exception:
                                        pass
                                try:
                                    submit_val = (comp[0].get_attribute("value") or "").strip()
                                except Exception:
                                    submit_val = val
                                sso_submits += 1
                                # Rotation order — measured on the live
                                # page (16:23 scheduler log): the #btnSubmit
                                # click and ENTER both produced "submit/
                                # network activity: NONE" while form
                                # .requestSubmit() POSTed to /Account/
                                # LoginCompany and advanced. So requestSubmit
                                # goes FIRST — it keeps native validation and
                                # the button=continue submitter value; the
                                # click / ENTER / raw-submit stay as fallbacks.
                                strategy = ("rsubmit", "submit", "enter", "formsubmit")[sso_submits - 1]
                                if strategy == "enter":
                                    from selenium.webdriver.common.keys import Keys as _Keys
                                    comp[0].send_keys(_Keys.RETURN)
                                    self.log(f"SSO Access Code page — ENTER sent on {label} field (attempt {sso_submits}, field='{submit_val}').")
                                elif strategy == "rsubmit":
                                    try:
                                        how = drv.execute_script(
                                            "const el = document.querySelector('#companyId');"
                                            "const btn = document.querySelector('#btnSubmit');"
                                            "const form = (el && el.form) || (btn && btn.form);"
                                            "if (!form) return 'noform';"
                                            "if (form.requestSubmit) { form.requestSubmit(btn || undefined); return 'requestSubmit'; }"
                                            "form.submit(); return 'form.submit';")
                                        self.log(f"SSO Access Code page — {label} submitted via form {how} (attempt {sso_submits}, field='{submit_val}').")
                                    except Exception as e:
                                        self.log(f"SSO Access Code page — form submit failed: {e}")
                                elif strategy == "formsubmit":
                                    try:
                                        how = drv.execute_script(
                                            "var el = document.querySelector('#companyId');"
                                            "var form = el ? el.form : null;"
                                            "if (!form) {"
                                            "  var btn = document.querySelector('#btnSubmit');"
                                            "  form = btn ? btn.form : null;"
                                            "}"
                                            "if (!form) {"
                                            "  var fs = document.querySelectorAll('form');"
                                            "  for (var i = 0; i < fs.length; i++) {"
                                            "    if (fs[i].querySelector('#companyId')) { form = fs[i]; break; }"
                                            "  }"
                                            "}"
                                            "if (!form) return 'noform';"
                                            "form.submit();"
                                            "return 'posted -> ' + String(form.action || '').slice(-60);")
                                        self.log(f"SSO Access Code page — raw form.submit() sent (attempt {sso_submits}, field='{submit_val}', {how}).")
                                    except Exception as e:
                                        self.log(f"SSO Access Code page — raw form.submit() failed: {e}")
                                else:
                                    btn_id = "btnSubmit"
                                    btns = drv.find_elements(By.ID, btn_id)
                                    if not btns:
                                        # Real button text is "Continue" —
                                        # fall back to a text match.
                                        try:
                                            btns = [b for b in drv.find_elements(By.TAG_NAME, "button")
                                                    if (b.text or "").strip().lower() == "continue"
                                                    and b.is_displayed()]
                                            if btns:
                                                btn_id = "button:Continue"
                                        except Exception:
                                            btns = []
                                    if btns:
                                        _js_click(btns[0])
                                        self.log(f"SSO Access Code page — {label} submitted via #{btn_id} (attempt {sso_submits}, field='{submit_val}').")
                                    else:
                                        self.log(f"SSO Access Code page — #{btn_id} not on page (attempt {sso_submits} skipped).")
                                # Give the SSO page up to 8s to advance before
                                # the next attempt — avoids double-submitting
                                # a value the server is still processing.
                                advanced = False
                                for _ in range(16):
                                    time.sleep(0.5)
                                    try:
                                        c2 = drv.find_elements(By.ID, "companyId")
                                        if not c2 or not c2[0].is_displayed():
                                            advanced = True
                                            break
                                    except Exception:
                                        break
                                if not advanced:
                                    # Page did not move — log WHY: field value
                                    # now, HTML5 validation message, visible
                                    # error text, current URL. Fail fast when
                                    # the server actually rejected the code.
                                    try:
                                        why = drv.execute_script("""
                                            const el = document.querySelector('#companyId');
                                            const out = {
                                                value: el ? el.value : null,
                                                validation: el ? el.validationMessage : null
                                            };
                                            out.errors = Array.from(document.querySelectorAll(
                                                '.validation-summary-errors, .field-validation-error, .error, .alert-danger, .text-danger, [class*="error" i]'
                                            )).map(e => (e.innerText || '').trim())
                                              .filter(t => t).slice(0, 3);
                                            out.body = (document.body && document.body.innerText
                                                ? document.body.innerText : '')
                                                .replace(/\\s+/g, ' ').trim().slice(0, 220);
                                            return out;
                                        """)
                                        err_txt = "; ".join(why.get("errors") or [])
                                        try:
                                            url_now = drv.current_url
                                        except Exception:
                                            url_now = "unknown"
                                        self.log(
                                            f"SSO page did not advance after attempt {sso_submits} "
                                            f"(URL: {url_now}) — field now '{why.get('value')}', "
                                            f"validation: '{why.get('validation') or 'none'}'"
                                            + (f", page error: {err_txt}" if err_txt else ""))
                                        try:
                                            _net = drv.execute_script(
                                                "return (window.__b2bNet || []).splice(0, 20);") or []
                                        except Exception:
                                            _net = []
                                        if _net:
                                            self.log("SSO page submit/network activity: " + " | ".join(str(x) for x in _net))
                                        else:
                                            self.log("SSO page submit/network activity: NONE — the submit never reached the form.")
                                        _body_txt = str(why.get("body") or "")
                                        if _body_txt:
                                            self.log(f"SSO page text: {_body_txt}")
                                        _vmsg = (why.get("validation") or "").lower()
                                        if "fill out" in _vmsg or "required" in _vmsg:
                                            self.log(
                                                "⚠ The browser blocked the submit — #companyId was EMPTY "
                                                "at submit time (the Access Code did not stick in the DOM).")
                                        if err_txt and any(k in err_txt.lower() for k in (
                                                "invalid", "incorrect", "wrong", "not found",
                                                "failed", "expire", "locked", "unable")):
                                            raise RuntimeError(
                                                f"B2B SSO rejected the Access Code: {err_txt} — "
                                                "verify the Company ID on the Portal Credentials tab "
                                                "(it doubles as the Access Code) and retry.")
                                    except RuntimeError:
                                        raise
                                    except Exception:
                                        pass
                                continue
                    except RuntimeError:
                        raise
                    except Exception:
                        pass
                if sso_stuck:
                    _fields = _visible_fields_snapshot()
                    raise RuntimeError(
                        "B2B login stalled on the SSO 'Access Code' page — tried 4 times "
                        "(requestSubmit → Submit → Enter → form.submit) without the page advancing. "
                        "The log lines above show the exact field value at every submit "
                        "attempt plus any page error text. Verify the Company ID on the "
                        "Portal Credentials tab (it doubles as the Access Code) and retry. "
                        + ("Visible fields: " + ", ".join(_fields) if _fields else "No visible fields on page.")
                    )
                now = time.time()
                if now - last_diag >= 15:
                    elapsed = int(now - (deadline - 150))
                    try:
                        url = drv.current_url
                    except Exception:
                        url = "unknown"
                    self.log(f"Waiting for Account ID ({elapsed}s elapsed, on {url})")
                    fields = _visible_fields_snapshot()
                    if fields:
                        self.log(f"Visible fields on page: {', '.join(fields)}")
                    last_diag = now
                time.sleep(0.5)
            if account_el is not None:
                if (self.account_id or "").strip():
                    _robust_type(account_el, self.account_id)
                    self.log(f"Account ID entered: {self.account_id}")
                else:
                    self.log("Account ID field found but no Account ID saved — leaving it empty.")
                time.sleep(1)
            elif not username_seen:
                fields = _visible_fields_snapshot()
                raise RuntimeError(
                    "B2B Account ID step failed: neither AccountId nor Username appeared within 150s. "
                    "If the page shows an 'Access Code' field, verify the Company ID on the "
                    "Portal Credentials tab (it doubles as the Access Code). "
                    + ("Visible fields: " + ", ".join(fields) if fields else "No visible input fields on page.")
                )
        except RuntimeError:
            raise
        except Exception as e:
            raise RuntimeError(f"B2B Step 2 (Account ID) failed: {e}")

        # Wait for the first visible field among candidate locators —
        # Cloudflare-aware, with the same 15s diagnostics as the ID waits.
        def _wait_for_any_field(candidates, label: str, timeout: int = 40):
            deadline = time.time() + timeout
            last_status_log = 0.0
            while time.time() < deadline:
                if stop_event is not None and stop_event.is_set():
                    raise RuntimeError("Cancelled by stop event.")
                if _b2b_is_human_verification_page(drv):
                    self.log(f"Cloudflare detected before {label} — waiting to clear…")
                    if not _b2b_wait_for_human_verification_clear(drv, stop_event=stop_event, log=self.log):
                        raise RuntimeError(f"Cloudflare not resolved before {label}.")
                for by, sel in candidates:
                    try:
                        els = drv.find_elements(by, sel)
                    except Exception:
                        continue
                    if els:
                        try:
                            if els[0].is_displayed():
                                return els[0]
                        except Exception:
                            pass
                now = time.time()
                if now - last_status_log >= 15:
                    elapsed = int(now - (deadline - timeout))
                    try:
                        url = drv.current_url
                    except Exception:
                        url = "unknown"
                    self.log(f"Waiting for {label} ({elapsed}s elapsed, on {url})")
                    _fields = _visible_fields_snapshot()
                    if _fields:
                        self.log(f"Visible fields on page: {', '.join(_fields)}")
                    last_status_log = now
                time.sleep(0.5)
            raise RuntimeError(f"B2B field ({label}) not visible after {timeout}s.")

        # ── Step 3: Username + Password — ported from VidaPay_Incentive_Extractor
        # login_store: multi-strategy field location (the SSO sign-in page
        # may not use #Username/#Password ids), submit via text-matched
        # button (Next / Sign In / Log in — force-enabled before clicking,
        # SSO renders it disabled until its JS validates the fields), and
        # invalid-credential re-entry.
        _user_candidates = [
            (By.ID, "Username"),
            (By.XPATH, "//label[contains(normalize-space(),'User Name')]/following::input[1]"),
            (By.XPATH, "//label[contains(normalize-space(),'Username')]/following::input[1]"),
            (By.XPATH, "//input[contains(@placeholder,'User') or contains(@name,'user') or contains(@id,'user')]"),
            (By.XPATH, "//input[@type='email']"),
            (By.XPATH, "(//input[not(@type='hidden') and not(@type='password')])[2]"),
        ]
        _pass_candidates = [
            (By.ID, "Password"),
            (By.XPATH, "//input[@type='password']"),
            (By.XPATH, "//label[contains(normalize-space(),'Password')]/following::input[1]"),
        ]

        def _fill_credentials() -> None:
            u = _wait_for_any_field(_user_candidates, "Username", timeout=60)
            got_u = _type_and_verify(u, self.username)
            self.log(f"Username entered: {got_u}")
            p = _wait_for_any_field(_pass_candidates, "Password", timeout=30)
            _type_and_verify(p, self.password)
            self.log("Password entered.")

        def _submit_credentials() -> bool:
            # Same button-matching chain as VidaPay_Incentive_Extractor: the SSO
            # sign-in button is force-enabled and clicked via JS (native
            # click first, MouseEvent fallback). "Next" first — that is the
            # button the SSO sign-in page actually shows. Waits are short:
            # the generic password-form submit below covers anything these
            # miss, so a wrong guess no longer burns half a minute.
            for _btn_label, kwargs in [
                ("#btnClick verify button", {"button_id": "btnClick", "timeout": 4}),
                ("Next button", {"text_contains": "next", "timeout": 2}),
                ("Sign In button", {"text_contains": "sign in", "timeout": 2}),
                ("Log In button", {"text_contains": "log in", "timeout": 2}),
                ("Login button", {"text_contains": "login", "timeout": 2}),
                ("Continue button", {"text_contains": "continue", "timeout": 2}),
            ]:
                if _b2b_click_button(drv, label=_btn_label, log=self.log, **kwargs):
                    return True
            # None of the known buttons matched (16:24 scheduler log: six
            # searches + input[type=submit] all missed while Account ID /
            # Username / Password were filled). The SSO credentials page
            # can render its submit button DISABLED until its JS validates
            # the typed fields — and _b2b_click_button skips disabled
            # controls — so the real button may be sitting right there
            # unmatchable. Log exactly which controls ARE on the page, then
            # submit the credentials form directly: force-enable every
            # submit control inside the password field's form and click the
            # best one, falling back to form.requestSubmit().
            try:
                inv = drv.execute_script("""
                    const vis = el => { const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0; };
                    const btns = Array.from(document.querySelectorAll(
                        'button, input[type=submit], input[type=button]'))
                        .map(b => ({
                            tag: b.tagName.toLowerCase(),
                            id: b.id || null,
                            type: b.getAttribute('type') || null,
                            text: ((b.innerText || b.value || '') + '').trim().slice(0, 30),
                            disabled: !!b.disabled,
                            visible: vis(b)
                        }));
                    const pass = Array.prototype.find.call(
                        document.querySelectorAll('input[type=password]'), vis);
                    return {btns: btns.slice(0, 10), hasPassword: !!pass,
                            formAction: (pass && pass.form)
                                ? String(pass.form.action || '') : null};
                """) or {}
                _parts = ["Password field: " + ("yes" if inv.get("hasPassword") else "NO")]
                if inv.get("formAction"):
                    _parts.append("password-form action=" + str(inv.get("formAction")))
                for _b in (inv.get("btns") or []):
                    _parts.append(
                        "<" + str(_b.get("tag"))
                        + ("#" + str(_b["id"]) if _b.get("id") else "")
                        + (" type=" + str(_b["type"]) if _b.get("type") else "")
                        + (" DISABLED" if _b.get("disabled") else "")
                        + (": '" + str(_b["text"]) + "'" if _b.get("text") else "")
                        + ("" if _b.get("visible") else " [hidden]") + ">")
                self.log("SSO credentials page controls: " + " | ".join(_parts))
            except Exception:
                pass
            try:
                _how = drv.execute_script("""
                    const vis = el => { if (!el) return false;
                        const r = el.getBoundingClientRect();
                        return r.width > 0 && r.height > 0; };
                    const pass = Array.prototype.find.call(
                        document.querySelectorAll('input[type=password]'), vis);
                    const form = pass ? (pass.form || pass.closest('form')) : null;
                    if (!form) return 'nopasswordform';
                    const cands = Array.from(form.querySelectorAll(
                        'button, input[type=submit], input[type=button]'));
                    cands.forEach(function (c) {
                        c.disabled = false;
                        c.removeAttribute('aria-disabled');
                    });
                    let pick = cands.find(c =>
                        (c.getAttribute('type') || '').toLowerCase() === 'submit' && vis(c));
                    if (!pick) pick = cands.find(c => vis(c));
                    if (pick) {
                        const desc = '<' + pick.tagName.toLowerCase()
                            + (pick.id ? '#' + pick.id : '') + '> '
                            + ((pick.innerText || pick.value || '') + '').trim().slice(0, 30);
                        pick.scrollIntoView({block: 'center'});
                        try { pick.focus(); } catch (e) {}
                        try { pick.click(); return 'clicked ' + desc; }
                        catch (e) {}
                        try {
                            pick.dispatchEvent(new MouseEvent('click',
                                {bubbles: true, cancelable: true, view: window}));
                            return 'mouseevent ' + desc;
                        } catch (e) {}
                    }
                    if (form.requestSubmit) { form.requestSubmit(); return 'requestSubmit'; }
                    form.submit();
                    return 'form.submit()';
                """)
                if _how and _how != "nopasswordform":
                    self.log(f"SSO credentials submitted via password-form fallback ({_how}).")
                    return True
                self.log("SSO password-form fallback: no visible password form found.")
            except Exception as e:
                self.log(f"SSO password-form fallback failed: {e}")
            # Last resort: ENTER inside the password field — the browser's
            # native implicit-submit path for a credentials form.
            try:
                p = drv.find_element(By.XPATH, "//input[@type='password']")
                from selenium.webdriver.common.keys import Keys as _Keys
                p.send_keys(_Keys.RETURN)
                self.log("SSO credentials submitted via ENTER in the password field.")
                return True
            except Exception:
                return False

        try:
            _fill_credentials()
            if not _submit_credentials():
                self.log("Warning: no login button found after credentials.")
            # Wait for URL change (page navigates away from login) — mirrors VidaPay.
            old_url = drv.current_url
            try:
                from selenium.webdriver.support.ui import WebDriverWait as _WDW
                _WDW(drv, 15).until(lambda d: d.current_url != old_url)
                self.log("URL changed after login.")
            except Exception:
                self.log("URL did not change after login click — continuing.")
            # Invalid-credentials detection + re-entry (extractor login_store).
            for _cred_retry in range(2):
                time.sleep(1.5)
                try:
                    invalid = drv.execute_script("""
                        const items = Array.from(document.querySelectorAll('li, .error, .alert, [class*="error"], [class*="invalid"], [class*="alert"]'));
                        for (const el of items) {
                            const t = (el.innerText || '').toLowerCase();
                            if (t.includes('invalid') || t.includes('incorrect') || t.includes('wrong') || t.includes('failed')) {
                                return el.innerText.trim();
                            }
                        }
                        return null;
                    """)
                except Exception:
                    invalid = None
                if not invalid:
                    break
                self.log(f"Login error detected: '{invalid}' — re-entering credentials (attempt {_cred_retry + 2})...")
                time.sleep(1)
                try:
                    _fill_credentials()
                    _submit_credentials()
                except Exception as ce:
                    self.log(f"Credential retry failed: {ce}")
                    break
            time.sleep(3)  # give SPA time to render heading/body before state machine reads
        except Exception as e:
            raise RuntimeError(f"B2B Step 3 (Username/Password) failed: {e}")

        # Post-login state machine — handles New Sign In, 2FA, Cloudflare, portal.
        # Mirrors VidaPay's finish_setup_steps approach.
        if not _b2b_finish_login_flow(self.driver, stop_event=stop_event, log=self.log, timeout=300):
            raise RuntimeError("B2B login flow did not reach portal — check credentials or complete 2FA.")

    def _is_authed(self) -> bool:
        from selenium.webdriver.common.by import By
        try:
            # Still on login form → not authed
            if self.driver.find_elements(By.ID, "companyId"):
                return False
            if self.driver.find_elements(By.ID, "AccountId"):
                return False
            if self.driver.find_elements(By.ID, "Username"):
                return False
            if self.driver.find_elements(By.ID, "btnSubmit"):
                return False
            if self.driver.find_elements(By.ID, "btnClick"):
                return False
            # URL left the login origin path or page has any content → authed
            url = self.driver.current_url.lower()
            if "wsreports.b2bsoft.com" not in url:
                return False
            body_text = (self.driver.find_element(By.TAG_NAME, "body").text or "").lower()
            if "new sign in" in body_text or "2-factor" in body_text:
                return False
            # No login fields present and on B2B domain → treat as authenticated
            return True
        except Exception:
            return False

    def navigate_to_report(self) -> bool:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        try:
            w = WebDriverWait(self.driver, 15)
            w.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.REPORT_SELECTOR))).click()
            time.sleep(2)
            self.log("✓ Navigated to Inventory Count Result Details")
            return True
        except Exception:
            self.log("Report tree item not found — using current view")
            return False

    def _select_date_range_today(self) -> None:
        """Select 'Today' in the report's TomSelect date-range dropdown.

        The Inventory Count Result Details page defaults the dropdown to
        'Month to Date' (user HTML: div.ts-control#tomselect-1-ts-control,
        item data-value="0"; the dropdown option data-value="2" is Today).
        Exporting without switching produces a month-to-date file
        (Inventory_Count_Result_Details_09012026-09102026.Xlsx) instead of
        today's counts.

        The 16:49 run showed the TomSelect is created only AFTER the
        report's initial data load finishes — the old 10s polling window
        expired before the widget existed ("No date-range dropdown
        found"). Now: poll up to 30s; drive an initialized TomSelect via
        its JS API when present, else set the RAW <select> value (the
        widget adopts it at init) and dispatch change; UI-click fallback
        on .ts-control / the 'Month to Date' item; and when everything
        misses, log the page's <select> inventory so the next run is
        diagnosable from the log alone.
        """
        from selenium.webdriver.common.by import By
        js_set_today = """
            var selects = Array.prototype.slice.call(document.querySelectorAll('select'));
            Array.prototype.slice.call(document.querySelectorAll('.ts-wrapper'))
                .forEach(function(w) {
                    var s = w.querySelector('select') ||
                        (w.parentElement ? w.parentElement.querySelector('select') : null);
                    if (s && selects.indexOf(s) < 0) selects.push(s);
                });
            var candidates = [];
            selects.forEach(function(s) {
                var ts = s.tomselect;
                var opts = Array.prototype.slice.call(s.options || []).map(function(o) {
                    return {v: o.value, t: (o.textContent || '').trim().toLowerCase()};
                });
                if (ts && !opts.some(function(o) { return o.t === 'today'; })) {
                    Object.keys(ts.options || {}).forEach(function(k) {
                        var o = ts.options[k];
                        if (o) opts.push({v: String(o.value), t: String(o.text || '').trim().toLowerCase()});
                    });
                }
                if (!opts.some(function(o) { return o.t === 'today'; })) return;
                var hasM2D = opts.some(function(o) { return o.t.indexOf('month to date') >= 0; });
                candidates.push({sel: s, ts: ts, opts: opts, m2d: hasM2D, init: !!ts});
            });
            if (!candidates.length) return 'no-ts';
            candidates.sort(function(a, b) {
                return ((b.init ? 2 : 0) + (b.m2d ? 1 : 0))
                     - ((a.init ? 2 : 0) + (a.m2d ? 1 : 0));
            });
            var c = candidates[0];
            var chosen = null;
            for (var i = 0; i < c.opts.length; i++) {
                if (c.opts[i].t === 'today') { chosen = c.opts[i]; break; }
            }
            if (!chosen) return 'no-option';
            if (c.ts) {
                if ((c.ts.getValue() + '') === (chosen.v + '')) return 'already';
                c.ts.setValue(chosen.v);
                var item = c.ts.wrapper ? c.ts.wrapper.querySelector('.item') : null;
                return 'ts-set:' + (item ? (item.textContent || '').trim() : chosen.v);
            }
            if ((c.sel.value + '') === (chosen.v + '')) return 'already-raw';
            c.sel.value = chosen.v;
            c.sel.dispatchEvent(new Event('input', {bubbles: true}));
            c.sel.dispatchEvent(new Event('change', {bubbles: true}));
            return 'raw-set:' + chosen.v;
        """
        result = ""
        deadline = time.time() + 30
        while time.time() < deadline:
            try:
                result = str(drv.execute_script(js_set_today) or "")
            except Exception:
                result = ""
            if result and result != "no-ts":
                break
            time.sleep(0.5)
        if result.startswith("ts-set:") or result.startswith("raw-set:"):
            self.log(
                f"✓ Date-range dropdown set to 'Today' ({result}) — "
                "waiting for the grid to refresh.")
            time.sleep(4)
            return
        if result in ("already", "already-raw"):
            self.log("Date-range dropdown already on 'Today'.")
            return
        # Still nothing — log what the page DOES have before the UI
        # fallback, so a variant page is diagnosable from the log alone.
        try:
            inv = drv.execute_script("""
                var sels = Array.prototype.slice.call(document.querySelectorAll('select'));
                var out = [];
                sels.slice(0, 8).forEach(function(s) {
                    var opts = Array.prototype.slice.call(s.options || [])
                        .slice(0, 6).map(function(o) { return (o.textContent || '').trim(); })
                        .filter(function(t) { return t; });
                    out.push((s.id ? '#' + s.id : (s.name ? '[name=' + s.name + ']' : 'select'))
                        + (s.tomselect ? '[ts]' : '[raw]')
                        + (opts.length ? '{' + opts.join(' | ') + '}' : '{}'));
                });
                return {n: sels.length, sels: out,
                        tsCtrl: document.querySelectorAll('.ts-control').length};
            """) or {}
            self.log(
                f"Date-range scan: {inv.get('n')} <select> element(s), "
                f"{inv.get('tsCtrl')} .ts-control: "
                + " ; ".join(inv.get("sels") or []))
        except Exception:
            pass
        # UI fallback. TomSelect opens its dropdown on MOUSEDOWN, not click
        # (user markup: .ts-control aria-expanded="false" and .ts-dropdown
        # style="display:none") — a bare synthetic .click() never opens it,
        # so the options never render. Dispatch the full mousedown/mouseup/
        # click sequence on the control, give the JS-API route one more
        # chance (opening can finish lazy init), then click the visible
        # 'Today' option — data-value="2", class="option", id
        # #tomselect-1-opt-3 in the user's markup.
        def _ts_full_click(el) -> None:
            drv.execute_script(
                "const el = arguments[0];"
                "const o = {bubbles: true, cancelable: true, view: window};"
                "el.dispatchEvent(new MouseEvent('mousedown', o));"
                "el.dispatchEvent(new MouseEvent('mouseup', o));"
                "el.dispatchEvent(new MouseEvent('click', o));",
                el)
        try:
            ctrl = None
            try:
                ctrl = drv.find_element(By.CSS_SELECTOR, ".ts-control")
            except Exception:
                for el in drv.find_elements(By.XPATH,
                        "//*[normalize-space(text())='Month to Date']"):
                    try:
                        if el.is_displayed():
                            ctrl = el
                            break
                    except Exception:
                        continue
            if ctrl is None:
                self.log("⚠ No date-range dropdown found — exporting with the current range.")
                return
            _ts_full_click(ctrl)
            time.sleep(0.8)
            try:
                result = str(drv.execute_script(js_set_today) or "")
            except Exception:
                result = ""
            if result.startswith("ts-set:") or result.startswith("raw-set:"):
                self.log(f"✓ Date-range dropdown set to 'Today' ({result}) — waiting for the grid to refresh.")
                time.sleep(4)
                return
            clicked = False
            opts = drv.find_elements(By.CSS_SELECTOR,
                    ".ts-dropdown .option, [id*='ts-dropdown'] .option")
            if not opts:
                # Options can render outside the expected classes — fall
                # back to any data-value element ('Today' carries
                # data-value="2"; the Month-to-Date item data-value="0"
                # never matches the 'today' text check).
                opts = drv.find_elements(By.CSS_SELECTOR, "[data-value]")
            for opt in opts:
                if "today" in (opt.text or "").strip().lower():
                    _ts_full_click(opt)
                    clicked = True
                    break
            if not clicked:
                self.log("⚠ 'Today' option not found in the date-range dropdown — exporting with the current range.")
                return
            time.sleep(1.2)
            try:
                chk = str(drv.execute_script(js_set_today) or "")
            except Exception:
                chk = ""
            if chk in ("already", "already-raw") or chk.startswith("ts-set:") or chk.startswith("raw-set:"):
                self.log(f"✓ Date-range dropdown set to 'Today' (UI click, {chk}) — waiting for the grid to refresh.")
                time.sleep(4)
            else:
                self.log(f"⚠ Clicked the Today option but the range check returned '{chk or 'no response'}' — exporting with the current range.")
        except Exception:
            self.log("⚠ No date-range dropdown found — exporting with the current range.")

    def download_xlsx(self, timeout: int = 45) -> Optional[Path]:
        """Try to trigger an XLSX download and return the file path."""
        from selenium.webdriver.common.by import By
        # Switch the report's date-range dropdown from 'Month to Date' to
        # 'Today' BEFORE triggering any export — the export honors whatever
        # range is currently selected.
        self._select_date_range_today()
        # Snapshot name → mtime: the scheduler re-exports the same file name
        # every cycle, so a re-download must be detected by mtime, not by a
        # new file name.
        existing = {p: p.stat().st_mtime for p in self.download_dir.glob("*.xlsx")}
        _export_started = time.time()
        # Try visible export/download buttons
        try:
            for text in ["count sheet", "countsheet", "download", "export", "xlsx", "excel"]:
                btns = self.driver.find_elements(By.XPATH,
                    f"//button[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                    f"'abcdefghijklmnopqrstuvwxyz'),'{text}')] | "
                    f"//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ',"
                    f"'abcdefghijklmnopqrstuvwxyz'),'{text}')]")
                for btn in btns:
                    try:
                        self.driver.execute_script("arguments[0].click();", btn)
                        self.log(f"Clicked export button: {text}")
                        time.sleep(3)
                        break
                    except Exception:
                        continue
        except Exception:
            pass
        # Try JS widget export
        try:
            src = self.driver.page_source
            m = re.search(r"window\['(Widget_\d+)'\]", src)
            if m:
                widget = m.group(1)
                self.driver.execute_script(f"window['{widget}'].WidgetPages.ExportReport('Xlsx', 1)")
                self.log("Triggered JS widget XLSX export")
                time.sleep(5)
        except Exception:
            pass
        # Wait for file — a new name OR the same name rewritten with a fresh
        # mtime (the scheduler exports the same file name on every cycle).
        deadline = time.time() + timeout
        while time.time() < deadline:
            for p in self.download_dir.glob("*.xlsx"):
                try:
                    mtime = p.stat().st_mtime
                except Exception:
                    continue
                if p not in existing or mtime > _export_started - 1:
                    self.log(f"✓ Downloaded: {p.name}")
                    return p
            time.sleep(1)
        self.log("No XLSX download detected within timeout")
        return None

    def quit(self):
        if getattr(self, "_using_debug_port", False):
            self.driver = None
            return
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.driver = None


# ─────────────────────────────────────────────────────────────────────────────
# Timesheet Scraper (gfh-telecom-app.web.app/timesheet)
# ─────────────────────────────────────────────────────────────────────────────
class TimesheetScraper:
    """Selenium scraper for the GFH Timesheet portal (Firebase auth)."""
    PORTAL_URL = "https://gfh-telecom-app.web.app/timesheet"

    def __init__(self, email: str, password: str,
                 download_dir: Path, log_fn=None):
        self.email = email
        self.password = password
        self.download_dir = Path(download_dir)
        self.log = log_fn or (lambda m: None)
        self.driver = None

    def _make_driver(self):
        self._using_debug_port = False
        # Try user's running Edge profile first (port 9227)
        try:
            self.driver = _edge_debug_driver()
            self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
                "behavior": "allow",
                "downloadPath": str(self.download_dir),
            })
            _find_or_open_tab(self.driver, _GFH_APP_URL)
            self._using_debug_port = True
            return
        except Exception:
            pass
        # Fallback: new Edge or Chrome instance
        try:
            from selenium import webdriver
            from selenium.webdriver.edge.options import Options as EdgeOptions
            opts = EdgeOptions()
            prefs = {
                "download.default_directory": str(self.download_dir),
                "download.prompt_for_download": False,
            }
            opts.add_experimental_option("prefs", prefs)
            self.driver = webdriver.Edge(options=opts)
        except Exception:
            from selenium import webdriver
            from selenium.webdriver.chrome.options import Options as ChromeOptions
            opts = ChromeOptions()
            prefs = {
                "download.default_directory": str(self.download_dir),
                "download.prompt_for_download": False,
            }
            opts.add_experimental_option("prefs", prefs)
            self.driver = webdriver.Chrome(options=opts)

    def _is_ts_logged_in(self) -> bool:
        """Return True if the TS tab shows no login form (Firebase session still active)."""
        from selenium.webdriver.common.by import By
        try:
            fields = self.driver.find_elements(
                By.CSS_SELECTOR, "input[type='email'], input[name='email'], #email")
            return not any(f.is_displayed() for f in fields)
        except Exception:
            return False

    def login(self) -> bool:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        if self.driver is None:
            self._make_driver()
        self.download_dir.mkdir(parents=True, exist_ok=True)
        # Switch to TS tab — prevents overwriting B2B tab.
        _find_or_open_tab(self.driver, self.PORTAL_URL)
        self.log(f"Opening {self.PORTAL_URL}")
        self.driver.get(self.PORTAL_URL)
        time.sleep(3)
        # If Firebase session active, skip login form entirely.
        if self._is_ts_logged_in():
            self.log("✓ Timesheet already logged in — skipping credentials.")
            return True
        # Not logged in — require credentials.
        if not (self.email and self.password):
            raise RuntimeError("Timesheet credentials incomplete — fill Portal Credentials tab.")
        wait = WebDriverWait(self.driver, 30)
        try:
            email_field = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "input[type='email'], input[name='email'], #email")))
            email_field.clear()
            email_field.send_keys(self.email)
            self.log(f"Entered email: {self.email}")
            pw_field = self.driver.find_element(
                By.CSS_SELECTOR, "input[type='password'], input[name='password'], #password")
            pw_field.clear()
            pw_field.send_keys(self.password)
            self.log("Password entered.")
            submit = self.driver.find_element(
                By.CSS_SELECTOR, "button[type='submit'], input[type='submit'], .login-btn, #loginBtn")
            submit.click()
            time.sleep(4)
            self.log("✓ Timesheet login submitted")
            return True
        except Exception as e:
            raise RuntimeError(f"Timesheet login failed: {e}")

    def download_xlsx(self, timeout: int = 60) -> Optional[Path]:
        """Click 'Today' filter then 'Export Excel' to download today's timesheet."""
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        # Ensure we are on the TS tab (B2B may have been active).
        _find_or_open_tab(self.driver, self.PORTAL_URL)
        # Navigate to timesheet route explicitly in case login redirected elsewhere.
        if not self.driver.current_url.startswith("https://gfh-telecom-app.web.app/timesheet"):
            self.driver.get(self.PORTAL_URL)
            time.sleep(4)
        else:
            time.sleep(2)  # let SPA finish rendering after tab switch

        # Snapshot name → mtime: the timesheet export writes the same file
        # name on every cycle, so a re-download must be detected by mtime.
        existing = {p: p.stat().st_mtime for p in self.download_dir.glob("*.xlsx")}
        _export_started = time.time()

        # Click "Today" date filter button — try up to 3 times (SPA may still render)
        _today_xpaths = [
            "//button[normalize-space(.)='Today']",
            "//button[contains(normalize-space(.),'Today') and not(contains(normalize-space(.),'Yesterday'))]",
            "//*[@role='button' and normalize-space(.)='Today']",
        ]
        _today_clicked = False
        for _attempt in range(3):
            for _xp in _today_xpaths:
                try:
                    _btn = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, _xp)))
                    try:
                        _btn.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", _btn)
                    self.log("Clicked 'Today' filter on timesheet.")
                    _today_clicked = True
                    break
                except Exception:
                    continue
            if _today_clicked:
                break
            time.sleep(2)
        if not _today_clicked:
            self.log("Warning: could not click 'Today' button — proceeding to export anyway.")
        else:
            time.sleep(3)

        # Click "Export Excel" (first match — not "Export B2B Hours" or "Export Bi-Weekly")
        _export_xpaths = [
            "//button[contains(normalize-space(.), 'Export Excel') and "
            "not(contains(normalize-space(.), 'B2B')) and "
            "not(contains(normalize-space(.), 'Bi-Weekly'))]",
            "//button[normalize-space(.)='Export Excel']",
        ]
        _export_clicked = False
        for _attempt in range(3):
            for _xp in _export_xpaths:
                try:
                    _btn = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, _xp)))
                    try:
                        _btn.click()
                    except Exception:
                        self.driver.execute_script("arguments[0].click();", _btn)
                    self.log("Clicked 'Export Excel' on timesheet.")
                    _export_clicked = True
                    break
                except Exception:
                    continue
            if _export_clicked:
                break
            time.sleep(2)
        if not _export_clicked:
            self.log("Warning: could not click 'Export Excel'.")

        deadline = time.time() + timeout
        while time.time() < deadline:
            for p in self.download_dir.glob("*.xlsx"):
                try:
                    mtime = p.stat().st_mtime
                except Exception:
                    continue
                if p not in existing or mtime > _export_started - 1:
                    self.log(f"✓ Timesheet downloaded: {p.name}")
                    return p
            time.sleep(1)
        self.log("No timesheet XLSX downloaded within timeout")
        return None

    def quit(self):
        if getattr(self, "_using_debug_port", False):
            self.driver = None
            return
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        self.driver = None


# ─────────────────────────────────────────────────────────────────────────────
# Audit Scheduler
# ─────────────────────────────────────────────────────────────────────────────
class AuditScheduler:
    """Per-district time-based audit scheduler.

    After the configured start time, sends the starting WhatsApp message,
    then every POLL_INTERVAL_MIN minutes: re-exports B2B + timesheet,
    reloads variances, and checks district completion.

    States: idle → running → held → running (resume) → stopped.
    """
    POLL_INTERVAL_MIN = 15

    def __init__(self, app: "MainWindow"):
        self.app = app
        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._hold_event = threading.Event()  # cleared = held, set = running
        self._hold_event.set()
        self._state = "idle"   # idle | running | held | stopped

    @property
    def state(self) -> str:
        return self._state

    def start(self, district_times: dict, stop_time=None) -> None:
        """district_times: {district_name: 'HH:MM'} or '' to start immediately.
        stop_time: datetime when scheduler auto-stops (None = run indefinitely)."""
        if self._state == "running":
            return
        self._stop_event.clear()
        self._hold_event.set()
        self._state = "running"
        self._thread = threading.Thread(
            target=self._run_loop, args=(district_times, stop_time), daemon=True, name="AuditScheduler")
        self._thread.start()
        self.app._log_scheduler(f"Scheduler started for {len(district_times)} districts.")

    def hold(self) -> None:
        if self._state == "running":
            self._hold_event.clear()
            self._state = "held"
            self.app._log_scheduler("Scheduler paused (Hold).")

    def resume(self) -> None:
        if self._state == "held":
            self._hold_event.set()
            self._state = "running"
            self.app._log_scheduler("Scheduler resumed.")

    def stop(self) -> None:
        self._stop_event.set()
        self._hold_event.set()  # unblock hold
        self._state = "stopped"
        self.app._log_scheduler("Scheduler stopped.")

    def _run_loop(self, district_times: dict, stop_time=None) -> None:
        import datetime as _dt
        # fired[district] = {"start_ts": float, "export_done": bool, "reminders_sent": int, "final_sent": bool}
        fired: dict = {}
        last_reextract: float = time.time()  # init to now so first re-extract waits full RE_EXTRACT_SEC
        INITIAL_EXPORT_SEC = 15 * 60   # 15 min after district starts
        RE_EXTRACT_SEC = 30 * 60       # re-extract B2B + GFH every 30 min
        REMINDER_DELAYS = [30 * 60, 60 * 60, 90 * 60]  # 30, 60, 90 min after start

        while not self._stop_event.is_set():
            self._hold_event.wait()
            if self._stop_event.is_set():
                break
            now = _dt.datetime.now()
            now_ts = time.time()

            # Check global stop time
            if stop_time and now >= stop_time:
                self.app._log_scheduler("⏹ Stop time reached. Sending final results.")
                for district in list(fired.keys()):
                    state = fired[district]
                    if not state.get("final_sent"):
                        state["final_sent"] = True
                        self.app._ui_post(lambda d=district: self.app._auto_send_final_result(d))
                break

            # Fire district starts
            for district, start_str in list(district_times.items()):
                if district in fired:
                    continue
                if start_str:
                    try:
                        h, m = map(int, start_str.split(":"))
                        target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                        if now < target:
                            continue
                    except Exception:
                        pass
                fired[district] = {"start_ts": now_ts, "export_done": False, "reminders_sent": 0, "final_sent": False}
                self.app._ui_post(lambda d=district: self.app._scheduler_start_district(d))

            # Per-district: initial export at 15 min, then reminders at 30/60/90 min
            for district, state in list(fired.items()):
                elapsed = now_ts - state["start_ts"]

                # Initial export + variance image at 15 min
                if not state.get("export_done") and elapsed >= INITIAL_EXPORT_SEC:
                    state["export_done"] = True
                    self.app._ui_post(lambda d=district: self.app._scheduler_run_export_cycle(district=d))

                # Reminders at 30, 60, 90 min
                reminders_sent = state.get("reminders_sent", 0)
                for i, delay in enumerate(REMINDER_DELAYS):
                    if reminders_sent <= i and elapsed >= delay:
                        state["reminders_sent"] = i + 1
                        reminder_num = i + 1
                        self.app._ui_post(lambda d=district, r=reminder_num: self.app._auto_send_reminder(d, r))
                        break

                # Final result at reminder 3 + 30 min (t+120 min) if not yet sent
                if not state.get("final_sent") and elapsed >= 120 * 60:
                    state["final_sent"] = True
                    self.app._ui_post(lambda d=district: self.app._auto_send_final_result(d))

            # Global re-extract B2B + GFH every 30 min (only after at least one district started)
            if fired and (now_ts - last_reextract) >= RE_EXTRACT_SEC:
                self.app._ui_post(self.app._scheduler_run_export_cycle)
                last_reextract = now_ts

            time.sleep(10)
        self._state = "stopped"



def _force_maximized(window):
    """Always open the window MAXIMIZED (user request) and KEEP it maximized
    while the window maps. Qt honors setWindowState before show(), but the
    maximized state is re-asserted several times across the first second
    (port of the Tk build's after(_delay, _zoom) re-assert loop)."""
    def _zoom():
        try:
            window.showMaximized()
        except Exception:
            try:
                window.setWindowState(window.windowState() | Qt.WindowMaximized)
            except Exception:
                pass
    _zoom()
    for _delay in (10, 80, 250, 600):
        try:
            QTimer.singleShot(_delay, _zoom)
        except Exception:
            pass


# ==========================================================
# QT UI LAYER
# ==========================================================
# The Tkinter UI class becomes MainWindow(QMainWindow). Same window
# title, same tab names/order, same button labels, same dialogs
# (Toplevel -> QDialog, messagebox -> QMessageBox,
# filedialog -> QFileDialog), same behaviors; worker-thread UI updates go
# through UIBridge(QObject) signals only.
from PySide6.QtCore import Qt, QObject, Signal, QTimer
from PySide6.QtGui import QIcon, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox, QRadioButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QTabWidget, QGroupBox, QSplitter, QMessageBox, QFileDialog,
    QInputDialog, QPlainTextEdit,
)

from header_manager_qt import FixedHeaderManagerQt
from native_titlebar_qt import sync_titlebar_qt


def _extract_embedded_icon(b64, filename):
    try:
        if not b64:
            return None
        import base64 as _b64
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None


def _set_window_icon_qt(window):
    """Set taskbar + titlebar icon - same resolution order as the Tkinter
    build's window iconbitmap chain: _MEIPASS ICO files first, then the
    .ico shipped beside the script, then the embedded base64 icon decoded
    to %TEMP%."""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        for ico_name in ("gfh_icon.ico", "gfh_telecom_llc_icon.ico"):
            ico_path = os.path.join(meipass, ico_name)
            if os.path.exists(ico_path):
                window.setWindowIcon(QIcon(ico_path))
                return
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, "gfh_icon.ico")
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    ico_path = os.path.join(base_dir, "gfh_telecom_llc_icon.ico")
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    try:
        ico_path = _extract_embedded_icon(EMBEDDED_ICON_B64, "gfh_audit_icon.ico")
        if ico_path:
            window.setWindowIcon(QIcon(ico_path))
    except Exception:
        pass


class UIBridge(QObject):
    """Thread-safe UI update delivery (scheduler / scraper / WhatsApp send
    worker threads -> GUI thread). Replaces the Tkinter build's direct
    (technically unsafe) cross-thread after()/StringVar.set/messagebox/
    Treeview calls."""
    status = Signal(str)
    dialog = Signal(str, str, str)          # kind: info|warning|error
    refresh_audit = Signal()
    refresh_status = Signal()
    post = Signal(object)                   # callable -> run on GUI thread


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        _vp_hide(self)
        self.setWindowTitle(APP_NAME + " — Timesheet Edition")
        self._app_icon = None
        # Dynamic screen resolution support: minimum size for Windows Snap,
        # then a true maximized state (re-asserted while mapping).
        self._apply_dynamic_geometry()
        _force_maximized(self)
        # Set the window/taskbar icon. Same resolution order as the Tkinter
        # build: _MEIPASS .ico first, the script-side .ico second, the
        # embedded base64 icon last.
        ensure_app_icon_files()
        _set_window_icon_qt(self)
        self.zoom_scale = 1.0
        APP_DIR.mkdir(parents=True, exist_ok=True)
        IMAGE_DIR.mkdir(parents=True, exist_ok=True)
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        migrate_legacy_app_data_if_needed()
        self.db = VarianceDatabase(DB_PATH)
        # Track file mtime for cross-instance sync.
        try:
            self._db_mtime: float = DB_PATH.stat().st_mtime if DB_PATH.exists() else 0.0
        except Exception:
            self._db_mtime = 0.0
        self._db_sync_paused: bool = False   # paused while THIS instance is writing
        self._db_sync_after_id = None        # cancellable after() handle
        self.master_store_records = self.db.store_master_records()
        self.current_inventory_records: List[Dict[str, str]] = []
        self.current_time_sheet_records: List[Dict[str, str]] = []

        # tk.StringVar/BooleanVar -> plain Python attributes (R4)
        self.store_district: str = ""
        self.store_name: str = ""
        self.rep_name: str = ""
        self.rep_phone: str = ""
        self.dm_district: str = ""
        self.dm_name: str = ""
        self.dm_phone: str = ""
        self.wg_district: str = ""
        self.wg_group_name: str = ""
        self.exclusion_district: str = ""
        self.exclusion_product: str = ""
        self.exclusion_imei: str = ""
        self.exclusion_comments: str = ""
        self.selected_exclusion_key: str = ""
        # File paths (Tk StringVars with no Entry widget in this build —
        # the scheduler fills them programmatically)
        self.inventory_path: str = ""
        self.time_sheet_path: str = ""
        self.status_text: str = "Select the Inventory_Count_Result_Details file, then click Load Variances. No data loaded yet."

        # ── Portal Credentials (stored in the existing SQLite DB) ───────────
        self.brs_company_id: str = "9909129"
        self.brs_account_id: str = ""
        self.brs_username: str = ""
        self.brs_password: str = ""
        self.ts_email: str = ""
        self.ts_password: str = ""

        # ── Scheduler state ─────────────────────────────────────────────────
        self._sched_time_vars: Dict[str, str] = {}   # district -> HH:MM value
        self._sched_ampm_vars: Dict[str, str] = {}   # district -> AM/PM value
        self._sched_log_var: str = "Scheduler idle."
        self._auto_import_done = False
        self.summary_text: str = "No data loaded"
        self.include_cleared: bool = False
        self.send_only_unsent: bool = SEND_ONLY_UNSENT_BY_DEFAULT

        self.loaded_keys: set = set()
        self.data_loaded = False
        self.key_by_iid: Dict[str, str] = {}

        self.status_rows: List[InventoryStatusRow] = []
        self.status_row_by_key: Dict[str, InventoryStatusRow] = {}
        self.status_key_by_iid: Dict[str, str] = {}
        self.status_checked_keys: set = set()
        self.audit_checked_keys: set = set()

        self._bridge = UIBridge()
        self._bridge.status.connect(self._set_status_label)
        self._bridge.dialog.connect(self._show_dialog)
        self._bridge.refresh_audit.connect(self.refresh_table)
        self._bridge.refresh_status.connect(self.refresh_status_table)
        self._bridge.post.connect(lambda fn: self._safe_post(fn))

        self._scheduler = AuditScheduler(self)
        self._load_saved_credentials()

        self.theme_manager = ThemeManagerQt(default="dark", app_name="VidaPay-GFH")
        # Always start dark — override any saved light preference
        self.theme_manager.current_theme = "dark"

        # ── Outer layout: header (navy) / body / footer copyright — exact
        # FixedHeaderManagerQt order required by the port pattern.
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header_mgr = FixedHeaderManagerQt(outer, title=APP_NAME)
        self.header_mgr.set_logo(logo_path=str(HEADER_LOGO_PATH), text="GFH")
        self.header_mgr.add_theme_toggle(self.theme_manager, callback=self._apply_theme)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(14, 12, 14, 10)
        body_layout.setSpacing(6)
        outer.addWidget(body, 1)

        self._build_ui_body(body_layout)

        self.header_mgr.add_copyright(self.theme_manager, outer)   # LAST

        self.set_status(f"Ready. Data folder: {APP_DIR}")
        self._start_db_sync_poll()
        self._apply_row_tag_colors()
        self._apply_theme()
        self.resize(1400, 900)
        # The Tk app opens maximized — replicate.
        QTimer.singleShot(0, self.showMaximized)

    def _ui_post(self, fn) -> None:
        """Run `fn` on the GUI thread (Tk app.after(0, fn) port). Safe to
        call from any worker thread — marshaled via the UIBridge."""
        self._bridge.post.emit(fn)

    def _safe_post(self, fn) -> None:
        try:
            fn()
        except Exception:
            pass

    # ── Window state save/restore for WhatsApp sending ──────────────────
    # When WhatsApp Desktop is brought to the foreground, Windows can push
    # this window behind it. These methods save the window geometry+state
    # before WhatsApp is focused and restore it after (Qt port of the Tk
    # geometry/state() pair).
    def _save_window_state(self) -> dict:
        """Save the current window geometry and state (maximized/normal)
        so it can be restored after WhatsApp steals focus."""
        try:
            return {
                "geometry": self.geometry(),
                "state": self.windowState(),
            }
        except Exception:
            return {}

    def _restore_window_state(self, saved: dict) -> None:
        """Restore window geometry + state saved by _save_window_state.
        Called after WhatsApp finishes sending — brings the app back to
        the exact position/size/state it was in before."""
        if not saved:
            return
        try:
            state = saved.get("state", Qt.WindowNoState)
            geom = saved.get("geometry", None)
            if state & Qt.WindowMaximized:
                self.setWindowState(Qt.WindowMaximized)
            elif geom is not None:
                self.setGeometry(geom)
            # Bring the app back to the foreground
            try:
                self.raise_()
                self.activateWindow()
            except Exception:
                pass
        except Exception:
            pass

    def _apply_dynamic_geometry(self) -> None:
        """Windows Snap-friendly minimum size; the true maximize is applied
        by _force_maximized() (Qt port of the Tk build's work-area geometry
        fill; the window stays resizable so Snap keeps working)."""
        try:
            screen = QApplication.primaryScreen()
            geo = screen.availableGeometry() if screen is not None else None
            if geo is not None:
                sw, sh = geo.width(), geo.height()
                self.setMinimumSize(min(1220, max(560, sw // 2)),
                                    min(740, max(420, sh // 2)))
        except Exception:
            pass

    def zoom_in(self, event=None) -> None:
        if self.zoom_scale < 2.0:
            self.zoom_scale = round(self.zoom_scale + 0.1, 1)
            self.apply_zoom()

    def zoom_out(self, event=None) -> None:
        if self.zoom_scale > 0.5:
            self.zoom_scale = round(self.zoom_scale - 0.1, 1)
            self.apply_zoom()

    def apply_zoom(self) -> None:
        """Scale the UI font (the Tk build scaled ttk style fonts; the Qt
        build appends one app-wide font-size rule on top of the theme QSS)."""
        self._apply_zoom_qss()

    def _apply_zoom_qss(self) -> None:
        app = QApplication.instance()
        if app is None:
            return
        try:
            size = max(6, round(13 * self.zoom_scale))
            base = re.sub(r"\n?/\* gfh-zoom \*/ QWidget \{ font-size: \d+px; \}", "",
                          app.styleSheet() or "")
            app.setStyleSheet(base + f"\n/* gfh-zoom */ QWidget {{ font-size: {size}px; }}")
        except Exception:
            pass

    def _build_ui_body(self, body_layout) -> None:
        """Zoom shortcuts + notebook (8 tabs) + status bar. Same widget
        labels and order as the Tkinter _build_ui()."""
        # Zoom keybindings (zoom buttons remain in the header area)
        sc_zoom_in = QShortcut(QKeySequence("Ctrl+="), self)
        sc_zoom_in.activated.connect(self.zoom_in)
        sc_zoom_plus = QShortcut(QKeySequence("Ctrl++"), self)
        sc_zoom_plus.activated.connect(self.zoom_in)
        sc_zoom_out = QShortcut(QKeySequence("Ctrl+-"), self)
        sc_zoom_out.activated.connect(self.zoom_out)

        self.tabs = QTabWidget()
        body_layout.addWidget(self.tabs, 1)
        self.status_tab = QWidget()
        self.audit_tab = QWidget()
        self.store_tab = QWidget()
        self.rep_tab = QWidget()
        self.dm_tab = QWidget()
        self.exclusion_tab = QWidget()
        self.credentials_tab = QWidget()
        self.scheduler_tab = QWidget()
        self.tabs.addTab(self.status_tab, "Inventory Audit Status")
        self.tabs.addTab(self.audit_tab, "Variance Audit")
        self.tabs.addTab(self.store_tab, "Store List")
        self.tabs.addTab(self.rep_tab, "Employees")
        self.tabs.addTab(self.dm_tab, "District DMs")
        self.tabs.addTab(self.exclusion_tab, "Excluded Devices")
        self.tabs.addTab(self.credentials_tab, "Portal Credentials")
        # Audit Scheduler tab removed — controls embedded in Inventory Audit Status tab

        self._build_status_tab()
        self._build_audit_tab()
        self._build_store_tab()
        self._build_rep_tab()
        self._build_dm_tab()
        self._build_exclusion_tab()
        self._build_credentials_tab()
        self._build_scheduler_tab()
        self.refresh_store_accounts_table()
        self.refresh_sales_reps_table()
        self.refresh_district_managers_table()
        self.refresh_whatsapp_groups_table()
        self.refresh_device_exclusions_table()

        # Status bar — fixed light bar with navy text, same colors as the Tk build.
        self.status_bar = QLabel(self.status_text)
        self.status_bar.setStyleSheet("background:#E9ECF5; color:#090d26; padding:6px;")
        body_layout.addWidget(self.status_bar)

    def _apply_theme(self, colors=None) -> None:
        """Apply theme colors to all widgets (Qt: one app-wide QSS;
        native title bars re-synced inside apply_theme_to_app)."""
        colors = self.theme_manager.get_colors()
        self.theme_manager.apply_theme_to_app(QApplication.instance())
        self.header_mgr.update_for_theme()
        # Re-apply native titlebar sync explicitly (apply_theme_to_app also
        # walks all top-level widgets; this pins the main window).
        try:
            sync_titlebar_qt(self, dark=(self.theme_manager.current_theme == "dark"))
        except Exception:
            pass
        self.COLOR_BG = colors["bg"]
        self.COLOR_CARD = colors["panel"]
        self.COLOR_TEXT = colors["text"]
        self.COLOR_MUTED = colors["text_dim"]
        self.COLOR_BORDER = colors["border"]
        self.COLOR_INPUT = colors.get("input", colors["panel"])
        self.COLOR_PANEL_ALT = colors.get("panel_alt", colors["panel"])
        # Row-highlight colors for the two audit tables MUST follow the
        # active theme (same palette as the Tk _apply_row_tag_colors).
        self._apply_row_tag_colors()
        self._apply_zoom_qss()
        # Recolor already-rendered rows with the new palette.
        try:
            if self.data_loaded:
                self.refresh_status_table()
                self.refresh_table()
        except Exception:
            pass

    def _apply_row_tag_colors(self) -> None:
        """Theme-aware row highlight colors for the two audit tables.

        Light theme keeps the classic pastel highlights:
            yellow = pending, blue = sent / completed-after-update, green = cleared.
        Dark theme swaps them for muted dark tints of the same families with
        light text (same values as the Tkinter tag_configure calls).
        """
        dark = getattr(self.theme_manager, "current_theme", "dark") == "dark"
        if dark:
            pending = ("#3a3117", "#ffd66e")   # dark yellow tint
            sent = ("#16324f", "#9ecbff")      # dark blue tint
            cleared = ("#17331f", "#8fe3a8")   # dark green tint
        else:
            pending = ("#FFF3CD", "#111827")
            sent = ("#D7ECFF", "#111827")
            cleared = ("#D9F7DF", "#111827")
        colors = self.theme_manager.get_colors()
        normal_bg = getattr(self, "COLOR_CARD", None) or colors.get("panel", ("#141b38" if dark else "#ffffff"))
        normal_fg = getattr(self, "COLOR_TEXT", None) or colors.get("text", ("#e8ecf7" if dark else "#111827"))
        self._row_tag_colors = {
            "status_pending": (pending[0], pending[1]),
            "status_completed_after_update": (sent[0], sent[1]),
            "status_completed_sent": (cleared[0], cleared[1]),
            "status_completed": (normal_bg, normal_fg),
            "variance_pending": (pending[0], pending[1]),
            "variance_sent": (sent[0], sent[1]),
            "variance_cleared": (cleared[0], cleared[1]),
        }

    def _paint_row(self, table, row: int, tag: str) -> None:
        """Apply the theme colors of a Tk row-tag to every item of one
        QTableWidget row."""
        colors = getattr(self, "_row_tag_colors", {})
        bg, fg = colors.get(tag, (None, None))
        if bg is None:
            return
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if item is not None:
                item.setBackground(QColor(bg))
                item.setForeground(QColor(fg))

    def _make_table(self, headers, widths, multi_select=False) -> QTableWidget:
        """ttk.Treeview(show='headings') -> QTableWidget. Same headings,
        same column widths, read-only, row selection, Qt-native scrollbars."""
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection if multi_select
                               else QAbstractItemView.SingleSelection)
        table.setWordWrap(False)
        table.setTextElideMode(Qt.ElideRight)
        for col, width in enumerate(widths):
            table.setColumnWidth(col, width)
        return table

    def _wire_sorting(self, table) -> None:
        """Heading-click sort with reverse toggle on re-click (port of the
        Tkinter sort_any_tree heading command binding)."""
        header = table.horizontalHeader()
        header.setSectionsClickable(True)
        header.setSortIndicatorShown(True)

        def _on_section(logical_index: int) -> None:
            state = getattr(table, "_gfh_sort_reverse", {})
            reverse = not state.get(logical_index, False)
            state[logical_index] = reverse
            table._gfh_sort_reverse = state
            self.sort_any_tree(table, logical_index, reverse)

        header.sectionClicked.connect(_on_section)

    @staticmethod
    def sort_any_tree(table, col: int, reverse: bool) -> None:
        """Sort a QTableWidget by column value (ttk.Treeview heading-sort
        port; repeated clicks reverse the order)."""
        table.sortItems(col, Qt.DescendingOrder if reverse else Qt.AscendingOrder)
        table.horizontalHeader().setSortIndicator(
            col, Qt.DescendingOrder if reverse else Qt.AscendingOrder)

    @staticmethod
    def _combo_set(combo, value: str) -> None:
        """Select `value` in a combo (Tk combobox.set() port); a value not
        in the list clears the selection to blank, same as setting a Tk
        textvariable to "".  """
        combo.setCurrentIndex(combo.findText(value))

    # ── Status / dialog helpers (worker threads -> UIBridge) ────────────
    def set_status(self, text: str) -> None:
        self.status_text = text
        self._bridge.status.emit(text)

    def _set_status_label(self, text: str) -> None:
        self.status_text = text
        self.status_bar.setText(text)
        try:
            self.status_bar.repaint()
        except Exception:
            pass

    def _worker_dialog(self, kind: str, title: str, text: str) -> None:
        """Show a messagebox from a worker thread — marshaled to the GUI
        thread through the UIBridge (Tk messagebox calls in the send
        threads were cross-thread already; Qt requires the GUI thread)."""
        self._bridge.dialog.emit(kind, title, text)

    def _show_dialog(self, kind: str, title: str, text: str) -> None:
        if kind == "error":
            QMessageBox.critical(self, title, text)
        elif kind == "warning":
            QMessageBox.warning(self, title, text)
        else:
            QMessageBox.information(self, title, text)


    # ── Portal Credentials Tab ──────────────────────────────────────────────
    def _build_credentials_tab(self) -> None:
        """Build the Portal Credentials tab UI."""
        tab = self.credentials_tab
        tab_outer = QVBoxLayout(tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)
        # BRS Section
        brs_box = QGroupBox("B2B Soft Portal (wsreports.b2bsoft.com)")
        brs_grid = QGridLayout(brs_box)
        fields_brs = [
            ("Company ID",  "brs_company_id_edit", False),
            ("Account ID",  "brs_account_id_edit", False),
            ("Username",    "brs_username_edit",   False),
            ("Password",    "brs_password_edit",   True),
        ]
        brs_initial = {
            "brs_company_id_edit": self.brs_company_id,
            "brs_account_id_edit": self.brs_account_id,
            "brs_username_edit": self.brs_username,
            "brs_password_edit": self.brs_password,
        }
        for row_i, (label, attr, is_pw) in enumerate(fields_brs):
            brs_grid.addWidget(QLabel(label), row_i, 0)
            edit = QLineEdit(brs_initial[attr])
            edit.setFixedWidth(300)
            if is_pw:
                edit.setEchoMode(QLineEdit.Password)
            setattr(self, attr, edit)
            brs_grid.addWidget(edit, row_i, 1)
        tab_outer.addWidget(brs_box)

        # Timesheet Section
        ts_box = QGroupBox("Timesheet Portal (gfh-telecom-app.web.app)")
        ts_grid = QGridLayout(ts_box)
        fields_ts = [
            ("Email",    "ts_email_edit",    False),
            ("Password", "ts_password_edit", True),
        ]
        ts_initial = {
            "ts_email_edit": self.ts_email,
            "ts_password_edit": self.ts_password,
        }
        for row_i, (label, attr, is_pw) in enumerate(fields_ts):
            ts_grid.addWidget(QLabel(label), row_i, 0)
            edit = QLineEdit(ts_initial[attr])
            edit.setFixedWidth(300)
            if is_pw:
                edit.setEchoMode(QLineEdit.Password)
            setattr(self, attr, edit)
            ts_grid.addWidget(edit, row_i, 1)
        tab_outer.addWidget(ts_box)

        # Save / Test buttons
        btn_row = QWidget()
        brow = QHBoxLayout(btn_row)
        brow.setContentsMargins(0, 0, 0, 0)
        save_btn = QPushButton("💾  Save Credentials")
        save_btn.clicked.connect(self._save_credentials)
        brow.addWidget(save_btn)
        test_brs_btn = QPushButton("🔌  Test B2B Login")
        test_brs_btn.clicked.connect(self._test_brs_login)
        brow.addWidget(test_brs_btn)
        test_ts_btn = QPushButton("🔌  Test Timesheet Login")
        test_ts_btn.clicked.connect(self._test_ts_login)
        brow.addWidget(test_ts_btn)
        brow.addStretch(1)
        tab_outer.addWidget(btn_row)
        tab_outer.addStretch(1)

    def _save_credentials(self) -> None:
        data = {
            "brs": {
                "company_id": self.brs_company_id_edit.text().strip(),
                "account_id": self.brs_account_id_edit.text().strip(),
                "username":   self.brs_username_edit.text().strip(),
                "password":   self.brs_password_edit.text(),
            },
            "timesheet": {
                "email":    self.ts_email_edit.text().strip(),
                "password": self.ts_password_edit.text(),
            },
        }
        try:
            self.db.save_portal_credentials(data)
            QMessageBox.information(self, "Saved", "Portal credentials saved to database.")
        except Exception as exc:
            QMessageBox.critical(self, "Save Error", str(exc))

    def _load_saved_credentials(self) -> None:
        try:
            data = self.db.load_portal_credentials()
            brs = data.get("brs", {})
            # Plain attrs are set first (they seed the tab widgets on build);
            # once the credentials tab exists the edits are updated too.
            self.brs_company_id = brs.get("company_id", "9909129") or "9909129"
            self.brs_account_id = brs.get("account_id", "") or ""
            self.brs_username = brs.get("username", "") or ""
            self.brs_password = brs.get("password", "") or ""
            ts = data.get("timesheet", {})
            self.ts_email = ts.get("email", "") or ""
            self.ts_password = ts.get("password", "") or ""
            if hasattr(self, "brs_company_id_edit"):
                self.brs_company_id_edit.setText(self.brs_company_id)
                self.brs_account_id_edit.setText(self.brs_account_id)
                self.brs_username_edit.setText(self.brs_username)
                self.brs_password_edit.setText(self.brs_password)
                self.ts_email_edit.setText(self.ts_email)
                self.ts_password_edit.setText(self.ts_password)
        except Exception:
            pass

    def _test_brs_login(self) -> None:
        self._save_credentials()
        scraper = B2BSoftScraper(
            company_id=self.brs_company_id_edit.text().strip(),
            account_id=self.brs_account_id_edit.text().strip(),
            username=self.brs_username_edit.text().strip(),
            password=self.brs_password_edit.text(),
            download_dir=EXPORT_DIR,
            log_fn=lambda m: self.set_status(f"[B2B] {m}"),
        )
        def _run():
            try:
                scraper.login()
                self._ui_post(lambda: QMessageBox.information(self, "B2B Login", "✓ Login successful!"))
            except Exception as exc:
                self._ui_post(lambda e=exc: QMessageBox.critical(self, "B2B Login Failed", str(e)))
            finally:
                scraper.quit()
        threading.Thread(target=_run, daemon=True).start()

    def _test_ts_login(self) -> None:
        self._save_credentials()
        scraper = TimesheetScraper(
            email=self.ts_email_edit.text().strip(),
            password=self.ts_password_edit.text(),
            download_dir=EXPORT_DIR,
            log_fn=lambda m: self.set_status(f"[TS] {m}"),
        )
        def _run():
            try:
                scraper.login()
                self._ui_post(lambda: QMessageBox.information(self, "Timesheet Login", "✓ Login successful!"))
            except Exception as exc:
                self._ui_post(lambda e=exc: QMessageBox.critical(self, "Timesheet Login Failed", str(e)))
            finally:
                scraper.quit()
        threading.Thread(target=_run, daemon=True).start()

    # ── Audit Scheduler Tab ─────────────────────────────────────────────────
    def _build_scheduler_tab(self) -> None:
        """Scheduler controls are now on the Inventory Audit Status tab (main tab).
        This tab shows a convenience redirect notice only."""
        tab = self.scheduler_tab
        tab_outer = QVBoxLayout(tab)
        tab_outer.setContentsMargins(20, 30, 20, 20)
        notice = QLabel(
            "The Audit Scheduler controls have moved to the Inventory Audit Status tab.\n\n"
            "Use the Start / Stop / Hold / Resume buttons and district time fields\n"
            "at the top of that tab to control the scheduler."
        )
        notice.setWordWrap(True)
        tab_outer.addWidget(notice, 0, Qt.AlignTop | Qt.AlignLeft)
        go_btn = QPushButton("Go to Inventory Audit Status →")
        go_btn.clicked.connect(lambda: self.tabs.setCurrentWidget(self.status_tab))
        tab_outer.addWidget(go_btn, 0, Qt.AlignTop | Qt.AlignLeft)
        tab_outer.addStretch(1)

    def _build_scheduler_district_rows(self) -> None:
        """Build one row per known district with HH:MM + AM/PM time-entry fields."""
        # Clear the previous rows (Qt: drop every child widget of the grid)
        while self._sched_grid.count():
            item = self._sched_grid.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._sched_time_vars.clear()
        self._sched_ampm_vars.clear()
        self._sched_time_edits = {}
        self._sched_ampm_combos = {}
        districts = self._known_districts_for_scheduler()
        if not districts:
            self._sched_grid.addWidget(
                QLabel("No districts found. Load an inventory file or add stores in the Store List tab."), 0, 0)
            return
        # Load saved times from DB
        import json as _json
        try:
            saved = _json.loads(self.db.get_setting("sched_district_times", "{}"))
        except Exception:
            saved = {}
        for idx, dist in enumerate(districts):
            col = (idx % 3) * 5
            row = idx // 3
            saved_val = saved.get(dist, {})
            self._sched_time_vars[dist] = saved_val.get("time", "")
            self._sched_ampm_vars[dist] = saved_val.get("ampm", "AM")
            self._sched_grid.addWidget(QLabel(dist), row, col)
            time_edit = QLineEdit(saved_val.get("time", ""))
            time_edit.setFixedWidth(60)
            self._sched_grid.addWidget(time_edit, row, col + 1)
            ampm_combo = QComboBox()
            ampm_combo.addItems(["AM", "PM"])
            self._combo_set(ampm_combo, saved_val.get("ampm", "AM"))
            ampm_combo.setFixedWidth(64)
            self._sched_grid.addWidget(ampm_combo, row, col + 2)
            hint = QLabel("HH:MM")
            hint.setStyleSheet("color:#8090b0;")
            self._sched_grid.addWidget(hint, row, col + 3)
            self._sched_time_edits[dist] = time_edit
            self._sched_ampm_combos[dist] = ampm_combo
            # col+4 is the spacer (stretch column)

    def _known_districts_for_scheduler(self) -> List[str]:
        """Return distinct districts from the DB store list."""
        try:
            return sorted(set(self.db.all_known_districts()))
        except Exception:
            return []

    @staticmethod
    def _to_24h(hhmm: str, ampm: str) -> str:
        """Convert HH:MM + AM/PM to HH:MM 24-hour string. Returns '' on bad input."""
        try:
            h, m = map(int, hhmm.strip().split(":"))
            ampm = ampm.strip().upper()
            if ampm == "PM" and h != 12:
                h += 12
            elif ampm == "AM" and h == 12:
                h = 0
            return f"{h:02d}:{m:02d}"
        except Exception:
            return ""

    def _sched_start(self) -> None:
        import datetime as _dt
        times_raw = {d: e.text().strip() for d, e in self._sched_time_edits.items()}
        if not times_raw:
            QMessageBox.warning(self, "No Districts", "No districts to schedule. Import inventory first.")
            return
        # Convert 12hr → 24hr for each district
        times: dict = {}
        for d, t in times_raw.items():
            if t:
                ampm_combo = self._sched_ampm_combos.get(d)
                ampm = ampm_combo.currentText() if ampm_combo is not None else "AM"
                times[d] = self._to_24h(t, ampm)
            else:
                times[d] = ""
        # Late-start: if configured start time already passed, start that district immediately
        now = _dt.datetime.now()
        adjusted_times: dict = {}
        for district, t in times.items():
            if t:
                try:
                    h, m = map(int, t.split(":"))
                    target = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if now >= target:
                        adjusted_times[district] = ""  # start immediately
                        self._log_scheduler(f"⚡ {district}: {t} already passed — starting immediately.")
                    else:
                        adjusted_times[district] = t
                except Exception:
                    adjusted_times[district] = t
            else:
                adjusted_times[district] = t
        # Compute stop time: explicit global stop time overrides duration
        stop_time_str = self._sched_stop_time_edit.text().strip()
        stop_time = None
        if stop_time_str:
            try:
                stop_ampm = self._sched_stop_ampm_combo.currentText()
                t24 = self._to_24h(stop_time_str, stop_ampm)
                h, m = map(int, t24.split(":"))
                stop_time = now.replace(hour=h, minute=m, second=0, microsecond=0)
                if stop_time <= now:
                    stop_time += _dt.timedelta(days=1)
            except Exception:
                pass
        if stop_time is None:
            try:
                dur_h = float(self._sched_duration_edit.text().strip() or "12")
            except Exception:
                dur_h = 12.0
            stop_time = now + _dt.timedelta(hours=dur_h)
        # Tracks districts whose scheduled start has fired this session.
        self._scheduler_fired_districts: set = set()
        # Tracks districts whose starting message has not yet been sent.
        self._scheduler_pending_messages: set = set()
        # Ensures only one export cycle runs at a time — prevents 7-district simultaneous B2B+TS scrapes.
        self._export_cycle_lock = threading.Lock()

        self._scheduler.start(adjusted_times, stop_time=stop_time)
        self._log_scheduler(f"▶ Started. Stops at {stop_time.strftime('%H:%M')}.")
        # Start WhatsApp notification OCR monitor for auto-IMEI clearing
        self._start_whatsapp_ocr_monitor()

        # Open Edge at port 9227 (launch if not running), open monitoring tabs,
        # then run initial export. Starting messages send INSIDE the export
        # cycle completion callback — never before download finishes.
        def _startup_sequence():
            def _log(m):
                self._ui_post(lambda mm=str(m): self._log_scheduler(mm))
            ready = _ensure_edge_open(log=_log)
            if ready:
                _log("✓ Edge ready at port 9227.")
                names = self._sched_open_tabs()
                if names:
                    self._ui_post(lambda n=list(names): self._log_tabs_opened(n))
            else:
                _log("⚠ Could not open Edge at port 9227.")
            self._scheduler_run_export_cycle()
        threading.Thread(target=_startup_sequence, daemon=True, name="SchedulerStartup").start()

    def _sched_save_times(self) -> None:
        import json as _json
        data = {}
        for dist, edit in self._sched_time_edits.items():
            ampm_combo = self._sched_ampm_combos.get(dist)
            data[dist] = {
                "time": edit.text().strip(),
                "ampm": ampm_combo.currentText() if ampm_combo is not None else "AM",
            }
        try:
            self.db.save_setting("sched_district_times", _json.dumps(data))
            self.set_status("Scheduler times saved.")
        except Exception as exc:
            self.set_status(f"Save failed: {exc}")

    def _sched_stop(self) -> None:
        self._scheduler.stop()
        self._wa_ocr_running = False

    def _sched_hold(self) -> None:
        self._scheduler.hold()

    def _sched_resume(self) -> None:
        self._scheduler.resume()

    def _log_scheduler(self, msg: str) -> None:
        """Append a message to the scheduler log text widget."""
        import datetime as _dt
        timestamp = _dt.datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {msg}"
        try:
            self._sched_log_text.appendPlainText(full_msg)
        except Exception:
            pass

    # ── Scheduler callbacks (run on main thread via _ui_post) ───────────────
    def _scheduler_start_district(self, district: str) -> None:
        """Mark district as fired. Starting message sends after export cycle completes."""
        self._log_scheduler(f"▶ Starting district: {district}")
        fired_set = getattr(self, "_scheduler_fired_districts", None)
        if fired_set is not None:
            fired_set.add(district)

    def _sched_open_tabs(self) -> list:
        """Open scheduler tabs in the automation Edge window: B2B, GFH app, and
        WhatsApp Web (only when WhatsApp Web mode is selected). Returns opened names."""
        try:
            mode_var = getattr(self, "wa_mode_var", None)
            include_wa = (mode_var == "web") if mode_var is not None else True
        except Exception:
            include_wa = True
        return open_monitoring_tabs(include_whatsapp=include_wa)

    def _log_tabs_opened(self, names: list) -> None:
        """Log 'Opened X, Y tabs.' — suppresses exact repeats within 90s so the
        scheduler-startup open and the first export cycle don't double-log."""
        msg = "Opened " + _humanize_list(list(names)) + " tabs."
        now = time.monotonic()
        last_msg, last_t = getattr(self, "_last_tabs_log", (None, 0.0))
        if msg == last_msg and (now - last_t) < 90:
            return
        self._last_tabs_log = (msg, now)
        self._log_scheduler(msg)

    def _scheduler_run_export_cycle(self, district: str = None) -> None:
        """Export B2B + Timesheet files, reload variances, optionally send variance image."""
        lock = getattr(self, "_export_cycle_lock", None)
        if lock is not None and not lock.acquire(blocking=False):
            self._log_scheduler("⟳ Export cycle already running — skipped duplicate call.")
            return
        self._log_scheduler("⟳ Running export cycle…")
        def _run():
            # Pre-open separate tabs so each scraper gets its own tab (and a
            # WhatsApp Web tab when WhatsApp Web mode is selected).
            try:
                names = self._sched_open_tabs()
                if names:
                    self._ui_post(lambda n=list(names): self._log_tabs_opened(n))
            except Exception:
                pass

            # Build both scrapers upfront — init tab handles sequentially before scraping.
            brs = B2BSoftScraper(
                company_id=self.brs_company_id_edit.text().strip(),
                account_id=self.brs_account_id_edit.text().strip(),
                username=self.brs_username_edit.text().strip(),
                password=self.brs_password_edit.text(),
                download_dir=EXPORT_DIR,
                log_fn=lambda m: self._ui_post(lambda mm=str(m): self._log_scheduler(f"[B2B] {mm}")),
            )
            brs._make_driver()

            ts = TimesheetScraper(
                email=self.ts_email_edit.text().strip(),
                password=self.ts_password_edit.text(),
                download_dir=EXPORT_DIR,
                log_fn=lambda m: self._ui_post(lambda mm=str(m): self._log_scheduler(f"[TS] {mm}")),
            )
            ts._make_driver()

            inv_file = None
            ts_file = None

            def _run_b2b():
                nonlocal inv_file
                try:
                    brs.login()
                    brs.navigate_to_report()
                    inv_file = brs.download_xlsx()
                except Exception as exc:
                    self._ui_post(lambda e=exc: self._log_scheduler(f"⚠ B2B export error: {e}"))
                finally:
                    brs.quit()

            def _run_ts():
                nonlocal ts_file
                try:
                    ts.login()
                    ts_file = ts.download_xlsx()
                except Exception as exc:
                    self._ui_post(lambda e=exc: self._log_scheduler(f"⚠ Timesheet export error: {e}"))
                finally:
                    ts.quit()

            # Run B2B then Timesheet sequentially — both share the same Edge browser
            # (port 9227). Parallel switch_to.window() calls on the same browser
            # race and corrupt each other's tab focus.
            try:
                _run_b2b()
                _run_ts()
            except Exception as exc:
                self._ui_post(lambda e=exc: self._log_scheduler(f"⚠ Export cycle error: {e}"))
            finally:
                # Reload variances and optionally send variance image — must
                # run on EVERY cycle. This block used to sit inside the except
                # handler, and _run_b2b/_run_ts swallow their own errors, so on
                # a normal successful export load_variances() never ran and the
                # Status tab stayed empty until a manual Load Variances click.
                _district_for_send = district
                def _reload_and_send():
                    try:
                        if inv_file and inv_file.exists():
                            self.inventory_path = str(inv_file)
                            self._log_scheduler(f"Loaded B2B file: {inv_file.name}")
                        if ts_file and ts_file.exists():
                            self.time_sheet_path = str(ts_file)
                            self._log_scheduler(f"Loaded timesheet: {ts_file.name}")
                        # Require both files to call load_variances (avoids missing-file dialog).
                        both_ready = (
                            inv_file and inv_file.exists() and
                            ts_file and ts_file.exists()
                        )
                        if both_ready:
                            self.load_variances()
                            self._auto_import_stores_from_inventory()

                        # Send starting message + status image for every fired district.
                        # This is the ONLY place starting messages send — after export attempt.
                        fired = list(getattr(self, "_scheduler_fired_districts", set()))
                        pending = getattr(self, "_scheduler_pending_messages", None)
                        if pending is None:
                            self._scheduler_pending_messages = set(fired)
                            pending = self._scheduler_pending_messages
                        unsent = [d for d in fired if d in pending]
                        if unsent:
                            # One thread sends all districts sequentially — no concurrent UI fight
                            # (critical for desktop mode; web mode also benefits from sequential tab use).
                            for d in unsent:
                                self._log_scheduler(f"💬 Sending starting message → {d}")
                            normalized = [normalize_district(d) for d in unsent]
                            threading.Thread(
                                target=self._send_starting_message_thread,
                                args=(normalized,),
                                daemon=True,
                                name="StartingMessages",
                            ).start()
                            for d in unsent:
                                if both_ready:
                                    self._auto_send_status_image(d)
                                pending.discard(d)

                        if _district_for_send:
                            self._auto_send_variance_image(_district_for_send)
                    except Exception as exc:
                        self._log_scheduler(f"⚠ Reload error: {exc}")
                self._ui_post(_reload_and_send)
                _lock = getattr(self, "_export_cycle_lock", None)
                if _lock is not None:
                    try:
                        _lock.release()
                    except RuntimeError:
                        pass
        threading.Thread(target=_run, daemon=True, name="ExportCycle").start()

    # ── Scheduler automation helpers (no-dialog versions) ───────────────────

    def _capture_tab_screenshot(self, tab_widget) -> Optional[Path]:
        """Take a screenshot of a specific tab widget and return the saved path."""
        try:
            import datetime as _dt
            ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = EXPORT_DIR / f"tab_screenshot_{ts}.png"
            # Qt-native widget grab (same output as the Tk build's
            # ImageGrab.grab of the widget's screen bounds)
            pixmap = tab_widget.grab()
            pixmap.save(str(out_path))
            return out_path
        except Exception as exc:
            self._log_scheduler(f"⚠ Screenshot error: {exc}")
            return None

    def _auto_send_status_image(self, district: str) -> None:
        """Send Inventory Audit Status image for a district without confirmation dialog."""
        try:
            rows = [row for row in self.status_rows if normalize_district(row.district) == normalize_district(district)]
            if not rows:
                self._log_scheduler(f"⚠ No status rows for district: {district}")
                return
            self._log_scheduler(f"📸 Sending status image → {district} ({len(rows)} rows)")
            threading.Thread(target=self._send_status_rows, args=(rows, "district"), daemon=True).start()
        except Exception as exc:
            self._log_scheduler(f"⚠ Status image error ({district}): {exc}")

    def _auto_send_starting_message(self, district: str) -> None:
        """Send starting message to a district WhatsApp group without confirmation dialog."""
        try:
            self._log_scheduler(f"💬 Sending starting message → {district}")
            districts = [normalize_district(district)]
            threading.Thread(target=self._send_starting_message_thread, args=(districts,), daemon=True).start()
        except Exception as exc:
            self._log_scheduler(f"⚠ Starting message error ({district}): {exc}")

    def _send_actions_panel_screenshot(self, district: str) -> None:
        """Capture the Variance Audit tab and send it to the district WhatsApp group."""
        try:
            screenshot_path = self._capture_tab_screenshot(self.audit_tab)
            if not screenshot_path or not screenshot_path.exists():
                self._log_scheduler(f"⚠ Actions panel screenshot failed for {district}")
                return
            self._log_scheduler(f"📷 Sending actions panel screenshot → {district}")
            def _send():
                try:
                    sender = WhatsAppSender(status_callback=self.set_status,
                                           mode=getattr(self, "wa_mode_var", "web") or "web")
                    group_name = group_name_for_district(district, self.db)
                    _win_state = self._save_window_state()
                    sender.send_image(group_name, screenshot_path, text_message="Audit actions panel.")
                    self._restore_window_state(_win_state)
                    self._log_scheduler(f"✓ Actions panel sent to {group_name}")
                except Exception as exc:
                    self._log_scheduler(f"⚠ Actions panel send error: {exc}")
            threading.Thread(target=_send, daemon=True).start()
        except Exception as exc:
            self._log_scheduler(f"⚠ Actions panel screenshot error ({district}): {exc}")

    def _auto_send_variance_image(self, district: str) -> None:
        """Send all pending variance rows for a district without confirmation dialog."""
        try:
            if not self.data_loaded:
                return
            all_rows = self.db.get_rows_by_keys(self.loaded_keys)
            rows = [row for row in all_rows
                    if normalize_district(row.district) == normalize_district(district) and not row.cleared]
            rows = self.filter_excluded_variance_rows(rows)
            if not rows:
                self._log_scheduler(f"✓ No pending variances for {district}")
                return
            self._log_scheduler(f"📊 Sending variance image → {district} ({len(rows)} rows)")
            threading.Thread(target=self._send_rows, args=(rows, "district", False), daemon=True).start()
        except Exception as exc:
            self._log_scheduler(f"⚠ Variance image error ({district}): {exc}")

    def _auto_send_reminder(self, district: str, reminder_number: int) -> None:
        """Send reminder N to a district without confirmation dialog."""
        reminder_messages = {
            1: "Please clear the pending variances.",
            2: "This is the second reminder. Please clear the pending variances immediately.",
            3: "Final reminder. Uncleared variances will be reported.",
        }
        try:
            if not self.data_loaded:
                return
            all_rows = self.db.get_rows_by_keys(self.loaded_keys)
            uncleared = [row for row in all_rows
                         if normalize_district(row.district) == normalize_district(district) and not row.cleared]
            uncleared = self.filter_excluded_variance_rows(uncleared)
            if not uncleared:
                self._log_scheduler(f"✓ {district}: all variances cleared — skipping reminder {reminder_number}")
                return
            message = reminder_messages.get(reminder_number, "Please clear the pending variances.")
            self._log_scheduler(f"🔔 Sending reminder {reminder_number}/3 → {district}")
            districts = [normalize_district(district)]
            threading.Thread(
                target=self._send_single_reminder_thread,
                args=(districts, uncleared, reminder_number, message), daemon=True).start()
        except Exception as exc:
            self._log_scheduler(f"⚠ Reminder {reminder_number} error ({district}): {exc}")

    def _auto_send_final_result(self, district: str) -> None:
        """Send final district audit result without confirmation dialog."""
        try:
            self._log_scheduler(f"🏁 Sending final result → {district}")
            districts = [normalize_district(district)]
            threading.Thread(target=self._send_final_district_result_thread, args=(districts,), daemon=True).start()
        except Exception as exc:
            self._log_scheduler(f"⚠ Final result error ({district}): {exc}")

    def _start_whatsapp_ocr_monitor(self) -> None:
        """Start background thread that reads WhatsApp notification screenshots via Tesseract OCR
        and auto-marks found IMEI numbers as cleared in the variance DB."""
        if getattr(self, "_wa_ocr_running", False):
            return
        self._wa_ocr_running = True
        threading.Thread(target=self._whatsapp_ocr_entry, daemon=True, name="WhatsAppOCR").start()

    def _whatsapp_ocr_entry(self) -> None:
        """OCR dependency auto-setup (VidaPay Transfer Bot style), then the monitor loop.

        Missing pieces are INSTALLED automatically instead of disabling the monitor:
        - Tesseract binary absent  -> winget, then the official silent installer.
        - pytesseract package      -> bundled in the frozen EXE (module-level import);
                                      auto-pip-installed when running from source.
        The monitor only disables itself if an install genuinely failed.
        """
        global pytesseract, PYTESSERACT_AVAILABLE

        def _olog(m):
            # Thread-safe: this entry runs on the WhatsAppOCR background thread;
            # Qt widgets must only be touched on the main thread.
            self._ui_post(lambda mm=str(m): self._log_scheduler(mm))

        _olog("👁 WhatsApp OCR monitor started.")
        if not _is_tesseract_installed():
            _olog("👁 Tesseract OCR not found — installing automatically…")
            ok = _install_tesseract_binary(lambda m: _olog(f"👁 {m}"))
            if not ok:
                _olog("⚠ Tesseract OCR binary not found and auto-install failed — OCR monitor disabled. "
                      "Install from https://github.com/UB-Mannheim/tesseract/wiki")
                self._wa_ocr_running = False
                return
            _refresh_tesseract_path()
            _olog("👁 Tesseract ready — OCR monitor active.")
        if not PYTESSERACT_AVAILABLE:
            # Running from source: try pip install once (frozen EXE bundles it).
            pip = _pip_cmd()
            installed = False
            if pip:
                _olog("👁 Installing pytesseract package…")
                ok, _out = _run_cmd_quiet(pip + ["install", "--quiet", "pytesseract"], timeout=600)
                installed = ok
                if installed:
                    try:
                        import importlib as _imp
                        pytesseract = _imp.import_module("pytesseract")
                        PYTESSERACT_AVAILABLE = True
                        _refresh_tesseract_path()
                    except Exception:
                        installed = False
            if not installed:
                _olog("⚠ pytesseract Python package not available — OCR monitor disabled. "
                      "Run: pip install pytesseract")
                self._wa_ocr_running = False
                return
        self._whatsapp_ocr_loop()

    def _whatsapp_ocr_loop(self) -> None:
        """Continuously grab the screen area where WhatsApp notifications appear and read IMEIs.
        Dependency checks/auto-install happen in _whatsapp_ocr_entry."""
        import re as _re
        import datetime as _dt
        if not (PYTESSERACT_AVAILABLE and pytesseract is not None):
            self._wa_ocr_running = False
            return

        seen_imeis: set = set()
        while getattr(self, "_wa_ocr_running", False):
            try:
                # Grab notification area at top of screen (Windows toast region)
                notif_img = ImageGrab.grab(bbox=(0, 0, 600, 200))
                text = pytesseract.image_to_string(notif_img)
                # Extract 15-digit IMEI patterns
                imei_pattern = _re.compile(r'\b\d{15}\b')
                found = imei_pattern.findall(text)
                for imei in found:
                    if imei in seen_imeis:
                        continue
                    seen_imeis.add(imei)
                    self._ui_post(lambda i=imei: self._ocr_clear_imei(i))
            except Exception:
                pass
            time.sleep(5)

    def _ocr_clear_imei(self, imei: str) -> None:
        """Mark a variance row with the given IMEI as cleared (found via OCR)."""
        try:
            if not self.data_loaded:
                return
            rows = self.db.get_rows_by_keys(self.loaded_keys)
            matches = [row for row in rows if row.imei and row.imei.strip() == imei and not row.cleared]
            if matches:
                def _do():
                    for row in matches:
                        self.db.set_cleared(row.key, True)
                self._db_write(_do)
                self.refresh_table()
                self._log_scheduler(f"👁 OCR: auto-cleared IMEI {imei} ({len(matches)} row(s))")
        except Exception as exc:
            self._log_scheduler(f"⚠ OCR clear error for IMEI {imei}: {exc}")

    # ── Auto-import stores from inventory count ──────────────────────────────
    def _auto_import_stores_from_inventory(self) -> None:
        """Scan the loaded inventory XLSX and auto-populate the Store List from it."""
        inv_path_str = self.inventory_path.strip()
        if not inv_path_str:
            self.set_status("Auto-import: no inventory file loaded yet.")
            return
        inv_path = Path(inv_path_str)
        if not inv_path.exists():
            self.set_status(f"Auto-import: file not found — {inv_path.name}")
            return
        try:
            wb = openpyxl.load_workbook(inv_path, read_only=True, data_only=True)
            ws = wb.active
            headers = []
            store_col = district_col = None
            for row in ws.iter_rows(max_row=3, values_only=True):
                if any(str(c or "").strip() for c in row):
                    headers = [str(c or "").strip().lower() for c in row]
                    break
            for i, h in enumerate(headers):
                if h == "store":
                    store_col = i
                if h in ("district", "region", "market"):
                    district_col = i

            if store_col is None:
                self.set_status("Auto-import: 'Store' column not found in inventory file.")
                return

            # Collect store-district pairs
            DISTRICT_MAP = {
                "az": "Arizona", "co": "Colorado", "la": "Louisiana",
                "tn": "Tennessee", "tx": "Texas",
            }
            AZ_VARIANTS = {"arizona - d1", "arizona - d2", "arizona d1", "arizona d2",
                           "arizona", "az"}

            seen: set = set()
            count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                store = str(row[store_col] if row[store_col] is not None else "").strip()
                if not store:
                    continue
                district = ""
                if district_col is not None:
                    district = str(row[district_col] if row[district_col] is not None else "").strip()
                # Normalize district
                d_lower = district.lower()
                if d_lower in AZ_VARIANTS:
                    district = "Arizona"
                else:
                    for abbr, full in DISTRICT_MAP.items():
                        if d_lower == abbr or d_lower == full.lower():
                            district = full
                            break
                key = (district.lower(), store.lower())
                if key in seen:
                    continue
                seen.add(key)
                try:
                    self.db.save_store_account(
                        district=district, store=store,
                        account_id="", username="", password="")
                    count += 1
                except Exception:
                    pass

            wb.close()
            self.refresh_store_accounts_table()
            self._build_scheduler_district_rows()
            self.set_status(f"Auto-import: added/updated {count} store(s) from inventory file.")
            self._log_scheduler(f"Auto-imported {count} stores from {inv_path.name}.")
        except Exception as exc:
            self.set_status(f"Auto-import error: {exc}")

    def _build_status_tab(self) -> None:
        tab_outer = QVBoxLayout(self.status_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        # ── Audit Scheduler panel (inline on main tab) ──────────────────────
        sched_box = QGroupBox("Audit Scheduler")
        sched_outer = QVBoxLayout(sched_box)

        # Control buttons row
        btn_row = QWidget()
        srow = QHBoxLayout(btn_row)
        srow.setContentsMargins(0, 0, 0, 0)
        self._sched_start_btn = QPushButton("▶  Start")
        self._sched_start_btn.clicked.connect(self._sched_start)
        self._sched_stop_btn = QPushButton("■  Stop")
        self._sched_stop_btn.clicked.connect(self._sched_stop)
        self._sched_hold_btn = QPushButton("⏸  Hold")
        self._sched_hold_btn.clicked.connect(self._sched_hold)
        self._sched_resume_btn = QPushButton("⏵  Resume")
        self._sched_resume_btn.clicked.connect(self._sched_resume)
        for btn in (self._sched_start_btn, self._sched_stop_btn,
                    self._sched_hold_btn, self._sched_resume_btn):
            srow.addWidget(btn)
            srow.addSpacing(8)
        save_times_btn = QPushButton("💾  Save Times")
        save_times_btn.clicked.connect(self._sched_save_times)
        srow.addWidget(save_times_btn)
        srow.addSpacing(8)
        sched_hint = QLabel("  Start times per district (HH:MM AM/PM):")
        sched_hint.setStyleSheet("color:#8090b0;")
        srow.addWidget(sched_hint)
        srow.addStretch(1)
        sched_outer.addWidget(btn_row)

        # Duration and global stop time row
        dur_row = QWidget()
        drow = QHBoxLayout(dur_row)
        drow.setContentsMargins(0, 0, 0, 0)
        drow.addWidget(QLabel("Duration (hrs):"))
        self._sched_duration_edit = QLineEdit("12")
        self._sched_duration_edit.setFixedWidth(60)
        drow.addWidget(self._sched_duration_edit)
        drow.addSpacing(12)
        drow.addWidget(QLabel("Global stop time:"))
        self._sched_stop_time_edit = QLineEdit("")
        self._sched_stop_time_edit.setFixedWidth(70)
        drow.addWidget(self._sched_stop_time_edit)
        self._sched_stop_ampm_combo = QComboBox()
        self._sched_stop_ampm_combo.addItems(["AM", "PM"])
        self._combo_set(self._sched_stop_ampm_combo, "PM")
        self._sched_stop_ampm_combo.setFixedWidth(64)
        drow.addWidget(self._sched_stop_ampm_combo)
        drow.addSpacing(4)
        dur_hint = QLabel("HH:MM AM/PM — overrides duration when set")
        dur_hint.setStyleSheet("color:#8090b0;")
        drow.addWidget(dur_hint)
        drow.addStretch(1)
        sched_outer.addWidget(dur_row)

        # District time inputs (compact, inline)
        self._sched_frame = QWidget()
        self._sched_grid = QGridLayout(self._sched_frame)
        self._sched_grid.setContentsMargins(0, 0, 0, 0)
        self._sched_grid.setHorizontalSpacing(4)
        sched_outer.addWidget(self._sched_frame)

        # Tiny status log
        self._sched_log_text = QPlainTextEdit()
        self._sched_log_text.setReadOnly(True)
        self._sched_log_text.setMaximumHeight(70)
        sched_outer.addWidget(self._sched_log_text)
        tab_outer.addWidget(sched_box)

        # ── Send Inventory Audit Status controls ────────────────────────────
        controls = QGroupBox("Send Inventory Audit Status")
        crow = QHBoxLayout(controls)
        crow.addWidget(QLabel("Send by:"))
        self.status_send_mode_combo = QComboBox()
        self.status_send_mode_combo.addItems(["District", "Store", "Sales Rep"])
        crow.addWidget(self.status_send_mode_combo)
        crow.addSpacing(12)
        crow.addWidget(QLabel("District filter:"))
        self.status_district_combo = QComboBox()
        self.status_district_combo.addItem("All Districts")
        # activated = user selection only (same trigger as <<ComboboxSelected>>)
        self.status_district_combo.activated.connect(lambda _i: self.on_status_district_change())
        crow.addWidget(self.status_district_combo)
        crow.addSpacing(12)
        crow.addWidget(QLabel("Store filter:"))
        self.status_store_combo = QComboBox()
        self.status_store_combo.addItem("All Stores")
        self.status_store_combo.activated.connect(lambda _i: self.on_status_store_change())
        crow.addWidget(self.status_store_combo)
        crow.addSpacing(6)
        for text, fn in (("Check Current Filter", self.auto_check_status_rows),
                         ("Check Pending Only", self.auto_check_pending_status_rows),
                         ("Clear Checkmarks", self.clear_status_checkmarks),
                         ("Send Status Image", self.send_checked_status),
                         ("Add Store", self.add_store_prompt),
                         ("Open Folder", self.open_app_folder)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            crow.addWidget(btn)
        crow.addStretch(1)
        tab_outer.addWidget(controls)

        search_box = QGroupBox("Search Inventory Audit Status")
        srow2 = QHBoxLayout(search_box)
        for label, attr, width in (("Any", "status_search_any", 100),
                                   ("District", "status_search_district", 85),
                                   ("Store", "status_search_store", 100),
                                   ("Status", "status_search_status", 85),
                                   ("Rep", "status_search_rep", 100)):
            srow2.addWidget(QLabel(label + ":"))
            edit = QLineEdit()
            edit.setFixedWidth(width)
            edit.textChanged.connect(lambda _t: self.refresh_status_table())
            setattr(self, attr + "_edit", edit)
            srow2.addWidget(edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_status_search)
        srow2.addWidget(clear_btn)
        srow2.addStretch(1)
        tab_outer.addWidget(search_box)

        legend = QLabel("Colors: Yellow = Pending, Blue = Completed after updated sheet load, Green = Completed and sent to WhatsApp.")
        tab_outer.addWidget(legend)

        self.status_tree = self._make_table(
            ["District", "Store", "Status", "Rep Name", "Checkbox"],
            [130, 200, 110, 170, 95])
        self._wire_sorting(self.status_tree)
        self.status_tree.itemClicked.connect(self.on_status_tree_click)
        tab_outer.addWidget(self.status_tree, 1)

    def _build_audit_tab(self) -> None:
        tab_outer = QVBoxLayout(self.audit_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        send_box = QGroupBox("Variance Audit Controls")
        grid = QGridLayout(send_box)
        # Row 0: filters
        grid.addWidget(QLabel("Send by:"), 0, 0)
        self.audit_send_mode_combo = QComboBox()
        self.audit_send_mode_combo.addItems(["District", "Store", "Sales Rep"])
        grid.addWidget(self.audit_send_mode_combo, 0, 1)
        grid.addWidget(QLabel("District filter:"), 0, 2)
        self.audit_district_combo = QComboBox()
        self.audit_district_combo.addItem("All Districts")
        self.audit_district_combo.activated.connect(lambda _i: self.on_audit_district_change())
        grid.addWidget(self.audit_district_combo, 0, 3)
        grid.addWidget(QLabel("Store filter:"), 0, 4)
        self.audit_store_combo = QComboBox()
        self.audit_store_combo.addItem("All Stores")
        self.audit_store_combo.activated.connect(lambda _i: self.on_audit_store_change())
        grid.addWidget(self.audit_store_combo, 0, 5)
        self.send_only_unsent_check = QCheckBox("Only unsent")
        self.send_only_unsent_check.setChecked(bool(SEND_ONLY_UNSENT_BY_DEFAULT))
        grid.addWidget(self.send_only_unsent_check, 0, 6)
        self.include_cleared_check = QCheckBox("Show cleared")
        self.include_cleared_check.toggled.connect(lambda _on: self.refresh_table())
        grid.addWidget(self.include_cleared_check, 0, 7)
        grid.setColumnStretch(8, 1)
        # Row 1: actions
        for col, (text, fn) in enumerate((
                ("Check Current Filter", self.auto_check_audit_rows),
                ("Clear Checkmarks", self.clear_audit_checkmarks),
                ("Send Checked Variance Image", self.send_checked_variances),
                ("Send Selected Image", self.send_selected),
                ("Send Pending", self.send_pending),
                ("Mark Cleared", lambda: self.mark_selected(True)),
                ("Mark Not Cleared", lambda: self.mark_selected(False)),
                ("Export Log", self.export_log),
                ("Open Folder", self.open_app_folder),
                ("Clear UI", self.clear_current_ui),
                ("Copy IMEI", self.copy_selected_imei)), start=0):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            grid.addWidget(btn, 1, col)
        tab_outer.addWidget(send_box)

        action_box = QGroupBox("Final Send Actions")
        arow = QHBoxLayout(action_box)
        arow.addWidget(QLabel("District:"))
        self.final_district_combo = QComboBox()
        self.final_district_combo.addItem("All Districts")
        arow.addWidget(self.final_district_combo)
        arow.addSpacing(12)
        start_btn = QPushButton("Send Starting Message")
        start_btn.clicked.connect(self.send_starting_message)
        arow.addWidget(start_btn)
        for n in (1, 2, 3):
            btn = QPushButton(f"Send Reminder {n}")
            btn.clicked.connect(lambda _c=False, num=n: self.send_reminder(num))
            arow.addWidget(btn)
        final_btn = QPushButton("Send Final District Result")
        final_btn.clicked.connect(self.send_final_district_result)
        arow.addWidget(final_btn)
        arow.addStretch(1)
        tab_outer.addWidget(action_box)

        search_box = QGroupBox("Search Variance Rows")
        srow2 = QHBoxLayout(search_box)
        for label, attr, width in (("Any", "audit_search_any", 100),
                                   ("District", "audit_search_district", 85),
                                   ("Store", "audit_search_store", 100),
                                   ("Product", "audit_search_product", 115),
                                   ("IMEI", "audit_search_imei", 100),
                                   ("Rep", "audit_search_rep", 85)):
            srow2.addWidget(QLabel(label + ":"))
            edit = QLineEdit()
            edit.setFixedWidth(width)
            edit.textChanged.connect(lambda _t: self.refresh_table())
            setattr(self, attr + "_edit", edit)
            srow2.addWidget(edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_audit_search)
        srow2.addWidget(clear_btn)
        srow2.addStretch(1)
        tab_outer.addWidget(search_box)

        legend = QLabel("Colors: Yellow = Pending variance, Blue = Sent to WhatsApp, Green = Cleared.")
        tab_outer.addWidget(legend)

        self.audit_tree = self._make_table(
            ["District", "Store", "Product", "IMEI", "Status", "Rep Name", "Clearance", "Checkbox"],
            [110, 150, 280, 130, 100, 170, 110, 85], multi_select=True)
        self._wire_sorting(self.audit_tree)
        self.audit_tree.itemClicked.connect(self.on_audit_tree_click)
        self.audit_tree.itemDoubleClicked.connect(lambda _item: self.on_audit_double_click())
        self.audit_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.audit_tree.customContextMenuRequested.connect(lambda _pos: self.copy_selected_imei())
        sc_copy = QShortcut(QKeySequence("Ctrl+C"), self.audit_tree)
        sc_copy.activated.connect(self.copy_selected_imei)
        tab_outer.addWidget(self.audit_tree, 1)

    def _build_store_tab(self) -> None:
        tab_outer = QVBoxLayout(self.store_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        form = QGroupBox("Add or Update Store")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("District"), 0, 0)
        self._store_district_combo = QComboBox()
        self._store_district_combo.addItems(self.db.all_known_districts())
        grid.addWidget(self._store_district_combo, 0, 1)
        grid.addWidget(QLabel("Store"), 0, 2)
        self.store_name_edit = QLineEdit()
        self.store_name_edit.setFixedWidth(230)
        grid.addWidget(self.store_name_edit, 0, 3)
        for col, (text, fn) in enumerate((
                ("Save Store", self.save_store_account_from_form),
                ("Import XLSX", self.import_store_accounts_file),
                ("Delete Selected", self.delete_selected_store_account),
                ("Clear Form", self.clear_store_form)), start=4):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            grid.addWidget(btn, 0, col)
        tab_outer.addWidget(form)

        store_search_box = QWidget()
        srow = QHBoxLayout(store_search_box)
        srow.setContentsMargins(0, 0, 0, 0)
        srow.addWidget(QLabel("Search:"))
        self.store_search_edit = QLineEdit()
        self.store_search_edit.setFixedWidth(290)
        self.store_search_edit.textChanged.connect(lambda _t: self.refresh_store_accounts_table())
        srow.addWidget(self.store_search_edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.store_search_edit.clear(), self.refresh_store_accounts_table()))
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(store_search_box)

        self.store_tree = self._make_table(["District", "Store"], [220, 360])
        self._wire_sorting(self.store_tree)
        self.store_tree.itemSelectionChanged.connect(self.on_store_account_select)
        tab_outer.addWidget(self.store_tree, 1)

    def _build_rep_tab(self) -> None:
        tab_outer = QVBoxLayout(self.rep_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        info = QGroupBox("About Employees")
        info_l = QVBoxLayout(info)
        info_lbl = QLabel(
            "Store each employee's full name and phone number here so WhatsApp reports can tag them.\n"
            "The 'Created By' field is optional — fill it in when the count file uses a short username "
            "(e.g. 'ArmanAli') that differs from the full Employee Name. "
            "The audit tries Created By first, then Employee Name, when resolving who to tag.\n"
            "Use 'Auto-detect Created By' to scan loaded variances and pre-fill the column automatically.")
        info_lbl.setWordWrap(True)
        info_l.addWidget(info_lbl)
        tab_outer.addWidget(info)

        form = QGroupBox("Add or Update Employee")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("Employee Name:"), 0, 0)
        self.rep_name_edit = QLineEdit()
        self.rep_name_edit.setFixedWidth(190)
        grid.addWidget(self.rep_name_edit, 0, 1)
        grid.addWidget(QLabel("Phone:"), 0, 2)
        self.rep_phone_edit = QLineEdit()
        self.rep_phone_edit.setFixedWidth(150)
        grid.addWidget(self.rep_phone_edit, 0, 3)
        grid.addWidget(QLabel("Created By (optional):"), 0, 4)
        self.rep_created_by_edit = QLineEdit()
        self.rep_created_by_edit.setFixedWidth(150)
        grid.addWidget(self.rep_created_by_edit, 0, 5)
        grid.setColumnStretch(6, 1)
        btn_frame = QWidget()
        brow = QHBoxLayout(btn_frame)
        brow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Save Employee", self.save_sales_rep_from_form),
                         ("Auto-detect Created By", self._rep_auto_detect_created_by),
                         ("Import XLSX", self.import_employees_file),
                         ("Export Excel", self.export_employees_file),
                         ("Delete Selected", self.delete_selected_sales_rep),
                         ("Clear Form", self.clear_rep_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            brow.addWidget(btn)
        brow.addStretch(1)
        grid.addWidget(btn_frame, 1, 0, 1, 7)
        tab_outer.addWidget(form)

        rep_search_box = QWidget()
        srow = QHBoxLayout(rep_search_box)
        srow.setContentsMargins(0, 0, 0, 0)
        srow.addWidget(QLabel("Search:"))
        self.rep_search_edit = QLineEdit()
        self.rep_search_edit.setFixedWidth(290)
        self.rep_search_edit.textChanged.connect(lambda _t: self.refresh_sales_reps_table())
        srow.addWidget(self.rep_search_edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.rep_search_edit.clear(), self.refresh_sales_reps_table()))
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(rep_search_box)

        self.rep_tree = self._make_table(
            ["Employee Name", "Phone Number", "Created By (count username)"],
            [280, 180, 200])
        self._wire_sorting(self.rep_tree)
        self.rep_tree.itemSelectionChanged.connect(self.on_sales_rep_select)
        tab_outer.addWidget(self.rep_tree, 1)

    # ------------------------------------------------------------
    # TAB 5: District DMs (WhatsApp send mode, group names, DM tags)
    # ------------------------------------------------------------
    def _build_dm_tab(self) -> None:
        tab_outer = QVBoxLayout(self.dm_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        # WhatsApp send mode selector at top of tab
        mode_box = QGroupBox("WhatsApp Send Mode")
        mrow = QHBoxLayout(mode_box)
        mrow.addWidget(QLabel("Send reports via:"))
        self.wa_mode_var = self.db.get_setting("whatsapp_mode", "web")
        self.wa_desktop_radio = QRadioButton("WhatsApp Desktop App (pyautogui)")
        self.wa_web_radio = QRadioButton("WhatsApp Web (browser)")
        if self.wa_mode_var == "desktop":
            self.wa_desktop_radio.setChecked(True)
        elif self.wa_mode_var == "web":
            self.wa_web_radio.setChecked(True)
        self.wa_desktop_radio.toggled.connect(lambda on: on and setattr(self, "wa_mode_var", "desktop"))
        self.wa_web_radio.toggled.connect(lambda on: on and setattr(self, "wa_mode_var", "web"))
        mrow.addWidget(self.wa_desktop_radio)
        mrow.addSpacing(12)
        mrow.addWidget(self.wa_web_radio)
        mrow.addSpacing(12)
        save_mode_btn = QPushButton("Save Mode")
        save_mode_btn.clicked.connect(lambda: (self.db.save_setting("whatsapp_mode", self.wa_mode_var),
                                               self.set_status(f"WhatsApp mode saved: {self.wa_mode_var}")))
        mrow.addWidget(save_mode_btn)
        mrow.addStretch(1)
        tab_outer.addWidget(mode_box)

        # Side-by-side layout (Tk PanedWindow -> QSplitter)
        splitter = QSplitter(Qt.Horizontal)
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(4, 4, 4, 4)
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(4, 4, 4, 4)
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        tab_outer.addWidget(splitter, 1)

        # ── Left: WhatsApp Group Names ──────────────────────────────────
        wg_form = QGroupBox("WhatsApp Group Names")
        wg_outer = QVBoxLayout(wg_form)
        wg_form_row = QWidget()
        wrow = QHBoxLayout(wg_form_row)
        wrow.setContentsMargins(0, 0, 0, 0)
        wrow.addWidget(QLabel("District:"))
        self._wg_district_combo = QComboBox()
        self._wg_district_combo.addItems(self.db.all_known_districts())
        wrow.addWidget(self._wg_district_combo)
        wrow.addSpacing(8)
        wrow.addWidget(QLabel("Group Name:"))
        self.wg_group_name_edit = QLineEdit()
        self.wg_group_name_edit.setFixedWidth(210)
        wrow.addWidget(self.wg_group_name_edit)
        wrow.addStretch(1)
        wg_outer.addWidget(wg_form_row)
        wg_btn_row = QWidget()
        wbrow = QHBoxLayout(wg_btn_row)
        wbrow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Save", self.save_whatsapp_group_from_form),
                         ("Delete", self.delete_selected_whatsapp_group),
                         ("Clear", self.clear_wg_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            wbrow.addWidget(btn)
        wbrow.addStretch(1)
        wg_outer.addWidget(wg_btn_row)
        lv.addWidget(wg_form)

        self.wg_tree = self._make_table(["District", "WhatsApp Group Name"], [160, 280])
        self._wire_sorting(self.wg_tree)
        self.wg_tree.itemSelectionChanged.connect(self.on_whatsapp_group_select)
        lv.addWidget(self.wg_tree, 1)

        # ── Right: District Manager WhatsApp Tags ───────────────────────
        form = QGroupBox("DM WhatsApp Tags")
        dm_outer = QVBoxLayout(form)
        dm_form_row = QWidget()
        drow = QHBoxLayout(dm_form_row)
        drow.setContentsMargins(0, 0, 0, 0)
        drow.addWidget(QLabel("District:"))
        self._dm_district_combo = QComboBox()
        self._dm_district_combo.addItems(self.db.all_known_districts())
        drow.addWidget(self._dm_district_combo)
        drow.addSpacing(8)
        drow.addWidget(QLabel("DM Name:"))
        self.dm_name_edit = QLineEdit()
        self.dm_name_edit.setFixedWidth(180)
        drow.addWidget(self.dm_name_edit)
        drow.addSpacing(8)
        drow.addWidget(QLabel("Phone:"))
        self.dm_phone_edit = QLineEdit()
        self.dm_phone_edit.setFixedWidth(150)
        drow.addWidget(self.dm_phone_edit)
        drow.addStretch(1)
        dm_outer.addWidget(dm_form_row)
        dm_btn_row = QWidget()
        dbrow = QHBoxLayout(dm_btn_row)
        dbrow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Save", self.save_district_manager_from_form),
                         ("Delete", self.delete_selected_district_manager),
                         ("Clear", self.clear_dm_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            dbrow.addWidget(btn)
        dbrow.addStretch(1)
        dm_outer.addWidget(dm_btn_row)
        rv.addWidget(form)

        self.dm_tree = self._make_table(["District", "DM Name", "Phone Number"], [160, 220, 180])
        self._wire_sorting(self.dm_tree)
        self.dm_tree.itemSelectionChanged.connect(self.on_district_manager_select)
        rv.addWidget(self.dm_tree, 1)

    def refresh_district_managers_table(self) -> None:
        if not hasattr(self, "dm_tree"):
            return
        self.dm_tree.setRowCount(0)
        for row in self.db.district_managers():
            r = self.dm_tree.rowCount()
            self.dm_tree.insertRow(r)
            values = (row.get("District", ""), row.get("DM Name", ""), row.get("Phone", ""))
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                self.dm_tree.setItem(r, col, item)
            self.dm_tree.item(r, 0).setData(Qt.UserRole, normalize_district(row.get("District", "")))

    def save_district_manager_from_form(self) -> None:
        try:
            self.db.save_district_manager(self._dm_district_combo.currentText(), self.dm_name_edit.text(), self.dm_phone_edit.text())
            self.refresh_district_managers_table()
            self.refresh_entry_form_districts()
            self.set_status(f"Saved district DM: {self._dm_district_combo.currentText()} | {self.dm_phone_edit.text()}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_district_manager(self) -> None:
        rows = {i.row() for i in self.dm_tree.selectedIndexes()} if hasattr(self, "dm_tree") else set()
        if not rows:
            QMessageBox.information(self, "No selection", "Select a district DM first.")
            return
        if QMessageBox.question(self, "Delete district DM", "Delete the selected district DM phone record?") != QMessageBox.Yes:
            return
        district = self.dm_tree.item(sorted(rows)[0], 0).data(Qt.UserRole)
        self.db.delete_district_manager(district)
        self.refresh_district_managers_table()
        self.refresh_entry_form_districts()
        self.clear_dm_form()
        self.set_status("Deleted selected district DM.")

    def clear_dm_form(self) -> None:
        self._combo_set(self._dm_district_combo, "")
        self.dm_name_edit.clear()
        self.dm_phone_edit.clear()

    def refresh_whatsapp_groups_table(self) -> None:
        if not hasattr(self, "wg_tree"):
            return
        self.wg_tree.setRowCount(0)
        for row in self.db.whatsapp_groups():
            r = self.wg_tree.rowCount()
            self.wg_tree.insertRow(r)
            values = (row.get("District", ""), row.get("Group Name", ""))
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                self.wg_tree.setItem(r, col, item)
            self.wg_tree.item(r, 0).setData(Qt.UserRole, normalize_district(row.get("District", "")))

    def save_whatsapp_group_from_form(self) -> None:
        try:
            self.db.save_whatsapp_group(self._wg_district_combo.currentText(), self.wg_group_name_edit.text())
            self.refresh_whatsapp_groups_table()
            self.refresh_entry_form_districts()
            self.set_status(f"Saved WhatsApp group: {self._wg_district_combo.currentText()} | {self.wg_group_name_edit.text()}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_whatsapp_group(self) -> None:
        rows = {i.row() for i in self.wg_tree.selectedIndexes()} if hasattr(self, "wg_tree") else set()
        if not rows:
            QMessageBox.information(self, "No selection", "Select a WhatsApp group first.")
            return
        if QMessageBox.question(self, "Delete WhatsApp group", "Delete the selected WhatsApp group record?") != QMessageBox.Yes:
            return
        district = self.wg_tree.item(sorted(rows)[0], 0).data(Qt.UserRole)
        self.db.delete_whatsapp_group(district)
        self.refresh_whatsapp_groups_table()
        self.refresh_entry_form_districts()
        self.clear_wg_form()
        self.set_status("Deleted selected WhatsApp group.")

    def clear_wg_form(self) -> None:
        self._combo_set(self._wg_district_combo, "")
        self.wg_group_name_edit.clear()

    def on_whatsapp_group_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.wg_tree.selectedIndexes()})
        if not rows:
            return
        row = rows[0]
        district = self.wg_tree.item(row, 0).text()
        for entry in self.db.whatsapp_groups():
            if normalize_district(entry.get("District", "")) == normalize_district(district):
                self._combo_set(self._wg_district_combo, entry.get("District", ""))
                self.wg_group_name_edit.setText(entry.get("Group Name", ""))
                return

    def on_district_manager_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.dm_tree.selectedIndexes()})
        if not rows:
            return
        row = rows[0]
        district = self.dm_tree.item(row, 0).text()
        for entry in self.db.district_managers():
            if normalize_district(entry.get("District", "")) == normalize_district(district):
                self._combo_set(self._dm_district_combo, entry.get("District", ""))
                self.dm_name_edit.setText(entry.get("DM Name", ""))
                self.dm_phone_edit.setText(entry.get("Phone", ""))
                return

    # ── Created By → Employee Mapping Tab ─────────────────────────────────
    # NOTE: exactly like the Tkinter build, this panel is implemented but
    # not wired into the notebook (self.mapping_tab is never created there
    # either). Ported verbatim-in-spirit for parity.
    def _build_mapping_tab(self) -> None:
        """Panel to map 'Created By' values from inventory count to
        Employee names + phone numbers from the new Timesheet file.

        This lets the audit correctly identify and WhatsApp the right
        person when the count's Created By username (e.g. 'ArmanAli')
        doesn't directly match the Timesheet Employee name
        (e.g. 'Arman Ali Mohammed').
        """
        # ── Instructions ───────────────────────────────────────────────
        info = QGroupBox("About This Panel")
        info_l = QVBoxLayout(info)
        info_lbl = QLabel(
            "The Inventory Count file uses a short 'Created By' username (e.g. 'ArmanAli'), "
            "while the Timesheet file uses the full 'Employee' name (e.g. 'Arman Ali Mohammed').\n"
            "Add a mapping here so the audit can look up the correct employee (and their phone) "
            "when sending WhatsApp reports.  Mappings are saved permanently in the database.")
        info_lbl.setWordWrap(True)
        info_l.addWidget(info_lbl)
        self.mapping_tab_layout = QVBoxLayout(self.mapping_tab)
        self.mapping_tab_layout.addWidget(info)

        # ── Auto-populate from loaded data ─────────────────────────────
        auto_frame = QWidget()
        arow = QHBoxLayout(auto_frame)
        arow.setContentsMargins(0, 0, 0, 0)
        auto_btn = QPushButton("Auto-detect Created By values from loaded variances")
        auto_btn.clicked.connect(self._mapping_auto_detect)
        arow.addWidget(auto_btn)
        refresh_btn = QPushButton("Refresh Table")
        refresh_btn.clicked.connect(self._mapping_refresh)
        arow.addWidget(refresh_btn)
        arow.addStretch(1)
        self.mapping_tab_layout.addWidget(auto_frame)

        # ── Add / Edit form ────────────────────────────────────────────
        form = QGroupBox("Add / Edit Mapping")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("Created By (from count file):"), 0, 0)
        self.map_created_by_edit = QLineEdit()
        self.map_created_by_edit.setFixedWidth(180)
        grid.addWidget(self.map_created_by_edit, 0, 1)
        grid.addWidget(QLabel("Employee Name (from timesheet):"), 0, 2)
        self.map_employee_edit = QLineEdit()
        self.map_employee_edit.setFixedWidth(230)
        grid.addWidget(self.map_employee_edit, 0, 3)
        grid.addWidget(QLabel("Phone (WhatsApp):"), 0, 4)
        self.map_phone_edit = QLineEdit()
        self.map_phone_edit.setFixedWidth(150)
        grid.addWidget(self.map_phone_edit, 0, 5)
        btn_row = QWidget()
        brow = QHBoxLayout(btn_row)
        brow.setContentsMargins(0, 0, 0, 0)
        save_btn = QPushButton("Save Mapping")
        save_btn.clicked.connect(self._mapping_save)
        brow.addWidget(save_btn)
        del_btn = QPushButton("Delete Selected")
        del_btn.clicked.connect(self._mapping_delete)
        brow.addWidget(del_btn)
        clear_btn = QPushButton("Clear Form")
        clear_btn.clicked.connect(self._mapping_clear_form)
        brow.addWidget(clear_btn)
        brow.addStretch(1)
        grid.addWidget(btn_row, 1, 0, 1, 6)
        self.mapping_tab_layout.addWidget(form)

        # ── Table ──────────────────────────────────────────────────────
        tbl_frame = QGroupBox("Saved Mappings")
        tf = QVBoxLayout(tbl_frame)
        self.mapping_tree = self._make_table(
            ["Created By (count file)", "Employee Name (timesheet)", "Phone (WhatsApp)"],
            [200, 280, 160])
        self.mapping_tree.itemSelectionChanged.connect(self._mapping_on_select)
        tf.addWidget(self.mapping_tree)
        self.mapping_tab_layout.addWidget(tbl_frame, 1)
        self._mapping_refresh()

    def _mapping_refresh(self) -> None:
        self.mapping_tree.setRowCount(0)
        for m in self.db.get_created_by_mappings():
            r = self.mapping_tree.rowCount()
            self.mapping_tree.insertRow(r)
            for col, value in enumerate((m["created_by"], m["employee_name"], m["phone"])):
                self.mapping_tree.setItem(r, col, QTableWidgetItem(str(value)))

    def _mapping_on_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.mapping_tree.selectedIndexes()})
        if not rows:
            return
        row = rows[0]
        self.map_created_by_edit.setText(self.mapping_tree.item(row, 0).text())
        self.map_employee_edit.setText(self.mapping_tree.item(row, 1).text())
        phone_item = self.mapping_tree.item(row, 2)
        self.map_phone_edit.setText(phone_item.text() if phone_item is not None else "")

    def _mapping_save(self) -> None:
        cb = self.map_created_by_edit.text().strip()
        em = self.map_employee_edit.text().strip()
        ph = self.map_phone_edit.text().strip()
        if not cb:
            QMessageBox.warning(self, "Missing Value", "Enter the 'Created By' value from the count file.")
            return
        self.db.upsert_created_by_mapping(cb, em, ph)
        self._mapping_refresh()
        self._mapping_clear_form()
        self.set_status(f"Mapping saved: '{cb}' → '{em}'")

    def _mapping_delete(self) -> None:
        rows = sorted({i.row() for i in self.mapping_tree.selectedIndexes()})
        if not rows:
            QMessageBox.warning(self, "No Selection", "Select a row to delete.")
            return
        cb = self.mapping_tree.item(rows[0], 0).text()
        if QMessageBox.question(self, "Confirm Delete", f"Delete mapping for '{cb}'?") == QMessageBox.Yes:
            self.db.delete_created_by_mapping(cb)
            self._mapping_refresh()
            self.set_status(f"Deleted mapping for '{cb}'")

    def _mapping_clear_form(self) -> None:
        self.map_created_by_edit.clear()
        self.map_employee_edit.clear()
        self.map_phone_edit.clear()

    def _mapping_auto_detect(self) -> None:
        """Scan all loaded variance rows for unique Created By values
        and add any not already mapped (leaving employee/phone blank
        for the user to fill in)."""
        if not self.data_loaded:
            QMessageBox.warning(self, "No Data", "Load variances first, then auto-detect.")
            return
        existing = {m["created_by"] for m in self.db.get_created_by_mappings()}
        added = 0
        try:
            with self.db.connect() as con:
                rows = con.execute(
                    "SELECT DISTINCT created_by FROM variances WHERE created_by IS NOT NULL AND created_by != ''"
                ).fetchall()
            for (cb,) in rows:
                if cb and cb not in existing:
                    self.db.upsert_created_by_mapping(cb, "", "")
                    added += 1
            self._mapping_refresh()
            self.set_status(f"Auto-detected {added} new Created By value(s). Fill in Employee Name and Phone.")
            if added == 0:
                QMessageBox.information(self, "Auto-detect", "All Created By values already have mappings (or none found in loaded data).")
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    # ------------------------------------------------------------
    # TAB 6: Excluded Devices
    # ------------------------------------------------------------
    def _build_exclusion_tab(self) -> None:
        tab_outer = QVBoxLayout(self.exclusion_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        form = QGroupBox("Exclude Devices From Variance Images")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("District"), 0, 0)
        self._exclusion_district_combo = QComboBox()
        self._exclusion_district_combo.addItems(self.db.all_known_districts())
        grid.addWidget(self._exclusion_district_combo, 0, 1)
        grid.addWidget(QLabel("Product"), 0, 2)
        self.exclusion_product_edit = QLineEdit()
        self.exclusion_product_edit.setFixedWidth(240)
        grid.addWidget(self.exclusion_product_edit, 0, 3)
        grid.addWidget(QLabel("IMEI"), 0, 4)
        self.exclusion_imei_edit = QLineEdit()
        self.exclusion_imei_edit.setFixedWidth(180)
        self.exclusion_imei_edit.editingFinished.connect(self.lookup_product_for_exclusion_imei)
        grid.addWidget(self.exclusion_imei_edit, 0, 5)
        grid.addWidget(QLabel("Comments"), 0, 6)
        self.exclusion_comments_edit = QLineEdit()
        self.exclusion_comments_edit.setFixedWidth(280)
        grid.addWidget(self.exclusion_comments_edit, 0, 7)
        btn_row = QWidget()
        brow = QHBoxLayout(btn_row)
        brow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Find Product", self.lookup_product_for_exclusion_imei),
                         ("Save", self.save_device_exclusion_from_form),
                         ("Update Comment", self.update_selected_exclusion_comment),
                         ("Import XLSX", self.import_excluded_imeis_file),
                         ("Export Excel", self.export_excluded_imeis_file),
                         ("Delete", self.delete_selected_device_exclusion),
                         ("Delete All", self.delete_all_device_exclusions_from_ui),
                         ("Clear", self.clear_device_exclusion_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            brow.addWidget(btn)
        brow.addStretch(1)
        grid.addWidget(btn_row, 1, 0, 1, 8)
        tab_outer.addWidget(form)

        # Search bar for Excluded Devices
        exclusion_search_box = QWidget()
        srow = QHBoxLayout(exclusion_search_box)
        srow.setContentsMargins(0, 0, 0, 0)
        srow.addWidget(QLabel("Search:"))
        self.exclusion_search_edit = QLineEdit()
        self.exclusion_search_edit.setFixedWidth(290)
        self.exclusion_search_edit.textChanged.connect(lambda _t: self.refresh_device_exclusions_table())
        srow.addWidget(self.exclusion_search_edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.exclusion_search_edit.clear(), self.refresh_device_exclusions_table()))
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(exclusion_search_box)

        self.exclusion_tree = self._make_table(
            ["District", "Product", "IMEI", "Comments"], [160, 400, 220, 420])
        self._wire_sorting(self.exclusion_tree)
        self.exclusion_tree.itemSelectionChanged.connect(self.on_device_exclusion_select)
        self.exclusion_tree.itemDoubleClicked.connect(lambda _item: self.edit_selected_exclusion_comment())
        tab_outer.addWidget(self.exclusion_tree, 1)

    def delete_all_device_exclusions_from_ui(self) -> None:
        if QMessageBox.question(
                self, "Delete All Exclusions",
                "Are you sure you want to completely remove ALL device exclusions from the database? This cannot be undone.") != QMessageBox.Yes:
            return
        self.db.delete_all_device_exclusions()
        self.refresh_device_exclusions_table()
        self.refresh_entry_form_districts()
        self.refresh_table()
        self.clear_device_exclusion_form()
        self.set_status("Deleted all device exclusions.")

    def refresh_device_exclusions_table(self) -> None:
        if not hasattr(self, "exclusion_tree"):
            return
        self.exclusion_tree.setRowCount(0)
        search = self.exclusion_search_edit.text().strip().lower() if hasattr(self, "exclusion_search_edit") else ""
        for row in self.db.device_exclusions():
            district = row.get("District", "")
            product = row.get("Product", "")
            imei = row.get("IMEI", "")
            comments = row.get("Comments", "")
            if search and search not in district.lower() and search not in product.lower() and search not in imei.lower() and search not in comments.lower():
                continue
            r = self.exclusion_tree.rowCount()
            self.exclusion_tree.insertRow(r)
            values = (district, product, imei, comments)
            for col, value in enumerate(values):
                self.exclusion_tree.setItem(r, col, QTableWidgetItem(str(value)))
            iid = row.get("RuleKey", device_rule_key(district + "|" + product + "|" + imei))
            self.exclusion_tree.item(r, 0).setData(Qt.UserRole, iid)

    def product_for_imei_from_loaded_rows(self, imei: str) -> str:
        imei_key = device_rule_key(imei)
        if not imei_key:
            return ""

        # Prefer current loaded variance rows.
        for row in self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []:
            if device_rule_key(row.imei) == imei_key:
                return safe_text(row.product)

        # Fallback to currently read inventory records if available.
        for rec in self.current_inventory_records:
            for key, value in rec.items():
                if device_rule_key(value) == imei_key:
                    product_col = find_column(rec, ["Product", "Product Name", "Device", "Model", "Item", "SKU Description"])
                    if product_col:
                        return safe_text(rec.get(product_col, ""))
        return ""

    def lookup_product_for_exclusion_imei(self, _event=None):
        imei = self.exclusion_imei_edit.text()
        if not safe_text(imei):
            return None
        product = self.product_for_imei_from_loaded_rows(imei)
        if product:
            self.exclusion_product_edit.setText(product)
            self.set_status(f"Product found for IMEI {imei}: {product}")
        else:
            self.set_status(f"No product found for IMEI {imei}. Product field left unchanged.")
        return None

    def save_device_exclusion_from_form(self) -> None:
        try:
            district = self._exclusion_district_combo.currentText()
            imei = self.exclusion_imei_edit.text()
            product = self.exclusion_product_edit.text()
            comments = self.exclusion_comments_edit.text()

            if safe_text(imei) and not safe_text(product):
                found_product = self.product_for_imei_from_loaded_rows(imei)
                if found_product:
                    product = found_product
                    self.exclusion_product_edit.setText(found_product)

            # IMEI-ONLY exclusions (user rule 2026-09-15): the IMEI is the
            # exclusion — whole products cannot be excluded. A product-only
            # entry is rejected and the app asks for the IMEI(s) instead.
            if not safe_text(imei):
                raise ValueError(
                    "IMEI is required — whole products cannot be excluded. "
                    "Enter the IMEI(s) of the specific device(s) to exclude."
                )
            # IMEI exclusions stay global and ignore district.
            district = ""
            self._combo_set(self._exclusion_district_combo, "")

            selected_key = self.selected_exclusion_key.strip()
            if selected_key:
                self.db.delete_device_exclusion(selected_key)

            self.db.save_device_exclusion(product, imei, comments, district)
            self.selected_exclusion_key = ""
            self.refresh_device_exclusions_table()
            self.refresh_entry_form_districts()
            self.refresh_table()
            self.set_status(
                f"Saved exclusion. District: {safe_text(district) or '-'} | Product: {safe_text(product) or '-'} | IMEI: {safe_text(imei) or '-'} | Comments: {safe_text(comments) or '-'}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def update_selected_exclusion_comment(self) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select an exclusion row first.")
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        try:
            self.db.update_device_exclusion_comment(rule_key, self.exclusion_comments_edit.text())
            self.selected_exclusion_key = rule_key
            self.refresh_device_exclusions_table()
            self.refresh_table()
            self.set_status("Updated selected exclusion comment.")
        except Exception as exc:
            QMessageBox.critical(self, "Update failed", str(exc))

    def edit_selected_exclusion_comment(self, _event=None):
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            return None

        current_comment = self.exclusion_comments_edit.text()
        new_comment, ok = QInputDialog.getText(
            self,
            "Edit Comments",
            "Enter comments for selected exclusion.",
            text=current_comment,
        )
        if not ok:
            return None

        self.exclusion_comments_edit.setText(new_comment)
        self.update_selected_exclusion_comment()
        return None

    def import_excluded_imeis_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self,
            "Import Excluded Devices XLSX",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)",
        )
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()
            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No exclusion rows found in the selected file.")
                return

            sample = records[0]
            district_col = find_column(sample, ["District", "Market", "Area"])
            product_col = find_column(sample, ["Product", "Product Name", "Device", "Model", "Item"])
            imei_col = find_column(sample, ["IMEI", "IMEI Number", "ESN", "ESN Number", "Serial", "Serial 1", "Device ID"])
            comments_col = find_column(sample, ["Comments", "Comment", "Notes", "Note", "Reason"])

            if not imei_col:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Import file needs an IMEI column — exclusions are IMEI-level only. Supported headers: District, Product, IMEI, Comments.",
                )
                return

            imported = 0
            skipped = 0
            for rec in records:
                district = normalize_district(rec.get(district_col, "")) if district_col else ""
                product = safe_text(rec.get(product_col, "")) if product_col else ""
                imei = safe_text(rec.get(imei_col, "")) if imei_col else ""
                comments = safe_text(rec.get(comments_col, "")) if comments_col else ""

                if imei and not product:
                    product = self.product_for_imei_from_loaded_rows(imei)

                # IMEI-ONLY exclusions (user rule 2026-09-15): rows without
                # an IMEI cannot be imported — whole products cannot be
                # excluded. The app asks for the IMEI(s) instead.
                if not imei:
                    skipped += 1
                    continue

                # IMEI exclusions ignore district.
                district = ""

                self.db.save_device_exclusion(product, imei, comments, district)
                imported += 1

            self.refresh_device_exclusions_table()
            self.refresh_table()
            QMessageBox.information(self, "Import complete", f"Imported {imported} exclusion row(s).\nSkipped {skipped} invalid row(s).\n\nExclusions are IMEI-level only — whole products cannot be excluded. Provide the IMEI(s) of any device that must be hidden.")
            self.set_status(f"Imported excluded devices. Imported: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Excluded devices import failed.")

    def export_excluded_imeis_file(self) -> None:
        rows = self.db.device_exclusions()
        if not rows:
            QMessageBox.information(self, "No exclusions", "No excluded devices are saved.")
            return
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Excluded_Devices_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Excluded Devices", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["District", "Product", "IMEI", "Comments"])
        for r_idx, row in enumerate(rows, start=2):
            ws.cell(row=r_idx, column=1, value=row.get("District", ""))
            ws.cell(row=r_idx, column=2, value=row.get("Product", ""))

            imei_cell = ws.cell(row=r_idx, column=3, value=row.get("IMEI", ""))
            imei_cell.data_type = 's'
            imei_cell.number_format = '@'

            ws.cell(row=r_idx, column=4, value=row.get("Comments", ""))
        wb.save(output_path)

        self.set_status(f"Exported excluded devices: {output_path}")
        QMessageBox.information(self, "Export complete", f"Excluded devices exported:\n{output_path}")

    def delete_selected_device_exclusion(self) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a device exclusion first.")
            return
        if QMessageBox.question(self, "Delete exclusion", "Delete the selected device exclusion from the database?") != QMessageBox.Yes:
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        self.db.delete_device_exclusion(rule_key)
        self.refresh_device_exclusions_table()
        self.refresh_entry_form_districts()
        self.refresh_table()
        self.clear_device_exclusion_form()
        self.set_status("Deleted selected device exclusion.")

    def clear_device_exclusion_form(self) -> None:
        self._combo_set(self._exclusion_district_combo, "")
        self.exclusion_product_edit.clear()
        self.exclusion_imei_edit.clear()
        self.exclusion_comments_edit.clear()
        self.selected_exclusion_key = ""

    def on_device_exclusion_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        self.selected_exclusion_key = rule_key
        for row in self.db.device_exclusions():
            if row.get("RuleKey") == rule_key:
                self._combo_set(self._exclusion_district_combo, row.get("District", ""))
                self.exclusion_product_edit.setText(row.get("Product", ""))
                self.exclusion_imei_edit.setText(row.get("IMEI", ""))
                self.exclusion_comments_edit.setText(row.get("Comments", ""))
                return

    def _rep_auto_detect_created_by(self) -> None:
        """Match 'Created By' usernames from the count file to full Employee
        names from the timesheet by finding who worked at the same Store.

        Example: Count file has 'BabarAli' at 'Mt View Store'.
                 Timesheet has 'Babar Ali' at 'Mt View Store'.
                 → auto-maps 'BabarAli' → 'Babar Ali'.
        """
        inv_records = getattr(self, "current_inventory_records", [])
        ts_records = getattr(self, "current_time_sheet_records", [])

        if not inv_records:
            QMessageBox.information(self, "No Data", "Load an Inventory Count file first (Load Variances).")
            return

        # Build store → employee map from timesheet
        sample_ts = ts_records[0] if ts_records else {}
        store_col_ts = find_column(sample_ts, ["Store"])
        emp_col_ts = find_column(sample_ts, ["Employee", "Employee Name"])
        email_col_ts = find_column(sample_ts, ["Email", "Email Address"])

        store_to_employee: Dict[str, str] = {}  # norm_store → full employee name
        store_to_email: Dict[str, str] = {}

        for rec in ts_records:
            emp_raw = safe_text(rec.get(emp_col_ts, "")) if emp_col_ts else ""
            if not emp_raw or "TOTAL" in emp_raw.upper() or "— " in emp_raw:
                continue
            store_raw = rec.get(store_col_ts, "") if store_col_ts else ""
            norm = normalize_store(store_raw)
            if norm:
                store_to_employee[norm] = emp_raw.strip()
                if email_col_ts:
                    store_to_email[norm] = safe_text(rec.get(email_col_ts, ""))

        # Build store → created_by map from count file
        sample_inv = inv_records[0] if inv_records else {}
        store_col_inv = find_column(sample_inv, ["Store"])
        cb_col = find_column(sample_inv, ["Created By", "Count By", "User Login"])

        store_to_created_by: Dict[str, str] = {}
        for rec in inv_records:
            store_raw = rec.get(store_col_inv, "") if store_col_inv else ""
            norm = normalize_store(store_raw)
            cb = safe_text(rec.get(cb_col, "")).strip() if cb_col else ""
            if norm and cb:
                store_to_created_by[norm] = cb

        # Match: same store = same person
        existing = {m["created_by"]: m for m in self.db.get_created_by_mappings()}
        added = updated = 0

        for norm_store, cb in store_to_created_by.items():
            emp_name = store_to_employee.get(norm_store, "")
            email = store_to_email.get(norm_store, "")
            if cb in existing:
                # Update employee name if now known and not already set
                if emp_name and not existing[cb]["employee_name"]:
                    self.db.upsert_created_by_mapping(cb, emp_name, existing[cb]["phone"])
                    updated += 1
            else:
                self.db.upsert_created_by_mapping(cb, emp_name, "")
                added += 1

        self.refresh_sales_reps_table()

        if added + updated == 0 and not store_to_created_by:
            QMessageBox.information(self, "Auto-detect", "No 'Created By' values found in the count file.")
            return

        msg = []
        if added:
            msg.append(f"{added} new mapping(s) added")
        if updated:
            msg.append(f"{updated} mapping(s) updated with employee name")

        summary_lines = []
        for norm_store, cb in sorted(store_to_created_by.items()):
            emp = store_to_employee.get(norm_store, "— not found in timesheet")
            summary_lines.append(f"  {cb}  →  {emp}  ({display_store(norm_store)})")

        self.set_status(f"Auto-detect: {', '.join(msg) if msg else 'all already mapped'}.")
        QMessageBox.information(
            self,
            "Auto-detect Complete",
            (f"{', '.join(msg)}.\n\n" if msg else "All already mapped.\n\n") +
            "Matches found:\n" + "\n".join(summary_lines[:20]) +
            ("\n..." if len(summary_lines) > 20 else "") +
            "\n\nFill in Phone numbers in the Employees tab, then Save.",
        )

    def import_store_accounts_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Import Store List XLSX", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()

            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No store rows found in the selected file.")
                return

            sample = records[0]
            district_col = find_column(sample, ["District", "Distrcit"])
            store_col = find_column(sample, ["Store", "Store Name", "Location"])

            missing = []
            if not district_col:
                missing.append("District")
            if not store_col:
                missing.append("Store")

            if missing:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Missing required column(s): " + ", ".join(missing) +
                    "\n\nRequired headers: District, Store",
                )
                return

            parsed_records: List[Dict[str, str]] = []
            skipped = 0
            for rec in records:
                district = normalize_district(rec.get(district_col, ""))
                store = display_store(rec.get(store_col, ""))

                if not store or not district or district == "Unknown":
                    skipped += 1
                    continue

                parsed_records.append({"District": district, "Store": store})

            imported = self.db.replace_store_accounts(parsed_records)

            self.master_store_records = self.db.store_master_records()
            self.refresh_store_accounts_table()
            self.refresh_entry_form_districts()
            self.populate_status_filters()
            self.populate_audit_filters()

            if self.data_loaded:
                source_name = os.path.basename(self.inventory_path.strip()) if self.inventory_path.strip() else ""
                self.status_rows, status_summary = build_inventory_status_rows(
                    self.current_inventory_records,
                    self.current_time_sheet_records,
                    master_store_records=self.master_store_records,
                    source_file=source_name,
                )
                self.status_row_by_key = {r.key: r for r in self.status_rows}
                self.db.upsert_inventory_status_rows(self.status_rows)
                self.refresh_status_table()

            QMessageBox.information(
                self,
                "Import complete",
                f"Replaced saved store list with {imported} row(s).\nSkipped {skipped} invalid row(s).\n\nThis list will stay saved until the next upload.",
            )
            self.set_status(f"Store list replaced from file. Saved: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Store list import failed.")

    def refresh_store_accounts_table(self) -> None:
        if not hasattr(self, "store_tree"):
            return
        self.store_tree.setRowCount(0)
        search = self.store_search_edit.text().strip().lower() if hasattr(self, "store_search_edit") else ""
        for row in self.db.store_accounts():
            district = row.get("District", "")
            store = row.get("Store", "")
            if search and search not in district.lower() and search not in store.lower():
                continue
            iid = row.get("StoreKey", normalize_store(store))
            r = self.store_tree.rowCount()
            self.store_tree.insertRow(r)
            for col, value in enumerate((district, store)):
                self.store_tree.setItem(r, col, QTableWidgetItem(str(value)))
            self.store_tree.item(r, 0).setData(Qt.UserRole, iid)

    def refresh_entry_form_districts(self) -> None:
        """Refresh all entry-form district comboboxes from the DB."""
        districts = self.db.all_known_districts()
        for combo in (self._store_district_combo, self._wg_district_combo, self._dm_district_combo, self._exclusion_district_combo):
            try:
                current = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                combo.addItems(districts)
                self._combo_set(combo, current)
                combo.blockSignals(False)
            except Exception:
                pass
        # Also refresh the audit and status filter combos
        if hasattr(self, 'audit_district_combo'):
            self.populate_audit_filters()
        if hasattr(self, 'status_district_combo'):
            self.populate_status_filters()

    def import_employees_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Import Employees XLSX", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()

            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No employee rows found in the selected file.")
                return

            sample = records[0]
            name_col = find_column(sample, [
                "Employee Name",
                "Employee",
                "Sales Rep Name",
                "Sales Rep",
                "Rep Name",
                "Salesperson",
                "Name",
                "Full Name",
            ])
            phone_col = find_column(sample, [
                "Phone Number",
                "Phone",
                "Mobile",
                "Mobile Number",
                "Contact",
                "Contact Number",
                "Number",
                "Whatsapp",
                "WhatsApp Number",
            ])

            missing = []
            if not name_col:
                missing.append("Employee Name")
            if not phone_col:
                missing.append("Phone Number")

            if missing:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Missing required column(s): " + ", ".join(missing) +
                    "\n\nRequired headers: Employee Name, Phone Number",
                )
                return

            imported = 0
            skipped = 0
            for rec in records:
                name = safe_text(rec.get(name_col, ""))
                phone = normalize_phone(rec.get(phone_col, ""))

                if not name or not phone:
                    skipped += 1
                    continue

                self.db.save_sales_rep(name, phone)
                imported += 1

            self.refresh_sales_reps_table()
            QMessageBox.information(
                self,
                "Import complete",
                f"Imported or updated {imported} employee phone row(s).\nSkipped {skipped} invalid row(s).\n\nEmployees stay saved until you delete them.",
            )
            self.set_status(f"Imported employees from file. Imported: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Employee import failed.")

    def export_employees_file(self) -> None:
        rows = self.db.sales_reps()
        if not rows:
            QMessageBox.information(self, "No employees", "No employees are saved in the database yet.")
            return

        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Employees_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Employees Excel", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return

        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Employee Name", "Phone Number"])
        for r_idx, row in enumerate(rows, start=2):
            ws.cell(row=r_idx, column=1, value=row.get("Rep Name", ""))
            ws.cell(row=r_idx, column=2, value=row.get("Phone", ""))
        wb.save(output_path)

        self.set_status(f"Exported employee list: {output_path}")
        QMessageBox.information(self, "Export complete", f"Employee list exported:\n{output_path}")

    def refresh_sales_reps_table(self) -> None:
        if not hasattr(self, "rep_tree"):
            return
        self.rep_tree.setRowCount(0)
        search = self.rep_search_edit.text().strip().lower() if hasattr(self, "rep_search_edit") else ""
        # Show standard sales_reps rows
        for row in self.db.sales_reps():
            rep_name = row.get("Rep Name", "")
            phone = row.get("Phone", "")
            created_by = ""
            if search and search not in rep_name.lower() and search not in phone.lower():
                continue
            iid = row.get("RepKey", person_name_key(rep_name))
            r = self.rep_tree.rowCount()
            self.rep_tree.insertRow(r)
            for col, value in enumerate((rep_name, phone, created_by)):
                self.rep_tree.setItem(r, col, QTableWidgetItem(str(value)))
            self.rep_tree.item(r, 0).setData(Qt.UserRole, iid)
        # Also show created_by mappings that don't have a matching sales_rep entry
        existing_names = {row.get("Rep Name", "").lower() for row in self.db.sales_reps()}
        for m in self.db.get_created_by_mappings():
            if m["employee_name"].lower() in existing_names:
                # Already shown above — just update its created_by column
                for r in range(self.rep_tree.rowCount()):
                    name_item = self.rep_tree.item(r, 0)
                    if name_item is not None and name_item.text().lower() == m["employee_name"].lower():
                        self.rep_tree.setItem(r, 2, QTableWidgetItem(m["created_by"]))
                        break
                continue
            if search and search not in m["created_by"].lower() and search not in m["employee_name"].lower():
                continue
            r = self.rep_tree.rowCount()
            self.rep_tree.insertRow(r)
            for col, value in enumerate((m["employee_name"], m["phone"], m["created_by"])):
                self.rep_tree.setItem(r, col, QTableWidgetItem(str(value)))

    def save_store_account_from_form(self) -> None:
        try:
            self.db.save_store_account(
                self._store_district_combo.currentText(),
                self.store_name_edit.text(),
                "",
                "",
                "",
            )
            self.master_store_records = self.db.store_master_records()
            self.refresh_store_accounts_table()
            self.refresh_entry_form_districts()
            self.populate_status_filters()
            self.populate_audit_filters()
            if self.data_loaded:
                source_name = os.path.basename(self.inventory_path.strip()) if self.inventory_path.strip() else ""
                self.status_rows, status_summary = build_inventory_status_rows(
                    self.current_inventory_records,
                    self.current_time_sheet_records,
                    master_store_records=self.master_store_records,
                    source_file=source_name,
                )
                self.status_row_by_key = {r.key: r for r in self.status_rows}
                self.db.upsert_inventory_status_rows(self.status_rows)
                self.refresh_status_table()
            self.set_status(f"Saved store: {self._store_district_combo.currentText()} | {self.store_name_edit.text()}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_store_account(self) -> None:
        rows = sorted({i.row() for i in self.store_tree.selectedIndexes()}) if hasattr(self, "store_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a store first.")
            return
        if QMessageBox.question(self, "Delete store", "Delete the selected store?") != QMessageBox.Yes:
            return
        store_key = self.store_tree.item(rows[0], 0).data(Qt.UserRole)
        self.db.delete_store_account(store_key)
        self.master_store_records = self.db.store_master_records()
        self.refresh_store_accounts_table()
        self.refresh_entry_form_districts()
        self.clear_store_form()
        self.set_status("Deleted selected store.")

    def clear_store_form(self) -> None:
        self._combo_set(self._store_district_combo, "")
        self.store_name_edit.clear()

    def on_store_account_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.store_tree.selectedIndexes()}) if hasattr(self, "store_tree") else []
        if not rows:
            return
        store_key = self.store_tree.item(rows[0], 0).data(Qt.UserRole)
        for row in self.db.store_accounts():
            if row.get("StoreKey") == store_key:
                self._combo_set(self._store_district_combo, row.get("District", ""))
                self.store_name_edit.setText(row.get("Store", ""))
                return

    def save_sales_rep_from_form(self) -> None:
        try:
            name = self.rep_name_edit.text().strip()
            phone = self.rep_phone_edit.text().strip()
            created_by = self.rep_created_by_edit.text().strip() if hasattr(self, "rep_created_by_edit") else ""
            self.db.save_sales_rep(name, phone)
            # Also save the created_by mapping if provided
            if created_by:
                self.db.upsert_created_by_mapping(created_by, name, phone)
            self.refresh_sales_reps_table()
            self.set_status(f"Saved: {name} | {phone}" + (f" | Created By: {created_by}" if created_by else ""))
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_sales_rep(self) -> None:
        rows = sorted({i.row() for i in self.rep_tree.selectedIndexes()}) if hasattr(self, "rep_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select an employee first.")
            return
        if QMessageBox.question(self, "Delete employee", "Delete the selected employee record?") != QMessageBox.Yes:
            return
        rep_key = self.rep_tree.item(rows[0], 0).data(Qt.UserRole)
        created_by_item = self.rep_tree.item(rows[0], 2)
        vals_created_by = created_by_item.text() if created_by_item is not None else ""
        if rep_key:
            self.db.delete_sales_rep(rep_key)
        # Also remove the created_by mapping if present
        if vals_created_by:
            self.db.delete_created_by_mapping(vals_created_by)
        self.refresh_sales_reps_table()
        self.clear_rep_form()
        self.set_status("Deleted selected employee.")

    def clear_rep_form(self) -> None:
        self.rep_name_edit.clear()
        self.rep_phone_edit.clear()
        if hasattr(self, "rep_created_by_edit"):
            self.rep_created_by_edit.clear()

    def on_sales_rep_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.rep_tree.selectedIndexes()}) if hasattr(self, "rep_tree") else []
        if not rows:
            return
        row = rows[0]
        name_item = self.rep_tree.item(row, 0)
        phone_item = self.rep_tree.item(row, 1)
        created_by_item = self.rep_tree.item(row, 2)
        if name_item is not None:
            self.rep_name_edit.setText(name_item.text() if name_item is not None else "")
            self.rep_phone_edit.setText(phone_item.text() if phone_item is not None else "")
            if hasattr(self, "rep_created_by_edit"):
                self.rep_created_by_edit.setText(created_by_item.text() if created_by_item is not None else "")

    def product_for_imei_from_loaded_rows(self, imei: str) -> str:
        imei_key = device_rule_key(imei)
        if not imei_key:
            return ""

        # Prefer current loaded variance rows.
        for row in self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []:
            if device_rule_key(row.imei) == imei_key:
                return safe_text(row.product)

        # Fallback to currently read inventory records if available.
        for rec in self.current_inventory_records:
            for key, value in rec.items():
                if device_rule_key(value) == imei_key:
                    product_col = find_column(rec, ["Product", "Product Name", "Device", "Model", "Item", "SKU Description"])
                    if product_col:
                        return safe_text(rec.get(product_col, ""))
        return ""

    def lookup_product_for_exclusion_imei(self, _event=None):
        imei = self.exclusion_imei_edit.text()
        if not safe_text(imei):
            return None
        product = self.product_for_imei_from_loaded_rows(imei)
        if product:
            self.exclusion_product_edit.setText(product)
            self.set_status(f"Product found for IMEI {imei}: {product}")
        else:
            self.set_status(f"No product found for IMEI {imei}. Product field left unchanged.")
        return None

    def update_selected_exclusion_comment(self) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select an exclusion row first.")
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        try:
            self.db.update_device_exclusion_comment(rule_key, self.exclusion_comments_edit.text())
            self.selected_exclusion_key = rule_key
            self.refresh_device_exclusions_table()
            self.refresh_table()
            self.set_status("Updated selected exclusion comment.")
        except Exception as exc:
            QMessageBox.critical(self, "Update failed", str(exc))

    def edit_selected_exclusion_comment(self, _event=None):
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            return None

        current_comment = self.exclusion_comments_edit.text()
        new_comment, ok = QInputDialog.getText(
            self,
            "Edit Comments",
            "Enter comments for selected exclusion.",
            text=current_comment,
        )
        if not ok:
            return None

        self.exclusion_comments_edit.setText(new_comment)
        self.update_selected_exclusion_comment()
        return None

    def export_excluded_imeis_file(self) -> None:
        rows = self.db.device_exclusions()
        if not rows:
            QMessageBox.information(self, "No exclusions", "No excluded devices are saved.")
            return
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Excluded_Devices_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Excluded Devices", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["District", "Product", "IMEI", "Comments"])
        for r_idx, row in enumerate(rows, start=2):
            ws.cell(row=r_idx, column=1, value=row.get("District", ""))
            ws.cell(row=r_idx, column=2, value=row.get("Product", ""))

            imei_cell = ws.cell(row=r_idx, column=3, value=row.get("IMEI", ""))
            imei_cell.data_type = 's'
            imei_cell.number_format = '@'

            ws.cell(row=r_idx, column=4, value=row.get("Comments", ""))
        wb.save(output_path)

        self.set_status(f"Exported excluded devices: {output_path}")
        QMessageBox.information(self, "Export complete", f"Excluded devices exported:\n{output_path}")

    def delete_selected_device_exclusion(self) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a device exclusion first.")
            return
        if QMessageBox.question(self, "Delete exclusion", "Delete the selected device exclusion from the database?") != QMessageBox.Yes:
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        self.db.delete_device_exclusion(rule_key)
        self.refresh_device_exclusions_table()
        self.refresh_entry_form_districts()
        self.refresh_table()
        self.clear_device_exclusion_form()
        self.set_status("Deleted selected device exclusion.")

    def clear_device_exclusion_form(self) -> None:
        self._combo_set(self._exclusion_district_combo, "")
        self.exclusion_product_edit.clear()
        self.exclusion_imei_edit.clear()
        self.exclusion_comments_edit.clear()
        self.selected_exclusion_key = ""

    def on_device_exclusion_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        self.selected_exclusion_key = rule_key
        for row in self.db.device_exclusions():
            if row.get("RuleKey") == rule_key:
                self._combo_set(self._exclusion_district_combo, row.get("District", ""))
                self.exclusion_product_edit.setText(row.get("Product", ""))
                self.exclusion_imei_edit.setText(row.get("IMEI", ""))
                self.exclusion_comments_edit.setText(row.get("Comments", ""))
                return

    def save_device_exclusion_from_form(self) -> None:
        try:
            district = self._exclusion_district_combo.currentText()
            imei = self.exclusion_imei_edit.text()
            product = self.exclusion_product_edit.text()
            comments = self.exclusion_comments_edit.text()

            if safe_text(imei) and not safe_text(product):
                found_product = self.product_for_imei_from_loaded_rows(imei)
                if found_product:
                    product = found_product
                    self.exclusion_product_edit.setText(found_product)

            # IMEI exclusions stay global and ignore district.
            if safe_text(imei):
                district = ""
                self._combo_set(self._exclusion_district_combo, "")

            selected_key = self.selected_exclusion_key.strip()
            if selected_key:
                self.db.delete_device_exclusion(selected_key)

            self.db.save_device_exclusion(product, imei, comments, district)
            self.selected_exclusion_key = ""
            self.refresh_device_exclusions_table()
            self.refresh_entry_form_districts()
            self.refresh_table()
            self.set_status(
                f"Saved exclusion. District: {safe_text(district) or '-'} | Product: {safe_text(product) or '-'} | IMEI: {safe_text(imei) or '-'} | Comments: {safe_text(comments) or '-'}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def import_excluded_imeis_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self,
            "Import Excluded Devices XLSX",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)",
        )
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()
            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No exclusion rows found in the selected file.")
                return

            sample = records[0]
            district_col = find_column(sample, ["District", "Market", "Area"])
            product_col = find_column(sample, ["Product", "Product Name", "Device", "Model", "Item"])
            imei_col = find_column(sample, ["IMEI", "IMEI Number", "ESN", "ESN Number", "Serial", "Serial 1", "Device ID"])
            comments_col = find_column(sample, ["Comments", "Comment", "Notes", "Note", "Reason"])

            if not product_col and not imei_col:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Import file needs at least Product or IMEI. Supported headers: District, Product, IMEI, Comments.",
                )
                return

            imported = 0
            skipped = 0
            for rec in records:
                district = normalize_district(rec.get(district_col, "")) if district_col else ""
                product = safe_text(rec.get(product_col, "")) if product_col else ""
                imei = safe_text(rec.get(imei_col, "")) if imei_col else ""
                comments = safe_text(rec.get(comments_col, "")) if comments_col else ""

                if imei and not product:
                    product = self.product_for_imei_from_loaded_rows(imei)

                # IMEI exclusions ignore district. Product exclusions require district.
                if imei:
                    district = ""
                elif product and (not district or district == "Unknown"):
                    skipped += 1
                    continue

                if not product and not imei:
                    skipped += 1
                    continue

                self.db.save_device_exclusion(product, imei, comments, district)
                imported += 1

            self.refresh_device_exclusions_table()
            self.refresh_table()
            QMessageBox.information(self, "Import complete", f"Imported {imported} exclusion row(s).\nSkipped {skipped} invalid row(s).\n\nProduct exclusions require District. IMEI exclusions are global.")
            self.set_status(f"Imported excluded devices. Imported: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Excluded devices import failed.")

    def pick_inventory(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Select Inventory_Count_Result_Details", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if path:
            self.inventory_path = path

    def pick_time_sheet(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Select Timesheet (timesheets_*.xlsx)", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if path:
            self.time_sheet_path = path

    def load_variances(self) -> None:
        inventory = self.inventory_path.strip()
        time_sheet = self.time_sheet_path.strip()
        if not inventory or not os.path.exists(inventory):
            QMessageBox.critical(self, "Missing file", "Select Inventory_Count_Result_Details.xlsx first.")
            return
        if not time_sheet or not os.path.exists(time_sheet):
            QMessageBox.critical(self, "Missing file", "Select Timesheet file (timesheets_*.xlsx) first.")
            return
        try:
            self.clear_current_ui(silent=True)
            self.set_status("Reading Excel files...")
            inv_records = read_xlsx_records(inventory)
            ts_records = read_xlsx_records(time_sheet)
            self.current_inventory_records = inv_records
            self.current_time_sheet_records = ts_records
            self.master_store_records = self.db.store_master_records()
            self.set_status("Building Inventory Status and Variance Audit rows...")

            self.status_rows, status_summary = build_inventory_status_rows(inv_records, ts_records, master_store_records=self.master_store_records, source_file=os.path.basename(inventory))
            self.status_row_by_key = {r.key: r for r in self.status_rows}
            self.db.upsert_inventory_status_rows(self.status_rows)
            rows, variance_summary = extract_variances(inv_records, ts_records, master_store_records=self.master_store_records, source_file=os.path.basename(inventory))
            self.db.upsert_rows(rows)
            self.loaded_keys = {row.key for row in rows}
            self.data_loaded = True

            self.populate_status_filters()
            self.populate_audit_filters()
            self.refresh_status_table()
            self.refresh_table()

            self.summary_text = (
                f"Inventory Status. Completed: {status_summary['completed']}   Pending: {status_summary['pending']}   "
                f"Variance Audit: {len(rows)}   SIMs skipped: {variance_summary['skipped_sims']}"
            )
            self.set_status(
                f"Loaded current session. Inventory Status rows: {len(self.status_rows)}. Variances: {len(rows)}. "
                f"Inventory rows read: {variance_summary['raw_inventory_rows']}. Latest rows used: {variance_summary['latest_inventory_rows']}. "
                f"Stale rows skipped: {variance_summary['stale_inventory_rows']}. SIM rows skipped: {variance_summary['skipped_sims']}."
            )
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Load failed", str(exc))
            self.set_status("Load failed. Check the Excel columns and try again.")

    def add_store_prompt(self) -> None:
        district, ok = QInputDialog.getText(self, "Add Store", "Enter district name.")
        if not ok:
            return
        district = normalize_district(district)
        if district == "Unknown":
            QMessageBox.critical(self, "Invalid district", "Please enter a valid district name.")
            return

        store, ok2 = QInputDialog.getText(self, "Add Store", "Enter store name.")
        if not ok2:
            return
        store = display_store(store)
        if not store:
            QMessageBox.critical(self, "Invalid store", "Store name is required.")
            return

        self.db.save_store_account(district, store, "", "", "")
        self.master_store_records = self.db.store_master_records()
        self.refresh_store_accounts_table()
        self.refresh_entry_form_districts()

        if self.data_loaded:
            source_name = os.path.basename(self.inventory_path.strip()) if self.inventory_path.strip() else ""
            self.status_rows, status_summary = build_inventory_status_rows(
                self.current_inventory_records,
                self.current_time_sheet_records,
                master_store_records=self.master_store_records,
                source_file=source_name,
            )
            self.status_row_by_key = {r.key: r for r in self.status_rows}
            self.populate_status_filters()
            self.refresh_status_table()
            self.summary_text = (
                f"Inventory Audit Status. Completed: {status_summary['completed']}   Pending: {status_summary['pending']}   Variance Audit: {len(self.loaded_keys)}"
            )
        else:
            self.populate_status_filters()

        QMessageBox.information(self, "Store saved", f"Saved store.\nDistrict: {district}\nStore: {store}")
        self.set_status(f"Added store: {district} | {store}")

    def populate_status_filters(self) -> None:
        districts = set()
        for r in self.status_rows:
            d = normalize_district(r.district or "Unknown")
            if d and d != "Unknown":
                districts.add(d)
        for d in self.db.all_known_districts():
            if d and d != "Unknown":
                districts.add(d)
        district_values = ["All Districts"] + sorted(districts)
        current = self.status_district_combo.currentText()
        self.status_district_combo.blockSignals(True)
        self.status_district_combo.clear()
        self.status_district_combo.addItems(district_values)
        self._combo_set(self.status_district_combo, current if current in district_values else "All Districts")
        self.status_district_combo.blockSignals(False)
        self.update_status_store_filter_values()

    def update_status_store_filter_values(self) -> None:
        selected_district = self.status_district_combo.currentText().strip()
        rows = self.status_rows
        if selected_district and selected_district != "All Districts":
            rows = [r for r in rows if normalize_district(r.district) == normalize_district(selected_district)]
        stores = sorted({r.store for r in rows if r.store})
        values = ["All Stores"] + stores
        current = self.status_store_combo.currentText()
        self.status_store_combo.blockSignals(True)
        self.status_store_combo.clear()
        self.status_store_combo.addItems(values)
        self._combo_set(self.status_store_combo, current if current in values else "All Stores")
        self.status_store_combo.blockSignals(False)

    def populate_audit_filters(self) -> None:
        districts = set()
        rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        for r in rows:
            d = normalize_district(r.district or "Unknown")
            if d and d != "Unknown":
                districts.add(d)
        for d in self.db.all_known_districts():
            if d and d != "Unknown":
                districts.add(d)
        district_values = ["All Districts"] + sorted(districts)
        current = self.audit_district_combo.currentText()
        self.audit_district_combo.blockSignals(True)
        self.audit_district_combo.clear()
        self.audit_district_combo.addItems(district_values)
        self._combo_set(self.audit_district_combo, current if current in district_values else "All Districts")
        self.audit_district_combo.blockSignals(False)
        # Also populate the final district combo
        if hasattr(self, "final_district_combo"):
            final_current = self.final_district_combo.currentText()
            self.final_district_combo.blockSignals(True)
            self.final_district_combo.clear()
            self.final_district_combo.addItems(district_values)
            self._combo_set(self.final_district_combo, final_current if final_current in district_values else "All Districts")
            self.final_district_combo.blockSignals(False)
        self.update_audit_store_filter_values()

    def update_audit_store_filter_values(self) -> None:
        rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        district = self.audit_district_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [r for r in rows if normalize_district(r.district) == normalize_district(district)]
        stores = sorted({r.store for r in rows if r.store})
        values = ["All Stores"] + stores
        current = self.audit_store_combo.currentText()
        self.audit_store_combo.blockSignals(True)
        self.audit_store_combo.clear()
        self.audit_store_combo.addItems(values)
        self._combo_set(self.audit_store_combo, current if current in values else "All Stores")
        self.audit_store_combo.blockSignals(False)

    def on_status_district_change(self, _event=None) -> None:
        self.update_status_store_filter_values()
        self.auto_check_status_rows()
        self.refresh_status_table()

    def on_status_store_change(self, _event=None) -> None:
        self.auto_check_status_rows()
        self.refresh_status_table()

    def on_audit_district_change(self, _event=None) -> None:
        self.update_audit_store_filter_values()
        self.auto_check_audit_rows()
        self.refresh_table()

    def on_audit_store_change(self, _event=None) -> None:
        self.auto_check_audit_rows()
        self.refresh_table()

    def matching_status_rows(self) -> List[InventoryStatusRow]:
        rows = list(self.status_rows)
        district = self.status_district_combo.currentText().strip()
        store = self.status_store_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [r for r in rows if normalize_district(r.district) == normalize_district(district)]
        if store and store != "All Stores":
            rows = [r for r in rows if r.store == store]

        any_q = safe_text(self.status_search_any_edit.text()).lower()
        district_q = safe_text(self.status_search_district_edit.text()).lower()
        store_q = safe_text(self.status_search_store_edit.text()).lower()
        status_q = safe_text(self.status_search_status_edit.text()).lower()
        rep_q = safe_text(self.status_search_rep_edit.text()).lower()

        if district_q:
            rows = [r for r in rows if district_q in safe_text(r.district).lower()]
        if store_q:
            rows = [r for r in rows if store_q in safe_text(r.store).lower()]
        if status_q:
            rows = [r for r in rows if status_q in safe_text(r.status).lower()]
        if rep_q:
            rows = [r for r in rows if rep_q in safe_text(r.rep_name).lower()]
        if any_q:
            rows = [
                r for r in rows
                if any_q in " ".join([
                    safe_text(r.district),
                    safe_text(r.store),
                    safe_text(r.status),
                    safe_text(r.rep_name),
                ]).lower()
            ]
        return rows

    def clear_status_search(self) -> None:
        self.status_search_any_edit.clear()
        self.status_search_district_edit.clear()
        self.status_search_store_edit.clear()
        self.status_search_status_edit.clear()
        self.status_search_rep_edit.clear()
        self.refresh_status_table()

    def refresh_status_table(self) -> None:
        self.status_tree.setRowCount(0)
        self.status_key_by_iid.clear()
        if not self.data_loaded:
            return
        rows = self.matching_status_rows()
        rows.sort(key=lambda r: (r.district, r.store, r.rep_name, r.status))
        state_map = self.db.inventory_status_state_map()
        for row in rows:
            checkbox = "☑ Send" if row.key in self.status_checked_keys else "☐ Skip"
            tag = self.status_row_tag(row, state_map)
            iid = self.status_tree.rowCount()
            self.status_tree.insertRow(iid)
            values = (row.district, row.store, row.status, row.rep_name, checkbox)
            for col, value in enumerate(values):
                self.status_tree.setItem(iid, col, QTableWidgetItem(str(value)))
            self._paint_row(self.status_tree, iid, tag)
            self.status_tree.item(iid, 0).setData(Qt.UserRole, row.key)
            self.status_key_by_iid[str(iid)] = row.key
        checked = len([r for r in rows if r.key in self.status_checked_keys])
        self.set_status(f"Inventory Status rows shown: {len(rows)}. Checked for send: {checked}.")

    def on_status_tree_click(self, item) -> None:
        if item.column() != 4:  # checkbox column
            return
        row = item.row()
        key = self.status_tree.item(row, 0).data(Qt.UserRole)
        if not key:
            return
        if key in self.status_checked_keys:
            self.status_checked_keys.remove(key)
        else:
            self.status_checked_keys.add(key)
        self.refresh_status_table()

    def matching_audit_rows(self) -> List[VarianceRow]:
        rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        if not self.include_cleared_check.isChecked():
            rows = [row for row in rows if not row.cleared]
        district = self.audit_district_combo.currentText().strip()
        store = self.audit_store_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [row for row in rows if normalize_district(row.district) == normalize_district(district)]
        if store and store != "All Stores":
            rows = [row for row in rows if row.store == store]

        any_q = safe_text(self.audit_search_any_edit.text()).lower()
        district_q = safe_text(self.audit_search_district_edit.text()).lower()
        store_q = safe_text(self.audit_search_store_edit.text()).lower()
        product_q = safe_text(self.audit_search_product_edit.text()).lower()
        imei_q = safe_text(self.audit_search_imei_edit.text()).lower()
        rep_q = safe_text(self.audit_search_rep_edit.text()).lower()

        if district_q:
            rows = [row for row in rows if district_q in safe_text(row.district).lower()]
        if store_q:
            rows = [row for row in rows if store_q in safe_text(row.store).lower()]
        if product_q:
            rows = [row for row in rows if product_q in safe_text(row.product).lower()]
        if imei_q:
            rows = [row for row in rows if imei_q in safe_text(row.imei).lower()]
        if rep_q:
            rows = [row for row in rows if rep_q in safe_text(row.rep_name).lower()]
        if any_q:
            rows = [
                row for row in rows
                if any_q in " ".join([
                    safe_text(row.district),
                    safe_text(row.store),
                    safe_text(row.product),
                    safe_text(row.imei),
                    safe_text(row.status),
                    safe_text(row.rep_name),
                ]).lower()
            ]

        rows = self.filter_excluded_variance_rows(rows)
        return rows

    def clear_audit_search(self) -> None:
        self.audit_search_any_edit.clear()
        self.audit_search_district_edit.clear()
        self.audit_search_store_edit.clear()
        self.audit_search_product_edit.clear()
        self.audit_search_imei_edit.clear()
        self.audit_search_rep_edit.clear()
        self.refresh_table()

    def copy_text_to_clipboard(self, text_value: str, status_message: str = "Copied to clipboard.") -> None:
        text_value = safe_text(text_value)
        if not text_value:
            return
        try:
            QApplication.clipboard().setText(text_value)
            self.set_status(status_message)
        except Exception as exc:
            QMessageBox.critical(self, "Copy failed", str(exc))

    def copy_selected_imei(self, event=None) -> None:
        rows = sorted({i.row() for i in self.audit_tree.selectedIndexes()}) if hasattr(self, "audit_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a variance row first.")
            return None
        imeis: List[str] = []
        for row in rows:
            imei_item = self.audit_tree.item(row, 3)
            if imei_item is not None:
                imei = safe_text(imei_item.text())
                if imei:
                    imeis.append(imei)
        if not imeis:
            QMessageBox.information(self, "No IMEI", "No IMEI found in selected row.")
            return None
        self.copy_text_to_clipboard("\n".join(imeis), f"Copied {len(imeis)} IMEI(s) to clipboard.")
        return None

    def refresh_table(self) -> None:
        self.audit_tree.setRowCount(0)
        self.key_by_iid.clear()

        if not self.data_loaded:
            self.summary_text = "No data loaded"
            self.set_status("UI cleared. Select both file locations, then click Load Variances.")
            return

        rows = self.matching_audit_rows()
        rows.sort(key=lambda r: (r.district, r.store, r.rep_name, r.status, r.product, r.imei))

        for row in rows:
            clearance = "Cleared" if row.cleared else ("Sent" if row.sent_count else "Not cleared")
            checkbox = "☑ Send" if row.key in self.audit_checked_keys else "☐ Skip"
            iid = self.audit_tree.rowCount()
            self.audit_tree.insertRow(iid)
            values = (row.district, row.store, row.product, row.imei, row.status, row.rep_name, clearance, checkbox)
            for col, value in enumerate(values):
                self.audit_tree.setItem(iid, col, QTableWidgetItem(str(value)))
            self._paint_row(self.audit_tree, iid, self.audit_row_tag(row))
            self.audit_tree.item(iid, 0).setData(Qt.UserRole, row.key)
            self.key_by_iid[str(iid)] = row.key
        cleared_count = sum(1 for r in rows if r.cleared)
        pending_count = sum(1 for r in rows if not r.cleared)
        checked = sum(1 for r in rows if r.key in self.audit_checked_keys)
        self.set_status(f"Variance Audit rows shown: {len(rows)}. Cleared: {cleared_count}. Not cleared: {pending_count}. Checked for send: {checked}.")

    def on_audit_double_click(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.audit_tree.selectedIndexes()})
        if not rows:
            return
        key = self.audit_tree.item(rows[0], 0).data(Qt.UserRole)
        if not key or key not in self.key_by_iid.values():
            return
        self.send_selected()

    def on_audit_tree_click(self, item) -> None:
        row = item.row()
        key = self.audit_tree.item(row, 0).data(Qt.UserRole)
        if item.column() == 3:  # IMEI column: click copies the IMEI
            imei_item = self.audit_tree.item(row, 3)
            if imei_item is not None:
                self.copy_text_to_clipboard(imei_item.text(), f"Copied IMEI: {imei_item.text()}")
            return
        if item.column() != 7:  # checkbox column
            return
        if not key:
            return
        if key in self.audit_checked_keys:
            self.audit_checked_keys.remove(key)
        else:
            self.audit_checked_keys.add(key)
        self.refresh_table()

    def selected_rows(self) -> List[VarianceRow]:
        rows = sorted({i.row() for i in self.audit_tree.selectedIndexes()})
        keys = [self.audit_tree.item(row, 0).data(Qt.UserRole) for row in rows]
        keys = [k for k in keys if k]
        return self.db.get_rows_by_keys(keys)

    def current_mode_key(self) -> str:
        label = self.audit_send_mode_combo.currentText().strip().lower()
        if label == "store":
            return "store"
        if label in {"sales rep", "rep", "salesperson"}:
            return "rep"
        return "district"

    def current_status_mode_key(self) -> str:
        label = self.status_send_mode_combo.currentText().strip().lower()
        if label == "store":
            return "store"
        if label in {"sales rep", "rep", "salesperson"}:
            return "rep"
        return "district"
    @staticmethod
    def _mix_hex(a: str, b: str, t: float) -> str:
        """Blend two #rrggbb colours (t=0 -> a, t=1 -> b)."""
        try:
            ar, ag, ab = (int(a.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
            br, bg, bb = (int(b.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
            return "#{:02x}{:02x}{:02x}".format(
                round(ar + (br - ar) * t),
                round(ag + (bg - ag) * t),
                round(ab + (bb - ab) * t))
        except Exception:
            return a
    def filter_excluded_variance_rows(self, rows: List[VarianceRow]) -> List[VarianceRow]:
        if not rows:
            return []
        return [row for row in rows if not self.db.is_device_excluded(row.district, row.product, row.imei)]
    def status_row_tag(self, row: InventoryStatusRow, state_map: Optional[Dict[str, Dict[str, str]]] = None) -> str:
        state_map = state_map if state_map is not None else self.db.inventory_status_state_map()
        state = state_map.get(normalize_store(row.store), {})
        status_norm = normalize_header(row.status)
        previous_norm = normalize_header(state.get("Previous Status", ""))
        sent_at = safe_text(state.get("Last Sent At", ""))

        if status_norm != "completed":
            return "status_pending"
        if sent_at:
            return "status_completed_sent"
        if previous_norm == "pending":
            return "status_completed_after_update"
        return "status_completed"

    def auto_check_pending_status_rows(self) -> None:
        rows = [r for r in self.matching_status_rows() if normalize_header(r.status) != "completed"]
        self.status_checked_keys = {r.key for r in rows}
        self.refresh_status_table()
        self.set_status(f"Prepared {len(rows)} pending inventory status row(s) for sending.")

    def auto_check_status_rows(self) -> None:
        rows = self.matching_status_rows()
        self.status_checked_keys = {r.key for r in rows}
        self.refresh_status_table()
        self.set_status(f"Prepared {len(rows)} inventory status row(s) for sending. Uncheck any row you want to skip.")

    def clear_status_checkmarks(self) -> None:
        self.status_checked_keys.clear()
        self.refresh_status_table()
        self.set_status("Inventory status checkmarks cleared.")
    def auto_check_audit_rows(self) -> None:
        rows = self.matching_audit_rows()
        self.audit_checked_keys = {r.key for r in rows}
        self.refresh_table()
        self.set_status(f"Prepared {len(rows)} variance row(s) for sending. Uncheck any row you want to skip.")

    def clear_audit_checkmarks(self) -> None:
        self.audit_checked_keys.clear()
        self.refresh_table()
        self.set_status("Variance checkmarks cleared.")
    @staticmethod
    def audit_row_tag(row: VarianceRow) -> str:
        if row.cleared:
            return "variance_cleared"
        if row.sent_count:
            return "variance_sent"
        return "variance_pending"
    @staticmethod
    def grouped_batches(rows: List[VarianceRow], mode: str) -> List[Tuple[str, str, List[VarianceRow]]]:
        grouped: Dict[Tuple[str, str], List[VarianceRow]] = {}
        for row in rows:
            district = normalize_district(row.district or "Unknown")
            if mode == "store":
                key = (district, row.store or "Unknown Store")
            elif mode == "rep":
                key = (district, row.rep_name or "Unknown Rep")
            else:
                key = (district, district)
            grouped.setdefault(key, []).append(row)
        batches: List[Tuple[str, str, List[VarianceRow]]] = []
        for (district, value), group_rows in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1])):
            title = district if mode == "district" else f"{value} | {district}"
            batches.append((title, district, group_rows))
        return batches

    @staticmethod
    def grouped_status_batches(rows: List[InventoryStatusRow], mode: str) -> List[Tuple[str, str, List[InventoryStatusRow]]]:
        grouped: Dict[Tuple[str, str], List[InventoryStatusRow]] = {}
        for row in rows:
            district = normalize_district(row.district or "Unknown")
            if mode == "store":
                key = (district, row.store or "Unknown Store")
            elif mode == "rep":
                key = (district, row.rep_name or "Unknown Rep")
            else:
                key = (district, district)
            grouped.setdefault(key, []).append(row)
        batches: List[Tuple[str, str, List[InventoryStatusRow]]] = []
        for (district, value), group_rows in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1])):
            title = district if mode == "district" else f"{value} | {district}"
            batches.append((title, district, group_rows))
        return batches

    @staticmethod
    def chunks(rows: List, size: int) -> Iterable[List]:
        for i in range(0, len(rows), size):
            yield rows[i:i + size]
    def inventory_status_message(self) -> str:
        today = dt.datetime.now()
        return f"Inventory Audit Status as of {today.month}/{today.day}/{today.year}."

    def variance_request_message(self, rows: List[VarianceRow]) -> str:
        """Build WhatsApp @mention message for variance request.

        Uses resolve_phone_for_rep() which tries:
        1. created_by_mappings by 'Created By' username (ArmanAli → phone)
        2. created_by_mappings by employee_name (Abdullah Hussain → phone)
        3. sales_reps table by rep_name
        """
        seen_phones: set[str] = set()
        phones: List[str] = []
        missing: List[str] = []

        for row in rows:
            rep_name = safe_text(row.rep_name)
            created_by = safe_text(row.created_by)
            phone = normalize_phone(
                self.db.resolve_phone_for_rep(rep_name, created_by)
            )
            if not phone:
                missing.append(created_by or rep_name or "Unknown")
                continue
            if phone in seen_phones:
                continue
            seen_phones.add(phone)
            phones.append(whatsapp_mention(phone))

        if missing:
            unique_missing = sorted(set(missing))
            self.set_status(
                "No phone found for: " + ", ".join(unique_missing[:5]) +
                (f"... (+{len(unique_missing)-5} more)" if len(unique_missing) > 5 else "") +
                " — add phone in the Employees tab (Created By column)."
            )

        if not phones:
            return ""

        return " ".join(phones) + " please share the images of the variances."
    def pending_inventory_count_message(self, rows: List[InventoryStatusRow]) -> str:
        """Build the per-store reminder sent after the Inventory Audit Status
        image.

        Each pending store is addressed by the PERSON responsible whenever
        the row has one (the Salesperson column — which also carries the
        timesheet-matched employee for stores absent from the count file):

        1. Rep has a phone on file (Employees tab / sales reps): tag it —
               @<phone>, please complete the inventory count ASAP.
        2. Rep known but no phone anywhere: address the person by name —
               <Rep Name>, please complete the inventory count ASAP.
        3. No rep on the row: fall back to the store name —
               <Store Name>, please complete the inventory count ASAP.

        One line per pending store, blocks separated by a blank line.
        """
        messages: List[str] = []
        seen_stores: set[str] = set()

        for row in rows:
            if safe_text(row.status).lower() == "completed":
                continue

            store_name = safe_text(row.store)
            if not store_name or store_name in seen_stores:
                continue
            seen_stores.add(store_name)

            rep_name = safe_text(row.rep_name)
            if rep_name:
                phone = normalize_phone(self.db.resolve_phone_for_rep(rep_name))
                if phone:
                    messages.append(
                        f"{whatsapp_mention(phone)}, please complete the inventory count ASAP."
                    )
                    continue
                messages.append(
                    f"{rep_name}, please complete the inventory count ASAP."
                )
                continue

            messages.append(f"{store_name}, please complete the inventory count ASAP.")

        return "\n\n".join(messages)
    def create_full_inventory_audit_log(self, final_results: List[Dict[str, str]]) -> Path:
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = EXPORT_DIR / f"GFH_Full_Inventory_Audit_Log_{stamp}.xlsx"
        variance_rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        final_results = final_results or []

        wb = openpyxl.Workbook()
        ws = wb.active
        headers = [
            "Section", "District", "Store", "Product", "IMEI", "Status", "Rep Name",
            "Clearance Status", "Sent Count", "Last Sent At", "Final Caption",
            "Pending Inventory Stores", "Final Image Path", "WhatsApp Group", "Generated At"
        ]
        ws.append(headers)
        
        imei_col_idx = headers.index("IMEI") + 1
        
        current_row = 2

        for row in sorted(self.status_rows, key=lambda r: (r.district, r.store, r.rep_name)):
            data = [
                "Inventory Status", row.district, row.store, "", "", row.status, row.rep_name,
                "", "", "", "", "", "", "", now_text()
            ]
            for c_idx, val in enumerate(data, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'
            current_row += 1

        for row in sorted(variance_rows, key=lambda r: (r.district, r.store, r.rep_name, r.product, r.imei)):
            data = [
                "Variance Audit", row.district, row.store, row.product, row.imei, row.status, row.rep_name,
                "Cleared" if row.cleared else "Not Cleared", row.sent_count, row.last_sent_at,
                "", "", "", "", now_text()
            ]
            for c_idx, val in enumerate(data, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'
            current_row += 1

        for item in final_results:
            data = [
                "Final District Send", item.get("district", ""), item.get("stores", ""), "", "",
                item.get("status", ""), "", "", "", "", item.get("caption", ""),
                item.get("pending_inventory_stores", ""), item.get("image_path", ""),
                item.get("group_name", ""), item.get("sent_at", "")
            ]
            for c_idx, val in enumerate(data, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'
            current_row += 1

        wb.save(output_path)
        return output_path
    def _send_rows_thread(self, rows: List[VarianceRow], mode: str, manual: bool) -> None:
        thread = threading.Thread(target=self._send_rows, args=(rows, mode, manual), daemon=True)
        thread.start()
    def _render_status_rows(self, batch_title: str, rows: List[InventoryStatusRow], mode: str = "status") -> Path:
        if Image is None:
            raise RuntimeError("Pillow is required. Install with: py -m pip install pillow")

        renderer = ImageRenderer()
        width = 1320
        margin = 28
        title_font = renderer._font(28, True)
        header_font = renderer._font(16, True)
        cell_font = renderer._font(15, False)
        cell_bold = renderer._font(15, True)
        sub_font = renderer._font(15, False)
        row_h = 34
        black = (18, 20, 43)
        border = (0, 0, 0)
        light_green = (210, 240, 214)
        pending_fill = (238, 230, 177)
        white = (255, 255, 255)
        # Standard GFH brand bar colors (matches gfh_inventory_aging_processor's
        # wa_render_table_image and every other GFH/VDS app's window header):
        # navy bar, logo + red divider on the left, title centered in white.
        navy = (9, 13, 38)
        red = (232, 48, 48)
        district_colors = {
            "Arizona": (196, 138, 230),
            "Atlanta": (189, 49, 120),
            "Colorado": (96, 194, 236),
            "Colorado East": (96, 194, 236),
            "Colorado West": (78, 186, 228),
            "Houston": (244, 128, 128),
            "Louisiana": (247, 199, 34),
            "Tennessee": (74, 145, 62),
        }

        logo_img = None
        logo_w = 0
        logo_h_used = 0
        if STATUS_LOGO_PATH.exists():
            try:
                logo_img = Image.open(STATUS_LOGO_PATH).convert("RGBA")
                scale = min(150 / logo_img.width, 56 / logo_img.height)
                size = (max(1, int(logo_img.width * scale)), max(1, int(logo_img.height * scale)))
                logo_img = logo_img.resize(size)
                logo_w, logo_h_used = size
            except Exception:
                logo_img = None
                logo_w = 0
                logo_h_used = 0

        brand_bar_h = 90   # matches FixedHeaderManager's standard GFH header height
        table_top = brand_bar_h + 18
        height = table_top + 26 + row_h * len(rows) + 24

        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        draw.rectangle([0, 0, width, brand_bar_h], fill=navy)
        logo_right_edge = margin
        if logo_img is not None:
            img.paste(logo_img, (margin, (brand_bar_h - logo_h_used) // 2), logo_img)
            logo_right_edge = margin + logo_w
        divider_x = logo_right_edge + 14
        draw.rectangle([divider_x, 14, divider_x + 3, brand_bar_h - 14], fill=red)

        title_text = "GFH Inventory Audit"
        today = dt.datetime.now()
        date_text = f"{today.month}/{today.day}/{today.year}"
        if safe_text(batch_title):
            date_text += f"   {safe_text(batch_title)}"

        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_w = title_bbox[2] - title_bbox[0]
        date_bbox = draw.textbbox((0, 0), date_text, font=sub_font)
        date_w = date_bbox[2] - date_bbox[0]
        draw.text(((width - title_w) // 2, 16), title_text, fill=white, font=title_font)
        draw.text(((width - date_w) // 2, 16 + 36), date_text, fill=white, font=sub_font)

        y = table_top
        x = margin
        col_widths = [160, 280, 145, 679]
        headers = ["District", "Store", "Status", "Salesperson"]
        draw.rectangle((x, y, width - margin, y + 26), outline=border, fill=white)
        cx = x
        for idx, h in enumerate(headers):
            draw.line((cx, y, cx, y + 26), fill=border, width=1)
            draw.text((cx + 4, y + 4), h, fill=black, font=header_font)
            cx += col_widths[idx]
        draw.line((width - margin, y, width - margin, y + 26), fill=border, width=1)
        draw.line((x, y + 26, width - margin, y + 26), fill=border, width=1)
        y += 26

        rows_sorted = sorted(rows, key=lambda r: (normalize_district(r.district), r.store))
        for row in rows_sorted:
            district = normalize_district(row.district)
            row_fill = district_colors.get(district, (245, 245, 245))
            draw.rectangle((x, y, width - margin, y + row_h), fill=row_fill, outline=border)
            vals = [district, row.store, row.status, row.rep_name or "-"]
            cx = x
            for cidx, val in enumerate(vals):
                if cidx == 2:
                    status_fill = light_green if safe_text(row.status).lower() == "completed" else pending_fill
                    draw.rectangle((cx, y, cx + col_widths[cidx], y + row_h), fill=status_fill, outline=border)
                draw.line((cx, y, cx, y + row_h), fill=border, width=1)
                text_val = safe_text(val) or "-"
                if len(text_val) > 42 and cidx == 3:
                    text_val = text_val[:39] + "..."
                if len(text_val) > 22 and cidx == 1:
                    text_val = text_val[:19] + "..."
                font = cell_bold if cidx in (0, 1, 3) else cell_font
                fill = (210, 0, 0) if (cidx == 2 and safe_text(row.status).lower() != "completed") else black
                draw.text((cx + 4, y + 6), text_val, fill=fill, font=font)
                cx += col_widths[cidx]
            draw.line((width - margin, y, width - margin, y + row_h), fill=border, width=1)
            y += row_h

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", batch_title)[:60] or "Inventory_Audit_Status"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = IMAGE_DIR / f"GFH_Inventory_Audit_Status_{safe_title}_{stamp}.png"
        img.save(path)
        return path
    @staticmethod
    def format_store_list_caption(stores: List[str]) -> str:
        clean = [safe_text(s) for s in stores if safe_text(s)]
        if not clean:
            return ""
        if len(clean) == 1:
            return clean[0]
        if len(clean) == 2:
            return f"{clean[0]} and {clean[1]}"
        return ", ".join(clean[:-1]) + f", and {clean[-1]}"

    def final_district_caption(self, district: str, rows: List[VarianceRow], pending_status_rows: Optional[List[InventoryStatusRow]] = None) -> str:
        parts: List[str] = []
        dm_phone = self.db.find_district_manager_phone(district)
        if dm_phone:
            parts.append(whatsapp_mention(dm_phone))

        stores = sorted({safe_text(r.store) for r in rows if safe_text(r.store)})
        if stores:
            store_text = self.format_store_list_caption(stores)
            parts.append(f"An Inventory audit has been completed, variance found in {store_text}.")
        else:
            parts.append("An Inventory audit has been completed and no variance found.")

        pending_stores = sorted({safe_text(r.store) for r in (pending_status_rows or []) if safe_text(r.store)})
        if pending_stores:
            pending_text = self.format_store_list_caption(pending_stores)
            verb = "did not complete" if len(pending_stores) == 1 else "did not complete"
            parts.append(f"{pending_text} {verb} an inventory count.")

        return " ".join(parts)
    def _render_no_variance_image(self, district: str) -> Path:
        if Image is None:
            raise RuntimeError("Pillow is required. Install with: py -m pip install pillow")

        renderer = ImageRenderer()
        width = 1360
        margin = 32
        title_font = renderer._font(34, True)
        sub_font = renderer._font(20, False)
        body_font = renderer._font(28, True)
        note_font = renderer._font(18, False)

        dark = (18, 20, 43)
        gray = (95, 95, 102)
        red = (233, 27, 47)
        border = (215, 218, 223)
        navy = (9, 13, 38)
        white = (255, 255, 255)

        logo_img = None
        logo_w = 0
        logo_h_used = 0
        if STATUS_LOGO_PATH.exists():
            try:
                logo_img = Image.open(STATUS_LOGO_PATH).convert("RGBA")
                scale = min(150 / logo_img.width, 56 / logo_img.height)
                size = (max(1, int(logo_img.width * scale)), max(1, int(logo_img.height * scale)))
                logo_img = logo_img.resize(size)
                logo_w, logo_h_used = size
            except Exception:
                logo_img = None
                logo_w = 0
                logo_h_used = 0

        # Standard GFH brand bar (matches gfh_inventory_aging_processor's
        # wa_render_table_image and every other GFH/VDS app's window header).
        brand_bar_h = 90
        box_top = brand_bar_h + 16
        height = box_top + 190 + 40
        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        draw.rectangle([0, 0, width, brand_bar_h], fill=navy)
        logo_right_edge = margin
        if logo_img is not None:
            img.paste(logo_img, (margin, (brand_bar_h - logo_h_used) // 2), logo_img)
            logo_right_edge = margin + logo_w
        divider_x = logo_right_edge + 14
        draw.rectangle([divider_x, 14, divider_x + 3, brand_bar_h - 14], fill=red)

        title_text = "GFH Inventory Variance"
        sub_text = f"{district}   Generated: {now_text()}"
        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_w = title_bbox[2] - title_bbox[0]
        sub_bbox = draw.textbbox((0, 0), sub_text, font=sub_font)
        sub_w = sub_bbox[2] - sub_bbox[0]
        draw.text(((width - title_w) // 2, 16), title_text, fill=white, font=title_font)
        draw.text(((width - sub_w) // 2, 16 + 38), sub_text, fill=white, font=sub_font)

        draw.rectangle((margin, box_top, width - margin, box_top + 160), outline=border, fill=(255, 255, 255))
        draw.rectangle((margin, box_top, width - margin, box_top + 46), fill=red)
        draw.text((margin + 14, box_top + 11), "Final Audit Result", fill="white", font=sub_font)
        draw.text((margin + 18, box_top + 78), "No variance found in this district.", fill=dark, font=body_font)
        draw.text((margin + 18, box_top + 118), "All variances were cleared or none were found.", fill=gray, font=note_font)

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", district)[:60] or "District"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = IMAGE_DIR / f"GFH_Final_Audit_{safe_title}_{stamp}.png"
        img.save(path)
        return path
    def mark_selected(self, cleared: bool) -> None:
        rows = self.selected_rows()
        if not rows:
            QMessageBox.information(self, "No selection", "Select one or more variance rows first.")
            return
        def _do():
            for row in rows:
                self.db.set_cleared(row.key, cleared)
        self._db_write(_do)
        self.refresh_table()
        action = "cleared" if cleared else "not cleared"
        self.set_status(f"Marked {len(rows)} variance row(s) as {action}.")

    def clear_current_ui(self, silent: bool = False) -> None:
        self.loaded_keys.clear()
        self.data_loaded = False
        self.status_rows = []
        self.status_row_by_key = {}
        self.status_checked_keys.clear()
        self.audit_checked_keys.clear()
        self.status_key_by_iid.clear()
        self.populate_status_filters()
        self.populate_audit_filters()
        self.refresh_status_table()
        self.refresh_table()
        if not silent:
            self.set_status("UI cleared. Select both file locations, then click Load Variances.")
    def _check_db_sync(self) -> None:
        """Called every 2 s.  If another instance modified the DB, refresh GUI."""
        try:
            if not self._db_sync_paused and DB_PATH.exists():
                mtime = DB_PATH.stat().st_mtime
                if mtime != self._db_mtime:
                    self._db_mtime = mtime
                    # Reload master data in case store/rep lists changed too
                    try:
                        self.master_store_records = self.db.store_master_records()
                    except Exception:
                        pass
                    # Refresh both visible tables silently
                    try:
                        self.refresh_table()
                    except Exception:
                        pass
                    try:
                        self.refresh_status_table()
                    except Exception:
                        pass
        except Exception:
            pass
        self._schedule_db_sync()
    def _db_write(self, fn, *args, **kwargs):
        """Wrap any DB write so the sync poller does not re-trigger on our own write."""
        self._db_sync_paused = True
        try:
            result = fn(*args, **kwargs)
        finally:
            # Update our own mtime baseline after writing, then re-enable polling
            try:
                self._db_mtime = DB_PATH.stat().st_mtime
            except Exception:
                pass
            self._db_sync_paused = False
        return result

    def open_app_folder(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(APP_DIR))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(APP_DIR)])
            else:
                subprocess.Popen(["xdg-open", str(APP_DIR)])
        except Exception as exc:
            QMessageBox.critical(self, "Open folder failed", str(exc))

    def send_checked_variances(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        rows = [row for row in self.db.get_rows_by_keys(self.loaded_keys) if row.key in self.audit_checked_keys]
        rows = [row for row in rows if self.include_cleared_check.isChecked() or not row.cleared]
        rows = self.filter_excluded_variance_rows(rows)
        if not rows:
            QMessageBox.information(self, "Nothing checked", "Select a district or store so variance rows get checked, then send them.")
            return
        mode = self.current_mode_key()
        batches = self.grouped_batches(rows, mode)
        if QMessageBox.question(self, "Send checked variances", f"Send {len(rows)} checked variance row(s) in {len(batches)} image batch(es)?") != QMessageBox.Yes:
            return
        self._send_rows_thread(rows, mode=mode, manual=True)

    def send_selected(self) -> None:
        rows = self.filter_excluded_variance_rows(self.selected_rows())
        if not rows:
            QMessageBox.information(self, "No selection", "Select one or more non-excluded variance rows first.")
            return
        mode = self.current_mode_key()
        self._send_rows_thread(rows, mode=mode, manual=True)

    def send_pending(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        mode = self.current_mode_key()
        rows = [row for row in self.db.get_rows_by_keys(self.loaded_keys) if not row.cleared]
        district = self.audit_district_combo.currentText().strip()
        store = self.audit_store_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [row for row in rows if normalize_district(row.district) == normalize_district(district)]
        if store and store != "All Stores":
            rows = [row for row in rows if row.store == store]
        rows = self.filter_excluded_variance_rows(rows)
        if self.send_only_unsent_check.isChecked():
            rows = [row for row in rows if not row.last_sent_at]
        if not rows:
            QMessageBox.information(self, "Nothing to send", "No pending unsent variances found in the current filter.")
            return
        label = SEND_MODE_LABELS.get(mode, mode.title())
        batches = self.grouped_batches(rows, mode)
        if QMessageBox.question(self, "Send pending", f"Send {len(rows)} pending variance rows in {len(batches)} image batch(es) by {label}?\n\nImages will route to district WhatsApp groups.") != QMessageBox.Yes:
            return
        self._send_rows_thread(rows, mode=mode, manual=False)

    def send_checked_status(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        rows = [self.status_row_by_key[k] for k in self.status_checked_keys if k in self.status_row_by_key]
        if not rows:
            QMessageBox.information(self, "Nothing checked", "Select a district or store so Inventory Audit Status rows get checked, then send them.")
            return
        mode = self.current_status_mode_key()
        batches = self.grouped_status_batches(rows, mode)
        if QMessageBox.question(self, "Send Inventory Audit Status", f"Send {len(rows)} inventory status row(s) in {len(batches)} image batch(es)?") != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_status_rows, args=(rows, mode), daemon=True)
        thread.start()

    def _send_rows(self, rows: List[VarianceRow], mode: str, manual: bool) -> None:
        try:
            renderer = ImageRenderer()
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", "web") or "web")
            batches = self.grouped_batches(rows, mode)
            send_mode_for_file = mode
            for batch_title, district, batch_rows in batches:
                group_name = group_name_for_district(district, self.db)
                for chunk_no, chunk_rows in enumerate(self.chunks(batch_rows, AUTO_SEND_CHUNK_SIZE), start=1):
                    title = batch_title if len(batch_rows) <= AUTO_SEND_CHUNK_SIZE else f"{batch_title} | Part {chunk_no}"
                    image_path = renderer.render_rows(title, chunk_rows, mode=send_mode_for_file)
                    try:
                        self.set_status(f"Sending {len(chunk_rows)} variance row(s) to {group_name}. {title}.")
                        _win_state = self._save_window_state()
                        sender.send_image(group_name, image_path, text_message=self.variance_request_message(chunk_rows))
                        self._restore_window_state(_win_state)
                        self.db.mark_sent(chunk_rows, group_name, title, str(image_path), mode=send_mode_for_file)
                    except Exception as exc:
                        self.db.mark_sent(chunk_rows, group_name, title, str(image_path), mode=send_mode_for_file, error=str(exc))
                        self._worker_dialog("error", "WhatsApp send failed", f"Image was created but WhatsApp send failed.\n\nGroup: {group_name}\nBatch: {title}\nImage: {image_path}\n\nError: {exc}")
                        self.set_status(f"Send failed. Image saved: {image_path}")
                        return
            self._bridge.refresh_audit.emit()
            self.set_status("Variance image sending completed.")
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status("Send failed.")

    def _send_status_rows(self, rows: List[InventoryStatusRow], mode: str) -> None:
        def _status(msg):
            # set_status itself is thread-safe in the Qt build (UIBridge emit)
            self.set_status(msg)
        try:
            sender = WhatsAppSender(status_callback=lambda m: _status(m),
                                    mode=getattr(self, "wa_mode_var", "web") or "web")
            batches = self.grouped_status_batches(rows, mode)
            for batch_title, district, batch_rows in batches:
                group_name = group_name_for_district(district, self.db)
                for chunk_no, chunk_rows in enumerate(self.chunks(batch_rows, 28), start=1):
                    title = batch_title if len(batch_rows) <= 28 else f"{batch_title} | Part {chunk_no}"
                    image_path = self._render_status_rows(title, chunk_rows, mode="inventory status")
                    _status(f"Sending Inventory Audit Status with {len(chunk_rows)} row(s) to {group_name}. {title}.")
                    _win_state = self._save_window_state()
                    sender.send_image(group_name, image_path, text_message=self.inventory_status_message())
                    self._restore_window_state(_win_state)
                    self.db.mark_status_sent(chunk_rows)
                    pending_message = self.pending_inventory_count_message(chunk_rows)
                    if pending_message:
                        _status(f"Sending incomplete inventory count reminder to {group_name}.")
                        _win_state2 = self._save_window_state()
                        sender.send_text(group_name, pending_message)
                        self._restore_window_state(_win_state2)
            self._bridge.refresh_status.emit()
            _status("Inventory Audit Status image sending completed.")
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            _status("Inventory Audit Status send failed.")

    def _send_starting_message_thread(self, districts: List[str]) -> None:
        try:
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", "web") or "web")
            message = "Please complete an Inventory count in 15 minutes."
            for district in districts:
                group_name = group_name_for_district(district, self.db)
                self.set_status(f"Sending starting message to {group_name}...")
                _win_state = self._save_window_state()
                sender.send_text(group_name, message)
                self._restore_window_state(_win_state)
            self.set_status("Starting message sent to all selected districts.")
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status(f"Starting message send failed: {exc}")

    def _send_single_reminder_thread(self, districts: List[str], uncleared_rows: List[VarianceRow], reminder_number: int, message: str) -> None:
        try:
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", "web") or "web")
            for district in districts:
                district_rows = [row for row in uncleared_rows if normalize_district(row.district) == normalize_district(district)]
                if not district_rows:
                    continue
                group_name = group_name_for_district(district, self.db)
                rep_mentions = self.variance_request_message(district_rows)
                full_message = f"Reminder {reminder_number}/3: {message}"
                if rep_mentions:
                    full_message = f"{rep_mentions} {message}"
                self.set_status(f"Sending reminder {reminder_number}/3 to {group_name}...")
                _win_state = self._save_window_state()
                sender.send_text(group_name, full_message)
                self._restore_window_state(_win_state)
                time.sleep(2)
            self.set_status(f"Reminder {reminder_number} sent successfully.")
            try:
                self._worker_dialog("info", "Reminder sent", f"Reminder {reminder_number} has been sent to the selected district group(s).")
            except Exception:
                pass
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status(f"Reminder {reminder_number} send failed: {exc}")

    def _send_final_district_result_thread(self, districts: List[str]) -> None:
        final_results: List[Dict[str, str]] = []
        try:
            sender = WhatsAppSender(self.set_status, mode=getattr(self, "wa_mode_var", "web") or "web")
            renderer = ImageRenderer()
            all_rows = self.db.get_rows_by_keys(self.loaded_keys)
            self.set_status("Sending final district audit results...")
            for district in districts:
                district_rows = [row for row in all_rows if normalize_district(row.district) == normalize_district(district)]
                district_rows = [row for row in district_rows if not row.cleared]
                district_rows = self.filter_excluded_variance_rows(district_rows)
                group_name = group_name_for_district(district, self.db)
                pending_status_rows = [
                    row for row in self.status_rows
                    if normalize_district(row.district) == normalize_district(district)
                    and normalize_header(row.status) != "completed"
                ]
                caption = self.final_district_caption(district, district_rows, pending_status_rows)
                if district_rows:
                    image_path = renderer.render_rows(district, district_rows, mode="district")
                    status = "Variance Found"
                else:
                    image_path = self._render_no_variance_image(district)
                    status = "No Variance Found"
                self.set_status(f"Sending final district result to {group_name}.")
                _win_state = self._save_window_state()
                sender.send_image(group_name, image_path, text_message=caption)
                self._restore_window_state(_win_state)
                final_results.append({
                    "district": district,
                    "stores": self.format_store_list_caption(sorted({safe_text(r.store) for r in district_rows if safe_text(r.store)})),
                    "pending_inventory_stores": self.format_store_list_caption(sorted({safe_text(r.store) for r in pending_status_rows if safe_text(r.store)})),
                    "status": status,
                    "caption": caption,
                    "image_path": str(image_path),
                    "group_name": group_name,
                    "sent_at": now_text(),
                })

            log_path = self.create_full_inventory_audit_log(final_results)
            self.set_status(f"Final district audit result sending completed. Full audit log created: {log_path}")
            try:
                self._worker_dialog("info", "Audit log created", f"Full inventory audit log created:\n{log_path}")
            except Exception:
                pass
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status(f"Final district sending failed: {exc}")

    def export_log(self) -> None:
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Variance_Clearance_Log_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Clearance Log", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return
        keys = self.loaded_keys if self.data_loaded else None
        self.db.export_xlsx(Path(path), keys=keys)
        self.set_status(f"Exported clearance log: {path}")
        QMessageBox.information(self, "Export complete", f"Log exported:\n{path}")

    # ── Cross-instance DB sync ───────────────────────────────────────────
    def _start_db_sync_poll(self) -> None:
        """Begin the 2-second mtime-polling loop (Tk after() -> single-shot
        QTimer re-armed by _check_db_sync)."""
        self._schedule_db_sync()

    def _schedule_db_sync(self) -> None:
        if getattr(self, "_db_sync_timer", None) is None:
            self._db_sync_timer = QTimer(self)
            self._db_sync_timer.setSingleShot(True)
            self._db_sync_timer.timeout.connect(self._check_db_sync)
        self._db_sync_timer.start(2000)

    def closeEvent(self, event) -> None:
        """Cancel background loops then close (Tk _on_close port)."""
        try:
            timer = getattr(self, "_db_sync_timer", None)
            if timer is not None:
                timer.stop()
            if self._db_sync_after_id is not None:
                self._db_sync_after_id = None
        except Exception:
            pass
        try:
            self._wa_ocr_running = False
        except Exception:
            pass
        try:
            self._scheduler.stop()
        except Exception:
            pass
        event.accept()

    def send_starting_message(self) -> None:
        """Send 'Please complete an Inventory count in 15 minutes.' to district WhatsApp groups."""
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        all_districts = sorted({normalize_district(row.district) for row in self.status_rows if safe_text(row.district)})
        if not all_districts:
            QMessageBox.information(self, "No districts", "No district data is available.")
            return
        selected = self.final_district_combo.currentText().strip()
        if selected and selected != "All Districts":
            districts = [normalize_district(selected)]
        else:
            districts = all_districts
        if QMessageBox.question(self, "Send starting message", f"Send starting message to {len(districts)} district group(s)?\n\nMessage: Please complete an Inventory count in 15 minutes.") != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_starting_message_thread, args=(districts,), daemon=True)
        thread.start()

    def send_reminder(self, reminder_number: int) -> None:
        """Send a single reminder message for uncleared variances to district WhatsApp groups."""
        reminder_messages = {
            1: "Please clear the pending variances.",
            2: "This is the second reminder. Please clear the pending variances immediately.",
            3: "Final reminder. Uncleared variances will be reported.",
        }
        if reminder_number not in reminder_messages:
            return
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        all_rows = self.db.get_rows_by_keys(self.loaded_keys)
        uncleared_rows = [row for row in all_rows if not row.cleared]
        uncleared_rows = self.filter_excluded_variance_rows(uncleared_rows)
        if not uncleared_rows:
            QMessageBox.information(self, "No uncleared variances", "All variances have been cleared. No reminder needed.")
            return
        all_districts = sorted({normalize_district(row.district) for row in uncleared_rows if safe_text(row.district)})
        if not all_districts:
            QMessageBox.information(self, "No districts", "No district data available for reminders.")
            return
        selected = self.final_district_combo.currentText().strip()
        if selected and selected != "All Districts":
            districts = [normalize_district(selected)]
        else:
            districts = all_districts
        message = reminder_messages[reminder_number]
        if QMessageBox.question(
                self,
                f"Send Reminder {reminder_number}",
                f"Send reminder {reminder_number} to {len(districts)} district group(s)?\n\n"
                f"Message: {message}\n\n"
                f"Total {len(uncleared_rows)} uncleared variance row(s).",
        ) != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_single_reminder_thread, args=(districts, uncleared_rows, reminder_number, message), daemon=True)
        thread.start()

    def send_final_district_result(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        all_districts = sorted({normalize_district(row.district) for row in self.status_rows if safe_text(row.district)})
        if not all_districts:
            all_districts = sorted({normalize_district(row.district) for row in self.db.get_rows_by_keys(self.loaded_keys) if safe_text(row.district)})
        if not all_districts:
            QMessageBox.information(self, "No districts", "No district data is available for final sending.")
            return
        selected = self.final_district_combo.currentText().strip()
        if selected and selected != "All Districts":
            districts = [normalize_district(selected)]
        else:
            districts = all_districts
        if QMessageBox.question(self, "Send final district result", f"Send final audit result to {len(districts)} district group(s)?") != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_final_district_result_thread, args=(districts,), daemon=True)
        thread.start()

    def _file_row(self, grid, row: int, label: str, edit, command) -> None:
        """File-chooser row helper (dead helper in the Tk build too — kept
        for parity; the scheduler fills the file paths programmatically)."""
        grid.addWidget(QLabel(label), row, 0)
        grid.addWidget(edit, row, 1)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(command)
        grid.addWidget(browse_btn, row, 2)


def _enable_dpi_awareness() -> None:
    """AppUserModelID + DPI awareness (Windows).

    Windows groups windows in the taskbar by process/AppUserModelID and
    needs physical-pixel reporting on high-DPI displays; without the
    shcore call the UI would render blurry on DPI-scaled laptops."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GFH.AuditAutomation")
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _vp_hide(win):
    """ANTI-FLASH: keep the window fully transparent while the UI is built, so a
    half-built window never flashes behind the boot splash."""
    try:
        win.setWindowOpacity(0.0)
    except Exception:
        pass
    try:
        import pyi_splash  # type: ignore
        pyi_splash.update_text("Loading interface...")
    except Exception:
        pass


def _vp_reveal(win):
    """Show the finished window in one step (called just before the splash closes)."""
    try:
        win.setWindowOpacity(1.0)
        win.raise_()
        win.activateWindow()
    except Exception:
        pass


def _close_pyi_splash():
    """Close the PyInstaller splash screen (frozen onefile builds only).

    Called once the main window is fully built and shown; harmless no-op
    in dev runs (pyi_splash only exists inside PyInstaller builds)."""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main() -> None:
    _enable_dpi_awareness()
    app = QApplication(sys.argv)
    _check_required_packages(app)
    win = MainWindow()
    win.show()
    _vp_reveal(win)
    _close_pyi_splash()
    sys.exit(app.exec())


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    try:
        main()
    except Exception as exc:
        traceback.print_exc()
        show_startup_error(exc)
        if sys.stdin and sys.stdin.isatty():
            input("Press Enter to close...")
