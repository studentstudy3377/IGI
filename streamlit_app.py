
from __future__ import annotations

import csv
import html
import io
import json
import numbers
import os
import re
import shutil
import sys
import threading
import time
import zlib
from dataclasses import asdict, dataclass, field
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional, Tuple
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

# -----------------------------------------------------------------------------
# Settings
# -----------------------------------------------------------------------------
API_URL = os.environ.get(
    "IGI_API_URL",
    "https://api.igi.org/ReportDetail.php?Printno={report_no}",
)
DIAGNOSIS_URL = os.environ.get(
    "IGI_DIAGNOSIS_URL",
    "https://www.igi.org/API-IGI/report-diagnosis.php?r={report_no}",
)
ENDPOINT_NAMES = {"api": "IGI API", "page": "report-diagnosis page"}
# The ways to reach IGI, tried in this order; the one that answers is tried first for the next report.
ROUTES = {
    "Auto": [("direct", "api"), ("direct", "page"), ("chrome", "page"), ("chrome", "api")],
    "Chrome": [("chrome", "page"), ("chrome", "api")],
    "Direct": [("direct", "api"), ("direct", "page")],
}
ON_WINDOWS = sys.platform == "win32"
# "uc": visible Chrome via undetected-chromedriver (Windows PC).
# "headless": hidden Chromium via Selenium (Streamlit Cloud / any Linux server).
BROWSER_KIND = os.environ.get("IGI_BROWSER") or ("uc" if ON_WINDOWS else "headless")
CHROMIUM_CANDIDATES = ["/usr/bin/chromium", "/usr/bin/chromium-browser", "chromium", "chromium-browser",
                       "google-chrome", "google-chrome-stable"]          # env CHROME_BINARY overrides
CHROMEDRIVER_CANDIDATES = ["/usr/bin/chromedriver", "/usr/lib/chromium/chromedriver", "chromedriver"]  # env CHROMEDRIVER

STREAMLIT_IPS_URL = "https://docs.streamlit.io/deploy/streamlit-community-cloud/status"
SAVE_FILE_NAME = "IGI Data {date}.xlsx"   # {date} = dd-mm-yyyy; an existing file is never overwritten
CHROME_PAGE_TIMEOUT = 60                   # seconds per report, including a Cloudflare check

# The app has no settings panel: these are fixed here.
READ_METHOD = "Auto"                       # "Auto", "Chrome" or "Direct" (see ROUTES above)
SHOW_CHROME = True                         # Windows PC: show the Chrome window (a hidden one is stopped by Cloudflare more often)
REPORT_NO_AS_IGI = False                   # False: Report No. as you typed it (852601380); True: as IGI prints it (LG852601380)
DATE_FORMAT = "dd/mm/yyyy"                 # Report Date: one of the names in DATE_FORMATS below
SKIP_MISSING = False                       # True: leave out reports IGI did not return
SAVE_FOLDER = r"D:\Gaurav\Download"        # Windows PC: where "Save Excel" writes

# Your sheet, column for column ("Culet  Size" keeps the two spaces of your header).
OUTPUT_COLUMNS = [
    "Lot no", "Document No / Name", "Report Date", "Report No.", "Shape", "Weight",
    "Clarity", "Col", "P/ct", "$ Rate", "Cost P/Ct", "MM 1", "MM 2", "MM 3", "Table",
    "Crown Height", "Pavilion Depth", "Total Depth", "Crown Angle", "Pavilion Angle",
    "Culet  Size", "Girdle Percent", "Girdle Name", "CUT-PROP", "POL or Pol/Sym", "SYM",
    "Flo", "Ratio",
]
NUMBER_COLUMNS = {
    "Weight", "P/ct", "$ Rate", "Cost P/Ct", "MM 1", "MM 2", "MM 3", "Table",
    "Crown Height", "Pavilion Depth", "Total Depth", "Crown Angle", "Pavilion Angle",
    "Girdle Percent", "Ratio",
}

# IGI "SHAPE AND CUT" -> your shape code; the first word found wins.
# Only OV comes from your sample row. Change the others to the codes your sheet uses.
# A shape not listed is written as IGI names it (for example "Old Mine Brilliant").
SHAPE_CODES = [
    ("OVAL", "OV"),
    ("PEAR", "PS"),
    ("MARQUISE", "MQ"),
    ("HEART", "HT"),
    ("PRINCESS", "PR"),
    ("SQUARE MODIFIED BRILLIANT", "PR"),
    ("CUSHION", "CU"),
    ("ASSCHER", "AS"),
    ("EMERALD", "EM"),
    ("RADIANT", "RAD"),
    ("CUT-CORNERED", "RAD"),
    ("CUT CORNERED", "RAD"),
    ("BAGUETTE", "BG"),
    ("TRILLIANT", "TR"),
    ("TRIANGULAR", "TR"),
    ("ROUND", "RD"),
]

# Report Date: shown/pasted as text in this format; the Excel file gets a real date.
DATE_FORMATS: Dict[str, Tuple[Optional[str], Optional[str]]] = {
    "dd/mm/yyyy": ("%d/%m/%Y", "dd/mm/yyyy"),
    "dd-mm-yyyy": ("%d-%m-%Y", "dd-mm-yyyy"),
    "dd-Mmm-yyyy": ("%d-%b-%Y", "dd-mmm-yyyy"),
    "As IGI writes it": (None, None),
}

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# -----------------------------------------------------------------------------
# Small text helpers
# -----------------------------------------------------------------------------
_NUM = r"\d+(?:\.\d+)?"


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def clean_text(value: Any) -> str:
    """One line of plain text: no <br>, no HTML entities, single spaces."""
    if is_blank(value):
        return ""
    text = html.unescape(str(value))
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
    text = re.sub(r"</?[A-Za-z][^>]*>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def short_error(exc: BaseException) -> str:
    text = str(exc).strip().splitlines()
    first = text[0] if text else exc.__class__.__name__
    return first[:200]


def _norm(text: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def title_words(text: str) -> str:
    """'MEDIUM TO SLIGHTLY THICK (FACETED)' -> 'Medium To Slightly Thick (Faceted)'."""
    text = clean_text(text)
    return text.title() if text else ""


def to_number(text: str) -> Optional[float]:
    match = re.search(_NUM, text or "")
    return float(match.group(0)) if match else None


def measurements(text: str) -> List[Optional[float]]:
    """'8.50 x 5.62 x 3.38 mm' -> [8.5, 5.62, 3.38]; '3.91 - 3.94 x 2.42 mm' -> [3.91, 3.94, 2.42]."""
    values: List[Optional[float]] = [float(n) for n in re.findall(_NUM, text or "")][:3]
    return values + [None] * (3 - len(values))


def percent_and_angle(text: str) -> Tuple[Optional[float], Optional[float]]:
    """'12.5% - 32.4°' -> (12.5, 32.4); '43.5%' -> (43.5, None)."""
    text = text or ""
    nums = [float(n) for n in re.findall(_NUM, text)]
    pct = re.search(rf"({_NUM})\s*%", text)
    ang = re.search(rf"({_NUM})\s*(?:°|º|˚|deg)", text, flags=re.I)
    percent = float(pct.group(1)) if pct else None
    angle = float(ang.group(1)) if ang else None
    rest = list(nums)
    for value in (percent, angle):
        if value is not None and value in rest:
            rest.remove(value)
    if percent is None and rest:
        percent = rest.pop(0)
    if angle is None and rest:
        angle = rest.pop(0)
    return percent, angle


def girdle_name_and_percent(text: str) -> Tuple[str, Optional[float]]:
    """'MEDIUM TO SLIGHTLY THICK (FACETED)' -> ('Medium To Slightly Thick(Faceted)', None)."""
    text = clean_text(text)
    pct = re.search(rf"({_NUM})\s*%", text)
    percent = float(pct.group(1)) if pct else None
    name = re.sub(rf"[,;]?\s*\(?\s*{_NUM}\s*%\s*\)?", "", text).strip(" ,;-")
    name = re.sub(r"\s+\(", "(", title_words(name))
    return name, percent


def shape_code(text: str) -> str:
    upper = clean_text(text).upper()
    if not upper:
        return ""
    for word, code in SHAPE_CODES:
        if re.search(rf"(?<![A-Z]){re.escape(word)}(?![A-Z])", upper):
            return code
    return title_words(upper)


def parse_report_date(text: str) -> Optional[datetime]:
    text = clean_text(text)
    for fmt in ("%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_report_date(text: str, date_fmt: str) -> str:
    raw = clean_text(text)
    python_fmt = DATE_FORMATS.get(date_fmt, (None, None))[0]
    if not raw or python_fmt is None:
        return raw
    parsed = parse_report_date(raw)
    return parsed.strftime(python_fmt) if parsed else raw


def ratio(mm1: Optional[float], mm2: Optional[float]) -> Optional[float]:
    """Length / width: the longer of MM 1 and MM 2 divided by the shorter one."""
    if not mm1 or not mm2:
        return None
    return max(mm1, mm2) / min(mm1, mm2)


# -----------------------------------------------------------------------------
# IGI record -> one row of your sheet
# -----------------------------------------------------------------------------
FIELD_KEYS: Dict[str, List[str]] = {
    "report_no": ["REPORT NUMBER", "Report No", "Report #"],
    "date": ["REPORT DATE", "Issue Date", "Date of Issue"],
    "shape": ["SHAPE AND CUT", "Shape and Cutting Style", "Shape & Cut", "Shape"],
    "weight": ["CARAT WEIGHT", "Weight"],
    "color": ["COLOR GRADE", "Colour Grade", "Color"],
    "clarity": ["CLARITY GRADE", "Clarity"],
    "cut": ["CUT GRADE", "Cut"],
    "polish": ["POLISH"],
    "symmetry": ["SYMMETRY"],
    "measurements": ["Measurements", "Measurement"],
    "table": ["Table Size", "Table"],
    "crown": ["Crown Height", "Crown Height - Angle"],
    "crown_angle": ["Crown Angle"],
    "pavilion": ["Pavilion Depth", "Pavillion Depth", "Pavilion Depth - Angle"],
    "pavilion_angle": ["Pavilion Angle", "Pavillion Angle"],
    "girdle": ["Girdle Thickness", "Girdle"],
    "girdle_pct": ["Girdle %", "Girdle Percentage", "Girdle Percent"],
    "culet": ["Culet"],
    "depth": ["Total Depth", "Depth"],
    "fluorescence": ["FLUORESCENCE"],
}


def rec_value(record: Dict[str, Any], name: str) -> str:
    """Value of an IGI field, matching the key without regard to case or spaces."""
    for alias in FIELD_KEYS[name]:
        value = record.get(_norm(alias))
        if not is_blank(value):
            return clean_text(value)
    return ""


def build_row(query: str, record: Dict[str, Any], igi_number: bool, date_fmt: str) -> Dict[str, Any]:
    rec = {_norm(k): v for k, v in (record or {}).items() if _norm(k)}
    row: Dict[str, Any] = {column: None for column in OUTPUT_COLUMNS}
    row["Report No."] = (rec_value(rec, "report_no") if igi_number else "") or query
    if not rec:
        return row

    mm1, mm2, mm3 = measurements(rec_value(rec, "measurements"))
    crown_height, crown_angle = percent_and_angle(rec_value(rec, "crown"))
    pavilion_depth, pavilion_angle = percent_and_angle(rec_value(rec, "pavilion"))
    girdle_name, girdle_pct = girdle_name_and_percent(rec_value(rec, "girdle"))

    row.update({
        "Report Date": format_report_date(rec_value(rec, "date"), date_fmt) or None,
        "Shape": shape_code(rec_value(rec, "shape")) or None,
        "Weight": to_number(rec_value(rec, "weight")),
        "Clarity": re.sub(r"\s+", "", rec_value(rec, "clarity")).upper() or None,
        "Col": rec_value(rec, "color") or None,
        "MM 1": mm1,
        "MM 2": mm2,
        "MM 3": mm3,
        "Table": to_number(rec_value(rec, "table")),
        "Crown Height": crown_height,
        "Pavilion Depth": pavilion_depth,
        "Total Depth": to_number(rec_value(rec, "depth")),
        "Crown Angle": crown_angle if crown_angle is not None else to_number(rec_value(rec, "crown_angle")),
        "Pavilion Angle": pavilion_angle if pavilion_angle is not None else to_number(rec_value(rec, "pavilion_angle")),
        "Culet  Size": title_words(rec_value(rec, "culet")) or None,
        "Girdle Percent": girdle_pct if girdle_pct is not None else to_number(rec_value(rec, "girdle_pct")),
        "Girdle Name": girdle_name or None,
        "CUT-PROP": title_words(rec_value(rec, "cut")) or None,
        "POL or Pol/Sym": title_words(rec_value(rec, "polish")) or None,
        "SYM": title_words(rec_value(rec, "symmetry")) or None,
        "Flo": title_words(rec_value(rec, "fluorescence")) or None,
        "Ratio": ratio(mm1, mm2),
    })
    return row


# -----------------------------------------------------------------------------
# Reading IGI's answer (API JSON or the report-diagnosis page)
# -----------------------------------------------------------------------------
@dataclass
class FetchResult:
    query: str
    status: str                       # "OK", "Not found" or "Error"
    record: Dict[str, Any] = field(default_factory=dict)
    message: str = ""
    source: str = ""                  # "direct" or "chrome"
    endpoint: str = ""                # "api" or "page"


def _json_body(text: str) -> str:
    """The text without Chrome's JSON-viewer 'Pretty-print' label in front."""
    return re.sub(r"^\s*Pretty-print\s*", "", text or "")


def page_kind(text: str) -> str:
    text = text or ""
    if "IGI Report Diagnosis" in text:
        return "diagnosis"
    body = _json_body(text).strip()
    if body[:1] in ("[", "{") or re.fullmatch(r'null|false|true|"[^"<>]{0,200}"', body):
        return "json"                     # the API's answer (null/false/"..." for an unknown number)
    low = text.lower()
    if "you have been blocked" in low or "attention required" in low:
        return "blocked"
    if any(s in low for s in ("just a moment", "verify you are human", "checking your browser",
                              "performing security verification")):
        return "challenge"
    if "err_" in low or "this site can" in low:
        return "browser_error"
    if not body:
        return "empty"
    return "unknown"


def is_report_record(record: Any) -> bool:
    if not isinstance(record, dict):
        return False
    keys = {_norm(k) for k in record}
    return bool(keys & {"reportnumber", "caratweight", "shapeandcut"})


def read_api_json(text: str) -> List[Dict[str, Any]]:
    """Report records from the API's JSON ([{...}], one {...}, or nothing for an unknown number)."""
    body = _json_body(text).strip()
    try:
        value, _ = json.JSONDecoder(strict=False).raw_decode(body)
    except ValueError:
        return []
    if isinstance(value, dict):
        value = [value]
    return [r for r in value if is_report_record(r)] if isinstance(value, list) else []


def _json_list_after(text: str, heading: str, stop: str = "") -> Optional[List[Any]]:
    """The JSON array that follows 'data:' under a heading of the diagnosis page."""
    start = text.find(heading)
    if start < 0:
        return None
    data_at = text.find("data:", start)
    end = text.find(stop, start + len(heading)) if stop else -1
    if data_at < 0 or (end >= 0 and data_at > end):
        return None
    bracket = text.find("[", data_at)
    if bracket < 0 or re.sub(r"<br\s*/?>|\s", "", text[data_at + 5:bracket], flags=re.I):
        return None
    try:
        value, _ = json.JSONDecoder(strict=False).raw_decode(text, bracket)
    except ValueError:
        return None
    return value if isinstance(value, list) else None


def read_diagnosis_page(text: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """(loose report records, jewellery/SWIFT records) from the page HTML or its visible text."""
    loose = _json_list_after(text, "LOOSE REPORT DATA", stop="SWIFT/JEWELRY DATA") or []
    swift = _json_list_after(text, "SWIFT/JEWELRY DATA") or []
    return [r for r in loose if isinstance(r, dict)], [r for r in swift if isinstance(r, dict)]


def merge_records(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    merged: Dict[str, Any] = {}
    for record in records:
        for key, value in record.items():
            if key not in merged or is_blank(merged[key]):
                merged[key] = value
    return merged


def interpret_page(query: str, text: str, source: str, endpoint: str = "page") -> FetchResult:
    kind = page_kind(text)

    def result(status: str, record: Optional[Dict[str, Any]] = None, message: str = "") -> FetchResult:
        return FetchResult(query, status, record or {}, message, source, endpoint)

    if kind == "blocked":
        return result("Error", message="blocked by IGI's Cloudflare")
    if kind == "challenge":
        return result("Error", message="stopped at IGI's Cloudflare check (Verify you are human)")
    if kind == "json":
        records = read_api_json(text)
        if records:
            return result("OK", merge_records(records))
        return result("Not found", message="IGI has no report with this number.")
    if kind == "empty":
        return result("Error", message="empty answer from IGI")
    if kind != "diagnosis":
        return result("Error", message="the IGI page did not open")
    loose, swift = read_diagnosis_page(text)
    if loose:
        return result("OK", merge_records(loose))
    if swift:
        return result("OK", merge_records(swift), "Taken from IGI's jewellery (SWIFT) data.")
    return result("Not found", message="IGI has no report with this number.")


def find_program(env_name: str, candidates: List[str]) -> Optional[str]:
    """A program path: the environment variable if set, else the first candidate that exists."""
    path = os.environ.get(env_name, "").strip()
    if path:
        return path
    for candidate in candidates:
        found = candidate if os.path.isabs(candidate) and os.path.exists(candidate) else shutil.which(candidate)
        if found:
            return found
    return None


class ChromePage:
    """One browser, reused for a whole batch.

    kind "uc": visible Chrome via undetected-chromedriver (Windows PC, as in the IGI report script).
    kind "headless": hidden Chromium via Selenium (Streamlit Cloud; packages.txt installs it).
    """

    def __init__(self, show_window: bool = True, kind: str = BROWSER_KIND):
        self.kind = kind
        self.show_window = show_window and kind == "uc"
        self.driver = None
        self.extra_args: List[str] = []           # additional Chrome switches, if ever needed
        self.extra_kwargs: Dict[str, Any] = {}    # e.g. driver_executable_path=... (uc only)

    def start(self) -> None:
        if self.kind == "uc":
            self._start_uc()
        else:
            self._start_headless()
        self.driver.set_page_load_timeout(CHROME_PAGE_TIMEOUT)

    def _start_headless(self) -> None:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service

        options = webdriver.ChromeOptions()
        browser = find_program("CHROME_BINARY", CHROMIUM_CANDIDATES)
        if browser:
            options.binary_location = browser
        for arg in ("--headless=new", "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
                    "--window-size=1100,850", *self.extra_args):
            options.add_argument(arg)
        driver_path = find_program("CHROMEDRIVER", CHROMEDRIVER_CANDIDATES)
        service = Service(executable_path=driver_path) if driver_path else Service()
        self.driver = webdriver.Chrome(service=service, options=options)

    def _options(self):
        import undetected_chromedriver as uc

        options = uc.ChromeOptions()
        if not self.show_window:
            options.add_argument("--headless=new")
        options.add_argument("--window-size=1100,850")
        options.add_argument("--disable-features=IsolateOrigins,site-per-process")
        options.add_argument("--disable-blink-features=AutomationControlled")
        for arg in self.extra_args:
            options.add_argument(arg)
        return options

    def _start_uc(self) -> None:
        import undetected_chromedriver as uc
        from selenium.common.exceptions import SessionNotCreatedException

        if not getattr(uc.Chrome, "_quiet_del", False):
            original_del = uc.Chrome.__del__

            def quiet_del(chrome):          # Windows: ignore "handle is invalid" on exit
                try:
                    original_del(chrome)
                except Exception:
                    pass

            uc.Chrome.__del__ = quiet_del
            uc.Chrome._quiet_del = True

        try:
            self.driver = uc.Chrome(options=self._options(), **self.extra_kwargs)
        except SessionNotCreatedException as exc:
            # ChromeDriver and Chrome versions differ: retry with Chrome's own major version.
            match = re.search(r"Current browser version is (\d+)\.", str(exc))
            if not match:
                raise
            self.driver = uc.Chrome(
                options=self._options(), version_main=int(match.group(1)), **self.extra_kwargs
            )

    def read(self, url: str, timeout: float = CHROME_PAGE_TIMEOUT,
             on_check: Optional[Callable[[], None]] = None) -> str:
        """Open the URL and return its visible text once IGI's answer (or a block page) is there."""
        if self.driver is None:
            self.start()
        self.driver.get(url)
        deadline = time.monotonic() + timeout
        told = False
        text = ""
        while True:
            try:
                state, text = self.driver.execute_script(
                    "return [document.readyState, document.body ? document.body.innerText : ''];"
                )
                text = text or ""
            except Exception:              # the page is changing (Cloudflare redirect)
                state, text = "loading", text
            kind = page_kind(text)
            if (kind in ("diagnosis", "json", "empty") and state == "complete") or kind in ("blocked", "browser_error"):
                return text
            if kind == "challenge" and not told and on_check is not None:
                on_check()
                told = True
            if time.monotonic() >= deadline:
                return text
            time.sleep(0.5)

    def close(self) -> None:
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None


def route_label(how: str, endpoint: str) -> str:
    way = {"direct": "direct request", "browser": "your browser"}.get(how, "Chrome")
    return f"{ENDPOINT_NAMES.get(endpoint, endpoint)} ({way})"


class IGIFetcher:
    """Reads report numbers one by one.

    It tries the ways to reach IGI in ROUTES order (Auto: API by direct request, the
    report-diagnosis page by direct request, the page in Chrome, the API in Chrome). A way that
    IGI's Cloudflare refuses is dropped for the rest of the batch; the way that answered is tried
    first for the next report.
    """

    def __init__(self, method: str = "Auto", show_chrome: bool = True,
                 on_status: Optional[Callable[[str], None]] = None):
        self.method = method
        self.on_status = on_status or (lambda message: None)
        self.routes: List[Tuple[str, str]] = list(ROUTES.get(method, ROUTES["Auto"]))
        self.dropped: Dict[Tuple[str, str], str] = {}      # way -> why it is not used any more
        uses_chrome = any(how == "chrome" for how, _ in self.routes)
        self.chrome = ChromePage(show_chrome) if uses_chrome else None
        self.session = requests.Session()

    def _direct(self, url: str) -> Tuple[Optional[str], str]:
        try:
            response = self.session.get(url, headers=BROWSER_HEADERS, timeout=20)
        except requests.RequestException as exc:
            return None, f"the request failed ({short_error(exc)})"
        return response.content.decode("utf-8", errors="replace"), ""

    def _in_chrome(self, url: str, endpoint: str) -> Tuple[Optional[str], str]:
        if self.chrome.driver is None:
            self.on_status("Starting Chrome…")
            try:
                self.chrome.start()
            except Exception as exc:
                message = f"Could not start Chrome: {short_error(exc)}"
                for route in self.routes:
                    if route[0] == "chrome":
                        self.dropped[route] = message
                return None, message
        visible = self.chrome.show_window
        note = ("IGI is showing a Cloudflare check in the Chrome window. If it asks, tick the box there."
                if visible else "IGI is showing a Cloudflare check; waiting for it to pass…")
        # Nobody can tick a check in a hidden browser, so the API (which asks for one) gets less time.
        timeout = CHROME_PAGE_TIMEOUT if (visible or endpoint == "page") else 20
        try:
            return self.chrome.read(url, timeout=timeout, on_check=lambda: self.on_status(note)), ""
        except Exception as exc:
            self.chrome.close()                # a fresh window for the next report
            return None, f"Chrome: {short_error(exc)}"

    def fetch(self, query: str) -> FetchResult:
        last: Optional[FetchResult] = None
        for route in list(self.routes):
            if route in self.dropped:
                continue
            how, endpoint = route
            url = (API_URL if endpoint == "api" else DIAGNOSIS_URL).format(report_no=quote(query, safe=""))
            text, problem = self._direct(url) if how == "direct" else self._in_chrome(url, endpoint)
            if text is None:
                if how == "direct":
                    self.dropped[route] = problem
                last = FetchResult(query, "Error", message=problem, source=how, endpoint=endpoint)
                continue
            result = interpret_page(query, text, how, endpoint)
            if result.status in ("OK", "Not found"):
                self.routes.remove(route)
                self.routes.insert(0, route)   # the way that answered goes first next time
                return result
            if page_kind(text) in ("blocked", "challenge"):
                self.dropped[route] = result.message
            last = result
        reasons = {route_label(*r): m for r, m in self.dropped.items()}
        if last is not None and last.status == "Error":
            reasons.setdefault(route_label(last.source, last.endpoint), last.message)
        message = "IGI did not answer: " + "; ".join(f"{k}: {v}" for k, v in reasons.items()) + "."
        if self.method == "Direct":
            message += " Set READ_METHOD to Auto or Chrome at the top of the file."
        return FetchResult(query, "Error", message=message, source=last.source if last else "",
                           endpoint=last.endpoint if last else "")

    def close(self) -> None:
        if self.chrome is not None:
            self.chrome.close()


# -----------------------------------------------------------------------------
# Reading IGI in the user's own browser (the "IGI → sheet" bookmark)
#
# IGI refuses Streamlit Cloud's servers, but it serves the user's own browser. The bookmark runs on
# a www.igi.org tab, reads /API-IGI/report-diagnosis.php for each number from that same site (so no
# cross-site rules apply), and shows the records as JSON to copy and paste back into the app.
# The numbers come from the link the app opens (…#reports=n1,n2,…) or from a prompt.
# -----------------------------------------------------------------------------
BOOKMARKLET_JS = r"""
(async function () {
  var ID = "igi-sheet-box";
  var old = document.getElementById(ID);
  if (old) { old.remove(); }
  var box = document.createElement("div");
  box.id = ID;
  box.style.cssText = "position:fixed;top:16px;right:16px;z-index:2147483647;width:400px;max-width:90vw;background:#fff;color:#111;border:2px solid #b4232c;border-radius:10px;padding:14px;font:14px/1.45 system-ui,Segoe UI,Arial,sans-serif;box-shadow:0 8px 30px rgba(0,0,0,.25);text-align:left";
  var title = document.createElement("div");
  title.textContent = "IGI → sheet";
  title.style.cssText = "font-weight:700;margin-bottom:6px;color:#b4232c";
  var msg = document.createElement("div");
  box.appendChild(title);
  box.appendChild(msg);
  document.body.appendChild(box);
  function say(text) { msg.textContent = text; }
  function button(label, onClick) {
    var b = document.createElement("button");
    b.textContent = label;
    b.style.cssText = "margin:8px 8px 0 0;padding:7px 14px;border:0;border-radius:6px;background:#b4232c;color:#fff;font:600 13px system-ui,Segoe UI,Arial,sans-serif;cursor:pointer";
    b.onclick = onClick;
    box.appendChild(b);
    return b;
  }
  function closeButton() { button("Close", function () { box.remove(); }); }
  var m = location.hash.match(/reports=([^&]*)/);
  var raw = m ? decodeURIComponent(m[1]) : "";
  if (!raw) { raw = prompt("IGI report numbers (any separator):", "") || ""; }
  var nums = (raw.match(/(?:LG)?\d{6,}/gi) || []).map(function (s) { return s.toUpperCase(); });
  nums = nums.filter(function (n, i) { return nums.indexOf(n) === i; });
  if (!nums.length) { say("No report numbers given."); closeButton(); return; }
  function section(html, head, stop) {
    var a = html.indexOf(head);
    if (a < 0) { return []; }
    var end = stop ? html.indexOf(stop, a + head.length) : -1;
    var d = html.indexOf("data:", a);
    if (d < 0 || (end >= 0 && d > end)) { return []; }
    var s = html.indexOf("[", d);
    if (s < 0) { return []; }
    var depth = 0, inString = false, escaped = false;
    for (var i = s; i < html.length; i++) {
      var c = html.charAt(i);
      if (inString) {
        if (escaped) { escaped = false; } else if (c === "\\") { escaped = true; } else if (c === '"') { inString = false; }
      } else if (c === '"') { inString = true; }
      else if (c === "[") { depth++; }
      else if (c === "]") {
        depth--;
        if (depth === 0) {
          try {
            var value = JSON.parse(html.slice(s, i + 1).replace(/[\r\n]/g, " "));
            return Array.isArray(value) ? value.filter(function (r) { return r && typeof r === "object"; }) : [];
          } catch (e) { return []; }
        }
      }
    }
    return [];
  }
  async function readOne(n) {
    try {
      var r = await fetch("/API-IGI/report-diagnosis.php?r=" + encodeURIComponent(n), { credentials: "same-origin", cache: "no-store" });
      var html = await r.text();
      if (html.indexOf("IGI Report Diagnosis") < 0) {
        return { query: n, status: "Error", message: "IGI did not return its report page (HTTP " + r.status + ")", notIgi: true };
      }
      var records = section(html, "LOOSE REPORT DATA", "SWIFT/JEWELRY DATA");
      if (!records.length) { records = section(html, "SWIFT/JEWELRY DATA", ""); }
      return records.length ? { query: n, status: "OK", records: records } : { query: n, status: "Not found" };
    } catch (e) {
      return { query: n, status: "Error", message: String(e) };
    }
  }
  var out = new Array(nums.length);
  say("Reading " + nums[0] + " (1 of " + nums.length + ")…");
  out[0] = await readOne(nums[0]);
  if (out[0].notIgi) {
    say("IGI did not return its report page here. Open IGI with the link in the app, let the page load, then click this bookmark on that tab.");
    closeButton();
    return;
  }
  var next = 1, count = 1;
  async function worker() {
    while (next < nums.length) {
      var i = next;
      next++;
      out[i] = await readOne(nums[i]);
      count++;
      say("Read " + count + " of " + nums.length + "…");
    }
  }
  await Promise.all([worker(), worker(), worker()]);
  var ok = 0, missing = 0, failed = 0;
  out.forEach(function (item) {
    delete item.notIgi;
    if (item.status === "OK") { ok++; } else if (item.status === "Not found") { missing++; } else { failed++; }
  });
  var text = JSON.stringify({ igi_reports: out, read_at: new Date().toISOString() });
  say("Done: " + ok + " read" + (missing ? ", " + missing + " not found" : "") + (failed ? ", " + failed + " failed" : "") + ". Click Copy result, then paste it into the app (step 3).");
  var area = document.createElement("textarea");
  area.value = text;
  area.readOnly = true;
  area.style.cssText = "display:block;width:100%;height:80px;margin-top:8px;font:11px/1.3 monospace;box-sizing:border-box";
  box.appendChild(area);
  var copy = button("Copy result", async function () {
    try { await navigator.clipboard.writeText(text); }
    catch (e) { area.focus(); area.select(); document.execCommand("copy"); }
    copy.textContent = "Copied ✓";
  });
  closeButton();
})();
"""

BOOKMARK_BUTTON_HTML = """
<div style="font-family:system-ui,'Segoe UI',Arial,sans-serif;display:flex;flex-wrap:wrap;align-items:center;gap:8px 14px;padding:6px 2px">
  <a href="__URL__" title="IGI → sheet"
     onclick="document.getElementById('hint').textContent='Drag it (don\\'t click it) onto your bookmarks bar.';return false;"
     style="background:#b4232c;color:#fff;padding:9px 18px;border-radius:8px;text-decoration:none;font-weight:600;font-size:15px">IGI → sheet</a>
  <span id="hint" style="color:#444;font-size:14px">← drag this button onto your bookmarks bar</span>
</div>
"""


def bookmarklet_url() -> str:
    """The bookmark's address: javascript: plus the code above on one line, percent-encoded."""
    code = " ".join(line.strip() for line in BOOKMARKLET_JS.strip().splitlines() if line.strip())
    return "javascript:" + quote(code, safe="")


def browser_link(queries: List[str]) -> str:
    """The IGI page to open, carrying the report numbers for the bookmark after the #."""
    first = DIAGNOSIS_URL.format(report_no=quote(queries[0], safe=""))
    return first.split("#")[0] + "#reports=" + ",".join(quote(q, safe="") for q in queries)


BAD_PASTE = "That isn't the bookmark's result. On the IGI tab click Copy result, then paste here again."


def parse_browser_result(text: str) -> Tuple[Dict[str, Dict[str, Any]], str]:
    """Results keyed by report number from the bookmark's JSON, and a problem message if any."""
    text = (text or "").strip()
    start = text.find("{")
    if start < 0:
        return {}, BAD_PASTE
    try:
        data, _ = json.JSONDecoder(strict=False).raw_decode(text[start:])
    except ValueError:
        return {}, BAD_PASTE
    items = data.get("igi_reports") if isinstance(data, dict) else None
    if not isinstance(items, list):
        return {}, BAD_PASTE
    results: Dict[str, Dict[str, Any]] = {}
    for item in items:
        if not isinstance(item, dict) or not str(item.get("query") or "").strip():
            continue
        query = clean_report_cell(item["query"]) or str(item["query"]).strip().upper()
        records = [r for r in (item.get("records") or []) if is_report_record(r)]
        if item.get("status") == "OK" and records:
            result = FetchResult(query, "OK", merge_records(records), "", "browser", "page")
        elif item.get("status") == "Not found":
            result = FetchResult(query, "Not found", {}, "IGI has no report with this number.", "browser", "page")
        else:
            result = FetchResult(query, "Error", {}, str(item.get("message") or "not read"), "browser", "page")
        results[query] = asdict(result)
    return results, ("" if results else BAD_PASTE)


# -----------------------------------------------------------------------------
# Input: pasted numbers, pasted Excel rows with a header, or an Excel/CSV file
# -----------------------------------------------------------------------------
REPORT_HEADERS = [  # normalised header names that hold the report number, in order of preference
    "reportno", "reportnumber", "igireportno", "igireportnumber", "inputreportnumber",
    "reportnos", "report", "reportid", "certificateno", "certificatenumber", "certno",
    "certnumber", "certificate", "cert", "igino", "igi", "id",
]
OUTPUT_BY_KEY = {_norm(column): column for column in OUTPUT_COLUMNS}
_TOKEN = re.compile(r"(?<![A-Za-z0-9])(LG)?\s?(\d{6,})(?![0-9])", re.I)


def report_tokens(text: str) -> List[str]:
    """Report numbers in free text: 6+ digits, keeping an LG prefix (lab-grown)."""
    return [(prefix or "").upper() + digits for prefix, digits in _TOKEN.findall(text or "")]


def clean_report_cell(value: Any) -> str:
    text = re.sub(r"[\s#\"']", "", clean_text(value)).upper()
    text = re.sub(r"\.0+$", "", text)
    if text.startswith("IGI") and text[3:4].isalnum():
        text = text[3:]
    return text if re.fullmatch(r"[A-Z]{0,3}\d{6,}", text) else ""


def excel_shortened(text: str) -> List[str]:
    return re.findall(r"(?<![\w.])\d(?:\.\d+)?[eE]\+\d{1,2}\b", text or "")


def find_report_column(header: List[str]) -> Optional[int]:
    keys = [_norm(cell) for cell in header]
    for wanted in REPORT_HEADERS:
        if wanted in keys:
            return keys.index(wanted)
    return None


def rows_to_entries(header: List[str], rows: List[List[str]], report_idx: int,
                    notes: List[str]) -> List[Dict[str, Any]]:
    column_of = {
        i: OUTPUT_BY_KEY[_norm(cell)]
        for i, cell in enumerate(header)
        if i != report_idx and _norm(cell) in OUTPUT_BY_KEY and OUTPUT_BY_KEY[_norm(cell)] != "Report No."
    }
    entries: List[Dict[str, Any]] = []
    skipped: List[str] = []
    for cells in rows:
        raw = cells[report_idx] if report_idx < len(cells) else ""
        query = clean_report_cell(raw)
        if not query:
            if any(str(c).strip() for c in cells):
                skipped.append(str(raw).strip() or "(blank)")
            continue
        extras = {
            column: str(cells[i]).strip()
            for i, column in column_of.items()
            if i < len(cells) and str(cells[i]).strip()
        }
        entries.append({"query": query, "extras": extras})
    if skipped:
        shown = ", ".join(skipped[:10]) + (" …" if len(skipped) > 10 else "")
        notes.append(f"Skipped {len(skipped)} row(s) with no report number in the report column: {shown}")
    return entries


def _shortened_note(text: str) -> List[str]:
    found = excel_shortened(text)
    if not found:
        return []
    return [
        "These look like numbers Excel shortened, and the real report number is lost: "
        + ", ".join(dict.fromkeys(found))
        + ". Format the column as Number with 0 decimals in Excel and copy again."
    ]


def parse_pasted(text: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    lines = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    rows = [line.split("\t") for line in lines if line.strip()]
    if not rows:
        return [], []
    notes = _shortened_note(text)
    report_idx = find_report_column(rows[0])
    if report_idx is not None:            # rows copied from Excel with their header row
        return rows_to_entries(rows[0], rows[1:], report_idx, notes), notes
    queries = list(dict.fromkeys(report_tokens(text)))
    return [{"query": q, "extras": {}} for q in queries], notes


def parse_upload(uploaded) -> Tuple[List[Dict[str, Any]], List[str]]:
    name = getattr(uploaded, "name", "file")
    try:
        if hasattr(uploaded, "seek"):
            uploaded.seek(0)
        if name.lower().endswith(".csv"):
            text = uploaded.read().decode("utf-8-sig", errors="replace")
            try:
                sep = csv.Sniffer().sniff(text[:5000], delimiters=",;\t|").delimiter
            except csv.Error:
                sep = ","
            df = pd.read_csv(io.StringIO(text), dtype=str, sep=sep)
        else:
            df = pd.read_excel(uploaded, dtype=str)
    except Exception as exc:
        return [], [f"Could not read {name}: {short_error(exc)}"]
    df = df.fillna("")
    header = [str(c) for c in df.columns]
    rows = [[str(v) for v in row] for row in df.values.tolist()]
    text = "\n".join("\t".join(r) for r in [header] + rows)
    notes = _shortened_note(text)
    report_idx = find_report_column(header)
    if report_idx is not None:
        return rows_to_entries(header, rows, report_idx, notes), notes
    queries = list(dict.fromkeys(report_tokens(text)))
    notes.append(f"{name} has no Report No. / Report Number / ID column, so every cell was searched for report numbers.")
    return [{"query": q, "extras": {}} for q in queries], notes


def collect_entries(pasted: str, uploaded: Any = None) -> Tuple[List[Dict[str, Any]], List[str]]:
    entries, notes = parse_pasted(pasted)
    if uploaded is not None:
        more, more_notes = parse_upload(uploaded)
        known = {e["query"] for e in entries}
        entries += [e for e in more if e["query"] not in known]
        notes += more_notes
    return entries, notes


# -----------------------------------------------------------------------------
# Table and exports
# -----------------------------------------------------------------------------
def number_text(value: Any) -> str:
    number = float(value)
    if number != number or number in (float("inf"), float("-inf")):
        return ""
    return str(int(number)) if number.is_integer() else repr(number)


def cell_text(value: Any) -> str:
    """A cell as it is shown and pasted: 1.00 -> '1', 8.50 -> '8.5', blanks -> ''."""
    if is_blank(value):
        return ""
    if isinstance(value, numbers.Number) and not isinstance(value, bool):
        return number_text(value)
    return re.sub(r"[\t\r\n]+", " ", str(value)).strip()


def as_number(value: Any) -> Optional[float]:
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return None


def to_cell_value(column: str, value: Any) -> Any:
    """A value you pasted (Lot no, P/ct …): numbers become numbers, anything else stays as typed."""
    text = clean_text(value)
    if column in NUMBER_COLUMNS and as_number(text) is not None:
        return as_number(text)
    return text or None


def build_table(entries: List[Dict[str, Any]], results: Dict[str, Dict[str, Any]],
                igi_number: bool, date_fmt: str, skip_missing: bool) -> pd.DataFrame:
    """The 28-column table as text, exactly as it will be pasted (blank cells stay blank)."""
    rows = []
    for entry in entries:
        result = results.get(entry["query"]) or {}
        if skip_missing and result.get("status") != "OK":
            continue
        row = build_row(entry["query"], result.get("record") or {}, igi_number, date_fmt)
        for column, value in (entry.get("extras") or {}).items():
            if is_blank(row.get(column)):         # IGI's value wins; your pasted value fills the rest
                row[column] = to_cell_value(column, value)
        rows.append({column: cell_text(row[column]) for column in OUTPUT_COLUMNS})
    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS, dtype=object)


def table_rows_text(df: pd.DataFrame) -> List[List[str]]:
    return [[cell_text(row[column]) for column in OUTPUT_COLUMNS] for _, row in df.iterrows()]


def tsv_text(df: pd.DataFrame, header: bool) -> str:
    lines = (["\t".join(OUTPUT_COLUMNS)] if header else []) + ["\t".join(r) for r in table_rows_text(df)]
    return "\n".join(lines)


def csv_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\r\n")
    writer.writerow(OUTPUT_COLUMNS)
    writer.writerows(table_rows_text(df))
    return buffer.getvalue().encode("utf-8-sig")


def excel_bytes(df: pd.DataFrame, date_fmt: str) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    python_fmt, excel_fmt = DATE_FORMATS.get(date_fmt, (None, None))
    wb = Workbook()
    ws = wb.active
    ws.title = "IGI"
    ws.append(OUTPUT_COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for _, row in df.iterrows():
        values = []
        for column in OUTPUT_COLUMNS:
            value = row[column]
            if is_blank(value):
                values.append(None)
            elif column in NUMBER_COLUMNS:
                number = as_number(value)
                values.append(number if number is not None else str(value))
            elif column == "Report No." and str(value).isdigit() and not str(value).startswith("0"):
                values.append(int(value))
            elif column == "Report Date" and python_fmt:
                try:
                    values.append(datetime.strptime(str(value), python_fmt))
                except ValueError:
                    values.append(str(value))
            else:
                values.append(str(value))
        ws.append(values)

    date_col = OUTPUT_COLUMNS.index("Report Date") + 1
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=date_col)
        if isinstance(cell.value, datetime) and excel_fmt:
            cell.number_format = excel_fmt

    texts = table_rows_text(df)
    for i, column in enumerate(OUTPUT_COLUMNS):
        longest = max([len(column)] + [len(r[i]) for r in texts])
        ws.column_dimensions[get_column_letter(i + 1)].width = min(max(longest + 2, 6), 45)
    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def free_path(folder: str, file_name: str) -> str:
    stem, ext = os.path.splitext(file_name)
    path = os.path.join(folder, file_name)
    n = 1
    while os.path.exists(path):
        path = os.path.join(folder, f"{stem} ({n}){ext}")
        n += 1
    return path


def save_excel(data: bytes, folder: str, file_name: str) -> str:
    os.makedirs(folder, exist_ok=True)
    path = free_path(folder, file_name)
    with open(path, "wb") as handle:
        handle.write(data)
    return path


# -----------------------------------------------------------------------------
# Streamlit page
# -----------------------------------------------------------------------------
BLOCKED_HELP = (
    "**IGI's Cloudflare is refusing this server** (Streamlit Community Cloud runs on Google Cloud in the US). "
    "Use **Read in your browser** in the sidebar: your own browser reads IGI, which IGI accepts, and you paste "
    "the result there.\n"
    "- To read on the server instead, IGI would have to allow Streamlit Cloud's outgoing IP addresses "
    f"(ask igi-api-help@igi.org; the addresses are listed at {STREAMLIT_IPS_URL})."
)


@st.cache_resource
def fetch_lock() -> threading.Lock:
    """One IGI batch at a time on this server: one browser at a time keeps the app's memory low."""
    return threading.Lock()


def run_fetch(entries: List[Dict[str, Any]], method: str, show_chrome: bool) -> None:
    state = st.session_state
    results: Dict[str, Dict[str, Any]] = state["results"]
    state["entries"] = entries
    todo = [q for q in dict.fromkeys(e["query"] for e in entries) if results.get(q, {}).get("status") != "OK"]
    if not todo:
        st.info("These reports were read already. Use 'Clear fetched reports' in the sidebar to read them again.")
        return

    lock = fetch_lock()
    if not lock.acquire(blocking=False):
        waiting = st.info("Someone else is reading from IGI on this app; your list starts when theirs is done…")
        lock.acquire()
        waiting.empty()
    fetcher = None
    try:
        bar = st.progress(0.0, text="Starting…")
        note = st.empty()
        fetcher = IGIFetcher(method, show_chrome, on_status=lambda message: note.info(message))
        for i, query in enumerate(todo):
            bar.progress(i / len(todo), text=f"Reading {query}  ({i + 1} of {len(todo)})")
            results[query] = asdict(fetcher.fetch(query))
            note.empty()
        bar.progress(1.0, text=f"Done: {len(todo)} report(s) read.")
    finally:
        if fetcher is not None:
            fetcher.close()
        lock.release()


def column_config() -> Dict[str, Any]:
    return {column: st.column_config.TextColumn(column) for column in OUTPUT_COLUMNS}


def show_results(entries, results, igi_number, date_fmt, skip_missing) -> Optional[pd.DataFrame]:
    """Summary, the editable 28-column table and the copy box; returns the table as edited."""
    queries = list(dict.fromkeys(e["query"] for e in entries))
    status = {q: (results.get(q) or {}).get("status", "") for q in queries}
    found = [q for q in queries if status[q] == "OK"]
    missing = [q for q in queries if status[q] == "Not found"]
    failed = [q for q in queries if status[q] not in ("OK", "Not found")]

    st.subheader("Result")
    summary = f"{len(found)} of {len(queries)} report(s) read from IGI."
    (st.success if not (missing or failed) else st.warning)(summary)
    if missing:
        st.write("**Not found at IGI:** " + ", ".join(missing))
    by_message: Dict[str, List[str]] = {}
    for q in failed:
        by_message.setdefault((results.get(q) or {}).get("message") or "Not read yet.", []).append(q)
    for message, numbers_ in by_message.items():
        shown = ", ".join(numbers_[:20]) + (f" … and {len(numbers_) - 20} more" if len(numbers_) > 20 else "")
        st.error(f"{shown}: {message}")
    if not ON_WINDOWS and any("Cloudflare" in message for message in by_message):
        st.info(BLOCKED_HELP)

    df = build_table(entries, results, igi_number, date_fmt, skip_missing)
    if df.empty:
        st.info("No rows to show.")
        return None

    st.caption("You can type in the table (Lot no, Document No / Name, P/ct, $ Rate, Cost P/Ct …); "
               "the downloads in the sidebar and the copy box below use what you see.")
    key = f"table_{zlib.crc32(df.to_json().encode('utf-8')):08x}"   # new table -> fresh editor
    edited = st.data_editor(df, key=key, hide_index=True, num_rows="fixed", column_config=column_config())

    st.markdown("**Copy for Excel**: click the copy icon at the top right of the box, then paste into your sheet.")
    with_header = st.checkbox("Include the header row", value=False)
    st.code(tsv_text(edited, with_header), language=None)

    with st.expander("What IGI returned, per report"):
        pick = st.selectbox("Report", queries)
        result = results.get(pick) or {}
        line = f"**{result.get('status') or 'Not read'}**"
        if result.get("message"):
            line += f" · {result['message']}"
        if result.get("source"):
            line += f" · read from the {route_label(result['source'], result.get('endpoint') or 'page')}"
        st.markdown(line)
        if result.get("record"):
            st.json(result["record"])
    return edited


def result_buttons(edited: pd.DataFrame, date_fmt: str, save_folder: str) -> None:
    """Download Excel / CSV and, on a Windows PC, Save Excel (shown in the sidebar)."""
    today = datetime.now().strftime("%d-%m-%Y")
    xlsx = excel_bytes(edited, date_fmt)
    st.download_button(
        "Download Excel", xlsx, file_name=SAVE_FILE_NAME.format(date=today),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button("Download CSV", csv_bytes(edited), file_name=f"IGI Data {today}.csv", mime="text/csv")
    if save_folder and st.button(f"Save Excel in {save_folder}"):     # Windows PC only
        try:
            st.success(f"Saved {save_excel(xlsx, save_folder, SAVE_FILE_NAME.format(date=today))}")
        except OSError as exc:
            st.error(f"Could not save in {save_folder}: {exc}")


def browser_steps(entries: List[Dict[str, Any]]) -> None:
    """Sidebar steps for Streamlit Cloud: IGI is read in the user's own browser with the bookmark."""
    state = st.session_state
    queries = list(dict.fromkeys(e["query"] for e in entries))
    st.subheader("Read in your browser")
    st.caption("IGI blocks Streamlit Cloud's server, so your own browser reads IGI:")
    if queries:
        st.link_button(f"1. Open IGI ({len(queries)} report(s)) ↗", browser_link(queries))
    else:
        st.caption("1. Paste report numbers; an Open IGI button appears here.")
    st.caption("2. On the IGI tab click your IGI → sheet bookmark, then Copy result.")
    with st.expander("Add the IGI → sheet bookmark (once)"):
        components.html(BOOKMARK_BUTTON_HTML.replace("__URL__", html.escape(bookmarklet_url(), quote=True)),
                        height=90)
        st.caption("Drag the red button onto your bookmarks bar (Ctrl+Shift+B shows the bar), "
                   "or add a bookmark by hand with this address:")
        st.code(bookmarklet_url(), language=None)
    pasted = st.text_area("3. Paste the result here", key="browser_result", height=90,
                          placeholder='{"igi_reports": [ … ]}')
    if pasted.strip():
        results, problem = parse_browser_result(pasted)
        if problem:
            st.warning(problem)
        if results:
            state["results"].update(results)
            known = [e["query"] for e in entries]
            state["entries"] = entries + [{"query": q, "extras": {}} for q in results if q not in known]
            st.success(f"Took {len(results)} report(s) from the result.")


def clear_numbers() -> None:
    st.session_state["numbers"] = ""


def clear_results() -> None:
    st.session_state["results"] = {}
    st.session_state["entries"] = []
    st.session_state["browser_result"] = ""


def main() -> None:
    st.set_page_config(page_title="IGI report data", page_icon="💎", layout="wide")
    state = st.session_state
    state.setdefault("results", {})
    state.setdefault("entries", [])
    state.setdefault("numbers", "")
    state.setdefault("browser_result", "")
    show_chrome = SHOW_CHROME and BROWSER_KIND == "uc"     # the server's Chromium is always hidden
    save_folder = SAVE_FOLDER if ON_WINDOWS else ""

    # The page: the report numbers, then the result. The buttons are in the sidebar; no settings are shown.
    st.title("IGI report data")
    pasted = st.text_area(
        "Report numbers", key="numbers", height=170,
        placeholder="852601380\nLG628475692\n…\n\nor paste rows from Excel including the header row",
    )
    entries, notes = collect_entries(pasted)
    for note in notes:
        st.warning(note)
    progress_area = st.container()

    with st.sidebar:
        unique = len(dict.fromkeys(e["query"] for e in entries))
        fetch = st.button(f"Fetch {unique} report(s) from IGI" if entries else "Fetch from IGI",
                          type="primary", disabled=not entries)
        st.button("Clear report numbers", on_click=clear_numbers, help="Empties the Report numbers box.")
        st.button("Clear fetched reports", on_click=clear_results, help="Removes the result table.")

        if not ON_WINDOWS:                  # Streamlit Cloud: IGI only answers your own browser
            st.divider()
            browser_steps(entries)

    if fetch:
        with progress_area:
            run_fetch(entries, READ_METHOD, show_chrome)

    edited = None
    if state["entries"]:
        edited = show_results(state["entries"], state["results"], REPORT_NO_AS_IGI, DATE_FORMAT, SKIP_MISSING)
    if edited is not None:
        with st.sidebar:
            st.divider()
            st.subheader("Result")
            result_buttons(edited, DATE_FORMAT, save_folder)


def _inside_streamlit() -> bool:
    try:
        from streamlit import runtime
        return runtime.exists()
    except Exception:
        return False


if __name__ == "__main__":
    if _inside_streamlit():
        main()
    else:                                   # "python streamlit_app.py": start Streamlit
        from streamlit.web import cli as stcli

        sys.argv = ["streamlit", "run", os.path.abspath(__file__)]
        sys.exit(stcli.main())
