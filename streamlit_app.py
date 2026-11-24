"""IGI Report Data: paste IGI report numbers, fetch them from IGI, download your 28-column sheet.

Runs on Streamlit Cloud (igi-data.streamlit.app) and on a PC (`streamlit run streamlit_app.py`).

How the reports are read when you click Fetch:
- The IGI Reader extension is in this browser (Chrome or Edge): your browser reads IGI. First IGI's API
  (api.igi.org/ReportDetail.php?Printno=...); if the API asks for its "Verify you are human" check
  instead of answering, IGI's report page (www.igi.org/API-IGI/report-diagnosis.php), which carries the
  same data.
- No extension, on a Windows PC: this app reads the same two IGI addresses itself.
- No extension, on Streamlit Cloud: IGI refuses Streamlit's servers (HTTP 403) and won't let a
  streamlit.app page call its API (no CORS), so the page shows how to add the extension. The extension
  itself can be downloaded from the page; its code is in this file (READER_* below).
"""
from __future__ import annotations

import html
import io
import json
import numbers
import os
import re
import sys
import traceback
import zipfile
import zlib
from dataclasses import asdict, dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st

# -----------------------------------------------------------------------------
# Settings
# -----------------------------------------------------------------------------
API_URL = os.environ.get("IGI_API_URL", "https://api.igi.org/ReportDetail.php?Printno={report_no}")
PAGE_URL = os.environ.get("IGI_PAGE_URL", "https://www.igi.org/API-IGI/report-diagnosis.php?r={report_no}")

# Without the extension, does this app read IGI itself? Yes on a Windows PC; not on Streamlit Cloud,
# whose servers IGI refuses. Set the environment variable IGI_READ_ON_SERVER=1 (or 0) to choose.
READ_ON_SERVER = os.environ.get("IGI_READ_ON_SERVER", "1" if sys.platform == "win32" else "0") == "1"

# Pages where the IGI Reader extension offers itself (its content-script matches). Ports don't matter.
APP_PAGES = ["https://igi-data.streamlit.app/*", "http://localhost/*", "http://127.0.0.1/*"]
READER_VERSION = "1.0.0"

SAVE_FILE_NAME = "IGI Data {date}.xlsx"   # the downloaded file; {date} = dd-mm-yyyy
REPORT_NO_AS_IGI = False                   # False: Report No. as you typed it; True: as IGI prints it
DATE_FORMAT = "dd/mm/yyyy"                 # Report Date: one of the names in DATE_FORMATS below
NO_DATA = "No Data"                        # shown for a report IGI did not return (not found, refused …)
NO_DATA_COLUMN = "Shape"                   # the column that says No Data on such a report's row

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

# Report Date: shown as text in this format; the Excel file gets a real date.
DATE_FORMATS: Dict[str, Tuple[Optional[str], Optional[str]]] = {
    "dd/mm/yyyy": ("%d/%m/%Y", "dd/mm/yyyy"),
    "dd-mm-yyyy": ("%d-%m-%Y", "dd-mm-yyyy"),
    "dd-Mmm-yyyy": ("%d-%b-%Y", "dd-mmm-yyyy"),
    "As IGI writes it": (None, None),
}

# Used only when this app reads IGI itself (a Windows PC without the extension).
BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/html;q=0.9, */*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
SERVER_TIMEOUT = (6, 25)                   # seconds: connect, read

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


def log_problem(text: str, error: bool = False) -> None:
    """The reason behind a No Data, for the console (PC) or the app's logs (Streamlit Cloud)."""
    print(f"[IGI] {text}", file=sys.stderr, flush=True)
    if error:
        traceback.print_exc()


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
    "report_no": ["REPORT NUMBER", "Report No", "Report #", "SUMMARY NO"],
    "date": ["REPORT DATE", "Issue Date", "Date of Issue"],
    "shape": ["SHAPE AND CUT", "Shape and Cutting Style", "Shape & Cut", "Shape", "SHAPE/CUT"],
    "weight": ["CARAT WEIGHT", "Weight", "ESTIMATED CARAT WEIGHT", "TOTAL ESTIMATED WEIGHT"],
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
# IGI's answers: the API's JSON, or the JSON on IGI's report page
# -----------------------------------------------------------------------------
@dataclass
class FetchResult:
    query: str
    status: str                       # "OK", "Not found" or "Error"
    record: Dict[str, Any]
    message: str = ""
    endpoint: str = ""                # "api" or "page"
    via: str = ""                     # "browser" (IGI Reader) or "server" (this app)
    check: bool = False               # IGI answered with its "Verify you are human" check


def has_data(result: Optional[Dict[str, Any]]) -> bool:
    """True when IGI returned the report; anything else (not found, refused, any error) is No Data."""
    return bool(result) and result.get("status") == "OK" and bool(result.get("record"))


REPORT_KEYS = {"reportnumber", "caratweight", "shapeandcut", "summaryno", "shapecut"}


def is_report_record(record: Any) -> bool:
    return isinstance(record, dict) and bool({_norm(k) for k in record} & REPORT_KEYS)


def records_from(value: Any) -> List[Dict[str, Any]]:
    """IGI's records from a JSON value: [{...}], one {...}, or jewellery data as [{"Label", "Value"}]
    pairs (turned into one record). null, false or a message for an unknown number gives []."""
    if isinstance(value, dict):
        value = [value]
    if not isinstance(value, list):
        return []
    items = [item for item in value if isinstance(item, dict)]
    if items and all("Label" in item and "Value" in item for item in items):
        flat: Dict[str, Any] = {}
        for item in items:
            label = clean_text(item.get("Label"))
            if label and label not in flat:
                flat[label] = item.get("Value")
        items = [flat] if flat else []
    return [item for item in items if is_report_record(item)]


def merge_records(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    merged: Dict[str, Any] = {}
    for record in records:
        for key, value in record.items():
            if key not in merged or is_blank(merged[key]):
                merged[key] = value
    return merged


def is_human_check(status: int, text: str) -> bool:
    """IGI (Cloudflare) answered with its 'Verify you are human' page instead of the data."""
    low = (text or "")[:20000].lower()
    markers = ("just a moment", "verify you are human", "performing security verification",
               "challenge-platform", "attention required", "you have been blocked")
    return any(marker in low for marker in markers) or (status in (403, 429, 503) and "<html" in low)


def read_api_text(text: str) -> Optional[List[Dict[str, Any]]]:
    """Records from the API's answer; None when the answer isn't JSON (a page, an error …)."""
    body = (text or "").lstrip("\ufeff").strip()
    if not body or body[:1] == "<":
        return None
    try:
        value, _ = json.JSONDecoder(strict=False).raw_decode(body)
    except ValueError:
        return None
    return records_from(value)


def json_list_after(text: str, heading: str, stop: str = "") -> Optional[List[Any]]:
    """The JSON array after 'data:' under a heading of IGI's report page."""
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
    for candidate in (text, html.unescape(text[bracket:])):
        at = bracket if candidate is text else 0
        try:
            value, _ = json.JSONDecoder(strict=False).raw_decode(candidate, at)
            return value if isinstance(value, list) else None
        except ValueError:
            continue
    return None


def read_page_text(text: str) -> Optional[List[Dict[str, Any]]]:
    """Records from IGI's report page (loose report data first, then jewellery data);
    None when the text isn't IGI's report page."""
    if "IGI Report Diagnosis" not in (text or ""):
        return None
    loose = records_from(json_list_after(text, "LOOSE REPORT DATA", stop="SWIFT/JEWELRY DATA"))
    return loose or records_from(json_list_after(text, "SWIFT/JEWELRY DATA"))


# -----------------------------------------------------------------------------
# Reading IGI from this app's own server (a Windows PC without the extension)
# -----------------------------------------------------------------------------
def server_get(session: requests.Session, endpoint: str, query: str) -> FetchResult:
    url = (API_URL if endpoint == "api" else PAGE_URL).format(report_no=quote(query, safe=""))

    def result(status: str, records: Optional[List[Dict[str, Any]]] = None,
               message: str = "", check: bool = False) -> FetchResult:
        record = merge_records(records) if records else {}
        return FetchResult(query, status, record, message, endpoint, "server", check)

    try:
        response = session.get(url, timeout=SERVER_TIMEOUT)
    except requests.RequestException as exc:
        return result("Error", message=f"IGI did not answer: {str(exc)[:200]}")
    text = response.text
    records = read_api_text(text) if endpoint == "api" else read_page_text(text)
    if records is not None and response.ok:
        return result("OK", records) if records else result("Not found", message="IGI has no report with this number.")
    if is_human_check(response.status_code, text):
        return result("Error", message=f"IGI asked for its human check (HTTP {response.status_code})", check=True)
    return result("Error", message=f"IGI's answer had no report data (HTTP {response.status_code})")


def server_read(session: requests.Session, query: str, memo: Dict[str, bool]) -> FetchResult:
    """The API first; IGI's report page when the API doesn't give the report."""
    api = None
    if not memo.get("api_refused"):
        api = server_get(session, "api", query)
        if api.status == "OK":
            return api
        memo["api_refused"] = api.check            # the API asks for its check: use the page from now on
    page = server_get(session, "page", query)
    if page.status == "Error" and api is not None and api.status == "Not found":
        return api
    return page


def run_server_fetch(queries: List[str]) -> Dict[str, Dict[str, Any]]:
    results: Dict[str, Dict[str, Any]] = {}
    bar = st.progress(0.0, text="Starting…")
    memo: Dict[str, bool] = {}
    with requests.Session() as session:
        session.headers.update(BROWSER_HEADERS)
        for i, query in enumerate(queries):
            bar.progress(i / len(queries), text=f"Reading {query}  ({i + 1} of {len(queries)})")
            try:
                result = server_read(session, query, memo)
            except Exception as exc:                  # any error: this report has No Data
                log_problem(f"{query}: error while reading", error=True)
                result = FetchResult(query, "Error", {}, str(exc)[:200], "", "server")
            results[query] = asdict(result)
    bar.progress(1.0, text=f"Done: {len(queries)} report(s) read.")
    return results


# -----------------------------------------------------------------------------
# The result sent back by the IGI Reader extension
# -----------------------------------------------------------------------------
def parse_reader_result(text: str) -> Dict[str, Dict[str, Any]]:
    """Results keyed by report number from the extension's JSON."""
    try:
        data = json.loads(text or "")
    except ValueError:
        log_problem("IGI Reader sent something that isn't JSON")
        return {}
    items = data.get("igi_reports") if isinstance(data, dict) else None
    results: Dict[str, Dict[str, Any]] = {}
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict) or not str(item.get("query") or "").strip():
            continue
        query = clean_report_cell(item["query"]) or str(item["query"]).strip().upper()
        records = records_from(item.get("records"))
        endpoint = str(item.get("endpoint") or "")
        if item.get("status") == "OK" and records:
            result = FetchResult(query, "OK", merge_records(records), "", endpoint, "browser")
        elif item.get("status") == "Not found":
            result = FetchResult(query, "Not found", {}, "IGI has no report with this number.", endpoint, "browser")
        else:
            result = FetchResult(query, "Error", {}, str(item.get("message") or "not read")[:300], endpoint,
                                 "browser", bool(item.get("check")))
        results[query] = asdict(result)
    return results


# -----------------------------------------------------------------------------
# Input: pasted report numbers, or rows copied from Excel with their header row
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


def parse_pasted(text: str) -> Tuple[List[Dict[str, Any]], List[str]]:
    lines = (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    rows = [line.split("\t") for line in lines if line.strip()]
    if not rows:
        return [], []
    notes: List[str] = []
    shortened = excel_shortened(text)
    if shortened:
        notes.append(
            "These look like numbers Excel shortened, and the real report number is lost: "
            + ", ".join(dict.fromkeys(shortened))
            + ". Format the column as Number with 0 decimals in Excel and copy again."
        )
    report_idx = find_report_column(rows[0])
    if report_idx is not None:            # rows copied from Excel with their header row
        return rows_to_entries(rows[0], rows[1:], report_idx, notes), notes
    queries = list(dict.fromkeys(report_tokens(text)))
    return [{"query": q, "extras": {}} for q in queries], notes


# -----------------------------------------------------------------------------
# Table and Excel file
# -----------------------------------------------------------------------------
def number_text(value: Any) -> str:
    number = float(value)
    if number != number or number in (float("inf"), float("-inf")):
        return ""
    return str(int(number)) if number.is_integer() else repr(number)


def cell_text(value: Any) -> str:
    """A cell as it is shown: 1.00 -> '1', 8.50 -> '8.5', blanks -> ''."""
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


def igi_row(query: str, result: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """The row from IGI's record, or None when the report has No Data."""
    if not has_data(result):
        return None
    try:
        return build_row(query, result["record"], REPORT_NO_AS_IGI, DATE_FORMAT)
    except Exception:                         # a record that cannot be read is No Data as well
        log_problem(f"{query} -> {NO_DATA}: IGI's record could not be read", error=True)
        return None


def build_table(entries: List[Dict[str, Any]], results: Dict[str, Dict[str, Any]]) -> Tuple[pd.DataFrame, List[str]]:
    """The 28-column table as text, as it is shown, and the reports with No Data.

    A report IGI did not return keeps its row with your pasted values, and NO_DATA_COLUMN says No Data.
    """
    rows = []
    no_data: Dict[str, None] = {}
    for entry in entries:
        query = entry["query"]
        row = igi_row(query, results.get(query))
        if row is None:
            no_data[query] = None
            row = build_row(query, {}, REPORT_NO_AS_IGI, DATE_FORMAT)
            row[NO_DATA_COLUMN] = NO_DATA
        for column, value in (entry.get("extras") or {}).items():
            if is_blank(row.get(column)):         # IGI's value wins; your pasted value fills the rest
                row[column] = to_cell_value(column, value)
        rows.append({column: cell_text(row[column]) for column in OUTPUT_COLUMNS})
    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS, dtype=object), list(no_data)


def excel_bytes(df: pd.DataFrame) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    python_fmt, excel_fmt = DATE_FORMATS.get(DATE_FORMAT, (None, None))
    wb = Workbook()
    ws = wb.active
    ws.title = "IGI"
    ws.append(OUTPUT_COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    texts = [[cell_text(row[column]) for column in OUTPUT_COLUMNS] for _, row in df.iterrows()]
    for values_text in texts:
        values: List[Any] = []
        for column, value in zip(OUTPUT_COLUMNS, values_text):
            if not value:
                values.append(None)
            elif column in NUMBER_COLUMNS:
                number = as_number(value)
                values.append(number if number is not None else value)
            elif column == "Report No." and value.isdigit() and not value.startswith("0"):
                values.append(int(value))
            elif column == "Report Date" and python_fmt:
                try:
                    values.append(datetime.strptime(value, python_fmt))
                except ValueError:
                    values.append(value)
            else:
                values.append(value)
        ws.append(values)

    date_col = OUTPUT_COLUMNS.index("Report Date") + 1
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=date_col)
        if isinstance(cell.value, datetime) and excel_fmt:
            cell.number_format = excel_fmt

    for i, column in enumerate(OUTPUT_COLUMNS):
        longest = max([len(column)] + [len(r[i]) for r in texts])
        ws.column_dimensions[get_column_letter(i + 1)].width = min(max(longest + 2, 6), 45)
    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# The IGI Reader extension (Chrome / Edge). The page offers it as a zip; load it unpacked.
#
# content.js runs on this app's pages. The Fetch button posts {igiSheet: "read", id, numbers} to its
# own window; content.js asks background.js to read each number (3 at a time) and posts progress and
# the result back. background.js does the reading: an extension may read the IGI addresses it lists
# in host_permissions, which a streamlit.app page may not (IGI sends no CORS header).
# -----------------------------------------------------------------------------
READER_BACKGROUND_JS = r"""
// IGI Reader: reads one IGI report per request from the IGI Report Data page.
// First IGI's API. If the API answers with its "Verify you are human" check instead of data,
// IGI's report page is used (it carries the same data), for the next 10 minutes as well.
"use strict";

const API_URL = (n) => "https://api.igi.org/ReportDetail.php?Printno=" + encodeURIComponent(n);
const PAGE_URL = (n) => "https://www.igi.org/API-IGI/report-diagnosis.php?r=" + encodeURIComponent(n);
const TIMEOUT_MS = 25000;
const API_PAUSE_MS = 10 * 60 * 1000;
let apiPausedUntil = 0;

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || message.type !== "igi-read" || sender.id !== chrome.runtime.id) {
    return false;
  }
  const number = String(message.number || "").trim().toUpperCase();
  readReport(number)
    .then(sendResponse)
    .catch((error) => sendResponse({ query: number, status: "Error", message: String((error && error.message) || error) }));
  return true;
});

async function getText(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const response = await fetch(url, { credentials: "include", cache: "no-store", signal: controller.signal });
    return { status: response.status, ok: response.ok, text: await response.text() };
  } finally {
    clearTimeout(timer);
  }
}

function isHumanCheck(answer) {
  const low = (answer.text || "").slice(0, 20000).toLowerCase();
  const markers = ["just a moment", "verify you are human", "performing security verification",
                   "challenge-platform", "attention required", "you have been blocked"];
  return markers.some((m) => low.includes(m)) || ([403, 429, 503].includes(answer.status) && low.includes("<html"));
}

// IGI's records from a JSON value: [{...}], one {...}, or jewellery data as Label/Value pairs.
function toRecords(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) { value = [value]; }
  if (!Array.isArray(value)) { return []; }
  let items = value.filter((item) => item && typeof item === "object" && !Array.isArray(item));
  if (items.length && items.every((item) => "Label" in item && "Value" in item)) {
    const flat = {};
    for (const item of items) {
      const label = String(item.Label || "").trim();
      if (label && !(label in flat)) { flat[label] = item.Value; }
    }
    items = Object.keys(flat).length ? [flat] : [];
  }
  return items.filter((item) => Object.values(item).some((v) => v !== null && v !== ""));
}

function parseJson(text) {
  return JSON.parse(text.replace(/^\uFEFF/, "").replace(/[\u0000-\u001F]+/g, " "));
}

function decodeEntities(text) {
  return text.replace(/&quot;/g, '"').replace(/&#0?39;|&apos;/g, "'").replace(/&lt;/g, "<")
             .replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}

// The JSON array after "data:" under a heading of IGI's report page.
function jsonAfter(html, heading, stop) {
  const start = html.indexOf(heading);
  if (start < 0) { return null; }
  const end = stop ? html.indexOf(stop, start + heading.length) : -1;
  const dataAt = html.indexOf("data:", start);
  if (dataAt < 0 || (end >= 0 && dataAt > end)) { return null; }
  const s = html.indexOf("[", dataAt);
  if (s < 0 || html.slice(dataAt + 5, s).replace(/<br\s*\/?>|\s/gi, "")) { return null; }
  let depth = 0, inString = false, escaped = false;
  for (let i = s; i < html.length; i++) {
    const c = html[i];
    if (inString) {
      if (escaped) { escaped = false; } else if (c === "\\") { escaped = true; } else if (c === '"') { inString = false; }
    } else if (c === '"') { inString = true; }
    else if (c === "[") { depth++; }
    else if (c === "]" && --depth === 0) {
      const raw = html.slice(s, i + 1);
      for (const candidate of [raw, decodeEntities(raw)]) {
        try { return parseJson(candidate); } catch (e) { /* try the next form */ }
      }
      return null;
    }
  }
  return null;
}

async function readApi(number) {
  const answer = await getText(API_URL(number));
  const body = answer.text.replace(/^\uFEFF/, "").trim();
  if (answer.ok && body && body[0] !== "<") {
    try {
      const records = toRecords(parseJson(body));
      return records.length
        ? { query: number, status: "OK", records: records, endpoint: "api" }
        : { query: number, status: "Not found", endpoint: "api" };
    } catch (e) { /* not JSON */ }
  }
  return { query: number, status: "Error", endpoint: "api", check: isHumanCheck(answer),
           message: "IGI's API gave no data (HTTP " + answer.status + ")" };
}

async function readPage(number) {
  let answer;
  try {
    answer = await getText(PAGE_URL(number));
  } catch (error) {
    return { query: number, status: "Error", endpoint: "page",
             message: "IGI did not answer: " + (error.name === "AbortError" ? "timed out" : error.message) };
  }
  if (!answer.text.includes("IGI Report Diagnosis")) {
    const check = isHumanCheck(answer);
    return { query: number, status: "Error", endpoint: "page", check: check,
             message: check ? "IGI asked for its human check (HTTP " + answer.status + ")"
                            : "IGI did not return its report page (HTTP " + answer.status + ")" };
  }
  let records = toRecords(jsonAfter(answer.text, "LOOSE REPORT DATA", "SWIFT/JEWELRY DATA"));
  if (!records.length) { records = toRecords(jsonAfter(answer.text, "SWIFT/JEWELRY DATA", "")); }
  return records.length
    ? { query: number, status: "OK", records: records, endpoint: "page" }
    : { query: number, status: "Not found", endpoint: "page" };
}

async function readReport(number) {
  if (!/^[A-Z]{0,3}\d{6,}$/.test(number)) {
    return { query: number, status: "Error", message: "not a report number" };
  }
  let api = null;
  if (Date.now() >= apiPausedUntil) {
    try { api = await readApi(number); } catch (error) { api = null; }
    if (api && api.status === "OK") { return api; }
    if (api && api.check) { apiPausedUntil = Date.now() + API_PAUSE_MS; }
  }
  const page = await readPage(number);
  if (page.status === "Error" && api && api.status === "Not found") { return api; }
  return page;
}
"""

READER_CONTENT_JS = r"""
// IGI Reader: on the IGI Report Data page. The page's Fetch button posts
// {igiSheet: "read", id, numbers}; this asks the extension to read each number from IGI
// (3 at a time) and posts the progress and the result back to the page.
(function () {
  "use strict";
  var version = chrome.runtime.getManifest().version;
  if (document.documentElement) {
    document.documentElement.setAttribute("data-igi-reader", version);
  }
  var running = {};

  function post(message) {
    window.postMessage(message, "*");
  }

  function readOne(number) {
    return new Promise(function (resolve) {
      try {
        chrome.runtime.sendMessage({ type: "igi-read", number: number }, function (answer) {
          var problem = chrome.runtime.lastError;
          if (problem || !answer) {
            resolve({ query: number, status: "Error", message: (problem && problem.message) || "no answer from IGI Reader" });
          } else {
            resolve(answer);
          }
        });
      } catch (e) {
        resolve({ query: number, status: "Error", message: "IGI Reader was reloaded: reload this page (F5)" });
      }
    });
  }

  window.addEventListener("message", function (event) {
    if (event.source !== window) { return; }
    var m = event.data;
    if (!m || m.igiSheet !== "read" || typeof m.id !== "string" || !Array.isArray(m.numbers) || running[m.id]) {
      return;
    }
    running[m.id] = true;
    var numbers = m.numbers.map(String).slice(0, 3000);
    var out = new Array(numbers.length);
    var next = 1;
    var done = 0;
    post({ igiReader: "progress", id: m.id, done: 0, total: numbers.length });
    function read(i) {
      return readOne(numbers[i]).then(function (answer) {
        out[i] = answer;
        done++;
        post({ igiReader: "progress", id: m.id, done: done, total: numbers.length });
      });
    }
    function worker() {
      if (next >= numbers.length) { return Promise.resolve(); }
      return read(next++).then(worker);
    }
    // The first report alone: if IGI's API asks for its check, the extension has switched to
    // IGI's report page before the rest start. Then 3 at a time.
    var first = numbers.length ? read(0) : Promise.resolve();
    first.then(function () {
      var workers = [];
      for (var w = 0; w < Math.min(3, numbers.length - 1); w++) { workers.push(worker()); }
      return Promise.all(workers);
    }).then(function () {
      delete running[m.id];
      post({ igiReader: "done", id: m.id, reports: out, via: "IGI Reader " + version });
    });
  });
})();
"""

READER_README = """IGI Reader {version}

Reads IGI report data for the IGI Report Data app ({pages}) through this browser.

Install (once per PC):
1. Unzip this folder to a place you will keep, for example C:\\IGI Reader
2. Chrome: open chrome://extensions  (Edge: edge://extensions) and switch on Developer mode.
3. Click "Load unpacked" and choose this folder (the one with manifest.json in it).
4. Open the app, press F5, paste report numbers and click Fetch.

What it reads: IGI's API (api.igi.org/ReportDetail.php?Printno=...) first; if IGI's API asks for its
"Verify you are human" check instead of answering, IGI's report page
(www.igi.org/API-IGI/report-diagnosis.php?r=...), which carries the same data.
It only acts on the app's pages listed above, and only when you click Fetch there.
"""


def reader_manifest() -> str:
    manifest = {
        "manifest_version": 3,
        "name": "IGI Reader",
        "version": READER_VERSION,
        "description": "Reads IGI report data for the IGI Report Data app, through this browser.",
        "icons": {"16": "icon16.png", "48": "icon48.png", "128": "icon128.png"},
        "host_permissions": ["https://api.igi.org/*", "https://www.igi.org/*"],
        "background": {"service_worker": "background.js"},
        "content_scripts": [{
            "matches": APP_PAGES,
            "js": ["content.js"],
            "run_at": "document_start",
            "all_frames": True,
        }],
    }
    return json.dumps(manifest, indent=2) + "\n"


def reader_icon_png(size: int) -> bytes:
    """A red diamond, drawn large and scaled down."""
    from PIL import Image, ImageDraw

    s = size * 8
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    top, girdle, tip = 0.20 * s, 0.42 * s, 0.92 * s
    outline = [(0.28 * s, top), (0.72 * s, top), (0.95 * s, girdle), (0.50 * s, tip), (0.05 * s, girdle)]
    draw.polygon(outline, fill=(255, 75, 75, 255))
    white, width = (255, 255, 255, 235), max(2, s // 28)
    draw.line([(0.05 * s, girdle), (0.95 * s, girdle)], fill=white, width=width)
    draw.line([(0.28 * s, top), (0.40 * s, girdle), (0.50 * s, top)], fill=white, width=width)
    draw.line([(0.50 * s, top), (0.60 * s, girdle), (0.72 * s, top)], fill=white, width=width)
    draw.line([(0.40 * s, girdle), (0.50 * s, tip), (0.60 * s, girdle)], fill=white, width=width)
    img = img.resize((size, size), Image.LANCZOS)
    out = io.BytesIO()
    img.save(out, "PNG")
    return out.getvalue()


def reader_files() -> Dict[str, bytes]:
    files = {
        "manifest.json": reader_manifest().encode("utf-8"),
        "background.js": READER_BACKGROUND_JS.lstrip().encode("utf-8"),
        "content.js": READER_CONTENT_JS.lstrip().encode("utf-8"),
        "README.txt": READER_README.format(
            version=READER_VERSION, pages=", ".join(APP_PAGES)).replace("\n", "\r\n").encode("utf-8"),
    }
    for size in (16, 48, 128):
        files[f"icon{size}.png"] = reader_icon_png(size)
    return files


def reader_zip_bytes() -> bytes:
    """The extension as a zip, its files at the root (built when asked for: a few milliseconds)."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in reader_files().items():
            info = zipfile.ZipInfo(name, date_time=(2026, 11, 24, 12, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, data)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# The Fetch button: a Streamlit components-v2 element (it runs in the app's page), styled like
# Streamlit's own red button. With IGI Reader in the browser it reads IGI there; otherwise it asks
# Python to read on the server (a Windows PC) or to show how to add the extension (Streamlit Cloud).
# -----------------------------------------------------------------------------
FETCH_HTML = '<button type="button" disabled>Fetch from IGI</button>'

FETCH_CSS = """
button {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
  min-height: 2.5rem;
  margin: 0;
  padding: 0.25rem 0.75rem;
  font-family: var(--st-font, inherit);
  font-size: 1rem;
  font-weight: var(--st-base-font-weight, 400);
  line-height: 1.6;
  color: #fff;
  background-color: var(--st-primary-color);
  border: 1px solid var(--st-primary-color);
  border-radius: var(--st-button-radius, 0.5rem);
  cursor: pointer;
  user-select: none;
}
button:hover:not(:disabled),
button:active:not(:disabled) {
  background-color: color-mix(in srgb, var(--st-primary-color) 70%, black);
  border-color: color-mix(in srgb, var(--st-primary-color) 70%, black);
}
button:focus-visible {
  outline: none;
  box-shadow: 0 0 0 0.2rem color-mix(in srgb, var(--st-primary-color) 50%, transparent);
}
button:disabled {
  color: color-mix(in srgb, var(--st-text-color) 40%, transparent);
  background-color: transparent;
  border-color: var(--st-border-color);
  cursor: not-allowed;
}
button.busy:disabled {
  color: #fff;
  background-color: color-mix(in srgb, var(--st-primary-color) 75%, white);
  border-color: color-mix(in srgb, var(--st-primary-color) 75%, white);
  cursor: progress;
}
"""

FETCH_JS = r"""
export default function (component) {
  const { data, parentElement, setTriggerValue } = component;
  const d = data || {};
  const numbers = Array.isArray(d.numbers) ? d.numbers : [];
  const todo = Array.isArray(d.todo) ? d.todo : [];
  const button = parentElement.querySelector("button");
  const g = window.__igiFetch || (window.__igiFetch = { job: null });
  const IDLE_MS = 70000;   // no word from IGI Reader for this long: give up
  const reader = () => document.documentElement.getAttribute("data-igi-reader") || "";

  g.setTriggerValue = setTriggerValue;     // the newest one, for a read that ends after a rerun
  g.render = () => {
    if (g.job) {
      button.disabled = true;
      button.classList.add("busy");
      button.textContent = "Reading " + g.job.done + " of " + g.job.total + "…";
      return;
    }
    button.classList.remove("busy");
    button.disabled = !numbers.length;
    button.textContent = d.label || "Fetch from IGI";
    button.title = reader()
      ? "Reads IGI through this browser (IGI Reader " + reader() + ")"
      : (d.server ? "Reads IGI from this computer" : "Needs the IGI Reader extension in this browser");
  };
  g.render();

  g.stalled = () => {
    if (!g.job) { return; }
    g.job = null;
    g.render();
    g.setTriggerValue("stalled", Date.now());
  };
  if (!g.listening) {
    g.listening = true;
    window.addEventListener("message", (event) => {
      const m = event.data;
      if (event.source !== window || !m || !m.igiReader || !g.job || m.id !== g.job.id) { return; }
      clearTimeout(g.job.timer);
      if (m.igiReader === "progress") {
        g.job.done = m.done;
        g.job.total = m.total;
        g.job.timer = setTimeout(() => g.stalled(), IDLE_MS);
        g.render();
      } else if (m.igiReader === "done") {
        g.job = null;
        g.render();
        g.setTriggerValue("reports", JSON.stringify({
          igi_reports: Array.isArray(m.reports) ? m.reports : [],
          via: m.via || "IGI Reader",
          read_at: new Date().toISOString(),
        }));
      }
    });
  }

  button.onclick = () => {
    if (g.job) { return; }
    if (!todo.length) {
      setTriggerValue("already", Date.now());
    } else if (reader()) {
      const id = Date.now().toString(36) + Math.random().toString(36).slice(2);
      g.job = { id: id, done: 0, total: todo.length, timer: setTimeout(() => g.stalled(), IDLE_MS) };
      g.render();
      window.postMessage({ igiSheet: "read", id: id, numbers: todo }, "*");
    } else {
      setTriggerValue(d.server ? "server" : "noreader", Date.now());
    }
  };
}
"""

TRIGGERS = ("reports", "server", "already", "noreader", "stalled")
FETCH_KEY = "igi_fetch"
PLAIN_KEY = "igi_fetch_plain"             # the plain-button fallback's click (a Streamlit without components v2)


def fired_this_run() -> Dict[str, Any]:
    """The Fetch button's signal for this run (reports, server, already, noreader or stalled).

    It is read from session state before the button is drawn, so that the button is drawn with the
    reports still to read *after* this run's result has been taken in.
    """
    state = st.session_state
    fired: Dict[str, Any] = {}
    try:
        values = state.get(FETCH_KEY)
        if values is not None:
            fired = {name: values.get(name) for name in TRIGGERS if values.get(name)}
    except Exception:
        log_problem("the Fetch button's signal could not be read", error=True)
    plain = state.pop(PLAIN_KEY, None)
    if plain:
        fired[plain] = True
    return fired


def _plain_click(has_todo: bool) -> None:
    st.session_state[PLAIN_KEY] = "server" if has_todo else "already"


def fetch_button(label: str, unique: List[str], todo: List[str]) -> None:
    """The Fetch button (sidebar)."""
    try:
        from streamlit.components import v2 as components_v2

        button = components_v2.component("igi_fetch_button", html=FETCH_HTML, css=FETCH_CSS, js=FETCH_JS)
        button(
            key=FETCH_KEY,
            data={"label": label, "numbers": unique, "todo": todo, "server": READ_ON_SERVER},
            **{f"on_{name}_change": (lambda: None) for name in TRIGGERS},
        )
    except Exception:                         # a plain button, reading on the server
        log_problem("the Fetch button could not be shown; using a plain button", error=True)
        st.button(label, type="primary", disabled=not unique, on_click=_plain_click, args=(bool(todo),))


# -----------------------------------------------------------------------------
# Streamlit page
# -----------------------------------------------------------------------------
def column_config() -> Dict[str, Any]:
    return {column: st.column_config.TextColumn(column) for column in OUTPUT_COLUMNS}


def show_results(entries: List[Dict[str, Any]], results: Dict[str, Dict[str, Any]]) -> pd.DataFrame:
    """Summary and the editable 28-column table; returns the table as edited."""
    df, no_data = build_table(entries, results)
    total = len(dict.fromkeys(e["query"] for e in entries))

    st.subheader("Result")
    summary = f"{total - len(no_data)} of {total} report(s) read from IGI."
    (st.success if not no_data else st.warning)(summary)
    if no_data:
        shown = ", ".join(no_data[:20]) + (f" … and {len(no_data) - 20} more" if len(no_data) > 20 else "")
        st.write(f"**{NO_DATA}:** {shown}")
    key = f"table_{zlib.crc32(df.to_json().encode('utf-8')):08x}"   # new table -> fresh editor
    return st.data_editor(df, key=key, hide_index=True, num_rows="fixed", column_config=column_config())


def show_reader_setup() -> None:
    """How to add the IGI Reader extension (shown after Fetch when it isn't in this browser)."""
    with st.container(border=True):
        st.markdown("**One-time setup: add the IGI Reader extension to this browser**")
        st.markdown(
            "IGI doesn't answer this app's server (it refuses Streamlit Cloud with HTTP 403), "
            "so the reports are read through your own browser by a small extension. "
            "Do this once on each PC:\n"
            "1. Download the extension below and unzip it to a folder you will keep, "
            "for example `C:\\IGI Reader`.\n"
            "2. In Chrome open `chrome://extensions` (in Edge: `edge://extensions`) and switch on "
            "**Developer mode**.\n"
            "3. Click **Load unpacked** and choose that folder.\n"
            "4. Come back to this page, press **F5**, and click **Fetch** again."
        )
        st.download_button(
            "Download IGI Reader extension", reader_zip_bytes(), file_name="IGI-Reader-extension.zip",
            mime="application/zip", on_click="ignore",
        )


def take_reader_result(text: str, entries: List[Dict[str, Any]]) -> Tuple[int, List[str]]:
    """Puts the extension's result into the table. Returns (reports read, reports IGI wanted its check for)."""
    state = st.session_state
    results = parse_reader_result(text)
    for query, result in results.items():
        if result["status"] != "OK":
            log_problem(f"{query} -> {NO_DATA} ({result['status']}, {result['endpoint'] or '-'}): {result['message']}")
    state["results"].update(results)
    known = {e["query"] for e in entries}
    state["entries"] = entries + [{"query": q, "extras": {}} for q in results if q not in known]
    checked = [q for q, r in results.items() if r.get("check")]
    return len(results), checked


def clear_all() -> None:
    """Start over: the report numbers and the result table."""
    state = st.session_state
    state["numbers"] = ""
    state["results"] = {}
    state["entries"] = []
    state["show_setup"] = False


def main() -> None:
    st.set_page_config(page_title="IGI Report Data", page_icon="💎", layout="wide")
    state = st.session_state
    state.setdefault("results", {})
    state.setdefault("entries", [])
    state.setdefault("numbers", "")
    state.setdefault("show_setup", False)

    st.title("IGI Report Data")
    pasted = st.text_area("Report numbers", key="numbers", height=140, placeholder="Paste report numbers here…")
    entries, notes = parse_pasted(pasted)
    for note in notes:
        st.warning(note)
    unique = list(dict.fromkeys(e["query"] for e in entries))

    # This run's Fetch signal first, so the button below knows which reports are still to read.
    fired = fired_this_run()
    if fired.get("server") and not [q for q in unique if not has_data(state["results"].get(q))]:
        fired = {"already": True}
    if fired.get("already"):
        state["entries"] = entries                # the table shows the numbers in the box
        st.info("These reports were read already. To read them again, click 'Clear all' in the sidebar "
                "and paste the numbers again.")
    elif fired.get("reports"):
        state["show_setup"] = False
        count, checked = take_reader_result(fired["reports"], entries)
        st.progress(1.0, text=f"Done: {count} report(s) read.")
        if checked:
            st.warning(
                f"IGI asked this browser to confirm it's a person before it shows {len(checked)} report(s). "
                f"Open [IGI's report page]({PAGE_URL.format(report_no=quote(checked[0], safe=''))}) in a new tab, "
                "let it load (tick the box if it asks), then click Fetch again."
            )
    elif fired.get("server"):
        state["entries"] = entries
        results = run_server_fetch([q for q in unique if not has_data(state["results"].get(q))])
        for query, result in results.items():
            if result["status"] != "OK":
                log_problem(f"{query} -> {NO_DATA} ({result['status']}, {result['endpoint'] or '-'}): {result['message']}")
        state["results"].update(results)
        if results and all(r.get("check") for r in results.values()):
            state["show_setup"] = True            # IGI refuses this computer too: read through the browser
    elif fired.get("noreader"):
        state["show_setup"] = True
    elif fired.get("stalled"):
        st.warning("IGI Reader stopped answering. Press F5 to reload this page, then click Fetch again. "
                   "If it happens again, check that IGI Reader is switched on in chrome://extensions.")

    todo = [q for q in unique if not has_data(state["results"].get(q))]
    with st.sidebar:
        label = f"Fetch {len(unique)} report(s) from IGI" if unique else "Fetch from IGI"
        fetch_button(label, unique, todo)
        st.button("Clear all", on_click=clear_all, help="Clears the report numbers and the result.")

    if state["show_setup"]:
        show_reader_setup()

    if state["entries"]:
        edited = show_results(state["entries"], state["results"])
        with st.sidebar:
            st.divider()
            st.subheader("File:")
            st.download_button(
                "Download Excel", excel_bytes(edited),
                file_name=SAVE_FILE_NAME.format(date=datetime.now().strftime("%d-%m-%Y")),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


def _inside_streamlit() -> bool:
    try:
        from streamlit import runtime
        return runtime.exists()
    except Exception:
        return False


if __name__ == "__main__":
    if _inside_streamlit():
        try:
            main()
        except Exception:                         # show No Data, not Streamlit's error box
            log_problem("error on the page", error=True)
            st.warning(NO_DATA)
    else:                                         # "python streamlit_app.py": start Streamlit
        from streamlit.web import cli as stcli

        sys.argv = ["streamlit", "run", os.path.abspath(__file__)]
        sys.exit(stcli.main())
