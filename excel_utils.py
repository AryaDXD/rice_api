"""Excel and CSV processing utilities for bulk student uploads."""

from __future__ import annotations

import csv
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from typing import Any

# Standard schema fields for the Add Student API
# (field_name, is_required, description, example_value)
STUDENT_FIELDS = [
    ("firstName", True, "Student First Name", "ALOKPARNA"),
    ("middleName", False, "Middle Name (Optional)", ""),
    ("lastName", True, "Student Last Name", "KAR"),
    ("primaryEmail", True, "Primary Email", "alokparnakar@gmail.com"),
    ("secondaryEmail", False, "Secondary Email (Optional)", ""),
    ("contactNo", True, "Contact Number", "9051496355"),
    ("address", False, "Address (Optional)", "kolkata"),
    ("batchUniqueId", True, "Batch Unique ID", "RSMSMPSC2627001"),
    ("studentUniqueId", True, "Student Unique ID", "2627/RICESMART/377"),
    ("secondaryLangauge", False, "Secondary Language (Optional)", "hindi"),
    ("classStartDate", False, "Class Start Date (MM/DD/YYYY)", "11/6/2026"),
]

# Flexible header aliases to map user spreadsheet columns to API schema
HEADER_ALIASES = {
    "firstname": "firstName",
    "first": "firstName",
    "fname": "firstName",
    "studentname": "firstName",
    "studentfirstname": "firstName",
    
    "lastname": "lastName",
    "last": "lastName",
    "lname": "lastName",
    "surname": "lastName",
    "studentlastname": "lastName",
    
    "middlename": "middleName",
    "middle": "middleName",
    "mname": "middleName",
    
    "primaryemail": "primaryEmail",
    "email": "primaryEmail",
    "emailid": "primaryEmail",
    "studentemail": "primaryEmail",
    "mail": "primaryEmail",
    
    "secondaryemail": "secondaryEmail",
    "altemail": "secondaryEmail",
    "alternateemail": "secondaryEmail",
    "otheremail": "secondaryEmail",
    
    "contactno": "contactNo",
    "contact": "contactNo",
    "contactnumber": "contactNo",
    "phone": "contactNo",
    "phoneno": "contactNo",
    "phonenumber": "contactNo",
    "mobile": "contactNo",
    "mobileno": "contactNo",
    "mobilenumber": "contactNo",
    "cell": "contactNo",
    
    "address": "address",
    "studentaddress": "address",
    "city": "address",
    "location": "address",
    
    "batchuniqueid": "batchUniqueId",
    "batchid": "batchUniqueId",
    "batch": "batchUniqueId",
    "batchcode": "batchUniqueId",
    "batchno": "batchUniqueId",
    
    "studentuniqueid": "studentUniqueId",
    "studentid": "studentUniqueId",
    "studentcode": "studentUniqueId",
    "rollno": "studentUniqueId",
    "regno": "studentUniqueId",
    "registrationno": "studentUniqueId",
    
    "secondarylangauge": "secondaryLangauge",
    "secondarylanguage": "secondaryLangauge",
    "language": "secondaryLangauge",
    "lang": "secondaryLangauge",
    "secondlanguage": "secondaryLangauge",
    
    "classstartdate": "classStartDate",
    "startdate": "classStartDate",
    "classdate": "classStartDate",
    "batchstartdate": "classStartDate",
}


def normalize_header(header: str) -> str:
    """Map flexible header variations to canonical student field name."""
    cleaned = re.sub(r"[^a-zA-Z0-9]", "", str(header).strip()).lower()
    return HEADER_ALIASES.get(cleaned, str(header).strip())


def col_to_index(col_letters: str) -> int:
    """Convert Excel column letters (A, B, ..., Z, AA, AB) to 0-based index."""
    idx = 0
    for ch in col_letters.upper():
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(0, idx - 1)


def index_to_col(idx: int) -> str:
    """Convert 0-based index to Excel column letters (0 -> A, 27 -> AB)."""
    res = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        res = chr(65 + rem) + res
    return res


def clean_cell_value(val: Any) -> str:
    """Format and clean a cell value, removing scientific notation or trailing .0 on IDs/phone numbers."""
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    # Strip trailing float '.0' from numbers (common when Excel exports integers)
    if s.endswith(".0"):
        prefix = s[:-2]
        if prefix.isdigit() or (prefix.startswith("+") and prefix[1:].isdigit()):
            return prefix
    # Handle serial dates from Excel if numeric
    try:
        f = float(s)
        # Typical Excel date serials fall between 1000 and 80000 (roughly 1902 to 2119)
        if 20000 <= f <= 80000 and ("." not in s or f.is_integer()):
            dt = datetime(1899, 12, 30) + timedelta(days=int(f))
            return dt.strftime("%m/%d/%Y")
    except (ValueError, OverflowError):
        pass
    return s


def parse_xlsx(file_bytes: bytes) -> list[list[str]]:
    """Parse .xlsx bytes using Python standard library (zipfile + xml.etree)."""
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        # 1. Read shared strings if present
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in tree.iter():
                if si.tag.endswith("si"):
                    text = "".join((t.text or "") for t in si.iter() if t.tag.endswith("t"))
                    shared_strings.append(text)

        # 2. Find first worksheet
        sheet_names = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml")]
        sheet_name = sheet_names[0] if sheet_names else "xl/worksheets/sheet1.xml"
        tree = ET.fromstring(z.read(sheet_name))

        # 3. Extract rows and cells
        table: list[list[str]] = []
        for row_el in tree.iter():
            if row_el.tag.endswith("row"):
                cells: dict[int, str] = {}
                for c in row_el.iter():
                    if c.tag.endswith("c"):
                        ref = c.get("r", "")
                        col_letters = "".join(ch for ch in ref if ch.isalpha())
                        c_idx = col_to_index(col_letters) if col_letters else len(cells)
                        t = c.get("t")
                        v_el = next((el for el in c if el.tag.endswith("v")), None)
                        val = v_el.text if v_el is not None else ""
                        if t == "s" and val.isdigit():
                            sid = int(val)
                            val = shared_strings[sid] if sid < len(shared_strings) else ""
                        elif t == "inlineStr":
                            is_t = next((el for el in c.iter() if el.tag.endswith("t")), None)
                            val = is_t.text if is_t is not None else ""
                        cells[c_idx] = clean_cell_value(val)
                if cells:
                    max_idx = max(cells.keys())
                    row_list = [cells.get(i, "") for i in range(max_idx + 1)]
                    if any(c for c in row_list):
                        table.append(row_list)
        return table


def parse_csv(file_bytes: bytes) -> list[list[str]]:
    """Parse CSV bytes handling UTF-8, UTF-8-BOM, and Latin-1."""
    text = ""
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if not text:
        return []

    reader = csv.reader(io.StringIO(text))
    table: list[list[str]] = []
    for row in reader:
        cleaned_row = [clean_cell_value(c) for c in row]
        if any(cleaned_row):
            table.append(cleaned_row)
    return table


def parse_spreadsheet(file_bytes: bytes, filename: str = "") -> list[list[str]]:
    """Auto-detect and parse either an Excel workbook or a CSV file."""
    # Check for ZIP/XLSX magic signature (PK\x03\x04)
    if file_bytes.startswith(b"PK\x03\x04") or filename.lower().endswith(".xlsx"):
        return parse_xlsx(file_bytes)
    return parse_csv(file_bytes)


def process_student_records(
    rows: list[list[str]], defaults: dict[str, str] | None = None
) -> dict[str, Any]:
    """Process 2D rows into normalized student dictionaries with validation."""
    if not rows:
        return {
            "success": False,
            "error": "The uploaded file contains no data.",
            "headers": [],
            "students": [],
            "total": 0,
            "valid_count": 0,
        }

    defaults = defaults or {}
    raw_headers = rows[0]
    column_mapping: dict[int, str] = {}
    mapped_canonical: set[str] = set()

    for idx, raw_h in enumerate(raw_headers):
        canonical = normalize_header(raw_h)
        if any(f[0] == canonical for f in STUDENT_FIELDS):
            column_mapping[idx] = canonical
            mapped_canonical.add(canonical)

    if not column_mapping:
        return {
            "success": False,
            "error": "No matching student columns found. Expected columns like firstName, lastName, primaryEmail, contactNo, studentUniqueId, batchUniqueId.",
            "headers": raw_headers,
            "students": [],
            "total": 0,
            "valid_count": 0,
        }

    students: list[dict[str, Any]] = []
    required_fields = [f[0] for f in STUDENT_FIELDS if f[1]]

    for row_num, row in enumerate(rows[1:], start=2):
        if not any(row):
            continue

        student_data: dict[str, str] = {}
        for f in STUDENT_FIELDS:
            student_data[f[0]] = ""

        for idx, canonical in column_mapping.items():
            if idx < len(row):
                student_data[canonical] = row[idx]

        # Apply default overrides if value is missing
        for def_key, def_val in defaults.items():
            if def_val and not student_data.get(def_key):
                student_data[def_key] = str(def_val).strip()

        # Validation
        missing = [rf for rf in required_fields if not student_data.get(rf)]
        is_valid = len(missing) == 0

        students.append({
            "row": row_num,
            "valid": is_valid,
            "missing_fields": missing,
            "data": student_data,
        })

    valid_count = sum(1 for s in students if s["valid"])

    return {
        "success": True,
        "headers_found": raw_headers,
        "mapped_columns": {raw_headers[i]: canonical for i, canonical in column_mapping.items() if i < len(raw_headers)},
        "missing_required_columns": [rf for rf in required_fields if rf not in mapped_canonical and not defaults.get(rf)],
        "students": students,
        "total": len(students),
        "valid_count": valid_count,
    }


def generate_csv_template() -> str:
    """Generate a CSV template prefilled with column headers and an example row."""
    headers = [f[0] for f in STUDENT_FIELDS]
    example = [f[3] for f in STUDENT_FIELDS]
    out = io.StringIO()
    writer = csv.writer(out)
    writer.writerow(headers)
    writer.writerow(example)
    return out.getvalue()


def generate_xlsx_template() -> bytes:
    """Generate a clean .xlsx workbook template with headers and example data in pure standard library."""
    headers = [f[0] for f in STUDENT_FIELDS]
    example = [f[3] for f in STUDENT_FIELDS]
    all_rows = [headers, example]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedString+xml"/>
</Types>""",
        )
        z.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>""",
        )
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="StudentUpload" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""",
        )

        strings: list[str] = []
        str_map: dict[str, int] = {}

        def get_sid(s: str) -> int:
            if s not in str_map:
                str_map[s] = len(strings)
                strings.append(s)
            return str_map[s]

        sheet_rows: list[str] = []
        for r_idx, row in enumerate(all_rows, 1):
            cell_xmls: list[str] = []
            for c_idx, val in enumerate(row):
                if val is not None and str(val) != "":
                    s_val = str(val)
                    sid = get_sid(s_val)
                    cell_ref = f"{index_to_col(c_idx)}{r_idx}"
                    cell_xmls.append(f'<c r="{cell_ref}" t="s"><v>{sid}</v></c>')
            sheet_rows.append(f'<row r="{r_idx}">{" ".join(cell_xmls)}</row>')

        sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    {" ".join(sheet_rows)}
  </sheetData>
</worksheet>"""
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

        escaped_strings = [
            s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") for s in strings
        ]
        sst_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(strings)}" uniqueCount="{len(strings)}">
  {"".join(f"<si><t>{s}</t></si>" for s in escaped_strings)}
</sst>"""
        z.writestr("xl/sharedStrings.xml", sst_xml)

    buf.seek(0)
    return buf.getvalue()

