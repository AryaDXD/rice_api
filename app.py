"""Local Swagger UI and credential-safe proxy backed by Postman collections, with Excel bulk student upload."""

from __future__ import annotations

import csv
import io
import json
import os
import re
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import excel_utils

ROOT = Path(__file__).resolve().parent
COLLECTIONS = ROOT / "postman" / "archive" / "collections"
SOURCES = [
    ("LMS API's FOR SMS (UAT).postman_collection.json", "UAT"),
    ("LMS API's FOR SMS.postman_collection 1.json", "Production / development"),
    ("Add student on Rice Edu.postman_collection 1.json", "Rice Edu"),
]
# These requests cover student, batch, batch-group, and defaulter workflows.
# Credential, password, and role administration endpoints are deliberately excluded.
ALLOWED = re.compile(r"student|batch|defaulter", re.I)
BLOCKED = re.compile(r"password|role|cco|sro|campaign", re.I)
STUDENT_REQUEST_EXAMPLE_PATH = ROOT / "postman" / "secrets" / "student_request_example.json"


def load_student_request_example():
    try:
        example = json.loads(STUDENT_REQUEST_EXAMPLE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return example if isinstance(example, dict) and isinstance(example.get("data"), dict) else None


def walk(items):
    for item in items or []:
        if "item" in item:
            yield from walk(item["item"])
        elif "request" in item:
            yield item


def resolve_url(url):
    if isinstance(url, str):
        return url
    return url.get("raw", "")


def blank_sample(value):
    """Keep the collection's JSON shape while dropping all sample/PII values."""
    if isinstance(value, dict):
        return {key: blank_sample(child) for key, child in value.items()}
    if isinstance(value, list):
        return [blank_sample(value[0])] if value else []
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return 0
    if value is None:
        return None
    return ""


def auth_header(request, collection):
    auth = request.get("auth") or collection.get("auth") or {}
    if auth.get("type") != "apikey":
        return None
    fields = {field.get("key"): field.get("value", "") for field in auth.get("apikey", [])}
    name = fields.get("key")
    if not name:
        return None
    val = (fields.get("value") or "").strip()
    # The collection request is the source of truth, just as it is in Postman.
    # Use the local key file when the request has no credential or has a placeholder.
    key_file = ROOT / "postman" / "secrets" / "api_key.txt"
    key = ""
    if key_file.exists():
        content = key_file.read_text(encoding="utf-8").strip()
        key = content.split(":", 1)[1].strip() if ":" in content else content
    if not val or val.upper() == "AUTHORIZATION" or val.startswith("{{"):
        val = key
    return (name, val or key)


def catalog():
    result = []
    for filename, environment in SOURCES:
        path = COLLECTIONS / filename
        if not path.exists():
            continue
        collection = json.loads(path.read_text(encoding="utf-8"))
        for item in walk(collection.get("item", [])):
            request = item["request"]
            method = request.get("method", "GET").upper()
            name = item.get("name", "API request")
            if method not in {"GET", "POST"} or not ALLOWED.search(name) or BLOCKED.search(name):
                continue
            url = resolve_url(request.get("url", {}))
            parsed = urlparse(url)
            if not parsed.hostname or parsed.scheme not in {"http", "https"}:
                continue

            item_env = environment
            item_name = name
            # Explicitly mark working production server
            if url == "https://lms.ricesmart.in/api/v1/rice/student":
                item_env = "Production"
                item_name = "Add Student SMS To LMS (Production)"
            elif url in {"http://lms.riceedu.in/api/v1/rice/student", "https://lms.riceedu.in/api/v1/rice/student"}:
                url = "https://lms.riceedu.in/api/v1/rice/student"
                item_env = "Rice Edu"
                item_name = "Add Student SMS To LMS (Rice Edu Production)"

            body = None
            raw = (request.get("body") or {}).get("raw")
            if method == "POST" and raw:
                try:
                    body = blank_sample(json.loads(raw))
                except (ValueError, TypeError):
                    body = {}
            if method == "POST" and "student" in url.lower() and url.rstrip("/").endswith("/api/v1/rice/student") and "ricesmart" in url.lower():
                body = load_student_request_example() or body
            result.append({
                "id": f"{len(result)}",
                "name": item_name,
                "environment": item_env,
                "method": method,
                "url": url,
                "body": body,
                "auth": auth_header(request, collection),
            })

    key_file = ROOT / "postman" / "secrets" / "api_key.txt"
    key = ""
    if key_file.exists():
        content = key_file.read_text(encoding="utf-8").strip()
        key = content.split(":", 1)[1].strip() if ":" in content else content
    shared_auth = ("AUTHORIZATION", key or "3c643f2d8c624292ae190df0beb644d7")

    # Ensure Rice Smart testing server (UAT) is present in catalog
    uat_student_url = "https://lms-uat.ricesmart.in/api/v1/rice/student"
    if not any(item["url"] == uat_student_url and item["method"] == "POST" for item in result):
        result.insert(0, {
            "id": f"{len(result)}",
            "name": "Add Student SMS To LMS (Testing / UAT)",
            "environment": "UAT",
            "method": "POST",
            "url": uat_student_url,
            "body": load_student_request_example() or {},
            "auth": shared_auth,
        })

    # Ensure Rice Edu testing server (UAT) is present in catalog
    riceedu_uat_student_url = "https://lms-uat.riceedu.in/api/v1/rice/student"
    if not any(item["url"] == riceedu_uat_student_url and item["method"] == "POST" for item in result):
        student_template = load_student_request_example() or {
            "data": {
                "firstName": "",
                "middleName": "",
                "lastName": "",
                "secondaryEmail": "",
                "contactNo": "",
                "address": "",
                "batchUniqueId": "",
                "studentUniqueId": "",
                "primaryEmail": "",
                "secondaryLangauge": "",
                "classStartDate": "",
            }
        }
        result.append({
            "id": f"{len(result)}",
            "name": "Add Student SMS To LMS (Rice Edu Testing / UAT)",
            "environment": "Rice Edu",
            "method": "POST",
            "url": riceedu_uat_student_url,
            "body": student_template,
            "auth": shared_auth,
        })

    # Re-index ids sequentially
    for idx, item in enumerate(result):
        item["id"] = str(idx)

    return result


def json_schema(sample):
    if isinstance(sample, dict):
        return {"type": "object", "properties": {key: json_schema(value) for key, value in sample.items()}}
    if isinstance(sample, list):
        return {"type": "array", "items": json_schema(sample[0]) if sample else {}}
    if isinstance(sample, bool):
        return {"type": "boolean"}
    if isinstance(sample, (int, float)):
        return {"type": "integer" if isinstance(sample, int) else "number"}
    if sample is None:
        return {"type": "string", "nullable": True}
    return {"type": "string"}


def openapi_document():
    paths = {}
    for api in catalog():
        operation = {
            "operationId": f"request_{api['id']}",
            "summary": api["name"],
            "description": f"Environment: {api['environment']}\nTarget: {api['url']}\nCredentials are applied by the local proxy.",
            "tags": [api["environment"]],
            "responses": {
                "default": {"description": "Response returned by the selected API."}
            },
        }
        if api["method"] == "POST":
            operation["requestBody"] = {
                "required": True,
                "content": {
                    "application/json": {
                        "schema": json_schema(api["body"] or {}),
                        "example": api["body"] or {},
                    }
                },
            }
        paths[f"/{api['id']}"] = {api["method"].lower(): operation}

    # Bulk Student Upload endpoints in OpenAPI schema
    paths["/api/bulk-upload/endpoints"] = {
        "get": {
            "summary": "List student creation endpoints",
            "description": "Returns candidate student endpoints available for bulk uploads.",
            "tags": ["Bulk Student Upload"],
            "responses": {"200": {"description": "List of available student endpoints."}},
        }
    }
    paths["/api/bulk-upload/template.xlsx"] = {
        "get": {
            "summary": "Download Excel (.xlsx) student upload template",
            "description": "Returns pre-formatted Excel workbook template with required student column names.",
            "tags": ["Bulk Student Upload"],
            "responses": {"200": {"description": "Binary .xlsx template file."}},
        }
    }
    paths["/api/bulk-upload/template.csv"] = {
        "get": {
            "summary": "Download CSV student upload template",
            "description": "Returns pre-formatted CSV template with required student column names.",
            "tags": ["Bulk Student Upload"],
            "responses": {"200": {"description": "Text .csv template file."}},
        }
    }
    paths["/api/bulk-upload/parse"] = {
        "post": {
            "summary": "Parse and validate student spreadsheet",
            "description": "Uploads an Excel (.xlsx) or CSV file and validates student columns and records.",
            "tags": ["Bulk Student Upload"],
            "responses": {"200": {"description": "Parsed and validated student records."}},
        }
    }
    paths["/api/bulk-upload/submit"] = {
        "post": {
            "summary": "Execute bulk student upload",
            "description": "Uploads an array of students to the selected LMS endpoint using the credential-safe proxy.",
            "tags": ["Bulk Student Upload"],
            "requestBody": {
                "required": True,
                "content": {
                    "application/json": {
                        "schema": {
                            "type": "object",
                            "properties": {
                                "endpoint_id": {"type": "string"},
                                "students": {"type": "array", "items": {"type": "object"}},
                            },
                            "required": ["endpoint_id", "students"],
                        }
                    }
                },
            },
            "responses": {"200": {"description": "Detailed upload execution report."}},
        }
    }

    env_tags = [{"name": name} for name in dict.fromkeys(api["environment"] for api in catalog())]
    tags = [{"name": "Bulk Student Upload"}] + env_tags

    return {
        "openapi": "3.0.3",
        "info": {
            "title": "Student Management API Tester & Bulk Upload",
            "version": "1.1.0",
            "description": "GET and POST requests imported from Postman collections with Excel/CSV bulk student upload.",
        },
        "servers": [{"url": "/api/proxy", "description": "Local credential-safe proxy"}],
        "tags": tags,
        "paths": paths,
    }


PAGE = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Student Management · API Tester & Bulk Upload</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/swagger-ui-dist@5/swagger-ui.css">
<style>
:root{
  --primary:#182b4d;
  --accent:#2563eb;
  --accent-hover:#1d4ed8;
  --bg:#f5f7fb;
  --card-bg:#ffffff;
  --border:#e2e8f0;
  --text-main:#1e293b;
  --text-muted:#64748b;
  --success:#16a34a;
  --success-bg:#dcfce7;
  --danger:#dc2626;
  --danger-bg:#fee2e2;
  --warn:#d97706;
  --warn-bg:#fef3c7;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text-main);font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}
.top{padding:20px max(24px,calc((100% - 1140px)/2));background:#fff;border-bottom:1px solid var(--border)}
.header-inner{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px}
.eyebrow{color:#5268ad;font:700 12px/1.4 system-ui,sans-serif;letter-spacing:.1em;text-transform:uppercase}
.top h1{font:700 24px/1.2 system-ui,sans-serif;color:var(--primary);margin:4px 0 6px}
.top p{font:14px/1.5 system-ui,sans-serif;color:var(--text-muted);margin:0}

/* Nav Tabs */
.nav-tabs{display:flex;background:#eef2f8;padding:4px;border-radius:10px;gap:4px}
.nav-tab{border:none;background:transparent;padding:9px 18px;border-radius:7px;font:600 13px system-ui,sans-serif;color:var(--text-muted);cursor:pointer;display:inline-flex;align-items:center;gap:8px;transition:all .15s ease}
.nav-tab:hover{color:var(--primary)}
.nav-tab.active{background:#fff;color:var(--accent);box-shadow:0 1px 4px rgba(0,0,0,.08)}
.tab-badge{background:#e0e7ff;color:#4338ca;font-size:10px;font-weight:700;padding:2px 7px;border-radius:12px;letter-spacing:.02em}

.notice{max-width:1140px;margin:18px auto 0;padding:12px 16px;border-radius:8px;background:#fff8e7;color:#715512;font:13px/1.5 system-ui,sans-serif;border-left:4px solid #f59e0b}
.shell{max-width:1140px;margin:20px auto 40px;padding:0 16px}

/* Cards */
.card{background:var(--card-bg);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.03)}
.card-header{display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:16px;margin-bottom:20px}
.card-header h2{font-size:18px;margin:0 0 4px;color:var(--primary)}
.card-desc{font-size:13px;color:var(--text-muted);margin:0}
.template-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.template-label{font-size:12px;font-weight:600;color:var(--text-muted)}

/* Buttons */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;font:600 13px system-ui,sans-serif;padding:8px 16px;border-radius:7px;border:1px solid transparent;cursor:pointer;text-decoration:none;transition:all .15s}
.btn-primary{background:var(--accent);color:#fff}
.btn-primary:hover{background:var(--accent-hover)}
.btn-primary:disabled{background:#94a3b8;cursor:not-allowed}
.btn-outline{background:#fff;border-color:var(--border);color:var(--text-main)}
.btn-outline:hover{background:#f8fafc;border-color:#cbd5e1}
.btn-danger{background:var(--danger);color:#fff}
.btn-danger:hover{background:#b91c1c}

/* Form Grid */
.form-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px}
.form-group{display:flex;flex-direction:column;gap:6px}
.form-group.full-width{grid-column:1/-1}
.form-group label{font-size:13px;font-weight:600;color:var(--text-main)}
.form-control{padding:9px 12px;border:1px solid var(--border);border-radius:7px;font:13px system-ui,sans-serif;background:#fff;color:var(--text-main)}
.form-control:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px rgba(37,99,235,.15)}

/* Drop Zone */
.drop-zone{border:2px dashed #cbd5e1;border-radius:10px;padding:32px 20px;text-align:center;background:#f8fafc;cursor:pointer;transition:all .2s ease}
.drop-zone:hover,.drop-zone.dragover{border-color:var(--accent);background:#eff6ff}
.drop-zone-icon{margin-bottom:8px}
.drop-zone-text{font-size:14px;color:var(--text-main);margin-bottom:4px}
.drop-zone-sub{font-size:12px;color:var(--text-muted)}

/* Stats Pill */
.card-stats{display:flex;gap:8px;flex-wrap:wrap}
.stat-pill{font-size:12px;font-weight:700;padding:4px 10px;border-radius:20px;display:inline-flex;align-items:center;gap:4px}
.stat-total{background:#f1f5f9;color:#334155}
.stat-valid{background:var(--success-bg);color:var(--success)}
.stat-warn{background:var(--danger-bg);color:var(--danger)}

/* Column Chips */
.column-chips{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px;padding:10px 12px;background:#f8fafc;border-radius:8px;border:1px solid var(--border)}
.col-chip{font-size:11px;padding:3px 8px;border-radius:6px;background:#fff;border:1px solid var(--border);color:var(--text-muted);display:inline-flex;align-items:center;gap:4px}
.col-chip.mapped{border-color:#bbf7d0;background:#f0fdf4;color:#15803d}
.col-chip.missing{border-color:#fecaca;background:#fef2f2;color:#b91c1c}

/* Table */
.table-container{max-height:420px;overflow:auto;border:1px solid var(--border);border-radius:8px;margin-bottom:16px}
.preview-table{width:100%;border-collapse:collapse;font-size:12px;text-align:left}
.preview-table th{position:sticky;top:0;background:#f8fafc;padding:9px 12px;border-bottom:1px solid var(--border);font-weight:700;color:#475569;z-index:1}
.preview-table td{padding:8px 12px;border-bottom:1px solid #edf2f7;color:#334155;white-space:nowrap}
.preview-table tr:hover td{background:#f8fafc}
.preview-table tr.row-invalid td{background:#fff1f2}
.status-tag{font-size:11px;font-weight:700;padding:2px 8px;border-radius:12px;display:inline-flex;align-items:center;gap:4px}
.tag-ready{background:#f1f5f9;color:#475569}
.tag-sending{background:#dbeafe;color:#1e40af;animation:pulse 1.2s infinite}
.tag-success{background:var(--success-bg);color:var(--success)}
.tag-failed{background:var(--danger-bg);color:var(--danger)}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}

/* Progress */
.progress-section{background:#f8fafc;border:1px solid var(--border);border-radius:8px;padding:12px 16px;margin-bottom:16px}
.progress-header{display:flex;justify-content:space-between;font-size:12px;font-weight:600;margin-bottom:6px;color:#334155}
.progress-bar-bg{height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden}
.progress-bar-fill{height:100%;background:var(--accent);border-radius:4px;transition:width .2s ease}

/* Completion Banner */
.completion-banner{padding:12px 16px;border-radius:8px;font-size:13px;font-weight:600;margin-bottom:16px;display:flex;align-items:center;justify-content:space-between}
.banner-success{background:var(--success-bg);color:#166534;border:1px solid #bbf7d0}
.banner-mixed{background:var(--warn-bg);color:#92400e;border:1px solid #fde68a}

.card-actions{display:flex;align-items:center;gap:10px;flex-wrap:wrap}

/* Swagger override */
.swagger-ui{max-width:1160px;margin:12px auto 36px}.swagger-ui .topbar{display:none}
@media(max-width:700px){.top{padding:16px}.header-inner{flex-direction:column;align-items:flex-start}}
</style>
</head>
<body>

<header class="top">
  <div class="header-inner">
    <div>
      <div class="eyebrow">Student Management · LMS Tools</div>
      <h1>Student Management & API Tester</h1>
      <p>Test individual student endpoints or upload students in bulk from Excel or CSV spreadsheets.</p>
    </div>
    <div class="nav-tabs">
      <button class="nav-tab active" id="btn-swagger" onclick="switchView('swagger')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1-2.5-2.5Z"/><path d="M6 6h10M6 10h10"/></svg>
        API Tester (Swagger)
      </button>
      <button class="nav-tab" id="btn-bulk" onclick="switchView('bulk')">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
        Bulk Student Upload
        <span class="tab-badge">Excel / CSV</span>
      </button>
    </div>
  </div>
</header>

<!-- View 1: Swagger UI -->
<div id="view-swagger">
  <div class="notice">Requests are sent through this local app using credentials from the Postman setup. API calls can create or update real records; choose the intended environment and use test data.</div>
  <div id="swagger-ui"></div>
</div>

<!-- View 2: Bulk Student Upload -->
<div id="view-bulk" style="display:none;">
  <div class="shell">
    <div class="notice">Upload student records in bulk from an Excel (.xlsx) or CSV file. Columns are mapped automatically to student schema fields.</div>

    <!-- Step 1 Card -->
    <div class="card">
      <div class="card-header">
        <div>
          <h2>1. Select Destination & Upload Spreadsheet</h2>
          <p class="card-desc">Choose the destination student API endpoint and provide your Excel or CSV file.</p>
        </div>
        <div class="template-actions">
          <span class="template-label">Download Template:</span>
          <a href="/api/bulk-upload/template.xlsx" class="btn btn-outline" download="student_upload_template.xlsx">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>
            Excel Template (.xlsx)
          </a>
          <a href="/api/bulk-upload/template.csv" class="btn btn-outline" download="student_upload_template.csv">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg>
            CSV Template (.csv)
          </a>
        </div>
      </div>

      <div class="form-grid">
        <div class="form-group full-width">
          <label for="bulk-endpoint">Target Student API Endpoint</label>
          <select id="bulk-endpoint" class="form-control">
            <option value="">Loading student endpoints...</option>
          </select>
        </div>

        <div class="form-group full-width">
          <label>Spreadsheet File (.xlsx, .csv)</label>
          <div class="drop-zone" id="drop-zone">
            <input type="file" id="file-input" accept=".xlsx,.xls,.csv" style="display:none" onchange="onFileChosen(this.files)">
            <div class="drop-zone-icon">
              <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="1.8"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"/></svg>
            </div>
            <div class="drop-zone-text" id="drop-zone-text">
              <strong>Click to upload</strong> or drag and drop your spreadsheet here
            </div>
            <div class="drop-zone-sub">Required columns: firstName, lastName, primaryEmail, contactNo, studentUniqueId, batchUniqueId.</div>
          </div>
        </div>

        <div class="form-group">
          <label for="def-batch">Default Batch ID (Optional)</label>
          <input type="text" id="def-batch" class="form-control" placeholder="e.g. RSMSMPSC2627001 (applied if row is empty)">
        </div>

        <div class="form-group">
          <label for="def-date">Default Class Start Date (Optional)</label>
          <input type="text" id="def-date" class="form-control" placeholder="e.g. 11/6/2026">
        </div>

        <div class="form-group">
          <label for="def-lang">Default Secondary Language (Optional)</label>
          <input type="text" id="def-lang" class="form-control" placeholder="e.g. hindi">
        </div>
      </div>
    </div>

    <!-- Step 2 Card: Preview & Validation -->
    <div class="card" id="preview-card" style="display:none;">
      <div class="card-header">
        <div>
          <h2>2. Preview & Validate Students</h2>
          <p class="card-desc" id="preview-desc">Review records and mapped columns before initiating upload.</p>
        </div>
        <div class="card-stats">
          <div class="stat-pill stat-total"><span id="stat-total">0</span> Total</div>
          <div class="stat-pill stat-valid"><span id="stat-valid">0</span> Ready</div>
          <div class="stat-pill stat-warn" id="pill-invalid" style="display:none;"><span id="stat-invalid">0</span> Incomplete</div>
        </div>
      </div>

      <div class="column-chips" id="column-chips"></div>

      <!-- Live progress bar -->
      <div id="upload-progress-container" style="display:none;" class="progress-section">
        <div class="progress-header">
          <span id="progress-status-text">Uploading students...</span>
          <span id="progress-percent-text">0%</span>
        </div>
        <div class="progress-bar-bg">
          <div class="progress-bar-fill" id="progress-bar-fill" style="width:0%"></div>
        </div>
      </div>

      <!-- Completion summary banner -->
      <div id="completion-banner" style="display:none;" class="completion-banner"></div>

      <div class="table-container">
        <table class="preview-table">
          <thead>
            <tr>
              <th width="36"><input type="checkbox" id="select-all" checked onchange="toggleSelectAll(this.checked)"></th>
              <th width="44">#</th>
              <th>Student Name</th>
              <th>Student ID</th>
              <th>Batch ID</th>
              <th>Email</th>
              <th>Contact</th>
              <th>Start Date</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody id="students-tbody"></tbody>
        </table>
      </div>

      <div class="card-actions">
        <button id="btn-start-upload" class="btn btn-primary" onclick="startBulkUpload()">
          🚀 Upload Selected Students (<span id="btn-upload-count">0</span>)
        </button>
        <button id="btn-cancel-upload" class="btn btn-danger" style="display:none;" onclick="cancelUpload()">
          ⏹ Cancel Upload
        </button>
        <button id="btn-export-results" class="btn btn-outline" style="display:none;" onclick="exportResultsCsv()">
          📥 Export Results Report (.csv)
        </button>
        <button id="btn-clear" class="btn btn-outline" onclick="resetBulkUpload()">
          ✖ Upload Another File
        </button>
      </div>
    </div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/swagger-ui-dist@5/swagger-ui-bundle.js"></script>
<script>
// --- Swagger Setup ---
const swaggerRoot = document.getElementById('swagger-ui');
const disableSpellcheck = () => swaggerRoot.querySelectorAll('textarea,[contenteditable="true"]').forEach(e => e.setAttribute('spellcheck', 'false'));
new MutationObserver(disableSpellcheck).observe(swaggerRoot, { childList: true, subtree: true });
window.ui = SwaggerUIBundle({
  url: '/openapi.json',
  dom_id: '#swagger-ui',
  deepLinking: true,
  displayRequestDuration: true,
  filter: true,
  tryItOutEnabled: true,
  docExpansion: 'list',
  defaultModelsExpandDepth: 1
});
disableSpellcheck();

// --- View Switcher ---
function switchView(tab) {
  document.getElementById('view-swagger').style.display = tab === 'swagger' ? 'block' : 'none';
  document.getElementById('view-bulk').style.display = tab === 'bulk' ? 'block' : 'none';
  document.getElementById('btn-swagger').classList.toggle('active', tab === 'swagger');
  document.getElementById('btn-bulk').classList.toggle('active', tab === 'bulk');
  if (tab === 'bulk') {
    loadEndpoints();
  }
}

// --- Bulk Upload State ---
let parsedStudents = [];
let uploadResults = [];
let isUploading = false;
let currentFile = null;

// --- Load Endpoints ---
async function loadEndpoints() {
  const sel = document.getElementById('bulk-endpoint');
  if (sel.options.length > 1 && sel.options[0].value !== '') return;
  try {
    const res = await fetch('/api/bulk-upload/endpoints');
    const list = await res.json();
    sel.innerHTML = '';
    list.forEach((ep, idx) => {
      const opt = document.createElement('option');
      opt.value = ep.id;
      opt.textContent = `[${ep.environment}] ${ep.name} (${ep.url})`;
      if (idx === 0) opt.selected = true;
      sel.appendChild(opt);
    });
  } catch (err) {
    sel.innerHTML = '<option value="">Failed to load endpoints</option>';
  }
}

// --- Drag & Drop ---
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  if (e.dataTransfer.files && e.dataTransfer.files.length) {
    onFileChosen(e.dataTransfer.files);
  }
});

function onFileChosen(files) {
  if (!files || !files.length) return;
  currentFile = files[0];
  document.getElementById('drop-zone-text').innerHTML = `Selected file: <strong>${escapeHtml(currentFile.name)}</strong> (${(currentFile.size / 1024).toFixed(1)} KB)`;
  parseCurrentFile();
}

// --- Re-parse when defaults change ---
['def-batch', 'def-date', 'def-lang'].forEach(id => {
  document.getElementById(id).addEventListener('change', () => {
    if (currentFile && !isUploading) parseCurrentFile();
  });
});

async function parseCurrentFile() {
  if (!currentFile) return;
  const formData = new FormData();
  formData.append('file', currentFile);
  formData.append('batchUniqueId', document.getElementById('def-batch').value.trim());
  formData.append('classStartDate', document.getElementById('def-date').value.trim());
  formData.append('secondaryLangauge', document.getElementById('def-lang').value.trim());

  document.getElementById('drop-zone-text').innerHTML = `Parsing <strong>${escapeHtml(currentFile.name)}</strong>...`;

  try {
    const res = await fetch('/api/bulk-upload/parse', { method: 'POST', body: formData });
    const data = await res.json();
    if (!data.success) {
      alert('Error parsing spreadsheet: ' + (data.error || 'Unknown error'));
      document.getElementById('drop-zone-text').innerHTML = `<strong>Click to upload</strong> or drag and drop your spreadsheet here`;
      return;
    }
    renderPreview(data);
  } catch (err) {
    alert('Upload/Parse failed: ' + err.message);
  }
}

function renderPreview(data) {
  document.getElementById('preview-card').style.display = 'block';
  document.getElementById('completion-banner').style.display = 'none';
  document.getElementById('upload-progress-container').style.display = 'none';
  document.getElementById('btn-export-results').style.display = 'none';

  parsedStudents = (data.students || []).map((s, idx) => ({
    ...s,
    idx: idx,
    selected: s.valid,
    status: s.valid ? 'ready' : 'incomplete',
    statusMsg: s.valid ? 'Ready' : `Missing: ${s.missing_fields.join(', ')}`,
    response: null
  }));

  // Update counters
  document.getElementById('stat-total').textContent = data.total;
  document.getElementById('stat-valid').textContent = data.valid_count;
  const invalidCount = data.total - data.valid_count;
  const pillInvalid = document.getElementById('pill-invalid');
  if (invalidCount > 0) {
    pillInvalid.style.display = 'inline-flex';
    document.getElementById('stat-invalid').textContent = invalidCount;
  } else {
    pillInvalid.style.display = 'none';
  }

  // Column chips
  const chipsContainer = document.getElementById('column-chips');
  chipsContainer.innerHTML = '';
  const mapped = data.mapped_columns || {};
  Object.keys(mapped).forEach(orig => {
    const chip = document.createElement('span');
    chip.className = 'col-chip mapped';
    chip.textContent = `${orig} → ${mapped[orig]}`;
    chipsContainer.appendChild(chip);
  });
  (data.missing_required_columns || []).forEach(req => {
    const chip = document.createElement('span');
    chip.className = 'col-chip missing';
    chip.textContent = `Missing Required: ${req}`;
    chipsContainer.appendChild(chip);
  });

  renderTableRows();
  updateSelectedCount();
}

function renderTableRows() {
  const tbody = document.getElementById('students-tbody');
  tbody.innerHTML = '';
  parsedStudents.forEach((st, i) => {
    const tr = document.createElement('tr');
    tr.id = `row-${i}`;
    if (!st.valid) tr.className = 'row-invalid';

    const d = st.data || {};
    const fullName = `${d.firstName || ''} ${d.middleName || ''} ${d.lastName || ''}`.replace(/\s+/g, ' ').trim();

    tr.innerHTML = `
      <td><input type="checkbox" ${st.selected ? 'checked' : ''} ${!st.valid ? 'disabled' : ''} onchange="toggleSelectRow(${i}, this.checked)"></td>
      <td>${st.row}</td>
      <td><strong>${escapeHtml(fullName || '—')}</strong></td>
      <td><code>${escapeHtml(d.studentUniqueId || '—')}</code></td>
      <td><code>${escapeHtml(d.batchUniqueId || '—')}</code></td>
      <td>${escapeHtml(d.primaryEmail || '—')}</td>
      <td>${escapeHtml(d.contactNo || '—')}</td>
      <td>${escapeHtml(d.classStartDate || '—')}</td>
      <td id="status-cell-${i}">
        <span class="status-tag ${st.valid ? 'tag-ready' : 'tag-failed'}">${escapeHtml(st.statusMsg)}</span>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function toggleSelectRow(idx, checked) {
  parsedStudents[idx].selected = checked;
  updateSelectedCount();
}

function toggleSelectAll(checked) {
  parsedStudents.forEach(st => {
    if (st.valid) st.selected = checked;
  });
  renderTableRows();
  updateSelectedCount();
}

function updateSelectedCount() {
  const count = parsedStudents.filter(s => s.selected && s.valid).length;
  document.getElementById('btn-upload-count').textContent = count;
  document.getElementById('btn-start-upload').disabled = count === 0 || isUploading;
}

// --- Execution ---
async function startBulkUpload() {
  const endpointId = document.getElementById('bulk-endpoint').value;
  if (!endpointId) {
    alert('Please select a target student API endpoint.');
    return;
  }
  const toUpload = parsedStudents.filter(s => s.selected && s.valid);
  if (!toUpload.length) {
    alert('No valid students selected.');
    return;
  }

  isUploading = true;
  document.getElementById('btn-start-upload').style.display = 'none';
  document.getElementById('btn-cancel-upload').style.display = 'inline-flex';
  document.getElementById('upload-progress-container').style.display = 'block';
  document.getElementById('completion-banner').style.display = 'none';

  uploadResults = [];
  let successful = 0, failed = 0;

  for (let i = 0; i < toUpload.length; i++) {
    if (!isUploading) break;
    const st = toUpload[i];
    const rowIdx = st.idx;

    // Update row status
    setRowState(rowIdx, 'sending', 'Uploading...');

    try {
      const res = await fetch(`/api/proxy/${endpointId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: st.data })
      });

      const text = await res.text();
      let respData;
      try { respData = JSON.parse(text); } catch(e) { respData = text; }

      const isOk = res.status >= 200 && res.status < 300;
      if (isOk) {
        successful++;
        setRowState(rowIdx, 'success', `✓ Success (${res.status})`, respData);
      } else {
        failed++;
        const errMsg = typeof respData === 'object' && respData ? (respData.error || respData.message || respData.detail || JSON.stringify(respData)) : respData;
        setRowState(rowIdx, 'failed', `✗ Failed (${res.status}): ${errMsg}`, respData);
      }
      uploadResults.push({ ...st.data, row: st.row, status: res.status, success: isOk, error: isOk ? '' : text });
    } catch (err) {
      failed++;
      setRowState(rowIdx, 'failed', `✗ Network error: ${err.message}`, err.message);
      uploadResults.push({ ...st.data, row: st.row, status: 0, success: false, error: err.message });
    }

    const pct = Math.round(((i + 1) / toUpload.length) * 100);
    document.getElementById('progress-bar-fill').style.width = pct + '%';
    document.getElementById('progress-percent-text').textContent = pct + '%';
    document.getElementById('progress-status-text').textContent = `Uploaded ${i + 1} of ${toUpload.length} students (${successful} successful, ${failed} failed)`;
  }

  isUploading = false;
  document.getElementById('btn-start-upload').style.display = 'inline-flex';
  document.getElementById('btn-cancel-upload').style.display = 'none';
  document.getElementById('btn-export-results').style.display = 'inline-flex';

  // Completion banner
  const banner = document.getElementById('completion-banner');
  banner.style.display = 'flex';
  if (failed === 0) {
    banner.className = 'completion-banner banner-success';
    banner.textContent = `🎉 All ${successful} student records were uploaded successfully!`;
  } else {
    banner.className = 'completion-banner banner-mixed';
    banner.textContent = `Upload complete: ${successful} succeeded, ${failed} failed. Click "Export Results Report" below for error details.`;
  }
}

function cancelUpload() {
  isUploading = false;
}

function setRowState(idx, type, text, detail) {
  const cell = document.getElementById(`status-cell-${idx}`);
  if (!cell) return;
  const tagClass = type === 'sending' ? 'tag-sending' : (type === 'success' ? 'tag-success' : 'tag-failed');
  cell.innerHTML = `<span class="status-tag ${tagClass}" title="${escapeHtml(JSON.stringify(detail || ''))}">${escapeHtml(text)}</span>`;
}

function resetBulkUpload() {
  currentFile = null;
  parsedStudents = [];
  uploadResults = [];
  isUploading = false;
  fileInput.value = '';
  document.getElementById('drop-zone-text').innerHTML = `<strong>Click to upload</strong> or drag and drop your spreadsheet here`;
  document.getElementById('preview-card').style.display = 'none';
}

function exportResultsCsv() {
  if (!uploadResults.length) return;
  const headers = ['Row', 'Student ID', 'Batch ID', 'First Name', 'Last Name', 'Email', 'Contact', 'Status', 'Success', 'Server Response'];
  const rows = uploadResults.map(r => [
    r.row || '',
    r.studentUniqueId || '',
    r.batchUniqueId || '',
    r.firstName || '',
    r.lastName || '',
    r.primaryEmail || '',
    r.contactNo || '',
    r.status || '',
    r.success ? 'TRUE' : 'FALSE',
    typeof r.error === 'string' ? r.error : JSON.stringify(r.error || '')
  ]);

  let csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(','))].join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', 'student_bulk_upload_report.csv');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
</script>
</body>
</html>'''


def extract_upload(headers, body_bytes):
    ctype = headers.get("Content-Type", "")
    if "multipart/form-data" in ctype:
        msg_bytes = f"Content-Type: {ctype}\r\n\r\n".encode() + body_bytes
        msg = BytesParser(policy=default).parsebytes(msg_bytes)
        fields = {}
        files = {}
        for part in msg.iter_parts():
            name = part.get_param("name", header="content-disposition")
            filename = part.get_filename()
            if filename:
                files[name or "file"] = {
                    "filename": filename,
                    "data": part.get_payload(decode=True) or b"",
                }
            else:
                fields[name or ""] = (part.get_payload(decode=True) or b"").decode("utf-8", errors="replace")
        return fields, files
    return {}, {"file": {"filename": "upload.xlsx", "data": body_bytes}}


def proxy_request(api, payload=None, query=""):
    headers = {"Accept": "application/json", "User-Agent": "StudentApiTester/1.0"}
    token = api.get("auth")
    if token and token[1]:
        headers[token[0]] = token[1]
    data = None
    if api["method"] == "POST":
        data = json.dumps(payload if payload is not None else {}).encode()
        headers["Content-Type"] = "application/json"
    target_url = api["url"]
    if query:
        separator = "&" if "?" in target_url else "?"
        target_url = f"{target_url}{separator}{query}"
    req = Request(target_url, data=data, headers=headers, method=api["method"])
    try:
        response = urlopen(req, timeout=25)
    except HTTPError as err:
        response = err
    except (URLError, TimeoutError) as exc:
        return 502, "application/json; charset=utf-8", json.dumps({"error": str(getattr(exc, "reason", exc))}).encode()
    except Exception as exc:
        return 502, "application/json; charset=utf-8", json.dumps({"error": f"Proxy request failed: {exc}"}).encode()
    try:
        status = getattr(response, "status", getattr(response, "code", 200))
        content_type = response.headers.get("Content-Type", "application/octet-stream")
        content = response.read(2_000_000)
    finally:
        if hasattr(response, "close"):
            response.close()
    return status, content_type, content


class Handler(BaseHTTPRequestHandler):
    server_version = "StudentApiTester/1.1"

    @property
    def effective_path(self) -> str:
        # Support reverse proxies and Vercel serverless URL rewrites
        forwarded = (
            self.headers.get("x-vercel-original-url")
            or self.headers.get("x-forwarded-uri")
            or self.headers.get("x-original-uri")
        )
        return forwarded if forwarded else self.path

    def send_json(self, data, status=200):
        encoded = json.dumps(data).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(encoded)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    def do_GET(self):
        parsed_url = urlparse(self.effective_path)
        path = parsed_url.path
        if path == "/":
            body = PAGE.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif path == "/openapi.json":
            self.send_json(openapi_document())
        elif path == "/api/bulk-upload/endpoints":
            endpoints = [
                {
                    "id": item["id"],
                    "name": item["name"],
                    "environment": item["environment"],
                    "url": item["url"],
                }
                for item in catalog()
                if item["method"] == "POST" and item["url"].rstrip("/").endswith("/api/v1/rice/student")
            ]
            self.send_json(endpoints)
        elif path in {"/api/bulk-upload/template.xlsx", "/api/bulk-upload/template"}:
            xlsx_bytes = excel_utils.generate_xlsx_template()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="student_upload_template.xlsx"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(xlsx_bytes)
        elif path == "/api/bulk-upload/template.csv":
            csv_bytes = excel_utils.generate_csv_template().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="student_upload_template.csv"')
            self.send_header("Content-Length", str(len(csv_bytes)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(csv_bytes)
        elif path.startswith("/api/proxy/"):
            self.run_proxy("GET")
        else:
            self.send_error(404)

    def do_POST(self):
        parsed_url = urlparse(self.effective_path)
        path = parsed_url.path
        if path == "/api/bulk-upload/parse":
            self.handle_bulk_parse()
        elif path == "/api/bulk-upload/submit":
            self.handle_bulk_submit()
        elif path.startswith("/api/proxy/"):
            self.run_proxy("POST")
        else:
            self.send_error(404)

    def handle_bulk_parse(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(min(length, 50_000_000))
            fields, files = extract_upload(self.headers, body)
            file_info = files.get("file")
            if not file_info or not file_info["data"]:
                self.send_json({"success": False, "error": "No file uploaded. Please upload an Excel (.xlsx) or CSV file."}, 400)
                return

            defaults = {
                "batchUniqueId": fields.get("batchUniqueId", "").strip(),
                "classStartDate": fields.get("classStartDate", "").strip(),
                "secondaryLangauge": fields.get("secondaryLangauge", fields.get("secondaryLanguage", "")).strip(),
            }
            rows = excel_utils.parse_spreadsheet(file_info["data"], file_info["filename"])
            res = excel_utils.process_student_records(rows, defaults)
            res["filename"] = file_info["filename"]
            self.send_json(res)
        except Exception as exc:
            self.send_json({"success": False, "error": f"Failed to parse file: {exc}"}, 500)

    def handle_bulk_submit(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(min(length, 50_000_000))
            ctype = self.headers.get("Content-Type", "")

            endpoint_id = ""
            students = []

            if "application/json" in ctype:
                payload = json.loads(body)
                endpoint_id = str(payload.get("endpoint_id", "")).strip()
                students = payload.get("students", [])
            elif "multipart/form-data" in ctype:
                fields, files = extract_upload(self.headers, body)
                endpoint_id = fields.get("endpoint_id", "").strip()
                file_info = files.get("file")
                if file_info and file_info["data"]:
                    defaults = {
                        "batchUniqueId": fields.get("batchUniqueId", "").strip(),
                        "classStartDate": fields.get("classStartDate", "").strip(),
                        "secondaryLangauge": fields.get("secondaryLangauge", fields.get("secondaryLanguage", "")).strip(),
                    }
                    rows = excel_utils.parse_spreadsheet(file_info["data"], file_info["filename"])
                    processed = excel_utils.process_student_records(rows, defaults)
                    students = [s["data"] for s in processed.get("students", []) if s.get("valid")]
            else:
                self.send_json({"success": False, "error": "Unsupported Content-Type. Please send application/json or multipart/form-data."}, 400)
                return

            if not endpoint_id:
                self.send_json({"success": False, "error": "Missing endpoint_id for destination API."}, 400)
                return

            api = next((item for item in catalog() if item["id"] == endpoint_id), None)
            if not api:
                self.send_json({"success": False, "error": f"Endpoint ID '{endpoint_id}' was not found in catalog."}, 404)
                return

            if not students:
                self.send_json({"success": False, "error": "No student records provided to upload."}, 400)
                return

            results = []
            success_count = 0
            fail_count = 0

            for idx, student in enumerate(students, start=1):
                st_data = student.get("data", student) if isinstance(student, dict) else {}
                student_payload = {"data": st_data}
                status, content_type, content_bytes = proxy_request(api, student_payload)

                resp_obj = None
                resp_text = content_bytes.decode("utf-8", errors="replace")
                try:
                    resp_obj = json.loads(resp_text)
                except ValueError:
                    resp_obj = resp_text

                is_ok = 200 <= status < 300
                if is_ok:
                    success_count += 1
                else:
                    fail_count += 1

                full_name = f"{st_data.get('firstName', '')} {st_data.get('lastName', '')}".strip()
                results.append({
                    "index": idx,
                    "name": full_name or f"Student #{idx}",
                    "studentUniqueId": st_data.get("studentUniqueId", ""),
                    "batchUniqueId": st_data.get("batchUniqueId", ""),
                    "status": status,
                    "success": is_ok,
                    "response": resp_obj,
                })

            self.send_json({
                "success": fail_count == 0,
                "total": len(students),
                "successful": success_count,
                "failed": fail_count,
                "target_endpoint": {
                    "id": api["id"],
                    "name": api["name"],
                    "environment": api["environment"],
                    "url": api["url"],
                },
                "results": results,
            })
        except Exception as exc:
            self.send_json({"success": False, "error": f"Bulk submit error: {exc}"}, 500)

    def run_proxy(self, method):
        try:
            parts = self.effective_path.removeprefix("/api/proxy/").split("?", 1)
            request_id = parts[0]
            query = parts[1] if len(parts) > 1 else ""
            api = next((item for item in catalog() if item["id"] == request_id), None)
            if not api or api["method"] != method:
                self.send_error(404)
                return
            payload = None
            if method == "POST":
                length = int(self.headers.get("Content-Length", "0"))
                if length > 0:
                    payload = json.loads(self.rfile.read(min(length, 1_000_000)))
                else:
                    payload = {}
            status, content_type, content = proxy_request(api, payload, query)
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(content)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(content)
        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            self.send_json({"error": f"Invalid request: {exc}"}, 400)
        except Exception as exc:
            self.send_json({"error": f"Proxy error: {exc}"}, 500)

    def log_message(self, fmt, *args):
        # Avoid logging request payloads, which can contain student information.
        super().log_message("%s", args[0] if args else "request")


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


if __name__ == "__main__":
    host = os.environ.get("API_TESTER_HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", os.environ.get("API_TESTER_PORT", "8000")))
    print(f"Student API Tester & Bulk Upload running at http://{host}:{port}")
    try:
        ReusableThreadingHTTPServer((host, port), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down Student API Tester.")
