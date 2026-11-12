"""Vercel Serverless Function entrypoint for Rice API & Student Bulk Upload."""

import sys
from pathlib import Path

# Add project root to sys.path so app and excel_utils can be imported
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app


class handler(app.Handler):
    """Vercel serverless request handler."""
    pass

