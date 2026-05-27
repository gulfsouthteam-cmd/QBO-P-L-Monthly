"""
QBO Profit & Loss (standard P&L) — parser + Flask endpoint.

Make.com POSTs the .xlsx as multipart/form-data (field name: 'file')
to /process, this returns JSON, Make writes a row to Google Sheets.
One QBO P&L export → one row in the sheet.
"""

from __future__ import annotations

import logging
import os
import re
from datetime import datetime
from io import BytesIO
from typing import Optional

from flask import Flask, jsonify, request
from openpyxl import load_workbook

log = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)

# Optional shared-secret auth. Set PIPELINE_API_KEY in Railway env vars.
# Make.com sends it in the 'X-Api-Key' header. Leave unset to skip auth.
API_KEY = os.environ.get("PIPELINE_API_KEY")


# ---------- Flask routes ----------

@app.route("/process", methods=["POST"])
def process():
    if API_KEY and request.headers.get("X-Api-Key") != API_KEY:
        return jsonify({"ok": False, "error": "unauthorized"}), 401

    upload = request.files.get("file")
    if upload is None:
        return jsonify({"ok": False, "error": "missing file"}), 400

    file_bytes = upload.read()
    if not file_bytes:
        return jsonify({"ok": False, "error": "empty file"}), 400

    try:
        parsed = parse(file_bytes)
    except ValueError as e:
        log.warning("parse failed: %s", e)
        return jsonify({"ok": False, "error": str(e)}), 422
    except Exception as e:  # noqa: BLE001
        log.exception("unexpected parse error")
        return jsonify({"ok": False, "error": f"parse failed: {e}"}), 500

    log.info("parsed report_date=%s fields=%d",
             parsed.get("report_date"), len(parsed) - 1)
    return jsonify(parsed), 200


@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "service": "pl-parser"}), 200


# ---------- Parser ----------

def parse(file_bytes: bytes) -> dict:
    wb = load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    report_date = _find_report_end_date(rows)
    if report_date is None:
        raise ValueError(
            "Could not find report date range in header. "
            "Is this actually a QBO Profit & Loss export?"
        )

    result: dict = {"report_date": report_date}

    for row in rows:
        if not row or len(row) < 2:
            continue
        label, value = row[0], row[1]
        if label is None or value is None:
            continue
        label_str = str(label).strip()
        if label_str in ("Total", ""):
            continue
        if isinstance(value, (int, float)):
            result[label_str] = round(float(value), 2)

    return result


_PERIOD_RE = re.compile(r"-\s*([A-Za-z]+ \d+,?\s*\d{4})\s*$")


def _find_report_end_date(rows: list) -> Optional[str]:
    """Pull end date from header (e.g. 'May 1, 2025-May 27, 2026' → '05/27/2026')."""
    for i in range(min(10, len(rows))):
        if not rows[i] or rows[i][0] is None:
            continue
        m = _PERIOD_RE.search(str(rows[i][0]))
        if not m:
            continue
        try:
            dt = datetime.strptime(m.group(1).replace(",", ""), "%B %d %Y")
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            continue
    return None


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
