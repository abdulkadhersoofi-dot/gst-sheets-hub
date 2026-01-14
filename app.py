from flask import Flask, jsonify, request, render_template
import gspread
from gspread.utils import ValueRenderOption
from gspread.exceptions import APIError
import os
import json
import time

# ---------- CONFIG ----------
MASTER_CONFIG_ID = "1ZAU_kvQEc6_B6-dwL6QdvbUpWkN52kE1zVQHcxBG7Lk"  # GST – Master Config

# ---------- GOOGLE SHEETS AUTH VIA ENV ----------
# On Render: set env var SERVICE_ACCOUNT_JSON = full JSON of service account
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON")
if not SERVICE_ACCOUNT_JSON:
    raise RuntimeError("SERVICE_ACCOUNT_JSON env var is not set")

creds_info = json.loads(SERVICE_ACCOUNT_JSON)
gc = gspread.service_account_from_dict(creds_info)  # gspread supports dict auth [web:46]

# ---------- FLASK APP ----------
app = Flask(__name__)

# --------- Simple caching for Master Config ---------
companies_cache = None
companies_cache_ts = 0
CACHE_TTL_SECONDS = 60


def load_companies():
    """
    Read all company rows from Master Config sheet.

    Each row must contain at least:
      - CompanyId
      - CompanyName
      - SpreadsheetId
    """
    global companies_cache, companies_cache_ts
    now = time.time()
    if companies_cache is not None and (now - companies_cache_ts) < CACHE_TTL_SECONDS:
        return companies_cache

    sh = gc.open_by_key(MASTER_CONFIG_ID)
    ws = sh.get_worksheet(0)
    records = ws.get_all_records()
    companies_cache = records
    companies_cache_ts = now
    return records


@app.route("/")
def index():
    """Serve main HTML page (company list first, then company detail)."""
    return render_template("index.html")


# ---------------- COMPANY + SHEET LISTING ---------------- #

@app.route("/companies", methods=["GET"])
def get_companies():
    """Return list of companies for the selection screen."""
    records = load_companies()
    data = [
        {
            "CompanyId": r.get("CompanyId"),
            "CompanyName": r.get("CompanyName"),
            "SpreadsheetId": r.get("SpreadsheetId"),
        }
        for r in records
        if r.get("CompanyId") and r.get("SpreadsheetId")
    ]
    return jsonify(data)


@app.route("/company/<company_id>/sheets", methods=["GET"])
def list_company_sheets(company_id):
    """List all worksheet names inside a company's Google Spreadsheet."""
    records = load_companies()
    record = next((r for r in records if r.get("CompanyId") == company_id), None)
    if not record:
        return jsonify({"error": "Company not found"}), 404

    spreadsheet_id = record.get("SpreadsheetId")
    if not spreadsheet_id:
        return jsonify({"error": "SpreadsheetId missing in Master Config"}), 400

    sh = gc.open_by_key
