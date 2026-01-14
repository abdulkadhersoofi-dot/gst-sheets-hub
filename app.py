from flask import Flask, jsonify, request, render_template
import gspread
from gspread.utils import ValueRenderOption  # for formula-aware reads
from gspread.exceptions import APIError

# ---------- CONFIG ----------
MASTER_CONFIG_ID = "1ZAU_kvQEc6_B6-dwL6QdvbUpWkN52kE1zVQHcxBG7Lk"  # GST – Master Config sheet ID
SERVICE_ACCOUNT_FILE = "service_account.json"
# Master Config columns expected:
#   CompanyId | CompanyName | SpreadsheetId | (optional extra columns)
# ----------------------------------------

app = Flask(__name__)
gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)


def load_companies():
    """
    Read all company rows from Master Config sheet.

    Each row must contain at least:
      - CompanyId
      - CompanyName
      - SpreadsheetId
    """
    sh = gc.open_by_key(MASTER_CONFIG_ID)
    ws = sh.get_worksheet(0)  # first tab
    records = ws.get_all_records()
    return records


@app.route("/")
def index():
    """Serve main HTML page (company list first, then company detail)."""
    return render_template("index.html")


# ---------------- COMPANY + SHEET LISTING ---------------- #

@app.route("/companies", methods=["GET"])
def get_companies():
    """
    Return list of companies for the selection screen.
    """
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
    """
    List all worksheet names inside a company's Google Spreadsheet.
    """
    records = load_companies()
    record = next((r for r in records if r.get("CompanyId") == company_id), None)
    if not record:
        return jsonify({"error": "Company not found"}), 404

    spreadsheet_id = record.get("SpreadsheetId")
    if not spreadsheet_id:
        return jsonify({"error": "SpreadsheetId missing in Master Config"}), 400

    sh = gc.open_by_key(spreadsheet_id)
    sheet_list = [
        {"sheetName": ws.title, "index": ws.index}
        for ws in sh.worksheets()
    ]  # worksheets() is the standard way to list tabs.[web:347]
    return jsonify(sheet_list)


# ---------------- SHEET READ ---------------- #

@app.route("/sheet/<company_id>", methods=["GET"])
def get_company_sheet(company_id):
    """
    Return full sheet values + editable mask for a given company + sheet.

    Query parameter:
      ?sheet=<sheet_name>
    """
    sheet_name = request.args.get("sheet")
    if not sheet_name:
        return jsonify({"error": "sheet parameter is required"}), 400

    records = load_companies()
    record = next((r for r in records if r.get("CompanyId") == company_id), None)
    if not record:
        return jsonify({"error": "Company not found"}), 404

    spreadsheet_id = record.get("SpreadsheetId")
    if not spreadsheet_id:
        return jsonify({"error": "SpreadsheetId missing in Master Config"}), 400

    sh = gc.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(sheet_name)  # select tab by title.[web:347]
    except Exception as e:
        return jsonify({"error": f"Sheet '{sheet_name}' not found: {e}"}), 404

    # 1) Display values (what user sees)
    values_display = ws.get_all_values()  # 2D list

    # 2) Same range, but formulas preserved
    values_formula = ws.get_values(
        value_render_option=ValueRenderOption.formula
    )  # formulas as text, e.g. "=SUM(A1:A5)".[web:342][web:345]

    # Build editable mask: False if formula, True otherwise
    editable_mask = []
    for row_disp, row_form in zip(values_display, values_formula):
        row_mask = []
        for disp, form in zip(row_disp, row_form):
            if isinstance(form, str) and form.startswith("="):
                row_mask.append(False)   # formula cell -> lock
            else:
                row_mask.append(True)    # normal value -> editable
        editable_mask.append(row_mask)

    return jsonify(
        {
            "company": {
                "CompanyId": record.get("CompanyId"),
                "CompanyName": record.get("CompanyName"),
                "SpreadsheetId": spreadsheet_id,
            },
            "sheet": sheet_name,
            "values": values_display,
            "editable": editable_mask,
        }
    )


# ---------------- SHEET UPDATE ---------------- #

@app.route("/sheet/<company_id>/update", methods=["POST"])
def update_company_sheet(company_id):
    """
    Overwrite only editable cells in a given sheet, keeping formulas untouched.

    URL:
      POST /sheet/<company_id>/update?sheet=<sheet_name>
    """
    sheet_name = request.args.get("sheet")
    if not sheet_name:
        return jsonify({"error": "sheet parameter is required"}), 400

    payload = request.get_json(force=True) or {}
    new_values = payload.get("values")
    editable = payload.get("editable")

    if not isinstance(new_values, list) or not isinstance(editable, list):
        return jsonify({"error": "values and editable must be 2D lists"}), 400

    records = load_companies()
    record = next((r for r in records if r.get("CompanyId") == company_id), None)
    if not record:
        return jsonify({"error": "Company not found"}), 404

    spreadsheet_id = record.get("SpreadsheetId")
    if not spreadsheet_id:
        return jsonify({"error": "SpreadsheetId missing in Master Config"}), 400

    sh = gc.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(sheet_name)
    except Exception as e:
        return jsonify({"error": f"Sheet '{sheet_name}' not found: {e}"}), 404

    # Get current sheet with formulas preserved
    current = ws.get_values(
        value_render_option=ValueRenderOption.formula
    )  # formulas remain as text so we don't lose them.[web:342]

    # Merge new values into current only where editable == True
    rows = min(len(current), len(new_values))
    for r in range(rows):
        cols = min(len(current[r]), len(new_values[r]))
        for c in range(cols):
            can_edit = (
                r < len(editable)
                and c < len(editable[r])
                and editable[r][c] is True
            )
            if can_edit:
                current[r][c] = new_values[r][c]

    # Write merged data back starting at A1, letting Sheets interpret input
    ws.update("A1", current, value_input_option="USER_ENTERED")
    return jsonify({"status": "ok", "rows": rows, "sheet": sheet_name})


# ---------------- SHEET CLONE (APR -> NEW MONTH) ---------------- #

@app.route("/sheet/<company_id>/clone", methods=["POST"])
def clone_company_sheet(company_id):
    """
    Create a new worksheet in the company's spreadsheet, based on a template
    sheet (e.g. 'APR 25'):

    - Copies all headings, formats, and formulas.
    - Clears only numeric values in the new sheet.
      * Formulas are kept.
      * Any text (headings, labels) is kept.
    """
    payload = request.get_json(force=True) or {}
    source_name = payload.get("source_sheet")
    new_name    = payload.get("new_sheet")

    if not source_name or not new_name:
        return jsonify({"error": "source_sheet and new_sheet are required"}), 400

    records = load_companies()
    record = next((r for r in records if r.get("CompanyId") == company_id), None)
    if not record:
        return jsonify({"error": "Company not found"}), 404

    spreadsheet_id = record.get("SpreadsheetId")
    if not spreadsheet_id:
        return jsonify({"error": "SpreadsheetId missing in Master Config"}), 400

    sh = gc.open_by_key(spreadsheet_id)

    # 1) Get source worksheet
    try:
        src_ws = sh.worksheet(source_name)
    except Exception as e:
        return jsonify({"error": f"Source sheet '{source_name}' not found: {e}"}), 404

    # 2) Duplicate entire sheet structure (formats + formulas + values)
    try:
        duplicated_ws = sh.duplicate_sheet(src_ws.id, new_sheet_name=new_name)
    except APIError as e:
        return jsonify({"error": f"Cannot create sheet '{new_name}': {e}"}), 400

    new_ws_obj = duplicated_ws  # Worksheet instance.[web:347]

    # 3) Get formulas/values from new sheet
    formulas = new_ws_obj.get_values(
        value_render_option=ValueRenderOption.formula
    )  # formulas as text, plain values otherwise.[web:342][web:345]

    # 4) Build cleaned matrix:
    #    - keep all formulas
    #    - keep all text (headings)
    #    - clear only numeric values
    cleaned = []
    for row in formulas:
        cleaned_row = []
        for val in row:
            # keep formulas
            if isinstance(val, str) and val.startswith("="):
                cleaned_row.append(val)
            else:
                # attempt to parse as number; if numeric, clear it
                s = str(val).strip()
                if s == "":
                    # empty already
                    cleaned_row.append("")
                else:
                    try:
                        float(s.replace(",", ""))
                        # numeric -> clear
                        cleaned_row.append("")
                    except (ValueError, TypeError):
                        # not numeric -> heading/label/text -> keep
                        cleaned_row.append(val)
        cleaned.append(cleaned_row)

    # 5) Write cleaned data back to new sheet
    new_ws_obj.update("A1", cleaned, value_input_option="USER_ENTERED")

    return jsonify({
        "status": "ok",
        "company": company_id,
        "source_sheet": source_name,
        "new_sheet": new_name
    })


if __name__ == "__main__":
    app.run(debug=True)
