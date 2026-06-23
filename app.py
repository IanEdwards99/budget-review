from __future__ import annotations

import json
import re
import sys
import urllib.request
from datetime import date
from pathlib import Path
from uuid import uuid4

APP_DIR = Path(__file__).resolve().parent
VENDOR = APP_DIR / "vendor"
if VENDOR.exists():
    sys.path.insert(0, str(VENDOR))

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from budget_engine import Period, build_workbook, read_abn, read_wise, workbook_payload_for_apps_script


app = Flask(__name__)
RUNS_DIR = APP_DIR / "runs"
UPLOADS_DIR = APP_DIR / "uploads"
RUNS_DIR.mkdir(exist_ok=True)
UPLOADS_DIR.mkdir(exist_ok=True)


def parse_sheet_id(text: str) -> str:
    text = (text or "").strip()
    match = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", text)
    return match.group(1) if match else text


def post_to_apps_script(script_url: str, spreadsheet_id: str, workbook_path: Path, sheet_names: list[str], title: str) -> dict:
    payload = {
        "spreadsheetId": parse_sheet_id(spreadsheet_id),
        "title": title,
        "sheets": workbook_payload_for_apps_script(workbook_path, sheet_names),
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(script_url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=60) as response:
        body = response.read().decode("utf-8")
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return {"ok": False, "message": body}


@app.get("/")
def index():
    return """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Budget Review Builder</title>
  <style>
    :root {
      --ink: #172033;
      --muted: #5f6b7a;
      --line: #d8e0ea;
      --blue: #1f5f9f;
      --blue-2: #2e75b6;
      --green: #16745c;
      --red: #a33a3a;
      --paper: #ffffff;
      --wash: #f5f8fb;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font: 14px/1.45 "Segoe UI", system-ui, sans-serif;
      color: var(--ink);
      background: var(--wash);
    }
    main {
      max-width: 1180px;
      margin: 0 auto;
      padding: 28px 20px 44px;
    }
    header {
      display: flex;
      align-items: end;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 18px;
    }
    h1 { margin: 0; font-size: 24px; letter-spacing: 0; }
    .sub { color: var(--muted); margin-top: 4px; }
    form {
      display: grid;
      grid-template-columns: minmax(0, 1.1fr) minmax(320px, .9fr);
      gap: 18px;
    }
    section, .result {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
    }
    h2 { margin: 0 0 12px; font-size: 15px; }
    .dropgrid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
    .drop {
      min-height: 118px;
      border: 1.5px dashed #9fb4ca;
      border-radius: 8px;
      padding: 14px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      gap: 8px;
      background: #fbfdff;
      cursor: pointer;
      transition: border-color .15s ease, background .15s ease, box-shadow .15s ease;
    }
    .drop.dragover {
      border-color: var(--blue-2);
      background: #eef6ff;
      box-shadow: 0 0 0 3px rgba(46, 117, 182, .12);
    }
    .drop strong { color: var(--blue); }
    .drop input { display: none; }
    .file-name { color: var(--muted); overflow-wrap: anywhere; min-height: 20px; }
    .fields { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; }
    label span { display: block; font-weight: 650; margin-bottom: 5px; }
    input[type="text"], input[type="date"], input[type="url"] {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 9px 10px;
      font: inherit;
      background: #fff;
    }
    .wide { grid-column: 1 / -1; }
    .actions {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-top: 14px;
    }
    button {
      border: 0;
      border-radius: 6px;
      background: var(--blue-2);
      color: white;
      padding: 10px 14px;
      font-weight: 700;
      cursor: pointer;
    }
    button:disabled { opacity: .55; cursor: wait; }
    .hint { color: var(--muted); font-size: 12px; }
    .help { display: block; color: var(--muted); font-size: 12px; margin-top: 5px; }
    .steps {
      margin: 8px 0 12px;
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fbfdff;
      color: var(--muted);
      font-size: 12px;
    }
    .steps ol { margin: 0; padding-left: 18px; }
    .steps li + li { margin-top: 4px; }
    .result { margin-top: 18px; display: none; }
    .result.show { display: block; }
    .ok { color: var(--green); font-weight: 700; }
    .bad { color: var(--red); font-weight: 700; }
    a.download {
      display: inline-block;
      margin-top: 10px;
      color: var(--blue);
      font-weight: 700;
    }
    @media (max-width: 880px) {
      form, .dropgrid, .fields { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div>
        <h1>Budget Review Builder</h1>
        <div class="sub">Drop ABN AMRO and Wise exports, pick the budget cycle, generate the formula-connected workbook.</div>
      </div>
    </header>

    <form id="budget-form">
      <section>
        <h2>Statements</h2>
        <div class="dropgrid">
          <label class="drop">
            <strong>ABN AMRO export</strong>
            <span class="hint">Drop the ABN file here: .xls or .xlsx with transactiondate, amount, description</span>
            <input required type="file" name="abn" accept=".xls,.xlsx">
            <span class="file-name" data-for="abn">No file selected</span>
          </label>
          <label class="drop">
            <strong>Wise transaction history</strong>
            <span class="hint">Drop the Wise file here: CSV or Excel export from Wise</span>
            <input required type="file" name="wise" accept=".csv,.xls,.xlsx">
            <span class="file-name" data-for="wise">No file selected</span>
          </label>
          <label class="drop wide">
            <strong>Existing workbook template</strong>
            <span class="hint">Optional. Drop an existing .xlsx workbook here, or leave empty.</span>
            <input type="file" name="template" accept=".xlsx">
            <span class="file-name" data-for="template">No file selected</span>
          </label>
        </div>
      </section>

      <section>
        <h2>Review period</h2>
        <div class="fields">
          <label class="wide"><span>Tab label</span><input name="label" type="text" value="June 2026" required></label>
          <label><span>Start</span><input name="start" type="date" value="2026-05-25" required></label>
          <label><span>End</span><input name="end" type="date" value="2026-06-24" required></label>
          <label><span>Data through</span><input name="through" type="date" value="2026-06-20" required><small class="help">The latest transaction date included in your files. If the month is unfinished, this helps the review show whether spending is on pace.</small></label>
        </div>

        <h2 style="margin-top:18px">Google Sheets sync</h2>
        <div class="steps">
          <ol>
            <li>Use a native Google Sheet URL that looks like <strong>docs.google.com/spreadsheets/d/...</strong>. If your file is an uploaded .xlsx, first open it in Sheets and choose <strong>File &gt; Save as Google Sheets</strong>.</li>
            <li>To update an existing Sheet, paste that native Sheet URL into <strong>Native Google Sheet URL or ID</strong>. To create a new Google Sheet instead, leave this field blank.</li>
            <li>Paste your Apps Script deployment URL, ending in <strong>/exec</strong>, into <strong>Apps Script Web App URL</strong>.</li>
            <li>Fill only the Apps Script URL to create a new Google Sheet, fill both URL fields to update an existing one, or leave both blank to only download the workbook.</li>
          </ol>
        </div>
        <div class="fields">
          <label class="wide"><span>Native Google Sheet URL or ID</span><input name="sheet_id" type="text" placeholder="Optional: https://docs.google.com/spreadsheets/d/..."><small class="help">Optional. Leave blank to create a new Google Sheet. To update an existing file, use a real Google Sheet, not an Excel .xlsx file opened in Google Drive.</small></label>
          <label class="wide"><span>Apps Script Web App URL</span><input name="script_url" type="url" placeholder="https://script.google.com/macros/s/.../exec"></label>
        </div>
        <div class="hint" style="margin-top:8px">Leave both sync fields blank to just download the workbook. Paste only the Apps Script URL to create a new Google Sheet. Paste both URLs to replace the named tabs in an existing native Google Sheet.</div>

        <div class="actions">
          <button id="submit" type="submit">Generate review</button>
          <span id="status" class="hint"></span>
        </div>
      </section>
    </form>

    <div id="result" class="result"></div>
  </main>

  <script>
    function updateFileName(input) {
      const name = input.files[0] ? input.files[0].name : 'No file selected';
      document.querySelector(`[data-for="${input.name}"]`).textContent = name;
    }

    document.querySelectorAll('input[type=file]').forEach(input => {
      input.addEventListener('change', () => updateFileName(input));
    });

    document.querySelectorAll('.drop').forEach(drop => {
      const input = drop.querySelector('input[type=file]');
      ['dragenter', 'dragover'].forEach(eventName => {
        drop.addEventListener(eventName, event => {
          event.preventDefault();
          event.stopPropagation();
          drop.classList.add('dragover');
        });
      });
      ['dragleave', 'dragend'].forEach(eventName => {
        drop.addEventListener(eventName, event => {
          event.preventDefault();
          event.stopPropagation();
          drop.classList.remove('dragover');
        });
      });
      drop.addEventListener('drop', event => {
        event.preventDefault();
        event.stopPropagation();
        drop.classList.remove('dragover');
        if (event.dataTransfer.files.length) {
          input.files = event.dataTransfer.files;
          updateFileName(input);
        }
      });
    });

    const form = document.getElementById('budget-form');
    const button = document.getElementById('submit');
    const status = document.getElementById('status');
    const result = document.getElementById('result');

    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      button.disabled = true;
      status.textContent = 'Working...';
      result.classList.remove('show');
      result.innerHTML = '';
      try {
        const response = await fetch('/generate', { method: 'POST', body: new FormData(form) });
        const payload = await response.json();
        if (!response.ok || !payload.ok) throw new Error(payload.error || 'Generation failed');
        const sheetLink = payload.google_sync && payload.google_sync.spreadsheet_url ? `<p><a class="download" href="${payload.google_sync.spreadsheet_url}" target="_blank" rel="noopener">Open Google Sheet</a></p>` : '';
        const sync = payload.google_sync ? `<p class="${payload.google_sync.ok ? 'ok' : 'bad'}">Google sync: ${payload.google_sync.message}</p>${sheetLink}` : '';
        result.innerHTML = `
          <h2>Review generated</h2>
          <p><strong>${payload.review_sheet}</strong> contains ${payload.rows} source rows. Expenses: EUR ${payload.expense_total.toFixed(2)}. Credits: EUR ${payload.credits_total.toFixed(2)}.</p>
          ${sync}
          <a class="download" href="${payload.download_url}">Download workbook</a>
        `;
        result.classList.add('show');
        status.textContent = 'Done';
      } catch (error) {
        result.innerHTML = `<p class="bad">${error.message}</p>`;
        result.classList.add('show');
        status.textContent = 'Needs attention';
      } finally {
        button.disabled = false;
      }
    });
  </script>
</body>
</html>
"""


@app.post("/generate")
def generate():
    run_id = uuid4().hex[:10]
    run_dir = RUNS_DIR / run_id
    upload_dir = UPLOADS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    upload_dir.mkdir(parents=True, exist_ok=True)

    try:
        abn_file = request.files["abn"]
        wise_file = request.files["wise"]
        template_file = request.files.get("template")

        abn_path = upload_dir / secure_filename(abn_file.filename)
        wise_path = upload_dir / secure_filename(wise_file.filename)
        abn_file.save(abn_path)
        wise_file.save(wise_path)

        template_path = None
        if template_file and template_file.filename:
            template_path = upload_dir / secure_filename(template_file.filename)
            template_file.save(template_path)

        period = Period(
            request.form["label"].strip(),
            date.fromisoformat(request.form["start"]),
            date.fromisoformat(request.form["end"]),
            date.fromisoformat(request.form["through"]),
        )

        rows = read_abn(abn_path, period) + read_wise(wise_path, period)
        rows.sort(key=lambda row: (row["date"], row["source"], row["source_id"]))

        output_path = run_dir / f"budget_review_{secure_filename(period.name).replace(' ', '_')}.xlsx"
        result = build_workbook(rows, period, output_path, template_path)
        result.update({"ok": True, "download_url": f"/download/{run_id}/{output_path.name}"})

        script_url = request.form.get("script_url", "").strip()
        sheet_id = request.form.get("sheet_id", "").strip()
        if sheet_id and not script_url:
            result["google_sync"] = {
                "ok": False,
                "message": "Google sync skipped: paste the Apps Script Web App URL too, or leave both sync fields blank.",
            }
        if script_url:
            try:
                sheet_names = ["Settings 2026", "Raw Txns 2026", result["review_sheet"]]
                sync_result = post_to_apps_script(script_url, sheet_id, output_path, sheet_names, f"Budget Review {period.name}")
                result["google_sync"] = {
                    "ok": bool(sync_result.get("ok")),
                    "message": sync_result.get("message", "Sheets updated" if sync_result.get("ok") else "Apps Script returned no message"),
                    "spreadsheet_id": sync_result.get("spreadsheetId") or parse_sheet_id(sheet_id),
                    "spreadsheet_url": sync_result.get("spreadsheetUrl"),
                    "tabs": sheet_names,
                }
            except Exception as exc:
                result["google_sync"] = {"ok": False, "message": str(exc), "spreadsheet_id": parse_sheet_id(sheet_id)}

        return jsonify(result)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/download/<run_id>/<filename>")
def download(run_id: str, filename: str):
    path = RUNS_DIR / secure_filename(run_id) / secure_filename(filename)
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5057, debug=False)
