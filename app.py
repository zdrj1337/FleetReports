"""
Web interface for the fleet payment report generator.

A thin Flask layer on top of fleet_logic.py, generate_reports.py and pdf_report.py
— it does not re-implement any business logic, it just calls them. Pick one of the
bundled sample weeks, generate the reports, view each driver's PDF in the browser,
and download the PDF or Excel files.

Run locally:
    pip install -r requirements.txt
    python app.py
    # open http://127.0.0.1:5000
"""

import io
import os
import re
import zipfile

from flask import Flask, abort, render_template, request, send_file

import fleet_logic
import generate_reports as gen
import pdf_report

app = Flask(__name__)
BASE = os.path.dirname(os.path.abspath(__file__))

PERIOD_RE = re.compile(r"^\d{4}-W\d{2}$")
FILE_RE = re.compile(r"^[A-Za-z0-9_]+\.(xlsx|pdf)$")


@app.template_filter("money")
def money(value):
    """Format a number as 1,234.56 (RON shown in the template)."""
    return "{:,.2f}".format(float(value))


def _periods():
    """All sample periods available under data/, newest first."""
    return list(reversed(gen._find_all_periods(BASE)))


def _dir_for_key(key):
    """Resolve a period key (e.g. 2026-W10) to its reports directory."""
    if not PERIOD_RE.match(key):
        abort(404)
    year, week = key.split("-")
    return os.path.join(BASE, "reports", year, week)


def _write_reports(df, label, out_dir):
    """Generate an Excel file and a PDF per driver, plus the fleet Excel summary."""
    os.makedirs(out_dir, exist_ok=True)
    for f in os.listdir(out_dir):
        if f.endswith((".xlsx", ".pdf")):
            os.remove(os.path.join(out_dir, f))
    for d in fleet_logic.drivers_in(df):
        stem = d.replace(" ", "_")
        gen.write_driver_report(d, df, label, os.path.join(out_dir, f"{stem}.xlsx"))
        pdf_report.write_driver_pdf(d, df, label, os.path.join(out_dir, f"{stem}.pdf"))
    gen.write_fleet_summary(df, label, os.path.join(out_dir, "_fleet_summary.xlsx"))
    pdf_report.write_fleet_summary_pdf(df, label, os.path.join(out_dir, "_fleet_summary.pdf"))


def _build_result(df, label, key):
    """Assemble the on-page summary (KPI totals + per-driver rows)."""
    s = fleet_logic.compute_summary(df)
    rows = [{
        "driver": t.driver,
        "gross": t.gross,
        "net": t.net,
        "bank_transfer": t.bank_transfer,
        "stem": t.driver.replace(" ", "_"),
    } for t in s.itertuples()]
    totals = {
        "gross": round(float(s["gross"].sum()), 2),
        "firm": round(float(s["firm_commission"].sum()), 2),
        "net": round(float(s["net"].sum()), 2),
        "transfer": round(float(s["bank_transfer"].sum()), 2),
    }
    return {"key": key, "label": label, "rows": rows, "totals": totals, "n": len(rows)}


@app.route("/")
def index():
    return render_template("index.html", periods=_periods(), result=None, error=None)


@app.route("/generate", methods=["POST"])
def generate():
    period = request.form.get("period", "")
    if not PERIOD_RE.match(period):
        return render_template("index.html", periods=_periods(), result=None,
                               error="Please choose a valid week.")
    data_dir = fleet_logic.resolve_data_dir(period, BASE)
    if not os.path.isdir(data_dir):
        return render_template("index.html", periods=_periods(), result=None,
                               error=f"No data found for {period}.")
    df = fleet_logic.load_period(data_dir)
    year, week = period.split("-")
    out_dir = os.path.join(BASE, "reports", year, week)
    _write_reports(df, period, out_dir)
    return render_template("index.html", periods=_periods(),
                           result=_build_result(df, period, key=period), error=None)


@app.route("/view/<key>/<path:filename>")
def view(key, filename):
    """Serve a PDF inline so it opens in the browser tab."""
    if not FILE_RE.match(filename) or not filename.endswith(".pdf"):
        abort(404)
    path = os.path.join(_dir_for_key(key), filename)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, mimetype="application/pdf", as_attachment=False, download_name=filename)


@app.route("/download/<key>/<path:filename>")
def download(key, filename):
    if not FILE_RE.match(filename):
        abort(404)
    path = os.path.join(_dir_for_key(key), filename)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=filename)


@app.route("/download_all/<key>")
def download_all(key):
    directory = _dir_for_key(key)
    if not os.path.isdir(directory):
        abort(404)
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(os.listdir(directory)):
            if f.endswith((".pdf", ".xlsx")):
                z.write(os.path.join(directory, f), f)
    mem.seek(0)
    return send_file(mem, as_attachment=True, mimetype="application/zip",
                     download_name=f"fleet-reports-{key}.zip")


if __name__ == "__main__":
    app.run(debug=True)
