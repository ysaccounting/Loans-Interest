"""
Flask web app for the daily-balance interest calculator.

Upload one or more QuickBooks transaction reports. The loan each report
belongs to (Affiliates 8.5%, YS Tix -> YS Affiliates 12%, or Mazel /
Damona & crew 13%) is auto-detected from the report's content. The app
returns a single workbook with a Loan Terms tab plus one calculation tab
per report.
"""

import os
import tempfile
import uuid

from flask import (Flask, jsonify, render_template, request,
                   send_file, abort)
from werkzeug.utils import secure_filename

import interest_calc

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB cap

WORK_DIR = os.path.join(tempfile.gettempdir(), "interest_calc_jobs")
os.makedirs(WORK_DIR, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    return jsonify(status="ok")


@app.route("/process", methods=["POST"])
def process():
    files = request.files.getlist("files")
    files = [f for f in files if f and f.filename]
    if not files:
        return jsonify(error="Please choose at least one .xlsx file."), 400

    loan_types = request.form.getlist("loan_types")  # "auto" or a specific key
    try:
        rates = {
            "affiliates": float(request.form.get("affiliates_rate", "8.5")) / 100.0,
            "ystix": float(request.form.get("ystix_rate", "12")) / 100.0,
            "mazel": float(request.form.get("mazel_rate", "13")) / 100.0,
        }
        days_in_year = int(float(request.form.get("days_in_year", "365")))
    except ValueError:
        return jsonify(error="Rates and days-in-year must be numbers."), 400

    job = uuid.uuid4().hex
    job_dir = os.path.join(WORK_DIR, job)
    os.makedirs(job_dir, exist_ok=True)

    jobs = []
    for i, f in enumerate(files):
        if not f.filename.lower().endswith((".xlsx", ".xlsm")):
            return jsonify(error=f"{f.filename}: must be .xlsx or .xlsm."), 400
        path = os.path.join(job_dir, secure_filename(f.filename))
        f.save(path)
        lt = loan_types[i] if i < len(loan_types) else "auto"
        jobs.append({"path": path, "loan_type": lt})

    out_name = "Loans_and_Interest.xlsx"
    out_path = os.path.join(job_dir, out_name)
    try:
        info = interest_calc.generate(jobs, out_path, rates, days_in_year)
    except Exception as exc:  # noqa: BLE001
        return jsonify(error=f"Could not process the files: {exc}"), 422

    info["download_url"] = f"/download/{job}/{out_name}"
    return jsonify(info)


@app.route("/download/<job>/<name>")
def download(job, name):
    safe = secure_filename(name)
    path = os.path.join(WORK_DIR, secure_filename(job), safe)
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=safe)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)
