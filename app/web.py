"""The local website: choose files -> check entries -> download certificates.

    python -m app.web          (or double-click Start.bat)

Runs only on this computer (127.0.0.1), no internet needed. Uploaded registers are kept in
uploads/<id>/ for a day so the review screen can be generated from; every download is also
saved in output/<date time>/ as a history.
"""
from __future__ import annotations

import io
import secrets
import shutil
import sys
import threading
import time
import webbrowser
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, render_template, request, send_file

from app.production import GRADE_ORDER, find_new_rows
from app.render_excel import write_excel, write_excel_combined
from app.report import build_report
from app.settings import SettingsError, companies, load_settings, with_company

ROOT = Path(__file__).resolve().parent.parent
UPLOADS = ROOT / "uploads"
OUTPUT = ROOT / "output"
ALLOWED = {".xls", ".xlsx"}
KEEP_UPLOADS_HOURS = 24
HOST, PORT = "127.0.0.1", 5050

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024   # registers are a few hundred KB


def row_id(row) -> str:
    """Stable id of a production row: grade + financial year + Sr. No. (same key as detection)."""
    grade, fy, sr = row.key
    return f"{grade.replace(' ', '')}-{fy}-{sr}"


def _clean_old_uploads():
    if not UPLOADS.exists():
        return
    cutoff = time.time() - KEEP_UPLOADS_HOURS * 3600
    for d in UPLOADS.iterdir():
        if d.is_dir() and d.stat().st_mtime < cutoff:
            shutil.rmtree(d, ignore_errors=True)


def _upload_dir(token: str) -> Path:
    if not token or not token.isalnum():
        abort(400)
    d = UPLOADS / token
    if not d.is_dir():
        abort(410, "These files are no longer here. Choose the registers again.")
    return d


def _files(d: Path) -> tuple[Path, Path]:
    old = next(d.glob("old.*"))
    new = next(d.glob("new.*"))
    return old, new


def _companies():
    try:
        return companies(load_settings())
    except SettingsError:
        return []


@app.get("/")
def choose_files():
    return render_template("upload.html", step=1, companies=_companies(), chosen=None, error=None)


@app.post("/check")
def check_entries():
    _clean_old_uploads()
    files = {k: request.files.get(k) for k in ("old", "new")}
    missing = [label for k, label in (("old", "old register"), ("new", "new register"))
               if not files[k] or not files[k].filename]
    company = request.form.get("company", "")
    names_by_id = dict(_companies())
    if company not in names_by_id:
        return _upload_error("Choose the company the certificates are for.")
    if missing:
        return _upload_error(f"Choose the {' and the '.join(missing)}.", company)
    for k, f in files.items():
        if Path(f.filename).suffix.lower() not in ALLOWED:
            return _upload_error(f"'{f.filename}' isn't an Excel file. Choose the .xls or .xlsx production register.",
                                 company)

    token = secrets.token_hex(8)
    d = UPLOADS / token
    d.mkdir(parents=True)
    names = {}
    for k, f in files.items():
        f.save(d / f"{k}{Path(f.filename).suffix.lower()}")
        names[k] = f.filename
    (d / "company.txt").write_text(company, encoding="utf-8")

    try:
        diff = find_new_rows(*_files(d))
    except Exception as e:   # unreadable / corrupt workbook
        shutil.rmtree(d, ignore_errors=True)
        return _upload_error(f"One of the files couldn't be opened as an Excel register ({e}). "
                             "Open it in Excel, save it again and choose it here.", company)
    return _review(token, names, names_by_id[company], diff)


def _upload_error(message, company=None):
    return render_template("upload.html", step=1, companies=_companies(), chosen=company, error=message), 400


def _review(token, names, company_name, diff):
    groups = []
    number = load_settings()["report"]["number_starts_at"]
    for grade in GRADE_ORDER:
        rows = [r for r in diff.new if r.grade == grade]
        good = []
        for r in rows:
            if r.ok:
                good.append((r, number))
                number += 1
        groups.append({"grade": grade, "rows": good,
                       "skipped": [r for r in rows if not r.ok]})
    return render_template("review.html", step=2, company_name=company_name, token=token, names=names,
                           groups=groups, errors=diff.errors, notes=diff.notes,
                           total=sum(len(g["rows"]) for g in groups), row_id=row_id)


@app.post("/download")
def download():
    token = request.form.get("token", "")
    d = _upload_dir(token)
    chosen = request.form.getlist("row")
    if not chosen:
        abort(400, "Tick at least one entry.")
    combined = request.form.get("files") == "one"

    settings = with_company(load_settings(), (d / "company.txt").read_text(encoding="utf-8"))
    diff = find_new_rows(*_files(d))
    by_id = {row_id(r): r for r in diff.new if r.ok}
    reports = []
    for rid in chosen:
        row = by_id.get(rid)
        if row is None:
            abort(400, "An entry is no longer in the register. Choose the files again.")
        number = request.form.get(f"no-{rid}", "").strip() or str(len(reports) + 1)
        reports.append(build_report(row, number, settings))
    _unique_names(reports)

    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    out = OUTPUT / stamp
    if combined:
        write_excel_combined(reports, out / f"Test certificates {stamp[:10]}.xlsx")
    else:
        for rep in reports:
            write_excel(rep, out / f"{rep.file_name}.xlsx")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(out.iterdir()):
            z.write(f, f.name)
    buf.seek(0)
    resp = send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"Test certificates {stamp[:10]}.zip")
    resp.headers["X-Certificates"] = str(len(reports))
    resp.headers["X-Saved-To"] = str(out)
    return resp


def _unique_names(reports):
    used = {}
    for rep in reports:
        k = rep.file_name.lower()
        used[k] = used.get(k, 0) + 1
        if used[k] > 1:
            rep.file_name = f"{rep.file_name}-{used[k]}"


@app.errorhandler(400)
@app.errorhandler(410)
def bad_request(e):
    return render_template("upload.html", step=1, companies=_companies(), chosen=None,
                           error=getattr(e, "description", None) or "Something was missing. Choose the files again."), e.code


def main():
    url = f"http://{HOST}:{PORT}/"
    if "--no-browser" not in sys.argv:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"\n  Test certificates is running at {url}\n  Keep this window open while you use it. Close it to stop.\n")
    app.run(host=HOST, port=PORT, debug=False)


if __name__ == "__main__":
    main()
