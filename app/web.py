"""The local website: choose files -> check entries -> download certificates.

    python -m app.web          (or double-click Start.bat)

Runs only on this computer (127.0.0.1), no internet needed. Uploaded registers are kept in
uploads/<id>/ for a day so the review screen can be generated from; every download is also
saved in output/<date time>/ as a history.

The ZIP is sorted by grade: a file per certificate goes in a "PE 63" / "PE 80" / "PE 100"
folder; "all in one file" makes one file per grade and format (PE 63.pdf, PE 80.pdf...).
Only grades with ticked entries appear.
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
from app.render_pdf import write_pdf, write_pdf_combined
from app.render_word import write_word, write_word_combined
from app.report import build_report
from app.report2 import Report2Error, write_report2
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
    for grade in GRADE_ORDER:
        rows = [r for r in diff.new if r.grade == grade]
        # Report No. starts as the register's column D (MP63/171); editable on the screen
        good = [(r, r.report_no) for r in rows if r.ok]
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
    # Report 1 = a certificate per entry, Report 2 = the testing record (entries side by side, one Excel
    # file). The check screen always sends report_choice; without it (older forms) Report 1 is made.
    which = set(request.form.getlist("report")) & {"1", "2"}
    if not which and "report_choice" not in request.form:
        which = {"1"}
    if not which:
        abort(400, "Tick Report 1, Report 2 or both.")
    # the check screen always sends format_choice; without it (older forms) Excel is made
    formats = [f for f in ("pdf", "excel", "word") if f in request.form.getlist("format")] if "1" in which else []
    if "1" in which and not formats and "format_choice" not in request.form:
        formats = ["excel"]
    if "1" in which and not formats:
        abort(400, "Tick at least one format: PDF, Excel or Word.")
    incomplete = None
    upload = request.files.get("incomplete")
    if "2" in which and upload and upload.filename:
        suffix = Path(upload.filename).suffix.lower()
        if suffix not in (".xls", ".xlsx"):
            abort(400, "The incomplete Report 2 must be an Excel file (.xls or .xlsx).")
        incomplete = d / f"incomplete report 2{suffix}"
        upload.save(incomplete)

    settings = with_company(load_settings(), (d / "company.txt").read_text(encoding="utf-8"))
    diff = find_new_rows(*_files(d))
    by_id = {row_id(r): r for r in diff.new if r.ok}
    reports = []
    for rid in chosen:
        row = by_id.get(rid)
        if row is None:
            abort(400, "An entry is no longer in the register. Choose the files again.")
        number = request.form.get(f"no-{rid}", row.report_no).strip()
        reports.append(build_report(row, number, settings))
    reports.sort(key=lambda rep: GRADE_ORDER.index(rep.row.grade))   # stable: register order within a grade
    _unique_names(reports)

    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    out = _new_folder(OUTPUT / stamp)      # two downloads in the same second get separate folders
    writers = {"excel": (".xlsx", write_excel, write_excel_combined),
               "pdf": (".pdf", write_pdf, write_pdf_combined),
               "word": (".docx", write_word, write_word_combined)}
    by_grade = {g: [rep for rep in reports if rep.row.grade == g] for g in GRADE_ORDER}
    made, report2 = [], None
    if "2" in which:
        try:
            report2, _ = write_report2(reports, out / "Report 2.xlsx", incomplete, settings)
        except Report2Error as e:
            shutil.rmtree(out, ignore_errors=True)
            abort(400, f"Report 2 not made: {e}.")
    for fmt in formats:
        ext, one, many = writers[fmt]
        for grade, reps in by_grade.items():
            if not reps:
                continue
            if combined:
                made.append(many(reps, out / f"{grade}{ext}"))
            else:
                for rep in reps:
                    made.append(one(rep, out / grade / f"{rep.file_name}{ext}"))

    if which == {"2"}:                     # Report 2 alone: just its Excel file, no ZIP
        resp = send_file(report2, as_attachment=True, download_name=f"Report 2 {stamp[:10]}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        resp.headers["X-Certificates"] = str(len(reports))
        resp.headers["X-Saved-To"] = str(out)
        return resp
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        order = {g: i for i, g in enumerate(GRADE_ORDER)}
        for f in sorted(made, key=lambda p: (order[p.stem if combined else p.parent.name], p.name)):
            z.write(f, f.relative_to(out).as_posix())     # only this download's files
        if report2:
            z.write(report2, report2.name)
    buf.seek(0)
    resp = send_file(buf, mimetype="application/zip", as_attachment=True,
                     download_name=f"Test certificates {stamp[:10]}.zip")
    resp.headers["X-Certificates"] = str(len(reports))
    resp.headers["X-Saved-To"] = str(out)
    return resp


def _new_folder(path: Path) -> Path:
    """path, or path (2), path (3)... if it already exists; created empty."""
    candidate, n = path, 2
    while True:
        try:
            candidate.mkdir(parents=True)
            return candidate
        except FileExistsError:
            candidate, n = path.with_name(f"{path.name} ({n})"), n + 1


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
