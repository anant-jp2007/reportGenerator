"""The local website: choose files -> check entries -> download certificates.

    python -m app.web          (or double-click Start.bat)

Runs only on this computer (127.0.0.1), no internet needed. Uploaded registers are kept in
uploads/<id>/ for a day so the review screen can be generated from; every download is also
saved in output/<date time>/ as a history.

New entries are found against the company's record of certified entries (ledger/<id>.json, see
app/ledger.py), so the register as it is now is enough. An old register can still be chosen: then
the two registers are compared as before, and the old one's entries are added to the record. A
company without a record needs the old register once. Every download adds its entries to the record.

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

from flask import Flask, abort, redirect, render_template, request, send_file, url_for

from app.ledger import LedgerError, find_new_in_ledger, ledger_path, load, readable_grades, set_aside
from app.production import GRADE_ORDER, find_new_rows, read_production
from app.render_excel import write_excel, write_excel_combined
from app.render_pdf import write_pdf, write_pdf_combined
from app.render_word import write_word, write_word_combined
from app.new_company import CompanyError, Draft, make_draft, preview_settings, read_sample, sample_row, save_company
from app.new_company import _png as to_png
from app.preview import certificate_html
from app.report import apply_values, build_report, editable_cells
from app.report2 import Report2Error, write_report2
from app.settings import SettingsError, companies, load_settings, with_company

ROOT = Path(__file__).resolve().parent.parent
UPLOADS = ROOT / "uploads"
OUTPUT = ROOT / "output"
LEDGER = ROOT / "ledger"
ALLOWED = {".xls", ".xlsx"}
KEEP_UPLOADS_HOURS = 24
HOST, PORT = "127.0.0.1", 5050

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024   # registers are a few hundred KB


def row_id(row) -> str:
    """Id of a production row: grade + financial year + Sr. No. (the detection key) + its row in the
    sheet, so two new rows with the same Sr. No. stay apart. The same file is read again on download."""
    grade, fy, sr = row.key
    return f"{grade.replace(' ', '')}-{fy}-{sr}-{row.excel_row}"


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


def _old_file(d: Path) -> Path | None:
    return next(d.glob("old.*"), None)


def _diff(d: Path):
    """The new entries of this upload: against the old register if one was chosen, else against the
    copy of the record taken when the files were checked (so a second download sees the same list)."""
    new, old = next(d.glob("new.*")), _old_file(d)
    if old:
        return find_new_rows(old, new)
    return find_new_in_ledger(load(d / "ledger.json"), new)


def _companies():
    try:
        return companies(load_settings())
    except SettingsError:
        return []


def _records(company_ids) -> dict[str, dict]:
    """Per company: is there a record to check against, and what it holds (for the first screen)."""
    out = {}
    for cid in company_ids:
        try:
            led = load(ledger_path(LEDGER, cid))
            out[cid] = {"started": led.started, "certified": led.certified, "last": led.last, "broken": False}
        except LedgerError:
            out[cid] = {"started": False, "certified": 0, "last": None, "broken": True}
    return out


def _start_page(chosen=None, error=None, code=200, added=None):
    listed = _companies()
    return render_template("upload.html", step=1, companies=listed, chosen=chosen, error=error, added=added,
                           records=_records([cid for cid, _ in listed])), code


@app.get("/")
def choose_files():
    added = dict(_companies()).get(request.args.get("added", ""))
    return _start_page(request.args.get("added") if added else None, added=added)


@app.post("/check")
def check_entries():
    _clean_old_uploads()
    files = {k: request.files.get(k) for k in ("old", "new")}
    files = {k: f for k, f in files.items() if f and f.filename}
    company = request.form.get("company", "")
    names_by_id = dict(_companies())
    if company not in names_by_id:
        return _upload_error("Choose the company the certificates are for.")
    record = None
    if "old" not in files:
        # one register: checked against the company's record, which must exist
        try:
            record = load(ledger_path(LEDGER, company))
        except LedgerError as e:
            why = str(e)
            return _upload_error(f"{why[0].upper()}{why[1:]}. Choose the old register too: the two registers are "
                                 "compared and a new record is started from them.", company)
        if not record.started:
            what = "the old register and the new register" if "new" not in files else "the old register too"
            return _upload_error(f"Choose {what}: there's no record yet of what was certified for "
                                 f"{names_by_id[company]}. After this download, the new register alone is enough.",
                                 company)
    if "new" not in files:
        return _upload_error("Choose the new register.", company)
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
    if record:
        shutil.copyfile(record.path, d / "ledger.json")

    try:
        diff = _diff(d)
    except Exception as e:   # unreadable / corrupt workbook
        shutil.rmtree(d, ignore_errors=True)
        return _upload_error(f"One of the files couldn't be opened as an Excel register ({e}). "
                             "Open it in Excel, save it again and choose it here.", company)
    return _review(token, names, names_by_id[company], diff, record)


def _upload_error(message, company=None):
    return _start_page(company, message, 400)


def _review(token, names, company_name, diff, record=None):
    groups = []
    for grade in GRADE_ORDER:
        rows = [r for r in diff.new if r.grade == grade]
        # Report No. starts as the register's column D (MP63/171); editable on the screen
        good = [(r, r.report_no) for r in rows if r.ok]
        groups.append({"grade": grade, "rows": good,
                       "skipped": [r for r in rows if not r.ok]})
    return render_template("review.html", step=2, company_name=company_name, token=token, names=names, record=record,
                           groups=groups, errors=diff.errors, notes=diff.notes,
                           total=sum(len(g["rows"]) for g in groups), row_id=row_id)


# ---------- add a company from one of its certificates ----------

PICTURES = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}


def _picture_upload(field: str) -> bytes | None:
    """An uploaded logo / stamp as PNG, None when none was chosen; CompanyError if it isn't a picture."""
    f = request.files.get(field)
    if not f or not f.filename:
        return None
    if Path(f.filename).suffix.lower() not in PICTURES:
        raise CompanyError(f"'{f.filename}' isn't a picture. Choose a .png or .jpg file")
    try:
        return to_png(f.read())
    except Exception as e:
        raise CompanyError(f"'{f.filename}' couldn't be read as a picture ({e})") from e


def _company_dir(token: str) -> Path:
    d = _upload_dir(token)
    if not (d / "draft.json").is_file():
        abort(410, "These files are no longer here. Choose the registers again.")
    return d


def _company_page(d: Path, error=None, code=200):
    """The add-company check screen: the five header rows to correct, and a sample certificate."""
    draft = Draft.from_json((d / "draft.json").read_text(encoding="utf-8"))
    logo, stamp = (p.read_bytes() if p.is_file() else None for p in (d / "logo.png", d / "stamp.png"))
    try:
        s = preview_settings(draft, d, logo, stamp)
        sheet = certificate_html(build_report(sample_row(), None, s), s, editable=False)
    except CompanyError as e:
        sheet, error = None, error or f"{str(e)[0].upper()}{str(e)[1:]}."
    notes = (d / "notes.txt").read_text(encoding="utf-8").splitlines() if (d / "notes.txt").is_file() else []
    return render_template("company_check.html", step=0, token=d.name, draft=draft, has_logo=bool(logo),
                           has_stamp=bool(stamp), sheet=sheet, error=error, notes=notes), code


@app.get("/company")
def new_company():
    return render_template("company_new.html", step=0, error=None, name=""), 200


@app.post("/company")
def read_company_sample():
    _clean_old_uploads()
    name = " ".join(request.form.get("name", "").split())
    sample = request.files.get("sample")

    def error(message):
        return render_template("company_new.html", step=0, error=message, name=name), 400

    if not name:
        return error("Type the company's name, as it should appear in the company menu.")
    if len(name) > 60:
        return error("The name is too long for the company menu (60 letters at most).")
    if name.lower() in {n.lower() for _, n in _companies()}:
        return error(f"There's already a company called {name}.")
    if not sample or not sample.filename:
        return error("Choose one of the company's test certificates (the Excel file).")
    suffix = Path(sample.filename).suffix.lower()
    if suffix not in ALLOWED:
        return error(f"'{sample.filename}' isn't an Excel file. Choose the company's certificate (.xls or .xlsx).")
    d = UPLOADS / secrets.token_hex(8)
    d.mkdir(parents=True)
    try:
        logo, stamp = _picture_upload("logo"), _picture_upload("stamp")
        sample.save(d / f"sample{suffix}")
        found = read_sample(d / f"sample{suffix}")
        logo, stamp = logo or found.logo, stamp or found.stamp
        draft = make_draft(name, found, has_stamp=bool(stamp))
    except CompanyError as e:
        shutil.rmtree(d, ignore_errors=True)
        return error(f"{str(e)[0].upper()}{str(e)[1:]}.")
    for data, file in ((logo, "logo.png"), (stamp, "stamp.png")):
        if data:
            (d / file).write_bytes(data)
    (d / "draft.json").write_text(draft.to_json(), encoding="utf-8")
    notes = [] if logo else found.notes
    (d / "notes.txt").write_text("\n".join(notes), encoding="utf-8")
    return _company_page(d)


@app.post("/company/update")
def update_company():
    d = _company_dir(request.form.get("token", ""))
    draft = Draft.from_json((d / "draft.json").read_text(encoding="utf-8"))
    draft.name = " ".join(request.form.get("name", draft.name).split()) or draft.name
    for i in range(5):
        draft.texts[i] = request.form.get(f"text{i + 1}", draft.texts[i]).replace("\r\n", "\n").strip()
        try:
            draft.sizes[i] = min(30.0, max(6.0, round(float(request.form.get(f"size{i + 1}", draft.sizes[i])) * 2) / 2))
        except ValueError:
            pass
        draft.bold[i] = f"bold{i + 1}" in request.form
    draft.tested_by = "tested_by" in request.form
    try:
        for kind in ("logo", "stamp"):
            new = _picture_upload(kind)
            if new:
                (d / f"{kind}.png").write_bytes(new)
            elif f"remove_{kind}" in request.form:
                (d / f"{kind}.png").unlink(missing_ok=True)
    except CompanyError as e:
        (d / "draft.json").write_text(draft.to_json(), encoding="utf-8")
        return _company_page(d, f"{str(e)[0].upper()}{str(e)[1:]}.", 400)
    (d / "draft.json").write_text(draft.to_json(), encoding="utf-8")
    (d / "notes.txt").unlink(missing_ok=True)
    return _company_page(d)


@app.post("/company/save")
def save_new_company():
    d = _company_dir(request.form.get("token", ""))
    draft = Draft.from_json((d / "draft.json").read_text(encoding="utf-8"))
    logo, stamp = (p.read_bytes() if p.is_file() else None for p in (d / "logo.png", d / "stamp.png"))
    try:
        cid = save_company(draft, logo, stamp)
    except CompanyError as e:
        return _company_page(d, f"{str(e)[0].upper()}{str(e)[1:]}.", 400)
    shutil.rmtree(d, ignore_errors=True)
    return redirect(url_for("choose_files", added=cid))


@app.get("/preview")
def preview():
    """One entry's certificate for the preview window, with new random test values. The page keeps
    what it shows (and any changes) and sends it with the download, so the files match the preview."""
    d = _upload_dir(request.args.get("token", ""))
    company = (d / "company.txt").read_text(encoding="utf-8")
    settings = with_company(load_settings(), company)
    row = {row_id(r): r for r in _diff(d).new if r.ok}.get(request.args.get("row", ""))
    if row is None:
        abort(400, "An entry is no longer in the register. Choose the files again.")
    number = request.args.get("no", row.report_no).strip()
    return certificate_html(build_report(row, number, settings), settings)


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

    company = (d / "company.txt").read_text(encoding="utf-8")
    settings = with_company(load_settings(), company)
    diff = _diff(d)
    by_id = {row_id(r): r for r in diff.new if r.ok}
    editable = editable_cells(settings)
    reports = []
    for rid in chosen:
        row = by_id.get(rid)
        if row is None:
            abort(400, "An entry is no longer in the register. Choose the files again.")
        number = request.form.get(f"no-{rid}", row.report_no).strip()
        rep = build_report(row, number, settings)
        # values seen (and maybe changed) in the preview: printed exactly as shown there
        apply_values(rep, {ref: request.form[f"v-{rid}-{ref}"] for ref in editable
                           if f"v-{rid}-{ref}" in request.form}, settings)
        reports.append(rep)
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

    recorded = _record(d, company, reports)
    if which == {"2"}:                     # Report 2 alone: just its Excel file, no ZIP
        resp = send_file(report2, as_attachment=True, download_name=f"Report 2 {stamp[:10]}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        resp.headers["X-Certificates"] = str(len(reports))
        resp.headers["X-Saved-To"] = str(out)
        resp.headers["X-Recorded"] = recorded
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
    resp.headers["X-Recorded"] = recorded
    return resp


def _record(d: Path, company: str, reports) -> str:
    """Add this download's entries (and the old register's, if one was compared) to the company's
    record. Returns "ok", or why the record couldn't be saved (shown on the done screen)."""
    path, old = ledger_path(LEDGER, company), _old_file(d)
    try:
        led = load(path)
    except LedgerError:
        # a damaged record is kept, renamed; a new one starts from the old register, or from the
        # copy taken when these files were checked
        set_aside(path)
        led = load(path) if old else load(d / "ledger.json")
        led.path = path
    try:
        if old:
            read = read_production(old)
            led.add(read.rows, "old register", grades=readable_grades(read))
        led.add([rep.row for rep in reports], "certified", {id(rep.row): rep.number for rep in reports})
        led.save()
    except OSError as e:
        return f"not saved ({e})"
    return "ok"


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
    return _start_page(None, getattr(e, "description", None) or "Something was missing. Choose the files again.", e.code)


def main():
    url = f"http://{HOST}:{PORT}/"
    if "--no-browser" not in sys.argv:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"\n  Test certificates is running at {url}\n  Keep this window open while you use it. Close it to stop.\n")
    app.run(host=HOST, port=PORT, debug=False)


if __name__ == "__main__":
    main()
