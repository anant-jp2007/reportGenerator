"""The record of what was certified, one file per company: ledger/<company id>.json.

With it, only the register as it is now needs uploading: every entry not in the record is new.
It uses the same matching as two registers (production.not_covered): key = grade + financial year
+ Sr. No., an identical entry first, each recorded entry covering one row.

The record is filled
- after every download, with the entries that got a certificate (and the Report No. used);
- from an old register, when one is compared: all its entries count as done, as they do in the
  comparison itself. That is how a company starts (its record is empty) and how a lost record
  is rebuilt.
Entries already in the record are never added twice, so downloading the same entries again
(another format, Report 2 later) changes nothing.

A grade only counts once the record has seen its sheet: a register with a grade the record doesn't
know (a renamed sheet, a new line) gets that grade skipped with a message, as two registers with
different sheets do.
"""
from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from app.production import GRADE_ORDER, Diff, ReadResult, Row, not_covered, read_production

VERSION = 1


class LedgerError(ValueError):
    """The record file can't be read; the message is shown to the user."""


@dataclass
class Ledger:
    path: Path
    grades: set[str] = field(default_factory=set)
    entries: list[dict] = field(default_factory=list)

    @property
    def started(self) -> bool:
        """False for a company never compared with an old register: everything would look new."""
        return bool(self.grades)

    @property
    def last(self) -> str | None:
        """When the last certificate was recorded, like 31-10-2026."""
        times = [e["added"] for e in self.entries if e["source"] == "certified"]
        return datetime.fromisoformat(max(times)).strftime("%d-%m-%Y") if times else None

    @property
    def certified(self) -> int:
        return sum(e["source"] == "certified" for e in self.entries)

    def done(self) -> list[tuple[tuple, tuple]]:
        out = []
        for e in self.entries:
            key = (e["grade"], e["fy"], e["sr"])
            d = date.fromisoformat(e["date"]) if e["date"] else None
            out.append((key, (key, d, e["size"], e["batch"])))
        return out

    def add(self, rows: list[Row], source: str, numbers: dict[int, str] | None = None,
            grades: set[str] | None = None) -> int:
        """Record the rows not recorded yet; numbers: id(row) -> Report No. used; grades: the sheets
        seen (an old register's, even empty ones), else the rows' grades. Returns how many were added."""
        added = datetime.now().isoformat(timespec="seconds")
        fresh = not_covered(self.done(), rows)
        for r in fresh:
            grade, fy, sr = r.key
            self.entries.append({"grade": grade, "fy": fy, "sr": sr,
                                 "date": r.date.isoformat() if r.date else None,
                                 "size": r.size_text, "batch": r.batch,
                                 "report_no": str((numbers or {}).get(id(r), r.report_no)),
                                 "source": source, "added": added})
        self.grades |= grades if grades is not None else {r.grade for r in rows}
        return len(fresh)

    def save(self) -> None:
        """Written to a temporary file first, then swapped in, so a crash never leaves half a file;
        the previous version is kept as <id>.previous.json."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        data = {"version": VERSION, "grades": [g for g in GRADE_ORDER if g in self.grades],
                "entries": self.entries}
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
        if self.path.exists():
            shutil.copyfile(self.path, self.path.with_name(f"{self.path.stem}.previous.json"))
        os.replace(tmp, self.path)


def readable_grades(read: ReadResult) -> set[str]:
    """The grades whose sheet was read: an unreadable sheet's entries aren't known, so its grade
    must not count as recorded (every entry of it would look new next time)."""
    return read.grades - read.errors.keys()


def ledger_path(folder: Path, company: str) -> Path:
    return Path(folder) / f"{company}.json"


def load(path: Path) -> Ledger:
    """The record at path; empty when there's none yet. A damaged file raises LedgerError: treating
    it as empty would make every entry look new."""
    path = Path(path)
    if not path.exists():
        return Ledger(path)
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        entries = data["entries"]
        for e in entries:
            for k in ("grade", "fy", "sr", "date", "size", "batch", "report_no", "source", "added"):
                e[k]
        return Ledger(path, set(data["grades"]), entries)
    except (OSError, ValueError, KeyError, TypeError) as e:
        raise LedgerError(f"the record of certified entries ({path.name}) can't be read ({e})") from e


def set_aside(path: Path) -> None:
    """Rename a damaged record so a new one can be started from an old register."""
    path = Path(path)
    if path.exists():
        path.rename(path.with_name(f"{path.stem}.damaged {datetime.now():%Y-%m-%d %H%M%S}.json"))


def find_new_in_ledger(ledger: Ledger, register) -> Diff:
    """The entries of the register that aren't in the record, like find_new_rows with an old register."""
    now = read_production(register)
    errors = [f"register: {e}" for e in now.errors.values()]
    if not now.grades:
        errors.append("no PE 63 / PE 80 / PE 100 sheets found in the register (sheets named like "
                      "'Production 63' or 'PE 63')")
    for g in GRADE_ORDER:
        if g in now.grades and g not in ledger.grades:
            errors.append(f"{g} isn't in the record of certified entries yet, so {g} can't be checked: "
                          f"compare with an old register that has the {g} sheet once")
        elif g in ledger.grades and g not in now.grades:
            errors.append(f"register: no sheet for {g}, so {g} can't be checked")
    blocked = now.errors.keys() | (now.grades - ledger.grades)
    fresh = not_covered(ledger.done(), [r for r in now.rows if r.grade not in blocked])
    fresh.sort(key=lambda r: (GRADE_ORDER.index(r.grade), r.excel_row))
    return Diff(fresh, errors, [f"register: {n}" for n in now.notes])
