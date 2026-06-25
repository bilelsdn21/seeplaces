"""
SeePlaces — Live Worklist
=========================
Turns a downloaded SeePlaces report into a persistent, checkable worklist so a
data-entry worker can see, at a glance, which excursion reservations are NEW
since the last check, which were CANCELLED, and which they have already entered
into the internal system.

Work-unit  : one excursion line, keyed by `Booking code` (unique per line,
             never null). One booking number can span several excursions, so the
             line — not the booking — is what gets entered and ticked off.

State store: worklist.db (SQLite, next to this file). Survives restarts, so a
             ticked line stays "done" forever.
"""

import os
import sqlite3
from datetime import datetime

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "worklist.db")


# ── DB ──────────────────────────────────────────────────────────────────────
def _conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def _init():
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS lines (
                key            TEXT PRIMARY KEY,
                booking_number TEXT,
                voucher        TEXT,
                status         TEXT,
                excursion      TEXT,
                excursion_date TEXT,
                booking_date   TEXT,
                hotel          TEXT,
                region         TEXT,
                pax            INTEGER,
                price          REAL,
                currency       TEXT,
                payment        TEXT,
                customer       TEXT,
                first_seen     TEXT,
                last_seen      TEXT,
                is_new         INTEGER DEFAULT 0,
                status_changed INTEGER DEFAULT 0,
                entered        INTEGER DEFAULT 0,
                entered_at     TEXT
            )
        """)
        c.execute("CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)")


_init()


def _meta_set(k, v):
    with _conn() as c:
        c.execute("INSERT INTO meta(k,v) VALUES(?,?) "
                  "ON CONFLICT(k) DO UPDATE SET v=excluded.v", (k, str(v)))


def _meta_get(k, default=None):
    with _conn() as c:
        r = c.execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
        return r["v"] if r else default


# ── Parsing ─────────────────────────────────────────────────────────────────
def _find_tnd_price_col(df: pd.DataFrame) -> str:
    """Gross-price column whose adjacent currency column holds TND."""
    cols = list(df.columns)
    for i, col in enumerate(cols):
        if "gross price" in str(col).lower() and i + 1 < len(cols):
            sample = df[cols[i + 1]].dropna().astype(str).str.strip().str.upper()
            if (sample == "TND").any():
                return col
    # fallback: first gross-price column
    for col in cols:
        if "gross price" in str(col).lower():
            return col
    raise ValueError("No 'Gross Price' column found in report.")


def _s(row, col, default=""):
    if col not in row or pd.isna(row[col]):
        return default
    return str(row[col]).strip()


def parse_lines(xlsx_path: str) -> list[dict]:
    """Read a SeePlaces report into a list of normalized excursion-line dicts."""
    df = pd.read_excel(xlsx_path)
    if "Booking number" not in df.columns:
        raise ValueError("This file does not look like a SeePlaces report "
                         "(missing 'Booking number' column).")

    price_col = _find_tnd_price_col(df)
    lines, seen = [], set()

    for _, r in df.iterrows():
        booking_code = _s(r, "Booking code")
        booking_no   = _s(r, "Booking number")
        excursion    = _s(r, "Excursion name")
        exc_date_raw = r.get("ExcursionDate")

        # Stable per-line key. Booking code is unique per line; fall back to a
        # composite if a file ever lacks it.
        key = booking_code or f"{booking_no}|{excursion}|{exc_date_raw}"
        if not key.strip() or key in seen:
            continue
        seen.add(key)

        status = _s(r, "Status").lower()
        status = "Cancelled" if status in ("canceled", "cancelled") else \
                 ("Confirmed" if status == "confirmed" else _s(r, "Status"))

        exc_date = pd.to_datetime(exc_date_raw, errors="coerce")
        bk_date  = pd.to_datetime(r.get("Booking date"), errors="coerce")

        first = _s(r, "Participant first name")
        last  = _s(r, "Participant last name")
        customer = (first + " " + last).strip()

        try:
            pax = int(pd.to_numeric(r.get("Participants count"), errors="coerce"))
        except (ValueError, TypeError):
            pax = 0
        price = pd.to_numeric(r.get(price_col), errors="coerce")
        price = float(price) if pd.notna(price) else 0.0

        lines.append({
            "key":            key,
            "booking_number": booking_no,
            "voucher":        _s(r, "VoucherNumber"),
            "status":         status,
            "excursion":      excursion or "(no excursion name)",
            "excursion_date": exc_date.strftime("%Y-%m-%d %H:%M") if pd.notna(exc_date) else "",
            "booking_date":   bk_date.strftime("%Y-%m-%d %H:%M") if pd.notna(bk_date) else "",
            "hotel":          _s(r, "Pickup name"),
            "region":         _s(r, "Region"),
            "pax":            pax,
            "price":          round(price, 2),
            "currency":       "TND",
            "payment":        _s(r, "Payment method"),
            "customer":       customer,
        })
    return lines


# ── Diff / sync ─────────────────────────────────────────────────────────────
_UPDATE_FIELDS = ("booking_number", "voucher", "status", "excursion",
                  "excursion_date", "booking_date", "hotel", "region",
                  "pax", "price", "currency", "payment", "customer")


def sync(lines: list[dict]) -> dict:
    """
    Merge freshly parsed lines into the store.
      • brand-new key            → inserted, flagged is_new
      • Confirmed → Cancelled    → flagged status_changed (a CANCEL alert)
    Returns a summary of what changed. Existing 'entered' ticks are preserved.
    """
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_lines, cancelled_lines = [], []

    with _conn() as c:
        existing = {row["key"]: row for row in c.execute(
            "SELECT key, status, entered FROM lines").fetchall()}

        # A new sync resets the "is_new" highlight from the previous round.
        c.execute("UPDATE lines SET is_new=0")

        for ln in lines:
            key = ln["key"]
            if key not in existing:
                c.execute(
                    """INSERT INTO lines
                       (key, booking_number, voucher, status, excursion,
                        excursion_date, booking_date, hotel, region, pax, price,
                        currency, payment, customer, first_seen, last_seen,
                        is_new, status_changed, entered, entered_at)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,0,0,NULL)""",
                    (key, ln["booking_number"], ln["voucher"], ln["status"],
                     ln["excursion"], ln["excursion_date"], ln["booking_date"],
                     ln["hotel"], ln["region"], ln["pax"], ln["price"],
                     ln["currency"], ln["payment"], ln["customer"], now, now))
                new_lines.append(ln)
            else:
                prev = existing[key]
                became_cancelled = (
                    prev["status"] == "Confirmed" and ln["status"] == "Cancelled")
                set_clause = ", ".join(f"{f}=?" for f in _UPDATE_FIELDS)
                params = [ln[f] for f in _UPDATE_FIELDS]
                params += [now]  # last_seen
                if became_cancelled:
                    c.execute(
                        f"UPDATE lines SET {set_clause}, last_seen=?, "
                        f"status_changed=1 WHERE key=?", params + [key])
                    cancelled_lines.append(ln)
                else:
                    c.execute(
                        f"UPDATE lines SET {set_clause}, last_seen=? WHERE key=?",
                        params + [key])

    _meta_set("last_sync", now)
    return {
        "new_count":       len(new_lines),
        "cancelled_count": len(cancelled_lines),
        "new":             new_lines,
        "cancelled":       cancelled_lines,
        "last_sync":       now,
    }


# ── Read for UI ─────────────────────────────────────────────────────────────
def _row_to_dict(row: sqlite3.Row) -> dict:
    d = dict(row)
    d["is_new"]         = bool(d.get("is_new"))
    d["status_changed"] = bool(d.get("status_changed"))
    d["entered"]        = bool(d.get("entered"))
    return d


def get_worklist() -> dict:
    """
    Returns the open worklist (anything not yet entered, plus any line that
    flipped to Cancelled so the worker can remove it), grouped by excursion.
    New lines and cancel-alerts float to the top.
    """
    with _conn() as c:
        rows = c.execute(
            """SELECT * FROM lines
               WHERE entered=0 OR status_changed=1
               ORDER BY is_new DESC, status_changed DESC, booking_date DESC"""
        ).fetchall()
        done_count = c.execute(
            "SELECT COUNT(*) n FROM lines WHERE entered=1").fetchone()["n"]

    open_lines = [_row_to_dict(r) for r in rows]

    # Group by excursion, keeping the new/alert-first ordering.
    groups, order = {}, []
    for ln in open_lines:
        g = ln["excursion"]
        if g not in groups:
            groups[g] = {"excursion": g, "lines": [], "pax": 0,
                         "new": 0, "cancelled": 0}
            order.append(g)
        grp = groups[g]
        grp["lines"].append(ln)
        if ln["status"] != "Cancelled":
            grp["pax"] += ln["pax"] or 0
        if ln["is_new"]:
            grp["new"] += 1
        if ln["status"] == "Cancelled":
            grp["cancelled"] += 1

    grouped = [groups[g] for g in order]
    grouped.sort(key=lambda x: (-x["new"], -x["cancelled"], x["excursion"]))

    return {
        "groups":      grouped,
        "open_count":  len(open_lines),
        "new_count":   sum(1 for l in open_lines if l["is_new"]),
        "cancel_count": sum(1 for l in open_lines if l["status"] == "Cancelled"),
        "open_pax":    sum(l["pax"] or 0 for l in open_lines
                           if l["status"] != "Cancelled"),
        "done_count":  done_count,
        "last_sync":   _meta_get("last_sync", ""),
        "last_error":  _meta_get("last_error", ""),
        "last_status": _meta_get("last_status", ""),
    }


def set_entered(key: str, entered: bool) -> dict:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if entered else None
    with _conn() as c:
        # Ticking a line also acknowledges a cancel-alert on it.
        c.execute(
            "UPDATE lines SET entered=?, entered_at=?, "
            "status_changed=CASE WHEN ?=1 THEN 0 ELSE status_changed END, "
            "is_new=CASE WHEN ?=1 THEN 0 ELSE is_new END "
            "WHERE key=?",
            (1 if entered else 0, now, 1 if entered else 0,
             1 if entered else 0, key))
    return {"ok": True}
