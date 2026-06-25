"""
SeePlaces — Weekly Guide Report
================================
Groups confirmed bookings by guide based on hotel (Pickup name).
Outputs one Excel with:
  • Guide Report sheet  — per-guide sections with booking rows + totals
  • By Region sheet     — bookings grouped by geographic region
  • Summary sheet       — KPI overview (generated first)

Usage:
    python guide_report.py <seeplaces_report.xlsx> <output.xlsx>
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import sys
import os
import json

# ── Guide email → name map (authoritative — from SeePlaces Expedient email col)
GUIDE_EMAIL_MAP = {
    "katarzyna.jaszcz@rep.itaka.pl": "kasha",
    "aneta.dymek@rep.itaka.pl":      "aneta",
    "joanna.kisiel@rep.itaka.pl":    "joanna",
}

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assignment_config.json")

_DEFAULT_CONFIG = {
    "region_guide": {
        "Hammamet": "joanna",
        "Korba":    "joanna",
        "Mahdia":   "aneta",
        "Monastir": "aneta",
        "Sousse":   "kasha",
    },
    "hotel_overrides": {
        "Iberostar Selection Kantaoui Bay": "aneta",
        "Marhaba Beach":                   "aneta",
        "Marhaba Royal Salem":             "aneta",
        "Marhaba Salem":                   "aneta",
        "Africa Jade Thalasso":            "kasha",
        "El Mouradi Palace":               "joanna",
    },
}

# ── Region mapping ────────────────────────────────────────────────────────────
_REGION_MAP = {
    "Africa Jade Thalasso":               "Korba",
    "Khayam Garden Beach & Spa":          "Korba",
    "Laico Hammamet":                     "Hammamet",
    "Le Royal Hammamet Hotel & Resort":   "Hammamet",
    "Mediterranee Thalasso Golf":         "Hammamet",
    "Occidental Marco Polo":              "Hammamet",
    "Regency Hammamet":                   "Hammamet",
    "Sentido Marillia Resort & Spa":      "Hammamet",
    "Yadis Hammamet":                     "Hammamet",
    "Barcelo Concorde Green Park Palace": "Sousse",
    "Concorde Green Park":                "Sousse",
    "El Mouradi Palace":                  "Sousse",
    "Iberostar Diar El Andalous":         "Sousse",
    "Iberostar Selection Kantaoui Bay":   "Sousse",
    "Jaz Tour Khalef":                    "Sousse",
    "Marhaba Beach":                      "Sousse",
    "Marhaba Club":                       "Sousse",
    "Marhaba Palace Sousse":              "Sousse",
    "Marhaba Royal Salem":                "Sousse",
    "Marhaba Salem":                      "Sousse",
    "Occidental Sousse Marhaba":          "Sousse",
    "Steigenberger Marhaba Palace":       "Sousse",
    "El Mouradi Mahdia":                  "Mahdia",
    "Thalassa Mahdia":                    "Mahdia",
    "Regency Monastir":                   "Monastir",
    "Royal Thalassa Monastir":            "Monastir",
}

_REGION_KW = [
    ("hammamet", "Hammamet"),
    ("nabeul",   "Korba"),
    ("korba",    "Korba"),
    ("mahdia",   "Mahdia"),
    ("monastir", "Monastir"),
    ("kantaoui", "Sousse"),
    ("marhaba",  "Sousse"),
    ("sousse",   "Sousse"),
]

_REGION_NORM = {k.lower().strip(): v for k, v in _REGION_MAP.items()}


def _get_region(hotel: str) -> str:
    if not hotel or str(hotel).lower().strip() in ("", "nan", "none"):
        return "Other"
    h = str(hotel).lower().strip()
    if h in _REGION_NORM:
        return _REGION_NORM[h]
    for kw, region in _REGION_KW:
        if kw in h:
            return region
    return "Other"


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    save_config(_DEFAULT_CONFIG)
    return _DEFAULT_CONFIG.copy()


def save_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def _assign_guide_by_region(region: str, hotel: str, cfg: dict) -> str:
    h = str(hotel).lower().strip()
    r = str(region).lower().strip()
    overrides = {k.lower().strip(): v for k, v in cfg.get("hotel_overrides", {}).items()}
    if h in overrides:
        return overrides[h]
    region_guide = {k.lower().strip(): v for k, v in cfg.get("region_guide", {}).items()}
    return region_guide.get(r, "unassigned")


def _assign_guide_row(payment: str, hotel: str, email: str, region: str, cfg: dict) -> str:
    """Offline bookings → expedient email. Online/unknown → region + override logic."""
    if str(payment).strip().lower() != "online":
        guide = GUIDE_EMAIL_MAP.get(str(email).strip().lower())
        if guide:
            return guide
    return _assign_guide_by_region(region, hotel, cfg)


# ── Style helpers ─────────────────────────────────────────────────────────────
def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="1E293B", size=9, italic=False) -> Font:
    return Font(bold=bold, color=color, size=size, italic=italic)

def _align(h="left", v="center", indent=0) -> Alignment:
    return Alignment(horizontal=h, vertical=v, indent=indent)


GUIDE_CONFIG = {
    "kasha":  {"bg": "4F46E5", "label": "KASHA"},
    "aneta":  {"bg": "0891B2", "label": "ANETA"},
    "joanna": {"bg": "059669", "label": "JOANNA"},
}

REGION_COLORS = {
    "Hammamet":         "7C3AED",
    "Nabeul":           "9333EA",
    "Sousse":           "0E7490",
    "Port El Kantaoui": "0369A1",
    "Mahdia":           "B45309",
    "Monastir":         "0F766E",
    "Other":            "64748B",
}

COLUMNS    = ["#", "Booking number", "VoucherNumber", "Status", "Excursion name", "Excursion Date", "Booking Date", "Gross Price", "Currency", "Payment method", "Hotel", "Pax", "Adults", "Child", "Guide"]
COL_WIDTHS = [5,   14,              13,              11,       42,               18,               18,             14,            10,         14,               34,       6,     8,        8,       10]

COL_HDR_BG = "E2E8F0"
ROW_ODD    = "FFFFFF"
ROW_EVEN   = "F8FAFC"
TOTAL_BG   = "FEF08A"
WHITE      = "FFFFFF"
MUTED      = "94A3B8"
UNASSIGNED = "64748B"
GUIDES     = ["kasha", "aneta", "joanna"]


# ── Data loading ──────────────────────────────────────────────────────────────
def _find_tnd_price_col(df: pd.DataFrame) -> str:
    """Find the Gross Price column whose adjacent currency column is TND."""
    cols = list(df.columns)
    for i, col in enumerate(cols):
        if "gross price" in str(col).lower() and i + 1 < len(cols):
            currency_col = cols[i + 1]
            sample = df[currency_col].dropna().astype(str).str.strip().str.upper()
            if (sample == "TND").any():
                return col
    raise ValueError("Could not find a Gross Price column with TND currency in the report.")


def load_report(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)

    df["status_norm"] = df["Status"].astype(str).str.strip().str.lower()
    df = df[df["status_norm"].isin(["confirmed", "canceled", "cancelled"])].copy()

    def _col(name, default=""):
        return df[name].astype(str).str.strip() if name in df.columns else pd.Series(default, index=df.index)

    df["idx"]          = pd.to_numeric(df["Index"], errors="coerce") if "Index" in df.columns else pd.Series(range(len(df)), index=df.index)
    df["booking_id"]   = df["Booking number"].astype(str).str.strip()
    df["voucher"]      = _col("VoucherNumber")
    df["hotel"]        = df["Pickup name"].astype(str).str.strip()
    df["confirmed"]    = df["status_norm"] == "confirmed"
    price_col          = _find_tnd_price_col(df)
    df["price"]        = pd.to_numeric(df[price_col], errors="coerce").fillna(0)
    df.loc[~df["confirmed"], "price"] = 0
    df["excursion"]    = df["Excursion name"].astype(str).str.strip()
    df["excursion_date"] = pd.to_datetime(df["ExcursionDate"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
    df["booking_date"] = pd.to_datetime(df.get("Booking date"), errors="coerce").dt.strftime("%Y-%m-%d %H:%M") if "Booking date" in df.columns else pd.Series("", index=df.index)
    df["payment"]      = df["Payment method"].astype(str).str.strip()
    df["pax_total"]    = pd.to_numeric(df.get("Participants count", 0), errors="coerce").fillna(0).astype(int)
    df["pax_adult"]    = pd.to_numeric(df.get("Adult", 0),              errors="coerce").fillna(0).astype(int)
    df["pax_child"]    = (
        pd.to_numeric(df["Child"], errors="coerce").fillna(0).astype(int)
        if "Child" in df.columns
        else (df["pax_total"] - df["pax_adult"]).clip(lower=0)
    )

    # Region: use column if present, else derive from hotel name
    df["region"] = _col("Region")
    empty_mask = df["region"].str.lower().isin(["", "nan", "none"])
    df.loc[empty_mask, "region"] = df.loc[empty_mask, "hotel"].apply(_get_region)

    # Guide assignment — use expedient email if available, else region logic
    exp_email = _col("Expedient email")
    cfg = load_config()
    df["guide"] = df.apply(
        lambda r: _assign_guide_row(r["Payment method"], r["hotel"],
                                    exp_email[r.name], r["region"], cfg), axis=1
    )

    return df[["idx", "booking_id", "voucher", "confirmed", "excursion", "excursion_date",
               "booking_date", "price", "payment", "hotel", "pax_total",
               "pax_adult", "pax_child", "guide", "region"]]


# ── Excel builder — Guide Report sheet ───────────────────────────────────────
def _write_guide_sheet(ws, df: pd.DataFrame):
    # Column indices (1-based):  #  BkNo  Vchr  St  Excursion  ExcDate  BkDate  Price  Cur  Pay  Hotel  Pax  Adults  Guide
    n         = len(COLUMNS)
    PRICE_COL = 8
    RIGHT_COLS = {PRICE_COL, 12, 13, 14}

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1

    def _write_headers(guide_bg):
        nonlocal row
        for ci, hdr in enumerate(COLUMNS, 1):
            c      = ws.cell(row=row, column=ci, value=hdr)
            c.font = _font(bold=True, color="1E293B", size=8)
            c.fill = _fill(COL_HDR_BG)
            c.alignment = _align(h="right" if ci in RIGHT_COLS else "center")
            if ci == n:
                c.fill = _fill("CBD5E1")
        ws.row_dimensions[row].height = 17
        row += 1

    def _write_booking_row(bk, bg, guide_accent):
        nonlocal row
        is_confirmed = bool(bk["confirmed"])
        row_bg  = bg if is_confirmed else "FEE2E2"
        status  = "Confirmed" if is_confirmed else "Cancelled"
        vals = [
            int(bk["idx"]) if pd.notna(bk["idx"]) else "",
            bk["booking_id"],
            bk["voucher"],
            status,
            bk["excursion"],
            bk["excursion_date"],
            bk["booking_date"],
            bk["price"] if is_confirmed else "",
            "TND" if is_confirmed else "",
            bk["payment"],
            bk["hotel"],
            bk["pax_total"] if bk["pax_total"] > 0 else "",
            bk["pax_adult"] if bk["pax_adult"] > 0 else "",
            bk["pax_child"] if bk["pax_child"] > 0 else "",
            bk["guide"].upper(),
        ]
        for ci, val in enumerate(vals, 1):
            c      = ws.cell(row=row, column=ci, value=val)
            c.font = _font(size=9, color="94A3B8" if not is_confirmed else "1E293B",
                           italic=not is_confirmed)
            c.alignment = _align(h="right" if ci in RIGHT_COLS else "left", indent=1)
            if ci == n:
                c.fill = _fill(guide_accent if is_confirmed else "CBD5E1")
                c.font = _font(bold=True, color=WHITE if is_confirmed else "64748B", size=8)
                c.alignment = _align(h="center")
            else:
                c.fill = _fill(row_bg)
            if ci == PRICE_COL:
                c.number_format = "#,##0.00"
        ws.row_dimensions[row].height = 15
        row += 1

    def _write_total_row(guide_df, cfg):
        nonlocal row
        confirmed_df = guide_df[guide_df["confirmed"]]
        total = confirmed_df["price"].sum()
        count = len(confirmed_df)
        for ci in range(1, n + 1):
            c      = ws.cell(row=row, column=ci)
            c.fill = _fill(TOTAL_BG)
            c.font = _font(bold=True, size=9, color="1E293B")
            if ci == 1:
                c.value     = f"{cfg['label']}  ·  {count} booking{'s' if count != 1 else ''}"
                c.alignment = _align(h="left", indent=1)
            elif ci == PRICE_COL:
                c.value         = total
                c.number_format = "#,##0.00"
                c.alignment     = _align(h="right")
            elif ci == PRICE_COL + 1:
                c.value     = "TND"
                c.alignment = _align(h="left", indent=1)
        ws.row_dimensions[row].height = 18
        row += 1

    def write_guide_section(guide: str):
        nonlocal row
        cfg      = GUIDE_CONFIG[guide]
        guide_df = df[df["guide"] == guide].sort_values("excursion_date")

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
        c           = ws.cell(row=row, column=1)
        c.value     = f"   {cfg['label']}"
        c.font      = _font(bold=True, color=WHITE, size=13)
        c.fill      = _fill(cfg["bg"])
        c.alignment = _align(h="left", v="center")
        ws.row_dimensions[row].height = 30
        row += 1

        if guide_df.empty:
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
            c           = ws.cell(row=row, column=1)
            c.value     = "No bookings for this period"
            c.font      = _font(italic=True, color=MUTED)
            c.fill      = _fill("F1F5F9")
            c.alignment = _align(h="left", indent=2)
            ws.row_dimensions[row].height = 18
            row += 1
        else:
            _write_headers(cfg["bg"])
            for i, (_, bk) in enumerate(guide_df.iterrows()):
                _write_booking_row(bk, ROW_ODD if i % 2 == 0 else ROW_EVEN, cfg["bg"])
            _write_total_row(guide_df, cfg)

        ws.row_dimensions[row].height = 10
        row += 1

    for guide in GUIDES:
        write_guide_section(guide)

    confirmed   = df[df["guide"].isin(GUIDES) & df["confirmed"]]
    grand_total = confirmed["price"].sum()
    grand_count = len(confirmed)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    c           = ws.cell(row=row, column=1)
    c.value     = f"   GRAND TOTAL  ·  {grand_count} booking{'s' if grand_count != 1 else ''}   {grand_total:,.2f} TND"
    c.font      = _font(bold=True, color=WHITE, size=11)
    c.fill      = _fill("1E293B")
    c.alignment = _align(h="left", v="center")
    ws.row_dimensions[row].height = 28
    row += 1

    unassigned = df[df["guide"] == "unassigned"]
    if not unassigned.empty:
        ws.row_dimensions[row].height = 10
        row += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
        c           = ws.cell(row=row, column=1)
        c.value     = "   UNASSIGNED — region not recognized"
        c.font      = _font(bold=True, color=WHITE, size=11)
        c.fill      = _fill(UNASSIGNED)
        c.alignment = _align(h="left", v="center")
        ws.row_dimensions[row].height = 26
        row += 1
        _write_headers(UNASSIGNED)
        for i, (_, bk) in enumerate(unassigned.iterrows()):
            _write_booking_row(bk, ROW_ODD if i % 2 == 0 else ROW_EVEN, UNASSIGNED)


# ── Excel builder — By Region sheet ──────────────────────────────────────────
def _write_region_sheet(ws, df: pd.DataFrame):
    reg_cols   = ["Region", "Hotel", "Guide", "Bookings", "Total TND"]
    reg_widths = [22, 42, 12, 12, 18]
    for i, w in enumerate(reg_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row     = 1
    n       = len(reg_cols)
    regions = sorted(df["region"].unique(), key=lambda r: (r == "Other", r))

    for region in regions:
        rdf   = df[df["region"] == region]
        color = REGION_COLORS.get(region, REGION_COLORS["Other"])

        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
        c           = ws.cell(row=row, column=1)
        c.value     = region.upper()
        c.font      = _font(bold=True, color=WHITE, size=11)
        c.fill      = _fill(color)
        c.alignment = _align(h="left", indent=1)
        ws.row_dimensions[row].height = 22
        row += 1

        for ci, hdr in enumerate(reg_cols, 1):
            c           = ws.cell(row=row, column=ci, value=hdr)
            c.font      = _font(bold=True, color="475569", size=8)
            c.fill      = _fill(COL_HDR_BG)
            c.alignment = _align(h="center")
        ws.row_dimensions[row].height = 15
        row += 1

        hotel_groups = (
            rdf.groupby("hotel")
               .agg(guide=("guide", "first"), count=("booking_id", "count"), total=("price", "sum"))
               .reset_index()
               .sort_values("hotel")
        )

        for idx, (_, hrow) in enumerate(hotel_groups.iterrows()):
            bg       = ROW_ODD if idx % 2 == 0 else ROW_EVEN
            row_vals = [region, hrow["hotel"], hrow["guide"].upper(), int(hrow["count"]), hrow["total"]]
            for ci, val in enumerate(row_vals, 1):
                c           = ws.cell(row=row, column=ci, value=val)
                c.fill      = _fill(bg)
                c.font      = _font()
                c.alignment = _align(h="right" if ci >= 4 else "left", indent=1)
                if ci == 5:
                    c.number_format = "#,##0.00"
            ws.row_dimensions[row].height = 15
            row += 1

        for ci in range(1, n + 1):
            c      = ws.cell(row=row, column=ci)
            c.fill = _fill(TOTAL_BG)
            c.font = _font(bold=True)
            if ci == n - 1:
                c.value     = f"SUBTOTAL  ({len(rdf)} booking{'s' if len(rdf) != 1 else ''})"
                c.alignment = _align(h="right")
            elif ci == n:
                c.value         = rdf["price"].sum()
                c.number_format = "#,##0.00"
                c.alignment     = _align(h="right")
        ws.row_dimensions[row].height = 16
        row += 1

        ws.row_dimensions[row].height = 8
        row += 1


# ── Excel builder — Summary sheet ────────────────────────────────────────────
def _write_summary_sheet(ws, meta: dict, df: pd.DataFrame):
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 20

    row = 1

    # Title
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    c           = ws.cell(row=row, column=1, value="SeePlaces Guide Report — Summary")
    c.font      = _font(bold=True, color=WHITE, size=14)
    c.fill      = _fill("4F46E5")
    c.alignment = _align(indent=1)
    ws.row_dimensions[row].height = 32
    row += 1

    if meta.get("date_from") or meta.get("date_to"):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c       = ws.cell(row=row, column=1,
                          value=f"Period: {meta.get('date_from','')} → {meta.get('date_to','')}")
        c.font  = _font(italic=True, color="94A3B8", size=9)
        c.fill  = _fill("0F172A")
        c.alignment = _align(indent=1)
        ws.row_dimensions[row].height = 16
        row += 1

    row += 1

    def section_header(label):
        nonlocal row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c           = ws.cell(row=row, column=1, value=label)
        c.font      = _font(bold=True, color=WHITE, size=10)
        c.fill      = _fill("1E293B")
        c.alignment = _align(indent=1)
        ws.row_dimensions[row].height = 20
        row += 1

        for ci, hdr in enumerate(["Category", "Count", "Total TND"], 1):
            c           = ws.cell(row=row, column=ci, value=hdr)
            c.font      = _font(bold=True, color="475569", size=8)
            c.fill      = _fill(COL_HDR_BG)
            c.alignment = _align(h="center")
        ws.row_dimensions[row].height = 14
        row += 1

    def data_row(label, count, total, bg="FFFFFF"):
        nonlocal row
        for ci, val in enumerate([label, count, total], 1):
            c           = ws.cell(row=row, column=ci, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(ci == 1))
            c.alignment = _align(h="right" if ci > 1 else "left", indent=1)
            if ci == 3 and isinstance(val, (int, float)):
                c.number_format = "#,##0.00"
        ws.row_dimensions[row].height = 16
        row += 1

    section_header("GUIDE BREAKDOWN")
    guide_bgs = {"kasha": "EEF2FF", "aneta": "ECFEFF", "joanna": "ECFDF5"}
    for g in GUIDES:
        s = meta.get(g, {"count": 0, "total": 0})
        data_row(g.upper(), s["count"], float(s["total"]), guide_bgs.get(g, "FFFFFF"))

    grand     = sum(meta.get(g, {}).get("count", 0) for g in GUIDES)
    grand_rev = sum(meta.get(g, {}).get("total", 0.0) for g in GUIDES)
    data_row("GRAND TOTAL", grand, float(grand_rev), TOTAL_BG)

    if meta.get("cancelled"):
        data_row("Cancelled (excluded)", meta["cancelled"], "", "FFF1F2")

    row += 1
    section_header("REGION BREAKDOWN")
    regions = meta.get("regions", {})
    for i, (region, rdata) in enumerate(sorted(regions.items(), key=lambda x: -x[1]["count"])):
        bg = ROW_ODD if i % 2 == 0 else ROW_EVEN
        data_row(region, rdata["count"], float(rdata["total"]), bg)


# ── Main Excel builder ────────────────────────────────────────────────────────
def build_excel(df: pd.DataFrame, out_path: str, meta: dict = None):
    wb       = openpyxl.Workbook()
    ws_guide = wb.active
    ws_guide.title = "Guide Report"

    ws_summary = wb.create_sheet("Summary", 0)
    ws_region  = wb.create_sheet("By Region")

    _write_guide_sheet(ws_guide, df)
    _write_region_sheet(ws_region, df)
    _write_summary_sheet(ws_summary, meta or {}, df)

    wb.active = ws_summary
    wb.save(out_path)


# ── Entry points ──────────────────────────────────────────────────────────────
def run(sp_path: str, out_path: str,
        date_from: str = None, date_to: str = None) -> dict:
    """Called programmatically (e.g. from the GUI). Returns summary dict."""
    print("Loading SeePlaces report…")

    df              = load_report(sp_path)
    total_confirmed = int(df["confirmed"].sum())
    cancelled_count = int((~df["confirmed"]).sum())
    print(f"  {total_confirmed} confirmed / {cancelled_count} cancelled")

    print("Building guide report…")

    conf_df       = df[df["confirmed"]]
    summary: dict = {"date_from": date_from or "", "date_to": date_to or ""}
    grand_total   = 0.0
    for guide in GUIDES:
        gdf   = conf_df[conf_df["guide"] == guide]
        total = round(float(gdf["price"].sum()), 2)
        grand_total += total
        summary[guide] = {"count": len(gdf), "total": total}
        print(f"  {guide.upper():<10}  {len(gdf):>3} booking(s)   {total:>10,.2f} TND")

    unassigned = conf_df[conf_df["guide"] == "unassigned"]
    summary["unassigned"] = {
        "count": len(unassigned),
        "total": round(float(unassigned["price"].sum()), 2),
    }

    regions: dict = {}
    for region, rdf in conf_df.groupby("region"):
        regions[region] = {
            "count": len(rdf),
            "total": round(float(rdf["price"].sum()), 2),
        }
    summary["regions"]        = regions
    summary["cancelled"]      = cancelled_count
    summary["grand_total"]    = round(grand_total, 2)
    summary["total_bookings"] = total_confirmed

    if not unassigned.empty:
        print(f"  {'UNASSIGNED':<10}  {len(unassigned):>3} booking(s)   {unassigned['price'].sum():>10,.2f} TND")
    print(f"  {'-'*42}")
    print(f"  {'TOTAL':<10}  {total_confirmed:>3} booking(s)   {grand_total:>10,.2f} TND")

    build_excel(df, out_path, meta=summary)

    try:
        import history as hist
        hist.save_run(summary, date_from=date_from, date_to=date_to)
    except Exception as e:
        print(f"  [history] {e}")

    return summary


def analyze(path: str, date_from: str = None, date_to: str = None, channel: str = "all") -> dict:
    """
    Returns rich analytics dict from a SeePlaces Excel file.
    Does NOT generate an output file — only returns data for the UI charts.
    """
    df_full = pd.read_excel(path)

    cancelled = int((df_full["Status"].str.strip().str.lower() != "confirmed").sum())

    df = df_full[df_full["Status"].str.strip().str.lower() == "confirmed"].copy()
    df["booking_id"]     = df["Booking number"].astype(str).str.strip()
    df["hotel"]          = df["Pickup name"].astype(str).str.strip()
    df["price"]          = pd.to_numeric(df[_find_tnd_price_col(df)], errors="coerce").fillna(0)
    df["excursion_date"] = pd.to_datetime(df["ExcursionDate"], errors="coerce")

    # Region: compute first (guide assignment depends on it)
    df["region"] = df["Region"].astype(str).str.strip()
    empty_mask = df["region"].str.lower().isin(["", "nan", "none"])
    df.loc[empty_mask, "region"] = df.loc[empty_mask, "hotel"].apply(_get_region)

    cfg = load_config()
    df["guide"] = df.apply(
        lambda r: _assign_guide_row(r["Payment method"], r["hotel"],
                                    r["Expedient email"], r["region"], cfg), axis=1
    )

    total       = len(df)
    grand_total = round(float(df["price"].sum()), 2)
    avg         = round(grand_total / total, 2) if total > 0 else 0.0

    # Guide breakdown
    guides: dict = {}
    for g in GUIDES:
        gdf   = df[df["guide"] == g]
        gtot  = round(float(gdf["price"].sum()), 2)
        gavg  = round(gtot / len(gdf), 2) if len(gdf) > 0 else 0.0
        guides[g] = {
            "count":       len(gdf),
            "total":       gtot,
            "avg":         gavg,
            "revenue_pct": round(gtot / grand_total * 100, 1) if grand_total > 0 else 0.0,
            "booking_pct": round(len(gdf) / total * 100, 1) if total > 0 else 0.0,
        }

    unassigned_df = df[df["guide"] == "unassigned"]
    guides["unassigned"] = {
        "count": len(unassigned_df),
        "total": round(float(unassigned_df["price"].sum()), 2),
    }

    # Region breakdown
    regions: dict = {}
    for region, rdf in df.groupby("region"):
        regions[region] = {"count": len(rdf), "total": round(float(rdf["price"].sum()), 2)}

    # Time series — auto-choose granularity based on period length
    valid_dates = df["excursion_date"].dropna()
    if len(valid_dates) > 0:
        span_days = (valid_dates.max() - valid_dates.min()).days
        if span_days <= 14:
            df["bucket"] = df["excursion_date"].dt.strftime("%Y-%m-%d")
            bucket_label = "day"
        elif span_days <= 90:
            df["bucket"] = df["excursion_date"].dt.to_period("W").apply(
                lambda p: p.start_time.strftime("%b %d") if pd.notna(p) else "?"
            )
            bucket_label = "week"
        else:
            df["bucket"] = df["excursion_date"].dt.to_period("M").apply(
                lambda p: p.start_time.strftime("%b %Y") if pd.notna(p) else "?"
            )
            bucket_label = "month"

        series: list = []
        for bucket, bdf in df.groupby("bucket"):
            entry = {"label": str(bucket), "total_count": len(bdf)}
            for g in GUIDES:
                entry[g] = int((bdf["guide"] == g).sum())
            series.append(entry)
        series.sort(key=lambda x: x["label"])
    else:
        series       = []
        bucket_label = "week"

    # Top 10 hotels by booking count
    hotel_stats = (
        df.groupby("hotel")
          .agg(guide=("guide","first"), region=("region","first"),
               count=("booking_id","count"), total=("price","sum"))
          .reset_index()
          .sort_values("count", ascending=False)
          .head(10)
    )
    top_hotels = [
        {"hotel": r["hotel"], "guide": r["guide"], "region": r["region"],
         "count": int(r["count"]), "total": round(float(r["total"]), 2)}
        for _, r in hotel_stats.iterrows()
    ]

    return {
        "date_from":        date_from or "",
        "date_to":          date_to   or "",
        "channel":          channel,
        "total":            total,
        "grand_total":      grand_total,
        "avg_per_booking":  avg,
        "cancelled":        cancelled,
        "guides":           guides,
        "regions":          regions,
        "series":           series,
        "bucket_label":     bucket_label,
        "top_hotels":       top_hotels,
    }


def main():
    if len(sys.argv) != 3:
        print("Usage: python guide_report.py <seeplaces_report.xlsx> <output.xlsx>")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])


if __name__ == "__main__":
    main()
