"""Extract the customer-visit deck figures from the SAP AR invoice export.

Everything the interactive deck draws comes out of this module, so no number in the
deck is ever retyped by hand. The output is a JSON payload where every chart datum
carries the invoice lines behind it -- that per-item detail is what the hover
tooltips read.

Framing note: the deck talks in UNITS (transfusion) and TESTS (immunoassay), never
ringgit, so every aggregate here is on Quantity, not Sales Amount.
"""

import argparse
import collections
import datetime
import json
import re
import warnings

import openpyxl

warnings.filterwarnings("ignore")

# --- period definitions ----------------------------------------------------------
# Calendar years. The export opens on 5 May 2025 and closes on 4 Sep 2026, so 2025 is
# only its last eight months and 2026 only its first eight -- equal-length windows,
# but different halves of the year, which is why the like-for-like pair below exists.
Y25 = [f"2025-{m:02d}" for m in range(5, 13)]
Y25_LABEL = "2025 · May–Dec"
Y26 = [f"2026-{m:02d}" for m in range(1, 9)]
Y26_LABEL = "2026 YTD · Jan–Aug"
# The one window present in both years: the part-month of Sep 2026 is excluded.
LFL_25 = ["2025-05", "2025-06", "2025-07", "2025-08"]
LFL_26 = ["2026-05", "2026-06", "2026-07", "2026-08"]
LAST_SIX = ["2026-03", "2026-04", "2026-05", "2026-06", "2026-07", "2026-08"]

TRANSFUSION = ("Bio-Rad", "Werfen Immucor")

# SAP product groups -> the names a customer would recognise.
GROUP_LABELS = {
    "DM-ID CARDS": "Bio-Rad ID gel cards",
    "DM-ID CELLS": "Bio-Rad ID cells",
    "DM-ID REAGENT": "Bio-Rad ID reagent",
    "DM-ID QC": "Bio-Rad ID QC",
    "DM-ACCS": "Bio-Rad accessories",
    "IMCG-REAGENT": "Werfen Immucor reagent",
    "IMCG-ACCS": "Werfen Immucor accessories",
    "DBL-REAGENT": "Werfen DBL reagent",
}

# Sarawak transfusion sites, matched on the SAP ship-to name.
SARAWAK_SITES = [
    ("Hospital Umum Sarawak", ("HOSPITAL UMUM SARAWAK",)),
    ("Hospital Sibu", ("HOSPITAL SIBU",)),
    ("Hospital Miri", ("HOSPITAL MIRI",)),
    ("Hospital Bintulu", ("HOSPITAL BINTULU",)),
    ("Hospital Sarikei", ("HOSPITAL SARIKEI",)),
    ("Hospital Kapit", ("HOSPITAL KAPIT",)),
    ("Hospital Limbang", ("HOSPITAL LIMBANG",)),
    ("Hospital Sri Aman", ("HOSPITAL SRI AMAN",)),
    ("Hospital Betong", ("HOSPITAL BETONG",)),
    ("Hospital Mukah", ("HOSPITAL MUKAH",)),
    ("Hospital Lawas", ("HOSPITAL LAWAS",)),
    ("Hospital Serian", ("HOSPITAL SERIAN",)),
]

# I-LOAN PDN tender achievement, reproduced from the supplied contract report.
TENDER = [
    ("PDN", 53.13), ("Melaka", 46.77), ("Perlis", 45.59), ("Sarawak", 41.52),
    ("Wil. Persekutuan", 40.26), ("P. Pinang", 33.10), ("Selangor", 31.79),
    ("N. Sembilan", 31.43), ("Sabah & Labuan", 27.96), ("Johor", 17.65),
    ("Terengganu", 17.08), ("Kelantan", 16.00), ("Kedah", 13.77),
    ("Perak", 12.19), ("Pahang", 11.81),
]
TENDER_TOTAL = 28.59

MONTH_LABEL = {
    "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
    "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
}


def mlabel(m):
    y, mm = m.split("-")
    return f"{MONTH_LABEL[mm]} {y[2:]}"


# --- load ------------------------------------------------------------------------
def load(src):
    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    it = wb["Export"].iter_rows(values_only=True)
    idx = {h: n for n, h in enumerate(next(it))}
    out = []
    ncols = max(idx.values()) + 1
    for r in it:
        if len(r) < ncols:
            continue
        d = r[idx["Date"]]
        if not isinstance(d, datetime.datetime):
            continue
        try:
            qty = float(r[idx["Quantity"]] or 0)
        except (TypeError, ValueError):
            qty = 0.0
        try:
            amt = float(r[idx["Sales Amount"]] or 0)
        except (TypeError, ValueError):
            amt = 0.0
        out.append({
            "month": d.strftime("%Y-%m"),
            "year": d.year,
            "brand": r[idx["Brand"]],
            "group": r[idx["Product Group"]] or "",
            "product": r[idx["Product Name"]] or "",
            "ship": (r[idx["Ship To Name (SAP)"]] or "").upper(),
            "qty": qty,
            "amt": amt,
        })
    wb.close()
    return out


# Title-casing a shouty SAP name mangles the parts that are already correct
# ("2ND GEN" -> "2Nd Gen"); these put them back.
CASE_FIXES = [
    (r"\b(\d+)(St|Nd|Rd|Th)\b", lambda m: m.group(1) + m.group(2).lower()),
    (r"\b(Id|Abo|Qc|Clia|Ebv|Ana|Iga|Igg|Igm|Ivd|Pdn|Diva|Bd)\b",
     lambda m: m.group(1).upper()),
    (r"\bI-Ii-Iii\b", "I-II-III"),
    (r"\bLiss\b", "LISS"),
    (r"\bIdqc(\d)\b", lambda m: "IDQC" + m.group(1)),
    (r"\b(I{2,3})\b", lambda m: m.group(1).upper()),
    (r"\bIi\b", "II"),
    (r"(\d)X(\d)", lambda m: m.group(1) + "x" + m.group(2)),
    (r"(\d)Ml\b", lambda m: m.group(1) + "ml"),
    (r"\bMl\b", "ml"),
]


def tidy(name):
    """SAP product names are shouty and carry pack sizes; make them readable."""
    n = re.sub(r"\s+", " ", str(name)).strip(" ;,")
    n = re.sub(r"^(CONSUMABLES|INFECTIOUS DISEASES)\s*[-;]\s*", "", n, flags=re.I)
    if n.isupper():
        n = n.title()
        for pat, rep in CASE_FIXES:
            n = re.sub(pat, rep, n)
    return n


def breakdown(rows, key=lambda r: r["product"], limit=8):
    """Collapse invoice lines into the {name, units} list a tooltip shows."""
    agg = collections.defaultdict(float)
    for r in rows:
        agg[tidy(key(r))] += r["qty"]
    items = sorted(((k, v) for k, v in agg.items() if v), key=lambda x: -x[1])
    head = [{"name": k, "units": v} for k, v in items[:limit]]
    if len(items) > limit:
        rest = sum(v for _, v in items[limit:])
        head.append({"name": f"{len(items) - limit} other lines", "units": rest})
    return head


def point(label, rows, sub=None):
    """One chart datum: its value, plus the invoice detail behind it."""
    return {
        "label": label,
        "value": sum(r["qty"] for r in rows),
        "sub": sub,
        "items": breakdown(rows),
    }


# --- Hospital Umum Sarawak -------------------------------------------------------
def build_hus(rows):
    tx = [r for r in rows
          if "HOSPITAL UMUM SARAWAK" in r["ship"] and r["brand"] in TRANSFUSION]

    # units by product line, 2025 (May-Dec) vs 2026 YTD (Jan-Aug)
    groups = collections.defaultdict(lambda: {"y25": [], "y26": []})
    for r in tx:
        if r["month"] in Y25:
            groups[r["group"]]["y25"].append(r)
        elif r["month"] in Y26:
            groups[r["group"]]["y26"].append(r)
    by_line = []
    for g, v in groups.items():
        q25 = sum(r["qty"] for r in v["y25"])
        q26 = sum(r["qty"] for r in v["y26"])
        if not (q25 or q26):
            continue
        by_line.append({
            "label": GROUP_LABELS.get(g, str(g).title()),
            "y25": q25,
            "y26": q26,
            # Both windows are eight months long, so this is a like-for-like ratio.
            "change": (q26 / q25 - 1) * 100 if q25 else None,
            "items25": breakdown(v["y25"]),
            "items26": breakdown(v["y26"]),
        })
    by_line.sort(key=lambda x: -(x["y25"] + x["y26"]))

    # Sarawak sites, like-for-like May-Aug in each year
    sites = []
    for name, keys in SARAWAK_SITES:
        site = [r for r in rows
                if r["brand"] in TRANSFUSION and any(k in r["ship"] for k in keys)]
        a = [r for r in site if r["month"] in LFL_25]
        b = [r for r in site if r["month"] in LFL_26]
        qa, qb = sum(r["qty"] for r in a), sum(r["qty"] for r in b)
        if not (qa or qb):
            continue
        sites.append({
            "label": name, "y25": qa, "y26": qb,
            "change": (qb / qa - 1) * 100 if qa else None,
            "items25": breakdown(a), "items26": breakdown(b),
        })
    sites.sort(key=lambda x: -(x["y25"] + x["y26"]))

    monthly = [point(mlabel(m), [r for r in tx if r["month"] == m]) for m in LAST_SIX]

    cards = [r for r in tx if r["group"] == "DM-ID CARDS"]
    bench = [r for r in tx if r["group"] in ("DM-ID CELLS", "DM-ID REAGENT", "DM-ID QC")]

    def q(rs, months):
        return sum(r["qty"] for r in rs if r["month"] in months)

    return {
        "y25_total": q(tx, Y25),
        "y26_total": q(tx, Y26),
        "lfl25": q(tx, LFL_25),
        "lfl26": q(tx, LFL_26),
        "by_line": by_line,
        "sites": sites,
        "monthly": monthly,
        "tender": [{"label": n, "value": v} for n, v in TENDER],
        "tender_total": TENDER_TOTAL,
        "cards25": q(cards, Y25),
        "cards26": q(cards, Y26),
        "bench26": q(bench, Y26),
        "last_card_month": max((r["month"] for r in cards if r["qty"]), default=None),
        "cards_monthly": [
            point(mlabel(m), [r for r in cards if r["month"] == m]) for m in LAST_SIX],
    }


# --- PIL Timberland (SNIBE Maglumi) ----------------------------------------------
TESTS_PER_KIT = re.compile(r"(\d+)\s*(?:T\b|TEST)", re.I)

# The three CLIA assays they run, under the names the lab uses for them.
ASSAY_NAMES = [
    ("H.PYLORI", "H. pylori IgG"),
    ("EBV", "EBV VCA IgA"),
    ("ANA SCREEN", "ANA Screen"),
]


def assay_name(product):
    up = str(product).upper()
    for key, label in ASSAY_NAMES:
        if key in up:
            return label
    return tidy(product)


def kit_size(name):
    """Tests per kit, read off the SAP product name (e.g. '...; 100 TEST')."""
    m = TESTS_PER_KIT.search(str(name))
    return int(m.group(1)) if m else 0


def build_pil(rows):
    sn = [r for r in rows if "TIMBERLAND" in r["ship"] and r["brand"] == "SNIBE"]
    for r in sn:
        r["tests"] = r["qty"] * kit_size(r["product"])

    assays = [r for r in sn if r["tests"]]          # billed reagent kits
    consumables = [r for r in sn if not r["tests"]]  # RM 0 rental consumables

    def tests(rs):
        return sum(r["tests"] for r in rs)

    def window(rs, months):
        return [r for r in rs if r["month"] in months]

    def tpoint(label, rs):
        agg = collections.defaultdict(float)
        for r in rs:
            agg[assay_name(r["product"])] += r["tests"]
        items = sorted(((k, v) for k, v in agg.items() if v), key=lambda x: -x[1])
        return {"label": label, "value": tests(rs), "sub": None,
                "items": [{"name": k, "units": v} for k, v in items]}

    mix = collections.defaultdict(lambda: {"y25": [], "y26": []})
    for r in assays:
        if r["month"] in Y25:
            mix[assay_name(r["product"])]["y25"].append(r)
        elif r["month"] in Y26:
            mix[assay_name(r["product"])]["y26"].append(r)
    by_assay = []
    for name, v in mix.items():
        size = kit_size(v["y25"][0]["product"] if v["y25"] else v["y26"][0]["product"])
        by_assay.append({
            "label": name,
            "y25": tests(v["y25"]), "y26": tests(v["y26"]),
            "kits25": sum(r["qty"] for r in v["y25"]),
            "kits26": sum(r["qty"] for r in v["y26"]),
            "size": size,
            "items25": [{"name": f"{sum(r['qty'] for r in v['y25']):.0f} kits "
                                 f"\u00d7 {size} tests", "units": tests(v["y25"])}],
            "items26": [{"name": f"{sum(r['qty'] for r in v['y26']):.0f} kits "
                                 f"\u00d7 {size} tests", "units": tests(v["y26"])}],
        })
    by_assay.sort(key=lambda x: -(x["y25"] + x["y26"]))

    monthly = [tpoint(mlabel(m), [r for r in assays if r["month"] == m])
               for m in LAST_SIX]

    return {
        "y25_tests": tests(window(assays, Y25)),
        "y26_tests": tests(window(assays, Y26)),
        "lfl25_tests": tests(window(assays, LFL_25)),
        "lfl26_tests": tests(window(assays, LFL_26)),
        "y25_kits": sum(r["qty"] for r in window(assays, Y25)),
        "y26_kits": sum(r["qty"] for r in window(assays, Y26)),
        "by_assay": by_assay,
        "monthly": monthly,
        "free_lines": len(consumables),
        "consumables": breakdown(consumables, limit=12),
    }


def build(src):
    rows = load(src)
    return {
        "generated": datetime.date.today().isoformat(),
        "hus": build_hus(rows),
        "pil": build_pil(rows),
    }


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("-o", "--out", default="deck_data.json")
    a = ap.parse_args()
    with open(a.out, "w") as fh:
        json.dump(build(a.src), fh, indent=1)
    print(f"wrote {a.out}")
