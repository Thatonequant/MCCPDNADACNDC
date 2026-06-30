#!/usr/bin/env python3
"""
NADAC vs. Cost Plus Drugs pipeline.

Downloads the current CMS NADAC dataset and the live Cost Plus Drugs catalog,
matches them by NDC, classifies drugs by estimated prior-authorization (PA)
burden, and produces:

  docs/nadac_vs_costplus.xlsx   - full multi-tab analysis workbook
  docs/index.html               - interactive search/scatter explorer

Run locally with:  python3 pipeline.py
Designed to also run unattended in GitHub Actions (see .github/workflows/update.yml).

IMPORTANT CAVEATS (carried over from the original analysis):
  - NADAC is acquisition cost, not what a cash-pay patient is billed.
  - The "Est. Retail (AWP)" columns use PUBLISHED INDUSTRY RATIOS
    (Brand approx NADAC x1.25, Generic approx NADAC x1.90), not measured retail prices.
  - PA Tier (Low/Moderate/High) is a CLINICAL HEURISTIC based on drug class,
    brand status, and cost -- NOT measured prior-authorization or denial data.
  - Cost Plus does not carry controlled substances or most cold-chain injectables;
    those will simply never appear in the matched set.
"""

import csv
import io
import json
import re
import sys
import statistics as st
from collections import defaultdict
from pathlib import Path

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

OUT_DIR = Path(__file__).parent / "docs"
OUT_DIR.mkdir(exist_ok=True)

CPD_API = "https://api.costplusdrugs.com/pricelist/cpd"

# Primary: try the well-known, stable direct-download URL pattern across recent dates.
# NADAC publishes weekly (every Wednesday) plus a monthly file on the first Monday on/after
# the 15th -- rather than guess which, we just try every day going back ~25 days until one hits.
# This avoids depending on any dataset ID, which changes every January (NADAC 2026 -> NADAC 2027 -> ...)
# and avoids depending on any API gateway, which has repeatedly proven less stable than the raw file URLs.
NADAC_CSV_PATTERN = "https://download.medicaid.gov/data/nadac-national-average-drug-acquisition-cost-{m:02d}-{d:02d}-{y}.csv"

# Fallback: scrape the (corrected) NADAC landing page for the current dated CSV link.
NADAC_LANDING_PAGE = "https://www.medicaid.gov/medicaid/nadac"


# --------------------------------------------------------------------------
# Step 1: Download NADAC
# --------------------------------------------------------------------------
def _parse_nadac_csv_text(rows_text):
    nadac = defaultdict(list)
    reader = csv.DictReader(io.StringIO(rows_text))
    for row in reader:
        ndc = (row.get("ndc") or row.get("NDC") or "").replace("-", "")
        price_raw = row.get("nadac_per_unit") or row.get("NADAC Per Unit")
        date_raw = row.get("effective_date") or row.get("Effective Date") or ""
        cls_raw = row.get("classification_for_rate_setting") or row.get("Classification for Rate Setting") or ""
        corresponding = (
            row.get("corresponding_generic_drug_nadac_per_unit")
            or row.get("Corresponding Generic Drug NADAC Per Unit")
            or ""
        )
        if not ndc or not price_raw:
            continue
        try:
            price = float(price_raw)
        except ValueError:
            continue
        nadac[ndc].append({
            "date": date_raw, "price": price, "classification": cls_raw,
            "corresponding_generic": corresponding,
        })
    return nadac


def download_nadac():
    """Returns a dict: normalized NDC -> list of row-dicts (date, price, classification, corresponding_generic)."""
    import datetime as _dt

    print("Downloading NADAC data...")
    headers = {"User-Agent": "Mozilla/5.0 (compatible; nadac-cpd-pipeline/1.0)"}

    # --- Primary: guess recent dates against the stable direct-download pattern ---
    today = _dt.date.today()
    for days_back in range(0, 28):
        d = today - _dt.timedelta(days=days_back)
        url = NADAC_CSV_PATTERN.format(m=d.month, d=d.day, y=d.year)
        try:
            resp = requests.get(url, headers=headers, timeout=120)
            if resp.status_code == 200 and len(resp.text) > 1_000_000:
                print(f"  Found NADAC file for {d.isoformat()}: {url}")
                nadac = _parse_nadac_csv_text(resp.text)
                print(f"  Parsed {len(nadac):,} unique NDCs from NADAC.")
                return nadac
        except requests.RequestException:
            continue
    print("  Date-guessing exhausted 28 days without a hit. Falling back to landing-page scrape...")

    # --- Fallback: scrape the landing page for whatever dated CSV link is currently posted ---
    resp = requests.get(NADAC_LANDING_PAGE, headers=headers, timeout=60)
    resp.raise_for_status()
    matches = re.findall(
        r'https://download\.medicaid\.gov/data/nadac-national-average-drug-acquisition-cost-[\d-]+\.csv',
        resp.text,
    )
    if not matches:
        raise RuntimeError(
            "Could not find a NADAC CSV link via date-guessing OR the landing page scrape. "
            f"CMS may have changed their file naming or page structure -- check {NADAC_LANDING_PAGE} "
            "manually and update NADAC_CSV_PATTERN / NADAC_LANDING_PAGE in pipeline.py."
        )
    csv_url = sorted(set(matches))[-1]
    print(f"  Found via landing page: {csv_url}")
    resp = requests.get(csv_url, headers=headers, timeout=120)
    resp.raise_for_status()
    nadac = _parse_nadac_csv_text(resp.text)
    print(f"  Parsed {len(nadac):,} unique NDCs from NADAC.")
    return nadac

HIGH_PA_CATEGORIES = {
    "Cancer", "Breast Cancer", "Leukemia", "HIV", "Organ Transplant",
    "Hidradenitis Suppurativa", "Psoriatic Arthritis", "Plaque Psoriasis",
    "Ulcerative Colitis", "Crohn's Disease", "Multiple sclerosis",
    "Pulmonary Fibrosis", "Huntington's Disease",
}
MODERATE_PA_CATEGORIES = {
    "Rheumatoid Arthritis", "Gout", "Fertility", "Erectile Dysfunction",
    "Weight Management", "Migraines", "Endometriosis", "Restless Leg Syndrome",
    "Overactive Bladder", "Iron Overload", "Wilson Disease", "ALS",
}
LOW_PA_CATEGORIES = {
    "Birth Control", "Infection", "Anti-bacterial", "High Blood Pressure",
    "High Cholesterol", "Diabetes", "Diuretic", "Steroid", "Pain & Inflammation",
    "Hormone Therapy", "Thyroid", "Vitamin Deficiency", "Low Blood Sugar",
}

AWP_MULTIPLIER = {"Branded": 1.25, "Generic": 1.90}
FONT_NAME = "Arial"


# --------------------------------------------------------------------------
# Step 2: Fetch Cost Plus Drugs catalog
# --------------------------------------------------------------------------
def fetch_cpd_catalog():
    print("Fetching Cost Plus Drugs catalog...")
    resp = requests.get(CPD_API, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    print(f"  Retrieved {len(data):,} drug entries.")

    drugs = []
    for entry in data:
        cpd = entry.get("cpd_channel") or {}
        if not cpd:
            continue
        ndcs = entry.get("equivalent_ndcs") or []
        drugs.append({
            "name": entry.get("medication_name"),
            "brand": entry.get("brand_name"),
            "type": entry.get("brand_generic"),
            "form": entry.get("form"),
            "strength": entry.get("strength"),
            "categories": "; ".join(entry.get("treatment_categories") or []),
            "pack_size": entry.get("sample_pack_size") or entry.get("pack_size") or 1,
            "price_per_unit": cpd.get("price_per_unit"),
            "fee": cpd.get("cash_dispensing_fee") if cpd.get("cash_dispensing_fee") is not None else 5,
            "shipping": cpd.get("cash_shipping_cost") if cpd.get("cash_shipping_cost") is not None else 5.25,
            "in_stock": "Yes" if cpd.get("in_stock") else "No",
            "ndcs": ndcs,
        })
    return drugs


# --------------------------------------------------------------------------
# Step 3: Match by NDC
# --------------------------------------------------------------------------
def norm(ndc):
    return (ndc or "").replace("-", "")


def match_drugs(cpd_drugs, nadac):
    print("Matching by NDC...")
    matched = []
    for d in cpd_drugs:
        if d["price_per_unit"] is None:
            continue
        chosen = None
        for ndc in d["ndcs"]:
            key = norm(ndc)
            if key in nadac:
                rows = sorted(nadac[key], key=lambda r: r["date"])
                latest = rows[-1]
                if latest["classification"] == "G":
                    chosen = latest["price"]
                    break
        if chosen is None:
            for ndc in d["ndcs"]:
                key = norm(ndc)
                if key in nadac:
                    rows = sorted(nadac[key], key=lambda r: r["date"])
                    latest = rows[-1]
                    if latest["classification"] == "B" and latest["corresponding_generic"]:
                        try:
                            chosen = float(latest["corresponding_generic"])
                            break
                        except ValueError:
                            pass
        if chosen is None:
            for ndc in d["ndcs"]:
                key = norm(ndc)
                if key in nadac:
                    rows = sorted(nadac[key], key=lambda r: r["date"])
                    chosen = rows[-1]["price"]
                    break
        if chosen is not None:
            matched.append({**d, "nadac": chosen})
    print(f"  Matched {len(matched):,} of {len(cpd_drugs):,} drugs ({len(matched)/len(cpd_drugs)*100:.1f}%).")
    return matched


# --------------------------------------------------------------------------
# Step 4: PA tier classification (heuristic, not measured data)
# --------------------------------------------------------------------------
def classify_pa(d):
    score = 0
    cats = set((d["categories"] or "").split("; "))
    if d["type"] == "Branded":
        score += 2
    if cats & HIGH_PA_CATEGORIES:
        score += 3
    if cats & MODERATE_PA_CATEGORIES:
        score += 1
    pack = d["pack_size"] or 1
    pack_cost = d["price_per_unit"] * pack
    if pack_cost > 200:
        score += 2
    elif pack_cost > 50:
        score += 1
    if cats & {"HIV"}:
        score += 1
    if (cats & LOW_PA_CATEGORIES) and d["type"] == "Generic" and pack_cost < 30:
        score -= 1
    if score <= 0:
        return "Low"
    elif score <= 2:
        return "Moderate"
    return "High"


def compute_metrics(d):
    pack = d["pack_size"] or 1
    nadac_cost = d["nadac"] * pack
    cpd_bare = d["price_per_unit"] * pack
    cpd_total = cpd_bare + d["fee"] + d["shipping"]
    mult = AWP_MULTIPLIER.get(d["type"], 1.90)
    est_retail = nadac_cost * mult
    sav_nadac = (nadac_cost - cpd_total) / nadac_cost * 100 if nadac_cost else 0
    sav_retail = (est_retail - cpd_total) / est_retail * 100 if est_retail else 0
    return {
        "nadac_cost": nadac_cost, "cpd_bare": cpd_bare, "cpd_total": cpd_total,
        "est_retail": est_retail, "sav_nadac": sav_nadac, "sav_retail": sav_retail,
    }


# --------------------------------------------------------------------------
# Step 5: Build the Excel workbook
# --------------------------------------------------------------------------
def build_workbook(matched, out_path):
    print("Building Excel workbook...")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", start_color="1B3A3A")
    body_font = Font(name=FONT_NAME, size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for d in matched:
        d["tier"] = classify_pa(d)
        d["metrics"] = compute_metrics(d)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Summary ---
    summary = wb.create_sheet("Summary")
    summary.cell(row=1, column=1, value="NADAC vs. Cost Plus Drugs — Summary").font = Font(name=FONT_NAME, size=14, bold=True)
    n = len(matched)
    sav_nadac_all = [d["metrics"]["sav_nadac"] for d in matched]
    sav_retail_all = [d["metrics"]["sav_retail"] for d in matched]
    pa_drugs = [d for d in matched if d["tier"] in ("Moderate", "High")]
    stats = [
        ("Matched drugs", n),
        ("Median savings vs. NADAC, incl. fees", f"{st.median(sav_nadac_all):.1f}%"),
        ("Median savings vs. Est. Retail (AWP), incl. fees", f"{st.median(sav_retail_all):.1f}%"),
        ("Moderate/High PA-burden drugs", len(pa_drugs)),
        ("Data refreshed", __import__("datetime").datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")),
    ]
    r = 3
    for label, val in stats:
        summary.cell(row=r, column=1, value=label).font = Font(name=FONT_NAME, bold=True, size=10)
        summary.cell(row=r, column=2, value=val).font = Font(name=FONT_NAME, size=10)
        r += 1
    r += 1
    caveats = [
        "NADAC is acquisition cost, not what a cash-pay patient is billed at the counter.",
        "Est. Retail (AWP) uses published industry ratios (Brand x1.25, Generic x1.90), not measured retail prices.",
        "PA Tier is a clinical heuristic (drug class, brand status, cost) -- NOT measured PA/denial data.",
        "Cost Plus does not carry controlled substances or most cold-chain injectables.",
        "This is a methodology demonstration, not medical or financial advice.",
    ]
    for cv in caveats:
        summary.cell(row=r, column=1, value="• " + cv).font = Font(name=FONT_NAME, size=10)
        summary.cell(row=r, column=1).alignment = Alignment(wrap_text=True)
        summary.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        summary.row_dimensions[r].height = 26
        r += 1
    summary.column_dimensions["A"].width = 42
    summary.column_dimensions["B"].width = 22

    # --- Matched (full data) ---
    ws = wb.create_sheet("Matched (Incl. Fees)")
    headers = ["Medication Name", "Brand Name", "Type", "Form", "Strength", "Treatment Categories",
               "Pack Size", "NADAC/Unit", "Cost Plus/Unit", "Fee", "Shipping", "PA Tier",
               "NADAC Pack Cost", "Cost Plus Total", "Savings vs NADAC (%)", "Savings vs Retail/AWP (%)", "In Stock"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font; c.fill = header_fill; c.border = border
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    matched_sorted = sorted(matched, key=lambda d: -d["metrics"]["sav_retail"])
    for i, d in enumerate(matched_sorted, start=2):
        m = d["metrics"]
        vals = [d["name"], d["brand"], d["type"], d["form"], d["strength"], d["categories"],
                d["pack_size"], d["nadac"], d["price_per_unit"], d["fee"], d["shipping"], d["tier"]]
        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=col, value=v)
            cell.font = body_font; cell.border = border
        for col, val, fmt in [(13, m["nadac_cost"], "$#,##0.00"), (14, m["cpd_total"], "$#,##0.00"),
                               (15, m["sav_nadac"]/100, "0.0%"), (16, m["sav_retail"]/100, "0.0%")]:
            cell = ws.cell(row=i, column=col, value=val); cell.font = body_font; cell.border = border; cell.number_format = fmt
        ws.cell(row=i, column=17, value=d["in_stock"]).font = body_font
    widths = [30, 20, 10, 20, 14, 28, 10, 12, 12, 8, 8, 12, 14, 14, 14, 16, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:Q{len(matched_sorted)+1}"
    rule = ColorScaleRule(start_type="min", start_color="F6E0DA", mid_type="num", mid_value=0,
                           mid_color="FFFFFF", end_type="max", end_color="C8E0CC")
    ws.conditional_formatting.add(f"P2:P{len(matched_sorted)+1}", rule)

    # --- Category Analysis ---
    ws2 = wb.create_sheet("Category Analysis")
    cat_data = defaultdict(list)
    for d in matched:
        for cat in (d["categories"] or "").split("; "):
            if cat:
                cat_data[cat].append(d)
    rows2 = []
    for cat, drugs in cat_data.items():
        if len(drugs) < 5:
            continue
        sav = [x["metrics"]["sav_nadac"] for x in drugs]
        rows2.append((cat, len(drugs), st.median(sav)))
    rows2.sort(key=lambda x: -x[2])
    headers2 = ["Category", "# Drugs", "Median Savings vs NADAC (%)"]
    for col, h in enumerate(headers2, 1):
        c = ws2.cell(row=1, column=col, value=h); c.font = header_font; c.fill = header_fill; c.border = border
    for i, (cat, cnt, med) in enumerate(rows2, start=2):
        ws2.cell(row=i, column=1, value=cat).font = body_font
        ws2.cell(row=i, column=2, value=cnt).font = body_font
        c = ws2.cell(row=i, column=3, value=med/100); c.font = body_font; c.number_format = "0.0%"
    ws2.column_dimensions["A"].width = 30
    ws2.column_dimensions["B"].width = 12
    ws2.column_dimensions["C"].width = 24
    ws2.freeze_panes = "A2"

    # --- PA Burden (Moderate/High) with AWP comparison ---
    ws3 = wb.create_sheet("PA Burden - Est. Retail (AWP)")
    note = ("Moderate/High PA-burden drugs (heuristic classification). Est. Retail uses published "
            "NADAC-to-AWP ratios (Brand x1.25, Generic x1.90) -- an industry rule of thumb, not measured pricing.")
    ws3.cell(row=1, column=1, value=note).font = Font(name=FONT_NAME, size=9, italic=True, color="7C6A4F")
    ws3.merge_cells("A1:H1")
    ws3.row_dimensions[1].height = 30
    headers3 = ["Medication Name", "Strength", "PA Tier", "NADAC Pack Cost", "Est. Retail (AWP)",
                "Cost Plus Total", "Savings vs NADAC (%)", "Savings vs Retail (%)"]
    for col, h in enumerate(headers3, 1):
        c = ws3.cell(row=3, column=col, value=h); c.font = header_font; c.fill = header_fill; c.border = border
    pa_sorted = sorted(pa_drugs, key=lambda d: -d["metrics"]["sav_retail"])
    for i, d in enumerate(pa_sorted, start=4):
        m = d["metrics"]
        vals = [d["name"], d["strength"], d["tier"]]
        for col, v in enumerate(vals, 1):
            ws3.cell(row=i, column=col, value=v).font = body_font
        for col, val, fmt in [(4, m["nadac_cost"], "$#,##0.00"), (5, m["est_retail"], "$#,##0.00"),
                               (6, m["cpd_total"], "$#,##0.00"), (7, m["sav_nadac"]/100, "0.0%"),
                               (8, m["sav_retail"]/100, "0.0%")]:
            cell = ws3.cell(row=i, column=col, value=val); cell.font = body_font; cell.number_format = fmt
    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 14
    for col_letter in ["C", "D", "E", "F", "G", "H"]:
        ws3.column_dimensions[col_letter].width = 16
    ws3.freeze_panes = "A4"

    wb.save(out_path)
    print(f"  Saved {out_path}")


# --------------------------------------------------------------------------
# Step 6: Build the interactive HTML explorer
# --------------------------------------------------------------------------
def build_html_explorer(matched, out_path):
    print("Building HTML explorer...")
    pa_data = []
    for d in matched:
        tier = classify_pa(d)
        if tier not in ("Moderate", "High"):
            continue
        m = compute_metrics(d)
        if m["cpd_bare"] <= 0 or m["nadac_cost"] <= 0:
            continue
        pa_data.append({
            "name": d["name"], "brand": d["brand"], "strength": d["strength"], "type": d["type"],
            "tier": tier, "categories": d["categories"],
            "cpd_cost": round(m["cpd_bare"], 2), "cpd_total": round(m["cpd_total"], 2),
            "nadac_cost": round(m["nadac_cost"], 2), "est_retail": round(m["est_retail"], 2),
            "sav_nadac": round(m["sav_nadac"], 1), "sav_retail": round(m["sav_retail"], 1),
        })

    data_json = json.dumps(pa_data)
    template_path = Path(__file__).parent / "explorer_template.html"
    html = template_path.read_text()
    html = html.replace("__DATA_JSON__", data_json)
    out_path.write_text(html)
    print(f"  Saved {out_path} ({len(pa_data)} drugs)")


# --------------------------------------------------------------------------
def main():
    nadac = download_nadac()
    cpd_drugs = fetch_cpd_catalog()
    matched = match_drugs(cpd_drugs, nadac)
    build_workbook(matched, OUT_DIR / "nadac_vs_costplus.xlsx")
    build_html_explorer(matched, OUT_DIR / "index.html")
    print("Done.")


if __name__ == "__main__":
    sys.exit(main())
