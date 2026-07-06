"""
setup_data.py
=============
Extracts real AGCM project data from Excel estimate workbooks and outputs:
  - estimation_agent/data/projects.json   → project-level records (from DATA sheet)
  - estimation_agent/data/divisions.json  → CSI division $/SF per project (from Estimate Summary sheet)

Usage:
  python setup_data.py

Place all Excel workbooks in the same folder as this script (or update EXCEL_DIR below).
Run once whenever new project files are added.
"""

import os
import json
import re
import openpyxl
from datetime import datetime

# ─── CONFIG ────────────────────────────────────────────────────────────────────
# Folder containing all the Excel estimate workbooks
EXCEL_DIR = "."

# Output folder (will be created if it doesn't exist)
OUTPUT_DIR = os.path.join("estimation_agent", "data")

# CSI division codes we care about (rows in Estimate Summary)
DIVISION_CODES = {
    "01 00 00": "General Requirements",
    "02 00 00": "Existing Conditions",
    "03 00 00": "Concrete",
    "04 00 00": "Masonry",
    "05 00 00": "Metals",
    "06 00 00": "Wood, Plastics, and Composites",
    "07 00 00": "Thermal and Moisture Protection",
    "08 00 00": "Openings",
    "09 00 00": "Finishes",
    "10 00 00": "Specialties",
    "11 00 00": "Equipment",
    "12 00 00": "Furnishings",
    "13 00 00": "Special Construction",
    "14 00 00": "Conveying Equipment",
    "21 00 00": "Fire Suppression",
    "22 00 00": "Plumbing",
    "23 00 00": "HVAC",
    "25 00 00": "Integrated Automation",
    "26 00 00": "Electrical",
    "27 00 00": "Communications",
    "28 00 00": "Electronic Safety and Security",
    "31 00 00": "Earthwork",
    "32 00 00": "Exterior Improvements",
    "33 00 00": "Utilities",
}

# Normalize division code key (strip trailing spaces)
def normalize_code(raw):
    if raw is None:
        return None
    return re.sub(r'\s+', ' ', str(raw).strip())


# ─── EXTRACT DATA SHEET ────────────────────────────────────────────────────────
def extract_data_sheet(ws):
    """
    Reads the DATA sheet and returns a project dict.
    The actual data row is always the row after the header row
    (which starts with 'Closing Date').
    """
    header = None
    for row in ws.iter_rows(values_only=True):
        if row[0] is not None and str(row[0]).strip() == "Closing Date":
            header = [str(c).strip() if c else "" for c in row]
            continue
        if header and any(c is not None for c in row):
            # This is the data row
            record = dict(zip(header, row))
            return record
    return None


# ─── EXTRACT ESTIMATE SUMMARY SHEET ───────────────────────────────────────────
def extract_estimate_summary(ws):
    """
    Reads the Estimate Summary sheet and returns:
      - project metadata (estimate number, name, area, province, city, etc.)
      - divisions: dict of {code: {total, cost_per_sf, me, labour, subcontract}}
      - totals: construction cost, OH&P, estimate price
    """
    meta = {}
    divisions = {}
    totals = {}

    rows = list(ws.iter_rows(values_only=True))

    for row in rows:
        # ── Header metadata rows ──────────────────────────────────────────
        if row[0] == "Closing Date :":
            val = row[1]
            if isinstance(val, datetime):
                meta["closing_date"] = val.strftime("%Y-%m-%d")
            else:
                meta["closing_date"] = str(val) if val else None

        elif row[0] == "Estimate # E-":
            meta["estimate_number"] = row[1]
            meta["area_sf"] = row[3]

        elif row[0] == "Province :":
            meta["province"] = row[1]
            meta["estimate_name"] = row[3]

        elif row[0] == "City/Town :":
            meta["city"] = row[1]
            meta["client"] = row[3]

        elif row[0] == "Budget Year :":
            meta["budget_year"] = row[1]
            meta["contract_type"] = row[3]

        elif row[0] == "Quarter:":
            meta["quarter"] = row[1]
            meta["payment_method"] = row[3]

        # ── Division data rows ────────────────────────────────────────────
        # Col 0 = CODE, Col 1 = Description, Col 2 = %, Col 3 = M&E,
        # Col 4 = M-HRS, Col 5 = LAB, Col 6 = SUB CONT, Col 7 = Total, Col 8 = $/SF string
        code_raw = row[0]
        if code_raw is not None:
            code = normalize_code(str(code_raw))
            # Match division codes like "01 00 00", "06 00 00", etc.
            if re.match(r'^\d{2} 00 00$', code):
                total = row[7] if row[7] is not None else 0
                me    = row[3] if row[3] is not None else 0
                labour = row[5] if row[5] is not None else 0
                subcon = row[6] if row[6] is not None else 0

                # Parse $/SF from string like "39.3$/SF" or numeric
                cost_per_sf = 0
                sf_raw = row[8]
                if sf_raw is not None:
                    sf_str = str(sf_raw).replace("$/SF", "").strip()
                    try:
                        cost_per_sf = float(sf_str)
                    except ValueError:
                        cost_per_sf = 0

                divisions[code] = {
                    "description": DIVISION_CODES.get(code, str(row[1]).strip() if row[1] else ""),
                    "total": round(float(total), 2),
                    "cost_per_sf": round(float(cost_per_sf), 2),
                    "me": round(float(me), 2),
                    "labour": round(float(labour), 2),
                    "subcontract": round(float(subcon), 2),
                }

        # ── Totals row ────────────────────────────────────────────────────
        if row[1] is not None and "Sub Totals" in str(row[1]):
            totals["construction_cost"] = round(float(row[7] or 0), 2)
            # Parse $/SF from col 8
            sf_raw = row[8]
            if sf_raw:
                sf_str = str(sf_raw).replace("$/SF", "").strip()
                try:
                    totals["cost_per_sf_total"] = float(sf_str)
                except ValueError:
                    totals["cost_per_sf_total"] = 0

        if row[1] is not None and "OH&P" in str(row[1]):
            totals["ohp"] = round(float(row[3] or 0), 2)
            totals["ohp_pct"] = round(float(row[4] or 0), 4)

    return meta, divisions, totals


# ─── INFER PROJECT TYPE ────────────────────────────────────────────────────────
def infer_project_type(name, type_code):
    """
    type_code in DATA sheet is stored as 0 (unknown/blank).
    Infer from project name keywords.
    """
    if type_code and type_code != 0:
        return str(type_code)

    name_lower = name.lower() if name else ""
    if any(k in name_lower for k in ["multi-res", "apartment", "residential", "condo", "ymca", "daycare"]):
        return "Multi-Residential / Institutional"
    if any(k in name_lower for k in ["dental", "medical", "clinic", "wellness", "health"]):
        return "Medical / Dental"
    if any(k in name_lower for k in ["office", "stantec", "title", "bioscript"]):
        return "Office"
    if any(k in name_lower for k in ["sobeys", "foodland", "grocery", "supermarket"]):
        return "Grocery / Food Retail"
    if any(k in name_lower for k in ["ford", "honda", "auto", "car", "dealership"]):
        return "Automotive"
    if any(k in name_lower for k in ["reno", "renovation", "expansion", "fit-up", "fit up", "tenant"]):
        return "Renovation / Tenant Fit-up"
    if any(k in name_lower for k in ["retail", "store", "shop", "mall"]):
        return "Retail"
    return "Commercial / Other"


# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_projects = []
    all_divisions = []

    xlsx_files = [
        f for f in os.listdir(EXCEL_DIR)
        if f.endswith(".xlsx") and not f.startswith("~")
    ]

    print(f"Found {len(xlsx_files)} Excel files in '{EXCEL_DIR}'\n")

    for fname in sorted(xlsx_files):
        fpath = os.path.join(EXCEL_DIR, fname)
        print(f"Processing: {fname}")

        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        except Exception as e:
            print(f"  ✗ Could not open: {e}\n")
            continue

        sheet_names = wb.sheetnames

        # ── DATA sheet ──────────────────────────────────────────────────
        if "DATA" not in sheet_names:
            print(f"  ✗ No DATA sheet found, skipping.\n")
            continue

        raw = extract_data_sheet(wb["DATA"])
        if not raw:
            print(f"  ✗ DATA sheet empty or unrecognized format.\n")
            continue

        estimate_name = raw.get("Estimate Name") or raw.get("Estimate Name ") or ""
        estimate_num  = raw.get("Estimate Number") or raw.get("Estimate Number2") or ""
        area_sf       = raw.get("Area") or 0
        price         = raw.get("Estimate Price") or 0
        margin        = raw.get("Margin") or 0
        schedule      = raw.get("Schedule") or 0
        cost_per_week = raw.get("Cost per Week") or 0
        soft_cost     = raw.get("Soft Cost") or 0
        province      = raw.get("Province") or ""
        city          = raw.get("City/Town") or ""
        quarter       = raw.get("Quarter") or ""
        budget_year   = raw.get("Budget Year") or ""
        contract_type = raw.get("Contract Type") or ""
        payment_method= raw.get("Payment Method") or ""
        price_cat     = raw.get("Price Category") or ""
        result        = raw.get("Successful / Unsuccessful") or ""
        status        = raw.get("Status") or ""
        type_code     = raw.get("Type of Construction") or 0

        project_type = infer_project_type(estimate_name, type_code)
        cost_per_sf  = round(float(price) / float(area_sf), 2) if area_sf and float(area_sf) > 0 else 0

        project_record = {
            "source_file":     fname,
            "estimate_number": str(estimate_num),
            "estimate_name":   str(estimate_name),
            "project_type":    project_type,
            "area_sf":         float(area_sf),
            "province":        str(province),
            "city":            str(city),
            "quarter":         str(quarter),
            "budget_year":     int(budget_year) if budget_year else None,
            "contract_type":   str(contract_type),
            "payment_method":  str(payment_method),
            "estimate_price":  round(float(price), 2),
            "price_category":  str(price_cat),
            "cost_per_sf":     cost_per_sf,
            "schedule_weeks":  int(schedule) if schedule else 0,
            "cost_per_week":   round(float(cost_per_week), 2),
            "soft_cost":       round(float(soft_cost), 2),
            "margin":          round(float(margin), 4),
            "margin_pct":      round(float(margin) * 100, 2),
            "status":          str(status),
            "result":          str(result),
        }

        all_projects.append(project_record)
        print(f"  ✓ Project: {estimate_name} | ${price:,.0f} | {area_sf:,.0f} SF | margin {float(margin)*100:.1f}%")

        # ── Estimate Summary sheet ───────────────────────────────────────
        if "Estimate Summary" not in sheet_names:
            print(f"  ✗ No Estimate Summary sheet.\n")
            continue

        meta, divisions, totals = extract_estimate_summary(wb["Estimate Summary"])

        # One division record per CSI division per project
        for code, div in divisions.items():
            if div["total"] == 0 and div["cost_per_sf"] == 0:
                continue  # skip empty divisions

            division_record = {
                "estimate_number": str(estimate_num),
                "estimate_name":   str(estimate_name),
                "project_type":    project_type,
                "area_sf":         float(area_sf),
                "province":        str(province),
                "city":            str(city),
                "budget_year":     int(budget_year) if budget_year else None,
                "division_code":   code,
                "division_name":   div["description"],
                "total":           div["total"],
                "cost_per_sf":     div["cost_per_sf"],
                "me":              div["me"],
                "labour":          div["labour"],
                "subcontract":     div["subcontract"],
                "pct_of_project":  round(div["total"] / float(price) * 100, 2) if price else 0,
            }
            all_divisions.append(division_record)

        active_divs = sum(1 for d in divisions.values() if d["total"] > 0)
        print(f"  ✓ Divisions: {active_divs} active CSI divisions extracted\n")

    # ── Write output files ───────────────────────────────────────────────────
    projects_path  = os.path.join(OUTPUT_DIR, "projects.json")
    divisions_path = os.path.join(OUTPUT_DIR, "divisions.json")

    with open(projects_path, "w", encoding="utf-8") as f:
        json.dump(all_projects, f, indent=2, ensure_ascii=False)

    with open(divisions_path, "w", encoding="utf-8") as f:
        json.dump(all_divisions, f, indent=2, ensure_ascii=False)

    # ── Summary ─────────────────────────────────────────────────────────────
    print("=" * 60)
    print(f"✅ Done!")
    print(f"   Projects extracted : {len(all_projects)}")
    print(f"   Division records   : {len(all_divisions)}")
    print(f"   Output → {projects_path}")
    print(f"   Output → {divisions_path}")
    print("=" * 60)

    # Print quick project summary table
    print("\nProject Summary:")
    print(f"{'#':<3} {'Name':<40} {'Price':>12} {'SF':>8} {'Margin':>8}")
    print("-" * 75)
    for i, p in enumerate(all_projects, 1):
        print(f"{i:<3} {p['estimate_name'][:39]:<40} ${p['estimate_price']:>11,.0f} {p['area_sf']:>8,.0f} {p['margin_pct']:>7.1f}%")


if __name__ == "__main__":
    main()