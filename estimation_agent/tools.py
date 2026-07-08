"""
tools.py
========
Four ADK tools for the AGCM Estimation Agent.
Loads real project data from projects.json and divisions.json

Tools:
  1. validate_project_inputs      — checks what info was provided
  2. find_similar_projects        — finds past AGCM projects using weighted similarity
  3. get_division_benchmarks      — returns $/SF by CSI division
  4. calculate_estimate_total     — computes final price, margin, category, schedule
  5. generate_estimate_excel      — creates professional Excel estimate file
"""

import os
import json
import math
from typing import Optional
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─── DATA LOADING ──────────────────────────────────────────────────────────────

_DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

def _load_projects() -> list[dict]:
    path = os.path.join(_DATA_DIR, "projects.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def _load_divisions() -> list[dict]:
    path = os.path.join(_DATA_DIR, "divisions.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

_PROJECTS = _load_projects()
_DIVISIONS = _load_divisions()

# ─── PROJECT TYPE ALIASES ───────────────────────────────────────────────────────

_TYPE_ALIASES: dict[str, str] = {
    "multi-res": "Multi-Residential / Institutional",
    "multi res": "Multi-Residential / Institutional",
    "multires": "Multi-Residential / Institutional",
    "multi-residential": "Multi-Residential / Institutional",
    "residential": "Multi-Residential / Institutional",
    "apartment": "Multi-Residential / Institutional",
    "condo": "Multi-Residential / Institutional",
    "ymca": "Multi-Residential / Institutional",
    "daycare": "Multi-Residential / Institutional",
    "school": "Multi-Residential / Institutional",
    "institutional": "Multi-Residential / Institutional",
    "community": "Multi-Residential / Institutional",
    "dental": "Medical / Dental",
    "medical": "Medical / Dental",
    "clinic": "Medical / Dental",
    "wellness": "Medical / Dental",
    "health": "Medical / Dental",
    "pharmacy": "Medical / Dental",
    "doctor": "Medical / Dental",
    "grocery": "Grocery / Food Retail",
    "sobeys": "Grocery / Food Retail",
    "supermarket": "Grocery / Food Retail",
    "food": "Grocery / Food Retail",
    "foodland": "Grocery / Food Retail",
    "atlanticsuperstore": "Grocery / Food Retail",
    "automotive": "Automotive",
    "auto": "Automotive",
    "ford": "Automotive",
    "dealership": "Automotive",
    "car": "Automotive",
    "vehicle": "Automotive",
    "renovation": "Renovation / Tenant Fit-up",
    "reno": "Renovation / Tenant Fit-up",
    "fit-up": "Renovation / Tenant Fit-up",
    "fit up": "Renovation / Tenant Fit-up",
    "tenant": "Renovation / Tenant Fit-up",
    "fitout": "Renovation / Tenant Fit-up",
    "expansion": "Renovation / Tenant Fit-up",
    "retrofit": "Renovation / Tenant Fit-up",
    "office": "Commercial / Other",
    "commercial": "Commercial / Other",
    "retail": "Commercial / Other",
    "warehouse": "Commercial / Other",
    "industrial": "Commercial / Other",
    "mixed": "Commercial / Other",
}

def _normalize_type(raw: str) -> Optional[str]:
    """Map a free-text project type to a canonical type."""
    if not raw:
        return None
    key = raw.lower().strip().replace("-", " ")
    for alias, canonical in _TYPE_ALIASES.items():
        if alias in key:
            return canonical
    return raw.strip()

# ─── PRICE CATEGORY LOGIC ──────────────────────────────────────────────────────

def _price_category(price: float) -> tuple[str, float]:
    """Returns (category_label, suggested_margin) based on AGCM's pricing tiers."""
    if price >= 5_000_000:
        return "A (>$5M)", 0.10
    elif price >= 2_000_000:
        return "B ($2M–$5M)", 0.15
    elif price >= 500_000:
        return "C ($500K–$2M)", 0.20
    else:
        return "D (<$500K)", 0.25

# ─── SIMILARITY SCORING ────────────────────────────────────────────────────────

def _similarity_score(project: dict, project_type: str, area_sf: float,
                       province: str, city: str) -> float:
    """Weighted similarity score (0–100) between a past project and the query."""
    score = 0.0

    # 1. Project type (40 pts)
    canonical = _normalize_type(project_type) or ""
    if project["project_type"] == canonical:
        score += 40
    elif _partial_type_match(project["project_type"], canonical):
        score += 20

    # 2. Area similarity (30 pts) — log-scale
    if area_sf and project["area_sf"]:
        ratio = area_sf / project["area_sf"]
        log_diff = abs(math.log(ratio))
        area_score = max(0.0, 30 * (1 - log_diff / math.log(5)))
        score += area_score

    # 3. Province (15 pts)
    if province and project["province"]:
        if project["province"].upper() == province.upper():
            score += 15

    # 4. City (10 pts)
    if city and project["city"]:
        if project["city"].lower() == city.lower():
            score += 10
        elif city.lower() in project["city"].lower() or project["city"].lower() in city.lower():
            score += 5

    # 5. Recency (5 pts)
    year = project.get("budget_year") or 2020
    recency = max(0, 5 - (2026 - int(year)))
    score += recency

    return round(score, 1)

def _partial_type_match(type_a: str, type_b: str) -> bool:
    """Check if two project types share a broad category."""
    broad = [
        {"Multi-Residential / Institutional"},
        {"Medical / Dental", "Commercial / Other"},
        {"Grocery / Food Retail", "Retail"},
        {"Automotive"},
        {"Renovation / Tenant Fit-up"},
    ]
    for group in broad:
        if type_a in group and type_b in group:
            return True
    return False

def _confidence_label(score: float) -> str:
    if score >= 70:
        return "HIGH"
    elif score >= 45:
        return "MEDIUM"
    elif score >= 25:
        return "LOW"
    else:
        return "VERY LOW"

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 1 — validate_project_inputs
# ══════════════════════════════════════════════════════════════════════════════

def validate_project_inputs(
    project_type: str,
    area_sf: Optional[float] = None,
    province: Optional[str] = None,
    city: Optional[str] = None,
    contract_type: Optional[str] = None,
) -> dict:
    """Validates the project inputs provided by the user."""
    fields_provided = []
    fields_missing = []
    warnings = []

    canonical_type = None
    if project_type:
        canonical_type = _normalize_type(project_type)
        fields_provided.append("project_type")
        if canonical_type == project_type.strip() and project_type.strip() not in _TYPE_ALIASES:
            warnings.append(
                f"Project type '{project_type}' is not a recognized AGCM category. "
                f"Will attempt to find closest match. Known types: "
                f"Multi-Residential, Medical/Dental, Grocery, Automotive, Renovation/Tenant Fit-up, Commercial."
            )
    else:
        fields_missing.append("project_type")

    if area_sf and area_sf > 0:
        fields_provided.append("area_sf")
        if area_sf < 500:
            warnings.append(f"Area of {area_sf:,.0f} SF is very small — estimate confidence will be low.")
        elif area_sf > 200_000:
            warnings.append(f"Area of {area_sf:,.0f} SF is larger than any project in the AGCM database — comparables will be extrapolated.")
    else:
        fields_missing.append("area_sf")

    if province:
        fields_provided.append("province")
        if province.upper() not in ["NB", "NS", "PEI", "NL", "PE"]:
            warnings.append(f"Province '{province}' is outside Atlantic Canada — no direct comparables, will use NB benchmarks.")
    else:
        fields_missing.append("province")
        warnings.append("Province not provided — defaulting to NB benchmarks.")

    if city:
        fields_provided.append("city")
    else:
        fields_missing.append("city")

    if contract_type:
        fields_provided.append("contract_type")
    else:
        fields_missing.append("contract_type")

    can_estimate = "project_type" in fields_provided and "area_sf" in fields_provided

    return {
        "fields_provided": fields_provided,
        "fields_missing": fields_missing,
        "can_estimate": can_estimate,
        "canonical_type": canonical_type,
        "warnings": warnings,
        "message": (
            "Sufficient information to generate an estimate."
            if can_estimate
            else f"Cannot estimate yet. Required fields missing: {', '.join(fields_missing)}. "
                 f"Please ask the user for: {' and '.join(fields_missing)}."
        ),
    }

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 2 — find_similar_projects
# ══════════════════════════════════════════════════════════════════════════════

def find_similar_projects(
    project_type: str,
    area_sf: float,
    province: str = "NB",
    city: str = "",
    top_n: int = 3,
) -> dict:
    """Finds the most similar past AGCM projects using weighted scoring."""
    projects_with_divs = set(str(d["estimate_number"]) for d in _DIVISIONS)
    projects_to_search = [p for p in _PROJECTS if str(p["estimate_number"]) in projects_with_divs]
    
    scored = []
    for p in projects_to_search:
        score = _similarity_score(p, project_type, area_sf, province, city)
        scored.append((score, p))

    scored.sort(key=lambda x: x[0], reverse=True)
    top = scored[:top_n]

    matches = []
    for score, p in top:
        size_ratio = area_sf / p["area_sf"] if p["area_sf"] else 1.0
        size_diff_pct = round((size_ratio - 1) * 100, 1)

        matches.append({
            "estimate_number":  p["estimate_number"],
            "estimate_name":    p["estimate_name"],
            "project_type":     p["project_type"],
            "area_sf":          p["area_sf"],
            "area_diff_pct":    size_diff_pct,
            "province":         p["province"],
            "city":             p["city"],
            "budget_year":      p["budget_year"],
            "estimate_price":   p["estimate_price"],
            "cost_per_sf":      p["cost_per_sf"],
            "margin_pct":       p["margin_pct"],
            "schedule_weeks":   p["schedule_weeks"],
            "contract_type":    p["contract_type"],
            "result":           p["result"],
            "similarity_score": score,
            "confidence":       _confidence_label(score),
        })

    total_weight = sum(s for s, _ in top)
    if total_weight > 0:
        avg_cost_per_sf = sum(s * p["cost_per_sf"] for s, p in top) / total_weight
    else:
        avg_cost_per_sf = 0

    best_score = top[0][0] if top else 0
    coverage_warning = None
    if best_score < 30:
        coverage_warning = (
            f"Best match score is only {best_score}/100 — "
            f"no strong comparable exists in the AGCM database for this project type/size. "
            f"Estimate confidence is LOW."
        )
    elif best_score < 55:
        coverage_warning = (
            f"Moderate match quality (score {best_score}/100). "
            f"Comparable projects differ in type or size. Treat as indicative only."
        )

    return {
        "matches": matches,
        "total_projects_in_database": len(projects_to_search),
        "avg_cost_per_sf_weighted": round(avg_cost_per_sf, 2),
        "coverage_warning": coverage_warning,
        "query": {
            "project_type": _normalize_type(project_type),
            "area_sf": area_sf,
            "province": province,
            "city": city,
        },
    }

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 3 — get_division_benchmarks
# ══════════════════════════════════════════════════════════════════════════════

def get_division_benchmarks(
    project_type: str,
    area_sf: float,
    province: str = "NB",
    comparable_estimate_numbers: Optional[list[str]] = None,
) -> dict:
    """Returns CSI division-level $/SF benchmarks for the new project."""
    if comparable_estimate_numbers:
        source_nums = set(str(x) for x in comparable_estimate_numbers)
        source_divs = [d for d in _DIVISIONS if str(d["estimate_number"]) in source_nums]
    else:
        result = find_similar_projects(project_type, area_sf, province, top_n=3)
        source_nums = set(str(m["estimate_number"]) for m in result["matches"])
        source_divs = [d for d in _DIVISIONS if str(d["estimate_number"]) in source_nums]

    if not source_divs:
        return {
            "error": "No division data found for the selected comparable projects.",
            "divisions": [],
            "total_construction_cost": 0,
            "total_cost_per_sf": 0,
        }

    division_groups: dict[str, list] = defaultdict(list)
    for d in source_divs:
        division_groups[d["division_code"]].append(d)

    benchmarks = []
    total_cost_per_sf = 0.0

    for code in sorted(division_groups.keys()):
        entries = division_groups[code]
        div_name = entries[0]["division_name"]

        weights = []
        for e in entries:
            if e["area_sf"] and e["area_sf"] > 0:
                ratio = area_sf / e["area_sf"]
                w = 1 / (1 + abs(math.log(ratio)))
            else:
                w = 0.5
            weights.append(w)

        total_w = sum(weights)
        if total_w > 0:
            avg_cost_per_sf = sum(e["cost_per_sf"] * w for e, w in zip(entries, weights)) / total_w
            avg_pct = sum(e["pct_of_project"] * w for e, w in zip(entries, weights)) / total_w
        else:
            avg_cost_per_sf = 0
            avg_pct = 0

        estimated_total = round(avg_cost_per_sf * area_sf, 2)

        cpf_values = [e["cost_per_sf"] for e in entries if e["cost_per_sf"] > 0]
        range_low = min(cpf_values) if cpf_values else 0
        range_high = max(cpf_values) if cpf_values else 0

        benchmarks.append({
            "division_code":        code,
            "division_name":        div_name,
            "cost_per_sf":          round(avg_cost_per_sf, 2),
            "estimated_total":      estimated_total,
            "pct_of_project":       round(avg_pct, 2),
            "range_low_per_sf":     round(range_low, 2),
            "range_high_per_sf":    round(range_high, 2),
            "source_project_count": len(entries),
            "comparable_projects":  [e["estimate_name"] for e in entries],
        })

        total_cost_per_sf += avg_cost_per_sf

    total_construction_cost = round(total_cost_per_sf * area_sf, 2)

    return {
        "divisions":                benchmarks,
        "total_construction_cost":  total_construction_cost,
        "total_cost_per_sf":        round(total_cost_per_sf, 2),
        "area_sf":                  area_sf,
        "source_projects":          [str(x) for x in source_nums],
        "division_count":           len(benchmarks),
    }

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 4 — calculate_estimate_total
# ══════════════════════════════════════════════════════════════════════════════

def calculate_estimate_total(
    construction_cost: float,
    area_sf: float,
    ohp_pct: Optional[float] = None,
    soft_cost: Optional[float] = None,
    override_margin_pct: Optional[float] = None,
) -> dict:
    """Calculates the final estimate price, margin, price category, and schedule."""
    
    ohp_pct_applied = ohp_pct if ohp_pct is not None else 5.0
    soft_cost_applied = soft_cost if soft_cost is not None else round(construction_cost * 0.05, 2)

    ohp_amount = round(construction_cost * (ohp_pct_applied / 100), 2)
    subtotal = construction_cost + ohp_amount + soft_cost_applied

    first_pass_price = subtotal / (1 - 0.10)
    category_label, suggested_margin = _price_category(first_pass_price)

    margin_pct = override_margin_pct if override_margin_pct is not None else suggested_margin * 100

    estimate_price = round(subtotal / (1 - margin_pct / 100), 2)
    category_label, _ = _price_category(estimate_price)

    cost_per_sf = round(estimate_price / area_sf, 2) if area_sf else 0

    avg_cost_per_week = 12000 if estimate_price < 2_000_000 else 13500
    schedule_weeks = max(4, round(estimate_price / avg_cost_per_week))

    return {
        "estimate_price":        estimate_price,
        "margin_pct":            margin_pct,
        "price_category":        category_label,
        "cost_per_sf":           cost_per_sf,
        "schedule_weeks":        schedule_weeks,
        "breakdown": {
            "construction_cost": construction_cost,
            "ohp_pct":           ohp_pct_applied,
            "ohp_amount":        ohp_amount,
            "soft_cost":         soft_cost_applied,
            "subtotal":          subtotal,
            "margin_pct":        margin_pct,
            "final_price":       estimate_price,
        },
        "notes": [
            f"Price category: {category_label} — AGCM standard margin for this range is {suggested_margin*100:.0f}%.",
            f"Schedule of {schedule_weeks} weeks is an estimate based on AGCM's historical cost-per-week ratio.",
            "Soft costs ({:.1f}% of construction) cover design, permits, surveys.".format(
                soft_cost_applied / construction_cost * 100 if construction_cost else 0
            ),
        ],
    }

# ══════════════════════════════════════════════════════════════════════════════
# TOOL 5 — generate_estimate_excel
# ══════════════════════════════════════════════════════════════════════════════

def generate_estimate_excel(
    project_type: str,
    area_sf: float,
    province: str,
    city: str,
    comparable_projects: list[dict],
    divisions: list[dict],
    construction_cost: float,
    ohp_amount: float,
    soft_cost: float,
    estimate_price: float,
    margin_pct: float,
    price_category: str,
    cost_per_sf: float,
    schedule_weeks: int,
) -> dict:
    """Generates a professional Excel spreadsheet of the estimate."""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Estimate"
    
    subheader_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    subheader_font = Font(bold=True, size=11)
    title_font = Font(bold=True, size=14, color="1F4E78")
    currency_format = "$#,##0.00"
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    
    row = 1
    
    ws[f'A{row}'] = f"AGCM ESTIMATE — {project_type}"
    ws[f'A{row}'].font = title_font
    row += 2
    
    ws[f'A{row}'] = "PROJECT DETAILS"
    ws[f'A{row}'].font = subheader_font
    row += 1
    
    details = [
        ("Project Type:", project_type),
        ("Area:", f"{area_sf:,.0f} SF"),
        ("Location:", f"{city}, {province}"),
        ("Estimate Date:", datetime.now().strftime("%B %d, %Y")),
    ]
    
    for label, value in details:
        ws[f'A{row}'] = label
        ws[f'B{row}'] = value
        ws[f'A{row}'].font = Font(bold=True, size=10)
        row += 1
    
    row += 1
    
    ws[f'A{row}'] = "COMPARABLE PROJECTS"
    ws[f'A{row}'].font = subheader_font
    row += 1
    
    headers = ["Project Name", "Area (SF)", "Cost/SF", "Margin %", "Similarity"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col)
        cell.value = header
        cell.fill = subheader_fill
        cell.font = subheader_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border
    
    row += 1
    
    for proj in comparable_projects:
        ws.cell(row=row, column=1).value = proj.get("estimate_name", "")
        ws.cell(row=row, column=2).value = proj.get("area_sf", 0)
        ws.cell(row=row, column=3).value = proj.get("cost_per_sf", 0)
        ws.cell(row=row, column=4).value = proj.get("margin_pct", 0)
        ws.cell(row=row, column=5).value = proj.get("similarity_score", 0)
        
        ws.cell(row=row, column=2).number_format = "#,##0"
        ws.cell(row=row, column=3).number_format = currency_format
        ws.cell(row=row, column=4).number_format = "0.00%"
        
        for col in range(1, 6):
            ws.cell(row=row, column=col).border = thin_border
        
        row += 1
    
    row += 1
    
    ws[f'A{row}'] = "CSI DIVISION BREAKDOWN"
    ws[f'A{row}'].font = subheader_font
    row += 1
    
    div_headers = ["Division Code", "Division Name", "$/SF", "Estimated Total"]
    for col, header in enumerate(div_headers, 1):
        cell = ws.cell(row=row, column=col)
        cell.value = header
        cell.fill = subheader_fill
        cell.font = subheader_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    
    row += 1
    
    for div in divisions:
        ws.cell(row=row, column=1).value = div.get("division_code", "")
        ws.cell(row=row, column=2).value = div.get("division_name", "")
        ws.cell(row=row, column=3).value = div.get("cost_per_sf", 0)
        ws.cell(row=row, column=4).value = div.get("estimated_total", 0)
        
        ws.cell(row=row, column=3).number_format = currency_format
        ws.cell(row=row, column=4).number_format = currency_format
        
        for col in range(1, 5):
            ws.cell(row=row, column=col).border = thin_border
        
        row += 1
    
    row += 1
    
    ws[f'A{row}'] = "ESTIMATE SUMMARY"
    ws[f'A{row}'].font = subheader_font
    row += 1
    
    ws[f'A{row}'] = "Construction Cost:"
    ws[f'B{row}'] = construction_cost
    ws[f'B{row}'].number_format = currency_format
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 1
    
    ws[f'A{row}'] = "OH&P (5%):"
    ws[f'B{row}'] = ohp_amount
    ws[f'B{row}'].number_format = currency_format
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 1
    
    ws[f'A{row}'] = "Soft Costs (5%):"
    ws[f'B{row}'] = soft_cost
    ws[f'B{row}'].number_format = currency_format
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 2
    
    ws[f'A{row}'] = "★ ESTIMATE PRICE:"
    ws[f'A{row}'].font = Font(bold=True, size=11, color="FF0000")
    ws[f'B{row}'] = estimate_price
    ws[f'B{row}'].number_format = currency_format
    ws[f'B{row}'].font = Font(bold=True, size=11, color="FF0000")
    row += 2
    
    ws[f'A{row}'] = "Price Category:"
    ws[f'B{row}'] = price_category
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 1
    
    ws[f'A{row}'] = "Margin:"
    ws[f'B{row}'] = margin_pct / 100
    ws[f'B{row}'].number_format = "0.0%"
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 1
    
    ws[f'A{row}'] = "Cost per SF:"
    ws[f'B{row}'] = cost_per_sf
    ws[f'B{row}'].number_format = currency_format
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 1
    
    ws[f'A{row}'] = "Estimated Schedule:"
    ws[f'B{row}'] = f"{schedule_weeks} weeks"
    ws[f'A{row}'].font = Font(bold=True, size=10)
    row += 1
    
    output_dir = os.path.join(os.path.dirname(__file__), "estimates")
    os.makedirs(output_dir, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"AGCM_Estimate_{project_type.replace(' ', '_')}_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    wb.save(filepath)
    
    return {
        "status": "success",
        "file_path": filepath,
        "file_name": filename,
        "message": f"✓ Excel estimate generated: {filename}",
        "download_ready": True,
    }

# ─── QUICK SELF-TEST ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=== Tool Self-Test ===\n")

    print("1. validate_project_inputs()")
    v = validate_project_inputs("dental office", 2000, "NB", "Moncton")
    print(json.dumps(v, indent=2))

    print("\n2. find_similar_projects()")
    s = find_similar_projects("dental", 2000, "NB", "Moncton")
    for m in s["matches"]:
        print(f"   [{m['similarity_score']:5.1f}] {m['estimate_name']} "
              f"({m['area_sf']:,.0f} SF, {m['margin_pct']}% margin, {m['confidence']})")

    print("\n3. get_division_benchmarks()")
    b = get_division_benchmarks("dental", 2000, "NB")
    print(f"   Total $/SF: ${b['total_cost_per_sf']:.2f}")
    print(f"   Total construction cost: ${b['total_construction_cost']:,.0f}")
    for d in b["divisions"]:
        if d["cost_per_sf"] > 0:
            print(f"   {d['division_code']} {d['division_name'][:30]:<30} "
                  f"${d['cost_per_sf']:>6.2f}/SF  ${d['estimated_total']:>10,.0f}")

    print("\n4. calculate_estimate_total()")
    c = calculate_estimate_total(b["total_construction_cost"], 2000)
    print(f"   Estimate price  : ${c['estimate_price']:,.0f}")
    print(f"   Price category  : {c['price_category']}")
    print(f"   Margin          : {c['margin_pct']}%")
    print(f"   Cost per SF     : ${c['cost_per_sf']:.2f}")
    print(f"   Schedule        : {c['schedule_weeks']} weeks")

    print("\n5. generate_estimate_excel()")
    e = generate_estimate_excel(
        project_type="Medical / Dental",
        area_sf=2000,
        province="NB",
        city="Moncton",
        comparable_projects=s["matches"],
        divisions=b["divisions"],
        construction_cost=b["total_construction_cost"],
        ohp_amount=c["breakdown"]["ohp_amount"],
        soft_cost=c["breakdown"]["soft_cost"],
        estimate_price=c["estimate_price"],
        margin_pct=c["margin_pct"],
        price_category=c["price_category"],
        cost_per_sf=c["cost_per_sf"],
        schedule_weeks=c["schedule_weeks"],
    )
    print(f"   Excel file: {e['file_name']}")
    print(f"   Status: {e['status']}")