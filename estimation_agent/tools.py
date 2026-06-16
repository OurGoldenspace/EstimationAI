# estimation_agent/tools.py
import json
import os

# Load data on import
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

with open(os.path.join(DATA_DIR, "projects.json")) as f:
    PROJECTS = json.load(f)

with open(os.path.join(DATA_DIR, "divisions.json")) as f:
    DIVISIONS = json.load(f)

# ─── Tool 1: Find similar projects ─────────────────────────────

def find_similar_projects(
    project_type: str,
    area_sf: int,
    province: str,
    city: str = "",
    max_results: int = 3
) -> dict:
    """
    Search the historical bid database for the most similar past projects.
    Returns the top matching projects ranked by similarity.
    
    Args:
        project_type: Type of construction (Multi-Residential, Office, Retail, Industrial, Institutional, Mixed Use)
        area_sf: Project area in square feet
        province: Province (NB, NS, PEI)
        city: City name (optional, improves matching)
        max_results: Maximum number of results to return (default 3)
    """
    scored = []
    
    for proj in PROJECTS:
        # Skip pending projects
        if proj.get("status") == "Pending":
            continue
        
        score = 0
        reasons = []
        
        # Type match — highest weight
        if proj["type"].lower() == project_type.lower():
            score += 40
            reasons.append(f"Same type: {proj['type']}")
        
        # Province match
        if proj["province"].lower() == province.lower():
            score += 25
            reasons.append(f"Same province: {proj['province']}")
        
        # City match
        if city and proj["city"].lower() == city.lower():
            score += 10
            reasons.append(f"Same city: {proj['city']}")
        
        # Area similarity — within 30% gets full points
        area_diff = abs(proj["area_sf"] - area_sf) / max(area_sf, 1)
        if area_diff <= 0.15:
            score += 20
            reasons.append(f"Very similar size: {proj['area_sf']:,} SF ({area_diff*100:.0f}% diff)")
        elif area_diff <= 0.30:
            score += 15
            reasons.append(f"Similar size: {proj['area_sf']:,} SF ({area_diff*100:.0f}% diff)")
        elif area_diff <= 0.50:
            score += 8
            reasons.append(f"Moderate size diff: {proj['area_sf']:,} SF ({area_diff*100:.0f}% diff)")
        
        # Recency
        if proj["year"] >= 2024:
            score += 5
            reasons.append("Recent (2024)")
        elif proj["year"] >= 2023:
            score += 3
            reasons.append("Fairly recent (2023)")
        
        if score > 0:
            scored.append({
                "project": proj,
                "similarity_score": score,
                "match_reasons": reasons,
                "confidence": "HIGH" if score >= 60 else "MEDIUM" if score >= 40 else "LOW"
            })
    
    scored.sort(key=lambda x: x["similarity_score"], reverse=True)
    
    results = scored[:max_results]
    
    return {
        "query": f"{project_type}, {area_sf:,} SF, {city}, {province}",
        "matches_found": len(results),
        "total_in_database": len(PROJECTS),
        "results": results
    }


# ─── Tool 2: Get division benchmarks ───────────────────────────

def get_division_benchmarks(
    project_type: str,
    area_sf: int,
    province: str
) -> dict:
    """
    Get historical cost per SF benchmarks for every CSI division 
    based on similar past projects. Returns suggested $/SF for 
    each division with confidence levels.
    
    Args:
        project_type: Type of construction (Multi-Residential, Office, Retail, Industrial, Institutional, Mixed Use)
        area_sf: Project area in square feet
        province: Province (NB, NS, PEI)
    """
    # Find matching project names
    matching_names = []
    for proj in PROJECTS:
        if proj["type"].lower() == project_type.lower():
            matching_names.append(proj["name"])
        elif proj["province"].lower() == province.lower():
            matching_names.append(proj["name"])
    
    # Aggregate division costs from matching projects
    div_benchmarks = {}
    
    for div in DIVISIONS:
        div_num = div["Division"]
        
        if div["Project"] in matching_names:
            if div_num not in div_benchmarks:
                div_benchmarks[div_num] = {
                    "division": div_num,
                    "name": div["Division_Name"],
                    "costs_per_sf": [],
                    "projects": []
                }
            div_benchmarks[div_num]["costs_per_sf"].append(div["Cost_Per_SF"])
            div_benchmarks[div_num]["projects"].append(div["Project"])
    
    # Calculate averages and confidence
    results = {}
    for div_num, data in div_benchmarks.items():
        costs = data["costs_per_sf"]
        avg = sum(costs) / len(costs)
        low = min(costs)
        high = max(costs)
        count = len(costs)
        
        suggested_total = round(avg * area_sf, 0)
        
        results[div_num] = {
            "division": div_num,
            "name": data["name"],
            "suggested_per_sf": round(avg, 2),
            "range_low": round(low, 2),
            "range_high": round(high, 2),
            "suggested_total": suggested_total,
            "data_points": count,
            "comparable_projects": list(set(data["projects"])),
            "confidence": "HIGH" if count >= 3 else "MEDIUM" if count >= 2 else "LOW"
        }
    
    return {
        "area_sf": area_sf,
        "project_type": project_type,
        "divisions": results,
        "total_suggested": sum(d["suggested_total"] for d in results.values()),
        "total_per_sf": round(sum(d["suggested_per_sf"] for d in results.values()), 2)
    }


# ─── Tool 3: Calculate estimate total ──────────────────────────

def calculate_estimate_total(
    construction_cost: float,
    ohp_percent: float = 5.5,
    hst_percent: float = 15.0,
    area_sf: int = 0
) -> dict:
    """
    Calculate the full estimate total from construction cost.
    Applies OH&P and HST using Benchmark's standard rates.
    Also determines the price category and recommended margin.
    
    Args:
        construction_cost: Total construction cost before OH&P
        ohp_percent: Overhead and Profit percentage (default 5.5%)
        hst_percent: HST percentage (default 15%)
        area_sf: Project area in SF for per-SF calculations
    """
    ohp = construction_cost * (ohp_percent / 100)
    estimated_price = construction_cost + ohp
    hst = estimated_price * (hst_percent / 100)
    total_with_hst = estimated_price + hst
    
    # Determine price category and margin
    if estimated_price > 5000000:
        category = "A"
        target_margin = 10.0
    elif estimated_price > 2000000:
        category = "B"
        target_margin = 15.0
    elif estimated_price > 500000:
        category = "C"
        target_margin = 20.0
    else:
        category = "D"
        target_margin = 25.0
    
    result = {
        "construction_cost": round(construction_cost, 2),
        "ohp_percent": ohp_percent,
        "ohp_amount": round(ohp, 2),
        "estimated_price_before_hst": round(estimated_price, 2),
        "hst_percent": hst_percent,
        "hst_amount": round(hst, 2),
        "total_with_hst": round(total_with_hst, 2),
        "price_category": category,
        "target_margin": target_margin
    }
    
    if area_sf > 0:
        result["construction_per_sf"] = round(construction_cost / area_sf, 2)
        result["price_per_sf"] = round(estimated_price / area_sf, 2)
        result["total_per_sf"] = round(total_with_hst / area_sf, 2)
    
    return result


# ─── Tool 4: Validate required inputs ──────────────────────────

def validate_project_inputs(
    project_type: str = "",
    area_sf: int = 0,
    province: str = "",
    city: str = "",
    contract_type: str = ""
) -> dict:
    """
    Validates that the user has provided enough information to 
    generate a reliable estimate. Checks 5 mandatory fields and 
    reports which are missing.
    
    Args:
        project_type: Type of construction
        area_sf: Project area in square feet
        province: Province code (NB, NS, PEI)
        city: City name
        contract_type: Contract type (CCDC 5B, CCDC 2, Design Build)
    """
    mandatory = {
        "project_type": {
            "value": project_type,
            "valid": project_type.lower() in [
                "multi-residential", "office", "retail",
                "industrial", "institutional", "mixed use"
            ],
            "hint": "Must be one of: Multi-Residential, Office, Retail, Industrial, Institutional, Mixed Use"
        },
        "area_sf": {
            "value": area_sf,
            "valid": area_sf > 0,
            "hint": "Project area in square feet (e.g. 50000)"
        },
        "province": {
            "value": province,
            "valid": province.upper() in ["NB", "NS", "PEI", "NL"],
            "hint": "Province code: NB, NS, PEI, or NL"
        },
        "city": {
            "value": city,
            "valid": len(city) > 0,
            "hint": "City name (e.g. Moncton, Fredericton, Halifax)"
        },
        "contract_type": {
            "value": contract_type,
            "valid": contract_type in ["CCDC 5B", "CCDC 2", "CCDC 5A", "Design Build", ""],
            "hint": "CCDC 5B, CCDC 2, CCDC 5A, or Design Build"
        }
    }
    
    provided = sum(1 for v in mandatory.values() if v["valid"])
    missing = {k: v["hint"] for k, v in mandatory.items() if not v["valid"]}
    
    return {
        "fields_provided": provided,
        "fields_required": 5,
        "is_complete": provided >= 5,
        "can_estimate": provided >= 3,
        "missing_fields": missing,
        "message": (
            "All required fields provided. Ready to generate estimate."
            if provided >= 5
            else f"Missing {5-provided} fields. Please provide: {', '.join(missing.keys())}. "
                 f"Need at least 3 of 5 for a rough estimate."
            if provided < 3
            else f"Have {provided}/5 fields. Can generate estimate but {', '.join(missing.keys())} would improve accuracy."
        )
    }