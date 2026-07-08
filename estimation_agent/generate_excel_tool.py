"""
FIXED Excel generation tool for AGCM estimates.
Replace the old generate_excel_tool.py with this version.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import os

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
    """
    Generates a professional Excel spreadsheet of the estimate.
    
    Args:
        project_type: Type of construction
        area_sf: Total area in SF
        province: Province
        city: City
        comparable_projects: List of comparable project dicts
        divisions: List of CSI division benchmarks
        construction_cost: Total construction cost
        ohp_amount: OH&P dollar amount
        soft_cost: Soft costs dollar amount
        estimate_price: Final estimate price
        margin_pct: Margin percentage
        price_category: Price category (A/B/C/D)
        cost_per_sf: Cost per square foot
        schedule_weeks: Estimated schedule
    
    Returns:
        dict with: file_path, file_name, message, download_ready
    """
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Estimate"
    
    # Styles
    subheader_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    subheader_font = Font(bold=True, size=11)
    title_font = Font(bold=True, size=14, color="1F4E78")
    currency_format = "$#,##0.00"
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Set column widths
    ws.column_dimensions['A'].width = 28
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    
    row = 1
    
    # Title
    ws[f'A{row}'] = f"AGCM ESTIMATE — {project_type}"
    ws[f'A{row}'].font = title_font
    row += 2
    
    # Project Details
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
    
    # Comparable Projects
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
    
    # CSI Division Breakdown
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
    
    # Estimate Summary
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
    
    # Save
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