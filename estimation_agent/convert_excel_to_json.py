import pandas as pd
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "estimation_agent" / "data"

DATA_DIR.mkdir(parents=True, exist_ok=True)

excel_file = BASE_DIR / "your_excel_file.xlsx"

# Convert Projects sheet
projects_df = pd.read_excel(excel_file, sheet_name="Projects")
projects = projects_df.to_dict(orient="records")

with open(DATA_DIR / "projects.json", "w", encoding="utf-8") as f:
    json.dump(projects, f, indent=2)

# Convert Divisions sheet
divisions_df = pd.read_excel(excel_file, sheet_name="Divisions")
divisions = divisions_df.to_dict(orient="records")

with open(DATA_DIR / "divisions.json", "w", encoding="utf-8") as f:
    json.dump(divisions, f, indent=2)

print("Converted Excel sheets to JSON successfully.")