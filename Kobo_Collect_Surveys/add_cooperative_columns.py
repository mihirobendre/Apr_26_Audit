import pandas as pd
import re
import shutil

INPUT_FILE = "REGENERATIVE_CARBON_ALLIANCE_FARMER_ONBOARDING_FORM_-_all_versions_-_English_en_-_2026-04-10-09-28-08.xlsx"
OUTPUT_FILE = "REGENERATIVE_CARBON_ALLIANCE_with_cooperative.xlsx"

# Mapping: keyword patterns → (standardised name, county)
COOPERATIVE_MAP = [
    ("mum",  "Mumberes Farmers Co-operative Society", "Baringo"),
    ("lan",   "Lanyuak Farmers Co-operative Society",  "Narok"),
    ("eor",       "Eor Emaa Farmers Co-operative Society", "Narok"),
    ("kab",  "Kabianga Dairy Farmers Co-operative Society", "Kericho"),
    ("kip",  "Kipsigis Highlands Co-operative Society", "Bomet"),
]

def classify(raw):
    if pd.isna(raw):
        return None, None
    normalised = str(raw).lower().strip()
    for keyword, standard_name, county in COOPERATIVE_MAP:
        if keyword in normalised:
            return standard_name, county
    return str(raw).strip(), None   # keep original if unrecognised

df = pd.read_excel(INPUT_FILE)

# Combine the two cooperative name columns (prefer '2. CO-OPERATIVE NAME', fallback to '2. C-OPERATIVE NAME')
combined = df["2. CO-OPERATIVE NAME"].where(df["2. CO-OPERATIVE NAME"].notna(), df["2. C-OPERATIVE NAME"])

cooperative_col = []
county_col      = []
for val in combined:
    name, county = classify(val)
    cooperative_col.append(name)
    county_col.append(county)

df["Cooperative"]        = cooperative_col
df["Predominant County"] = county_col

df.to_excel(OUTPUT_FILE, index=False)
print(f"Done. Saved to {OUTPUT_FILE}")
print("\nCooperative value counts:")
print(df["Cooperative"].value_counts().to_string())
print("\nPredominant County value counts:")
print(df["Predominant County"].value_counts().to_string())
