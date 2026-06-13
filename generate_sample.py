"""
generate_sample.py — synthetic, clearly-fake sample CSV for the PACES generators.
Matches the 28-column schema in INPUT_SPECIFICATION.md. All facility names and
numbers are invented; this is for smoke-testing the scripts, not real APCD data.
"""
import csv
from pathlib import Path

OUT = Path("/sessions/serene-dazzling-tesla/mnt/outputs/sample_data.csv")

HEADER = ["Clinical Chapter","Episode Name","Facility Name","Facility Hospital System",
 "Facility County Designation","Facility County","Facility HSR","Facility RAE Region",
 "Facility DOI Category","Medicaid Volume","High Risk Volume","Normal Risk Volume",
 "2021 IP Base Rate","2021 OP Base Rate","2022 IP Base Rate","2022 OP Base Rate",
 "2023 IP Base Rate","2023 OP Base Rate","2024 IP Base Rate","2024 OP Base Rate",
 "2025 IP Base Rate","2025 OP Base Rate","CDPS Facility Risk Score","Total Actual Cost",
 "Average Actual Cost","Average Base Rate Adjusted (Observed) Cost",
 "Average Risk Adjusted Cost","Facility O:E Cost Ratio"]

# facility -> (system, county designation, county, HSR, RAE, DOI, CDPS)
FAC = {
 "Example Regional Medical Center": ("Example Health","Urban","Sample County A",2,"4","Large Metro",1.35),
 "Sample Community Hospital North": ("Example Health","Urban","Sample County A",2,"4","Metro",1.10),
 "Demo Memorial Hospital":          ("Demo Care Network","Urban","Sample County B",4,"3","Metro",1.55),
 "Sample Valley Medical Center":    ("Demo Care Network","Urban","Sample County C",3,"2","Metro",1.20),
 "Example University Hospital":     ("Sample University Health","Urban","Sample County A",2,"4","Large Metro",1.70),
 "Example Rural Clinic":            ("No System","Rural","Sample County D",1,"1","Rural",0.95),
 "Demo Micro Hospital":             ("No System","Rural","Sample County E",1,"1","Micro",1.05),
 "Sample CEAC Facility":            ("No System","Rural","Sample County F",4,"3","CEAC",1.15),
}
# episode -> (clinical chapter, expected per-episode risk-adjusted cost $)
EP = {
 "CABG (including Cardiac catheterization)": ("Cardio-Vascular System",58000),
 "Open heart valve surgery (including Cardiac catheterization)": ("Cardio-Vascular System",72000),
 "Percutaneous cardiac intervention (including Cardiac catheterization)": ("Cardio-Vascular System",30000),
 "Thyroidectomy": ("Endocrine System",12500),
 "Cholecystectomy": ("Digestive System",14000),
 "Colectomy": ("Digestive System",33000),
 "Esophagogastroduodenoscopy (Upper Endoscopy)": ("Digestive System",4200),
 "Repair Ventral Hernia": ("Digestive System",16500),
 "Mastectomy": ("Breast",18500),
}

# (facility, episode, medicaid_volume, high_risk_volume, O:E)
# Designed coverage:
#   Demo Memorial Hospital -> outlier (O:E>1.10) in 4 episodes  => SYSTEMIC
#   Example University Hospital -> outlier in 2 episodes (not systemic)
#   mix of non-outliers; low/zero-volume + no-O:E edge cases for DQ flags
ROWS = [
 # --- CABG ---
 ("Example Regional Medical Center","CABG (including Cardiac catheterization)",34,12,1.07),
 ("Demo Memorial Hospital","CABG (including Cardiac catheterization)",28,14,1.42),
 ("Example University Hospital","CABG (including Cardiac catheterization)",22,11,1.55),
 ("Sample Valley Medical Center","CABG (including Cardiac catheterization)",18,6,0.94),
 # --- Open heart valve ---
 ("Demo Memorial Hospital","Open heart valve surgery (including Cardiac catheterization)",16,9,1.33),
 ("Example University Hospital","Open heart valve surgery (including Cardiac catheterization)",14,8,1.18),
 ("Example Regional Medical Center","Open heart valve surgery (including Cardiac catheterization)",12,5,1.02),
 # --- PCI ---
 ("Example Regional Medical Center","Percutaneous cardiac intervention (including Cardiac catheterization)",57,20,1.05),
 ("Demo Memorial Hospital","Percutaneous cardiac intervention (including Cardiac catheterization)",41,18,1.26),
 ("Sample Community Hospital North","Percutaneous cardiac intervention (including Cardiac catheterization)",33,10,0.98),
 ("Sample Valley Medical Center","Percutaneous cardiac intervention (including Cardiac catheterization)",26,8,1.12),
 # --- Thyroidectomy ---
 ("Example University Hospital","Thyroidectomy",24,5,0.91),
 ("Sample Community Hospital North","Thyroidectomy",19,4,1.04),
 ("Demo Memorial Hospital","Thyroidectomy",12,3,1.21),
 # --- Cholecystectomy ---
 ("Sample Community Hospital North","Cholecystectomy",61,14,1.03),
 ("Example Regional Medical Center","Cholecystectomy",48,16,1.16),
 ("Sample Valley Medical Center","Cholecystectomy",37,9,0.99),
 ("Demo Micro Hospital","Cholecystectomy",8,2,1.25),
 # --- Colectomy ---
 ("Example University Hospital","Colectomy",29,15,1.34),
 ("Demo Memorial Hospital","Colectomy",24,13,1.48),
 ("Example Regional Medical Center","Colectomy",21,9,1.07),
 ("Sample CEAC Facility","Colectomy",6,2,1.13),
 # --- EGD ---
 ("Sample Community Hospital North","Esophagogastroduodenoscopy (Upper Endoscopy)",128,30,1.02),
 ("Example Regional Medical Center","Esophagogastroduodenoscopy (Upper Endoscopy)",96,28,1.19),
 ("Sample Valley Medical Center","Esophagogastroduodenoscopy (Upper Endoscopy)",74,15,0.96),
 ("Example Rural Clinic","Esophagogastroduodenoscopy (Upper Endoscopy)",11,2,1.22),
 # --- Repair Ventral Hernia ---
 ("Example Regional Medical Center","Repair Ventral Hernia",44,13,1.10),
 ("Demo Memorial Hospital","Repair Ventral Hernia",31,12,1.29),
 ("Sample Community Hospital North","Repair Ventral Hernia",27,7,1.01),
 # --- Mastectomy ---
 ("Example University Hospital","Mastectomy",33,12,1.22),
 ("Sample Valley Medical Center","Mastectomy",26,8,1.06),
 ("Sample CEAC Facility","Mastectomy",9,3,1.14),
 # --- Data-quality edge cases ---
 ("Demo Micro Hospital","Thyroidectomy",4,1,1.30),        # low_vol (excluded from analytical sheets)
 ("Example Rural Clinic","Cholecystectomy",2,0,1.40),      # very_low_vol
 ("Sample CEAC Facility","Esophagogastroduodenoscopy (Upper Endoscopy)",0,0,0.0),  # zero_vol;no_OE
 ("Example Rural Clinic","Repair Ventral Hernia",10,3,0.0),# no_OE (kept, non-outlier)
]

def br(expected, ip_share, yr):
    # plausible historical base rate, trending ~3%/yr from 2021
    base = expected*ip_share
    return round(base*(1.03**(yr-2021)))

rows_out=[]
for fac,ep,vol,hr,oe in ROWS:
    system,desig,county,hsr,rae,doi,cdps = FAC[fac]
    chapter,expected = EP[ep]
    normal = max(vol-hr,0)
    observed_per = round(expected*oe) if oe>0 else 0       # Avg Base Rate Adjusted (Observed)
    avg_actual = observed_per                               # actual ≈ observed per episode
    total_actual = avg_actual*vol
    rows_out.append([
        chapter, ep, fac, system, desig, county, hsr, rae, doi,
        vol, hr, normal,
        br(expected,0.62,2021), br(expected,0.20,2021),
        br(expected,0.62,2022), br(expected,0.20,2022),
        br(expected,0.62,2023), br(expected,0.20,2023),
        br(expected,0.62,2024), br(expected,0.20,2024),
        br(expected,0.62,2025), br(expected,0.20,2025),
        cdps,
        f"${total_actual:,}", f"${avg_actual:,}", f"${observed_per:,}", f"${expected:,}",
        f"{oe:.2f}",
    ])

with open(OUT,"w",newline="",encoding="utf-8") as f:
    w=csv.writer(f); w.writerow(HEADER); w.writerows(rows_out)
print(f"wrote {OUT}  rows={len(rows_out)} facilities={len(FAC)} episodes={len(EP)}")
