# Input Specification — PACES CFO Workbook Generator

This document describes the input file requirements for `build_cfo_workbook.py`. KPMG analysts running future APCD data cycles should validate their CSV against this spec before running the script.

---

## File Format

- **Format:** CSV (comma-separated values), UTF-8 or UTF-8 with BOM
- **First row:** Column headers (case-sensitive, exact match required)
- **Subsequent rows:** One row per facility-episode combination
- **Blank rows:** Permitted between groups (e.g., between episodes); the script filters them out automatically based on Episode Name being populated

---

## Required Columns

The script will exit with a clear error if any of these columns are missing. Header names must match exactly, case-sensitive, including spaces and punctuation.

| Column Header (exact) | Type | Description | Used For |
|---|---|---|---|
| `Episode Name` | string | Full episode label from PACES grouper | Primary grouping key; matched against config for classification |
| `Facility Name` | string | Hospital/facility name | Facility identification across episodes |
| `Facility Hospital System` | string | System affiliation (e.g., "UCHealth", "CommonSpirit") or "No System" for independents | Worksheet 10 rollup |
| `Facility RAE Region` | integer (1–4) | Regional Accountable Entity catchment | Worksheet 8 rollup |
| `Facility DOI Category` | string | Urbanity tier — typically "Metro", "Large Metro", "Micro", "Rural", "CEAC" | Worksheet 9 rollup |
| `Medicaid Volume` | number | Number of Medicaid episodes at this facility | Volume filter (default ≥ 5) |
| `High Risk Volume` | number | Subset of Medicaid Volume classified as high-risk | % High Risk Volume computation |
| `Total Actual Cost` | number (can include `$` and `,`) | Annual actual spend in dollars | Outlier excess computation |
| `Facility O:E Cost Ratio` | number | Observed / Expected cost ratio (risk-adjusted) | Outlier threshold check; rank within episode |

### Number formatting

The parser accepts:
- Plain numbers: `1234.56`
- With dollar sign: `$1,234.56`
- With commas: `1,234.56`
- With percent: `45.0%`
- Quoted strings: `"$1,234.56 "`

Whitespace, dollar signs, commas, and percent signs are stripped before parsing.

---

## Optional Columns

If present, these columns are preserved in the All Data and Analytical View worksheets but are not required for computation:

| Column Header | Description |
|---|---|
| `Clinical Chapter` | High-level grouping (e.g., "Cardio-Vascular System") |
| `Facility County Designation` | Urban / Rural designation |
| `Facility County` | County name — used in Worksheet 5 if present |
| `Facility HSR` | Health Service Region |
| `Normal Risk Volume` | Subset of Medicaid Volume classified as normal-risk |
| `2021 IP Base Rate`, `2021 OP Base Rate`, ... `2025 IP Base Rate`, `2025 OP Base Rate` | Historical base rates per year |
| `CDPS Facility Risk Score` | Chronic disability payment system facility risk |
| `Average Actual Cost` | Per-episode actual cost |
| `Average Base Rate Adjusted (Observed) Cost` | Per-episode observed cost |
| `Average Risk Adjusted Cost` | Per-episode expected cost |

The script reads these columns through without modification and they appear in worksheets 2–4.

---

## What the Script Computes (you do NOT provide these)

These six columns are added by the script in worksheets 2, 3, and 4:

| Computed Column | Formula / Source |
|---|---|
| `Episode Classification` | Looked up in `config.json` → `episode_classifications` |
| `% High Risk Volume` | `High Risk Volume ÷ Medicaid Volume × 100` |
| `O:E Rank (within Episode)` | Rank ascending by O:E within each episode group (1 = best) |
| `Outlier Excess $` | `Total Actual Cost × (1 − 1/O:E)` if `O:E > threshold`, else blank |
| `Is Outlier (O:E>threshold)` | "Yes" / "No" |
| `Data Quality Flag` | `OK`, `low_vol`, `very_low_vol`, `zero_vol`, or `no_OE` |

---

## Data Quality Considerations

The script applies these defaults; they can be overridden via CLI flags or config:

| Concern | Default Treatment |
|---|---|
| Medicaid Volume < 5 | Flagged `low_vol`; excluded from analytical worksheets (3–10) |
| Medicaid Volume < 3 | Flagged `very_low_vol` |
| Medicaid Volume = 0 | Flagged `zero_vol` |
| Missing or zero O:E | Flagged `no_OE`; treated as non-outlier |
| Negative O:E | Treated as non-outlier (excess = 0) |

The volume filter exists because facilities with very small Medicaid volume produce unstable O:E ratios — a single high-cost case at a facility with n=2 will produce an apparent outlier that does not reflect facility performance. The vol ≥ 5 threshold is conservative and consistent with published Medicaid episode-program practice.

---

## Episode Classification Maintenance

When KPMG adds episodes to a future data cycle:

1. Open `config.json`
2. Under `episode_classifications`, add the new episode name as a key and assign `"Standard"` or `"Heterogeneous"` as the value
3. Re-run the script

If an episode appears in the data that is not in the classification config, the script defaults to `Heterogeneous` and prints a warning at runtime listing every unclassified episode. The output workbook still generates correctly, but the addressable % defaults to the Heterogeneous value (50% by default).

### Classification framework

- **Standard** — Episodes with a tight procedure list and where the diagnosis is not the cost driver. Outliers reflect real performance gaps (acuity not captured by risk adjustment + complications + post-procedural service overuse). Examples: CABG, Open Heart Valve Surgery, PCI, Thyroidectomy. Default addressable: 70%.
- **Heterogeneous** — Episodes that bundle different procedures or diagnoses into one definition, so case-mix variance contributes to apparent outliers. Engagement can address part but not all of the excess. Examples: EGD (diagnostic + therapeutic + surveillance), Mastectomy (lumpectomy vs full mastectomy), Colectomy (cancer vs IBD vs diverticular). Default addressable: 50%.

---

## Running the Script

### Minimum invocation

```bash
python build_cfo_workbook.py "Latest_data.csv"
```

Output: `Latest_data_workbook.xlsx` in the same folder as the input.

### Specifying output location and config

```bash
python build_cfo_workbook.py "Latest_data.csv" \
  --output "Q3_2026_CFO_Workbook.xlsx" \
  --config "config.json"
```

### Overriding filters at runtime

```bash
python build_cfo_workbook.py "Latest_data.csv" --min-volume 10 --oe-threshold 1.05
```

### Help

```bash
python build_cfo_workbook.py --help
```

---

## Dependencies

- **Python 3.9 or newer** (uses f-strings, walrus operator features)
- **openpyxl ≥ 3.0** — `pip install openpyxl`

No other dependencies. All other parsing uses the Python standard library.

---

## Output

A single `.xlsx` workbook with 10 worksheets:

1. **README** — methodology, definitions, key totals
2. **All Data** — complete dataset with 6 added analytical columns
3. **Analytical View** — filtered to vol ≥ filter (default 5)
4. **Outliers Only** — O:E > threshold within analytical view
5. **Systemic Outliers** — facilities outlier in 3+ episodes
6. **Top-30 Concentration** — facility-episodes ranked by outlier excess $
7. **Episode Summary** — per-episode rollup with classification + addressable
8. **By RAE Region** — outlier excess concentration by RAE catchment
9. **By DOI Urbanity** — outlier rate and excess by facility category
10. **By Hospital System** — outlier excess by system affiliation

Each filtered worksheet has a colored banner row at the top stating exactly what filter is applied so end users can see the analytical context at a glance.

---

## Troubleshooting

**"ERROR: Input file is missing required column(s)"**
Your CSV header names do not match the required spec exactly. Most common cause: extra space in header, or column renamed in source system. Open the CSV in a text editor (not Excel, which may auto-format), inspect row 1, and rename headers to match the spec exactly.

**"WARNING: Episodes not in classification config"**
A new episode appeared in the data that isn't in `config.json`. The script proceeded with the default classification. Update `config.json` to assign the episode explicitly, then re-run.

**Output XLSX file is open in Excel**
Excel locks files while open. Close the file in Excel before re-running, or use a different `--output` path.

**Worksheet 4 (Outliers Only) is empty**
No facility-episodes exceed the O:E threshold. Either the threshold is set too high for this data cycle, or risk adjustment has compressed the distribution. Try `--oe-threshold 1.05` to see the next tier of facilities.

---

## File Manifest

Folder: `PACES_Workbook_Generator/`

| File | Purpose |
|---|---|
| `build_cfo_workbook.py` | The script |
| `config.json` | Episode classifications, RAE descriptors, thresholds — edit to customize |
| `INPUT_SPECIFICATION.md` | This document |

Drop this folder anywhere; the script and config travel together.
