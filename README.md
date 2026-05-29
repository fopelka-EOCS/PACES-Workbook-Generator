# PACES CFO Workbook Generator

Standalone Python tool that converts a Colorado APCD facility-episode risk-adjusted cost CSV into a 10-worksheet Excel workbook organized for HCPF / CFO decision-making.

Built for the **EOCS · KPMG Colorado Medicaid engagement** so the analytical workbook can be regenerated on every data refresh without manual recreation.

---

## Quick start

```bash
# Install the one dependency (one-time setup)
pip install openpyxl

# Run against any APCD CSV that matches the input spec
python build_cfo_workbook.py "Your_APCD_Data.csv"
```

Output: a `.xlsx` file beside the input with the suffix `_workbook.xlsx`.

---

## What gets generated

Ten worksheets, each filtered to its analytical purpose with a banner at the top stating exactly what filter is applied:

| # | Sheet | Filter |
|---|---|---|
| 1 | **README** | n/a |
| 2 | **All Data** | None — full dataset |
| 3 | **Analytical View** | Volume ≥ 5 |
| 4 | **Outliers Only** | Vol ≥ 5 AND O:E > 1.10 |
| 5 | **Systemic Outliers** | Facilities outlier in 3+ episodes |
| 6 | **Top-30 Concentration** | Ranked by outlier excess $ |
| 7 | **Episode Summary** | Per-episode rollup with classification |
| 8 | **By RAE Region** | Outlier rollup by RAE catchment |
| 9 | **By DOI Urbanity** | Outlier rollup by facility urbanity |
| 10 | **By Hospital System** | Outlier rollup by system affiliation |

Color highlighting:
- **Red** rows / cells — outlier or systemic-outlier indicators
- **Yellow** — low-volume data quality flag
- **Green** — high-priority engagement target (top-10, RAE > 30% concentration, top-5 systems by cumulative %)

---

## Files in this folder

| File | Purpose | Edit it? |
|---|---|---|
| `build_cfo_workbook.py` | The script | No — unless you want to change worksheet structure |
| `config.json` | Episode classifications, RAE descriptors, thresholds | **Yes** — add new episodes here as they appear in future data cycles |
| `INPUT_SPECIFICATION.md` | Required and optional input columns; data quality rules | Reference only |
| `README.md` | This file | Reference only |

---

## When KPMG runs a new data cycle

1. Confirm the CSV headers match `INPUT_SPECIFICATION.md`. The script will exit with a clear error if any required column is missing.
2. If new episodes have been added to the PACES grouper, edit `config.json` and add them under `episode_classifications` with `"Standard"` or `"Heterogeneous"` as the value. The script will warn at runtime if any unclassified episode is found and will default it to Heterogeneous.
3. Run:
   ```bash
   python build_cfo_workbook.py "Q3_2026_data.csv" --output "Q3_2026_workbook.xlsx"
   ```
4. Open the resulting workbook. The README sheet shows the headline totals at the bottom; verify they're consistent with prior cycles (or expected refresh deltas).

---

## Adjusting filters

Defaults in `config.json`:

```json
{
  "min_volume_filter": 5,
  "oe_outlier_threshold": 1.10,
  "systemic_min_episodes": 3,
  "top_n_concentration": 30
}
```

Override at runtime without editing config:

```bash
python build_cfo_workbook.py data.csv --min-volume 10 --oe-threshold 1.05
```

The CLI overrides take precedence over `config.json` values.

---

## Methodology notes

The script implements **Frank's face-validity framework** for distinguishing addressable from definitional outlier excess:

- **Standard episodes** (tight procedure list, diagnosis not the driver): outliers reflect real performance gaps — 70% addressable
- **Heterogeneous episodes** (bundled procedures or diagnoses): part of the variance reflects case-mix or definitional issues — 50% addressable

Default addressability percentages are configurable in `config.json` if BPCI Advanced / CJR / state Medicaid program evidence suggests refinement.

---

## Dependencies

- **Python 3.9 or newer**
- **openpyxl ≥ 3.0** (`pip install openpyxl`)
- No other dependencies. All other parsing uses the Python standard library.

---

## Support

This tool is a one-time deliverable from the EOCS engagement. If KPMG needs to adapt it (additional worksheets, different filters, new data sources), the script is well-commented and modular — each worksheet is built by its own `build_*` function that can be modified independently.
