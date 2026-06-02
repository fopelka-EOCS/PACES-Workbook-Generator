# PACES CFO Workbook & Decision Tool Generator

Two standalone Python tools that convert a Colorado APCD facility-episode risk-adjusted cost CSV into:

1. **A 10-worksheet Excel workbook** organized for HCPF / CFO analytical review (`build_cfo_workbook.py`)
2. **An interactive CFO Decision Tool HTML page** for HCPF Finance leadership (`build_cfo_html.py`)

Both share the same analytics module (`paces_analytics.py`) and the same input CSV format, so a single data refresh produces both deliverables.

Built for the **EOCS · KPMG Colorado Medicaid engagement**.

---

## Quick start

```bash
# One-time setup
pip install -r requirements.txt

# Generate the 10-sheet Excel workbook
python build_cfo_workbook.py "Your_APCD_Data.csv"

# Generate the interactive CFO Decision Tool HTML
python build_cfo_html.py "Your_APCD_Data.csv"
```

Outputs land beside the input file by default. Both scripts accept `--output` to redirect.

---

## What gets generated

### `build_cfo_workbook.py` — Excel workbook (10 worksheets)

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

### `build_cfo_html.py` — Interactive HTML Decision Tool

A single self-contained `.html` file that opens in any modern browser. Sections:

1. **Three Variables That Drive the Number** — Standard vs Heterogeneous classification, facility outlier persistence, top-N concentration
2. **The Levers** — Informed Referral Steering (RAE + PCP) and Outlier-Targeted QI Engagement
3. **The Calculator** — Per-episode editable inputs, lever intensity sliders, top-N selector, live recompute, sensitivity grid
4. **Where the Team Engages** — Systemic outlier facilities table + top-30 concentration table
5. **Where the Excess Lives** — RAE Region / DOI Urbanity / Hospital System concentration tables + case-mix defensibility note
6. **The CFO's Defensible Goal** — Live progress bar vs $25M ceiling, model-grounded savings range, suggested Year-2 target
7. **Defensibility & Team Function** — Why the methodology holds up

The HTML is fully self-contained (CSS and JavaScript inlined) — no external server required, no JavaScript libraries, no installation. Open the file in Chrome, Edge, Safari, or Firefox.

---

## Files in this folder

| File | Purpose | Edit it? |
|---|---|---|
| `paces_analytics.py` | Shared computation module — imported by both generators | No — unless adding new analytics |
| `build_cfo_workbook.py` | Excel workbook generator | No — modify worksheets if needed |
| `build_cfo_html.py` | HTML Decision Tool generator | No — modify HTML output if needed |
| `templates/cfo_decision_tool.html` | HTML template with `@@MARKER@@` placeholders | Yes — to change visual layout / copy |
| `config.json` | Episode classifications, RAE descriptors, thresholds | **Yes** — add new episodes here |
| `INPUT_SPECIFICATION.md` | Required and optional input columns; data quality rules | Reference only |
| `requirements.txt` | Python dependencies | Reference only |
| `LICENSE` | MIT | Reference only |
| `README.md` | This file | Reference only |

---

## When KPMG runs a new data cycle

1. Confirm the CSV headers match `INPUT_SPECIFICATION.md`. Both scripts exit with a clear error if any required column is missing.
2. If new episodes have been added to the PACES grouper, edit `config.json` and add them under `episode_classifications` with `"Standard"` or `"Heterogeneous"` as the value. Both scripts warn at runtime if any unclassified episode is found and default it to Heterogeneous.
3. Run both generators:
   ```bash
   python build_cfo_workbook.py "Q3_2026_data.csv"
   python build_cfo_html.py "Q3_2026_data.csv"
   ```
4. Open the resulting workbook and HTML. Verify the README sheet's headline totals are consistent with expected refresh deltas.

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

Override at runtime without editing the config:

```bash
python build_cfo_workbook.py data.csv --min-volume 10 --oe-threshold 1.05
python build_cfo_html.py data.csv --min-volume 10 --oe-threshold 1.05
```

The CLI overrides take precedence over `config.json` values.

---

## Modifying the HTML output

If KPMG wants to adjust the visual layout, copy changes, or styling of the Decision Tool HTML:

1. Open `templates/cfo_decision_tool.html` in any editor
2. Markers in the template use the format `@@MARKER_NAME@@` — these are substituted at generation time
3. The full list of available markers and what they contain is documented inline in `build_cfo_html.py` (search for `build_substitutions`)
4. CSS and JavaScript inside the template can be edited freely; they don't use markers

Common modifications:
- Color scheme: edit the `:root` CSS variables at the top of the template
- Section ordering: rearrange the section blocks; the `@@SECTION_5_CONTENT@@` marker generates the entire Where-the-Excess-Lives block from Python
- Default goal target: edit the `state.goalTarget` value in the embedded JavaScript

---

## Methodology notes

Both generators implement **Frank's face-validity framework** for distinguishing addressable from definitional outlier excess:

- **Standard episodes** (tight procedure list, diagnosis not the driver): outliers reflect real performance gaps — 70% addressable
- **Heterogeneous episodes** (bundled procedures or diagnoses): part of the variance reflects case-mix or definitional issues — 50% addressable

Default addressability percentages are configurable in `config.json` if BPCI Advanced / CJR / state Medicaid program evidence suggests refinement.

---

## Dependencies

- **Python 3.9 or newer**
- **openpyxl ≥ 3.0** (`pip install openpyxl`) — only required for `build_cfo_workbook.py`
- **No dependencies** beyond stdlib for `build_cfo_html.py`

`requirements.txt` includes openpyxl for the workbook generator. If you only need the HTML generator, no install step is needed beyond Python itself.

---

## Support

These tools are deliverables from the EOCS engagement. The codebase is well-commented and modular — each generator imports from `paces_analytics.py`, so adding new dimensions or modifying calculations happens in one place and both deliverables update automatically.
