"""
build_cfo_html.py
─────────────────────────────────────────────────────────────────────────────
PACES CFO Decision Tool HTML Generator
Colorado Medicaid · KPMG Engagement · Episodes of Care Solutions

Takes a Colorado APCD facility-episode risk-adjusted cost CSV as input and
produces the interactive CFO Decision Tool HTML page used by HCPF Finance.

USAGE
─────────────────────────────────────────────────────────────────────────────
  python build_cfo_html.py INPUT.csv
  python build_cfo_html.py INPUT.csv --output Medicaid_CFO_Decision_Tool.html
  python build_cfo_html.py INPUT.csv --config config.json --template path/to/template.html
  python build_cfo_html.py INPUT.csv --min-volume 5 --oe-threshold 1.10
  python build_cfo_html.py --help

OUTPUT
─────────────────────────────────────────────────────────────────────────────
  A single self-contained .html file. Opens in any modern browser.
  Sections:
    1. Three Variables That Drive the Number (Standard vs Heterogeneous,
       persistence, concentration)
    2. The Levers (Referral Steering, QI Engagement)
    3. The Calculator (per-episode editable inputs, top-N selector, lever
       intensities, live recompute, sensitivity grid)
    4. Where the Team Engages (systemic outliers, top-N concentration)
    5. Where the Excess Lives (RAE Region, DOI urbanity, Hospital System)
    6. The CFO's Defensible Goal (live progress vs ceiling)
    7. Defensibility & Team Function

DEPENDENCIES
─────────────────────────────────────────────────────────────────────────────
  Python 3.9+
  paces_analytics.py    (in this folder)
  templates/cfo_decision_tool.html  (in this folder)
  config.json           (in this folder, optional)

  No external dependencies. Standard library only.

INPUT REQUIREMENTS
─────────────────────────────────────────────────────────────────────────────
  See INPUT_SPECIFICATION.md. Same input as build_cfo_workbook.py.
─────────────────────────────────────────────────────────────────────────────
"""

import argparse
import json
import sys
from pathlib import Path

# Local imports
sys.path.insert(0, str(Path(__file__).parent))
import paces_analytics as pa


# ────────────────────────────────────────────────────────────────────────────
# HELPERS
# ────────────────────────────────────────────────────────────────────────────
def fmt_dollars_m(v, places=1):
    """Format a number in millions, e.g. 25.18 -> '25.2'."""
    return f"{v / 1_000_000:.{places}f}"


def fmt_k_or_m(v, m_places=2):
    """Format as $XXK if under $1M, else $X.XXM. Returns string with sign and units."""
    if abs(v) >= 1_000_000:
        return f"{v / 1_000_000:.{m_places}f}M"
    return f"{round(v / 1_000):,}K"


def js_array_pretty(items, indent=2):
    """
    Render a list of dicts as a JS array literal (slightly nicer than json.dumps).
    Strings use double quotes; numbers and booleans pass through.
    """
    if not items:
        return "[]"
    lines = ["["]
    for i, item in enumerate(items):
        suffix = "," if i < len(items) - 1 else ""
        # Use JSON serialization for the object body, but unquote known numeric keys
        body = json.dumps(item, ensure_ascii=False)
        lines.append(" " * indent + body + suffix)
    lines.append("]")
    return "\n".join(lines)


def js_dict_int_keys(d):
    """Render a dict {int_or_string: float} as a JS object literal."""
    if not d:
        return "{}"
    pairs = []
    for k, v in sorted(d.items(), key=lambda x: int(x[0])):
        pairs.append(f"  {int(k)}: {float(v):.3f}")
    return "{\n" + ",\n".join(pairs) + "\n}"


# ────────────────────────────────────────────────────────────────────────────
# MODEL-GROUNDED SAVINGS RANGE (mirrors HTML JS computeEpisode logic)
# ────────────────────────────────────────────────────────────────────────────
def compute_scenario_savings(episodes, levers, concentration, addr_pcts):
    """Mirror the JS computeEpisode/computeAll function exactly."""
    a_total = c_total = 0
    mig_a = levers["migA"] / 100
    dl_a = levers["dlA"] / 100
    eng_c = levers["engC"] / 100
    red_c = levers["redC"] / 100
    for ep in episodes:
        # addressable %: use the episode's default (Standard 70% / Heterogeneous 50%)
        # exactly as the HTML JS does on load
        addr_pct = addr_pcts.get(ep["cls_short"], 0.5)
        base_addressable = ep["outlier_excess"] * addr_pct * concentration
        elective_factor = 1 if ep["elective"] else 0
        lever_a = ep["outlier_actual"] * concentration * mig_a * dl_a * elective_factor
        non_mig_share = 1 - mig_a * elective_factor
        lever_c = base_addressable * non_mig_share * eng_c * red_c
        a_total += lever_a
        c_total += lever_c
    return a_total + c_total


def compute_model_grounded_range(episodes, concentration, config):
    """Return dict of named scenarios → savings dollars."""
    addr_pcts_short = {
        "std": config["addressable_pct_by_classification"]["Standard"],
        "het": config["addressable_pct_by_classification"]["Heterogeneous"],
    }
    presets = {
        "floor": {"migA": 5, "dlA": 10, "engC": 30, "redC": 15},
        "balanced": {"migA": 10, "dlA": 10, "engC": 50, "redC": 20},
        "stretch": {"migA": 15, "dlA": 10, "engC": 65, "redC": 25},
    }
    return {
        "floor_top20": compute_scenario_savings(
            episodes, presets["floor"], concentration.get(20, 0), addr_pcts_short
        ),
        "balanced_top20": compute_scenario_savings(
            episodes, presets["balanced"], concentration.get(20, 0), addr_pcts_short
        ),
        "balanced_top30": compute_scenario_savings(
            episodes, presets["balanced"], concentration.get(30, 0), addr_pcts_short
        ),
        "stretch_top30": compute_scenario_savings(
            episodes, presets["stretch"], concentration.get(30, 0), addr_pcts_short
        ),
        "stretch_all": compute_scenario_savings(
            episodes, presets["stretch"], 1.0, addr_pcts_short
        ),
    }


def suggest_year2_target_range(model_range):
    """
    Suggest a defensible Year-2 CFO target range.
    Low end: Stretch at Top-30 (focused-engagement realistic stretch)
    High end: midpoint of Stretch-Top-30 and Stretch-All (achievable with upside)
    Both rounded to nearest $500K for clean public language.
    """
    low = model_range["stretch_top30"]
    high = (model_range["stretch_top30"] + model_range["stretch_all"]) / 2
    def round_to_500k(x):
        return round(x / 500_000) * 500_000
    return round_to_500k(low), round_to_500k(high)


# ────────────────────────────────────────────────────────────────────────────
# SECTION 5 HTML (Where the Excess Lives)
# ────────────────────────────────────────────────────────────────────────────
def build_section_5_html(rae, doi, systems, totals):
    """Generate the entire Section 5 block."""
    excess_m = fmt_dollars_m(totals["outlier_excess"])
    # Top systems for the cumulative figure
    top5_cum_pct = sum(s["pct_of_total"] for s in systems[:5])

    # Find Top-2 systems for narrative
    top_rae = rae[0] if rae else None
    second_rae = rae[1] if len(rae) > 1 else None

    # DOI Top-2 by outlier rate
    doi_high_rate = sorted(doi, key=lambda d: -d["outlier_rate_pct"])
    metro = next((d for d in doi if d["doi"].lower() == "metro"), None)

    # Find independent / "No System"
    independent = next(
        (s for s in systems if s["system"].lower() in ("no system", "(unclassified)")),
        None,
    )

    # RAE narrative
    rae_narrative_parts = []
    rae_color_map = {
        # Top-2 highlighted in green (concentration of opportunity)
    }
    for r in rae:
        rae_narrative_parts.append(
            f"<b>RAE Region {r['rae']}</b>: <b>${fmt_dollars_m(r['outlier_excess'], 2)}M "
            f"({r['pct_of_total']}%)</b>"
        )
    rae_narrative = " · ".join(rae_narrative_parts)

    # System narrative
    system_narrative_parts = []
    for s in systems[:5]:
        system_narrative_parts.append(
            f"<b>{s['system']}</b>: ${fmt_dollars_m(s['outlier_excess'], 2)}M "
            f"({s['pct_of_total']}%)"
        )
    system_narrative = " · ".join(system_narrative_parts)

    # Build RAE table rows
    rae_rows = []
    for i, r in enumerate(rae):
        highlight = ' style="background:#e8f5e9;"' if i == 0 else ""
        bold_pct = f"<b>{r['pct_of_total']}%</b>" if i == 0 else f"{r['pct_of_total']}%"
        rae_rows.append(
            f'<tr{highlight}><td><b>Region {r["rae"]}</b><br>'
            f'<span style="font-size:7pt;color:var(--muted);">{r["descriptor"]}</span></td>'
            f'<td class="right">${fmt_dollars_m(r["outlier_excess"], 2)}M</td>'
            f'<td class="right">{bold_pct}</td></tr>'
        )

    # Build DOI table rows
    doi_rows = []
    for d in doi:
        highlight = ""
        if d["outlier_rate_pct"] >= 50:
            highlight = ' style="background:#fce8e6;"'
        bold_rate = (
            f'<b>{d["outlier_rate_pct"]}%</b>'
            if d["outlier_rate_pct"] >= 50
            else f'{d["outlier_rate_pct"]}%'
        )
        doi_rows.append(
            f'<tr{highlight}><td><b>{d["doi"]}</b><br>'
            f'<span style="font-size:7pt;color:var(--muted);">{d["n_facilities"]} facilities</span></td>'
            f'<td class="right">${fmt_dollars_m(d["outlier_excess"], 2)}M</td>'
            f'<td class="right">{bold_rate}</td></tr>'
        )

    # Build System table rows (top 5)
    system_rows = []
    for s in systems[:5]:
        descriptor = (
            f"{s['n_facilities']} independent"
            + ("s" if s["n_facilities"] != 1 else "")
            if s["system"].lower() in ("no system", "(unclassified)")
            else f"{s['n_facilities']} facilit"
            + ("ies" if s["n_facilities"] != 1 else "y")
        )
        system_rows.append(
            f'<tr><td><b>{s["system"]}</b><br>'
            f'<span style="font-size:7pt;color:var(--muted);">{descriptor}</span></td>'
            f'<td class="right">${fmt_dollars_m(s["outlier_excess"], 2)}M</td>'
            f'<td class="right">{s["pct_of_total"]}%</td></tr>'
        )
    system_rows.append(
        f'<tr style="background:#e8f5e9;"><td colspan="2"><b>Top 5 cumulative</b></td>'
        f'<td class="right"><b>{top5_cum_pct:.1f}%</b></td></tr>'
    )

    # RAE var-card narrative
    top_pct = top_rae["pct_of_total"] if top_rae else 0
    second_pct = second_rae["pct_of_total"] if second_rae else 0
    rae_var_title = (
        f"{top_pct}% in RAE {top_rae['rae']}, {second_pct}% in RAE {second_rae['rae']}"
        if top_rae and second_rae
        else "Concentration by RAE"
    )

    # DOI var-card narrative — pick the two highest outlier-rate categories
    high_rates = [d for d in doi_high_rate if d["outlier_rate_pct"] >= 50][:2]
    if len(high_rates) >= 2:
        doi_var_title = (
            f"{high_rates[1]['doi']} & {high_rates[0]['doi']} have "
            f"{high_rates[1]['outlier_rate_pct']:.0f}–"
            f"{high_rates[0]['outlier_rate_pct']:.0f}% outlier rates"
        )
    else:
        doi_var_title = "Outlier rates vary by urbanity"

    # System var-card narrative
    system_var_title = f"Top 5 systems = {top5_cum_pct:.0f}% of excess"
    top5_sum_m = fmt_dollars_m(sum(s["outlier_excess"] for s in systems[:5]), 1)
    independent_text = ""
    if independent:
        independent_text = (
            f"<b>Independents (\"{independent['system']}\")</b>: "
            f"${fmt_dollars_m(independent['outlier_excess'], 2)}M "
            f"({independent['pct_of_total']}%) — "
        )

    # CEAC / Rural specific text
    ceac_text = ""
    metro_rate = metro["outlier_rate_pct"] if metro else 32
    rural = next((d for d in doi if d["doi"].lower() == "rural"), None)
    ceac = next((d for d in doi if d["doi"].lower() == "ceac"), None)
    rural_rate = rural["outlier_rate_pct"] if rural else 0
    ceac_rate = ceac["outlier_rate_pct"] if ceac else 0
    metro_excess_m = (
        fmt_dollars_m(metro["outlier_excess"], 1) if metro else "0"
    )

    return f"""<div class="section-label">Section 5 · Where the Excess Lives — Geographic, RAE &amp; System View</div>
<h2>The Strategic Map: Who Holds the ${excess_m}M</h2>

<p style="font-size:9pt; color:var(--muted); margin-bottom:14px;">The APCD output enables three strategic dimensions for engagement planning. Each addresses a different "where do we put our hands" question from the CFO's vantage point.</p>

<div class="vars-grid">
  <div class="var-card v1">
    <div class="var-label">By RAE Region</div>
    <div class="var-title">{rae_var_title}</div>
    <p>Outlier excess concentrates dramatically by Regional Accountable Entity catchment. {rae_narrative}. Lever A (referral steering) operates through RAE contracts — engaging RAE {top_rae['rae'] if top_rae else ''} leadership alone touches more than {round(top_pct)}% of the opportunity.</p>
  </div>
  <div class="var-card v2">
    <div class="var-label">By DOI Urbanity</div>
    <div class="var-title">{doi_var_title}</div>
    <p>Metro facilities hold the most outlier <i>dollars</i> (${metro_excess_m}M, {metro_rate}% outlier rate), but <b>CEAC</b> facilities (Critical Essential Access Centers) and <b>Rural</b> facilities have outlier rates of <b>{ceac_rate}%</b> and <b>{rural_rate}%</b> respectively — 2–3× the metro rate. These small facilities individually carry less excess but are statistically far more likely to be outliers. Engagement strategy must distinguish: large-metro outliers are likely real performance gaps; rural/CEAC outliers often reflect access tradeoffs and small-volume variance — different QI playbook.</p>
  </div>
  <div class="var-card v3">
    <div class="var-label">By Hospital System</div>
    <div class="var-title">{system_var_title}</div>
    <p>Five entities contain ${top5_sum_m}M of the ${excess_m}M outlier excess. {independent_text}{system_narrative}. Engagement at the system C-suite level for the named chains plus a coordinated approach to independents covers {top5_cum_pct:.0f}% of the recoverable opportunity in five conversations.</p>
  </div>
</div>

<h3>Geographic / System Concentration — Full Tables</h3>

<div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap:14px; margin-top:8px;">
  <div>
    <table class="conc">
      <thead><tr><th>RAE</th><th class="right">Outl $</th><th class="right">% of Total</th></tr></thead>
      <tbody>
        {chr(10).join('        ' + r for r in rae_rows)}
      </tbody>
    </table>
  </div>
  <div>
    <table class="conc">
      <thead><tr><th>DOI Category</th><th class="right">Outl $</th><th class="right">Out Rate</th></tr></thead>
      <tbody>
        {chr(10).join('        ' + r for r in doi_rows)}
      </tbody>
    </table>
  </div>
  <div>
    <table class="conc">
      <thead><tr><th>System</th><th class="right">Outl $</th><th class="right">% Total</th></tr></thead>
      <tbody>
        {chr(10).join('        ' + r for r in system_rows)}
      </tbody>
    </table>
  </div>
</div>

<div style="background:var(--bg-alt); border-left:4px solid var(--savings); padding:12px 16px; margin:14px 0; border-radius:0 3px 3px 0; font-size:9pt;">
  <b style="color:var(--savings);">Defensibility note — case mix is identical.</b> Outliers (O:E &gt; {totals.get('oe_threshold', 1.10):.2f}) and non-outliers (O:E ≤ {totals.get('oe_threshold', 1.10):.2f}) both have <b>{totals['outlier_hr_pct']:.1f}% high-risk volume</b> (outliers) and <b>{totals['non_outlier_hr_pct']:.1f}%</b> (non-outliers). The risk adjustment is doing its job — outliers are outliers on <i>cost</i>, not on patient acuity. This is the single most important defensibility point: there is no "but they take sicker patients" rejoinder available to outlier facilities under questioning.
</div>"""


# ────────────────────────────────────────────────────────────────────────────
# BUILD SUBSTITUTIONS DICT
# ────────────────────────────────────────────────────────────────────────────
def build_substitutions(result, generated_date=None):
    """Build the dict of @@MARKER@@ → value substitutions for the template."""
    config = result["config"]
    totals = result["totals"]
    episodes = result["episodes"]
    conc = result["concentration"]

    # Useful derived numbers
    total_actual = totals["actual"]
    total_excess = totals["outlier_excess"]
    total_addr = totals["addressable"]
    n_outliers = totals["n_outliers"]
    n_filtered = totals["n_filtered"]
    n_episodes = totals["n_episodes"]
    outlier_pct_of_spend = (total_excess / total_actual * 100) if total_actual else 0
    addr_pct_of_outlier = (
        (total_addr / total_excess * 100) if total_excess else 0
    )

    # Top-N percentages
    top_10 = conc.get(10, 0) * 100
    top_20 = conc.get(20, 0) * 100
    top_30 = conc.get(30, 0) * 100
    top_50 = conc.get(50, 0) * 100

    # Capture dollars at each top-N
    top_10_m = total_excess * conc.get(10, 0) / 1_000_000
    top_20_m = total_excess * conc.get(20, 0) / 1_000_000
    top_30_m = total_excess * conc.get(30, 0) / 1_000_000
    top_50_m = total_excess * conc.get(50, 0) / 1_000_000

    # Addressable percentage in config
    std_addr_pct = int(
        config["addressable_pct_by_classification"]["Standard"] * 100
    )
    het_addr_pct = int(
        config["addressable_pct_by_classification"]["Heterogeneous"] * 100
    )

    # Model-grounded scenario savings
    model_range = compute_model_grounded_range(episodes, conc, config)
    year2_low, year2_high = suggest_year2_target_range(model_range)

    # Pass oe_threshold into totals for Section 5
    totals["oe_threshold"] = config["oe_outlier_threshold"]

    # Section 5 HTML (generated as block)
    section_5 = build_section_5_html(
        result["rae"], result["doi"], result["systems"], totals
    )

    # JS data arrays
    # EPISODES — only the fields the JS calculator reads
    episodes_js = [
        {
            "name": e["name"],
            "cls": e["cls_short"],
            "actual": int(e["actual"]),
            "volume": int(e["volume"]),
            "n_facilities": int(e["n_facilities"]),
            "n_outliers": int(e["n_outliers"]),
            "outlier_excess": int(e["outlier_excess"]),
            "outlier_actual": int(e["outlier_actual"]),
            "elective": bool(e["elective"]),
            "electiveLocked": bool(e["electiveLocked"]),
        }
        for e in episodes
    ]

    # SYSTEMIC — facilities outlier in 3+ episodes
    systemic_js = [
        {
            "fac": s["facility"],
            "sys": s["system"] or "No System",
            "outlier_eps": s["outlier_eps"],
            "total_eps": s["total_eps"],
            "excess": s["excess"],
            "episodes": s["episodes"],
        }
        for s in result["systemic"]
    ]

    # TOP_OUTLIERS — top-30 facility-episodes
    top_outliers_js = [
        {
            "fac": o["facility"],
            "ep": o["episode_short"],
            "oe": round(o["oe"], 2),
            "excess": o["excess"],
        }
        for o in result["top_outliers"]
    ]

    # Concentration dict — JS keys as integer keys
    concentration_js = {int(k): round(v, 3) for k, v in conc.items()}

    # Generated date — pulled from config, file mtime, or current date
    if generated_date is None:
        generated_date = "May 2026"

    return {
        # Top-line numbers
        "@@TOTAL_ACTUAL_M@@": fmt_dollars_m(total_actual, 1),
        "@@TOTAL_OUTLIER_EXCESS_M@@": fmt_dollars_m(total_excess, 1),
        "@@TOTAL_ADDRESSABLE_M@@": fmt_dollars_m(total_addr, 1),
        "@@TOTAL_ADDRESSABLE_M_PRECISE@@": fmt_dollars_m(total_addr, 2),
        "@@TOTAL_OUTLIER_EXCESS_NUM@@": str(int(total_excess)),
        "@@OUTLIER_PCT_OF_SPEND@@": f"{outlier_pct_of_spend:.1f}",
        "@@ADDR_PCT_OF_OUTLIER@@": f"{addr_pct_of_outlier:.1f}",
        # Counts
        "@@N_EPISODES@@": str(n_episodes),
        "@@N_OUTLIERS@@": str(n_outliers),
        "@@ANALYTICAL_ROWS@@": str(n_filtered),
        "@@TOTAL_RAW_ROWS@@": str(len(result["data"])),
        # Filters
        "@@MIN_VOLUME@@": str(config["min_volume_filter"]),
        "@@OE_THRESHOLD@@": f"{config['oe_outlier_threshold']:.2f}",
        # Addressable percentages
        "@@STD_ADDR_PCT@@": str(std_addr_pct),
        "@@HET_ADDR_PCT@@": str(het_addr_pct),
        # Concentration percentages — short (integer) and precise (1 decimal)
        "@@TOP_10_PCT@@": f"{top_10:.1f}",
        "@@TOP_20_PCT@@": f"{top_20:.1f}",
        "@@TOP_30_PCT@@": f"{top_30:.1f}",
        "@@TOP_50_PCT@@": f"{top_50:.1f}",
        "@@TOP_10_PCT_INT@@": str(round(top_10)),
        "@@TOP_20_PCT_INT@@": str(round(top_20)),
        "@@TOP_30_PCT_INT@@": str(round(top_30)),
        "@@TOP_50_PCT_INT@@": str(round(top_50)),
        # Concentration dollar amounts in millions
        "@@TOP_10_M@@": fmt_dollars_m(top_10_m * 1_000_000, 2),
        "@@TOP_20_M@@": fmt_dollars_m(top_20_m * 1_000_000, 2),
        "@@TOP_30_M@@": fmt_dollars_m(top_30_m * 1_000_000, 2),
        "@@TOP_50_M@@": fmt_dollars_m(top_50_m * 1_000_000, 2),
        # Model-grounded scenario range
        "@@FLOOR_TOP20@@": fmt_k_or_m(model_range["floor_top20"]),
        "@@BALANCED_TOP20@@": fmt_k_or_m(model_range["balanced_top20"]),
        "@@BALANCED_TOP30@@": fmt_k_or_m(model_range["balanced_top30"]),
        "@@STRETCH_TOP30@@": fmt_k_or_m(model_range["stretch_top30"]),
        "@@STRETCH_ALL@@": fmt_k_or_m(model_range["stretch_all"]),
        "@@YEAR2_TARGET_LOW@@": fmt_k_or_m(year2_low, m_places=1),
        "@@YEAR2_TARGET_HIGH@@": fmt_k_or_m(year2_high, m_places=1),
        # Metadata
        "@@GENERATED_DATE@@": generated_date,
        # Section 5 — entire generated block
        "@@SECTION_5_CONTENT@@": section_5,
        # JS data arrays
        "@@EPISODES_JSON@@": js_array_pretty(episodes_js),
        "@@SYSTEMIC_JSON@@": js_array_pretty(systemic_js),
        "@@TOP_OUTLIERS_JSON@@": js_array_pretty(top_outliers_js),
        "@@CONCENTRATION_JSON@@": js_dict_int_keys(concentration_js),
    }


# ────────────────────────────────────────────────────────────────────────────
# MAIN
# ────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Generate the CFO Decision Tool HTML from a Colorado APCD CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input_csv", help="Path to input CSV file")
    parser.add_argument(
        "--output",
        "-o",
        help="Output HTML path (default: Medicaid_CFO_Episode_Team_Decision_Tool.html "
        "beside the input)",
    )
    parser.add_argument(
        "--template",
        help="Path to template HTML (default: templates/cfo_decision_tool.html beside script)",
    )
    parser.add_argument(
        "--config",
        "-c",
        help="Optional path to config.json (default: looks beside script)",
    )
    parser.add_argument(
        "--generated-date",
        help='String to embed for "Source" metadata (default: "May 2026")',
    )
    parser.add_argument(
        "--min-volume", type=int, help="Override volume filter (default in config)"
    )
    parser.add_argument(
        "--oe-threshold",
        type=float,
        help="Override O:E outlier threshold (default in config)",
    )
    args = parser.parse_args()

    input_path = Path(args.input_csv)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}", file=sys.stderr)
        sys.exit(2)

    output_path = (
        Path(args.output)
        if args.output
        else input_path.with_name(
            "Medicaid_CFO_Episode_Team_Decision_Tool.html"
        )
    )

    template_path = (
        Path(args.template)
        if args.template
        else Path(__file__).parent / "templates" / "cfo_decision_tool.html"
    )
    if not template_path.exists():
        print(f"ERROR: Template not found: {template_path}", file=sys.stderr)
        sys.exit(2)

    print(f"\nPACES CFO Decision Tool HTML Generator")
    print(f"  Input:    {input_path}")
    print(f"  Output:   {output_path}")
    print(f"  Template: {template_path}")

    # Run full analysis via shared analytics module
    cli_overrides = {
        "min_volume_filter": args.min_volume,
        "oe_outlier_threshold": args.oe_threshold,
    }
    result = pa.run_full_analysis(
        input_path, config_path=args.config, cli_overrides=cli_overrides
    )

    print(
        f"  Filter:   vol >= {result['config']['min_volume_filter']}, "
        f"O:E > {result['config']['oe_outlier_threshold']}"
    )
    print(f"  Episodes: {result['totals']['n_episodes']}")
    print(
        f"  Outlier excess in scope: ${result['totals']['outlier_excess']:,.0f} "
        f"across {result['totals']['n_outliers']} facility-episodes"
    )

    # Build substitutions
    print(f"\nBuilding substitution dictionary...")
    subs = build_substitutions(result, generated_date=args.generated_date)
    print(f"  {len(subs)} substitutions prepared")

    # Load template
    with open(template_path, encoding="utf-8") as f:
        template = f.read()

    # Substitute. Important: longer markers first to avoid partial matches.
    # @@TOTAL_OUTLIER_EXCESS_NUM@@ would otherwise be a prefix-mismatch for
    # @@TOTAL_OUTLIER_EXCESS_M@@.
    for marker in sorted(subs.keys(), key=lambda k: -len(k)):
        template = template.replace(marker, subs[marker])

    # Sanity check: did any unsubstituted markers remain?
    import re

    leftover = re.findall(r"@@([A-Z0-9_]+)@@", template)
    if leftover:
        print(
            f"\n  WARNING: {len(set(leftover))} unsubstituted markers in output:",
            file=sys.stderr,
        )
        for m in sorted(set(leftover)):
            print(f"    - @@{m}@@", file=sys.stderr)

    # Write output
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(template)

    print(f"\nDone. Saved HTML to:\n  {output_path}\n")


if __name__ == "__main__":
    main()
