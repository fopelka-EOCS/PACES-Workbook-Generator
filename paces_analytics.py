"""
paces_analytics.py
─────────────────────────────────────────────────────────────────────────────
PACES Episode-of-Care shared analytics module.

Pure computation functions that read a Colorado APCD facility-episode CSV
and produce the data structures consumed by both:
  - build_cfo_workbook.py (generates the 10-sheet Excel workbook)
  - build_cfo_html.py     (generates the CFO Decision Tool HTML)

This module has NO file-format dependencies (no openpyxl, no Jinja2).
It uses only the Python standard library so it can be imported by
either generator script without additional installs.
─────────────────────────────────────────────────────────────────────────────
"""

import csv
import json
import sys
from collections import defaultdict
from pathlib import Path

# ────────────────────────────────────────────────────────────────────────────
# DEFAULT CONFIGURATION
# ────────────────────────────────────────────────────────────────────────────
DEFAULT_CONFIG = {
    "episode_classifications": {
        "CABG (including Cardiac catheterization)": "Standard",
        "Coronary Artery Bypass Graft (CABG)": "Standard",
        "Open heart valve surgery (including Cardiac catheterization)": "Standard",
        "Percutaneous cardiac intervention (including Cardiac catheterization)": "Standard",
        "Thyroidectomy": "Standard",
        "Esophagogastroduodenoscopy (Upper Endoscopy)": "Heterogeneous",
        "Cholecystectomy": "Heterogeneous",
        "Bariatric Surgery": "Heterogeneous",
        "Mastectomy": "Heterogeneous",
        "Colectomy": "Heterogeneous",
        "Repair Ventral Hernia": "Heterogeneous",
        "Leg Amputation": "Heterogeneous",
    },
    "addressable_pct_by_classification": {
        "Standard": 0.70,
        "Heterogeneous": 0.50,
    },
    "default_classification_for_unknown": "Heterogeneous",
    "rae_descriptors": {
        "1": "Northeast / Western Slope",
        "2": "Boulder / Larimer",
        "3": "Pueblo / Colorado Springs",
        "4": "Denver Metro / Adams / Arapahoe",
    },
    "min_volume_filter": 5,
    "oe_outlier_threshold": 1.10,
    "systemic_min_episodes": 3,
    "top_n_concentration": 30,
    # HTML-only:
    "elective_default_by_episode": {
        # If absent, defaults to True (elective). True if elective, False if not.
        "Vaginal Delivery": False,
        "C-Section": False,
        "Leg Amputation": False,
        "Fracture/Dislocation Lower Leg/Ankle/Foot": False,
    },
    "elective_locked_by_episode": {
        "Vaginal Delivery": True,
        "C-Section": True,
        "Leg Amputation": True,
        "Fracture/Dislocation Lower Leg/Ankle/Foot": True,
    },
    # Short labels used in HTML (e.g. "EGD" instead of full episode name in dense tables)
    "short_label_by_episode": {
        "Esophagogastroduodenoscopy (Upper Endoscopy)": "EGD",
        "Percutaneous cardiac intervention (including Cardiac catheterization)": "PCI (incl. Cath)",
        "CABG (including Cardiac catheterization)": "CABG (w/ Cath)",
        "Open heart valve surgery (including Cardiac catheterization)": "Open Heart Valve",
        "Coronary Artery Bypass Graft (CABG)": "CABG",
    },
}

REQUIRED_COLUMNS = [
    "Episode Name",
    "Facility Name",
    "Facility Hospital System",
    "Facility RAE Region",
    "Facility DOI Category",
    "Medicaid Volume",
    "High Risk Volume",
    "Total Actual Cost",
    "Facility O:E Cost Ratio",
]


# ────────────────────────────────────────────────────────────────────────────
# UTILITIES
# ────────────────────────────────────────────────────────────────────────────
def parse_number(s):
    """Parse a string that may contain $, comma, %, or whitespace into a float."""
    if s is None:
        return 0.0
    s = str(s).strip().replace(",", "").replace("$", "").replace("%", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def validate_schema(header_row, required_columns=None):
    """Return col_idx dict, or exit(2) listing missing columns."""
    if required_columns is None:
        required_columns = REQUIRED_COLUMNS
    col_idx = {}
    missing = []
    for col in required_columns:
        if col in header_row:
            col_idx[col] = header_row.index(col)
        else:
            missing.append(col)
    if missing:
        print("ERROR: Input file is missing required column(s):", file=sys.stderr)
        for col in missing:
            print(f"  - {col!r}", file=sys.stderr)
        print(
            "\nExpected exact case-sensitive header names. See INPUT_SPECIFICATION.md.",
            file=sys.stderr,
        )
        sys.exit(2)
    return col_idx


def load_config(config_path=None):
    """Load config.json layered over DEFAULT_CONFIG. Returns merged config dict."""
    config = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
    if config_path is None:
        config_path = Path(__file__).parent / "config.json"
    if config_path and Path(config_path).exists():
        with open(config_path, encoding="utf-8") as f:
            user_config = json.load(f)
        for k, v in user_config.items():
            if k.startswith("_"):
                continue  # skip comment fields
            if isinstance(v, dict) and isinstance(config.get(k), dict):
                config[k] = {**config[k], **v}
            else:
                config[k] = v
    return config


def read_input_csv(path, col_idx):
    """Read CSV. Returns (header_row, [list of non-blank data rows])."""
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)
    header = rows[0]
    ep_idx = col_idx["Episode Name"]
    data = [r for r in rows[1:] if len(r) > ep_idx and r[ep_idx].strip()]
    return header, data


def compute_oe_ranks_within_episode(data, col_idx):
    """Within each episode, rank facilities by O:E ascending (1 = best)."""
    by_ep = defaultdict(list)
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    ep_idx = col_idx["Episode Name"]
    for i, r in enumerate(data):
        by_ep[r[ep_idx].strip()].append((i, parse_number(r[oe_idx])))
    ranks = {}
    for _ep, items in by_ep.items():
        for rank, (orig_i, _) in enumerate(sorted(items, key=lambda x: x[1]), start=1):
            ranks[orig_i] = rank
    return ranks


def augment_row(orig_idx, row, col_idx, config, oe_ranks):
    """Add 6 computed analytical columns to a row, return augmented list."""
    vol = parse_number(row[col_idx["Medicaid Volume"]])
    hr = parse_number(row[col_idx["High Risk Volume"]])
    oe = parse_number(row[col_idx["Facility O:E Cost Ratio"]])
    actual = parse_number(row[col_idx["Total Actual Cost"]])
    ep_name = row[col_idx["Episode Name"]].strip()

    pct_hr = (hr / vol * 100) if vol > 0 else 0
    cls = config["episode_classifications"].get(
        ep_name, config["default_classification_for_unknown"]
    )
    threshold = config["oe_outlier_threshold"]
    is_outlier = oe > threshold
    excess = (actual * (1 - 1 / oe)) if (oe > 0 and is_outlier) else 0

    flags = []
    if vol == 0:
        flags.append("zero_vol")
    elif vol < 3:
        flags.append("very_low_vol")
    elif vol < config["min_volume_filter"]:
        flags.append("low_vol")
    if oe == 0:
        flags.append("no_OE")
    dq = ";".join(flags) if flags else "OK"

    return list(row) + [
        cls,
        f"{pct_hr:.1f}%",
        oe_ranks.get(orig_idx, 0),
        f"{excess:.2f}" if is_outlier else "",
        "Yes" if is_outlier else "No",
        dq,
    ]


def warn_unclassified_episodes(data, col_idx, config):
    """Print warning if input contains episodes not in classification config."""
    classified = set(config["episode_classifications"].keys())
    found = {r[col_idx["Episode Name"]].strip() for r in data}
    unknown = found - classified
    if unknown:
        default = config["default_classification_for_unknown"]
        print("\n  WARNING: Episodes not in classification config:")
        for ep in sorted(unknown):
            print(f"    - {ep!r}  -> defaulted to {default!r}")
        print(
            "  Update config.json -> 'episode_classifications' "
            "to classify these episodes explicitly.\n"
        )


# ────────────────────────────────────────────────────────────────────────────
# ANALYTICS — return data structures consumed by both generators
# ────────────────────────────────────────────────────────────────────────────
def apply_volume_filter(augmented_rows, col_idx, config):
    """Return only rows where Medicaid Volume >= min_volume_filter."""
    min_vol = config["min_volume_filter"]
    vol_idx = col_idx["Medicaid Volume"]
    return [r for r in augmented_rows if parse_number(r[vol_idx]) >= min_vol]


def compute_episode_summaries(filtered_rows, col_idx, config):
    """
    Per-episode rollup. Returns list of dicts:
      { name, cls, n_facilities, n_outliers, volume, actual,
        outlier_excess, outlier_actual, elective, electiveLocked,
        addressable }
    Sorted by outlier_excess descending.
    """
    ep_idx = col_idx["Episode Name"]
    vol_idx = col_idx["Medicaid Volume"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]

    ep_data = defaultdict(
        lambda: {
            "actual": 0,
            "vol": 0,
            "outlier_excess": 0,
            "outlier_actual": 0,
            "n_facilities": 0,
            "n_outliers": 0,
        }
    )
    for r in filtered_rows:
        ep = r[ep_idx].strip()
        vol = parse_number(r[vol_idx])
        actual = parse_number(r[actual_idx])
        oe = parse_number(r[oe_idx])
        ep_data[ep]["actual"] += actual
        ep_data[ep]["vol"] += vol
        ep_data[ep]["n_facilities"] += 1
        if oe > threshold:
            ep_data[ep]["outlier_excess"] += actual * (1 - 1 / oe)
            ep_data[ep]["outlier_actual"] += actual
            ep_data[ep]["n_outliers"] += 1

    addr_map = config["addressable_pct_by_classification"]
    elec_def = config.get("elective_default_by_episode", {})
    elec_lock = config.get("elective_locked_by_episode", {})
    result = []
    for ep, d in ep_data.items():
        cls = config["episode_classifications"].get(
            ep, config["default_classification_for_unknown"]
        )
        addr_pct = addr_map.get(cls, 0.5)
        result.append(
            {
                "name": ep,
                "classification": cls,
                "cls_short": "std" if cls == "Standard" else "het",
                "n_facilities": d["n_facilities"],
                "n_outliers": d["n_outliers"],
                "volume": int(d["vol"]),
                "actual": round(d["actual"]),
                "outlier_excess": round(d["outlier_excess"]),
                "outlier_actual": round(d["outlier_actual"]),
                "addressable": round(d["outlier_excess"] * addr_pct),
                "addressable_pct": addr_pct,
                "elective": elec_def.get(ep, True),
                "electiveLocked": elec_lock.get(ep, False),
            }
        )
    result.sort(key=lambda e: -e["outlier_excess"])
    return result


def compute_systemic_outliers(filtered_rows, col_idx, config):
    """
    Facilities outlier in N+ distinct episodes (N=systemic_min_episodes).
    Returns list of dicts sorted by (-outlier_eps, -excess):
      { facility, system, county, doi, rae,
        outlier_eps, total_eps, persistence_pct, excess, episodes (list) }
    """
    fac_idx = col_idx["Facility Name"]
    sys_idx = col_idx["Facility Hospital System"]
    ep_idx = col_idx["Episode Name"]
    rae_idx = col_idx["Facility RAE Region"]
    doi_idx = col_idx["Facility DOI Category"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    county_idx = col_idx.get("Facility County")
    threshold = config["oe_outlier_threshold"]
    min_eps = config["systemic_min_episodes"]
    short_labels = config.get("short_label_by_episode", {})

    fac_eps = defaultdict(set)
    fac_out_eps = defaultdict(set)
    fac_excess = defaultdict(float)
    fac_meta = {}
    for r in filtered_rows:
        fac = r[fac_idx].strip()
        ep = r[ep_idx].strip()
        oe = parse_number(r[oe_idx])
        fac_eps[fac].add(ep)
        if fac not in fac_meta:
            fac_meta[fac] = {
                "system": r[sys_idx].strip(),
                "county": r[county_idx].strip() if county_idx is not None else "",
                "doi": r[doi_idx].strip(),
                "rae": r[rae_idx].strip(),
            }
        if oe > threshold:
            fac_out_eps[fac].add(ep)
            fac_excess[fac] += parse_number(r[actual_idx]) * (1 - 1 / oe)

    out = []
    for fac, eps in fac_out_eps.items():
        if len(eps) >= min_eps:
            m = fac_meta[fac]
            short_eps = sorted([short_labels.get(e, e) for e in eps])
            out.append(
                {
                    "facility": fac,
                    "system": m["system"],
                    "county": m["county"],
                    "doi": m["doi"],
                    "rae": m["rae"],
                    "outlier_eps": len(eps),
                    "total_eps": len(fac_eps[fac]),
                    "persistence_pct": round(len(eps) / len(fac_eps[fac]) * 100),
                    "excess": round(fac_excess[fac]),
                    "episodes": short_eps,
                }
            )
    out.sort(key=lambda x: (-x["outlier_eps"], -x["excess"]))
    return out


def compute_top_outliers(filtered_rows, col_idx, config, n=30):
    """
    Top-N facility-episode combinations ranked by outlier excess $.
    Returns list of dicts: { facility, episode, episode_short, oe, excess }
    """
    fac_idx = col_idx["Facility Name"]
    ep_idx = col_idx["Episode Name"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]
    short_labels = config.get("short_label_by_episode", {})

    out = []
    for r in filtered_rows:
        oe = parse_number(r[oe_idx])
        if oe > threshold:
            excess = parse_number(r[actual_idx]) * (1 - 1 / oe)
            ep = r[ep_idx].strip()
            out.append(
                {
                    "facility": r[fac_idx].strip(),
                    "episode": ep,
                    "episode_short": short_labels.get(ep, ep),
                    "oe": round(oe, 2),
                    "excess": round(excess),
                }
            )
    out.sort(key=lambda x: -x["excess"])
    return out[:n]


def compute_concentration_curve(filtered_rows, col_idx, config):
    """
    Returns dict mapping {N: captured_share_of_total_excess} for the
    cumulative top-N facility-episodes. Always includes total at the high end.
    """
    fac_idx = col_idx["Facility Name"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]

    all_out = []
    for r in filtered_rows:
        oe = parse_number(r[oe_idx])
        if oe > threshold:
            excess = parse_number(r[actual_idx]) * (1 - 1 / oe)
            all_out.append(excess)
    all_out.sort(reverse=True)
    total = sum(all_out)
    if total == 0:
        return {}

    curve = {}
    for N in [10, 15, 20, 30, 50]:
        if N <= len(all_out):
            curve[N] = sum(all_out[:N]) / total
    curve[len(all_out)] = 1.0  # "all"
    return curve


def compute_rae_breakdown(filtered_rows, col_idx, config, total_excess):
    """Outlier excess rollup by RAE Region. Returns list of dicts sorted by RAE."""
    fac_idx = col_idx["Facility Name"]
    rae_idx = col_idx["Facility RAE Region"]
    vol_idx = col_idx["Medicaid Volume"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]

    rae_excess = defaultdict(float)
    rae_vol = defaultdict(float)
    rae_facs = defaultdict(set)
    rae_out = defaultdict(int)
    for r in filtered_rows:
        rae = r[rae_idx].strip()
        rae_vol[rae] += parse_number(r[vol_idx])
        rae_facs[rae].add(r[fac_idx].strip())
        oe = parse_number(r[oe_idx])
        if oe > threshold:
            rae_excess[rae] += parse_number(r[actual_idx]) * (1 - 1 / oe)
            rae_out[rae] += 1

    descriptors = config["rae_descriptors"]
    result = []
    for rae in sorted(rae_excess.keys(), key=lambda r: -rae_excess[r]):
        pct = rae_excess[rae] / total_excess * 100 if total_excess else 0
        result.append(
            {
                "rae": rae,
                "descriptor": descriptors.get(rae, ""),
                "n_facilities": len(rae_facs[rae]),
                "volume": int(rae_vol[rae]),
                "n_outliers": rae_out[rae],
                "outlier_excess": round(rae_excess[rae]),
                "pct_of_total": round(pct, 1),
            }
        )
    return result


def compute_doi_breakdown(filtered_rows, col_idx, config, total_excess):
    """Outlier rate and excess by DOI urbanity. Returns list sorted by excess desc."""
    fac_idx = col_idx["Facility Name"]
    doi_idx = col_idx["Facility DOI Category"]
    vol_idx = col_idx["Medicaid Volume"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]

    doi_excess = defaultdict(float)
    doi_vol = defaultdict(float)
    doi_facs = defaultdict(set)
    doi_out = defaultdict(int)
    doi_total = defaultdict(int)
    for r in filtered_rows:
        doi = r[doi_idx].strip()
        doi_vol[doi] += parse_number(r[vol_idx])
        doi_facs[doi].add(r[fac_idx].strip())
        doi_total[doi] += 1
        oe = parse_number(r[oe_idx])
        if oe > threshold:
            doi_excess[doi] += parse_number(r[actual_idx]) * (1 - 1 / oe)
            doi_out[doi] += 1

    result = []
    for doi in sorted(doi_excess.keys(), key=lambda d: -doi_excess[d]):
        out_rate = doi_out[doi] / doi_total[doi] * 100 if doi_total[doi] else 0
        pct = doi_excess[doi] / total_excess * 100 if total_excess else 0
        result.append(
            {
                "doi": doi,
                "n_facilities": len(doi_facs[doi]),
                "volume": int(doi_vol[doi]),
                "facility_episodes": doi_total[doi],
                "n_outliers": doi_out[doi],
                "outlier_rate_pct": round(out_rate, 1),
                "outlier_excess": round(doi_excess[doi]),
                "pct_of_total": round(pct, 1),
            }
        )
    return result


def compute_system_breakdown(filtered_rows, col_idx, config, total_excess, top_n=10):
    """Outlier excess by hospital system. Returns top-N sorted by excess desc."""
    fac_idx = col_idx["Facility Name"]
    sys_idx = col_idx["Facility Hospital System"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]

    sys_excess = defaultdict(float)
    sys_facs = defaultdict(set)
    sys_out = defaultdict(int)
    sys_total = defaultdict(int)
    for r in filtered_rows:
        s = r[sys_idx].strip() or "(unclassified)"
        sys_facs[s].add(r[fac_idx].strip())
        sys_total[s] += 1
        oe = parse_number(r[oe_idx])
        if oe > threshold:
            sys_excess[s] += parse_number(r[actual_idx]) * (1 - 1 / oe)
            sys_out[s] += 1

    result = []
    for s in sorted(sys_excess.keys(), key=lambda s: -sys_excess[s])[:top_n]:
        out_rate = sys_out[s] / sys_total[s] * 100 if sys_total[s] else 0
        pct = sys_excess[s] / total_excess * 100 if total_excess else 0
        result.append(
            {
                "system": s,
                "n_facilities": len(sys_facs[s]),
                "facility_episodes": sys_total[s],
                "n_outliers": sys_out[s],
                "outlier_rate_pct": round(out_rate, 1),
                "outlier_excess": round(sys_excess[s]),
                "pct_of_total": round(pct, 1),
            }
        )
    return result


def compute_totals(filtered_rows, col_idx, config):
    """Top-line totals across the analytical view."""
    fac_idx = col_idx["Facility Name"]
    ep_idx = col_idx["Episode Name"]
    actual_idx = col_idx["Total Actual Cost"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    threshold = config["oe_outlier_threshold"]
    addr_map = config["addressable_pct_by_classification"]
    cls_map = config["episode_classifications"]
    default_cls = config["default_classification_for_unknown"]

    total_actual = 0
    total_excess = 0
    total_addressable = 0
    outlier_hr_vol = 0
    outlier_vol = 0
    non_hr_vol = 0
    non_vol = 0
    n_outliers = 0
    for r in filtered_rows:
        actual = parse_number(r[actual_idx])
        oe = parse_number(r[oe_idx])
        vol = parse_number(r[col_idx["Medicaid Volume"]])
        hr = parse_number(r[col_idx["High Risk Volume"]])
        ep = r[ep_idx].strip()
        total_actual += actual
        if oe > threshold:
            excess = actual * (1 - 1 / oe)
            total_excess += excess
            cls = cls_map.get(ep, default_cls)
            total_addressable += excess * addr_map.get(cls, 0.5)
            outlier_hr_vol += hr
            outlier_vol += vol
            n_outliers += 1
        else:
            non_hr_vol += hr
            non_vol += vol

    return {
        "actual": total_actual,
        "outlier_excess": total_excess,
        "addressable": total_addressable,
        "n_filtered": len(filtered_rows),
        "n_outliers": n_outliers,
        "n_facilities": len({r[fac_idx].strip() for r in filtered_rows}),
        "n_episodes": len({r[ep_idx].strip() for r in filtered_rows}),
        "outlier_hr_pct": (outlier_hr_vol / outlier_vol * 100) if outlier_vol else 0,
        "non_outlier_hr_pct": (non_hr_vol / non_vol * 100) if non_vol else 0,
    }


# ────────────────────────────────────────────────────────────────────────────
# CONVENIENCE — full pipeline that returns everything a generator needs
# ────────────────────────────────────────────────────────────────────────────
def run_full_analysis(input_csv_path, config_path=None, cli_overrides=None):
    """
    End-to-end pipeline. Returns a dict containing everything needed to build
    a workbook or HTML deliverable:
      {
        config, header, data, augmented_rows, sorted_aug, filtered_aug,
        col_idx, totals, episodes, systemic, top_outliers, concentration,
        rae, doi, systems, outliers_only
      }
    """
    config = load_config(config_path)
    if cli_overrides:
        for k, v in cli_overrides.items():
            if v is not None:
                config[k] = v

    # Validate
    with open(input_csv_path, encoding="utf-8-sig", newline="") as f:
        first_line = next(csv.reader(f))
    col_idx = validate_schema(first_line)

    # Read
    header, data = read_input_csv(input_csv_path, col_idx)
    warn_unclassified_episodes(data, col_idx, config)

    # Augment
    oe_ranks = compute_oe_ranks_within_episode(data, col_idx)
    augmented = [
        augment_row(i, r, col_idx, config, oe_ranks) for i, r in enumerate(data)
    ]

    # Sort augmented: by total episode spend desc, then by O:E desc within episode
    actual_idx = col_idx["Total Actual Cost"]
    ep_idx = col_idx["Episode Name"]
    oe_idx = col_idx["Facility O:E Cost Ratio"]
    vol_idx = col_idx["Medicaid Volume"]

    ep_spend = defaultdict(float)
    for r in augmented:
        ep_spend[r[ep_idx].strip()] += parse_number(r[actual_idx])
    sorted_aug = sorted(
        augmented,
        key=lambda r: (
            -ep_spend[r[ep_idx].strip()],
            r[ep_idx],
            -parse_number(r[oe_idx]),
            -parse_number(r[vol_idx]),
        ),
    )

    # Filter
    filtered_aug = apply_volume_filter(sorted_aug, col_idx, config)

    # Outliers as a sorted list (by excess $ desc)
    outliers_only = [r for r in filtered_aug if r[-2] == "Yes"]
    outliers_only.sort(key=lambda r: -parse_number(r[-3]))

    # Rollups
    totals = compute_totals(filtered_aug, col_idx, config)
    episodes = compute_episode_summaries(filtered_aug, col_idx, config)
    systemic = compute_systemic_outliers(filtered_aug, col_idx, config)
    top_outliers = compute_top_outliers(filtered_aug, col_idx, config, n=30)
    concentration = compute_concentration_curve(filtered_aug, col_idx, config)
    rae = compute_rae_breakdown(filtered_aug, col_idx, config, totals["outlier_excess"])
    doi = compute_doi_breakdown(filtered_aug, col_idx, config, totals["outlier_excess"])
    systems = compute_system_breakdown(
        filtered_aug, col_idx, config, totals["outlier_excess"]
    )

    return {
        "config": config,
        "header": header,
        "data": data,
        "augmented_rows": augmented,
        "sorted_aug": sorted_aug,
        "filtered_aug": filtered_aug,
        "outliers_only": outliers_only,
        "col_idx": col_idx,
        "totals": totals,
        "episodes": episodes,
        "systemic": systemic,
        "top_outliers": top_outliers,
        "concentration": concentration,
        "rae": rae,
        "doi": doi,
        "systems": systems,
    }
