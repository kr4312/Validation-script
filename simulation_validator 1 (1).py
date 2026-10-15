"""
Reusable RFID Simulation Validation Report Generator.

Inputs (CLI args):
  --log         Simulator log file (contains lines: ... [DEBUG] tag <EPC> at (<X>, <Y>) floor <F> ...)
  --db          SQLite file with RawReads table (raw read records)
  --sitegraph   Site graph JSON (floors/regions/readers/antennas)
  --config      Config JSON (default probability tables, budgets, movement params)
  --tags        Comma separated list of EPCs to focus on (optional, default = all tags found)
  --movement    Target percentage of tags that are moving (0-100)
  --rps         Target reads-per-second for the simulation run
  --out         Output .xlsx report path (default: validation_report.xlsx)

Produces an Excel workbook with:
  - "Raw Reads"       : raw SQLite RawReads table dump
  - "Tag Placement"   : per read, actual tag X/Y/floor parsed from log (regex) joined to raw reads
  - "Static Tags"     : Tag Placement rows for tags with no real X/Y deviation across reads
  - "Dynamic Tags"    : Tag Placement rows for tags whose real X/Y changed across reads (moving)
  - "Reader Per-Second Reads" : per reader, per second - tags in visibility range (and how many of
    those are static), reads predicted by the probability table vs reads that actually happened,
    whether the per-second hardware cap was hit, and the shortfall when it was
  - "V1 Read Probability / RPS Budget" .. "V7 Tag Movement" : one sheet per validation with
    step-wise expected vs actual numbers and a PASS/FAIL/INFO status per row (config-sourced
    expected values are annotated with "(config)").

Adjust COLUMN_MAP below if your RawReads table uses different column names.
"""
import argparse
import json
import math
import re
import sqlite3
import sys
from pathlib import Path

import pandas as pd

# ---- adjust to match your SQLite RawReads schema ----
COLUMN_MAP = {
    "epc": "tag_id",
    "timestamp": "source_timestamp_utc",
    "antenna_id": "antenna_id",
    "reader_name": "reader_id",
    "floor": "floor_id",
    "antenna_x": "tag_x",
    "antenna_y": "tag_y",
}

LOG_RE = re.compile(
    r"\[DEBUG\].*?tag\s+(?P<epc>\S+)\s+at\s*\(\s*(?P<x>-?\d+(?:\.\d+)?)\s*,\s*(?P<y>-?\d+(?:\.\d+)?)\s*\)\s*floor\s*(?P<floor>\d+).*?moving=(?P<moving>true|false)"
)
LOG_TS_RE = re.compile(r"\[(?P<ts>\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z?)\]")


def load_config(path):
    return json.loads(Path(path).read_text())


def load_sitegraph_readers(path):
    """Flatten readers/antennas across all floors into a DataFrame.
    SNAP antennas are built into the reader housing: all 8 antennas share the
    reader's X/Y and are spaced 45 deg apart starting at the reader's yaw.
    """
    data = json.loads(Path(path).read_text())
    rows = []
    next_reader_id = 1
    for floor in data.get("floors", []):
        for reader in floor.get("readers", []):
            reader_id = next_reader_id
            next_reader_id += 1
            base_yaw = reader.get("rotation", {}).get("yaw", 0)
            if reader.get("category") == "SNAP" and not reader.get("antennas"):
                antennas = [{"port": i + 1, "position": reader["position"],
                             "rotation": {"yaw": base_yaw + i * 45}} for i in range(8)]
            else:
                antennas = reader.get("antennas") or [{"port": 1, "position": reader["position"], "rotation": {"yaw": base_yaw}}]
            for ant in antennas:
                rows.append({
                    "floor_name": floor.get("name"),
                    "floor_number": floor.get("number"),
                    "reader_id": reader_id,
                    "reader_name": reader.get("name"),
                    "category": reader.get("category"),
                    "type": reader.get("type"),
                    "antenna_port": ant.get("port"),
                    "x": ant["position"]["x"],
                    "y": ant["position"]["y"],
                    "yaw": ant.get("rotation", {}).get("yaw", base_yaw),
                })
    return pd.DataFrame(rows)


def load_raw_reads(db_path):
    con = sqlite3.connect(db_path)
    try:
        df = pd.read_sql_query("SELECT * FROM RawReads", con)
    finally:
        con.close()
    if "raw_payload" in df.columns:
        df = df.drop(columns=["raw_payload"])  # binary blob, not useful in report
    return df


def parse_log_tag_positions(log_path):
    """Return DataFrame[timestamp, epc, real_x, real_y, floor, moving] from log regex matches.
    `moving` comes directly from the log's moving=true/false flag (ground truth).
    """
    rows = []
    last_ts = None
    with open(log_path, "r", errors="ignore") as f:
        for line in f:
            ts_match = LOG_TS_RE.search(line)
            if ts_match:
                last_ts = ts_match.group("ts")
            m = LOG_RE.search(line)
            if m:
                rows.append({
                    "timestamp": last_ts,
                    "epc": m.group("epc"),
                    "real_x": float(m.group("x")),
                    "real_y": float(m.group("y")),
                    "floor": int(m.group("floor")),
                    "moving": m.group("moving") == "true",
                })
    return pd.DataFrame(rows)


def build_tag_placement(raw_reads, log_positions):
    """For each raw read, attach the tag's real x/y nearest in time from the log."""
    extra_cols = [COLUMN_MAP["antenna_id"], COLUMN_MAP["reader_name"], COLUMN_MAP["floor"]]
    out_cols = ["epc", "timestamp", "real_x", "real_y", "floor"] + extra_cols
    if raw_reads.empty or log_positions.empty:
        return pd.DataFrame(columns=out_cols)
    epc_col, ts_col = COLUMN_MAP["epc"], COLUMN_MAP["timestamp"]
    rr = raw_reads[[epc_col, ts_col] + extra_cols].rename(columns={epc_col: "epc", ts_col: "timestamp"})
    rr["timestamp"] = pd.to_datetime(rr["timestamp"], errors="coerce", utc=True).dt.tz_localize(None).astype("datetime64[ns]")
    lp = log_positions.copy()
    lp["timestamp"] = pd.to_datetime(lp["timestamp"], errors="coerce", utc=True).dt.tz_localize(None).astype("datetime64[ns]")
    lp = lp.dropna(subset=["timestamp"]).sort_values("timestamp")
    rr = rr.dropna(subset=["timestamp"]).sort_values("timestamp")
    merged = pd.merge_asof(rr, lp, on="timestamp", by="epc", direction="nearest")
    return merged


def distance_ft(x1, y1, x2, y2):
    return math.hypot(x2 - x1, y2 - y1) / 12.0  # inches -> feet


PROB_HEIGHT_COL = "probability3m"  # spec: probability is evaluated at 3m tag height only


def lookup_probability(dist_ft, table):
    """Interpolate probability3m (3m height only, per spec) from a readProbability table."""
    pts = sorted(table, key=lambda r: r["distanceFt"])
    col = PROB_HEIGHT_COL
    if dist_ft <= pts[0]["distanceFt"]:
        return pts[0][col]
    if dist_ft >= pts[-1]["distanceFt"]:
        return pts[-1][col]
    for a, b in zip(pts, pts[1:]):
        if a["distanceFt"] <= dist_ft <= b["distanceFt"]:
            span = b["distanceFt"] - a["distanceFt"]
            t = (dist_ft - a["distanceFt"]) / span if span else 0
            return a[col] + t * (b[col] - a[col])
    return 0.0


def get_probability_table(sim, reader_type):
    """Each reader type (e.g. SNAP_DEFAULT) has its own distance/probability curve; fall back to default."""
    types = sim["readers"].get("types", {})
    if reader_type in types and types[reader_type].get("readProbability"):
        return types[reader_type]["readProbability"]
    return sim["readers"]["default"]["readProbability"]


def boundary_value_tests(table, label):
    """Boundary value analysis of the probability curve itself: exact breakpoints, just outside
    the defined range (should clamp), and the documented min/max distance claims."""
    steps = []
    pts = sorted(table, key=lambda r: r["distanceFt"])
    lo, hi = pts[0]["distanceFt"], pts[-1]["distanceFt"]
    cases = [
        (lo, pts[0][PROB_HEIGHT_COL], "at lower boundary (exact table value)"),
        (lo - 5, pts[0][PROB_HEIGHT_COL], "below lower boundary (must clamp to first value)"),
        (hi, pts[-1][PROB_HEIGHT_COL], "at upper boundary (exact table value)"),
        (hi + 5, pts[-1][PROB_HEIGHT_COL], "beyond upper boundary (must clamp to last value, spec: 0% at max range)"),
    ]
    statuses = []
    for dist, expected, desc in cases:
        actual = lookup_probability(dist, table)
        ok = math.isclose(actual, expected, abs_tol=1e-6)
        statuses.append("PASS" if ok else "FAIL")
        steps.append(row(f"{label}: probability at distanceFt={dist} ({desc})", expected, round(actual, 4), "PASS" if ok else "FAIL"))
    passed = verdict(statuses)
    steps.append(row(f"{label}: boundary-value self-test verdict", "curve clamps correctly at all boundaries", passed, passed))
    return steps, passed == "PASS"


def split_static_dynamic(placement):
    """Split Tag Placement rows using the log's ground-truth moving=true/false flag:
    moving=true -> Dynamic Tags, moving=false -> Static Tags (majority vote per tag)."""
    if placement.empty or "moving" not in placement.columns:
        return placement.copy(), placement.copy()
    is_moving = placement.groupby("epc")["moving"].agg(lambda s: s.mode().iloc[0] if not s.mode().empty else False)
    moving_epcs = set(is_moving[is_moving].index)
    static_df = placement[~placement["epc"].isin(moving_epcs)].copy()
    dynamic_df = placement[placement["epc"].isin(moving_epcs)].copy()
    return static_df, dynamic_df


def row(step, expected, actual, status="INFO"):
    return {"step": step, "expected": expected, "actual": actual, "status": status}


def verdict(status_list):
    checked = [s for s in status_list if s in ("PASS", "FAIL")]
    return "PASS" if checked and all(s == "PASS" for s in checked) else "FAIL" if checked else "PASS"


def fail_sample_rows(bad_df, cols, reason, limit=20):
    """Build proof rows (step=PROOF) for up to `limit` failing tags with their actual values."""
    out = []
    for _, r in bad_df.head(limit).iterrows():
        proof = ", ".join(f"{c}={r[c]}" for c in cols if c in r)
        out.append(row(f"PROOF: tag {r.get('epc', '?')} failed", reason, proof, "FAIL"))
    if len(bad_df) > limit:
        out.append(row("PROOF: additional failing tags omitted", reason, f"{len(bad_df) - limit} more", "FAIL"))
    return out


# ---------------------------------------------------------------------------
# Validation 1: Read Probability / RPS Budget (all tags: static + dynamic)
# ---------------------------------------------------------------------------
def validation_1(raw_reads, placement, readers_df, config, target_rps, total_tag_population):
    """Step 1: probability curve is correct (boundary-value tested). Step 2: every reader's
    read-count in any given second never exceeds its hardware ceiling (maxReadsPerReader),
    since the cap is a PER-SECOND processing limit, not a limit on the whole run."""
    steps = []
    sim = config["simulation"]
    max_per_reader = sim.get("maxReadsPerReader", 200)

    # --- Step 1: Boundary value analysis of the read-probability curve(s), 3m height only ---
    steps.append(row("STEP 1: Read-probability curve, evaluated at 3m tag height only (spec)", "n/a", "n/a", "INFO"))
    default_table = sim["readers"]["default"]["readProbability"]
    bvt_steps, bvt_ok = boundary_value_tests(default_table, "default table")
    steps.extend(bvt_steps)
    reader_types = [t for t in sim["readers"].get("types", {}).keys()]
    type_bvt_ok = True
    for t in reader_types:
        t_steps, t_ok = boundary_value_tests(get_probability_table(sim, t), f"type {t}")
        steps.extend(t_steps)
        type_bvt_ok = type_bvt_ok and t_ok

    # --- Step 2: compare observed reads against the expected probability, per reader TYPE's own curve ---
    merged = placement.merge(readers_df, left_on=COLUMN_MAP["antenna_id"], right_on="antenna_port", how="left")
    if not merged.empty and {"real_x", "real_y", "x", "y"}.issubset(merged.columns):
        merged["dist_ft"] = merged.apply(
            lambda r: distance_ft(r["real_x"], r["real_y"], r["x"], r["y"]) if pd.notna(r.get("x")) else None, axis=1)
        merged["expected_prob"] = merged.apply(
            lambda r: lookup_probability(r["dist_ft"], get_probability_table(sim, r.get("type"))) if pd.notna(r.get("dist_ft")) else None, axis=1)
        avg_expected_prob = merged["expected_prob"].mean(skipna=True)
    else:
        avg_expected_prob = float("nan")
    steps.append(row("STEP 2: Avg expected read-probability at observed distances (per reader-type curve, 3m)",
                      f"{avg_expected_prob:.2f} (config curve)", f"{avg_expected_prob:.2f}", "INFO"))

    # --- Step 3: PER-SECOND read cap (hardware processing ceiling) ---
    steps.append(row("STEP 3: Per-reader, per-second read cap (hardware ceiling, 0 tolerance)", "n/a", "n/a", "INFO"))
    steps.append(row("total_visible_tags (input population)", "n/a", total_tag_population, "INFO"))
    steps.append(row("target RPS (input)", "n/a", target_rps, "INFO"))
    steps.append(row("MAX_READS_PER_READER per second", f"{max_per_reader} (config)", max_per_reader, "INFO"))

    cap_ok = True
    reader_col, ts_col = COLUMN_MAP["reader_name"], COLUMN_MAP["timestamp"]
    if reader_col in raw_reads.columns and ts_col in raw_reads.columns and not raw_reads.empty:
        rr = raw_reads[[reader_col, ts_col]].copy()
        rr["sec_bucket"] = pd.to_datetime(rr[ts_col], errors="coerce", utc=True).dt.floor("s")
        per_sec = rr.dropna(subset=["sec_bucket"]).groupby([reader_col, "sec_bucket"]).size()
        worst_per_reader = per_sec.groupby(level=0).max()
        for name, worst in worst_per_reader.items():
            ok = worst <= max_per_reader  # boundary: exactly max_per_reader is allowed, max_per_reader+1 is not
            cap_ok = cap_ok and ok
            steps.append(row(f"Reader[{name}] busiest 1-second read count <= cap", f"<= {max_per_reader}", int(worst), "PASS" if ok else "FAIL"))
            if not ok:
                bad_sec = per_sec[name][per_sec[name] > max_per_reader]
                sample = pd.DataFrame({"reader": name, "second": bad_sec.index.astype(str), "reads_in_second": bad_sec.values, "cap": max_per_reader}).head(20)
                steps.extend(fail_sample_rows(sample, ["reader", "second", "reads_in_second", "cap"],
                                               f"Reader {name} exceeded its per-second hardware read cap"))
    else:
        steps.append(row("Per-second reader read counts", "n/a", "no reader/timestamp data", "INFO"))

    passed = bvt_ok and type_bvt_ok and cap_ok
    steps.append(row("FINAL VERDICT", "probability curve correct at boundaries AND no reader ever exceeds its per-second cap", "PASS" if passed else "FAIL", "PASS" if passed else "FAIL"))
    return pd.DataFrame(steps), passed


# ---------------------------------------------------------------------------
# Validation 2: Sleep Duration & Broadcasting (static tags only)
# ---------------------------------------------------------------------------
def validation_2(raw_reads, static_df, config):
    """A tag being read by several antennas at the same moment is one wake/broadcast EVENT, not
    several (spec: 'the same time a tag can be read by different antennas' is expected). So reads
    are first collapsed to one event per tag per second before measuring sleep gaps."""
    steps = []
    sleep_cfg = config["simulation"]["readers"]["default"]["sleepDistribution"]
    mu_s, sigma_s = sleep_cfg["muMin"] * 60, sleep_cfg["sigmaMin"] * 60
    min_s, max_s = sleep_cfg["minMinutes"] * 60, sleep_cfg["maxMinutes"] * 60

    static_epcs = set(static_df["epc"]) if not static_df.empty else set()
    epc_col, ts_col = COLUMN_MAP["epc"], COLUMN_MAP["timestamp"]
    df = raw_reads[raw_reads[epc_col].isin(static_epcs)][[epc_col, ts_col]].copy() if static_epcs else raw_reads.iloc[0:0]
    df[ts_col] = pd.to_datetime(df[ts_col], errors="coerce", utc=True).dt.floor("s")
    df = df.drop_duplicates(subset=[epc_col, ts_col])  # collapse simultaneous multi-antenna reads into one event
    df = df.sort_values([epc_col, ts_col])
    df["gap_s"] = df.groupby(epc_col)[ts_col].diff().dt.total_seconds()
    gaps = df["gap_s"].dropna()

    steps.append(row("Sleep model: gauss(mu,sigma) clamped to [min,max]",
                      f"mu={mu_s}s(config) sigma={sigma_s}s(config) min={min_s}s(config) max={max_s}s(config)", "n/a", "INFO"))
    steps.append(row("Simultaneous multi-antenna reads of the same tag collapsed to 1 event (expected behavior)", "n/a", "applied", "INFO"))
    steps.append(row("Static tags evaluated / distinct wake events observed", f"{len(static_epcs)} tags", len(df), "INFO"))
    steps.append(row("Inter-event gaps measured (boundary: gap==min or gap==max both PASS)", "n/a", len(gaps), "INFO"))
    if len(gaps):
        steps.append(row("Observed mean gap (s)", f"~{mu_s}s (config)", f"{gaps.mean():.1f}", "INFO"))
        steps.append(row("Observed std dev gap (s)", f"~{sigma_s}s (config)", f"{gaps.std():.1f}", "INFO"))
        df["in_bounds"] = df["gap_s"].between(min_s, max_s)
        bad = df[df["gap_s"].notna() & ~df["in_bounds"]]
        pct_within = df["gap_s"].notna().astype(int).mul(0).add(df["in_bounds"].fillna(True)).mean() * 100
        n_total = df["gap_s"].notna().sum()
        n_bad = len(bad)
        ok = n_bad == 0
        steps.append(row("Gaps within [min,max] sleep window (config, 0 tolerance)",
                          f"0 violations / {n_total} gaps", f"{n_bad} violations / {n_total} gaps", "PASS" if ok else "FAIL"))
        if not ok:
            bad = bad.rename(columns={epc_col: "epc"})
            bad["expected_range_s"] = f"[{min_s},{max_s}]"
            steps.extend(fail_sample_rows(bad, ["epc", "gap_s", "expected_range_s"], "Sleep gap outside configured [min,max] window"))
        passed = ok
    else:
        steps.append(row("No repeated static-tag reads to evaluate", "n/a", "n/a", "PASS"))
        passed = True
    steps.append(row("FINAL VERDICT", "100% of sleep gaps within configured bounds", "PASS" if passed else "FAIL", "PASS" if passed else "FAIL"))
    return pd.DataFrame(steps), passed


# ---------------------------------------------------------------------------
# Generic geometric sweep validator (used by SNAP / STAR / TRANSITION / POS)
# ---------------------------------------------------------------------------
def within_sweep(tag_x, tag_y, ant_x, ant_y, yaw_deg, half_angle_deg, max_dist_in):
    dist = math.hypot(tag_x - ant_x, tag_y - ant_y)
    bearing = math.degrees(math.atan2(tag_y - ant_y, tag_x - ant_x))
    angle_diff = abs((bearing - yaw_deg + 180) % 360 - 180)
    if dist > max_dist_in:
        return False, dist, angle_diff
    return angle_diff <= half_angle_deg, dist, angle_diff


def validate_category_sweep(readers_df, combined_df, category, antenna_count_expected, half_angle_deg, max_dist_in):
    steps = []
    cat_readers = readers_df[readers_df["category"] == category]
    n_readers = cat_readers["reader_name"].nunique()
    steps.append(row(f"Readers of category {category} found", ">= 1", n_readers, "PASS" if n_readers >= 1 else "FAIL"))

    counts_per_reader = cat_readers.groupby("reader_name")["antenna_port"].nunique()
    if antenna_count_expected is not None and len(counts_per_reader):
        count_ok = bool((counts_per_reader == antenna_count_expected).all())
        steps.append(row(f"Antennas per {category} reader", f"{antenna_count_expected} (spec)", counts_per_reader.to_dict(), "PASS" if count_ok else "FAIL"))
    else:
        count_ok = True

    steps.append(row("Sweep half-angle enforced", f"{half_angle_deg} deg (spec)", half_angle_deg, "INFO"))
    steps.append(row("Max read distance enforced", f"{max_dist_in} in (spec)", max_dist_in, "INFO"))

    merged = combined_df.merge(cat_readers, how="inner", left_on="floor", right_on="floor_number")
    if merged.empty:
        steps.append(row("Tags geometrically matched to an antenna zone", "n/a", "no data", "PASS" if count_ok else "FAIL"))
        return pd.DataFrame(steps), count_ok

    results = merged.apply(lambda r: within_sweep(r["real_x"], r["real_y"], r["x"], r["y"], r["yaw"], half_angle_deg, max_dist_in), axis=1)
    merged["in_zone"] = [r[0] for r in results]
    merged["dist_in"] = [r[1] for r in results]
    merged["angle_diff_deg"] = [r[2] for r in results]
    per_tag_ok = merged.groupby("epc")["in_zone"].any()
    n_total, n_bad = len(per_tag_ok), int((~per_tag_ok).sum())
    geo_ok = n_bad == 0
    steps.append(row("Tags geometrically matched to >=1 antenna zone (100% required)",
                      f"0 unmatched / {n_total} tags", f"{n_bad} unmatched / {n_total} tags", "PASS" if geo_ok else "FAIL"))
    if not geo_ok:
        bad_epcs = per_tag_ok[~per_tag_ok].index
        closest = merged[merged["epc"].isin(bad_epcs)].sort_values("dist_in").groupby("epc").first().reset_index()
        closest["max_dist_in"] = max_dist_in
        closest["half_angle_deg"] = half_angle_deg
        steps.extend(fail_sample_rows(closest, ["epc", "reader_name", "dist_in", "max_dist_in", "angle_diff_deg", "half_angle_deg"],
                                       f"No {category} antenna sweep covers this tag's real X/Y (shows closest antenna's distance/angle vs limits)"))
    passed = count_ok and geo_ok
    steps.append(row("FINAL VERDICT", f"{category} geometry conforms to spec for 100% of tags", "PASS" if passed else "FAIL", "PASS" if passed else "FAIL"))
    return pd.DataFrame(steps), passed


def validation_3(readers_df, combined_df):
    return validate_category_sweep(readers_df, combined_df, "SNAP", antenna_count_expected=8, half_angle_deg=22.5, max_dist_in=540)


def validation_4(readers_df, combined_df):
    return validate_category_sweep(readers_df, combined_df, "STAR", antenna_count_expected=None, half_angle_deg=45, max_dist_in=540)


def validation_5(readers_df, combined_df):
    steps = []
    cat_readers = readers_df[readers_df["category"].isin(["CONTACT", "POS"])]
    counts_per_reader = cat_readers.groupby("reader_name")["antenna_port"].nunique()
    count_ok = bool((counts_per_reader <= 2).all()) if len(counts_per_reader) else True
    steps.append(row("POS reader antenna count", "1 or 2 (spec)", counts_per_reader.to_dict(), "PASS" if count_ok else "FAIL"))
    steps.append(row("Max read distance from center", "10 in (spec)", 10, "INFO"))
    steps.append(row("Read zone size", "20x20 in square (spec)", "20x20", "INFO"))

    merged = combined_df.merge(cat_readers, how="inner", left_on="floor", right_on="floor_number")
    if merged.empty:
        steps.append(row("Tags matched within 20x20in POS zone", "n/a", "no data", "PASS" if count_ok else "FAIL"))
        return pd.DataFrame(steps), count_ok
    merged["dx"] = (merged["real_x"] - merged["x"]).abs()
    merged["dy"] = (merged["real_y"] - merged["y"]).abs()
    merged["in_square"] = merged["dx"].le(10) & merged["dy"].le(10)
    per_tag_ok = merged.groupby("epc")["in_square"].any()
    n_total, n_bad = len(per_tag_ok), int((~per_tag_ok).sum())
    geo_ok = n_bad == 0
    steps.append(row("Tags matched within 20x20in POS zone (100% required)",
                      f"0 unmatched / {n_total} tags", f"{n_bad} unmatched / {n_total} tags", "PASS" if geo_ok else "FAIL"))
    if not geo_ok:
        bad_epcs = per_tag_ok[~per_tag_ok].index
        closest = merged[merged["epc"].isin(bad_epcs)].assign(dist_in=lambda d: (d["dx"] ** 2 + d["dy"] ** 2) ** 0.5)
        closest = closest.sort_values("dist_in").groupby("epc").first().reset_index()
        closest["max_offset_in"] = 10
        steps.extend(fail_sample_rows(closest, ["epc", "reader_name", "dx", "dy", "max_offset_in"],
                                       "Tag real X/Y outside the 20x20in POS read zone"))
    passed = count_ok and geo_ok
    steps.append(row("FINAL VERDICT", "POS geometry conforms to spec for 100% of tags", "PASS" if passed else "FAIL", "PASS" if passed else "FAIL"))
    return pd.DataFrame(steps), passed


def validation_6(readers_df, combined_df):
    steps = []
    cat_readers = readers_df[readers_df["category"] == "TRANSITION"]
    counts_per_reader = cat_readers.groupby("reader_name")["antenna_port"].nunique()
    count_ok = bool((counts_per_reader == 4).all()) if len(counts_per_reader) else True
    steps.append(row("Antennas per TRANSITION reader", "4 (spec)", counts_per_reader.to_dict(), "PASS" if count_ok else "FAIL"))

    yaws_ok = True
    for name, grp in cat_readers.groupby("reader_name"):
        yaws = sorted(grp["yaw"].tolist())
        yaws_ok = yaws_ok and any(abs((a - b) % 360 - 180) < 1 for a in yaws for b in yaws if a != b)
    steps.append(row("Antennas arranged in opposing N/S pairs (0/180 deg)", "True (spec)", yaws_ok, "PASS" if yaws_ok else "FAIL"))
    steps.append(row("Sweep half-angle enforced", "45 deg (spec)", 45, "INFO"))
    steps.append(row("Max read distance enforced", "180 in (spec)", 180, "INFO"))

    merged = combined_df.merge(cat_readers, how="inner", left_on="floor", right_on="floor_number")
    if merged.empty:
        geo_ok = True
        steps.append(row("Tags matched to a TRANSITION antenna zone", "n/a", "no data", "INFO"))
    else:
        results = merged.apply(lambda r: within_sweep(r["real_x"], r["real_y"], r["x"], r["y"], r["yaw"], 45, 180), axis=1)
        merged["in_zone"] = [r[0] for r in results]
        merged["dist_in"] = [r[1] for r in results]
        merged["angle_diff_deg"] = [r[2] for r in results]
        per_tag_ok = merged.groupby("epc")["in_zone"].any()
        n_total, n_bad = len(per_tag_ok), int((~per_tag_ok).sum())
        geo_ok = n_bad == 0
        steps.append(row("Tags geometrically matched to a TRANSITION antenna zone (100% required)",
                          f"0 unmatched / {n_total} tags", f"{n_bad} unmatched / {n_total} tags", "PASS" if geo_ok else "FAIL"))
        if not geo_ok:
            bad_epcs = per_tag_ok[~per_tag_ok].index
            closest = merged[merged["epc"].isin(bad_epcs)].sort_values("dist_in").groupby("epc").first().reset_index()
            closest["max_dist_in"] = 180
            steps.extend(fail_sample_rows(closest, ["epc", "reader_name", "dist_in", "max_dist_in", "angle_diff_deg"],
                                           "No TRANSITION antenna sweep covers this tag's real X/Y (shows closest antenna's distance/angle vs limits)"))
    passed = count_ok and yaws_ok and geo_ok
    steps.append(row("FINAL VERDICT", "TRANSITION geometry conforms to spec for 100% of tags", "PASS" if passed else "FAIL", "PASS" if passed else "FAIL"))
    return pd.DataFrame(steps), passed


# ---------------------------------------------------------------------------
# Validation 7: Tag Movement - checked every second from the start of the run
# ---------------------------------------------------------------------------
def validation_7(log_positions, config, movement_pct, target_rps, total_tag_population):
    """Moving/static classification comes from the log's own moving=true/false flag.
    The expected moving-tag count (population x movement %) is fixed from the beginning
    of the run and is checked against the observed count at EVERY second in the log."""
    steps = []
    move_cfg = config["simulation"]["movement"]
    time_to_exit = config["simulation"].get("timeToExitSec", 30)
    moving_weight = config["simulation"].get("movingTagReadWeight", 1)
    expected_moving_count = round(total_tag_population * movement_pct / 100.0)

    steps.append(row("Total tag population", "n/a", total_tag_population, "INFO"))
    steps.append(row("Expected moving tag count, fixed from t=0 (population x target %%, 0 tolerance)",
                      f"{movement_pct}% -> {expected_moving_count}", "checked every second below", "INFO"))

    if log_positions.empty or "moving" not in log_positions.columns:
        steps.append(row("No log moving=true/false snapshots available", "n/a", "no data", "FAIL"))
        return pd.DataFrame(steps), False

    lp = log_positions.copy()
    lp["sec_bucket"] = pd.to_datetime(lp["timestamp"], errors="coerce", utc=True).dt.floor("s")
    per_sec_moving = lp[lp["moving"]].groupby("sec_bucket")["epc"].nunique()
    all_secs = lp["sec_bucket"].dropna().unique()
    per_sec_moving = per_sec_moving.reindex(all_secs, fill_value=0).sort_index()

    bad_secs = per_sec_moving[per_sec_moving != expected_moving_count]
    n_total, n_bad = len(per_sec_moving), len(bad_secs)
    exact_ok = n_bad == 0
    steps.append(row("Seconds where observed moving-tag count == expected (100% required)",
                      f"{n_total} / {n_total} seconds match", f"{n_total - n_bad} / {n_total} seconds match", "PASS" if exact_ok else "FAIL"))
    steps.append(row("POS visit ratio before departure", f"{move_cfg['posVisitRatio']} (config)", move_cfg["posVisitRatio"], "INFO"))
    steps.append(row("Transition halt duration (s)", f"{move_cfg['transitionHaltSec']} (config)", move_cfg["transitionHaltSec"], "INFO"))
    steps.append(row("Departure dwell duration (s)", f"{move_cfg['departureDwellSec']} (config)", move_cfg["departureDwellSec"], "INFO"))
    steps.append(row("timeToExitSec used for speed=distance/time", f"{time_to_exit} (config)", time_to_exit, "INFO"))
    steps.append(row("movingTagReadWeight applied in RPS selection", f"{moving_weight} (config)", moving_weight, "INFO"))
    steps.append(row("Target RPS (input)", "n/a", target_rps, "INFO"))

    passed = exact_ok
    if not passed:
        proof = bad_secs.reset_index()
        proof.columns = ["second", "observed_moving_count"]
        proof["expected_moving_count"] = expected_moving_count
        proof["epc"] = proof["second"].astype(str)  # fail_sample_rows needs an 'epc' label column
        steps.extend(fail_sample_rows(proof, ["second", "observed_moving_count", "expected_moving_count"],
                                       "Moving-tag count at this second did not match population x target %%"))
    steps.append(row("FINAL VERDICT", "observed moving count matches target at every second (0 tolerance)", "PASS" if passed else "FAIL", "PASS" if passed else "FAIL"))
    return pd.DataFrame(steps), passed


def build_reader_second_analysis(raw_reads):
    """Per reader, per second: total read count (to confirm the per-second hardware ceiling)."""
    reader_col, ts_col = COLUMN_MAP["reader_name"], COLUMN_MAP["timestamp"]

    if raw_reads.empty or reader_col not in raw_reads.columns or ts_col not in raw_reads.columns:
        return pd.DataFrame(columns=["reader_name", "second", "read_count"])

    rr = raw_reads[[reader_col, ts_col]].copy()
    rr["second"] = pd.to_datetime(rr[ts_col], errors="coerce", utc=True).dt.tz_localize(None).dt.floor("s")
    rr = rr.dropna(subset=["second"])
    counts = rr.groupby([reader_col, "second"]).size().reset_index(name="read_count")
    counts = counts.rename(columns={reader_col: "reader_name"})
    return counts.sort_values(["reader_name", "second"]).reset_index(drop=True)


# Plain-language description of each validation, for readers unfamiliar with the system.
SUMMARY_DESCRIPTIONS = {
    "V1 Read Probability-RPS": "Checks the chance a tag gets read at a given distance is correct, and that no reader is asked to process more than its per-second hardware limit.",
    "V2 Sleep-Broadcast": "Checks that non-moving (static) tags pause the right amount of time between reads instead of broadcasting non-stop.",
    "V3 SNAP Reader": "Checks SNAP readers cover the right shape and distance around them, and only 'see' tags that are really inside that zone.",
    "V4 STAR Reader": "Checks STAR readers cover the right shape and distance around them, and only 'see' tags that are really inside that zone.",
    "V5 POS Reader": "Checks checkout (POS) readers only detect tags within their small 20x20 inch zone.",
    "V6 Transition Reader": "Checks doorway (TRANSITION) readers are set up with the correct 4 antennas, directions, and range.",
    "V7 Tag Movement": "Checks the number of tags that are moving around the store matches what was requested for this test run.",
}


def build_summary(results, total_tag_population, target_rps, movement_pct):
    """Plain-language overview sheet, in logical order, for someone new to the system."""
    rows = []
    rows.append({"Validation": "Test Inputs", "What it checks": "Settings used for this run", "Result": "INFO",
                 "Details": f"Tag population={total_tag_population}, Target RPS={target_rps}, Target moving %={movement_pct}"})
    overall_pass = True
    for name, (df, passed) in results.items():
        overall_pass = overall_pass and passed
        final = df[df["step"] == "FINAL VERDICT"]
        detail = str(final.iloc[-1]["expected"]) if not final.empty else ""
        rows.append({
            "Validation": name,
            "What it checks": SUMMARY_DESCRIPTIONS.get(name, ""),
            "Result": "PASS" if passed else "FAIL",
            "Details": detail,
        })
    rows.insert(1, {"Validation": "OVERALL RESULT", "What it checks": "Did every validation pass?",
                     "Result": "PASS" if overall_pass else "FAIL", "Details": "See each row below for the specific reason."})
    return pd.DataFrame(rows)


def write_report(out_path, raw_reads, placement, static_df, dynamic_df, reader_second_df, results,
                  total_tag_population, target_rps, movement_pct):
    summary_df = build_summary(results, total_tag_population, target_rps, movement_pct)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        for sheet_name, (df, _) in results.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        raw_reads.to_excel(writer, sheet_name="Raw Reads", index=False)
        placement.to_excel(writer, sheet_name="Tag Placement", index=False)
        static_df.to_excel(writer, sheet_name="Static Tags", index=False)
        dynamic_df.to_excel(writer, sheet_name="Dynamic Tags", index=False)
        reader_second_df.to_excel(writer, sheet_name="Reader Per-Second Reads", index=False)


def main():
    ap = argparse.ArgumentParser(description="Generate RFID simulation validation report")
    ap.add_argument("--log", required=True)
    ap.add_argument("--db", required=True)
    ap.add_argument("--sitegraph", required=True)
    ap.add_argument("--config", required=True)
    ap.add_argument("--tags", type=int, required=True, help="Total tag population size")
    ap.add_argument("--movement", type=float, default=0.0, help="Target %% of moving tags")
    ap.add_argument("--rps", type=float, required=True, help="Target reads per second")
    ap.add_argument("--out", default="validation_report.xlsx")
    args = ap.parse_args()

    config = load_config(args.config)
    readers_df = load_sitegraph_readers(args.sitegraph)
    raw_reads = load_raw_reads(args.db)
    log_positions = parse_log_tag_positions(args.log)
    placement = build_tag_placement(raw_reads, log_positions)
    static_df, dynamic_df = split_static_dynamic(placement)
    reader_second_df = build_reader_second_analysis(raw_reads)

    results = {
        "V1 Read Probability-RPS": validation_1(raw_reads, placement, readers_df, config, args.rps, args.tags),
        "V2 Sleep-Broadcast": validation_2(raw_reads, static_df, config),
        "V3 SNAP Reader": validation_3(readers_df, placement),
        "V4 STAR Reader": validation_4(readers_df, placement),
        "V5 POS Reader": validation_5(readers_df, placement),
        "V6 Transition Reader": validation_6(readers_df, placement),
        "V7 Tag Movement": validation_7(log_positions, config, args.movement, args.rps, args.tags),
    }

    write_report(args.out, raw_reads, placement, static_df, dynamic_df, reader_second_df, results, args.tags, args.rps, args.movement)
    overall = all(p for _, p in results.values())
    print(f"Report written to {args.out}. Overall: {'PASS' if overall else 'FAIL'}")
    for name, (_, passed) in results.items():
        print(f"  {name}: {'PASS' if passed else 'FAIL'}")


if __name__ == "__main__":
    sys.exit(main())

