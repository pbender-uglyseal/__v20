"""
Mileage Ratio reconciliation probe
==================================

Purpose: determine whether the UI column "Mileage Ratio" in the Generator
Energy Charge Details - 5 minute export is the same metric as
MarketRegulationResults.RegMileage from the PTP API.

Approach:
  1. Pull UI Mileage Ratio for one generator on one flowday (5-min ET).
  2. Pull MarketRegulationResults.RegMileage for the same UTC window (30-min).
  3. Roll the UI 5-min values up to 30-min ET buckets and align with API intervals.
  4. Print a side-by-side with ratio (API / UI_mean) so any consistent
     scaling factor is obvious.

Usage
-----
    python mileage_ratio_probe.py
    python mileage_ratio_probe.py --date 2026-04-12 --element "V20 Wantage"

Dependencies: requests, pandas, openpyxl (plus tzdata on Windows if zoneinfo
is missing America/New_York).
"""

import argparse
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional

import pandas as pd

try:
    from zoneinfo import ZoneInfo
except ImportError:  # Python < 3.9 fallback
    from backports.zoneinfo import ZoneInfo  # type: ignore

from tenaska_generator_energy_charges import (
    BASE_URL,
    PASSWORD,
    PtpClient,
    USERNAME,
    _encode_segment,
)

UI_FILE_DEFAULT = Path(
    r"C:\Users\PaulBender\Claude_LOCAL\V20_Energy\V20_BatteryMonitoring\tenaska_data"
    "\\Generator Energy Charge Details - 5 minute_2026-04-15_2026-04-16.xlsx"
)

ET = ZoneInfo("America/New_York")
UTC = timezone.utc


# ---------------------------------------------------------------------------
# UI side
# ---------------------------------------------------------------------------

def load_ui_window(path: Path, date_str: str, element: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Data_Generator")
    flow = pd.to_datetime(df["Flowday (Eastern)"]).dt.strftime("%Y-%m-%d")
    mask = (df["Element"] == element) & (flow == date_str)
    sub = df.loc[mask, ["Flowday (Eastern)", "Interval", "Mileage Ratio"]].copy()

    # Build a real ET-aware datetime from Flowday + Interval.
    interval_str = sub["Interval"].astype(str)
    combined = sub["Flowday (Eastern)"].astype(str) + " " + interval_str
    ts = pd.to_datetime(combined, format="%Y-%m-%d %H:%M", errors="coerce")
    sub["ts_et"] = ts.dt.tz_localize(ET, nonexistent="shift_forward",
                                     ambiguous="NaT")
    sub["ts_utc"] = sub["ts_et"].dt.tz_convert(UTC)

    # 30-min bucket start (ET) — floor to 00 or 30.
    sub["bucket_et"] = sub["ts_et"].dt.floor("30min")
    return sub.sort_values("ts_et").reset_index(drop=True)


def bucket_ui_to_30min(ui: pd.DataFrame) -> pd.DataFrame:
    g = ui.groupby("bucket_et")["Mileage Ratio"]
    return pd.DataFrame({
        "bucket_et": g.mean().index,
        "ui_n": g.size().values,
        "ui_min": g.min().values,
        "ui_mean": g.mean().values,
        "ui_max": g.max().values,
    })


# ---------------------------------------------------------------------------
# API side
# ---------------------------------------------------------------------------

def fetch_regmileage(client: PtpClient, market: str,
                     begin_iso: str, end_iso: str) -> pd.DataFrame:
    """Pull RegMileage timeseries via the nested /query endpoint."""
    path = f"/ptp/{_encode_segment(market)}/MarketRegulationResults/query"
    resp = client.get(path, params={
        "begin": begin_iso,
        "end": end_iso,
        "dataPoints": "RegMileage",
    })
    if not resp.ok:
        raise RuntimeError(f"/query failed: HTTP {resp.status_code} - {resp.text[:400]}")
    body = resp.json()

    rows = []
    for el in body.get("data", []) or []:
        for dp in el.get("dataPoints", []) or []:
            if dp.get("keyName") != "RegMileage":
                continue
            for v in dp.get("values", []) or []:
                val = _first_value(v.get("data"))
                rows.append({
                    "intervalStartUtc": v.get("intervalStartUtc"),
                    "intervalEndUtc": v.get("intervalEndUtc"),
                    "RegMileage": val,
                })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["intervalStartUtc"] = pd.to_datetime(df["intervalStartUtc"], utc=True)
    df["bucket_et"] = df["intervalStartUtc"].dt.tz_convert(ET).dt.floor("30min")
    return df.sort_values("intervalStartUtc").reset_index(drop=True)


def _first_value(data_list) -> Optional[float]:
    if not isinstance(data_list, list) or not data_list:
        return None
    first = data_list[0]
    if isinstance(first, dict):
        return first.get("value", first.get("data"))
    return first


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[2])
    ap.add_argument("--date", default="2026-04-12",
                    help="Flowday (Eastern) to compare, YYYY-MM-DD. Default: 2026-04-12")
    ap.add_argument("--element", default="V20 Wantage",
                    help='UI "Element" value to filter. Default: "V20 Wantage"')
    ap.add_argument("--market", default="PJM")
    ap.add_argument("--ui-file", type=Path, default=UI_FILE_DEFAULT,
                    help=f"Path to the UI export xlsx. Default: {UI_FILE_DEFAULT}")
    return ap.parse_args()


def main() -> int:
    args = parse_args()

    # Convert the ET flowday to UTC begin/end for the API query.
    start_et = datetime.strptime(args.date, "%Y-%m-%d").replace(tzinfo=ET)
    end_et = start_et + timedelta(days=1)
    begin_iso = start_et.astimezone(UTC).strftime("%Y-%m-%dT%H:%MZ")
    end_iso = end_et.astimezone(UTC).strftime("%Y-%m-%dT%H:%MZ")

    print(f"UI file:   {args.ui_file.name}")
    print(f"Flowday:   {args.date} ({args.element})")
    print(f"UTC range: {begin_iso} -> {end_iso}")
    print()

    # UI
    ui = load_ui_window(args.ui_file, args.date, args.element)
    if ui.empty:
        print(f"!! No UI rows for element={args.element!r} on {args.date}. "
              f"Check the Element column for exact spelling.")
        elems = pd.read_excel(args.ui_file, sheet_name="Data_Generator")["Element"].unique()
        print(f"   Elements in file: {list(elems)}")
        return 1

    ui_bucketed = bucket_ui_to_30min(ui)
    print(f"UI: {len(ui)} 5-min rows -> {len(ui_bucketed)} 30-min buckets")
    print(f"UI Mileage Ratio range across day: "
          f"min={ui['Mileage Ratio'].min():.4f}, "
          f"mean={ui['Mileage Ratio'].mean():.4f}, "
          f"max={ui['Mileage Ratio'].max():.4f}")
    print()

    # API
    print(f"Authenticating to {BASE_URL} ...")
    client = PtpClient(BASE_URL, USERNAME, PASSWORD)
    client.authenticate()
    print("Fetching MarketRegulationResults.RegMileage ...")
    api = fetch_regmileage(client, args.market, begin_iso, end_iso)
    if api.empty:
        print("!! No RegMileage rows returned from API for this window.")
        return 2

    print(f"API: {len(api)} 30-min intervals")
    print(f"API RegMileage range: "
          f"min={api['RegMileage'].min():.4f}, "
          f"mean={api['RegMileage'].mean():.4f}, "
          f"max={api['RegMileage'].max():.4f}")
    print()

    # Align and print side-by-side
    merged = api.merge(ui_bucketed, on="bucket_et", how="outer").sort_values("bucket_et")
    merged["ratio_api_over_ui_mean"] = merged["RegMileage"] / merged["ui_mean"]

    display = merged[[
        "bucket_et", "RegMileage",
        "ui_n", "ui_min", "ui_mean", "ui_max",
        "ratio_api_over_ui_mean",
    ]].copy()
    display["bucket_et"] = display["bucket_et"].dt.strftime("%Y-%m-%d %H:%M ET")

    print("Side-by-side (30-min buckets):")
    print(display.to_string(index=False, float_format=lambda v: f"{v:0.4f}"))
    print()

    # Verdict heuristic
    valid = merged.dropna(subset=["RegMileage", "ui_mean"])
    if valid.empty:
        print("VERDICT: cannot compare — no overlapping intervals with values in both sources.")
        return 0
    ratios = valid["ratio_api_over_ui_mean"].dropna()
    if ratios.std() < 0.05 * abs(ratios.mean() or 1.0):
        print(f"VERDICT: likely same metric scaled by ~{ratios.mean():.3f} "
              f"(std={ratios.std():.4f} across {len(ratios)} buckets).")
    else:
        print(f"VERDICT: likely DIFFERENT metrics. Ratio API/UI varies widely "
              f"(mean={ratios.mean():.3f}, std={ratios.std():.4f}).")
        print("Interpretation: UI 'Mileage Ratio' is probably not MarketRegulationResults.RegMileage.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
