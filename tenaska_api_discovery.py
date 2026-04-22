"""
Tenaska PowerTools Platform (PTP) - API Discovery
=================================================

For every endpoint under a market that your API role can see, this script
pulls:

  1. Endpoint metadata (datapoint schema + element definitions)
  2. Element list for a sample day (what entities are in scope)
  3. A one-day sample query (actual data shape)

...and writes it all to a single Excel workbook that acts as a data
dictionary. Use this to compare what the API exposes against what the
PT Portal / Data Exchange UI shows so you can plan an ingestion pipeline.

Usage
-----
    python tenaska_api_discovery.py
    python tenaska_api_discovery.py --sample-date 2026-04-20
    python tenaska_api_discovery.py --market PJM --output discovery.xlsx

Output workbook sheets
----------------------
    Index      - one row per endpoint: counts + access status
    Schemas    - every datapoint across every endpoint (flat)
    Elements   - every element (entity) across every endpoint (flat)
    <Endpoint> - one sheet per endpoint with a one-day sample

Dependencies: requests, pandas, openpyxl
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# Reuse the auth client, BASE_URL, and helpers from the sibling script.
from tenaska_generator_energy_charges import (
    BASE_URL,
    PASSWORD,
    PtpClient,
    USERNAME,
    _columnar_response_to_df,
    _encode_segment,
    _nested_response_to_df,
    list_endpoints,
)

REQUEST_DELAY_SEC = 1.1
MAX_SAMPLE_ROWS_PER_SHEET = 5000

_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


# ---------------------------------------------------------------------------
# Fetchers
# ---------------------------------------------------------------------------

def fetch_metadata(client: PtpClient, market: str, endpoint: str) -> Tuple[bool, Dict[str, Any]]:
    """GET /ptp/{market}/{endpoint} - returns (ok, body_or_error)."""
    r = client.get(f"/ptp/{_encode_segment(market)}/{_encode_segment(endpoint)}")
    if r.ok:
        try:
            return True, r.json()
        except Exception as e:
            return False, {"status": r.status_code, "error": f"json parse: {e}", "raw": r.text[:500]}
    return False, {"status": r.status_code, "error": r.text[:500]}


def fetch_elements(
    client: PtpClient, market: str, endpoint: str, begin: str
) -> Tuple[bool, Dict[str, Any]]:
    r = client.get(
        f"/ptp/{_encode_segment(market)}/{_encode_segment(endpoint)}/elements",
        params={"begin": begin},
    )
    if r.ok:
        try:
            return True, r.json()
        except Exception as e:
            return False, {"status": r.status_code, "error": f"json parse: {e}", "raw": r.text[:500]}
    return False, {"status": r.status_code, "error": r.text[:500]}


def fetch_sample(
    client: PtpClient, market: str, endpoint: str, begin: str, end: str
) -> Tuple[bool, pd.DataFrame, str]:
    """
    One-day sample. Tries /query-columnar first, falls back to /query.
    Returns (ok, dataframe, note).
    """
    m = _encode_segment(market)
    e = _encode_segment(endpoint)
    params = {"begin": begin, "end": end}

    r = client.get(f"/ptp/{m}/{e}/query-columnar", params=params)
    if r.ok:
        try:
            df = _columnar_response_to_df(r.json())
            if not df.empty:
                return True, df, "columnar"
        except Exception as ex:
            pass  # fall through to nested

    time.sleep(REQUEST_DELAY_SEC)
    r = client.get(f"/ptp/{m}/{e}/query", params=params)
    if r.ok:
        try:
            df = _nested_response_to_df(r.json())
            return True, df, "nested"
        except Exception as ex:
            return False, pd.DataFrame(), f"nested parse error: {ex}"
    return False, pd.DataFrame(), f"HTTP {r.status_code}: {r.text[:300]}"


# ---------------------------------------------------------------------------
# Flatteners
# ---------------------------------------------------------------------------

def extract_datapoints(metadata: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Pull datapoint schema entries out of an endpoint's metadata payload."""
    data = metadata.get("data") if isinstance(metadata, dict) else metadata
    if not isinstance(data, dict):
        return []

    # Try common keys.
    candidates: List[Any] = []
    for key in ("dataPoints", "datapoints", "DataPoints"):
        v = data.get(key)
        if isinstance(v, list):
            candidates = v
            break

    rows: List[Dict[str, Any]] = []
    for dp in candidates:
        if not isinstance(dp, dict):
            continue
        rows.append({
            "keyName": dp.get("keyName") or dp.get("name"),
            "dataType": dp.get("dataType") or dp.get("type"),
            "periodicity": dp.get("periodicity"),
            "objectType": dp.get("objectType") or dp.get("elementDefinition"),
            "dataPointType": dp.get("dataPointType"),
            "description": dp.get("description") or dp.get("summary"),
            "units": dp.get("units") or dp.get("unit"),
        })
    return rows


def extract_elements(elements_body: Dict[str, Any]) -> List[Dict[str, Any]]:
    data = elements_body.get("data") if isinstance(elements_body, dict) else elements_body
    if not isinstance(data, list):
        return []
    rows: List[Dict[str, Any]] = []
    for el in data:
        if not isinstance(el, dict):
            continue
        rows.append({
            "name": el.get("name") or el.get("element"),
            "identifier": el.get("identifier"),
            "elementDefinition": el.get("elementDefinition") or el.get("definition"),
            "elementDefinitionIdentifier": el.get("elementDefinitionIdentifier"),
            "parent": el.get("parent"),
            "parentIdentifier": el.get("parentElementIdentifier") or el.get("parentIdentifier"),
            "goLive": el.get("goLive") or el.get("goLiveDate"),
            "expiration": el.get("expiration") or el.get("expirationDate"),
        })
    return rows


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

def _safe_sheet_name(name: str, existing: set) -> str:
    clean = _INVALID_SHEET_CHARS.sub(" ", name).strip() or "Sheet"
    base = clean[:31]
    name_try = base
    i = 2
    while name_try in existing:
        suffix = f"_{i}"
        name_try = base[: 31 - len(suffix)] + suffix
        i += 1
    existing.add(name_try)
    return name_try


def write_workbook(
    output_path: Path,
    index_rows: List[Dict[str, Any]],
    schema_rows: List[Dict[str, Any]],
    element_rows: List[Dict[str, Any]],
    samples: Dict[str, pd.DataFrame],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    used_names: set = set()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(index_rows).to_excel(writer, sheet_name=_safe_sheet_name("Index", used_names), index=False)

        if schema_rows:
            pd.DataFrame(schema_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("Schemas", used_names), index=False
            )
        else:
            pd.DataFrame({"note": ["No schemas discovered."]}).to_excel(
                writer, sheet_name=_safe_sheet_name("Schemas", used_names), index=False
            )

        if element_rows:
            pd.DataFrame(element_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("Elements", used_names), index=False
            )
        else:
            pd.DataFrame({"note": ["No elements discovered."]}).to_excel(
                writer, sheet_name=_safe_sheet_name("Elements", used_names), index=False
            )

        for endpoint, df in samples.items():
            sheet = _safe_sheet_name(endpoint, used_names)
            if df is None or df.empty:
                pd.DataFrame({"note": [f"No sample rows returned for {endpoint}."]}).to_excel(
                    writer, sheet_name=sheet, index=False
                )
            else:
                trimmed = df.head(MAX_SAMPLE_ROWS_PER_SHEET)
                trimmed.to_excel(writer, sheet_name=sheet, index=False)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _valid_date(s: str) -> str:
    try:
        datetime.strptime(s, "%Y-%m-%d")
    except ValueError as e:
        raise argparse.ArgumentTypeError(f"Expected YYYY-MM-DD, got {s!r}") from e
    return s


def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Discover every accessible PTP API endpoint and dump schema + samples to Excel."
    )
    p.add_argument("--market", default="PJM", help="Market / root to explore (default: PJM).")
    default_sample = (datetime.utcnow() - timedelta(days=3)).strftime("%Y-%m-%d")
    p.add_argument(
        "--sample-date",
        type=_valid_date,
        default=default_sample,
        help=(
            "Date used for the one-day sample query (YYYY-MM-DD). "
            f"Default: {default_sample} (3 days ago, so full data is likely settled)."
        ),
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx path. Default: Tenaska_API_Discovery_<market>_<date>.xlsx",
    )
    p.add_argument(
        "--only",
        nargs="*",
        default=None,
        help="Optional: only explore these endpoint names (space-separated).",
    )
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    market = args.market
    sample_begin = args.sample_date
    sample_end = (
        datetime.strptime(sample_begin, "%Y-%m-%d") + timedelta(days=1)
    ).strftime("%Y-%m-%d")

    output = args.output or Path(
        f"Tenaska_API_Discovery_{market}_{sample_begin}.xlsx"
    )

    print(f"Authenticating to {BASE_URL} ...")
    client = PtpClient(base_url=BASE_URL, username=USERNAME, password=PASSWORD)
    client.authenticate()
    print("  token acquired.")

    print(f"Listing endpoints under market '{market}' ...")
    eps = list_endpoints(client, market)
    if args.only:
        wanted = set(args.only)
        eps = [e for e in eps if (e.get("name") or "") in wanted]
    if not eps:
        print("  no endpoints returned. Exiting.", file=sys.stderr)
        return 1
    print(f"  found {len(eps)} endpoints.")

    index_rows: List[Dict[str, Any]] = []
    schema_rows: List[Dict[str, Any]] = []
    element_rows: List[Dict[str, Any]] = []
    samples: Dict[str, pd.DataFrame] = {}

    for ep in eps:
        name = ep.get("name") or ep.get("identifier") or "<unknown>"
        identifier = ep.get("identifier") or ""
        print(f"\n--- {name} ---")

        row: Dict[str, Any] = {
            "endpoint": name,
            "identifier": identifier,
            "metadata_ok": False,
            "datapoint_count": 0,
            "elements_ok": False,
            "element_count": 0,
            "sample_ok": False,
            "sample_rows": 0,
            "sample_cols": 0,
            "sample_source": "",
            "error": "",
        }

        # Metadata
        ok, meta = fetch_metadata(client, market, name)
        row["metadata_ok"] = ok
        if ok:
            dps = extract_datapoints(meta)
            row["datapoint_count"] = len(dps)
            for dp in dps:
                schema_rows.append({"endpoint": name, **dp})
            print(f"  metadata: {len(dps)} datapoints")
        else:
            row["error"] = f"metadata: {meta.get('error','')[:200]}"
            print(f"  metadata FAILED: {meta}")
        time.sleep(REQUEST_DELAY_SEC)

        # Elements
        ok, elbody = fetch_elements(client, market, name, sample_begin)
        row["elements_ok"] = ok
        if ok:
            els = extract_elements(elbody)
            row["element_count"] = len(els)
            for el in els:
                element_rows.append({"endpoint": name, **el})
            print(f"  elements: {len(els)}")
        else:
            err = elbody.get("error", "") if isinstance(elbody, dict) else ""
            if not row["error"]:
                row["error"] = f"elements: {str(err)[:200]}"
            print(f"  elements FAILED: {elbody}")
        time.sleep(REQUEST_DELAY_SEC)

        # Sample query (one day)
        ok, df, note = fetch_sample(client, market, name, sample_begin, sample_end)
        row["sample_ok"] = ok
        row["sample_source"] = note
        if ok:
            row["sample_rows"] = len(df)
            row["sample_cols"] = len(df.columns)
            samples[name] = df
            print(f"  sample: {len(df):,} rows x {len(df.columns)} cols (via {note})")
        else:
            if not row["error"]:
                row["error"] = f"sample: {note[:200]}"
            samples[name] = pd.DataFrame({"error": [note]})
            print(f"  sample FAILED: {note}")
        time.sleep(REQUEST_DELAY_SEC)

        index_rows.append(row)

    print(f"\nWriting {output} ...")
    write_workbook(output, index_rows, schema_rows, element_rows, samples)
    print(f"Done. Wrote {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())