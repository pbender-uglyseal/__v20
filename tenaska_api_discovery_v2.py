"""
Tenaska PowerTools Platform (PTP) - API Discovery v2 (with request logging)
===========================================================================

Same crawl as tenaska_api_discovery.py: for every endpoint under a market the
caller's API role can see, this pulls (1) endpoint metadata, (2) the element
list for a sample day, and (3) a one-day sample query, and writes the
results to an Excel workbook.

What this version adds
----------------------
Two new sheets in the same output workbook capture the actual HTTP requests
the script issues, so anyone reading the workbook can reproduce every call
by hand or in code:

    Requests_Logical
        One row per call. Human-readable view of *what was asked for*:
        endpoint name + identifier, phase (authenticate / list-endpoints /
        metadata / elements / sample-columnar / sample-nested-fallback),
        market, begin, end, plus the response status / size / duration.

    Requests_HTTP
        One row per call. Full HTTP envelope of *what was sent on the wire*:
        method, full URL, query params (JSON), headers (Authorization
        elided), request body, response status, response size in bytes,
        duration in milliseconds.

Use Requests_Logical for the dev review pass; use Requests_HTTP for the
implementation team to replicate calls verbatim.

Usage
-----
    python tenaska_api_discovery_v2.py
    python tenaska_api_discovery_v2.py --sample-date 2026-04-20
    python tenaska_api_discovery_v2.py --market PJM --output discovery.xlsx
    python tenaska_api_discovery_v2.py --only Market_Price DART-Basis

Output workbook sheets
----------------------
    Index             - one row per endpoint: counts + access status
    Schemas           - every datapoint across every endpoint (flat)
    Elements          - every element across every endpoint (flat)
    Requests_Logical  - one row per HTTP call, business view (NEW)
    Requests_HTTP     - one row per HTTP call, wire-level view (NEW)
    <Endpoint>        - one sheet per endpoint with a one-day sample

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
import requests

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
REQUEST_TIMEOUT = 120
MAX_SAMPLE_ROWS_PER_SHEET = 5000

_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")

# Headers that should never appear in the request log in the clear.
_SENSITIVE_HEADER_NAMES = {"authorization", "x-api-key", "cookie"}


# ---------------------------------------------------------------------------
# Logging client
# ---------------------------------------------------------------------------

def _redact_headers(headers: Dict[str, str]) -> Dict[str, str]:
    """Return a copy of `headers` with sensitive values masked."""
    return {
        k: ("<REDACTED>" if k.lower() in _SENSITIVE_HEADER_NAMES else v)
        for k, v in headers.items()
    }


class LoggingPtpClient(PtpClient):
    """
    Drop-in replacement for PtpClient that records every HTTP call it makes.

    Each entry in `request_log` is a dict with both the logical context
    (endpoint, phase, market, begin, end) set via `set_phase()` and the
    wire-level details (method, url, params, headers, body, status, size,
    duration) captured at request time.

    `set_phase()` is called by the discovery code before each call so the
    log entry carries the human-readable "what were we trying to do?".
    """

    def __post_init__(self) -> None:
        # PtpClient is a @dataclass; preserve its __post_init__ behavior.
        super().__post_init__()
        self.request_log: List[Dict[str, Any]] = []
        self._current_phase: Dict[str, Any] = {
            "endpoint_name": "",
            "endpoint_identifier": "",
            "phase": "",
            "market": "",
            "begin": None,
            "end": None,
        }

    def set_phase(
        self,
        endpoint_name: str,
        endpoint_identifier: str,
        phase: str,
        market: str,
        begin: Optional[str] = None,
        end: Optional[str] = None,
    ) -> None:
        """Tag the next HTTP call with logical context for the log."""
        self._current_phase = {
            "endpoint_name": endpoint_name,
            "endpoint_identifier": endpoint_identifier,
            "phase": phase,
            "market": market,
            "begin": begin,
            "end": end,
        }

    # ------------------------------------------------------------------
    # Authentication is captured manually because the parent's
    # authenticate() bypasses self.get() and goes straight to the session.
    # ------------------------------------------------------------------
    def authenticate(self) -> str:
        import base64

        creds = f"{self.username}:{self.password}".encode("utf-8")
        basic = base64.b64encode(creds).decode("ascii")
        headers = {"Authorization": f"Basic {basic}"}
        url = f"{self.base_url}/authentication/token"

        t0 = time.perf_counter()
        resp = self._session.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        self.request_log.append({
            **{
                "endpoint_name": "(auth)",
                "endpoint_identifier": "",
                "phase": "authenticate",
                "market": "",
                "begin": None,
                "end": None,
            },
            "method": "GET",
            "url": url,
            "query_params": "",
            "headers": json.dumps(_redact_headers(headers), sort_keys=True),
            "request_body": "",
            "response_status": resp.status_code,
            "response_size_bytes": len(resp.content) if resp.content is not None else 0,
            "duration_ms": round(elapsed_ms, 1),
        })

        resp.raise_for_status()
        body = resp.json()
        token = body.get("data") if isinstance(body, dict) else None
        if not isinstance(token, str) or not token:
            raise RuntimeError(f"Unexpected token response shape: {body!r}")
        self._token = token
        return token

    # ------------------------------------------------------------------
    # All other reads go through get(); override to capture every call.
    # ------------------------------------------------------------------
    def get(self, path: str, params: Optional[Dict[str, Any]] = None) -> requests.Response:
        url = f"{self.base_url}{path}"
        headers = self._auth_headers()

        t0 = time.perf_counter()
        resp = self._session.get(
            url, headers=headers, params=params, timeout=REQUEST_TIMEOUT
        )
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        ctx = dict(self._current_phase)  # snapshot per call
        self.request_log.append({
            **ctx,
            "method": "GET",
            "url": url,
            "query_params": json.dumps(params, sort_keys=True) if params else "",
            "headers": json.dumps(_redact_headers(headers), sort_keys=True),
            "request_body": "",  # GETs don't carry bodies in this API
            "response_status": resp.status_code,
            "response_size_bytes": len(resp.content) if resp.content is not None else 0,
            "duration_ms": round(elapsed_ms, 1),
        })
        return resp


# ---------------------------------------------------------------------------
# Fetchers (logical phase is set on the client before each call)
# ---------------------------------------------------------------------------

def fetch_metadata(
    client: LoggingPtpClient, market: str, endpoint_name: str, identifier: str
) -> Tuple[bool, Dict[str, Any]]:
    client.set_phase(endpoint_name, identifier, "metadata", market)
    r = client.get(f"/ptp/{_encode_segment(market)}/{_encode_segment(endpoint_name)}")
    if r.ok:
        try:
            return True, r.json()
        except Exception as e:
            return False, {"status": r.status_code, "error": f"json parse: {e}", "raw": r.text[:500]}
    return False, {"status": r.status_code, "error": r.text[:500]}


def fetch_elements(
    client: LoggingPtpClient,
    market: str,
    endpoint_name: str,
    identifier: str,
    begin: str,
) -> Tuple[bool, Dict[str, Any]]:
    client.set_phase(endpoint_name, identifier, "elements", market, begin=begin)
    r = client.get(
        f"/ptp/{_encode_segment(market)}/{_encode_segment(endpoint_name)}/elements",
        params={"begin": begin},
    )
    if r.ok:
        try:
            return True, r.json()
        except Exception as e:
            return False, {"status": r.status_code, "error": f"json parse: {e}", "raw": r.text[:500]}
    return False, {"status": r.status_code, "error": r.text[:500]}


def fetch_sample(
    client: LoggingPtpClient,
    market: str,
    endpoint_name: str,
    identifier: str,
    begin: str,
    end: str,
) -> Tuple[bool, pd.DataFrame, str]:
    """One-day sample. Tries /query-columnar first, falls back to /query."""
    m = _encode_segment(market)
    e = _encode_segment(endpoint_name)
    params = {"begin": begin, "end": end}

    client.set_phase(endpoint_name, identifier, "sample-columnar", market, begin=begin, end=end)
    r = client.get(f"/ptp/{m}/{e}/query-columnar", params=params)
    if r.ok:
        try:
            df = _columnar_response_to_df(r.json())
            if not df.empty:
                return True, df, "columnar"
        except Exception:
            pass  # fall through to nested

    time.sleep(REQUEST_DELAY_SEC)
    client.set_phase(
        endpoint_name, identifier, "sample-nested-fallback", market, begin=begin, end=end
    )
    r = client.get(f"/ptp/{m}/{e}/query", params=params)
    if r.ok:
        try:
            df = _nested_response_to_df(r.json())
            return True, df, "nested"
        except Exception as ex:
            return False, pd.DataFrame(), f"nested parse error: {ex}"
    return False, pd.DataFrame(), f"HTTP {r.status_code}: {r.text[:300]}"


# ---------------------------------------------------------------------------
# Flatteners (unchanged from v1)
# ---------------------------------------------------------------------------

def extract_datapoints(metadata: Dict[str, Any]) -> List[Dict[str, Any]]:
    data = metadata.get("data") if isinstance(metadata, dict) else metadata
    if not isinstance(data, dict):
        return []

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
# Excel writer (adds Requests_Logical + Requests_HTTP sheets)
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


_LOGICAL_COLS = [
    "endpoint_name",
    "endpoint_identifier",
    "phase",
    "market",
    "begin",
    "end",
    "response_status",
    "response_size_bytes",
    "duration_ms",
]

_HTTP_COLS = [
    "endpoint_name",
    "phase",
    "method",
    "url",
    "query_params",
    "headers",
    "request_body",
    "response_status",
    "response_size_bytes",
    "duration_ms",
]


def write_workbook(
    output_path: Path,
    index_rows: List[Dict[str, Any]],
    schema_rows: List[Dict[str, Any]],
    element_rows: List[Dict[str, Any]],
    samples: Dict[str, pd.DataFrame],
    request_log: List[Dict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    used_names: set = set()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Index
        pd.DataFrame(index_rows).to_excel(
            writer, sheet_name=_safe_sheet_name("Index", used_names), index=False
        )

        # Schemas
        if schema_rows:
            pd.DataFrame(schema_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("Schemas", used_names), index=False
            )
        else:
            pd.DataFrame({"note": ["No schemas discovered."]}).to_excel(
                writer, sheet_name=_safe_sheet_name("Schemas", used_names), index=False
            )

        # Elements
        if element_rows:
            pd.DataFrame(element_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("Elements", used_names), index=False
            )
        else:
            pd.DataFrame({"note": ["No elements discovered."]}).to_excel(
                writer, sheet_name=_safe_sheet_name("Elements", used_names), index=False
            )

        # NEW: Request logs (logical view first for human readers, HTTP view for engineers)
        if request_log:
            log_df = pd.DataFrame(request_log)
            logical_cols_present = [c for c in _LOGICAL_COLS if c in log_df.columns]
            http_cols_present = [c for c in _HTTP_COLS if c in log_df.columns]
            log_df[logical_cols_present].to_excel(
                writer,
                sheet_name=_safe_sheet_name("Requests_Logical", used_names),
                index=False,
            )
            log_df[http_cols_present].to_excel(
                writer,
                sheet_name=_safe_sheet_name("Requests_HTTP", used_names),
                index=False,
            )
        else:
            pd.DataFrame({"note": ["No requests logged."]}).to_excel(
                writer,
                sheet_name=_safe_sheet_name("Requests_Logical", used_names),
                index=False,
            )

        # Per-endpoint sample sheets
        for endpoint_name, df in samples.items():
            sheet = _safe_sheet_name(endpoint_name, used_names)
            if df is None or df.empty:
                pd.DataFrame({"note": [f"No sample rows returned for {endpoint_name}."]}).to_excel(
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
        description=(
            "Discover every accessible PTP API endpoint and dump schema + samples + "
            "request log to Excel."
        )
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
        help=(
            "Output .xlsx path. Default: "
            "Tenaska_API_Discovery_<market>_<date>_with_requests.xlsx"
        ),
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
        f"Tenaska_API_Discovery_{market}_{sample_begin}_with_requests.xlsx"
    )

    print(f"Authenticating to {BASE_URL} ...")
    client = LoggingPtpClient(base_url=BASE_URL, username=USERNAME, password=PASSWORD)
    client.authenticate()
    print("  token acquired.")

    print(f"Listing endpoints under market '{market}' ...")
    client.set_phase("(market)", "", "list-endpoints", market)
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
        ok, meta = fetch_metadata(client, market, name, identifier)
        row["metadata_ok"] = ok
        if ok:
            dps = extract_datapoints(meta)
            row["datapoint_count"] = len(dps)
            for dp in dps:
                schema_rows.append({"endpoint": name, **dp})
            print(f"  metadata: {len(dps)} datapoints")
        else:
            row["error"] = f"metadata: {str(meta.get('error',''))[:200]}"
            print(f"  metadata FAILED: {meta}")
        time.sleep(REQUEST_DELAY_SEC)

        # Elements
        ok, elbody = fetch_elements(client, market, name, identifier, sample_begin)
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
        ok, df, note = fetch_sample(
            client, market, name, identifier, sample_begin, sample_end
        )
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

    print(f"\nLogged {len(client.request_log)} HTTP calls.")
    print(f"Writing {output} ...")
    write_workbook(
        output, index_rows, schema_rows, element_rows, samples, client.request_log
    )
    print(f"Done. Wrote {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
