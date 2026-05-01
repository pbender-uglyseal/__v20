"""
Tenaska PowerTools Platform (PTP) - API Discovery v3
====================================================

Same crawl as v2 with three new capabilities folded in, all derived from the
PowerTools API wiki (https://tenaska.atlassian.net/wiki/external/YjFjNzRhMmZhMTA5NDY3MGEzOGI2OThkOWU0YzE0MGM):

  1. /oas capture
       Hits the OpenAPI Specification endpoint and saves the response as a
       JSON sidecar next to the workbook. Adds an OAS sheet listing path,
       method, summary, tags, and operationId for every operation in the spec.
       Useful for the implementation team as a per-role API contract reference.

  2. Element batching for large endpoints
       Endpoints with element_count > --large-endpoint-threshold (default 500,
       e.g. Market_Price has 16,339 pNodes) get sampled in batches using the
       elementIdentifiers query parameter. Each batch is logged as its own
       phase in the request log so the dev team can see exactly how the
       chunking was performed. This works around the HTTP 400 / 10 MB header
       limit failure we hit in v1/v2.

  3. /options probing
       For every Reference-typed datapoint discovered in metadata, this hits
       /ptp/{market}/{endpoint}/options?elementDefinitions={defId} to fetch
       the lookup values (settlement points, transmission zones, etc.) the
       datapoint references. Captured in a new Options sheet. Best-effort:
       if /options returns 4xx for a definition, the failure is recorded in
       the request log and the script keeps going.

Output workbook sheets
----------------------
    Index             - one row per endpoint: counts + access status (extended)
    Schemas           - every datapoint across every endpoint (flat)
    Elements          - every element across every endpoint (flat)
    OAS               - operations from /oas (NEW)
    Options           - lookup values for Reference datapoints (NEW)
    Requests_Logical  - one row per HTTP call, business view
    Requests_HTTP     - one row per HTTP call, wire-level view
    <Endpoint>        - one sheet per endpoint with a one-day sample

Sidecar artifacts
-----------------
    <output>_oas.json - raw OpenAPI spec response, saved alongside the workbook

Usage
-----
    python tenaska_api_discovery_v3.py
    python tenaska_api_discovery_v3.py --sample-date 2026-04-20
    python tenaska_api_discovery_v3.py --batch-size 50 --large-endpoint-threshold 200
    python tenaska_api_discovery_v3.py --no-oas --no-options    # turn new features off

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

# Tenaska's stated scope ceiling: a single /query response must return
# fewer than ~10,000,000 records. The misleading "Estimated header size
# 20,407,411 exceeds limit of 10,000,000" error from v1/v2 runs was actually
# the records-count limit being hit. We default to half that as a safety
# margin so we don't graze the ceiling.
DEFAULT_RECORDS_BUDGET = 5_000_000

# Datapoint periodicity (in minutes) → intervals per day. periodicity=None
# (Meta datapoints) is treated as 1 record per day per element.
def _intervals_per_day(periodicity_minutes: Any) -> int:
    if periodicity_minutes in (None, 0, "None"):
        return 1
    try:
        m = int(periodicity_minutes)
    except (TypeError, ValueError):
        return 1
    if m <= 0:
        return 1
    return max(1, 1440 // m)


def _records_per_element(datapoints: List[Dict[str, Any]]) -> int:
    """Sum intervals-per-day across the endpoint's datapoints."""
    return sum(_intervals_per_day(dp.get("periodicity")) for dp in datapoints)


def _safe_batch_size(
    requested: int, records_per_element: int, budget: int
) -> int:
    """Cap batch size so projected records-per-batch stays under budget."""
    if records_per_element <= 0:
        return requested
    max_by_budget = max(1, budget // records_per_element)
    return min(requested, max_by_budget)

_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")

# Headers that should never appear in the request log in the clear.
_SENSITIVE_HEADER_NAMES = {"authorization", "x-api-key", "cookie"}


# ---------------------------------------------------------------------------
# Logging client (carried over from v2)
# ---------------------------------------------------------------------------

def _redact_headers(headers: Dict[str, str]) -> Dict[str, str]:
    return {
        k: ("<REDACTED>" if k.lower() in _SENSITIVE_HEADER_NAMES else v)
        for k, v in headers.items()
    }


class LoggingPtpClient(PtpClient):
    """PtpClient subclass that records every HTTP call it makes."""

    def __post_init__(self) -> None:
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
        self._current_phase = {
            "endpoint_name": endpoint_name,
            "endpoint_identifier": endpoint_identifier,
            "phase": phase,
            "market": market,
            "begin": begin,
            "end": end,
        }

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
            "endpoint_name": "(auth)",
            "endpoint_identifier": "",
            "phase": "authenticate",
            "market": "",
            "begin": None,
            "end": None,
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

    def get(self, path: str, params: Optional[Dict[str, Any]] = None) -> requests.Response:
        url = f"{self.base_url}{path}"
        headers = self._auth_headers()

        t0 = time.perf_counter()
        resp = self._session.get(
            url, headers=headers, params=params, timeout=REQUEST_TIMEOUT
        )
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        ctx = dict(self._current_phase)
        self.request_log.append({
            **ctx,
            "method": "GET",
            "url": url,
            "query_params": json.dumps(params, sort_keys=True, default=str) if params else "",
            "headers": json.dumps(_redact_headers(headers), sort_keys=True),
            "request_body": "",
            "response_status": resp.status_code,
            "response_size_bytes": len(resp.content) if resp.content is not None else 0,
            "duration_ms": round(elapsed_ms, 1),
        })
        return resp


# ---------------------------------------------------------------------------
# Fetchers
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


# --- NEW in v3: /oas capture ---------------------------------------------

def fetch_oas(client: LoggingPtpClient) -> Tuple[bool, Optional[Dict[str, Any]]]:
    """GET /oas - returns the per-role OpenAPI spec."""
    client.set_phase("(oas)", "", "oas", "")
    r = client.get("/oas")
    if r.ok:
        try:
            return True, r.json()
        except Exception:
            return False, None
    return False, None


def extract_oas_paths(oas: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Flatten OpenAPI spec into one row per (path, method)."""
    if not isinstance(oas, dict):
        return []
    paths = oas.get("paths", {})
    if not isinstance(paths, dict):
        return []

    rows: List[Dict[str, Any]] = []
    for path, ops in paths.items():
        if not isinstance(ops, dict):
            continue
        for method, op in ops.items():
            # Skip OpenAPI metadata keys like 'parameters', '$ref', 'x-*'
            if method.startswith("x-") or method in ("parameters", "summary", "description", "$ref"):
                continue
            if not isinstance(op, dict):
                continue
            params = op.get("parameters") or []
            param_names = [
                p.get("name") for p in params if isinstance(p, dict) and p.get("name")
            ]
            rows.append({
                "path": path,
                "method": method.upper(),
                "operationId": op.get("operationId", ""),
                "summary": (op.get("summary") or "")[:300],
                "tags": ", ".join(op.get("tags", []) if isinstance(op.get("tags"), list) else []),
                "parameters": ", ".join(param_names),
                "deprecated": bool(op.get("deprecated", False)),
            })
    return rows


# --- NEW in v3: batched sample fetch -------------------------------------

def fetch_sample_simple(
    client: LoggingPtpClient,
    market: str,
    endpoint_name: str,
    identifier: str,
    begin: str,
    end: str,
) -> Tuple[bool, pd.DataFrame, str]:
    """Original v1/v2 sample path: try columnar, fall back to nested. No batching."""
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
            pass

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


def fetch_sample_batched(
    client: LoggingPtpClient,
    market: str,
    endpoint_name: str,
    identifier: str,
    begin: str,
    end: str,
    element_ids: List[str],
    batch_size: int,
) -> Tuple[bool, pd.DataFrame, str]:
    """
    Sample data fetched in batches, filtering by elementIdentifiers.

    Used for endpoints where the unfiltered sample exceeds the API's response
    header limit (notably Market_Price with 16k+ pNodes).

    Each batch is logged as its own phase in the request log so the dev team
    can reproduce the chunking exactly.
    """
    m = _encode_segment(market)
    e = _encode_segment(endpoint_name)
    n_batches = (len(element_ids) + batch_size - 1) // batch_size
    frames: List[pd.DataFrame] = []
    successes = 0

    for i in range(0, len(element_ids), batch_size):
        batch = element_ids[i : i + batch_size]
        batch_idx = i // batch_size + 1
        phase = f"sample-batch-{batch_idx}-of-{n_batches}"
        client.set_phase(endpoint_name, identifier, phase, market, begin=begin, end=end)

        # requests serializes lists as repeated query params: ?elementIdentifiers=a&elementIdentifiers=b
        params = {
            "begin": begin,
            "end": end,
            "elementIdentifiers": batch,
        }
        r = client.get(f"/ptp/{m}/{e}/query-columnar", params=params)
        if r.ok:
            try:
                df = _columnar_response_to_df(r.json())
                if not df.empty:
                    frames.append(df)
                    successes += 1
            except Exception:
                pass  # batch failed; keep going

        time.sleep(REQUEST_DELAY_SEC)

    if not frames:
        return False, pd.DataFrame(), f"all {n_batches} batches returned empty or failed"

    combined = pd.concat(frames, ignore_index=True)
    return True, combined, f"batched-columnar ({successes}/{n_batches} batches succeeded)"


# --- NEW in v3: /options probing -----------------------------------------

def extract_reference_definitions(metadata: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Pull (datapoint_keyName, definition_id, options_link) for every Reference
    datapoint in an endpoint's metadata.

    The wiki shows datapoint metadata may include either a numeric definition
    id or a links.options URL; we capture whichever is available.
    """
    data = metadata.get("data") if isinstance(metadata, dict) else metadata
    if not isinstance(data, dict):
        return []

    candidates: List[Any] = []
    for key in ("dataPoints", "datapoints", "DataPoints"):
        v = data.get(key)
        if isinstance(v, list):
            candidates = v
            break

    refs: List[Dict[str, Any]] = []
    for dp in candidates:
        if not isinstance(dp, dict):
            continue
        # Recognize as Reference if objectType says so OR a links.options exists
        is_ref = False
        obj_type = dp.get("objectType") or dp.get("dataType") or ""
        if isinstance(obj_type, str) and "reference" in obj_type.lower():
            is_ref = True

        links = dp.get("links") if isinstance(dp.get("links"), dict) else {}
        options_link = links.get("options") if isinstance(links, dict) else None
        if options_link:
            is_ref = True

        if not is_ref:
            continue

        # Definition ID may live under several names
        def_id = (
            dp.get("elementDefinitionIdentifier")
            or dp.get("definitionIdentifier")
            or dp.get("definition")
            or dp.get("elementDefinition")
        )

        refs.append({
            "keyName": dp.get("keyName") or dp.get("name"),
            "definition_id": def_id,
            "options_link": options_link,
        })
    return refs


def fetch_options(
    client: LoggingPtpClient,
    market: str,
    endpoint_name: str,
    identifier: str,
    definition_id: Any,
) -> Tuple[bool, List[Dict[str, Any]], str]:
    """
    GET /ptp/{market}/{endpoint}/options?elementDefinitions={definition_id}.

    Best-effort: if the call fails or the response shape isn't what we expect,
    return [] and let the caller move on. Failures are visible in the
    request log because every call goes through the LoggingPtpClient.
    """
    if definition_id is None or definition_id == "":
        return False, [], "no definition id"

    m = _encode_segment(market)
    e = _encode_segment(endpoint_name)
    client.set_phase(
        endpoint_name, identifier, f"options-def-{definition_id}", market
    )
    r = client.get(
        f"/ptp/{m}/{e}/options",
        params={"elementDefinitions": str(definition_id)},
    )
    if not r.ok:
        return False, [], f"HTTP {r.status_code}"

    try:
        body = r.json()
    except Exception as ex:
        return False, [], f"json parse: {ex}"

    # Response shape (per wiki): typically { data: [ {name, identifier, ...}, ... ] }
    data = body.get("data") if isinstance(body, dict) else body
    rows: List[Dict[str, Any]] = []
    if isinstance(data, list):
        for opt in data:
            if not isinstance(opt, dict):
                continue
            rows.append({
                "name": opt.get("name"),
                "identifier": opt.get("identifier"),
                "elementDefinition": opt.get("elementDefinition") or opt.get("definition"),
                "elementDefinitionIdentifier": opt.get("elementDefinitionIdentifier"),
            })
    elif isinstance(data, dict):
        # Sometimes /options returns the same shape as the endpoint metadata —
        # in that case we extract from data.elements if present.
        elements = data.get("elements") or data.get("Elements") or []
        if isinstance(elements, list):
            for opt in elements:
                if not isinstance(opt, dict):
                    continue
                rows.append({
                    "name": opt.get("name"),
                    "identifier": opt.get("identifier"),
                    "elementDefinition": opt.get("elementDefinition"),
                    "elementDefinitionIdentifier": opt.get("elementDefinitionIdentifier"),
                })

    return True, rows, "ok" if rows else "empty response"


# ---------------------------------------------------------------------------
# Flatteners (unchanged from v1/v2)
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


_LOGICAL_COLS = [
    "endpoint_name", "endpoint_identifier", "phase", "market",
    "begin", "end", "response_status", "response_size_bytes", "duration_ms",
]

_HTTP_COLS = [
    "endpoint_name", "phase", "method", "url", "query_params",
    "headers", "request_body", "response_status",
    "response_size_bytes", "duration_ms",
]


def write_workbook(
    output_path: Path,
    index_rows: List[Dict[str, Any]],
    schema_rows: List[Dict[str, Any]],
    element_rows: List[Dict[str, Any]],
    samples: Dict[str, pd.DataFrame],
    request_log: List[Dict[str, Any]],
    oas_rows: List[Dict[str, Any]],
    options_rows: List[Dict[str, Any]],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    used_names: set = set()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(index_rows).to_excel(
            writer, sheet_name=_safe_sheet_name("Index", used_names), index=False
        )

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

        # NEW: OAS sheet
        if oas_rows:
            pd.DataFrame(oas_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("OAS", used_names), index=False
            )
        else:
            pd.DataFrame({"note": ["No OAS captured (skipped or unavailable)."]}).to_excel(
                writer, sheet_name=_safe_sheet_name("OAS", used_names), index=False
            )

        # NEW: Options sheet
        if options_rows:
            pd.DataFrame(options_rows).to_excel(
                writer, sheet_name=_safe_sheet_name("Options", used_names), index=False
            )
        else:
            pd.DataFrame({"note": ["No /options data probed (skipped or no Reference datapoints)."]}).to_excel(
                writer, sheet_name=_safe_sheet_name("Options", used_names), index=False
            )

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

        for endpoint_name, df in samples.items():
            sheet = _safe_sheet_name(endpoint_name, used_names)
            if df is None or df.empty:
                pd.DataFrame({"note": [f"No sample rows returned for {endpoint_name}."]}).to_excel(
                    writer, sheet_name=sheet, index=False
                )
            else:
                trimmed = df.head(MAX_SAMPLE_ROWS_PER_SHEET)
                trimmed.to_excel(writer, sheet_name=sheet, index=False)


def write_oas_sidecar(output_path: Path, oas: Dict[str, Any]) -> Path:
    """Save the raw OAS JSON next to the workbook."""
    sidecar = output_path.with_name(output_path.stem + "_oas.json")
    sidecar.write_text(json.dumps(oas, indent=2, default=str), encoding="utf-8")
    return sidecar


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
            "Discover every accessible PTP API endpoint and dump schema, samples, "
            "OAS, options, and request log to Excel."
        )
    )
    p.add_argument("--market", default="PJM")
    default_sample = (datetime.utcnow() - timedelta(days=3)).strftime("%Y-%m-%d")
    p.add_argument("--sample-date", type=_valid_date, default=default_sample)
    p.add_argument("--output", type=Path, default=None)
    p.add_argument("--only", nargs="*", default=None)

    # NEW v3 flags
    p.add_argument("--no-oas", action="store_true",
                   help="Skip the /oas capture (default: on).")
    p.add_argument("--no-options", action="store_true",
                   help="Skip /options probing for Reference datapoints (default: on).")
    p.add_argument("--no-batch", action="store_true",
                   help="Disable element batching for large endpoints (default: on).")
    p.add_argument("--batch-size", type=int, default=100,
                   help="Max elementIdentifiers per batched sample call (default: 100).")
    p.add_argument("--large-endpoint-threshold", type=int, default=500,
                   help="If element_count exceeds this, switch to batched sampling (default: 500).")
    p.add_argument("--records-budget", type=int, default=DEFAULT_RECORDS_BUDGET,
                   help=(
                       f"Max projected records per single API call (default: "
                       f"{DEFAULT_RECORDS_BUDGET:,}). Tenaska's stated ceiling "
                       f"is 10,000,000; we default to half for safety margin. "
                       f"Batch size is auto-shrunk per-endpoint to honor this."
                   ))
    return p.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    market = args.market
    sample_begin = args.sample_date
    sample_end = (
        datetime.strptime(sample_begin, "%Y-%m-%d") + timedelta(days=1)
    ).strftime("%Y-%m-%d")

    output = args.output or Path(
        f"Tenaska_API_Discovery_{market}_{sample_begin}_v3.xlsx"
    )

    print(f"Authenticating to {BASE_URL} ...")
    client = LoggingPtpClient(base_url=BASE_URL, username=USERNAME, password=PASSWORD)
    client.authenticate()
    print("  token acquired.")

    # NEW v3: capture /oas before walking endpoints
    oas_rows: List[Dict[str, Any]] = []
    oas_body: Optional[Dict[str, Any]] = None
    if not args.no_oas:
        print("Fetching /oas ...")
        ok, oas_body = fetch_oas(client)
        if ok and oas_body is not None:
            oas_rows = extract_oas_paths(oas_body)
            print(f"  OAS captured: {len(oas_rows)} operations")
        else:
            print("  OAS capture FAILED (see Requests_HTTP for details)")
        time.sleep(REQUEST_DELAY_SEC)

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
    options_rows: List[Dict[str, Any]] = []
    samples: Dict[str, pd.DataFrame] = {}
    seen_options_keys: set = set()

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
            "options_probed": 0,        # NEW v3
            "options_captured": 0,      # NEW v3
            "error": "",
        }

        # Metadata
        ok, meta = fetch_metadata(client, market, name, identifier)
        row["metadata_ok"] = ok
        ref_defs: List[Dict[str, Any]] = []
        records_per_element = 0
        if ok:
            dps = extract_datapoints(meta)
            row["datapoint_count"] = len(dps)
            for dp in dps:
                schema_rows.append({"endpoint": name, **dp})
            ref_defs = extract_reference_definitions(meta)
            records_per_element = _records_per_element(dps)
            print(
                f"  metadata: {len(dps)} datapoints  "
                f"({len(ref_defs)} Reference, ~{records_per_element} records/element/day)"
            )
        else:
            row["error"] = f"metadata: {str(meta.get('error',''))[:200]}"
            print(f"  metadata FAILED: {meta}")
        time.sleep(REQUEST_DELAY_SEC)

        # Elements
        element_ids: List[str] = []
        ok, elbody = fetch_elements(client, market, name, identifier, sample_begin)
        row["elements_ok"] = ok
        if ok:
            els = extract_elements(elbody)
            row["element_count"] = len(els)
            for el in els:
                element_rows.append({"endpoint": name, **el})
                if el.get("identifier"):
                    element_ids.append(el["identifier"])
            print(f"  elements: {len(els)}")
        else:
            err = elbody.get("error", "") if isinstance(elbody, dict) else ""
            if not row["error"]:
                row["error"] = f"elements: {str(err)[:200]}"
            print(f"  elements FAILED: {elbody}")
        time.sleep(REQUEST_DELAY_SEC)

        # Sample query — choose batched vs simple based on element count.
        # Also auto-shrink batch size for endpoints whose schema would push a
        # full batch over the records budget (Tenaska's stated 10M ceiling).
        use_batch = (
            not args.no_batch
            and row["element_count"] > args.large_endpoint_threshold
            and len(element_ids) > 0
        )
        # Even non-large endpoints get budget-checked: if a full unbatched
        # query would exceed the budget, we switch to batched mode anyway.
        projected_unbatched = row["element_count"] * max(records_per_element, 1)
        if not args.no_batch and projected_unbatched > args.records_budget and len(element_ids) > 0:
            use_batch = True

        if use_batch:
            effective_batch = _safe_batch_size(
                args.batch_size, records_per_element, args.records_budget
            )
            note_extras = ""
            if effective_batch < args.batch_size:
                note_extras = (
                    f" (auto-shrunk from {args.batch_size} to fit "
                    f"{args.records_budget:,} records/call budget)"
                )
            print(
                f"  sampling in batches: {len(element_ids)} elements / "
                f"{effective_batch} per batch{note_extras}"
            )
            ok, df, note = fetch_sample_batched(
                client, market, name, identifier,
                sample_begin, sample_end, element_ids, effective_batch,
            )
        else:
            ok, df, note = fetch_sample_simple(
                client, market, name, identifier, sample_begin, sample_end,
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

        # NEW v3: options probing
        if not args.no_options and ref_defs:
            for ref in ref_defs:
                def_id = ref.get("definition_id")
                key = (name, str(def_id))
                if def_id is None or key in seen_options_keys:
                    continue
                seen_options_keys.add(key)
                row["options_probed"] += 1
                ok, opt_rows, note = fetch_options(
                    client, market, name, identifier, def_id
                )
                if ok:
                    for opt in opt_rows:
                        options_rows.append({
                            "endpoint": name,
                            "datapoint": ref.get("keyName"),
                            "definition_id": def_id,
                            **opt,
                        })
                    row["options_captured"] += len(opt_rows)
                time.sleep(REQUEST_DELAY_SEC)

        index_rows.append(row)

    print(f"\nLogged {len(client.request_log)} HTTP calls.")
    print(f"Writing {output} ...")
    write_workbook(
        output, index_rows, schema_rows, element_rows, samples,
        client.request_log, oas_rows, options_rows,
    )
    if oas_body is not None:
        sidecar = write_oas_sidecar(output, oas_body)
        print(f"  also wrote {sidecar}")
    print(f"Done. Wrote {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
