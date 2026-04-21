"""
Tenaska PowerTools Platform (PTP) - Generator Energy Charge Details Exporter
============================================================================

Pulls the following PJM datasets from the Tenaska PowerTools Platform API
(https://api.ptp.energy) for a user-specified date range and writes the
results to an Excel workbook with one sheet per dataset:

    - Generator Energy Charge Details - 5 minute
    - Generator Hourly Energy Charge Details

API docs: https://tenaska.atlassian.net/wiki/external/YjFjNzRhMmZhMTA5NDY3MGEzOGI2OThkOWU0YzE0MGM

Usage
-----
    python tenaska_generator_energy_charges.py --start 2026-04-01 --end 2026-04-15
    python tenaska_generator_energy_charges.py --start 2026-04-01 --end 2026-04-15 --output my_report.xlsx

    # Discover the exact endpoint names available under a market:
    python tenaska_generator_energy_charges.py --list-endpoints
    python tenaska_generator_energy_charges.py --list-endpoints --market PJM

Notes
-----
- Credentials are hardcoded below per request. For anything beyond a quick,
  one-off script you should move them out of source (env vars, .env, or a
  secret manager) and rotate the password once done.
- The API rate-limits at roughly 1 call / second on a sliding window, so a
  short delay is applied between requests.
- The token endpoint returns a JWT valid for 24 hours; the script fetches
  one fresh token per run and reuses it.
- Both /query-columnar and /query are attempted for each dataset; the
  columnar endpoint returns tidier tabular output, while the nested /query
  endpoint is used as a fallback and flattened into rows.

Dependencies
------------
    pip install requests pandas openpyxl
"""

import argparse
import base64
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
from urllib.parse import quote

import pandas as pd
import requests


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# NOTE: Hardcoded per user request. Replace with env vars for production use.
USERNAME = "pbender@medullus.com"
PASSWORD = "fyk1YRZ-pue-cdw3jyk"

BASE_URL = "https://api.ptp.energy"
TOKEN_PATH = "/authentication/token"

# Market / root under which the datasets live (PJM per the screenshots).
MARKET = "PJM"

# Datasets to pull. Keys become Excel sheet names; values are the endpoint
# names exactly as the API expects them. If a name is rejected with
# "not found or inaccessible" (status code 2303), run the script with
# --list-endpoints to see the exact names available under the market and
# update this dict.
DATASETS: Dict[str, str] = {
    "Gen Energy Charge 5min": "Generator Energy Charge Details - 5 minute",
    "Gen Hourly Energy Charge": "Generator Hourly Energy Charge Details",
}

# Seconds between requests to stay under the rate limit.
REQUEST_DELAY_SEC = 1.1

# Request timeout (seconds).
REQUEST_TIMEOUT = 120


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

@dataclass
class PtpClient:
    """Minimal PowerTools Platform API client."""

    base_url: str
    username: str
    password: str
    _token: Optional[str] = None
    _session: requests.Session = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        self._session = requests.Session()

    def authenticate(self) -> str:
        """Exchange Basic Auth credentials for a JWT (valid 24h)."""
        creds = f"{self.username}:{self.password}".encode("utf-8")
        basic = base64.b64encode(creds).decode("ascii")
        headers = {"Authorization": f"Basic {basic}"}
        url = f"{self.base_url}{TOKEN_PATH}"

        resp = self._session.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        body = resp.json()

        token = body.get("data") if isinstance(body, dict) else None
        if not isinstance(token, str) or not token:
            raise RuntimeError(f"Unexpected token response shape: {body!r}")
        self._token = token
        return token

    def _auth_headers(self) -> Dict[str, str]:
        if not self._token:
            self.authenticate()
        return {"Authorization": f"Bearer {self._token}"}

    def get(self, path: str, params: Optional[Dict[str, Any]] = None) -> requests.Response:
        url = f"{self.base_url}{path}"
        resp = self._session.get(
            url,
            headers=self._auth_headers(),
            params=params,
            timeout=REQUEST_TIMEOUT,
        )
        return resp


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _encode_segment(segment: str) -> str:
    """URL-encode a path segment."""
    return quote(segment, safe="")


# ---------------------------------------------------------------------------
# Discovery
# ---------------------------------------------------------------------------

def list_markets(client: PtpClient) -> List[Dict[str, Any]]:
    resp = client.get("/ptp")
    resp.raise_for_status()
    body = resp.json()
    data = body.get("data") if isinstance(body, dict) else body
    return data if isinstance(data, list) else []


def list_endpoints(client: PtpClient, market: str) -> List[Dict[str, Any]]:
    """Return the list of endpoints available under a market."""
    resp = client.get(f"/ptp/{_encode_segment(market)}")
    resp.raise_for_status()
    body = resp.json()
    data = body.get("data") if isinstance(body, dict) else body

    candidates: List[Any] = []
    if isinstance(data, dict):
        for key in ("endpoints", "Endpoints", "links", "children"):
            v = data.get(key)
            if isinstance(v, list):
                candidates = v
                break
    elif isinstance(data, list):
        candidates = data

    endpoints: List[Dict[str, Any]] = []
    for item in candidates:
        if isinstance(item, dict):
            endpoints.append({
                "name": item.get("name") or item.get("endpointName") or item.get("title"),
                "identifier": item.get("identifier") or item.get("endpointIdentifier") or item.get("id"),
                "url": item.get("url") or item.get("href"),
            })
    return endpoints


def print_endpoints(client: PtpClient, market: str) -> None:
    print(f"Endpoints available under market '{market}':\n")
    eps = list_endpoints(client, market)
    if not eps:
        print("  (no endpoints returned - printing raw response for debugging)\n")
        resp = client.get(f"/ptp/{_encode_segment(market)}")
        print(resp.text[:4000])
        return

    name_width = max((len(e.get("name") or "") for e in eps), default=4)
    id_width = max((len(e.get("identifier") or "") for e in eps), default=10)
    header = f"  {'name'.ljust(name_width)}  {'identifier'.ljust(id_width)}"
    print(header)
    print("  " + "-" * (len(header) - 2))
    for e in eps:
        print(f"  {(e.get('name') or '').ljust(name_width)}  {(e.get('identifier') or '').ljust(id_width)}")


# ---------------------------------------------------------------------------
# Querying
# ---------------------------------------------------------------------------

def query_dataset(
    client: PtpClient,
    market: str,
    endpoint_name: str,
    begin: str,
    end: str,
) -> pd.DataFrame:
    """
    Fetch a dataset for a date range and return a flat DataFrame.

    Tries /query-columnar first. If that fails, falls back to /query and
    flattens the nested element/datapoint/value tree.
    """
    market_seg = _encode_segment(market)
    ep_seg = _encode_segment(endpoint_name)
    params = {"begin": begin, "end": end}

    columnar_path = f"/ptp/{market_seg}/{ep_seg}/query-columnar"
    resp = client.get(columnar_path, params=params)
    if resp.ok:
        try:
            df = _columnar_response_to_df(resp.json())
            if not df.empty:
                return df
        except Exception as exc:  # noqa: BLE001
            print(f"  columnar parse failed ({exc}); falling back to /query", file=sys.stderr)
    else:
        # Surface the server's error message so the caller can act on it.
        detail = resp.text[:500]
        print(
            f"  columnar query returned HTTP {resp.status_code}: {detail}",
            file=sys.stderr,
        )
        # If this was a "not found" error, there's no point falling back.
        if resp.status_code in (403, 404) and "not found" in detail.lower():
            raise RuntimeError(
                f"Query failed for '{endpoint_name}': "
                f"HTTP {resp.status_code} - {detail}"
            )

    time.sleep(REQUEST_DELAY_SEC)

    nested_path = f"/ptp/{market_seg}/{ep_seg}/query"
    resp = client.get(nested_path, params=params)
    if not resp.ok:
        raise RuntimeError(
            f"Query failed for '{endpoint_name}': "
            f"HTTP {resp.status_code} - {resp.text[:500]}"
        )
    return _nested_response_to_df(resp.json())


# ---------------------------------------------------------------------------
# Response parsing
# ---------------------------------------------------------------------------

def _columnar_response_to_df(body: Dict[str, Any]) -> pd.DataFrame:
    """Flatten the /query-columnar response into a DataFrame."""
    data = body.get("data") if isinstance(body, dict) else body

    if isinstance(data, list):
        return pd.json_normalize(data)

    if isinstance(data, dict):
        array_cols = {k: v for k, v in data.items() if isinstance(v, list)}
        if array_cols and len({len(v) for v in array_cols.values()}) == 1:
            return pd.DataFrame(array_cols)

        rows: List[Dict[str, Any]] = []
        for element_name, payload in data.items():
            if isinstance(payload, list):
                for row in payload:
                    if isinstance(row, dict):
                        rows.append({"element": element_name, **row})
        if rows:
            return pd.json_normalize(rows)

    return pd.DataFrame()


def _nested_response_to_df(body: Dict[str, Any]) -> pd.DataFrame:
    """Flatten the nested /query response."""
    elements = body.get("data", []) if isinstance(body, dict) else []
    if not isinstance(elements, list):
        return pd.DataFrame()

    indexed: Dict[tuple, Dict[str, Any]] = {}

    for el in elements:
        if not isinstance(el, dict):
            continue
        el_meta = {
            "element": el.get("element"),
            "identifier": el.get("identifier"),
            "definition": el.get("definition"),
            "parent": el.get("parent"),
            "parentDefinition": el.get("parentDefinition"),
            "parentIdentifier": el.get("parentIdentifier"),
            "goLiveDate": el.get("goLiveDate"),
            "expirationDate": el.get("expirationDate"),
        }
        for dp in el.get("dataPoints", []) or []:
            key_name = dp.get("keyName")
            for interval in dp.get("values", []) or []:
                start = interval.get("intervalStartUtc")
                end = interval.get("intervalEndUtc")
                data_values = interval.get("data")

                if isinstance(data_values, list):
                    for dv in data_values:
                        value, coords = _extract_datavalue(dv)
                        row_key = (
                            el_meta["identifier"],
                            start,
                            end,
                            _coords_key(coords),
                        )
                        row = indexed.setdefault(
                            row_key,
                            {
                                **el_meta,
                                "intervalStartUtc": start,
                                "intervalEndUtc": end,
                                **(coords or {}),
                            },
                        )
                        row[key_name] = value
                else:
                    row_key = (el_meta["identifier"], start, end, None)
                    row = indexed.setdefault(
                        row_key,
                        {
                            **el_meta,
                            "intervalStartUtc": start,
                            "intervalEndUtc": end,
                        },
                    )
                    row[key_name] = data_values

    if not indexed:
        return pd.DataFrame()

    df = pd.DataFrame(list(indexed.values()))
    lead = [c for c in ("element", "identifier", "intervalStartUtc", "intervalEndUtc") if c in df.columns]
    rest = [c for c in df.columns if c not in lead]
    return df[lead + rest]


def _extract_datavalue(dv: Any):
    if not isinstance(dv, dict):
        return dv, None
    value = dv.get("value", dv.get("data"))
    coords = dv.get("coords")
    if isinstance(coords, list):
        coords = {f"coord_{i}": c for i, c in enumerate(coords)}
    elif not isinstance(coords, dict):
        coords = None
    return value, coords


def _coords_key(coords: Optional[Dict[str, Any]]):
    if not coords:
        return None
    return tuple(sorted(coords.items()))


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


def _safe_sheet_name(name: str) -> str:
    clean = _INVALID_SHEET_CHARS.sub(" ", name).strip()
    return (clean[:31]) or "Sheet"


def write_excel(frames: Dict[str, pd.DataFrame], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        wrote_any = False
        for sheet, df in frames.items():
            safe = _safe_sheet_name(sheet)
            if df is None or df.empty:
                pd.DataFrame({"note": [f"No data returned for '{sheet}'."]}).to_excel(
                    writer, sheet_name=safe, index=False
                )
            else:
                df.to_excel(writer, sheet_name=safe, index=False)
            wrote_any = True
        if not wrote_any:
            pd.DataFrame({"note": ["No datasets requested."]}).to_excel(
                writer, sheet_name="Empty", index=False
            )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _valid_date(s: str) -> str:
    try:
        datetime.strptime(s, "%Y-%m-%d")
    except ValueError as e:
        raise argparse.ArgumentTypeError(f"Expected YYYY-MM-DD, got {s!r}") from e
    return s


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=(
            "Pull PJM Generator Energy Charge Details (5-minute and hourly) "
            "from the Tenaska PowerTools Platform API and save to Excel."
        )
    )
    p.add_argument("--start", required=False, type=_valid_date,
                   help="Start date (YYYY-MM-DD), inclusive. Required unless --list-endpoints.")
    p.add_argument("--end", required=False, type=_valid_date,
                   help="End date (YYYY-MM-DD), inclusive. Required unless --list-endpoints.")
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help=(
            "Output .xlsx path. Defaults to "
            "'V20_Generator_Energy_Charges_<start>_to_<end>.xlsx' in the cwd."
        ),
    )
    p.add_argument(
        "--market",
        default=MARKET,
        help=f"Market / root name to query (default: {MARKET}).",
    )
    p.add_argument(
        "--list-endpoints",
        action="store_true",
        help=(
            "Print the endpoints available under the given market and exit. "
            "Use this to find the exact endpoint names to paste into the "
            "DATASETS dict at the top of this script."
        ),
    )
    args = p.parse_args(argv)
    if not args.list_endpoints and (not args.start or not args.end):
        p.error("--start and --end are required unless --list-endpoints is used")
    return args


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)

    print(f"Authenticating to {BASE_URL} ...")
    client = PtpClient(base_url=BASE_URL, username=USERNAME, password=PASSWORD)
    client.authenticate()
    print("  token acquired.")

    if args.list_endpoints:
        print_endpoints(client, args.market)
        return 0

    output = args.output or Path(
        f"V20_Generator_Energy_Charges_{args.start}_to_{args.end}.xlsx"
    )

    frames: Dict[str, pd.DataFrame] = {}
    for sheet_label, endpoint_name in DATASETS.items():
        print(f"Fetching '{endpoint_name}' ({args.start} -> {args.end}) ...")
        try:
            df = query_dataset(
                client,
                market=args.market,
                endpoint_name=endpoint_name,
                begin=args.start,
                end=args.end,
            )
            print(f"  got {len(df):,} rows, {len(df.columns)} columns.")
            frames[sheet_label] = df
        except Exception as e:  # noqa: BLE001
            msg = str(e)
            print(f"  ERROR fetching '{endpoint_name}': {msg}", file=sys.stderr)
            if "not found or inaccessible" in msg or "2303" in msg:
                print(
                    "  HINT: run with --list-endpoints to see the exact "
                    "endpoint names available under this market, then "
                    "update the DATASETS dict at the top of this script.",
                    file=sys.stderr,
                )
            frames[sheet_label] = pd.DataFrame(
                {"error": [msg], "endpoint": [endpoint_name]}
            )
        time.sleep(REQUEST_DELAY_SEC)

    print(f"Writing {output} ...")
    write_excel(frames, output)
    print(f"Done. Wrote {output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())