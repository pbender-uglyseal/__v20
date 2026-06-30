#!/usr/bin/env python3
"""
webvisu_capture.py
==================

Safe data-capture automation for a browser-based WebVisu / SCADA HMI that is
rendered on HTML5 <canvas> elements (#background / #foreground) rather than
semantic DOM controls.

Because the real UI is painted on canvas, normal selector-based clicking is not
available for the on-screen controls. We therefore click by *position* -- but we
do so DEFENSIVELY:

  * Calibration coordinates were captured once, in a known 1280x1024 setup.
  * Those raw pixels are converted into RATIOS relative to the live canvas.
  * Before every click we re-measure the canvas bounding box and recompute the
    click point from the current rendered size.

This is the key safeguard: if the viewport, zoom, window decorations, or page
scaling ever change, a hard-coded absolute coordinate could silently land on the
WRONG element. A ratio that is re-applied to the live geometry tracks the target
control instead, so the click stays on the intended spot (or fails loudly if the
canvas is missing/resized unexpectedly).

Pipeline
--------
    1. open page (self-signed cert bypassed)  ->
    2. proportional navigation clicks         ->
    3. screenshot of JUST the final page      ->
    4. OCR that screenshot with Tesseract     ->
    5. reconstruct the on-screen grid -> CSV

Dependencies
------------
    python -m pip install --upgrade pip
    pip install playwright pandas pytesseract pillow
    python -m playwright install chromium

Tesseract is a separate native program (pytesseract is just a wrapper):
    Windows: install the UB-Mannheim build, then either add it to PATH or pass
             --tesseract-cmd "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
    The script auto-detects the default install path if you don't pass one.

Usage
-----
    python webvisu_capture.py --url https://192.168.x.x/webvisu.htm
    python webvisu_capture.py --url https://... --headless          # optional later
    python webvisu_capture.py --url https://... --tesseract-cmd "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

Playwright needs both the Python package AND the separate browser download
(`python -m playwright install chromium`).
"""

from __future__ import annotations

import argparse
import logging
import os
import shutil
import statistics
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from playwright.sync_api import (
    Browser,
    BrowserContext,
    Error as PlaywrightError,
    Page,
    sync_playwright,
)

# pytesseract is a thin wrapper around the native `tesseract` binary. Import it
# lazily-tolerantly so the navigation half of the script still works even if OCR
# deps aren't installed yet; we raise a clear error only when OCR is attempted.
try:
    import pytesseract
    from PIL import Image
    _OCR_IMPORT_ERROR: Optional[str] = None
except ImportError as exc:  # pragma: no cover
    pytesseract = None  # type: ignore[assignment]
    Image = None  # type: ignore[assignment]
    _OCR_IMPORT_ERROR = str(exc)

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

# The viewport used for calibration. Fixing the viewport gives us consistent,
# repeatable canvas rendering so the ratio math below stays valid. These are the
# DEFAULTS for --width/--height, set to match the current LS Energy calibration
# (1920x1080), so the script can be run with no viewport flags at all.
CALIB_WIDTH = 1920
CALIB_HEIGHT = 1080

# The interaction layer for a WebVisu app is the #foreground canvas (it sits on
# top of #background and receives pointer events).
INTERACTION_CANVAS = "#foreground"
BACKGROUND_CANVAS = "#background"


@dataclass(frozen=True)
class CalibrationPoint:
    """A click position measured during calibration, plus the coordinate space
    it was measured in. Stored as raw pixels; converted to ratios at runtime."""

    name: str
    x: float
    y: float
    ref_width: int = CALIB_WIDTH
    ref_height: int = CALIB_HEIGHT


# The two on-screen controls that must be clicked, in sequence, to reach the
# target data view. Measured in the exact 1280x1024 calibration setup.
#
# IMPORTANT: these are CALIBRATION values, not the final design. They are only
# meaningful as ratios of the reference coordinate space (see ref_width/height).
# Measured with calibrate_clicks.py at viewport 1920x1080. Run capture with the
# SAME viewport: --width 1920 --height 1080.
CALIBRATION_POINTS = (
    CalibrationPoint(name="nav_step_1", x=55, y=71),
    CalibrationPoint(name="nav_step_2", x=1757, y=87),
)

# Pause between clicks so the canvas app can repaint the new view.
INTER_CLICK_DELAY_S = 1.5

logger = logging.getLogger("webvisu_capture")


# --------------------------------------------------------------------------- #
# Click geometry
# --------------------------------------------------------------------------- #

def calculate_click_position(
    cal: CalibrationPoint,
    box: dict,
    ref_width: int,
    ref_height: int,
) -> tuple[float, float]:
    """Convert a calibration pixel into a live click point on the current canvas.

    The calibration point is expressed as a *proportion* of the reference
    coordinate space, then that same proportion is applied to the LIVE bounding
    box of the canvas. This means:

        - If the canvas renders at exactly the calibration size, the click lands
          on the original pixel.
        - If the canvas was scaled, moved, or the window was resized, the click
          tracks the intended control instead of a stale absolute pixel.

    `ref_width`/`ref_height` are the viewport the calibration pixels were
    measured in -- which MUST equal the --width/--height used at capture time.
    Calibrate and capture at the same viewport (calibrate_clicks.py uses the
    same flags) and these ratios stay valid even as you change resolution.

    `box` is a Playwright bounding box: {"x", "y", "width", "height"} in CSS
    pixels relative to the top-left of the viewport.
    """
    # Step 1: normalize the calibration pixel to a 0..1 ratio of the reference
    # space it was captured in.
    x_ratio = cal.x / ref_width
    y_ratio = cal.y / ref_height

    # Guardrail: a ratio outside 0..1 means the calibration pixel sits OUTSIDE
    # the reference space it was supposedly measured in (e.g. x=1602 against a
    # 1280-wide reference -> 1.25). That would project the click off the live
    # element. Warn loudly and clamp into bounds rather than click into the void.
    if not (0.0 <= x_ratio <= 1.0) or not (0.0 <= y_ratio <= 1.0):
        logger.warning(
            "[%s] calibration ratio out of bounds (x=%.3f, y=%.3f). The point "
            "(%.0f,%.0f) lies outside the %dx%d reference. Clamping to edge -- "
            "did you calibrate at a different --width/--height than capture?",
            cal.name, x_ratio, y_ratio, cal.x, cal.y, ref_width, ref_height,
        )
    x_ratio = min(max(x_ratio, 0.0), 1.0)
    y_ratio = min(max(y_ratio, 0.0), 1.0)

    # Step 2: re-apply that ratio to the LIVE rendered geometry.
    live_x = box["x"] + box["width"] * x_ratio
    live_y = box["y"] + box["height"] * y_ratio

    logger.info(
        "[%s] ratio=(%.4f, %.4f) box=(x=%.1f y=%.1f w=%.1f h=%.1f) -> click=(%.1f, %.1f)",
        cal.name, x_ratio, y_ratio,
        box["x"], box["y"], box["width"], box["height"],
        live_x, live_y,
    )
    return live_x, live_y


# --------------------------------------------------------------------------- #
# Browser lifecycle
# --------------------------------------------------------------------------- #

def open_page(context: BrowserContext, url: str, load_wait: float = 5.0) -> Page:
    """Open the target page and confirm it loaded before any interaction.

    `load_wait` is the pause AFTER the document loads, before we click. A heavy
    WebVisu HMI keeps initializing its canvas event handlers for several seconds
    after 'load' fires -- clicking too early gets silently ignored (the view
    never navigates). When automation clicks miss but manual clicks work, this
    is almost always the cause: a human is slower than a 2s timer. Give it room.
    """
    page = context.new_page()
    logger.info("Navigating to %s", url)

    # wait_until="load" ensures the document and its sub-resources (including the
    # WebVisu canvas bootstrap) have finished loading before we touch anything.
    page.goto(url, wait_until="load", timeout=30_000)

    # Give the canvas app time to run its onload Webvisu(...) initializer AND
    # wire up its pointer handlers before we start clicking.
    logger.info("Waiting %.1fs for the WebVisu canvas to become interactive ...",
                load_wait)
    page.wait_for_timeout(int(load_wait * 1000))
    logger.info("Page loaded: title=%r", page.title())
    return page


def _get_canvas_box(page: Page, selector: str) -> dict:
    """Return the live bounding box of a canvas, failing loudly if absent."""
    locator = page.locator(selector)

    if locator.count() == 0:
        raise RuntimeError(
            f"Expected canvas {selector!r} not found. The page may not have "
            f"loaded, or this is not the WebVisu interface we calibrated against."
        )

    box = locator.bounding_box()
    if box is None or box["width"] == 0 or box["height"] == 0:
        raise RuntimeError(
            f"Canvas {selector!r} has no measurable geometry "
            f"(box={box}). Refusing to click blind."
        )
    return box


def _get_reference_box(page: Page) -> dict:
    """Return the coordinate space the calibration ratios apply to.

    Prefers the live #foreground canvas (the real WebVisu interaction layer). If
    that canvas isn't present -- e.g. when pointed at a plain-DOM test page --
    we fall back to the viewport so the same proportional clicks still work.
    """
    locator = page.locator(INTERACTION_CANVAS)
    if locator.count() > 0:
        box = locator.bounding_box()
        if box and box["width"] and box["height"]:
            return box

    size = page.viewport_size or {"width": CALIB_WIDTH, "height": CALIB_HEIGHT}
    logger.warning(
        "%s canvas not found; falling back to viewport %dx%d for click math.",
        INTERACTION_CANVAS, size["width"], size["height"],
    )
    return {"x": 0.0, "y": 0.0,
            "width": float(size["width"]), "height": float(size["height"])}


def perform_navigation_clicks(
    page: Page,
    artifact_dir: Path,
    ref_width: int,
    ref_height: int,
    inter_click_delay: float = INTER_CLICK_DELAY_S,
    click_dwell_ms: float = 120.0,
) -> None:
    """Click the two required controls, re-measuring geometry before each click.

    Guardrails on every step:
      * confirm the interaction canvas still exists and has real size
      * sanity-check the live box against the calibration aspect so we notice if
        the page was unexpectedly resized/zoomed
      * log the bounding box and computed click point
      * screenshot before each click AND after the sequence

    `click_dwell_ms` holds the mouse button down briefly so a WebVisu canvas
    button actually latches the press (a too-fast synthetic click can be missed).
    `inter_click_delay` lets each resulting view finish rendering before the
    next click -- the 2nd control may not exist until the 1st view transitions.
    """
    save_debug_artifacts(page, artifact_dir, "01_before_clicks")

    for idx, cal in enumerate(CALIBRATION_POINTS, start=1):
        # Re-measure EVERY time -- never reuse a stale box. A view transition
        # between clicks can re-layout the canvas.
        box = _get_reference_box(page)

        # Optional drift warning: the calibration assumed a specific aspect
        # ratio. If the live canvas differs a lot, the proportional mapping is
        # still applied, but we flag it so a human can re-verify the target.
        _warn_on_scale_drift(box, ref_width, ref_height)

        live_x, live_y = calculate_click_position(cal, box, ref_width, ref_height)

        logger.info("Click %d/%d (%s) at (%.1f, %.1f) dwell=%.0fms",
                    idx, len(CALIBRATION_POINTS), cal.name, live_x, live_y,
                    click_dwell_ms)
        # move -> press -> dwell -> release, so the canvas registers a real click.
        page.mouse.move(live_x, live_y)
        page.mouse.down()
        page.wait_for_timeout(int(click_dwell_ms))
        page.mouse.up()

        # Let the canvas repaint the resulting view before the next action, and
        # snapshot the intermediate state for debugging the navigation path.
        time.sleep(inter_click_delay)
        save_debug_artifacts(page, artifact_dir, f"01b_after_click_{idx}")

    save_debug_artifacts(page, artifact_dir, "02_after_clicks")


def _warn_on_scale_drift(
    box: dict, ref_width: int, ref_height: int, tolerance: float = 0.05,
) -> None:
    """Warn (do not abort) if the live canvas size diverges from calibration.

    We keep going -- the ratio approach is exactly what makes a resized canvas
    survivable -- but a large divergence is worth surfacing in case the page is
    showing something other than what we calibrated against.
    """
    expected_aspect = ref_width / ref_height
    live_aspect = box["width"] / box["height"]
    if abs(live_aspect - expected_aspect) / expected_aspect > tolerance:
        logger.warning(
            "Canvas aspect ratio drift: live=%.3f vs calibration=%.3f. "
            "Proportional mapping still applied, but re-verify the target.",
            live_aspect, expected_aspect,
        )


# --------------------------------------------------------------------------- #
# Screenshots
# --------------------------------------------------------------------------- #

def save_debug_artifacts(page: Page, artifact_dir: Path, label: str) -> Path:
    """Save a full-page screenshot for verification/troubleshooting."""
    artifact_dir.mkdir(parents=True, exist_ok=True)
    shot_path = artifact_dir / f"{label}.png"
    page.screenshot(path=str(shot_path), full_page=True)
    logger.info("Saved screenshot: %s", shot_path)
    return shot_path


def save_final_page_screenshot(page: Page, artifact_dir: Path) -> Path:
    """Capture ONLY the final page (the visible viewport), not a stitched
    full-page scroll. This is the image we feed to OCR -- a single, exactly
    1280x1024 frame keeps text crisp and the geometry predictable.
    """
    artifact_dir.mkdir(parents=True, exist_ok=True)
    shot_path = artifact_dir / "final_page.png"
    # full_page=False -> just the current viewport, i.e. the final view as shown.
    page.screenshot(path=str(shot_path), full_page=False)
    logger.info("Saved final-page screenshot: %s", shot_path)
    return shot_path


def crop_to_region(src: Path, dst: Path, crop: tuple[float, float, float, float]) -> Path:
    """Crop an image to a fractional (left, top, right, bottom) region and save.

    Fractions are 0..1 of the image's own dimensions, so the crop is independent
    of resolution / device scale factor. We crop OUT the left-hand fault-group
    panels and gauges on the LS Energy 'Summary of all BRiCs' view so OCR sees
    only the parameter-label column + the BRiC #1..#N value columns. The cropped
    image is saved so you can eyeball exactly what OCR reads and tweak --crop.
    """
    left, top, right, bottom = crop
    img = Image.open(src)
    w, h = img.width, img.height
    box = (int(w * left), int(h * top), int(w * right), int(h * bottom))
    img.crop(box).save(dst)
    logger.info("Cropped %s -> %s  (fractions L=%.3f T=%.3f R=%.3f B=%.3f -> px %s)",
                src.name, dst.name, left, top, right, bottom, box)
    return dst


# --------------------------------------------------------------------------- #
# OCR  ->  CSV
# --------------------------------------------------------------------------- #

def configure_tesseract(tesseract_cmd: Optional[str]) -> None:
    """Point pytesseract at the native tesseract binary.

    Order of preference: explicit --tesseract-cmd, then PATH, then the default
    Windows install location. Fails with an actionable message if not found.
    """
    if pytesseract is None:
        raise RuntimeError(
            f"OCR dependencies missing ({_OCR_IMPORT_ERROR}). "
            f"Run: pip install pytesseract pillow"
        )

    candidates = []
    if tesseract_cmd:
        candidates.append(tesseract_cmd)
    on_path = shutil.which("tesseract")
    if on_path:
        candidates.append(on_path)
    candidates.append(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    candidates.append(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe")

    for cand in candidates:
        if cand and os.path.exists(cand):
            pytesseract.pytesseract.tesseract_cmd = cand
            logger.info("Using tesseract binary: %s", cand)
            return
        if cand == on_path and on_path:  # found on PATH, no need to verify path
            pytesseract.pytesseract.tesseract_cmd = cand
            logger.info("Using tesseract binary from PATH: %s", cand)
            return

    raise RuntimeError(
        "Could not locate the tesseract binary. Install the UB-Mannheim build "
        "and pass --tesseract-cmd \"C:\\Program Files\\Tesseract-OCR\\tesseract.exe\"."
    )


def _cluster_rows(words: list[dict], y_tol_factor: float = 0.6) -> list[list[dict]]:
    """Group OCR words into visual rows by their vertical center.

    Two words belong to the same row if their y-centers are within a tolerance
    derived from the median word height (handles slight baseline jitter).
    """
    if not words:
        return []
    heights = [w["h"] for w in words]
    y_tol = max(6.0, statistics.median(heights) * y_tol_factor)

    words_sorted = sorted(words, key=lambda w: w["cy"])
    rows: list[list[dict]] = [[words_sorted[0]]]
    for w in words_sorted[1:]:
        # Compare against the running mean y-center of the current row.
        row_cy = statistics.mean(x["cy"] for x in rows[-1])
        if abs(w["cy"] - row_cy) <= y_tol:
            rows[-1].append(w)
        else:
            rows.append([w])
    return rows


def _detect_column_gutters(words: list[dict], img_width: int) -> list[float]:
    """Find vertical whitespace gutters that separate columns.

    We build an x-occupancy mask from every word box (across ALL rows), then look
    for runs of empty x where no text ever appears. Splitting on those gutters --
    rather than on per-word gaps -- is what keeps multi-word cells like
    "Output kW" or "All Inverters" together in one column instead of fracturing.

    Returns a sorted list of x boundaries (gutter midpoints) used to bin words
    into columns.
    """
    if not words:
        return []

    occupied = bytearray(img_width + 1)
    for w in words:
        left = max(0, int(w["x"]))
        right = min(img_width, int(w["x"] + w["w"]))
        for x in range(left, right + 1):
            occupied[x] = 1

    # A gutter must be wider than a normal intra-cell space. Tie the threshold to
    # the median word height so it scales with font size / zoom.
    median_h = statistics.median(w["h"] for w in words)
    min_gutter = max(12, median_h * 1.0)

    boundaries: list[float] = []
    run_start: Optional[int] = None
    for x in range(img_width + 1):
        if occupied[x] == 0:
            if run_start is None:
                run_start = x
        else:
            if run_start is not None:
                run_len = x - run_start
                # Ignore the leading/trailing margins; only interior gaps split.
                if run_len >= min_gutter and run_start > 0:
                    boundaries.append((run_start + x) / 2.0)
                run_start = None
    return boundaries


def _assign_column(cx: float, boundaries: list[float]) -> int:
    """Return the column index for a word center, given gutter boundaries."""
    col = 0
    for b in boundaries:
        if cx > b:
            col += 1
        else:
            break
    return col


def ocr_image_to_dataframe(
    image_path: Path,
    min_conf: int = 40,
    drop_sparse_below: int = 2,
    upscale: float = 2.0,
) -> Optional[pd.DataFrame]:
    """OCR a screenshot and reconstruct the on-screen grid into a DataFrame.

    Strategy:
      1. tesseract image_to_data -> per-word text + bounding boxes + confidence.
      2. Detect column gutters from global whitespace (stable across rows).
      3. Cluster words into rows by vertical center.
      4. Place each word into (row, column); join multi-word cells by x order.
      5. Use the first well-populated row as the header.

    `drop_sparse_below` removes near-empty rows (e.g. a stray page title above
    the table) that would otherwise become junk lines in the CSV.
    """
    # Preprocess for accuracy: grayscale + upscale. Screen captures are only
    # ~96 DPI; Tesseract reads small digits (and decimal points!) far more
    # reliably when enlarged. All downstream geometry is relative, so the
    # scale-up is transparent to the row/column reconstruction. When the capture
    # was already taken at a high device-scale-factor, pass a smaller --upscale
    # so we don't blow the image up needlessly (slow, no accuracy gain).
    img = Image.open(image_path).convert("L")
    if upscale and upscale != 1.0:
        img = img.resize(
            (int(img.width * upscale), int(img.height * upscale)), Image.LANCZOS
        )

    data = pytesseract.image_to_data(
        img,
        output_type=pytesseract.Output.DICT,
    )

    words: list[dict] = []
    for i in range(len(data["text"])):
        text = (data["text"][i] or "").strip()
        try:
            conf = float(data["conf"][i])
        except (ValueError, TypeError):
            conf = -1.0
        if not text or conf < min_conf:
            continue
        x, y, w, h = (data["left"][i], data["top"][i],
                      data["width"][i], data["height"][i])
        words.append({
            "text": text, "x": x, "y": y, "w": w, "h": h,
            "cx": x + w / 2.0, "cy": y + h / 2.0,
        })

    if not words:
        logger.warning("OCR found no text above confidence %d.", min_conf)
        return None

    img_width = max(w["x"] + w["w"] for w in words) + 5
    boundaries = _detect_column_gutters(words, int(img_width))
    n_cols = len(boundaries) + 1
    logger.info("OCR: %d words, %d columns detected.", len(words), n_cols)

    rows = _cluster_rows(words)

    grid: list[list[str]] = []
    for row_words in rows:
        cells: list[list[dict]] = [[] for _ in range(n_cols)]
        for w in row_words:
            cells[_assign_column(w["cx"], boundaries)].append(w)
        # Join words within a cell left-to-right.
        row_cells = [
            " ".join(t["text"] for t in sorted(c, key=lambda t: t["x"]))
            for c in cells
        ]
        non_empty = sum(1 for c in row_cells if c)
        if non_empty < drop_sparse_below:
            logger.info("Dropping sparse row (%d cells): %r",
                        non_empty, [c for c in row_cells if c])
            continue
        grid.append(row_cells)

    if not grid:
        logger.warning("No populated rows after reconstruction.")
        return None

    header, *body = grid
    # Guard against duplicate/blank header names so pandas stays happy.
    header = [h if h else f"col_{i}" for i, h in enumerate(header)]
    seen: dict[str, int] = {}
    uniq_header = []
    for h in header:
        if h in seen:
            seen[h] += 1
            uniq_header.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            uniq_header.append(h)

    df = pd.DataFrame(body, columns=uniq_header)
    logger.info("Reconstructed table: %d data rows x %d cols",
                len(df), df.shape[1])
    return df


def write_csv(df: pd.DataFrame, out_path: Path) -> None:
    """Write a DataFrame to CSV without the pandas index column."""
    df.to_csv(out_path, index=False)
    logger.info("Wrote %d rows to %s", len(df), out_path)


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #

def run(
    url: str,
    headless: bool,
    out_dir: Path,
    tesseract_cmd: Optional[str],
    min_conf: int,
    width: int,
    height: int,
    scale: float,
    settle_delay: float,
    load_wait: float,
    inter_click_delay: float,
    click_dwell_ms: float,
    crop: tuple[float, float, float, float],
) -> int:
    run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    artifact_dir = out_dir / f"run_{run_stamp}"
    csv_path = artifact_dir / "captured_data.csv"

    with sync_playwright() as p:
        browser: Browser = p.chromium.launch(headless=headless)

        # The browser CONTEXT carries the settings that matter most here:
        #
        #   ignore_https_errors=True
        #     The HMI is served over HTTPS with a SELF-SIGNED certificate on a
        #     private LAN address. Chromium would normally block the page with a
        #     cert error and the automation could never reach the canvas. This
        #     flag suppresses cert validation *for this context only* so the
        #     page loads. It is appropriate for a trusted device on a closed
        #     network; do not point this script at public/untrusted hosts.
        #
        #   viewport=width x height
        #     A fixed viewport gives deterministic canvas rendering, which keeps
        #     the calibration ratios meaningful run-to-run. Widen it (and
        #     re-calibrate at the same size) if the data view is clipped.
        #
        #   device_scale_factor=scale
        #     Renders at higher pixel density WITHOUT changing layout or click
        #     coordinates -- the screenshot comes out scale-x larger, so OCR
        #     reads small HMI digits far more reliably. This is real high-DPI
        #     rendering, not a blurry post-hoc upscale.
        context: BrowserContext = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": width, "height": height},
            device_scale_factor=scale,
        )

        exit_code = 0
        try:
            page = open_page(context, url, load_wait=load_wait)

            perform_navigation_clicks(
                page, artifact_dir, width, height,
                inter_click_delay=inter_click_delay,
                click_dwell_ms=click_dwell_ms,
            )

            # Let the final view fully render before we capture it. A WebVisu
            # transition can lag behind the click, so wait before screenshotting.
            if settle_delay > 0:
                logger.info("Waiting %.1fs for the final view to settle ...",
                            settle_delay)
                page.wait_for_timeout(int(settle_delay * 1000))

            # Screenshot ONLY the final page (viewport), then OCR it to CSV.
            final_png = save_final_page_screenshot(page, artifact_dir)

            # Crop to the table region (drop left-hand fault/gauge panels) so OCR
            # only reads the parameter labels + BRiC columns. crop=(0,0,1,1)
            # disables it. The cropped image is saved for visual verification.
            ocr_png = final_png
            if crop != (0.0, 0.0, 1.0, 1.0):
                ocr_png = crop_to_region(
                    final_png, artifact_dir / "final_page_cropped.png", crop,
                )

            configure_tesseract(tesseract_cmd)
            # The capture is already `scale`-x high-DPI; only add software
            # upscale if that leaves us below ~2x effective resolution.
            sw_upscale = max(1.0, 2.0 / scale)
            df = ocr_image_to_dataframe(
                ocr_png, min_conf=min_conf, upscale=sw_upscale,
            )

            if df is not None and not df.empty:
                write_csv(df, csv_path)
                logger.info("Done. CSV: %s", csv_path)
            else:
                logger.warning(
                    "OCR produced no table. Review %s and try lowering "
                    "--min-conf.", final_png,
                )
                exit_code = 1
        except (PlaywrightError, RuntimeError) as exc:
            logger.error("Aborted safely: %s", exc)
            # Best-effort failure screenshot.
            try:
                save_debug_artifacts(page, artifact_dir, "99_error_state")
            except Exception:
                pass
            exit_code = 1
        finally:
            context.close()
            browser.close()

    return exit_code


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--url", required=True,
        help="HTTPS URL of the WebVisu page (self-signed cert OK).",
    )
    parser.add_argument(
        "--headless", action="store_true",
        help="Run Chromium headless. Default is headed for debugging.",
    )
    parser.add_argument(
        "--out-dir", default="./webvisu_runs", type=Path,
        help="Directory for screenshots and CSV output.",
    )
    parser.add_argument(
        "--tesseract-cmd", default=None,
        help=r'Path to tesseract.exe (e.g. "C:\Program Files\Tesseract-OCR\tesseract.exe"). '
             "Auto-detected from PATH / default install if omitted.",
    )
    parser.add_argument(
        "--min-conf", type=int, default=40,
        help="Minimum OCR word confidence (0-100). Lower it if cells are dropped.",
    )
    parser.add_argument(
        "--width", type=int, default=CALIB_WIDTH,
        help=f"Viewport width (default {CALIB_WIDTH}). Widen if the data view is "
             "clipped. MUST match the width used in calibrate_clicks.py.",
    )
    parser.add_argument(
        "--height", type=int, default=CALIB_HEIGHT,
        help=f"Viewport height (default {CALIB_HEIGHT}). MUST match calibration.",
    )
    parser.add_argument(
        "--scale", type=float, default=2.0,
        help="Device scale factor / pixel density for OCR clarity (default 2.0). "
             "Try 3.0 for very small HMI text. Does not affect click coordinates.",
    )
    parser.add_argument(
        "--settle-delay", type=float, default=2.0,
        help="Seconds to wait after the last click before the final screenshot "
             "(default 2.0). Increase if the data view renders slowly.",
    )
    parser.add_argument(
        "--load-wait", type=float, default=5.0,
        help="Seconds to wait after page load before the first click (default "
             "5.0). Increase if navigation clicks are ignored -- the WebVisu "
             "canvas needs time to become interactive.",
    )
    parser.add_argument(
        "--inter-click-delay", type=float, default=3.0,
        help="Seconds to wait between the two navigation clicks (default 3.0). "
             "Increase if the 2nd click fires before the 1st view has rendered.",
    )
    parser.add_argument(
        "--click-dwell", type=float, default=120.0,
        help="Milliseconds to hold each click down (default 120) so a WebVisu "
             "canvas button latches the press. Increase if clicks don't register.",
    )
    parser.add_argument(
        "--crop", type=str, default="0.35,0.0,1.0,1.0",
        help="Fractional crop 'left,top,right,bottom' (0..1) applied before OCR "
             "to drop the left-hand fault/gauge panels. Default '0.35,0,1,1' "
             "keeps the parameter labels + BRiC columns. Use '0,0,1,1' for none. "
             "Check final_page_cropped.png and nudge the left value if needed.",
    )
    return parser.parse_args(argv)


def _parse_crop(text: str) -> tuple[float, float, float, float]:
    """Parse a 'left,top,right,bottom' crop string into a validated tuple."""
    try:
        parts = tuple(float(p) for p in text.split(","))
    except ValueError:
        raise SystemExit(f"--crop must be four numbers, got {text!r}")
    if len(parts) != 4:
        raise SystemExit(f"--crop needs exactly 4 values, got {len(parts)}")
    left, top, right, bottom = parts
    if not (0.0 <= left < right <= 1.0) or not (0.0 <= top < bottom <= 1.0):
        raise SystemExit(f"--crop must satisfy 0<=left<right<=1 and "
                         f"0<=top<bottom<=1, got {parts}")
    return parts  # type: ignore[return-value]


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )
    args = parse_args()
    sys.exit(run(
        url=args.url,
        headless=args.headless,
        out_dir=args.out_dir,
        tesseract_cmd=args.tesseract_cmd,
        min_conf=args.min_conf,
        width=args.width,
        height=args.height,
        scale=args.scale,
        settle_delay=args.settle_delay,
        load_wait=args.load_wait,
        inter_click_delay=args.inter_click_delay,
        click_dwell_ms=args.click_dwell,
        crop=_parse_crop(args.crop),
    ))


if __name__ == "__main__":
    main()
