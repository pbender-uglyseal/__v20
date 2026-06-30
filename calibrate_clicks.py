#!/usr/bin/env python3
"""
calibrate_clicks.py
===================

Interactive click-calibration helper for webvisu_capture.py.

Opens the target page in a Playwright Chromium window at EXACTLY the same
1280x1024 viewport the automation uses, then records the precise viewport
coordinates of every click you make -- and draws a marker where you clicked.

Why this is the right way to calibrate
--------------------------------------
The automation clicks with page.mouse.click(x, y), which uses VIEWPORT
coordinates (the page content area), not desktop/screen pixels. Measuring on
your monitor with a ruler or screenshot tool includes window chrome, title
bars, and any OS scaling -- which is exactly how x=1602 crept in against a
1280-wide page. Capturing the click here removes all of that: what you click is
what the automation will click.

Usage
-----
    python calibrate_clicks.py --url https://192.168.x.x/webvisu.htm

Then, in the browser window that opens:
    1. Click the FIRST control (nav step 1).
    2. Click the SECOND control (nav step 2).
    3. Switch back to this terminal and press Enter.

It prints a ready-to-paste CALIBRATION_POINTS block. Copy it into
webvisu_capture.py (keeping the reference at 1280x1024).

Dependencies: same as the main script (playwright). No OCR needed here.
"""

from __future__ import annotations

import argparse
from typing import Optional

from playwright.sync_api import sync_playwright

# Defaults match webvisu_capture.py (LS Energy HMI calibrated at 1920x1080).
CALIB_WIDTH = 1920
CALIB_HEIGHT = 1080

# JS injected into the page: on every click, record viewport coords and draw a
# small numbered dot so you can see exactly where each click landed.
_MARKER_JS = """
() => {
  window.__clicks = [];
  document.addEventListener('click', (e) => {
    const x = Math.round(e.clientX);
    const y = Math.round(e.clientY);
    window.__clicks.push({x, y});
    window.recordClick(x, y);
    const dot = document.createElement('div');
    dot.textContent = window.__clicks.length;
    Object.assign(dot.style, {
      position: 'fixed', left: (x - 10) + 'px', top: (y - 10) + 'px',
      width: '20px', height: '20px', borderRadius: '50%',
      background: 'rgba(255,0,0,0.65)', color: '#fff', font: 'bold 12px sans-serif',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: 2147483647, pointerEvents: 'none', border: '2px solid #fff',
    });
    document.body.appendChild(dot);
  }, true);  // capture phase: record even if the canvas swallows the event
}
"""


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--url", required=True,
        help="HTTPS/HTTP URL of the page to calibrate against.",
    )
    parser.add_argument(
        "--width", type=int, default=CALIB_WIDTH,
        help=f"Viewport width (default {CALIB_WIDTH}). Use the SAME value you "
             "pass to webvisu_capture.py --width.",
    )
    parser.add_argument(
        "--height", type=int, default=CALIB_HEIGHT,
        help=f"Viewport height (default {CALIB_HEIGHT}). Match capture --height.",
    )
    args = parser.parse_args()
    width, height = args.width, args.height

    collected: list[dict] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            ignore_https_errors=True,
            viewport={"width": width, "height": height},
        )
        page = context.new_page()

        # Bridge: the page calls window.recordClick(x, y); we capture it here.
        def record_click(x: int, y: int) -> None:
            collected.append({"x": x, "y": y})
            print(f"  recorded click #{len(collected)}: x={x}, y={y}")

        page.expose_function("recordClick", record_click)

        print(f"Opening {args.url} at {width}x{height} ...")
        page.goto(args.url, wait_until="load", timeout=30_000)
        page.evaluate(_MARKER_JS)

        print("\n" + "=" * 64)
        print("Click your controls IN ORDER in the browser window:")
        print("  1) the first nav control")
        print("  2) the second nav control")
        print("Each click prints its coordinates here and drops a numbered dot.")
        print("When done, come back and press Enter.")
        print("=" * 64 + "\n")

        try:
            input("Press Enter when finished clicking... ")
        except (EOFError, KeyboardInterrupt):
            pass

        context.close()
        browser.close()

    _print_result(collected, width, height)


def _print_result(points: list[dict], width: int, height: int) -> None:
    if not points:
        print("\nNo clicks were recorded.")
        return

    print("\n" + "=" * 64)
    print(f"Calibrated at {width}x{height}. Run capture with the SAME viewport:")
    print(f"    python webvisu_capture.py --url <URL> --width {width} --height {height}")
    print("Paste this block into webvisu_capture.py:")
    print("=" * 64)
    print("CALIBRATION_POINTS = (")
    for i, pt in enumerate(points[:2], start=1):
        print(f'    CalibrationPoint(name="nav_step_{i}", '
              f'x={pt["x"]}, y={pt["y"]}),')
    print(")")
    if len(points) > 2:
        print(f"\n(Note: {len(points)} clicks recorded; used the first two. "
              f"All: {points})")


if __name__ == "__main__":
    main()
