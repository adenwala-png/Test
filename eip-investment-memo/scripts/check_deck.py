#!/usr/bin/env python3
"""Catch the layout faults you would otherwise only find by opening the deck.

A memo is usually assembled without anyone looking at it until it is finished, and the
failure mode is always the same: a paragraph that was written long runs past the bottom
of its box, a table with more rows than the slide has room for, a shape nudged off the
canvas. PowerPoint renders all of these silently -- the text is simply gone from view.

So this checks the geometry arithmetic that a human eye would otherwise have to do:

  offcanvas   a shape sits partly or wholly outside the 13.33 x 7.5in slide
  overflow    estimated text height exceeds the box it was placed in
  collision   two text boxes overlap enough that one will cover the other
  font        a run in a font other than Nunito (usually an inherited default)
  palette     a color outside the EIP palette, which reads as off-brand
  empty       a slide with a title and nothing else, or an empty placeholder

Run it after building and fix what it reports before sending the deck on:

    python scripts/check_deck.py memo.pptx

Overflow is estimated, not measured -- there is no font metric available here, so it uses
an average character width. It is tuned to under-report rather than cry wolf: a flagged
box is very likely genuinely overfull, but a clean report is not proof of a clean deck.
Exits 1 if any error-level finding is present.
"""

import argparse
import sys
from collections import defaultdict

from pptx import Presentation
from pptx.util import Emu

EMU_IN = 914400.0
SLIDE_W, SLIDE_H = 13.333, 7.5
TOL = 0.02  # inches of slack before calling something off-canvas

EIP_PALETTE = {
    "005C2A", "007742", "2F854C", "919191", "595959", "717171",
    "BFBFBF", "C00000", "FFFFFF", "000000", "D9D9D9", "EEEEEE",
    "F2F7F4", "DDE8E1", "9DC3AE", "A4A3A4",
}

# Average glyph width as a fraction of point size, and line height as a multiple of it.
# Nunito is a fairly narrow humanist sans; 0.48 tracks it well enough to spot real overruns.
CHAR_W_RATIO = 0.48
LINE_H_RATIO = 1.22


def _in(v):
    return (v or 0) / EMU_IN


def estimate_text_height(shape):
    """Rough rendered height, in inches, of the text in a shape."""
    tf = shape.text_frame
    usable_w = _in(shape.width) - _in(tf.margin_left) - _in(tf.margin_right)
    if usable_w <= 0.05:
        return 0.0
    total = _in(tf.margin_top) + _in(tf.margin_bottom)
    for para in tf.paragraphs:
        runs = para.runs
        if not runs:
            total += 0.10
            continue
        size = max((r.font.size.pt for r in runs if r.font.size), default=10)
        text = "".join(r.text for r in runs)
        char_w = size * CHAR_W_RATIO / 72.0
        chars_per_line = max(1, int(usable_w / char_w))
        lines = max(1, -(-len(text) // chars_per_line))
        total += lines * size * LINE_H_RATIO / 72.0
        total += (para.space_after.pt if para.space_after else 0) / 72.0
    return total


def iter_text_shapes(slide):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            yield sh


def overlap_area(a, b):
    ax1, ay1, ax2, ay2 = _in(a.left), _in(a.top), _in(a.left) + _in(a.width), _in(a.top) + _in(a.height)
    bx1, by1, bx2, by2 = _in(b.left), _in(b.top), _in(b.left) + _in(b.width), _in(b.top) + _in(b.height)
    w = min(ax2, bx2) - max(ax1, bx1)
    h = min(ay2, by2) - max(ay1, by1)
    return max(0.0, w) * max(0.0, h)


def check(path, strict_palette=False):
    prs = Presentation(path)
    findings = []

    def add(level, slide_no, kind, msg):
        findings.append((level, slide_no, kind, msg))

    for n, slide in enumerate(prs.slides, 1):
        text_shapes = []

        for sh in slide.shapes:
            if sh.left is None or sh.top is None:
                continue
            l, t = _in(sh.left), _in(sh.top)
            r, b = l + _in(sh.width), t + _in(sh.height)
            label = f"{sh.shape_type}:{sh.name!r}"

            if l < -TOL or t < -TOL or r > SLIDE_W + TOL or b > SLIDE_H + TOL:
                add("error", n, "offcanvas",
                    f"{label} spans ({l:.2f},{t:.2f})-({r:.2f},{b:.2f}); slide is {SLIDE_W}x{SLIDE_H}")

            if getattr(sh, "has_table", False):
                rows = len(sh.table.rows)
                needed = sum(_in(row.height) for row in sh.table.rows)
                if t + needed > SLIDE_H + TOL:
                    add("error", n, "overflow",
                        f"table {label} with {rows} rows needs {needed:.2f}in from top {t:.2f}in "
                        f"-- runs {t + needed - SLIDE_H:.2f}in past the slide")

            if sh.has_text_frame and sh.text_frame.text.strip():
                text_shapes.append(sh)
                est = estimate_text_height(sh)
                box_h = _in(sh.height)
                # Autofit/auto-grow boxes legitimately exceed their nominal height, but a
                # box overrunning the slide bottom is a real problem either way.
                if est > box_h * 1.12 and t + est > SLIDE_H - 0.1:
                    add("error", n, "overflow",
                        f"{label} text needs ~{est:.2f}in in a {box_h:.2f}in box at top {t:.2f}in "
                        f"-- likely runs off the slide")
                elif est > box_h * 1.35:
                    add("warn", n, "overflow",
                        f"{label} text needs ~{est:.2f}in in a {box_h:.2f}in box")

                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.name and run.font.name != "Nunito" and not run.font.name.startswith("+"):
                            add("warn", n, "font", f"{label} uses {run.font.name!r}, expected 'Nunito'")
                        try:
                            rgb = run.font.color.rgb
                        except (AttributeError, TypeError):
                            rgb = None
                        if rgb and str(rgb).upper() not in EIP_PALETTE:
                            add("warn" if strict_palette else "info", n, "palette",
                                f"{label} uses #{rgb} -- outside the EIP palette")

        for i, a in enumerate(text_shapes):
            for b in text_shapes[i + 1:]:
                area = overlap_area(a, b)
                smaller = min(_in(a.width) * _in(a.height), _in(b.width) * _in(b.height))
                if smaller > 0 and area / smaller > 0.35:
                    add("warn", n, "collision",
                        f"{a.name!r} and {b.name!r} overlap over {area / smaller:.0%} of the smaller box")

        # Judge emptiness by content, not shape count: a section divider is legitimately
        # a single text box, while a slide holding only its title is an unfinished stub.
        body_chars = sum(
            len(sh.text_frame.text.strip())
            for sh in text_shapes
            if _in(sh.top) > 0.45  # exclude the title band
        )
        has_exhibit = any(
            getattr(sh, "has_table", False) or getattr(sh, "has_chart", False) or sh.shape_type == 13
            for sh in slide.shapes
        )
        if body_chars < 25 and not has_exhibit:
            add("warn", n, "empty", "slide has a title but almost no content -- is it finished?")

    return prs, findings


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--strict-palette", action="store_true",
                    help="treat off-palette colors as warnings rather than notes")
    ap.add_argument("--quiet", action="store_true", help="suppress info-level notes")
    args = ap.parse_args()

    prs, findings = check(args.deck, args.strict_palette)

    levels = defaultdict(int)
    for level, *_ in findings:
        levels[level] += 1

    print(f"{args.deck}: {len(prs.slides)} slides")
    shown = [f for f in findings if not (args.quiet and f[0] == "info")]
    if not shown:
        print("  no layout problems found")
    for level, slide_no, kind, msg in sorted(shown, key=lambda f: (f[1], f[0])):
        print(f"  [{level:5}] slide {slide_no:>2} {kind}: {msg}")

    print(f"\n  {levels['error']} error(s), {levels['warn']} warning(s), {levels['info']} note(s)")
    return 1 if levels["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
