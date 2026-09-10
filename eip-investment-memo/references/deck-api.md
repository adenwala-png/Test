# `eip_deck.py` API reference

Import it by path from the skill directory:

```python
import sys; sys.path.insert(0, "scripts")
from eip_deck import Memo, MARGIN_L, CONTENT_W, COL_W, COL_R_L, BODY_T, BODY_B, GREEN, RED
```

Everything is in inches on a 13.333 × 7.5in canvas.

## Contents

- [Memo](#memo) — deck-level
- [Slide methods](#slide-methods) — content placement
- [Grid constants](#grid-constants)
- [Palette](#palette)
- [Recipes](#recipes)

---

## Memo

```python
Memo(company, memo_type="Final Approval", sections=None, template=None)
```

`sections` defaults to the five standard sections; pass your own list to change the
divider slides. `template` defaults to `assets/eip-memo-template.pptx`.

| Method | Returns | Purpose |
|---|---|---|
| `title_slide(headline, terms=None, deal_team=None, sourced=None, date=None, subtitle=None)` | slide | Cover. `headline` is the one-sentence ask; `terms` is `[(label, value), ...]` |
| `divider(active_index)` | slide | Section divider with the active section picked out in green |
| `slide(title_bold, title_rest="")` | slide | Titled content slide — the workhorse |
| `blank_slide()` | slide | Untitled, for full-bleed exhibits |
| `save(path)` | Path | Writes the .pptx |

---

## Slide methods

### Text

**`title(bold_part, rest="")`** — the two-tone title. Called for you by `Memo.slide()`.

**`labeled_block(items, left=MARGIN_L, top=BODY_T, width=CONTENT_W, height=None, size=10, space_after=6)`**

The memo's basic prose unit. `items` is `[(label, body), ...]`; pass `label=None` for an
unlabelled paragraph. Renders as **Label**: body, with the label bold green and the body
grey.

**`bullets(lines, left, top, width, height=None, size=10, color=BODY_GREY)`**

For lists of facts. Use `labeled_block` when the content is an argument — bullets flatten
claim and evidence into the same visual weight.

**`heading(text, left, top, width, size=14, color=GREEN, height=0.34)`** — a sub-heading.

**`source(text)`** — grey source or note line pinned to the bottom. Every exhibit slide
should have one.

### Composites

**`two_columns(left_heading=None, right_heading=None, left_color=GREEN, right_color=GREEN, top=BODY_T)`**

Returns `(left_x, right_x, content_top, col_width)` and draws the divider rule. Place
anything into either side without recomputing the grid.

**`pros_cons(thesis, risks, top=BODY_T)`**

The thesis/risks slide. Both arguments are `[(label, body), ...]`. Risks are headed in
red deliberately — an IC reader should find the bear case without hunting.

**`commentary(text, left=COL_R_L, top=BODY_T, width=COL_W, height=None, heading="Commentary")`**

The rail that reads an exhibit. `text` may be a string or a list of paragraphs. Pass
`heading=None` to drop the "Commentary" label.

**`kpi_tiles(tiles, left=MARGIN_L, top=BODY_T, width=CONTENT_W, height=0.95)`**

A row of headline figures from `[(value, label), ...]`. Four to five tiles is the
comfortable maximum across the full width.

### Exhibits

**`table(rows, left, top, width, height=None, col_widths=None, size=10, header_size=11, align=None, first_col_left=True)`**

`rows[0]` is the header, rendered white on EIP green. `col_widths` are relative weights
(`[3, 1.2, 1.5]`). `align` is a per-column list of `PP_ALIGN` values; the default puts
the first column left and the rest centred.

Roughly 12 rows fit a full-height slide. Continue onto a second slide rather than
shrinking type below 9pt.

**`chart(chart_type, categories, series, left, top, width, height=None, title=None, number_format='#,##0', show_values=True, legend=None)`**

Native PowerPoint charts, EIP-coloured. `chart_type` is one of `column`,
`stacked_column`, `bar`, `line`, `pie`. `series` is `[(name, values), ...]`.

Native rather than images so the deck stays editable — an IC member can click into a
number, and the deal team can revise a forecast without regenerating a picture. Only use
an image for chart types PowerPoint has no native form for.

**`image(path, left, top, width=None, height=None)`** — pass one of width/height to
preserve aspect ratio.

**`textbox(left, top, width, height)`** — escape hatch. Returns a python-pptx textbox
with margins and wrapping already set; style runs yourself with the palette constants.

**`rule(x=None, top=1.05, bottom=BODY_B)`** — vertical divider, defaults to the column gutter.

---

## Grid constants

| Constant | Value | Meaning |
|---|---|---|
| `SLIDE_W`, `SLIDE_H` | 13.333, 7.5 | Canvas |
| `MARGIN_L`, `MARGIN_R` | 0.23 | Side margins |
| `CONTENT_W` | 12.87 | Full content width |
| `TITLE_T`, `TITLE_H` | 0.24, 0.34 | Title band |
| `BODY_T` | 0.68 | First usable row under the title |
| `BODY_B` | 7.05 | Last usable row above the source line |
| `SOURCE_T` | 7.12 | Source line |
| `COL_W` | 6.30 | One of two equal columns |
| `COL_R_L` | 6.79 | Left edge of the right column |
| `GUTTER` | 0.26 | Between columns |

For an asymmetric split — wide prose with a narrow commentary rail — 8.6in / 4.05in
starting at 9.05in is the common pairing.

---

## Palette

Applied as direct run colours. The template's theme is stock Office; EIP's identity
lives in the runs, so set these explicitly.

| Constant | Hex | Use |
|---|---|---|
| `GREEN` | `#005C2A` | Headlines, lead-in labels, emphasis |
| `GREEN_ACCENT` | `#007742` | Separators, figures worth the eye |
| `GREEN_DATE` | `#2F854C` | Cover date |
| `TITLE_GREY` | `#919191` | Descriptive half of a title |
| `BODY_GREY` | `#595959` | Body copy |
| `MUTED_GREY` | `#717171` | Deal team, secondary metadata |
| `SOURCE_GREY` | `#BFBFBF` | Source and note lines |
| `RED` | `#C00000` | Risks, downside cases |

Typeface is Nunito throughout. Body copy is 10pt, titles 20pt, section headings 14pt,
table headers 11pt, source lines 9pt.

---

## Recipes

**Exhibit with commentary rail**

```python
s = memo.slide("Backlog", "Recovery Sustained Through Four Consecutive Quarters")
s.chart("column", quarters, [("Backlog", values)],
        left=0.23, top=1.0, width=7.6, height=5.6, title="Backlog ($mm)")
s.commentary(["Backlog bottomed in 4Q23 and has grown every quarter since.",
              "Growth is award-driven rather than a change in recognition timing."],
             left=8.1, top=1.0, width=5.0, height=5.6)
s.source("Note: a project is treated as won when revenue is first recognised.")
```

**KPI row above an exhibit**

```python
s.kpi_tiles([("$62.4M", "FY25 Revenue"), ("16.0%", "FY25 Adj. EBITDA Margin")])
s.chart("column", years, series, top=1.85, height=4.9, width=7.6)
```

**Forecast case table** — same column shape in every case slide, so cases compare line
by line:

```python
s.table([["($mm)", "FY26E", "FY27E", "FY28E", "FY29E"],
         ["Revenue", "71.8", "80.4", "88.9", "97.2"],
         ["Adj. EBITDA", "12.1", "14.0", "15.8", "17.6"],
         ["Margin %", "16.9%", "17.4%", "17.8%", "18.1%"]],
        col_widths=[2.4, 1, 1, 1, 1], top=1.0, height=2.6)
```
