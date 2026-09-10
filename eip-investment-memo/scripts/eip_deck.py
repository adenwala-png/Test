"""Build EIP investment memo decks on EIP's own template.

The point of this module is that a memo slide is not a blank canvas. EIP memos have a
settled visual grammar -- a two-tone title, 10pt Nunito body copy where the lead-in label
is bold green and the explanation is grey, a commentary rail beside every exhibit, a grey
source line pinned to the bottom. Rebuilding that from raw python-pptx on every memo wastes
effort and drifts: one deck ends up 11pt, another puts the title at 0.3" instead of 0.24",
and the set stops looking like it came from one firm.

So this module owns the geometry and the run styling, and the caller owns the argument.
Write `slide.labeled("Company", "Acme, founded in 2009 ...")` and the colors, font, size and
position are already right.

Typical use:

    from eip_deck import Memo

    memo = Memo(company="Acme Utility Services", memo_type="Final Approval")
    memo.title_slide(
        headline="EIP to invest up to $18.0 million in a first lien term loan",
        terms=[("Credit (Debt Investment)", "$18.0 million"), ("Reserve", "$4.0 million")],
        deal_team=["Jane Doe", "Sam Roe"], date="April 2026")

    memo.divider(0)
    s = memo.slide("Acme", "Opportunity Summary")
    s.labeled_block([("Situation Overview", "Sponsor is acquiring ..."),
                     ("Company", "Acme, founded in 2009 ...")])
    s.source("Source: management, L.E.K. research")

    memo.save("Acme_Final_Approval_Memo.pptx")

Geometry is expressed in inches throughout, matching how the source decks were laid out.
"""

from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# --- EIP brand -------------------------------------------------------------------
# These are applied as direct run colors, not theme colors. The template's theme is
# stock Office; EIP's identity lives in the runs, so setting them explicitly is what
# makes a slide look like EIP rather than like PowerPoint.
GREEN = RGBColor(0x00, 0x5C, 0x2A)       # headlines, lead-in labels, emphasis
GREEN_ACCENT = RGBColor(0x00, 0x77, 0x42)  # separators, figures worth the eye
GREEN_DATE = RGBColor(0x2F, 0x85, 0x4C)
TITLE_GREY = RGBColor(0x91, 0x91, 0x91)  # the descriptive half of a title
BODY_GREY = RGBColor(0x59, 0x59, 0x59)   # body copy
MUTED_GREY = RGBColor(0x71, 0x71, 0x71)
SOURCE_GREY = RGBColor(0xBF, 0xBF, 0xBF)  # source and note lines
RED = RGBColor(0xC0, 0x00, 0x00)         # risks, downside cases
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RULE_GREY = RGBColor(0xD9, 0xD9, 0xD9)

FONT = "Nunito"

# Series colors for native charts, ordered so a two-series chart reads green vs grey.
CHART_COLORS = [GREEN, TITLE_GREY, GREEN_ACCENT, RGBColor(0x9D, 0xC3, 0xAE),
                RGBColor(0xBF, 0xBF, 0xBF), GREEN_DATE]

# --- canvas ----------------------------------------------------------------------
SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN_L, MARGIN_R = 0.23, 0.23
CONTENT_W = SLIDE_W - MARGIN_L - MARGIN_R      # 12.87
TITLE_T, TITLE_H = 0.24, 0.34
BODY_T = 0.68                                   # first usable row under the title
BODY_B = 7.05                                   # keep clear of the source line
SOURCE_T = 7.12
GUTTER = 0.26
COL_W = (CONTENT_W - GUTTER) / 2                # 6.30 -- the two-column workhorse
COL_R_L = MARGIN_L + COL_W + GUTTER

DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "eip-memo-template.pptx"


def _style(run, size=10, bold=False, color=BODY_GREY, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return run


class _Slide:
    """One memo slide. Methods place content on the EIP grid and return the shape."""

    def __init__(self, memo, slide):
        self._memo = memo
        self.slide = slide

    # -- primitives ----------------------------------------------------------------
    def textbox(self, left, top, width, height):
        box = self.slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.04)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        return box

    def title(self, bold_part, rest=""):
        """The EIP two-tone title: bold green subject, grey description.

        The grey half is where the slide earns its keep -- "Pros & Cons" says nothing,
        "Pros & Cons | Risks Mitigated by Proven Operating History" tells a reader the
        conclusion before they look at the table.
        """
        box = self.textbox(MARGIN_L, TITLE_T, CONTENT_W, TITLE_H)
        p = box.text_frame.paragraphs[0]
        _style(p.add_run(), 20, True, GREEN).text = bold_part + (" " if rest else "")
        if rest:
            _style(p.add_run(), 20, False, TITLE_GREY).text = rest
        return box

    def heading(self, text, left, top, width, size=14, color=GREEN, height=0.34):
        box = self.textbox(left, top, width, height)
        p = box.text_frame.paragraphs[0]
        _style(p.add_run(), size, True, color).text = text
        return box

    def labeled_block(self, items, left=MARGIN_L, top=BODY_T, width=CONTENT_W,
                      height=None, size=10, space_after=6):
        """Paragraphs of the form **Label**: explanation -- the memo's basic unit of prose.

        `items` is a list of (label, body) pairs; pass label=None for a paragraph with no
        lead-in. The label carries the claim and the body carries the evidence, which is
        what lets a reader skim the bold text and still follow the argument.
        """
        height = height or (BODY_B - top)
        box = self.textbox(left, top, width, height)
        tf = box.text_frame
        for i, (label, body) in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(space_after)
            if label:
                _style(p.add_run(), size, True, GREEN).text = label
                _style(p.add_run(), size, True, GREEN_ACCENT).text = ": "
            _style(p.add_run(), size, False, BODY_GREY).text = body
        return box

    def bullets(self, lines, left=MARGIN_L, top=BODY_T, width=CONTENT_W, height=None,
                size=10, color=BODY_GREY, space_after=4):
        """Bulleted lines. Use for lists of facts; use labeled_block for argument."""
        height = height or (BODY_B - top)
        box = self.textbox(left, top, width, height)
        tf = box.text_frame
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(space_after)
            _style(p.add_run(), size, False, color).text = "•  " + line
        return box

    def rule(self, x=None, top=1.05, bottom=BODY_B):
        """The thin vertical divider between two columns."""
        x = COL_R_L - GUTTER / 2 if x is None else x
        line = self.slide.shapes.add_connector(1, Inches(x), Inches(top), Inches(x), Inches(bottom))
        line.line.color.rgb = RULE_GREY
        line.line.width = Pt(0.75)
        return line

    def source(self, text):
        """Grey source/note line pinned to the bottom of the slide."""
        box = self.textbox(MARGIN_L, SOURCE_T, CONTENT_W, 0.26)
        p = box.text_frame.paragraphs[0]
        _style(p.add_run(), 9, False, SOURCE_GREY).text = text
        return box

    # -- composites ----------------------------------------------------------------
    def two_columns(self, left_heading=None, right_heading=None,
                    left_color=GREEN, right_color=GREEN, top=BODY_T):
        """Split the body into two columns with headings and a divider.

        Returns (left_x, right_x, content_top, col_width) so the caller can place
        anything into either side without recomputing the grid.
        """
        content_top = top
        if left_heading or right_heading:
            if left_heading:
                self.heading(left_heading, MARGIN_L, top, COL_W, 14, left_color)
            if right_heading:
                self.heading(right_heading, COL_R_L, top, COL_W, 14, right_color)
            content_top = top + 0.46
            self.rule(top=content_top - 0.06)
        else:
            self.rule(top=content_top)
        return MARGIN_L, COL_R_L, content_top, COL_W

    def pros_cons(self, thesis, risks, top=BODY_T):
        """The Investment Thesis / Key Risks & Mitigants slide.

        Risks are headed in red on purpose: an IC reader should be able to find the
        bear case without hunting, and a memo that buries it loses credibility.
        `thesis` and `risks` are lists of (label, body) pairs.
        """
        lx, rx, ct, cw = self.two_columns("Investment Thesis", "Key Risks & Potential Mitigants",
                                          GREEN, RED, top)
        self.labeled_block(thesis, lx, ct, cw, BODY_B - ct)
        self.labeled_block(risks, rx, ct, cw, BODY_B - ct)

    def commentary(self, text, left=None, top=BODY_T, width=None, height=None, heading="Commentary"):
        """The commentary rail that sits beside an exhibit.

        Every chart and table in an EIP memo is accompanied by the reading of it. The
        exhibit shows what happened; this box says why it matters.
        """
        left = COL_R_L if left is None else left
        width = COL_W if width is None else width
        height = (BODY_B - top) if height is None else height
        if heading:
            self.heading(heading, left, top, width, 11, GREEN, 0.26)
            top, height = top + 0.30, height - 0.30
        box = self.textbox(left, top, width, height)
        tf = box.text_frame
        paras = text if isinstance(text, (list, tuple)) else [text]
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.space_after = Pt(6)
            _style(p.add_run(), 10, False, BODY_GREY).text = para
        return box

    def kpi_tiles(self, tiles, left=MARGIN_L, top=BODY_T, width=CONTENT_W, height=0.95):
        """A row of headline figures: [(value, label), ...]."""
        n = len(tiles)
        gap = 0.14
        tw = (width - gap * (n - 1)) / n
        for i, (value, label) in enumerate(tiles):
            x = left + i * (tw + gap)
            shape = self.slide.shapes.add_shape(5, Inches(x), Inches(top), Inches(tw), Inches(height))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(0xF2, 0xF7, 0xF4)
            shape.line.color.rgb = RGBColor(0xDD, 0xE8, 0xE1)
            shape.shadow.inherit = False
            tf = shape.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            _style(p.add_run(), 20, True, GREEN).text = str(value)
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            _style(p2.add_run(), 9, False, BODY_GREY).text = label
        return None

    def table(self, rows, left=MARGIN_L, top=BODY_T, width=CONTENT_W, height=None,
              col_widths=None, size=10, header_size=11, align=None, first_col_left=True):
        """A styled table. `rows[0]` is the header.

        `col_widths` are relative weights (e.g. [2, 1, 4]); `align` is a per-column list
        of PP_ALIGN values. Numbers read best right- or centre-aligned and the first
        column left-aligned, which is the default.
        """
        nrows, ncols = len(rows), len(rows[0])
        height = height or min(BODY_B - top, 0.32 * nrows)
        shape = self.slide.shapes.add_table(nrows, ncols, Inches(left), Inches(top),
                                            Inches(width), Inches(height))
        table = shape.table
        table.first_row = True
        table.horz_banding = True

        if col_widths:
            total = sum(col_widths)
            for i, w in enumerate(col_widths):
                table.columns[i].width = Emu(int(Inches(width) * w / total))

        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                cell = table.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(0.06)
                cell.margin_top = cell.margin_bottom = Inches(0.03)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                if align:
                    p.alignment = align[c]
                elif c == 0 and first_col_left:
                    p.alignment = PP_ALIGN.LEFT
                else:
                    p.alignment = PP_ALIGN.CENTER
                if r == 0:
                    _style(p.add_run(), header_size, True, WHITE).text = str(val)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = GREEN
                else:
                    _style(p.add_run(), size, False, BODY_GREY).text = str(val)
        return shape

    def chart(self, chart_type, categories, series, left=MARGIN_L, top=BODY_T,
              width=COL_W, height=None, title=None, number_format='#,##0',
              show_values=True, legend=None):
        """A native PowerPoint chart, styled to the EIP palette.

        Native rather than an image so the deck stays editable -- an IC member who wants
        to see the underlying number can click into it, and the deal team can revise a
        forecast without regenerating a picture.

        `series` is a list of (name, values) pairs. `chart_type` accepts the friendly
        names 'column', 'bar', 'line', 'stacked_column', 'pie'.
        """
        kinds = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "stacked_column": XL_CHART_TYPE.COLUMN_STACKED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
        }
        height = height or (BODY_B - top)
        data = CategoryChartData()
        data.categories = categories
        for name, values in series:
            data.add_series(name, values)

        gf = self.slide.shapes.add_chart(kinds[chart_type], Inches(left), Inches(top),
                                         Inches(width), Inches(height), data)
        chart = gf.chart
        chart.font.name = FONT
        chart.font.size = Pt(9)
        chart.font.color.rgb = BODY_GREY

        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
            _style(chart.chart_title.text_frame.paragraphs[0].runs[0], 11, True, GREEN)
        else:
            chart.has_title = False

        show_legend = len(series) > 1 if legend is None else legend
        chart.has_legend = show_legend
        if show_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False

        for i, plot_series in enumerate(chart.series):
            color = CHART_COLORS[i % len(CHART_COLORS)]
            if chart_type == "line":
                plot_series.format.line.color.rgb = color
                plot_series.format.line.width = Pt(2.25)
            else:
                plot_series.format.fill.solid()
                plot_series.format.fill.fore_color.rgb = color

        if show_values and chart_type != "pie":
            plot = chart.plots[0]
            plot.has_data_labels = True
            labels = plot.data_labels
            labels.number_format = number_format
            labels.number_format_is_linked = False
            labels.font.size = Pt(8)
            labels.font.name = FONT
            labels.font.color.rgb = BODY_GREY

        if chart_type != "pie":
            for axis in (chart.category_axis, chart.value_axis):
                axis.format.line.color.rgb = RULE_GREY
                axis.tick_labels.font.size = Pt(9)
                axis.tick_labels.font.name = FONT
                axis.tick_labels.font.color.rgb = BODY_GREY
            chart.value_axis.has_major_gridlines = True
            chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xEE, 0xEE, 0xEE)
            chart.category_axis.has_major_gridlines = False
        return chart

    def image(self, path, left, top, width=None, height=None):
        kw = {}
        if width:
            kw["width"] = Inches(width)
        if height:
            kw["height"] = Inches(height)
        return self.slide.shapes.add_picture(str(path), Inches(left), Inches(top), **kw)


class Memo:
    """A memo deck under construction."""

    def __init__(self, company, memo_type="Final Approval", sections=None, template=None):
        self.company = company
        self.memo_type = memo_type
        self.sections = sections or [
            "Opportunity, Transaction, and Diligence Overview",
            "Company and Market Overview",
            "Financials",
            "Sponsor Overview",
            "Appendix",
        ]
        self.prs = Presentation(str(template or DEFAULT_TEMPLATE))
        self._layouts = {l.name: l for l in self.prs.slide_masters[0].slide_layouts}

    # -- slide creation ------------------------------------------------------------
    def _blank(self):
        """A blank slide with the layout's empty date/footer placeholders removed.

        They inherit from the master and would otherwise stamp a stale date on every
        slide; the slide-number field on the master is kept.
        """
        slide = self.prs.slides.add_slide(self._layouts["Blank"])
        for shape in list(slide.placeholders):
            if shape.placeholder_format.idx in (10, 11):  # DATE, FOOTER
                shape._element.getparent().remove(shape._element)
        return slide

    def slide(self, title_bold, title_rest=""):
        """A titled content slide -- the workhorse. Returns a _Slide."""
        s = _Slide(self, self._blank())
        s.title(title_bold, title_rest)
        return s

    def blank_slide(self):
        """An untitled slide, for full-bleed exhibits."""
        return _Slide(self, self._blank())

    def title_slide(self, headline, terms=None, deal_team=None, sourced=None,
                    date=None, subtitle=None):
        """The cover: what EIP is being asked to approve, stated in one sentence.

        `headline` should name the amount and the instrument, because that is the single
        fact every reader wants first. `terms` is a list of (label, value) pairs.
        """
        s = _Slide(self, self._blank())

        box = s.textbox(MARGIN_L, 1.15, CONTENT_W, 0.9)
        p = box.text_frame.paragraphs[0]
        _style(p.add_run(), 28, True, GREEN).text = f"{self.company} "
        _style(p.add_run(), 28, False, TITLE_GREY).text = f"{self.memo_type} Memo"
        if subtitle:
            p2 = box.text_frame.add_paragraph()
            _style(p2.add_run(), 13, False, MUTED_GREY).text = subtitle

        hb = s.textbox(MARGIN_L, 2.15, CONTENT_W, 0.7)
        _style(hb.text_frame.paragraphs[0].add_run(), 13, True, GREEN_ACCENT).text = headline

        if terms:
            s.labeled_block(terms, MARGIN_L, 2.95, CONTENT_W, 2.0, size=11, space_after=4)

        line = s.slide.shapes.add_connector(1, Inches(MARGIN_L), Inches(5.19),
                                            Inches(SLIDE_W - MARGIN_R), Inches(5.19))
        line.line.color.rgb = RULE_GREY
        line.line.width = Pt(1)

        if deal_team:
            tb = s.textbox(MARGIN_L, 5.40, 9.6, 1.2)
            tp = tb.text_frame.paragraphs[0]
            _style(tp.add_run(), 11, True, MUTED_GREY).text = "Deal Team:   "
            _style(tp.add_run(), 11, False, MUTED_GREY).text = ",  ".join(deal_team)
            if sourced:
                sp = tb.text_frame.add_paragraph()
                _style(sp.add_run(), 11, False, MUTED_GREY).text = f"Sourced: {sourced}"

        if date:
            db = s.textbox(MARGIN_L, 6.70, CONTENT_W, 0.35)
            _style(db.text_frame.paragraphs[0].add_run(), 14, True, GREEN_DATE).text = date
        return s

    def divider(self, active_index):
        """The recurring section-divider slide, with the current section picked out.

        Repeating the whole contents list on every divider is deliberate: a memo is read
        in pieces, often out of order, and the divider re-orients whoever just opened it.
        """
        s = _Slide(self, self._blank())
        box = s.textbox(1.1, 2.1, SLIDE_W - 2.2, 3.4)
        roman = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
        for i, name in enumerate(self.sections):
            p = box.text_frame.paragraphs[0] if i == 0 else box.text_frame.add_paragraph()
            p.space_after = Pt(14)
            active = i == active_index
            _style(p.add_run(), 20 if active else 17, active, GREEN if active else TITLE_GREY
                   ).text = f"Section {roman[i]}.  {name}"
        return s

    def save(self, path):
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(path))
        return path
