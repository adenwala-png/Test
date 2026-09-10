---
name: eip-investment-memo
description: "Build EIP investment approval memos as branded PowerPoint decks from diligence inputs — call notes and transcripts, sponsor and management decks, QofE reports, financial models in Excel, and market research. Use this skill whenever the user mentions an investment memo, IC memo, approval memo, preliminary or final approval, a credit memo, a diligence write-up, or asks to turn deal diligence into slides — and also when they hand over deal materials and ask to 'put together the deck', 'write this up for IC', or 'summarize diligence for the committee' without naming a format. Also use it to update, extend, or restructure an existing EIP memo, to draft individual memo sections (opportunity summary, pros and cons, forecast cases, sponsor overview), or to check a memo for gaps before circulation."
---

# EIP investment approval memos

An EIP memo is the document an investment committee approves a transaction from. That
purpose sets everything else: a reader needs to reach a decision, find the bear case
without hunting for it, and trust every number on the page. A memo that reads well but
contains one figure nobody can source is worse than useless, because it spends the deal
team's credibility.

You are assembling that document from raw diligence. The work is synthesis, not
transcription — the inputs are long and repetitive, and the memo is the argument that
survives after they are read.

## Two memo types

EIP approves in two stages, and the type determines what the deck must contain.

| | **Preliminary Approval** | **Final Approval** |
|---|---|---|
| Question it answers | Should we spend real diligence money on this? | Should we fund it? |
| Typical length | 15–25 slides | 35–60 slides |
| Diligence state | Management calls, sponsor materials, desk research | Completed QofE, customer and expert calls, market study, final terms |
| Distinguishing content | Preliminary terms, thesis, key risks, diligence plan | **Updates Since Preliminary Approval**, call summaries, final terms and covenants, full forecast case set |

Ask which one if the user hasn't said and it isn't obvious. The tell for a Final Approval
memo is that preliminary approval already happened — if the user mentions IC votes, a
prior approval date, or "what's changed since," it is Final. If they are still deciding
whether to pursue the deal, it is Preliminary.

If the ask is smaller than either — "just the pros and cons slide," "redo the forecast
cases" — build that section only, on the same template, and don't pad it into a full memo.

## Workflow

### 1. Inventory the inputs before reading them closely

List what you have and what it can support. Call notes give you customer sentiment and
qualitative risk; a QofE gives you defensible EBITDA; a model gives you the forecast
cases; market research gives you TAM and competitive framing. Knowing which input backs
which slide up front stops you from later writing a claim you cannot source.

Note what is *missing* too. Absent inputs become the diligence-gaps section, which is a
real deliverable rather than an admission of failure — IC would rather see a named open
item than a confident sentence covering it up.

### 2. Read for the argument, not for coverage

Extract into a working notes file as you go, and record where each fact came from
(`file, page/tab/speaker`). You will need those attributions for source lines, and they
are what lets you tell a management assertion from a verified figure.

What you are looking for:

- **The transaction**: purchase price, entry multiple, instrument, EIP's amount and
  position, sources and uses, sponsor and co-investors.
- **The earnings base**: reported vs adjusted EBITDA, what the adjustments are, and who
  diligenced them. Adjusted EBITDA with unexplained bridges is the single most common
  weak point in a memo.
- **Revenue durability**: contracted vs repeat vs one-time, backlog, customer tenure,
  concentration, renewal history.
- **The bear case**: what has to be true for this to work, and what breaks it.

### 3. Plan the slide inventory, then confirm it

Draft the section-by-section slide list from `references/memo-outline.md`, adapted to
this deal and to what the inputs can actually support. Show it to the user before
building. A 45-slide deck is expensive to rebuild, and a two-minute check on the outline
catches "we don't present it that way" long before it becomes forty slides of rework.

### 4. Build the deck

Use `scripts/eip_deck.py`. It owns the EIP grid, the Nunito type, the green/grey palette
and the slide archetypes, so you write content and not coordinates:

```python
import sys; sys.path.insert(0, "scripts")
from eip_deck import Memo

memo = Memo(company="Acme Utility Services", memo_type="Final Approval")
memo.title_slide(headline="EIP to invest up to $22.0 million ...",
                 terms=[("Credit (Debt Investment)", "$18.0 million")],
                 deal_team=["A. Rivera"], date="April 2026")
memo.divider(0)
s = memo.slide("Acme", "Opportunity Summary")
s.labeled_block([("Situation Overview", "..."), ("Company", "...")], width=8.6)
s.commentary("...", left=9.05, width=4.05)
s.source("Source: Management, CBIZ Quality of Earnings")
memo.save("Acme_Final_Approval_Memo.pptx")
```

Full API, geometry constants and palette: `references/deck-api.md`. Read it before
placing anything by hand — most of what you would hand-position already has a method,
and the ones that don't have named grid constants to sit on.

Build in section order and save as you go, so a failure late in a long deck doesn't cost
the whole run.

### 5. Check before you hand it over

```bash
python scripts/check_deck.py memo.pptx
```

This catches what nobody sees until the deck is opened: text overrunning its box, shapes
off the canvas, tables longer than the slide, off-brand fonts and colors. Fix every
error. There is usually no working renderer available, so this check is standing in for
the human eye — treat a clean report as necessary, not sufficient, and say so when you
hand the deck over.

To read the deck back — either to proof what you just built, or to see what is already in
a memo you have been asked to update:

```bash
python scripts/deck_to_markdown.py memo.pptx -o memo.md
```

### 6. Report what you built and what you couldn't

Tell the user, briefly: the slide count by section, which claims rest on management
assertion rather than diligenced fact, and every gap you left. Do not quietly drop a
planned slide because the input was thin — an empty slide flagged as empty is
information; a silently missing one is a hole in the argument nobody knows about.

## Writing the memo

**Titles carry the conclusion.** Every content slide has a two-tone title: bold green
subject, grey descriptive half. The grey half is where the slide earns its place —
"Pros & Cons" tells a reader nothing, "Pros & Cons | Risks Mitigated by Long-Dated MSAs
and a Tenured Field Workforce" tells them the finding before they read the table. If you
cannot write a grey half that says something, the slide probably has no point yet.

**Lead with the label.** Body copy is `**Label**: explanation` (`labeled_block` renders
it). The label carries the claim, the body the evidence, so a reader skimming bold text
still follows the argument.

**Never invent a number.** Every figure traces to an input. If a number is needed and
absent, write the label with `[TBD — source]` and list it in the diligence gaps rather
than interpolating something plausible. A fabricated figure in an IC memo is not a
cosmetic error: it can be relied on in a funding decision, and it is nearly undetectable
once it is formatted like every other number on the page. This matters more than
finishing every slide.

**Attribute honestly.** "Management reports" and "per the CBIZ QofE" are different
claims with different weight, and the source line at the bottom of each exhibit should
make clear which one you are making.

**Give risks their due.** Each risk gets a real mitigant or an honest "no mitigant
identified." A risks section made of strawmen is transparent to an IC and costs the
whole memo its credibility.

More on prose conventions, section-by-section: `references/writing.md`.

## Regenerating the template

`assets/eip-memo-template.pptx` carries EIP's master, layouts, theme and logo, with no
deal content. If EIP's branding changes, re-derive it from a current memo:

```bash
python scripts/make_template.py NEWEST_MEMO.pptx assets/eip-memo-template.pptx \
  --deny CompanyName SponsorName
```

It strips slides, author names, revision history, SharePoint and add-in metadata, the
cached slide-title list, and the target company's logo off the master — then scans the
result and refuses to write a template that still names the source deal. Pass the deal
and sponsor names to `--deny` so that scan is meaningful.

**Never commit a real memo, or a template derived from one, without running that scan.**
These decks contain live transaction terms and counterparty names.

## References

- `references/memo-outline.md` — slide-by-slide inventory for both memo types
- `references/deck-api.md` — `eip_deck.py` API, grid geometry, palette
- `references/writing.md` — prose conventions and evidence discipline

## Scripts

| Script | Purpose |
|---|---|
| `scripts/eip_deck.py` | The builder. Import it; don't place shapes by hand |
| `scripts/check_deck.py` | Layout QA — run before handing anything over |
| `scripts/deck_to_markdown.py` | Read a deck back as text |
| `scripts/make_template.py` | Re-derive the template from a current memo |
