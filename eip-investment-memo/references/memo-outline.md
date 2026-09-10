# Memo slide inventory

Derived from EIP Final Approval memos. Treat it as the menu, not the mandate: include a
slide when the deal and the diligence support it, drop it when they don't, and add
deal-specific slides freely. A memo padded with thin slides reads as weaker than a
shorter one where every page carries weight.

Section numbering and the divider slide before each section are conventions worth
keeping — the memo gets read in pieces and out of order, and the dividers re-orient
whoever just opened it.

## Contents

- [Final Approval memo](#final-approval-memo)
- [Preliminary Approval memo](#preliminary-approval-memo)
- [Slide patterns](#slide-patterns)

---

## Final Approval memo

Typically 35–60 slides across five sections plus appendix.

### Cover

One slide. States what is being approved: EIP's total commitment, split by instrument,
with the deal structure beneath it. Also carries deal team, sourcing attribution, and
date. A reader who sees only this slide should know the ask.

### Section I — Opportunity, Transaction, and Diligence Overview

| Slide | Contents | Notes |
|---|---|---|
| Updates Since Preliminary Approval | Prelim approval date, IC votes by member, diligence completed since, material changes to terms or performance | Final-only. The first question IC asks is what changed |
| Due Diligence Conducted | Workstream table: area, provider, status, key finding | Shows the diligence was real |
| Customer Call Summaries | Table per call: customer, contact and role, score, key insights | Usually 2–3 slides. Score is the reference's strength of endorsement |
| Expert Call Summaries | Same format, third-party experts (GLG or similar) | |
| [Company] Opportunity Summary | Situation overview, company description, sponsor, equity co-investors, strategic relevance to EIP, deal allocation | The single densest slide in the memo |
| Pros & Cons | Investment thesis (left, green) vs key risks and potential mitigants (right, red) | Title's grey half should state the overall risk conclusion |
| Sources & Uses and PF Capitalization | Sources, uses, pro forma cap structure with leverage multiples | |
| Terms, SBA Compliance and Returns Analysis | Term sheet summary, SBA/SBIC eligibility, base returns | |
| Covenants | Financial covenants with headroom vs base case | |
| Revolver / Other Lender Terms | Only when there is a third-party working capital facility | |
| Organization & Equity Structure | Legal structure diagram and common equity cap table | |

### Section II — Company and Market Overview

| Slide | Contents |
|---|---|
| Company Overview | Founding, HQ, headcount, what it does, end markets, why it wins. Often runs to a second "Continued" slide |
| Service Offerings | Each line of business with % of revenue and gross margin |
| Select Project Case Studies | Two or three representative projects with outcomes |
| Company History & Evolution | Timeline of inflections — founding, expansions, pivots |
| Leadership Overview | Each executive: role, tenure, background, what they own |
| Organization / Footprint | Org chart, offices, geographic coverage |
| Top Customers | Customer, tenure, revenue, % of total, contract expiry |
| Customer & Project Analysis | Mix shifts by segment, project type, end market |
| Customer Splits by Contract Win Year | Cohort view — are new customers replacing churned ones |
| Contract Gains and Fades by Vintage | Whether contracts grow or shrink after award |
| Wins & Backlog Trending | Quarterly bookings and backlog, with the methodology note |

### Section III — Financials

| Slide | Contents |
|---|---|
| Quality of Earnings Summary | Provider, scope, conclusions, diligence-adjusted EBITDA |
| EBITDA Adjustments | Reported-to-adjusted bridge, each adjustment lettered and explained |
| Historical Financials | Three-plus years: revenue, gross profit, adj. EBITDA, margins, working capital |
| Interim / YTD Comparison | Current-year actuals vs budget and prior year |
| Forecast Case Narrative | What distinguishes each case — the assumptions, not the outputs |
| Financial KPIs | Revenue and margin trend exhibits |
| Financial Forecast — Sponsor Base Case | Sponsor's model as delivered |
| Financial Forecast — Sponsor Lender Case | Sponsor's downside |
| Financial Forecast — EIP Base Case | EIP's own view |
| Financial Forecast — EIP Flat Case | No growth from current run-rate |
| Financial Forecast — EIP Downside Case | The case that tests the covenants |
| Returns Sensitivity | IRR and MOIC grids across the driver pair that matters |

The case set is the analytical core of a credit memo. Each case gets the same table
shape so cases can be compared line by line, and the downside case must be genuinely
adverse — a downside that still comfortably clears covenants has not been stress-tested.

### Section IV — Sponsor Overview

| Slide | Contents |
|---|---|
| Sponsor Overview | Firm, founding, AUM, strategy, select investments, portfolio |
| Sponsor Vision / Value-Add | What they intend to do with this asset |
| Principal Track Record | Deal-by-deal history of the lead partner |
| Third-Party Equity Investors | Co-investors, their history with EIP |

### Section V — Appendix

Pull forward anything IC will ask about but that would break the argument's flow:
market overview and sizing, industry M&A multiple trends, valuations by service line and
firm size, public comparables, precedent transactions, customer capex history, cohort
analyses, bonding exposure, drivers of margin change, ESG assessment, third-party market
study excerpts, voice-of-customer survey results, and additional expert call summaries.

---

## Preliminary Approval memo

15–25 slides. Same template and conventions, narrower scope: the question is whether to
spend real diligence money, not whether to fund.

Keep:

- Cover, with preliminary terms
- Opportunity Summary
- Pros & Cons
- Preliminary Sources & Uses and structure
- Company Overview, service offerings, leadership
- Top customers and revenue durability
- Historical financials and management forecast
- Sponsor overview
- Market overview

Replace *Updates Since Preliminary Approval* and *Due Diligence Conducted* with a
**Proposed Diligence Plan**: what will be diligenced, by whom, at what cost, and which
open questions each workstream is meant to close. That slide is the actual ask.

Omit the full forecast case set, covenant detail, QofE summary, and call summaries —
that work hasn't happened yet. Saying so plainly is better than a thin version of it.

---

## Slide patterns

Four shapes cover most of the deck. Matching content to the right one is most of what
makes a memo readable.

**Prose slide** — `labeled_block` across the full width, or across ~8.6in with a
`commentary` rail at right. For opportunity summaries and narrative sections.

**Exhibit + commentary** — chart or table on the left (~7.6in), commentary at right
(~5.0in). Every exhibit in an EIP memo is accompanied by its reading: the exhibit shows
what happened, the commentary says why it matters. An exhibit alone leaves the
interpretation to the reader, which for a downside case is exactly the wrong thing.

**Split comparison** — `pros_cons`, or `two_columns` for any left/right contrast.

**Full-width table** — top customers, covenants, terms, call summaries. Keep to roughly
12 rows per slide; continue onto "(2/3)" slides rather than shrinking the type.
