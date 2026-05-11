# Radar — Prioritized Feature Backlog

> Generated: 2026-05-11 | References: `docs/jobs-to-be-done.md`, `docs/personas.md`

---

## Framework Selection

**Context assessment:**
| Dimension | Radar's Situation |
|---|---|
| Product stage | Post-PMF, growing adoption — ~16 active teams, internal tool with clear users |
| Team context | Small (1–2 engineers), low stakeholder complexity, single product owner |
| Decision needs | Filtering a manageable backlog (~10 items), not building consensus across orgs |
| Data availability | Limited usage metrics; good qualitative signal from daily use patterns |

**Recommended framework: ICE (Impact × Confidence × Ease)**

Rationale: RICE requires reliable reach/usage data we don't have. ICE is the right fit for a small-team internal tool in the post-PMF growth stage — it's lightweight, fast to apply, and forces the three questions that matter most right now: *how much does it help users? how sure are we? how hard is it to build?*

Scores are 1–10. ICE = Impact × Confidence × Ease. Use as directional input, not absolute truth.

---

## Scored Backlog

| # | Feature | Impact | Confidence | Ease | ICE Score | Tier |
|---|---|---|---|---|---|---|
| 1 | Advertiser health score / churn prediction | 9 | 7 | 5 | **315** | Now |
| 2 | Multi-CSV comparison (period over period) | 9 | 8 | 6 | **432** | Now |
| 3 | Custom alert thresholds per advertiser | 8 | 8 | 7 | **448** | Now |
| 4 | PDF report export | 7 | 9 | 7 | **441** | Now |
| 5 | Email digest of weekly highlights | 7 | 7 | 6 | **294** | Next |
| 6 | Team-level dashboards with aggregated metrics | 8 | 6 | 4 | **192** | Next |
| 7 | Mobile-responsive layout | 6 | 8 | 5 | **240** | Next |
| 8 | Annotation / notes on trends | 5 | 7 | 8 | **280** | Next |
| 9 | Scheduled CSV auto-import | 7 | 5 | 3 | **105** | Later |
| 10 | API integration with Yektanet's CRM | 8 | 4 | 2 | **64** | Later |

---

## Rationale by Feature

### 1. Multi-CSV Comparison — ICE: 432 (Now)
**Impact 9:** Directly addresses the top underserved job: tracking portfolio performance over time. Every user faces this gap daily.
**Confidence 8:** The need is unambiguous from the problem statements; comparison is the missing primitive.
**Ease 6:** Requires a data persistence layer but the comparison UI can be built on top of existing chart components.
*Builds the foundation for health scores and team dashboards.*

### 2. Custom Alert Thresholds — ICE: 448 (Now)
**Impact 8:** Transforms reactive scanning into proactive monitoring. High daily workflow impact for Sara persona.
**Confidence 8:** Alerts already exist in the app; this extends them with per-advertiser configuration rather than global thresholds.
**Ease 7:** Alert engine already built — adding per-advertiser config is incremental, not a rewrite.
*Quick win that makes the existing alerts feature actually useful.*

### 3. PDF Report Export — ICE: 441 (Now)
**Impact 7:** Directly unblocks the "prepare for client conversations" job. Ali persona's highest pain point.
**Confidence 9:** The need is explicit and universal across account managers.
**Ease 7:** Libraries like `jsPDF` or `html2canvas` can render existing UI components. No new data logic needed.
*High confidence, low risk, immediate user delight.*

### 4. Advertiser Health Score — ICE: 315 (Now)
**Impact 9:** The single most-requested capability pattern across all personas.
**Confidence 7:** We understand the need clearly; the algorithm design (what signals to weight) needs validation.
**Ease 5:** Requires defining and tuning a scoring model; risk of low trust if the score feels opaque.
*Start with a transparent, simple formula (spend trend + share change + session frequency) before adding ML.*

### 5. Email Digest of Weekly Highlights — ICE: 294 (Next)
**Impact 7:** Extends Radar's value outside the app — surfaces insights even when users don't log in.
**Confidence 7:** High value in theory; actual open/engagement rates unknown without testing.
**Ease 6:** Requires email infrastructure (SMTP/sendgrid) and a scheduled job — new operational surface area.

### 6. Annotations / Notes on Trends — ICE: 280 (Next)
**Impact 5:** Nice-to-have for power users tracking their own context alongside data.
**Confidence 7:** Low risk of building the wrong thing; clear bounded scope.
**Ease 8:** Purely front-end, localStorage-backed. Very fast to ship.
*Low effort, delivers disproportionate value to the daily user who wants to leave themselves notes.*

### 7. Mobile-Responsive Layout — ICE: 240 (Next)
**Impact 6:** Moderate — users primarily work at desktops; mobile is for quick lookups between meetings.
**Confidence 8:** RTL + data-dense layout is genuinely hard to use on small screens today.
**Ease 5:** Significant CSS refactoring across a large component tree; risk of visual regressions.

### 8. Team-Level Dashboards — ICE: 192 (Next)
**Impact 8:** Reza persona's highest-priority need; strategic value for leadership.
**Confidence 6:** The data model (CSV is advertiser-level, not team-level) requires careful design.
**Ease 4:** Requires a new aggregation data model and a separate dashboard screen.
*High impact but non-trivial — sequence after multi-CSV comparison lays the data foundation.*

### 9. Scheduled CSV Auto-Import — ICE: 105 (Later)
**Impact 7:** Would eliminate daily manual upload friction entirely.
**Confidence 5:** Requires Yektanet's data pipeline to expose a stable API or SFTP endpoint — external dependency.
**Ease 3:** Backend infrastructure, auth, scheduling, error handling. Significant scope.

### 10. API Integration with CRM — ICE: 64 (Later)
**Impact 8:** Highest strategic value long-term — linking ad performance to sales outcomes.
**Confidence 4:** CRM structure, data ownership, and integration contracts are unknown.
**Ease 2:** Cross-system integration with unknown API, auth complexity, data mapping work.
*High potential but too many unknowns. Validate the CRM API surface before estimating.*

---

## Recommended Sequencing

```
Now (Q2 2026):
  ├── Custom alert thresholds (quick win, existing infra)
  ├── PDF report export (unblocks client conversations immediately)
  ├── Multi-CSV comparison (foundational — unlocks team dashboards later)
  └── Advertiser health score v1 (simple transparent formula, iterate)

Next (Q3 2026):
  ├── Annotations / notes (low effort, high daily-user value)
  ├── Email weekly digest (expand reach beyond active sessions)
  ├── Mobile-responsive layout (accessibility and field use)
  └── Team-level dashboards (builds on multi-CSV comparison foundation)

Later (Q4 2026+):
  ├── Scheduled CSV auto-import (needs external data pipeline work)
  └── CRM API integration (validate feasibility first)
```

---

*Scores are directional — re-score quarterly as usage data and team capacity change.*
