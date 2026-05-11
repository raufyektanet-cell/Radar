# Radar — Product Roadmap (Q2–Q4 2026)

> Generated: 2026-05-11 | Living document — revisit at start of each quarter
> References: `docs/prioritized-backlog.md`, `docs/jobs-to-be-done.md`, `docs/problem-statements.md`

---

## Strategic Goals

| Goal | Metric | Baseline | Target |
|---|---|---|---|
| Reduce advertiser churn | At-risk advertisers flagged before they churn | 0% proactive | >60% caught proactively |
| Increase Yektanet revenue | Sales manager prep time → more client conversations | ~45 min/day on data prep | <15 min/day |
| Better client conversations | Account managers using Radar data in client calls | Unknown | >80% of AMs weekly |

---

## Now — Q2 2026 (May–July)

*High-confidence, high-impact work that builds on existing infrastructure.*

---

### Initiative 1: Custom Alert Thresholds

**Hypothesis:** If we let sales managers configure per-advertiser alert thresholds (e.g., "notify me when Yektanet share drops >10 pts"), they will catch at-risk advertisers before churn events — reducing the reactive scanning that currently takes 30–45 minutes per day.

**Success metric:** ≥50% of active users configure at least one custom alert within 30 days of launch; "time spent scanning alerts manually" drops in qualitative feedback.

**Effort:** S (alerts engine already exists; adding per-advertiser config is incremental)

**Dependencies:** None — builds directly on existing AlertsScreen component

**Outcome connection:** Addresses Problem Statement 1 (at-risk triage), JTBD Job 1 (monitor advertiser health)

---

### Initiative 2: PDF / Image Report Export

**Hypothesis:** If account managers can export a single-advertiser market brief as a PDF, they will use Radar data directly in client conversations — replacing manual slide preparation and reducing prep time per client meeting from ~30 min to <5 min.

**Success metric:** PDF export used in ≥30% of sessions within 30 days of launch; qualitative confirmation from AMs that they shared it with a client.

**Effort:** S–M (html2canvas/jsPDF on existing UI components; no new data logic)

**Dependencies:** Design: decide which view(s) to include in export (trends chart + market share table + summary stats)

**Outcome connection:** Addresses Problem Statement 2 (client-ready output), JTBD Job 2 (prepare for client conversations)

---

### Initiative 3: Multi-CSV Comparison (Period-over-Period)

**Hypothesis:** If users can upload a second CSV and compare it side-by-side with the current dataset, they will be able to track portfolio trends over time — eliminating the need to manually reconstruct history from saved reports.

**Success metric:** ≥40% of active users use the comparison view within 60 days; "I can't see trends over time" removed from top-5 complaints.

**Effort:** M (requires persisting a second dataset in state; new comparison UI layer on top of existing charts)

**Dependencies:** Architecture decision: store second CSV in localStorage or memory only? (localStorage recommended for persistence across sessions)

**Outcome connection:** Addresses JTBD Job 4 (track portfolio performance over time); foundational for Team Dashboards in Q3

---

### Initiative 4: Annotations / Notes on Trends

**Hypothesis:** If users can attach a short note to an advertiser (e.g., "called client 5/10, pausing for 2 weeks"), they will reduce context-switching between Radar and external notes tools — improving their ability to act on insights days later.

**Success metric:** ≥25% of active users create at least one annotation within 30 days of launch.

**Effort:** S (localStorage-backed, front-end only; no backend required)

**Dependencies:** None

**Outcome connection:** Addresses Ali persona's need for personal context alongside data

---

## Next — Q3 2026 (August–October)

*High-value work that requires design, infrastructure, or the Q2 foundation to be in place.*

---

### Initiative 5: Advertiser Health Score v1

**Hypothesis:** If Radar surfaces a simple, transparent health score per advertiser (based on spend trend + Yektanet share change + session frequency over 30 days), sales managers will open Radar in the morning knowing immediately which 2–3 advertisers need attention — without manual scanning.

**Success metric:** ≥70% of users interact with the health score view in their first session; qualitative feedback that "it surfaces the right advertisers."

**Effort:** M (algorithm design + new ranked view; scoring logic straightforward but must be explainable)

**Dependencies:** Multi-CSV comparison (Initiative 3) provides historical delta needed for trend component of score; custom alerts (Initiative 1) can be seeded from health score thresholds

**Outcome connection:** Addresses Problem Statement 1 root cause; JTBD Opportunity 1 (proactive health monitoring)

**Risk:** Low trust if scoring feels like a black box. Mitigate by showing the formula inline: "سهم یکتانت ↓۱۲٪ در ۳۰ روز — امتیاز: ۳۵/۱۰۰"

---

### Initiative 6: Weekly Email Digest

**Hypothesis:** If sales managers receive a weekly Persian-language email summarizing their top 5 at-risk advertisers and key market movements, they will engage with Radar insights even on days they don't log in — increasing overall retention and proactive follow-up rate.

**Success metric:** ≥60% open rate; ≥30% of email recipients click through to Radar within 24 hours.

**Effort:** M (requires email infrastructure: SMTP/SendGrid, scheduled job, email template in Persian)

**Dependencies:** Health score (Initiative 5) provides the ranked list of at-risk advertisers; requires server-side scheduling capability

**Outcome connection:** Extends Radar's value outside the session; addresses Sara persona's aspiration to catch churn before it happens

---

### Initiative 7: Team-Level Dashboard

**Hypothesis:** If department heads can see an aggregated view of their team's advertiser portfolio (total sessions, avg Yektanet share, at-risk count by team), they will be able to identify systemic risks and intervene earlier — reducing their reliance on managers to surface problems reactively.

**Success metric:** Reza-persona users (team leads) log in weekly; qualitative feedback that "I can now see my team's health without asking managers."

**Effort:** L (requires data aggregation layer; team-to-advertiser mapping; new dashboard screen)

**Dependencies:** Multi-CSV comparison (Initiative 3) for historical context; teams.ts already defines team/manager structure — can be used to filter advertiser data by team ownership

**Outcome connection:** Addresses Problem Statement 3 (cross-team visibility), JTBD Job 4 (track portfolio performance)

---

### Initiative 8: Mobile-Responsive Layout

**Hypothesis:** If Radar renders correctly on mobile, account managers will be able to pull up a quick market share number between client meetings — reducing the "let me follow up" moments in client conversations.

**Success metric:** Mobile session share increases from ~0% to ≥15%; qualitative reports of mobile use in the field.

**Effort:** M (significant RTL + data-dense CSS refactoring; risk of visual regressions)

**Dependencies:** None, but sequence after major UI changes from Q2 to avoid re-doing work

**Risk:** The existing component tree is desktop-first and complex. Consider a dedicated mobile "quick view" screen rather than full responsive overhaul.

---

## Later — Q4 2026+

*Strategic bets with significant dependencies or unknowns. Validate feasibility before committing.*

---

### Initiative 9: Scheduled CSV Auto-Import

**Hypothesis:** If Radar can pull fresh data automatically from Yektanet's data pipeline (via API or SFTP), it will eliminate the daily manual upload friction entirely — making Radar a real-time intelligence tool rather than a batch analysis tool.

**Effort:** L (requires Yektanet data pipeline integration, auth, error handling, scheduling infrastructure)

**Dependencies:** Yektanet engineering to expose a stable data API or SFTP endpoint — external dependency, unknown timeline

**Prerequisite action:** Spike with data engineering team to assess feasibility and timeline before adding to roadmap

---

### Initiative 10: CRM API Integration

**Hypothesis:** If Radar can pull advertiser relationship data (contact history, deal stage, last contact date) from Yektanet's CRM, account managers will have a single view combining ad performance and relationship health — enabling proactive outreach before advertisers churn.

**Effort:** L (cross-system integration, unknown CRM API surface, data mapping)

**Dependencies:** CRM API access, data ownership decisions, privacy/security review

**Prerequisite action:** Discovery sprint — identify CRM system in use, API availability, data model, and whether integration is organizationally feasible

---

## Dependency Map

```
Q2 (Now)                    Q3 (Next)                   Q4 (Later)
─────────────────────────   ─────────────────────────   ─────────────────
Custom Alerts ──────────────► Health Score v1
Multi-CSV Comparison ───────► Team Dashboard
                            ► Weekly Email Digest ───────► Auto-Import
PDF Export                  Mobile Layout               CRM Integration
Annotations
```

---

## What We're NOT Doing (and Why)

| Excluded | Reason |
|---|---|
| Gamification (badges, streaks) | No evidence this is a user need; distraction from core workflow |
| Public advertiser profiles | Privacy and competitive sensitivity concerns |
| Multi-language support beyond Persian | Internal tool; Persian-first is correct priority |
| Real-time data streaming | Batch daily is sufficient for current use case; complexity not justified |

---

## Review Cadence

- **Monthly:** Review initiative progress, update confidence/effort estimates
- **Quarterly:** Re-prioritize based on usage data, user feedback, and new discoveries
- **Before each initiative kick-off:** Run the relevant skill (epic-hypothesis, user-story) to validate before building

---

*This roadmap is a hypothesis, not a contract. Initiatives should be validated with lightweight experiments before full build-out.*
