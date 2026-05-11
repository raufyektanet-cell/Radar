# Radar — Problem Statements

> Grounded in the proto-personas from `docs/personas.md` and codebase analysis. These statements guide discovery and feature validation — not solution design.

---

## Problem Statement 1: At-Risk Advertiser Triage

```
I am:      A sales manager overseeing 20–40 advertisers at Yektanet,
           responsible for catching churn before it affects revenue.

Trying to: Start each morning knowing exactly which 2–3 advertisers
           need my attention today — before I open a single spreadsheet.

But:       I have to manually upload yesterday's CSV, wait for it to
           parse, then scroll through every advertiser card one by one
           to spot declining trends or unusual competitor activity.

Because:   Radar has no proactive health scoring or ranked priority
           view — it surfaces all advertisers equally, leaving the
           identification work entirely to me.

Which makes me feel: Anxious that I'll miss a critical churn signal
           buried in the noise, and frustrated that a tool meant to
           save time still requires the same mental effort as a raw
           spreadsheet.
```

**Summary sentence:** A sales manager responsible for dozens of advertisers can't start her day with confidence because Radar treats every advertiser as equally important, forcing her to manually scan all data rather than act on what matters most.

---

## Problem Statement 2: Client-Ready Market Intelligence

```
I am:      An account manager preparing for a client check-in call,
           needing to show competitive positioning data that builds
           trust and justifies the client's Yektanet investment.

Trying to: Walk into a client meeting with a clear, credible narrative:
           "Here's your market share, here's where your competitors
           stand, and here's why now is the right time to increase spend."

But:       The data I need is spread across Radar's heatmap, the trends
           chart, and raw CSV numbers — none of it formatted in a way
           I could show a client directly, and exporting gives me an
           unstyled CSV that I still have to turn into a slide.

Because:   Radar is built for internal analysis, not client-facing
           communication — there's no shareable snapshot, PDF summary,
           or presentation-ready output that a non-technical client
           could read.

Which makes me feel: Underprepared and unprofessional when I have to
           say "let me follow up with the numbers" instead of showing
           them live, and like I'm leaving value on the table in every
           client conversation.
```

**Summary sentence:** An account manager sitting on valuable competitive intelligence can't use it in client conversations because Radar produces analysis for internal eyes only, creating a gap between the insight and the moment it could actually change a client's decision.

---

## Problem Statement 3: Cross-Team Performance Visibility

```
I am:      A sales department head overseeing 4–6 sales teams,
           accountable to leadership for Yektanet's overall market
           position across our advertiser portfolio.

Trying to: Identify systemic risks and team-level performance patterns
           early enough to intervene — before they show up as missed
           monthly targets.

But:       Radar only shows data at the individual advertiser level.
           To get a team-level or department-level view, I have to ask
           each manager to extract their own data and send it to me,
           then manually piece it together.

Because:   There is no aggregation layer in Radar — no team dashboard,
           no cross-advertiser trend rollup, no way to filter "show me
           everything owned by Team X" without re-uploading a
           specially filtered CSV.

Which makes me feel: Out of the loop between weekly review meetings,
           and dependent on my managers' initiative to surface problems
           rather than being able to spot them myself.
```

**Summary sentence:** A department head who needs to lead proactively is forced into reactive management because Radar surfaces no team-level patterns, leaving strategic decisions based on secondhand summaries rather than live data.

---

## Cross-Cutting Root Cause

All three problems share a common root: **Radar is optimized for individual advertiser deep-dives, not for the triage, communication, and oversight workflows that dominate daily usage.** The tool answers "what happened to advertiser X?" well, but struggles with "what should I pay attention to right now?" and "how do I act on what I found?"

---

*Generated: 2026-05-11 | Status: Hypothesis (not yet validated with user research) | References: `docs/personas.md`*
