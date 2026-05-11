# PRD Development Skill

## Context: Radar (Market Analyzer)
Radar is a Vite+React SPA used by Yektanet sales managers to analyze advertiser CSV data — viewing spend trends, market share heatmaps, industry breakdowns, and AI chat insights.

## PRD Development Workflow

Eight phases to transform discovery insights into engineering-ready specifications:

### Phase 1: Executive Summary (30 min)
One paragraph: problem, solution, impact.

### Phase 2: Problem Statement (60 min)
Who experiences it, what it is, why it matters, supporting evidence.

### Phase 3: Target Users & Personas (30 min)
Primary and secondary personas with goals and pain points.

### Phase 4: Strategic Context (45 min)
Link to business OKRs, market opportunity, explain timing.

### Phase 5: Solution Overview (60 min)
High-level approach without prescribing exact implementation.

### Phase 6: Success Metrics (30 min)
Primary, secondary, and guardrail metrics with targets.

### Phase 7: User Stories & Requirements (90-120 min)
Testable user stories with Gherkin acceptance criteria.

### Phase 8: Out of Scope & Dependencies (30 min)
Explicitly define what's excluded, identify blockers.

### Key Principles
- PRDs are living documents, not waterfall contracts
- Collaborate during story breakdown to prevent isolation
- Explicit scope boundaries combat creep

## Your Task
Write a PRD for Radar's most impactful next feature: **Advertiser Health Score & Proactive Alerts**.

The feature would automatically calculate a health score per advertiser based on spend trend, market share change, and session frequency — then surface at-risk advertisers proactively without requiring manual data scanning.

Follow all 8 phases. For Phase 7, write at least 5 user stories with Gherkin acceptance criteria.

Save output to `docs/prd-health-score.md`.
