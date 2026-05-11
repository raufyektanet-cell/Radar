# Roadmap Planning Skill

## Context: Radar (Market Analyzer)
Radar is a Vite+React SPA used by Yektanet sales managers to analyze advertiser CSV data — viewing spend trends, market share heatmaps, industry breakdowns, and AI chat insights.

## Strategic Roadmap Planning Workflow

Five phases to transform strategy into executable roadmaps:

### Phase 1: Gather Inputs
Collect business goals, customer problems, technical constraints, and stakeholder requests from discovery, OKRs, and engineering leadership.

### Phase 2: Define Initiatives
Convert inputs into epics with hypotheses, success metrics, and effort estimates.

### Phase 3: Prioritize Initiatives
Rank epics using RICE or ICE, adjusting for strategic alignment beyond pure metrics.

### Phase 4: Sequence Roadmap
Organize into quarters (Now/Next/Later), mapping dependencies and validating feasibility.

### Phase 5: Communicate Roadmap
Present to stakeholders through structured presentations, incorporating feedback before publishing.

### Key Principles
- Outcome-driven framing (not feature lists)
- Tie initiatives to business metrics and customer problems
- Maintain flexibility for learning-based adjustments

### Critical Pitfalls
- Treating roadmaps as waterfall contracts
- Prioritizing by executive preference over data
- Mapping features without connecting to outcomes
- Excluding stakeholder input during development

## Your Task
Create a 3-quarter roadmap for Radar. Use this as input:

**Business goals:** Help sales managers reduce churn, increase Yektanet revenue, prepare better client conversations

**Known problems (from codebase analysis):**
- No automated data refresh (manual CSV upload each time)
- No proactive alerting for at-risk advertisers
- Limited mobile support
- AI chat is new and underutilized

**Format:**
- Now (Q2 2026): High-priority, in-progress or immediately actionable
- Next (Q3 2026): Important but requires preparation
- Later (Q4 2026+): Strategic bets and longer-horizon work

For each initiative include: hypothesis, success metric, estimated effort (S/M/L), dependencies.

Save output to `docs/roadmap.md`.
