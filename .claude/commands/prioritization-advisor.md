# Prioritization Advisor Skill

## Context: Radar (Market Analyzer)
Radar is a Vite+React SPA used by Yektanet sales managers to analyze advertiser CSV data — viewing spend trends, market share heatmaps, industry breakdowns, and AI chat insights. The product is in active use by ~16 sales teams at Yektanet.

## Prioritization Advisor Framework

Select the right prioritization framework by assessing four key dimensions:

### Assessment Dimensions
1. **Product Stage** — Pre-PMF experimentation → mature multi-product portfolio
2. **Team Context** — Team size, alignment levels, stakeholder complexity
3. **Decision-Making Needs** — Filtering ideas / building consensus / data-driven rigor
4. **Data Availability** — Gut-based / lightweight / metrics-heavy

### Framework Options
- **RICE** (Reach × Impact × Confidence ÷ Effort) — scaling teams with moderate data
- **ICE** (Impact × Confidence × Ease) — lightweight for resource-constrained teams
- **Value/Effort Matrix** — quick visual for early-stage products
- **Kano Model** — distinguishing delighters from must-haves
- **MoSCoW** (Must/Should/Could/Won't) — stakeholder alignment
- **Cost of Delay** — time-sensitive decisions

### Anti-Patterns
- Framework whiplash (switching frameworks constantly)
- Applying frameworks that don't match your stage
- Treating scores as absolute rather than directional

## Your Task
Radar context: ~10 potential features in backlog, active internal tool, ~16 teams using it, small dev team (1-2 engineers).

1. Assess which framework best fits Radar's current stage and context
2. Apply that framework to prioritize these potential features:
   - Scheduled CSV auto-import
   - Advertiser health score / churn prediction
   - Email digest of weekly highlights
   - Mobile-responsive layout
   - Multi-CSV comparison (period over period)
   - Custom alert thresholds per advertiser
   - Team-level dashboards with aggregated metrics
   - API integration with Yektanet's CRM
   - PDF report export
   - Annotation / notes on trends

3. Output a prioritized backlog with scores and rationale

Save output to `docs/prioritized-backlog.md`.
