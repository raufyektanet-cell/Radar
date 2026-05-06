# RADAR Feature Roadmap

You are a senior PM at Yektanet. Analyze the current **Market Analyzer** codebase and generate a prioritized feature roadmap.

## Step 1 — Audit the codebase

Read `src/App.tsx` and identify:
1. **Existing features** that are incomplete, buggy, or have TODO/suppressed warnings (e.g., the `void ResponsiveContainer; void LineChart; void Line;` on line 519 — charts are imported but unused)
2. **UX friction points** in the current flow (upload → analyze → preview)
3. **Missing features** a sales manager would obviously want
4. **Technical debt** that blocks future work

## Step 2 — Generate the roadmap

Output a roadmap table with 3 horizons:

### Now (Next 2 weeks) — Quick wins, high value
| # | Feature | Why | Effort | Impact |
|---|---------|-----|--------|--------|
| 1 | ... | ... | S/M/L | High/Med/Low |

### Next (1–2 months) — Meaningful improvements
| # | Feature | Why | Effort | Impact |

### Later (3–6 months) — Strategic bets
| # | Feature | Why | Effort | Impact |

## Step 3 — Top recommendation

Pick ONE feature from the "Now" horizon and write a 3-sentence pitch: what it is, why it matters to sales managers at Yektanet, and what file/component to start with.

## Constraints to consider
- No backend — changes must work with Vite proxy or client-side only
- Persian RTL UI — any new component must support `direction: rtl`
- localStorage budget: currently stores up to 10 reports and 5 history entries
- OpenRouter API: streaming, 4000 max_tokens, model `anthropic/claude-sonnet-4-5`
- The app is deployed to GitHub Pages: `https://raufyektanet-cell.github.io/Radar/`

$ARGUMENTS
