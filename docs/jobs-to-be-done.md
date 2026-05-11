# Radar — Jobs-to-Be-Done Analysis

> Based on codebase analysis, team structure, and proto-personas from `docs/personas.md`.
> Importance and satisfaction ratings are hypotheses — validate with user interviews.

---

## Job 1: Monitor Advertiser Health

**Functional:** Know which advertisers are declining, recovering, or at risk — before it becomes a problem.
**Social:** Be seen as a proactive manager who catches issues early rather than reacting to complaints.
**Emotional:** Feel in control of the portfolio without having to manually inspect every account daily.

**Importance:** High | **Current Satisfaction:** Low

### Pains
1. No automated health scoring — must visually scan all advertiser cards to spot patterns
2. Alerts exist but are reactive (surface after the fact) and require the user to navigate to a separate screen
3. No "change from last week" view — trends require visual interpretation of charts, not clear signals

### Gains
1. A ranked priority list on login: "3 advertisers need attention today"
2. Configurable thresholds per advertiser (e.g., alert when Yektanet share drops >10 pts)
3. Health score history — see if an at-risk advertiser is recovering or still declining

---

## Job 2: Prepare for Client Conversations

**Functional:** Assemble a clear, credible competitive positioning snapshot for a specific advertiser before a call.
**Social:** Appear data-driven and authoritative in client meetings — not scrambling for numbers.
**Emotional:** Feel confident walking into a conversation, knowing the data tells a compelling story.

**Importance:** High | **Current Satisfaction:** Low

### Pains
1. No client-ready output — CSV export is raw and unstyled, requires manual slide preparation
2. Data is spread across heatmap, trends chart, and table — no single "advertiser brief" view
3. [ASSUMPTION—VALIDATE] No Persian-language narrative summary suitable for sharing externally

### Gains
1. One-click advertiser brief: market share, top competitors, 30-day trend — formatted for sharing
2. PDF or image export with Yektanet branding
3. AI-generated Persian narrative summary: "این آگهی‌دهنده در ۳۰ روز گذشته..."

---

## Job 3: Understand Competitive Landscape in an Industry

**Functional:** See which agencies dominate a given industry vertical and how that's shifting over time.
**Social:** Position Yektanet as the expert on a client's competitive context, not just a vendor.
**Emotional:** Feel knowledgeable enough to have strategic conversations, not just tactical ones.

**Importance:** High | **Current Satisfaction:** Medium

### Pains
1. Industry treemap shows current state but no trend — can't see if a vertical is shifting toward competitors
2. Filtering by industry and viewing competitor breakdown requires multiple manual steps
3. No "new entrant" signal — can't tell when a new agency starts taking share in a category

### Gains
1. Industry trend view: share of each agency within a vertical over time (30/60/90 days)
2. Auto-flagging: "تپسل در دسته بانک‌ها در ۷ روز گذشته ۸٪ رشد داشته"
3. Benchmarking: how does a specific advertiser compare to category average?

---

## Job 4: Track Portfolio Performance Over Time

**Functional:** Compare this week's advertiser portfolio performance to last week, last month, and last quarter.
**Social:** Demonstrate measurable progress to team leads and leadership in weekly reviews.
**Emotional:** Feel like the work is compounding — that past effort shows up in the numbers.

**Importance:** High | **Current Satisfaction:** Low

### Pains
1. Each CSV upload is a one-time snapshot — no persistent historical database across uploads
2. Saved reports show past analyses but can't be compared side-by-side
3. No team-level or portfolio-level aggregate — only individual advertiser views

### Gains
1. Automatic period-over-period comparison when a new CSV is uploaded
2. Portfolio dashboard: aggregate Yektanet share, total sessions, at-risk count across all advertisers
3. Team-level view for managers: how is Team X's portfolio performing vs. last month?

---

## Job 5: Answer Specific Data Questions Quickly

**Functional:** Get a precise answer to a question like "which advertiser in the finance category has the highest competitor exposure?" without building a manual filter chain.
**Social:** Respond instantly to ad-hoc requests from team leads or clients without saying "let me look into that."
**Emotional:** Feel capable and resourceful rather than dependent on engineers for data queries.

**Importance:** Medium | **Current Satisfaction:** Medium (AI chat addresses this partially)

### Pains
1. AI chat is powerful but users may not know what to ask — no guided prompts for common questions
2. Chat context resets on re-upload — losing conversation history when new data is loaded
3. [ASSUMPTION—VALIDATE] Response time for complex queries can feel slow during peak hours

### Gains
1. Suggested questions based on what the data shows: "آیا می‌خواهید درباره ۳ آگهی‌دهنده‌ای که بیشترین افت را داشتند بیشتر بدانید؟"
2. Chat history persists across sessions tied to a specific dataset
3. Structured output for common queries (tables, not just prose)

---

## Underserved Jobs — Prime Product Opportunities

These three jobs score **high importance / low current satisfaction** — the biggest unmet needs in Radar today:

### Opportunity 1: Proactive Health Monitoring
**Job:** Monitor advertiser health without manual daily scanning
**Gap:** Radar is reactive (user must look) not proactive (tool surfaces signals)
**Suggested direction:** Health score algorithm + daily digest / push notification + ranked priority inbox

### Opportunity 2: Client-Ready Output
**Job:** Prepare for client conversations with shareable competitive data
**Gap:** All output is for internal analysis; nothing is formatted for external communication
**Suggested direction:** Single-advertiser "brief" view with PDF/image export + AI-generated Persian summary

### Opportunity 3: Historical Portfolio Tracking
**Job:** Track portfolio performance trends over time and across teams
**Gap:** Each upload is stateless; no persistence, comparison, or aggregation layer exists
**Suggested direction:** Persistent datastore (localStorage or backend) that accumulates uploads and enables period-over-period views

---

*Generated: 2026-05-11 | Status: Hypothesis | References: `docs/personas.md`, `docs/problem-statements.md`*
