# Radar — Customer Journey Map

> Persona: Sara — Sales Manager (see `docs/personas.md`)
> Generated: 2026-05-11 | Revisit: quarterly | References: `docs/problem-statements.md`, `docs/jobs-to-be-done.md`

---

## Journey Overview

| | **Stage 1: Morning Setup** | **Stage 2: Analysis** | **Stage 3: Insight** | **Stage 4: Action** | **Stage 5: Follow-Up** |
|---|---|---|---|---|---|
| **Customer Actions** | Opens Radar, uploads CSV, waits for parse | Reviews trends chart, heatmap, industry treemap | Asks AI chat a specific question | Exports CSV or saves report for client call | Returns days later, reviews saved reports, compares to new data |
| **Touchpoints** | Login screen → Upload screen → Ready screen | TrendsScreen, heatmap, industry map, advertiser cards | ChatScreen, quick-question chips | Export button, save report dialog, Topbar save indicator | Reports page, History page, saved report cards |
| **Customer Experience** | Anxious and rushed — "I need to know what happened yesterday before standup" | Focused but manually scanning — "I have to look at every card to find problems" | Cautiously curious — "Will it actually understand my question, or give me a generic answer?" | Relieved if export is clean; frustrated if it needs reformatting | Uncertain — "I can't remember what I was looking at last week" |
| **KPIs** | Time-to-first-insight after login (target: <2 min) | % of sessions where user opens ≥3 screens (engagement depth) | AI chat usage rate per session (target: >40%) | Export/save rate per session (target: >30%) | Report return rate within 7 days (target: >50%) |
| **Business Goals** | Reduce daily data prep friction to <10 min | Surface the right advertisers without manual scanning | Make AI chat the default question-answering layer | Enable client-ready output from Radar directly | Build habitual daily/weekly usage pattern |
| **Teams Involved** | Product (upload UX), Data (CSV format stability) | Product (chart UX), Data (CSV schema) | Product (chat UX), AI/LLM infrastructure | Product (export), Sales (client format needs) | Product (reports UX), Sales (workflow integration) |

---

## Stage 1: Morning Setup

**Scenario:** Sara arrives at her desk at 8:45am. Standup is at 9:30. She needs to know if anything significant happened with her advertisers yesterday.

### Customer Actions
- Opens Radar in browser (or checks if last session is still live)
- Navigates to upload screen
- Locates yesterday's CSV export from internal data pipeline (separate tool)
- Drags file onto upload zone
- Waits for parsing, checks file stats before clicking Analyze

### Touchpoints
- Browser bookmark / direct URL
- Login screen (if session expired)
- Upload zone (drag-drop or click-to-select)
- File stats card (row count, date range, advertiser count)
- "تحلیل بازار" analyze button

### Customer Experience
> "هر روز باید همین کارو بکنم. CSV رو پیدا کنم، آپلود کنم، صبر کنم."

- **Emotion:** Resigned routine mixed with low-level anxiety about what the data will show
- **Thought:** "If something went wrong yesterday I want to know before someone else tells me"
- **Friction point:** Finding the right CSV file from the data pipeline tool adds 2–3 minutes of context-switching before Radar even opens

### KPIs
- Upload-to-analyze time: target <60 seconds from file drop
- Session start rate (daily active users / total users): target >80% on working days

### Business Goals
- Make the daily startup feel effortless — zero-friction entry into the day's intelligence

### Teams Involved
- Product: upload UX, loading states
- Data/Engineering: CSV export availability and naming consistency

---

## Stage 2: Analysis

**Scenario:** Sara clicks Analyze. She has ~20 minutes before standup. She needs to identify which 2–3 advertisers require action today.

### Customer Actions
- Scans the advertiser cards on the main screen for alert badges
- Opens TrendsScreen to check market share movement for flagged advertisers
- Switches to heatmap to spot patterns across the portfolio
- Checks industry treemap to see if any verticals shifted
- Manually notes advertiser names she wants to follow up on

### Touchpoints
- Advertiser cards with alert indicators
- Trends line chart (stacked bars + share heatmap)
- Industry treemap
- Filter bar (by industry, team, alert type)
- Sort controls on the advertiser table

### Customer Experience
> "باید کارت به کارت بگردم تا ببینم کجا مشکل داریم."

- **Emotion:** Focused but time-pressured; mild frustration at the volume of cards to scan
- **Thought:** "There must be a better way than looking at every single advertiser manually"
- **Delight moment:** When an alert badge immediately surfaces an advertiser she was already worried about — "یکتانت sees it too"

### KPIs
- Average advertiser cards viewed per session: currently unknown (instrument this)
- Alert interaction rate: % of sessions where user opens ≥1 alert detail
- Filter usage rate: % of sessions using industry/team filter

### Business Goals
- Surface the 2–3 actionable advertisers without requiring full-portfolio scan
- Reduce average time-to-first-actionable-insight to <5 minutes

### Teams Involved
- Product: alert ranking, card layout, filter UX
- Data: alert threshold calibration

---

## Stage 3: Insight

**Scenario:** Sara sees that "آگهی‌دهنده X" dropped from 42% to 28% Yektanet share in 3 days. She wants to understand why before calling the account manager.

### Customer Actions
- Opens AI chat screen
- Either types a specific question or taps a quick-question chip
- Reads the streamed response
- Asks a follow-up question for clarification
- Copies key numbers from the chat to use in her follow-up message

### Touchpoints
- ChatScreen with message bubbles (RTL layout)
- Quick-question chips ("چه آگهی‌دهنده‌ای بیشترین افت را داشته؟")
- Streaming response indicator
- Chat input field + send button

### Customer Experience
> "وقتی جواب درستی می‌ده خیلی وقت می‌بره. وقتی جواب اشتباه می‌ده نمی‌دونم باید بهش اعتماد کنم یا نه."

- **Emotion:** Cautiously optimistic — the chat is useful but trust is not yet fully established
- **Thought:** "I need to verify what it says against the raw numbers — I can't just repeat this to my manager without checking"
- **Frustration:** Chat history disappears if she re-uploads a new CSV — loses context between sessions

### KPIs
- AI chat engagement rate: % of sessions where chat is opened (target: >40%)
- Messages per chat session: target >2 (indicates productive back-and-forth)
- Quick-chip usage rate: % of chat sessions starting with a chip vs. free text

### Business Goals
- Establish AI chat as the trusted first-stop for "why did this happen?" questions
- Reduce time spent cross-referencing chat answers with raw table data

### Teams Involved
- Product: chat UX, quick-chip curation, trust signals (source attribution)
- AI/LLM: response quality, latency, context window management

---

## Stage 4: Action

**Scenario:** Sara has identified 3 advertisers to discuss in standup and 1 that needs a client-call prep briefing. She needs to capture this before leaving her desk.

### Customer Actions
- Clicks "خروجی CSV" to export filtered data
- Opens the exported file, realizes she needs to reformat it for the client
- Alternatively: saves a report in Radar via the save dialog
- Writes notes about what she found in a separate messaging app or email

### Touchpoints
- Export button (table view)
- Save report dialog (name input, confirmation)
- Topbar "ذخیره نشده" amber chip (unsaved state indicator)
- External tools: Excel, Google Sheets, Telegram/email for client prep

### Customer Experience
> "خروجی CSV می‌گیرم ولی باز باید توی اکسل کار کنم تا بشه پیش مشتری برد."

- **Emotion:** Mild disappointment — the insight was found in Radar but the communication still happens outside it
- **Thought:** "I wish I could just share this directly without reformatting"
- **Partial relief:** Saving a report in Radar at least means she can come back to this snapshot

### KPIs
- Export rate: % of sessions ending with a CSV export (instrument this)
- Save report rate: % of sessions with a saved report
- Post-Radar tool switch rate: how often users open Excel within 10 minutes of a Radar session (proxy for reformatting friction)

### Business Goals
- Make Radar the end-to-end workflow — insight → action → output — without leaving the app
- Increase saved report rate to >50% of analysis sessions

### Teams Involved
- Product: PDF export, report save UX, export formatting
- Sales: define what "client-ready" output looks like

---

## Stage 5: Follow-Up

**Scenario:** Three days later, Sara wants to check if the advertiser she flagged has recovered. She also wants to compare this week's portfolio to last week's.

### Customer Actions
- Returns to Radar, uploads new CSV
- Navigates to Reports page to find her saved report
- Tries to compare the saved snapshot to today's data
- Checks History page for advertisers she previously flagged
- Finds it difficult to see what changed vs. her last session

### Touchpoints
- Reports page (saved report cards)
- History page (previously reported advertisers)
- Compare mode toggle (if enabled)
- Re-upload flow (starts a fresh analysis, no connection to prior session)

### Customer Experience
> "می‌خوام ببینم آگهی‌دهنده‌ای که هفته پیش مشکل داشت الان کجاست — ولی هر بار که CSV آپلود می‌کنم همه چیز از نو شروع می‌شه."

- **Emotion:** Frustrated by statelessness — each session feels disconnected from the last
- **Thought:** "I'm rebuilding context from scratch every time. I know I saved a report but I can't easily compare it to now."
- **Aspiration:** "I want Radar to remember what I was tracking"

### KPIs
- 7-day return rate: % of users who return to Radar within 7 days of a session (target: >70%)
- Report retrieval rate: % of sessions that include opening a saved report
- Compare mode usage rate (once built): target >30% of returning sessions

### Business Goals
- Build habitual daily/weekly usage pattern — Radar as the persistent intelligence layer, not a one-off lookup tool
- Reduce re-orientation time at session start (user should resume where they left off)

### Teams Involved
- Product: reports UX, history page, compare mode, session persistence
- Engineering: localStorage or backend persistence strategy

---

## Top 3 Pain Points with Improvement Opportunities

### Pain Point 1: Manual Portfolio Scanning (Stage 2)
**Current experience:** Sara spends 15–25 minutes scanning all advertiser cards to find the 2–3 that matter.
**Root cause:** Radar surfaces all advertisers equally — no ranked priority view.
**Measurable impact:** Estimated 15–20 minutes of avoidable work per day across 16 sales managers = ~4–5 hours of collective time lost daily.
**Improvement opportunity:** Advertiser health score + ranked "needs attention" list on the analysis screen. Success metric: reduce average time-to-first-actionable-insight from ~15 min to <5 min.

### Pain Point 2: No Client-Ready Output (Stage 4)
**Current experience:** After finding an insight in Radar, Sara must reformat data in Excel before sharing with a client.
**Root cause:** CSV export is raw and unstyled; no single-advertiser briefing view exists.
**Measurable impact:** ~30 min of reformatting work per client call prep; estimated 2–3 client calls per manager per week = 60–90 min of avoidable work weekly per manager.
**Improvement opportunity:** PDF export of a single-advertiser brief (market share + trends + AI summary). Success metric: % of sessions ending with PDF export instead of CSV export; target >30%.

### Pain Point 3: Session Statelessness (Stage 5)
**Current experience:** Each new CSV upload resets the analysis state — no persistent view of what changed since last session.
**Root cause:** Radar holds only one dataset in memory; no historical persistence layer.
**Measurable impact:** Users must re-orient from scratch every session; 7-day return rate likely lower than it should be because the tool doesn't "remember" the user's context.
**Improvement opportunity:** Multi-CSV comparison (period-over-period). Success metric: 7-day return rate; target >70% of active users returning within a week.

---

*Next review: 2026-08-11 | Update triggers: major feature launch, user research findings, significant usage pattern change*
