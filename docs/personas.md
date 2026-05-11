# Radar — Proto-Personas

> These are proto-personas based on codebase analysis, team structure (teams.ts), and observed UX patterns. All assumptions are marked `[ASSUMPTION—VALIDATE]`.

---

## Persona 1: سارا مدیر — Sales Manager Sara

**Role:** Sales Team Manager (e.g., Managed Service / Customer Success)
**Demographics:** 28–35 years old, university-educated, Tehran-based, 3–6 years at Yektanet
**Behavioral Context:** Manages a portfolio of 15–40 advertisers. Starts her day by reviewing yesterday's data before morning team standups. Alternates between Radar and spreadsheets. Comfortable with data but not technical.

**Representative Quote:**
> "هر روز باید CSV رو دستی آپلود کنم، بعد به تک‌تک آگهی‌دهنده‌ها نگاه کنم تا بفهمم کجا ریزش داریم. تا وقتی همه رو چک کنم نصف روزم رفته."

**Key Pain Points:**
1. Manual CSV upload every single day — no automation
2. No at-a-glance signal for which advertisers are at risk right now
3. Hard to prepare a concise briefing for a client call from raw data
4. [ASSUMPTION—VALIDATE] Frequently switches between Radar and a separate CRM or messaging tool to follow up
5. Alert volume can be noisy — hard to distinguish signal from noise

**Short-Term Goals (daily/weekly):**
- Quickly identify the 2–3 advertisers that need attention today
- Generate a market share snapshot for a client call
- Track whether a previously flagged advertiser has recovered

**Aspirational Goals:**
- Become the go-to person on her team for market intelligence
- Proactively catch churn before it happens, not react after
- [ASSUMPTION—VALIDATE] Reduce time spent on manual data preparation by 60%

**Influence Mapping:**
- **Shapes her decisions:** Team lead / department head (sets KPIs), account managers who flag issues, client feedback
- **She influences:** Junior account managers who rely on her analysis, client-facing decisions
- **Beliefs affecting adoption:** Trusts data she prepared herself more than auto-generated summaries; needs to understand *how* a score is calculated before trusting it

---

## Persona 2: علی حسابدار — Account Manager Ali

**Role:** Account Manager within a sales team (reports to a Sales Manager)
**Demographics:** 24–30 years old, newer to Yektanet (1–3 years), goal-oriented, competitive
**Behavioral Context:** Responsible for a subset of advertisers within the team. Uses Radar primarily to prep for advertiser check-in calls and to spot opportunities to upsell. [ASSUMPTION—VALIDATE] Uses mobile occasionally to check data between client meetings.

**Representative Quote:**
> "وقتی می‌رم پیش مشتری می‌خوام بگم «شما الان ۲۸٪ سهم بازار داری، رقیبت ۳۵٪» — ولی این عدد رو باید خودم از چند جا جمع‌آوری کنم. یه جای مرکزی نیست."

**Key Pain Points:**
1. No single view combining Yektanet spend, competitor sessions, and market share for one advertiser
2. [ASSUMPTION—VALIDATE] Exports are not formatted for direct sharing with clients
3. No way to annotate or bookmark a specific trend to revisit
4. Session is shared (admin-only login) — can't personalize views for his advertiser subset
5. [ASSUMPTION—VALIDATE] Mobile experience is poor for quick lookups

**Short-Term Goals (daily/weekly):**
- Pull a quick market share number before a client call
- Identify if a competitor recently entered an advertiser's category
- Show a client their trend vs. competitors in a presentable format

**Aspirational Goals:**
- Be seen as a trusted advisor who brings proactive insights to clients, not just reactive reports
- [ASSUMPTION—VALIDATE] Share live or PDF snapshots directly with clients from Radar
- Track personal performance metrics tied to advertiser health

**Influence Mapping:**
- **Shapes his decisions:** Sales manager's directives, client relationship dynamics, monthly quota pressure
- **He influences:** Client advertising decisions (budget allocation), junior team members
- **Beliefs affecting adoption:** Wants speed over depth — a tool that takes more than 30 seconds to give him an answer loses him

---

## Persona 3: رضا رهبر تیم — Team Lead Reza

**Role:** Department Head / Business Unit Lead overseeing multiple sales teams
**Demographics:** 33–42 years old, senior, manages 3–6 sales managers, presents to VP/C-level
**Behavioral Context:** Doesn't use Radar daily but reviews aggregated output weekly. Primary concern is cross-team performance trends and identifying systemic issues (e.g., a whole industry vertical declining). [ASSUMPTION—VALIDATE] Accesses Radar via shared screen in team reviews, not personally.

**Representative Quote:**
> "من نمی‌خوام تک‌تک آگهی‌دهنده‌ها رو ببینم — می‌خوام بدونم کدوم تیم‌ها خوب عمل می‌کنن و کجا باید مداخله کنم."

**Key Pain Points:**
1. No team-level or department-level aggregated view (current tool is per-advertiser)
2. No way to compare performance across multiple time periods in one view
3. [ASSUMPTION—VALIDATE] Has to ask managers to extract and reformat data for leadership presentations
4. No executive summary — must scroll through full analysis to find key signals
5. Can't track historical performance without re-uploading old CSVs manually

**Short-Term Goals (daily/weekly):**
- Weekly: spot which teams are over/underperforming vs. prior week
- Monthly: prepare a market position summary for leadership
- Identify systemic risks (e.g., a competitor gaining share across multiple advertisers)

**Aspirational Goals:**
- Make Radar the single source of truth for Yektanet's market intelligence
- [ASSUMPTION—VALIDATE] Integrate Radar data into leadership OKR reviews
- Reduce the manual slide-deck preparation work for monthly business reviews

**Influence Mapping:**
- **Shapes his decisions:** Business unit KPIs, VP/C-level expectations, competitive pressure
- **He influences:** Feature prioritization for internal tools, team workflows, what metrics matter
- **Beliefs affecting adoption:** Skeptical of tools that show too much — wants executive-level distillation, not raw data

---

## Key Assumptions to Validate

| Assumption | How to Validate |
|---|---|
| Sara switches between Radar and CRM/messaging daily | Shadow session or 15-min interview |
| Ali uses mobile for quick lookups | Check analytics / ask team |
| Exports need to be client-shareable | Show current export to 3 AMs, ask if they'd share it directly |
| Reza accesses via shared screen, not personally | Ask 2 team leads how they consume Radar output |
| Ali wants personal bookmark/annotation feature | Ask if they currently use workarounds (starred rows, screenshots) |
| Reza influences internal tool prioritization | Confirm org chart / decision-making process |

---

*Generated: 2026-05-11 | Status: Proto (unvalidated) | Next step: Schedule 3 user interviews to validate top assumptions*
