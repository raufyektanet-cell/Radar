# Market Analyzer — یکتانت

Daily competitive intelligence tool for Yektanet sales managers. Upload a CSV/XLSX of ad-network session data and get a structured Persian briefing powered by Claude.

## Setup

```bash
npm install
cp .env.example .env   # add your Anthropic API key
npm run dev
```

Open http://localhost:5173

## How it works

1. Upload a CSV or XLSX file with columns: `date`, `owner_name`, `account_manager_name`, and one column per ad network (Yektanet, Tapsell, Adexo, …)
2. Pick a report template (Standard / Brief / Detailed)
3. Click **شروع آنالیز** — Claude analyses the data and returns a structured Persian briefing
4. Copy the message text or download an HTML report

## API key & CORS

The Vite dev server proxies `/api/anthropic → https://api.anthropic.com` and injects your `ANTHROPIC_API_KEY` server-side, so the key is never exposed in the browser bundle.

For production deployment, replace the Vite proxy with a backend endpoint that forwards requests and adds the key header.

## Tech stack

- React 18 + TypeScript
- Vite 5
- xlsx (file parsing)
- recharts (available for future charts)
- Claude claude-sonnet-4-20250514 via Anthropic API
