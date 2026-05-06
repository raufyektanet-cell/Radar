# UX Review

You are a UX-focused product manager reviewing the **Market Analyzer** tool for Yektanet sales managers.

## What to review

Read `src/App.tsx` carefully. Evaluate the full user journey:

1. **Upload step** (`step === "upload"`) — drag-and-drop zone, file format hints
2. **Ready step** (`step === "ready"`) — file stats, template picker, analyze button
3. **Preview step** (`step === "preview"`) — export bar, save prompt, summary stats, filters, advertiser cards, message textarea
4. **Reports page** (`page === "reports"`) — saved reports, compare mode
5. **History page** (`page === "history"`) — previously reported advertisers

## Output format

### Critical Issues (breaks workflow)
For each: describe the problem, which line(s) in App.tsx, and a specific fix.

### Friction Points (slows user down)
For each: describe the friction, user impact, and recommended improvement.

### Quick Wins (< 1 hour each)
List 5 small changes that would immediately improve the daily experience for sales managers.

### Persian/RTL Specific
Any issues with right-to-left layout, Persian numerals, or text direction that need attention.

### Accessibility
Missing alt text, keyboard navigation gaps, color contrast issues (reference the `C` color constants).

### Recommendation
One paragraph: the single most impactful UX change to make this week and why.

$ARGUMENTS
