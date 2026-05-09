import { useState, useCallback, useEffect, useMemo, useRef, useContext, createContext } from "react";
import * as XLSX from "xlsx";
import { TEAMS, ADMIN_USER, type SessionUser } from "./teams";

const MODEL = "anthropic/claude-sonnet-4-5";
const OPENROUTER_KEY = import.meta.env.VITE_OPENROUTER_KEY as string | undefined;
const HISTORY_KEY = "analyzer:reported_advertisers";
const REPORTS_KEY = "analyzer:saved_reports";
const THEME_KEY = "analyzer:theme";
const SESSION_KEY = "radar:session";
const APP_PASSWORD = (import.meta.env.VITE_APP_PASSWORD as string | undefined) || "radar1403";

// ── Design Tokens v2 ──────────────────────────────────────────────────────────

function makeTokens(isDark: boolean) {
  const a = "#D4623A", ar = "212,98,58";
  if (isDark) return {
    bg: "#070B12", bgDeep: "#040709", surface: "#0D1420", card: "#111927", cardHov: "#161F30",
    border: "rgba(255,255,255,.06)", border2: "rgba(255,255,255,.11)", border3: "rgba(255,255,255,.18)",
    accent: a, accentDim: `rgba(${ar},.12)`, accentBrd: `rgba(${ar},.28)`, accentGlow: `rgba(${ar},.4)`, accentHov: "#C45530",
    green: "#1DB87E", greenDim: "rgba(29,184,126,.12)", greenBrd: "rgba(29,184,126,.28)",
    amber: "#D97B2A", amberDim: "rgba(217,123,42,.12)", amberBrd: "rgba(217,123,42,.28)",
    red: "#E05252", redDim: "rgba(224,82,82,.12)", redBrd: "rgba(224,82,82,.28)",
    blue: "#4A8EDB", blueDim: "rgba(74,142,219,.12)",
    t1: "#EEF2FF", t2: "rgba(238,242,255,.60)", t3: "rgba(238,242,255,.32)", t4: "rgba(238,242,255,.14)",
    mono: "'JetBrains Mono', monospace",
    // legacy aliases kept for upload/login/modal screens
    surface2: "#0D1420", text1: "#EEF2FF", text2: "rgba(238,242,255,.60)", text3: "rgba(238,242,255,.32)",
    coral: a, coralDim: `rgba(${ar},.12)`, coralBorder: `rgba(${ar},.28)`,
    greenBorder: "rgba(29,184,126,.28)", amberBorder: "rgba(217,123,42,.28)",
    danger: "#E05252", dangerBg: "rgba(224,82,82,.10)", dangerBorder: "rgba(224,82,82,.28)",
  };
  return {
    bg: "#F5F2ED", bgDeep: "#EDE8DF", surface: "#FFFFFF", card: "#FFFFFF", cardHov: "#FAF7F2",
    border: "rgba(0,0,0,.07)", border2: "rgba(0,0,0,.13)", border3: "rgba(0,0,0,.22)",
    accent: a, accentDim: `rgba(${ar},.10)`, accentBrd: `rgba(${ar},.22)`, accentGlow: `rgba(${ar},.25)`, accentHov: "#C45530",
    green: "#15956A", greenDim: "rgba(21,149,106,.10)", greenBrd: "rgba(21,149,106,.22)",
    amber: "#A0600A", amberDim: "rgba(160,96,10,.10)", amberBrd: "rgba(160,96,10,.22)",
    red: "#C93535", redDim: "rgba(201,53,53,.10)", redBrd: "rgba(201,53,53,.22)",
    blue: "#1D5FAA", blueDim: "rgba(29,95,170,.10)",
    t1: "#0F0E0C", t2: "#4A4540", t3: "#9B9490", t4: "#C8C3BD",
    mono: "'JetBrains Mono', monospace",
    // legacy aliases
    surface2: "#F5F2ED", text1: "#0F0E0C", text2: "#4A4540", text3: "#9B9490",
    coral: a, coralDim: `rgba(${ar},.10)`, coralBorder: `rgba(${ar},.22)`,
    greenBorder: "rgba(21,149,106,.22)", amberBorder: "rgba(160,96,10,.22)",
    danger: "#C93535", dangerBg: "rgba(201,53,53,.08)", dangerBorder: "rgba(201,53,53,.22)",
  };
}

type Theme = ReturnType<typeof makeTokens>;
const ThemeCtx = createContext<Theme>(makeTokens(true));
function useD() { return useContext(ThemeCtx); }

const AGENCY_COLORS: Record<string, string> = {
  "یکتانت": "#D4623A", "تپسل": "#3B82F6", "ادکسو": "#8B5CF6", "آپارات": "#EF4444",
  "یلو ادوایز": "#F59E0B", "بله": "#10B981", "روبیکا": "#EC4899", "دیما": "#6366F1",
  "طاووس": "#14B8A6", "دارت": "#F97316", "چاووش": "#84CC16", "تلوبیون": "#06B6D4",
  "ایتا": "#A78BFA", "سروش": "#FB923C", "نجوا": "#34D399", "بازار": "#64748B", "مایکت": "#78716C",
  "تریبون": "#0EA5E9", "جریان": "#D946EF", "ادورج": "#22D3EE",
};

const TEMPLATES: Record<string, { label: string; desc: string }> = {
  standard: { label: "استاندارد", desc: "پیام روزانه معمول" },
  brief: { label: "خلاصه", desc: "فقط مهم‌ترین نکات" },
  detailed: { label: "تفصیلی", desc: "تحلیل کامل با جزئیات" },
};

const TEMPLATE_INSTRUCTIONS: Record<string, string> = {
  standard: "",
  brief: "Be very concise — max 1 sentence per advertiser. Only the single most important change. IMPORTANT: You MUST still use the exact ##ADVERTISER##/##LEAD##/##END## block format — do not skip it.",
  detailed: "Be detailed — up to 4 sentences per advertiser. Include all agencies and full trend. IMPORTANT: You MUST still use the exact ##ADVERTISER##/##LEAD##/##END## block format for every entry.",
};

const SYSTEM_PROMPT = `You are a market intelligence analyst at Yektanet (یکتانت). You write a daily Persian briefing for sales managers.

CSV STRUCTURE
- Date (date), Owner_id (id), Advertiser_name (copy EXACTLY as-is), Category_level_1 (industry), Category_level_2 (sub-industry), Team, Account_manager_name, Performance_manager_name, Supervisor_name, Daily_spend (تومان spent on Yektanet), Total_sessions
- Agency session columns: Yektanet, Tapsell, Deema, Tavoos, Adexo, Chavosh, Aparat, Daart, Yellowadwise, Najva, Triboon, Jaryan, Telewebion, Adverge, Soroush, Soroush_ny, Bale_ny, Rubika_ny, Eitaa_ny, Bazaar, Myket
- OUR NETWORK: Yektanet only | COMPETITORS: all others | Empty/blank=0
- _ny suffix = non-Yektanet traffic on that platform

SPECIAL: همکده،5040،بیکوپلاس،Owner_id 9868=SKIP. Aparat/Tavoos/Telewebion=VIDEO only. Aparat-owned=NOT leads.

PRIORITY: +5 reactivated | +4 very high sessions | +3 new/cut agency or Yektanet declining >20pts | +2 3+ agencies/high sessions | +1 dominant changed | -999 id 9868 | -3 همکده/5040/بیکوپلاس | -2 minor. Aim 8-12.

AGENCY TRANSLATIONS: Yektanet→یکتانت|Tapsell→تپسل|Adexo→ادکسو|Aparat→آپارات|Deema→دیما|Tavoos→طاووس|Yellowadwise→یلو ادوایز|Daart→دارت|Chavosh→چاووش|Telewebion→تلوبیون|Bale_ny→بله|Rubika_ny→روبیکا|Eitaa_ny→ایتا|Soroush_ny→سروش|Soroush→سروش|Najva→نجوا|Triboon→تریبون|Jaryan→جریان|Adverge→ادورج|Bazaar→بازار|Myket→مایکت

NUMBERS: "حدود ۳۰ هزار سشن" | shares nearest 5٪ | "یکتانت (۴۸٪)|تپسل (۱۸٪)" — NO session counts next to percentages. End each SUMMARY with: حجم کل دیروز: حدود X هزار سشن
DATE: Latest→دیروز|Before→پریروز|Earlier→اوایل هفته. Never English month names.
NO hashtags. NO کمپین. NO translated Advertiser_name. NO English in output. Start with ##

OUTPUT:
##ADVERTISER##
NAME: [wallex.com](http://wallex.com)
OWNERID: 12345
AGENCIES: Yektanet:4000,Tapsell:5500
SUMMARY: یکتانت (۴۰٪) | تپسل (۵۵٪). سهم یکتانت از ۷۰٪ به ۴۰٪ کاهش یافته. حجم کل دیروز: حدود ۱۰ هزار سشن
##END##
##LEAD##
NAME: [x.ir](http://x.ir)
OWNERID: 999
AGENCIES: Bale_ny:1500
NOTE: یکتانت صفره. قابل اپروچه. حجم کل دیروز: حدود ۱.۵ هزار سشن
##END##
##COMPETITOR##
PLATFORM: تپسل
NEWCLIENTS: ندارد
TOPCLIENTS: [wallex.com](http://wallex.com)
NOTE: در چند اکانت سهم گرفته.
##END##
##MARKET##
NOTE: وضعیت کلی بازار.
##END##
##DISCLAIMER##
داده‌ها بر اساس کراول وب هستن و تخمینی‌اند — روندها قابل اعتمادند اما اعداد دقیق نیستن.
##END##

RULES: No hashtags. Max 2 sentences per advertiser. Only top 2 agencies. LEADS mandatory. No کمپین. No translated Advertiser_name. No English. Start with ##`;

// ── Utilities ────────────────────────────────────────────────────────────────

function normalizeDate(val: unknown): string {
  if (!val) return "";
  if (typeof val === "number") return new Date(Math.round((val - 25569) * 86400 * 1000)).toISOString().slice(0, 10);
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (mdy) return `${mdy[3]}-${mdy[1].padStart(2, "0")}-${mdy[2].padStart(2, "0")}`;
  const d = new Date(s); return isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
}

function readFile(file: File, onSuccess: (rows: Record<string, string>[], csvText: string) => void, onError: (err: unknown) => void) {
  const isXlsx = /\.(xlsx|xls)$/i.test(file.name);
  const reader = new FileReader();
  reader.onload = e => {
    try {
      if (isXlsx) {
        const wb = XLSX.read((e.target as FileReader).result, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false }) as Record<string, string>[];
        onSuccess(rows, XLSX.utils.sheet_to_csv(ws));
      } else {
        const text = (e.target as FileReader).result as string;
        const lines = text.trim().split("\n"), sep = lines[0].includes("\t") ? "\t" : ",";
        const headers = lines[0].split(sep).map(h => h.trim().replace(/^"|"$/g, ""));
        const rows = lines.slice(1).map(line => {
          const vals: string[] = []; let cur = "", inQ = false;
          for (const c of line) { if (c === '"') { inQ = !inQ; continue; } if (c === sep && !inQ) { vals.push(cur.trim()); cur = ""; continue; } cur += c; }
          vals.push(cur.trim());
          const obj: Record<string, string> = {}; headers.forEach((h, i) => { obj[h] = (vals[i] || "").replace(/^"|"$/g, "").trim(); }); return obj;
        }).filter(r => r[headers[0]]);
        onSuccess(rows, text);
      }
    } catch (err) { onError(err); }
  };
  if (isXlsx) reader.readAsArrayBuffer(file); else reader.readAsText(file, "UTF-8");
}

function getStats(rows: Record<string, string>[]) {
  const dates = [...new Set(rows.map(r => normalizeDate(r.Date || r.date)).filter(Boolean))].sort();
  return { dates, lastDate: dates[dates.length - 1] || "", totalRows: rows.length, advertisers: [...new Set(rows.map(r => r.Advertiser_name || r.owner_name).filter(Boolean))].length };
}

function todayLabel(): string {
  const d = new Date(), m = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  return `${d.getDate()} ${m[d.getMonth()]} ${d.getFullYear()}`;
}

function extractBlocks(tag: string, text: string): string[] {
  const re = new RegExp(`##${tag}##([\\s\\S]*?)##END##`, "g");
  const out: string[] = []; let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) out.push(m[1].trim()); return out;
}

function parseFields(block: string): Record<string, string> {
  const obj: Record<string, string> = {}, known = ["NAME", "OWNERID", "AGENCIES", "SUMMARY", "NOTE", "PLATFORM", "NEWCLIENTS", "TOPCLIENTS"];
  const lines = block.split("\n"); let key: string | null = null, val: string[] = [];
  for (const line of lines) {
    const ci = line.indexOf(":"), pk = ci > -1 ? line.slice(0, ci).trim() : null;
    if (pk && known.includes(pk)) { if (key) obj[key] = val.join("\n").trim(); key = pk; val = [line.slice(ci + 1).trim()]; }
    else if (key) val.push(line);
  }
  if (key) obj[key] = val.join("\n").trim(); return obj;
}

interface Agency { name: string; value: number; color: string; }

function parseAgencies(str: string): Agency[] {
  if (!str) return [];
  const TR: Record<string, string> = {
    Yektanet: "یکتانت", Tapsell: "تپسل", Adexo: "ادکسو", Aparat: "آپارات", Deema: "دیما",
    Tavoos: "طاووس", Yellowadwise: "یلو ادوایز", YellowAdwise: "یلو ادوایز", Daart: "دارت",
    Chavosh: "چاووش", Telewebion: "تلوبیون", Bale_ny: "بله", Bale_NonYektanet: "بله",
    Rubika_ny: "روبیکا", Rubika_NonYektanet: "روبیکا", Eitaa_ny: "ایتا", Eitaa_NonYektanet: "ایتا",
    Soroush_ny: "سروش", Soroush_NonYektanet: "سروش", Soroush: "سروش",
    Najva: "نجوا", Triboon: "تریبون", Jaryan: "جریان", Adverge: "ادورج",
    Bazaar: "بازار", Myket: "مایکت",
  };
  return str.split(",").map(s => { const [k, v] = s.trim().split(":"); const name = TR[k?.trim()] || k?.trim() || ""; return { name, value: parseInt(v) || 0, color: AGENCY_COLORS[name] || "#999" }; }).filter(a => a.name && a.value > 0);
}

function formatNumber(n: number): string {
  if (n >= 1000000) return `${(n / 1000000).toFixed(1)}M`;
  if (n >= 1000) return `${Math.round(n / 1000)}K`;
  return String(n);
}

function fmtMoney(n: number): string {
  if (n >= 1000000000) return `${(n / 1000000000).toFixed(1)}B`;
  if (n >= 1000000) return `${Math.round(n / 1000000)}M`;
  return `${Math.round(n / 1000)}K`;
}

function toJalali(date: Date): string {
  const months = ["فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور","مهر","آبان","آذر","دی","بهمن","اسفند"];
  const gy = date.getFullYear(), gm = date.getMonth() + 1, gd = date.getDate();
  const leap = (gy % 4 === 0 && gy % 100 !== 0) || gy % 400 === 0;
  const gDays = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  let gDno = 365 * gy + Math.floor((gy + 3) / 4) - Math.floor((gy + 99) / 100) + Math.floor((gy + 399) / 400);
  for (let i = 0; i < gm - 1; i++) gDno += gDays[i];
  gDno += gd - 1;
  let jDno = gDno - 79;
  const jNp = Math.floor(jDno / 12053); jDno %= 12053;
  let jy = 979 + 33 * jNp + 4 * Math.floor(jDno / 1461); jDno %= 1461;
  if (jDno >= 366) { jy += Math.floor((jDno - 1) / 365); jDno = (jDno - 1) % 365; }
  const jMd = [31,31,31,31,31,31,30,30,30,30,30,29];
  let jm = 0;
  for (; jm < 12 && jDno >= jMd[jm]; jm++) jDno -= jMd[jm];
  return `${jDno + 1} ${months[jm]} ${jy}`;
}

interface Advertiser { name: string; ownerid: string; manager: string; performanceManager?: string; supervisor?: string; team?: string; summary?: string; note?: string; agencies: Agency[]; industry1?: string; industry2?: string; }
interface Competitor { platform: string; newclients: string; topclients: string; note: string; }
interface AnalysisResult { dateLabel: string; advertisers: Advertiser[]; leads: Advertiser[]; competitors: Competitor[]; market: string; disclaimer: string; id?: string; savedAt?: number; }
interface HistoryEntry { date: string; dateLabel: string; names: string[]; }

function buildMessage(result: AnalysisResult): string {
  let msg = `🟥 از مارکت چه خبر؟ ${result.dateLabel}\n`;
  for (const adv of result.advertisers) { const mgr = adv.manager ? ` | ${adv.manager}` : ""; msg += `\n🔹 **${adv.name} | ${adv.ownerid}${mgr}**\n${adv.summary}\n`; }
  if (result.leads.length > 0) { msg += "\n"; for (const l of result.leads) { const mgr = l.manager ? ` | ${l.manager}` : ""; msg += `\n⭕ **${l.name} | ${l.ownerid}${mgr}**\n${l.note}\n`; } }
  for (const c of result.competitors) { msg += `\n\n🔸 **${c.platform}**\n${c.note}`; if (c.topclients) msg += `\nمهم‌ترین مشتریان: ${c.topclients}`; if (c.newclients && c.newclients !== "ندارد") msg += `\nمشتریان جدید: ${c.newclients}`; }
  if (result.market) msg += `\n\n📊 ${result.market}`;
  msg += `\n\n_${result.disclaimer}_\n\n📍 Powered by Claude`;
  return msg;
}

function generateHTML(result: AnalysisResult): string {
  const agBar = (agencies: Agency[]) => {
    const total = agencies.reduce((s, a) => s + a.value, 0);
    if (!total) return "";
    return `<div style="display:flex;border-radius:4px;overflow:hidden;height:5px;margin:8px 0 4px">${agencies.map(a => `<div style="flex:${a.value};background:${a.color}"></div>`).join("")}</div><div>${agencies.map(a => `<span style="font-size:11px;color:#888;display:inline-flex;align-items:center;gap:4px;margin-left:8px"><span style="width:8px;height:8px;border-radius:50%;background:${a.color};display:inline-block"></span>${a.name} ${Math.round(a.value / total * 100)}٪</span>`).join("")}</div>`;
  };
  const advCards = result.advertisers.map(adv => {
    const total = adv.agencies.reduce((s, a) => s + a.value, 0), ykt = adv.agencies.find(a => a.name === "یکتانت"), yPct = total > 0 && ykt ? Math.round(ykt.value / total * 100) : 0;
    return `<div style="flex:1;background:#fff;border:0.5px solid #D3CFC7;border-radius:12px;padding:12px 14px;margin-bottom:10px"><div style="font-weight:500;font-size:13px;direction:ltr">${adv.name} <span style="color:#888;font-size:11px">#${adv.ownerid}</span></div>${total > 0 ? `<div style="display:flex;align-items:center;gap:8px;margin-top:5px"><div style="flex:1;height:4px;border-radius:2px;background:#EDE8DF;overflow:hidden"><div style="width:${yPct}%;height:100%;background:#D4623A"></div></div><span style="font-size:11px;font-weight:500;color:#D4623A">یکتانت ${yPct}٪</span></div>` : ""}<p style="font-size:13px;color:#5F5E5A;line-height:1.9;margin:8px 0">${adv.summary || ""}</p>${agBar(adv.agencies)}</div>`;
  }).join("");
  return `<!DOCTYPE html><html lang="fa" dir="rtl"><head><meta charset="UTF-8"><title>گزارش بازار — ${result.dateLabel}</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:system-ui,sans-serif;background:#F5F0E8;color:#2C2C2A;direction:rtl;padding:1.5rem 1rem}.c{max-width:720px;margin:0 auto}</style></head><body><div class="c"><h2 style="margin-bottom:1rem">از مارکت چه خبر؟ ${result.dateLabel}</h2>${advCards}<p style="font-size:11px;color:#888;margin-top:1.5rem;font-style:italic">${result.disclaimer}</p></div></body></html>`;
}

function downloadHTML(result: AnalysisResult) {
  const blob = new Blob([generateHTML(result)], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob), a = document.createElement("a");
  a.href = url; a.download = `market-report-${result.dateLabel.replace(/ /g, "-")}.html`; a.click();
  URL.revokeObjectURL(url);
}

// ── Global CSS ────────────────────────────────────────────────────────────────

const GLOBAL_CSS = `
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
html,body{overflow-y:auto;overflow-x:hidden}
@keyframes fu{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
@keyframes radarSweep{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
@keyframes radarPulse{0%{r:5;opacity:.6}100%{r:22;opacity:0}}
@keyframes dotPop{0%,100%{opacity:0;r:0}50%{opacity:1;r:4}}
@keyframes barGrow{from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}
@keyframes shimmer{0%{background-position:200% 0}100%{background-position:-200% 0}}
@keyframes slideInModal{from{opacity:0;transform:translateY(-14px) scale(.97)}to{opacity:1;transform:translateY(0) scale(1)}}
.fu{animation:fu .3s ease both}
`;

// ── Shared UI Components v2 ───────────────────────────────────────────────────

function Sparkline({ data, color, w = 80, h = 24, filled = false }: { data: number[]; color?: string; w?: number; h?: number; filled?: boolean }) {
  const D = useD();
  if (!color) color = D.accent;
  if (!data || data.length < 2) return <div style={{ width: w, height: h }} />;
  const max = Math.max(...data, 1), min = Math.min(...data, 0), rng = max - min || 1;
  const pts = data.map((v, i) => [i / (data.length - 1) * w, h - 2 - ((v - min) / rng) * (h - 4)]);
  const line = "M" + pts.map(p => p.join(",")).join(" L");
  if (filled) {
    const fill = line + ` L${pts[pts.length - 1][0]},${h} L0,${h} Z`;
    return (
      <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ display: "block", overflow: "visible" }}>
        <defs><linearGradient id={`sg${color.replace(/[^a-z0-9]/gi, "")}`} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={color} stopOpacity=".3" /><stop offset="100%" stopColor={color} stopOpacity="0" /></linearGradient></defs>
        <path d={fill} fill={`url(#sg${color.replace(/[^a-z0-9]/gi, "")})`} />
        <path d={line} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round" />
      </svg>
    );
  }
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ display: "block", overflow: "visible" }}>
      <path d={line} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round" />
    </svg>
  );
}

function MiniLineChart({ data, labels, color, h = 80 }: { data: number[]; labels?: string[]; color?: string; h?: number }) {
  const D = useD();
  if (!color) color = D.accent;
  if (!data || data.length < 2) return null;
  const W = 300;
  const max = Math.max(...data, 1), min = Math.min(...data, 0), rng = max - min || 1;
  const pts = data.map((v, i) => [i / (data.length - 1) * (W - 24) + 12, h - 16 - ((v - min) / rng) * (h - 28)]);
  const line = "M" + pts.map(p => p.join(",")).join(" L");
  const fill = line + ` L${pts[pts.length - 1][0]},${h - 4} L12,${h - 4} Z`;
  const gId = `mlg${color.replace(/[^a-z0-9]/gi, "")}`;
  return (
    <svg width={W} height={h} viewBox={`0 0 ${W} ${h}`} style={{ display: "block", width: "100%" }}>
      <defs><linearGradient id={gId} x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={color} stopOpacity=".35" /><stop offset="100%" stopColor={color} stopOpacity="0" /></linearGradient></defs>
      <path d={fill} fill={`url(#${gId})`} />
      <path d={line} fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
      {labels && pts.map((p, i) => <text key={i} x={p[0]} y={h - 2} textAnchor="middle" fontSize="9" fill={D.t3} fontFamily={D.mono}>{labels[i]}</text>)}
      <circle cx={pts[pts.length - 1][0]} cy={pts[pts.length - 1][1]} r="3" fill={color} />
    </svg>
  );
}

function DonutChart({ agencies, size = 72 }: { agencies: Agency[]; size?: number }) {
  const D = useD();
  const r = size * 0.36, cx = size / 2, cy = size / 2, circ = 2 * Math.PI * r;
  const total = agencies.reduce((s, a) => s + a.value, 0);
  if (!total) return <div style={{ width: size, height: size }} />;
  let off = 0;
  const slices = agencies.map(a => {
    const frac = a.value / total, dash = frac * circ, gap = circ - dash, dashoffset = -(off * circ);
    off += frac;
    return { ...a, dash, gap, dashoffset };
  });
  const ykt = agencies.find(a => a.name === "یکتانت");
  const yPct = ykt ? Math.round(ykt.value / total * 100) : 0;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ flexShrink: 0 }}>
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={D.border2} strokeWidth={size * 0.11} />
      {slices.map((s, i) => (
        <circle key={i} cx={cx} cy={cy} r={r} fill="none" stroke={s.color} strokeWidth={size * 0.11}
          strokeDasharray={`${s.dash} ${s.gap}`} strokeDashoffset={s.dashoffset}
          transform={`rotate(-90 ${cx} ${cy})`} strokeLinecap="round" />
      ))}
      <text x={cx} y={cy} textAnchor="middle" dominantBaseline="middle"
        fontSize={size * 0.17} fontWeight="700" fill={ykt ? D.accent : D.t3} fontFamily={D.mono}>
        {yPct}٪
      </text>
    </svg>
  );
}

function AgencyBar({ agencies, h = 4 }: { agencies: Agency[]; h?: number }) {
  const total = agencies.reduce((s, a) => s + a.value, 0);
  if (!total) return null;
  return (
    <div style={{ display: "flex", borderRadius: h, overflow: "hidden", height: h, margin: "4px 0" }}>
      {agencies.map((a, i) => (
        <div key={i} style={{ flex: a.value, background: a.color, minWidth: 1 }} />
      ))}
    </div>
  );
}

function AgencyLegend({ agencies, compact = false }: { agencies: Agency[]; compact?: boolean; T?: Theme }) {
  const D = useD();
  const total = agencies.reduce((s, a) => s + a.value, 0);
  if (!total) return null;
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: compact ? "2px 8px" : "3px 10px", marginTop: 4 }}>
      {agencies.map((a, i) => (
        <span key={i} style={{ fontSize: compact ? 10 : 11, color: D.t3, display: "flex", alignItems: "center", gap: 4 }}>
          <span style={{ width: 6, height: 6, borderRadius: "50%", background: a.color, display: "inline-block", flexShrink: 0 }} />
          {a.name} {Math.round(a.value / total * 100)}٪
        </span>
      ))}
    </div>
  );
}

function Badge({ label, color }: { label: string; color: string }) {
  return (
    <span style={{ fontSize: 9.5, color, background: color + "22", border: `1px solid ${color}44`, padding: "1px 7px", borderRadius: 20, fontWeight: 600, whiteSpace: "nowrap" }}>
      {label}
    </span>
  );
}

function StatusBadge({ status }: { status: string }) {
  const D = useD();
  const MAP: Record<string, { l: string; c: string }> = {
    declining: { l: "کاهشی", c: D.red }, stable: { l: "پایدار", c: D.green },
    growing: { l: "رشد", c: D.green }, reactivated: { l: "بازگشت", c: D.amber },
    inactive: { l: "غیرفعال", c: D.t3 }, lead: { l: "لید", c: D.blue },
  };
  const s = MAP[status] || { l: status, c: D.t3 };
  return <Badge label={s.l} color={s.c} />;
}

function KpiCard({ label, value, sub, color, delta, delay = 0, icon }: { label: string; value: string | number; sub?: string; color?: string; delta?: number; delay?: number; icon?: React.ReactNode }) {
  const D = useD();
  const c = color || D.accent;
  return (
    <div className="fu" style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "14px 16px", animationDelay: `${delay}ms`, position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", top: 0, right: 0, width: 80, height: 80, borderRadius: "0 0 0 80px", background: c, opacity: .06, pointerEvents: "none" }} />
      {icon && <div style={{ color: c, marginBottom: 8, opacity: .7 }}>{icon}</div>}
      <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 600, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 26, fontWeight: 800, color: c, fontFamily: D.mono, lineHeight: 1, letterSpacing: "-1px" }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: D.t3, marginTop: 5 }}>{sub}</div>}
      {delta != null && (
        <div style={{ fontSize: 11, color: delta >= 0 ? D.green : D.red, fontFamily: D.mono, marginTop: 4 }}>
          {delta >= 0 ? "▲" : "▼"} {Math.abs(delta)}٪ هفته گذشته
        </div>
      )}
    </div>
  );
}

function SectionHeader({ title, sub, action }: { title: string; sub?: string; action?: React.ReactNode }) {
  const D = useD();
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, paddingBottom: 10, borderBottom: `1px solid ${D.border}` }}>
      <div>
        <div style={{ fontSize: 15, fontWeight: 700, color: D.t1, letterSpacing: "-.2px" }}>{title}</div>
        {sub && <div style={{ fontSize: 11, color: D.t3, marginTop: 2, fontFamily: D.mono }}>{sub}</div>}
      </div>
      {action}
    </div>
  );
}

function Pill({ label, active, onClick }: { label: string; active: boolean; onClick: () => void }) {
  const D = useD();
  return (
    <button onClick={onClick} style={{ padding: "5px 12px", borderRadius: 99, border: `1px solid ${active ? D.accentBrd : D.border}`, background: active ? D.accentDim : "transparent", color: active ? D.accent : D.t2, fontSize: 11, fontWeight: active ? 600 : 400, cursor: "pointer", whiteSpace: "nowrap", transition: "all .14s", fontFamily: "Vazirmatn,sans-serif" }}>
      {label}
    </button>
  );
}

function RadarHero() {
  const D = useD();
  return (
    <div style={{ position: "relative", width: 180, height: 180, margin: "0 auto 20px" }}>
      <svg width="180" height="180" viewBox="0 0 200 200">
        <circle cx="100" cy="100" r="85" fill="none" stroke={D.accent} strokeWidth="0.5" opacity="0.15" />
        <circle cx="100" cy="100" r="60" fill="none" stroke={D.accent} strokeWidth="0.5" opacity="0.22" />
        <circle cx="100" cy="100" r="36" fill="none" stroke={D.accent} strokeWidth="0.5" opacity="0.32" />
        <line x1="100" y1="15" x2="100" y2="185" stroke={D.accent} strokeWidth="0.4" opacity="0.1" />
        <line x1="15" y1="100" x2="185" y2="100" stroke={D.accent} strokeWidth="0.4" opacity="0.1" />
        <g style={{ transformOrigin: "100px 100px", animation: "radarSweep 3s linear infinite" }}>
          <path d="M100,100 L100,15 A85,85 0 0,1 185,100 Z" fill={D.accent} opacity="0.08" />
          <line x1="100" y1="100" x2="100" y2="15" stroke={D.accent} strokeWidth="1.5" opacity="0.6" />
        </g>
        <circle cx="142" cy="68" r="4" fill={D.accent} style={{ animation: "dotPop 3s 0.9s ease-in-out infinite" }} />
        <circle cx="72" cy="140" r="3" fill={D.green} style={{ animation: "dotPop 3s 1.7s ease-in-out infinite" }} />
        <circle cx="158" cy="128" r="2.5" fill={D.amber} style={{ animation: "dotPop 3s 2.3s ease-in-out infinite" }} />
        <circle cx="100" cy="100" r="5" fill={D.accent} />
        <circle cx="100" cy="100" r="5" fill={D.accent} opacity="0.6" style={{ animation: "radarPulse 2.4s ease-out infinite" }} />
      </svg>
    </div>
  );
}

function Toast({ msg, type, onDone }: { msg: string; type: "success" | "error"; onDone: () => void }) {
  const D = useD();
  useEffect(() => { const t = setTimeout(onDone, 3000); return () => clearTimeout(t); }, [onDone]);
  return (
    <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: type === "success" ? D.green : D.red, color: "#fff", padding: "10px 20px", borderRadius: 12, fontSize: 13, fontWeight: 500, zIndex: 9999, animation: "fadeUp 0.2s ease both", whiteSpace: "nowrap", boxShadow: "0 4px 20px rgba(0,0,0,0.3)" }}>
      {msg}
    </div>
  );
}

function DetailModal({ adv, type, onClose, onRegen, regenLoading }: { adv: Advertiser; type: string; onClose: () => void; onRegen: () => void; regenLoading: boolean }) {
  const D = useD();
  useEffect(() => {
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", esc);
    return () => document.removeEventListener("keydown", esc);
  }, [onClose]);
  const total = adv.agencies.reduce((s, a) => s + a.value, 0);
  return (
    <div role="dialog" aria-modal="true" aria-label={`جزئیات ${adv.name}`} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }} onClick={onClose}>
      <div style={{ background: D.card, border: `1px solid ${D.border2}`, borderRadius: 20, padding: "24px 28px", width: "100%", maxWidth: 480, animation: "slideInModal 0.22s ease both", direction: "rtl" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 600, color: D.t1, direction: "ltr" }}>{adv.name}</div>
            <div style={{ fontSize: 12, color: D.t3, marginTop: 3, direction: "ltr" }}>#{adv.ownerid}{adv.manager ? ` · ` : ""}{adv.manager && <span style={{ color: D.accent }}>{adv.manager}</span>}</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onRegen} disabled={regenLoading} aria-label="بازسازی تحلیل" style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${D.border}`, background: "transparent", color: D.t2, fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 5 }}>
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={regenLoading ? { animation: "spin 0.8s linear infinite" } : {}}><polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10" /></svg>
              بازسازی
            </button>
            <button onClick={onClose} aria-label="بستن" style={{ width: 40, height: 40, borderRadius: 8, border: `1px solid ${D.border}`, background: "transparent", color: D.t3, cursor: "pointer", fontSize: 20, lineHeight: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>×</button>
          </div>
        </div>
        {total > 0 && (
          <div style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 24, marginBottom: 14 }}>
              <DonutChart agencies={adv.agencies} size={100} />
              <AgencyLegend agencies={adv.agencies} />
            </div>
            <AgencyBar agencies={adv.agencies} h={6} />
          </div>
        )}
        {(adv.industry1 || adv.industry2) && (
          <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
            {adv.industry1 && <Badge label={adv.industry1} color={D.blue} />}
            {adv.industry2 && <Badge label={adv.industry2} color={D.t3} />}
          </div>
        )}
        {(adv.team || adv.manager || adv.performanceManager || adv.supervisor) && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 14 }}>
            {adv.team && <div style={{ padding: "8px 12px", borderRadius: 10, background: D.surface, border: `1px solid ${D.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: D.t3 }}>تیم</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: D.t1 }}>{adv.team}</p></div>}
            {adv.manager && <div style={{ padding: "8px 12px", borderRadius: 10, background: D.surface, border: `1px solid ${D.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: D.t3 }}>اکانت منیجر</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: D.accent }}>{adv.manager}</p></div>}
            {adv.performanceManager && <div style={{ padding: "8px 12px", borderRadius: 10, background: D.surface, border: `1px solid ${D.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: D.t3 }}>پرفورمنس منیجر</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: D.blue }}>{adv.performanceManager}</p></div>}
            {adv.supervisor && <div style={{ padding: "8px 12px", borderRadius: 10, background: D.surface, border: `1px solid ${D.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: D.t3 }}>سوپروایزر</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: D.amber }}>{adv.supervisor}</p></div>}
          </div>
        )}
        <div style={{ background: D.surface, borderRadius: 12, padding: "14px 16px" }}>
          <p style={{ margin: 0, fontSize: 14, color: D.t2, lineHeight: 2, direction: "rtl" }}>
            {type === "lead" ? adv.note : adv.summary}
          </p>
        </div>
      </div>
    </div>
  );
}

// ── Login Page ────────────────────────────────────────────────────────────────

function LoginPage({ onLogin }: { onLogin: (u: SessionUser) => void }) {
  const T = useD();
  const [selected, setSelected] = useState<SessionUser | null>(null);
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const [showPw, setShowPw] = useState(false);

  const allUsers: SessionUser[] = [
    ...TEAMS.map(t => ({ username: t.username, managerFa: t.managerFa, managerEn: t.managerEn, teamFa: t.teamFa, teamEn: t.teamEn, isAdmin: false })),
    { username: ADMIN_USER.username, managerFa: ADMIN_USER.managerFa, managerEn: ADMIN_USER.managerEn, teamFa: ADMIN_USER.teamFa, teamEn: ADMIN_USER.teamEn, isAdmin: true },
  ];

  const byDept: Record<string, SessionUser[]> = {};
  TEAMS.forEach((t, i) => {
    if (!byDept[t.dept]) byDept[t.dept] = [];
    byDept[t.dept].push(allUsers[i]);
  });

  const handleLogin = () => {
    if (!selected) { setErr("لطفاً نام خود را انتخاب کنید"); return; }
    if (pw !== APP_PASSWORD) { setErr("رمز عبور اشتباه است"); return; }
    localStorage.setItem(SESSION_KEY, JSON.stringify({ ...selected, loginAt: Date.now() }));
    onLogin(selected);
  };

  const deptColors: Record<string, string> = { "بیزینس": T.accent, "محصول": T.blue, "فروش": T.green, "": T.t3 };

  return (
    <div style={{ minHeight: "100dvh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, direction: "rtl", fontFamily: "'Vazirmatn', system-ui, sans-serif" }}>
      <div style={{ width: "100%", maxWidth: 560, animation: "fadeUp 0.3s ease both" }}>
        {/* Header */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ width: 52, height: 52, borderRadius: 16, background: T.coral, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 14px" }}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="10" stroke="#fff" strokeWidth="1.5" /><circle cx="12" cy="12" r="6" stroke="#fff" strokeWidth="1" opacity="0.6" /><circle cx="12" cy="12" r="2.5" fill="#fff" /></svg>
          </div>
          <h1 style={{ margin: "0 0 6px", fontSize: 24, fontWeight: 700, color: T.text1 }}>Radar</h1>
          <p style={{ margin: 0, fontSize: 14, color: T.text3 }}>آنالیز رقابتی روزانه یکتانت</p>
        </div>

        {/* Manager grid */}
        <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 20, padding: 20, marginBottom: 16 }}>
          <p style={{ margin: "0 0 14px", fontSize: 12, fontWeight: 600, color: T.text3, textTransform: "uppercase", letterSpacing: "0.08em" }}>نام خود را انتخاب کنید</p>
          {Object.entries(byDept).map(([dept, users]) => (
            <div key={dept} style={{ marginBottom: 16 }}>
              <p style={{ margin: "0 0 8px", fontSize: 11, color: deptColors[dept] || T.text3, fontWeight: 600 }}>{dept}</p>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {users.map(u => {
                  const active = selected?.username === u.username;
                  return (
                    <button key={u.username} onClick={() => { setSelected(u); setErr(""); }}
                      style={{ padding: "8px 14px", borderRadius: 10, border: `1.5px solid ${active ? T.coral : T.border}`, background: active ? T.coralDim : "transparent", color: active ? T.coral : T.text2, fontSize: 13, fontWeight: active ? 600 : 400, cursor: "pointer", transition: "all 0.15s" }}>
                      {u.managerFa}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
          {/* Admin */}
          <div style={{ borderTop: `1px solid ${T.border}`, paddingTop: 12, marginTop: 4 }}>
            {(() => {
              const adminUser = allUsers[allUsers.length - 1];
              const active = selected?.username === "admin";
              return (
                <button onClick={() => { setSelected(adminUser); setErr(""); }}
                  style={{ padding: "8px 14px", borderRadius: 10, border: `1.5px solid ${active ? T.coral : T.border}`, background: active ? T.coralDim : "transparent", color: active ? T.coral : T.text3, fontSize: 13, fontWeight: active ? 600 : 400, cursor: "pointer", transition: "all 0.15s" }}>
                  ادمین (همه تیم‌ها)
                </button>
              );
            })()}
          </div>
        </div>

        {/* Password */}
        <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16, padding: 20 }}>
          <p style={{ margin: "0 0 10px", fontSize: 12, fontWeight: 600, color: T.text3, textTransform: "uppercase", letterSpacing: "0.08em" }}>رمز عبور</p>
          <div style={{ display: "flex", gap: 10 }}>
            <div style={{ flex: 1, position: "relative" }}>
              <input type={showPw ? "text" : "password"} value={pw} onChange={e => { setPw(e.target.value); setErr(""); }}
                onKeyDown={e => e.key === "Enter" && handleLogin()}
                placeholder="رمز عبور را وارد کنید"
                style={{ width: "100%", fontSize: 14, padding: "11px 14px", paddingLeft: 40, borderRadius: 11, border: `1px solid ${err ? T.danger : T.border}`, background: T.surface2, color: T.text1, outline: "none", direction: "rtl" }} />
              <button onClick={() => setShowPw(v => !v)} aria-label={showPw ? "مخفی کردن رمز عبور" : "نمایش رمز عبور"}
                style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: T.text3, padding: 4, display: "flex", minWidth: 32, minHeight: 32, alignItems: "center", justifyContent: "center" }}>
                {showPw
                  ? <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94"/><path d="M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19"/><line x1="1" y1="1" x2="23" y2="23"/></svg>
                  : <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>}
              </button>
            </div>
            <button onClick={handleLogin}
              style={{ padding: "11px 22px", borderRadius: 11, border: "none", background: T.coral, color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>
              ورود
            </button>
          </div>
          {err && <p style={{ margin: "8px 0 0", fontSize: 12, color: T.danger }}>{err}</p>}
          {selected && <p style={{ margin: "10px 0 0", fontSize: 12, color: T.text3 }}>
            ورود به عنوان <span style={{ color: T.coral, fontWeight: 600 }}>{selected.managerFa}</span>
            {!selected.isAdmin && <span> — تیم {selected.teamFa}</span>}
          </p>}
        </div>
      </div>
    </div>
  );
}

// ── Sidebar v2 ────────────────────────────────────────────────────────────────

const NAV_ITEMS = [
  { id: "upload",    fa: "آنالیز جدید",   icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg> },
  { id: "dashboard", fa: "داشبورد",        icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg> },
  { id: "results",   fa: "آنالیز",         icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg> },
  { id: "competitor",fa: "رقبا",           icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg> },
  { id: "marketmap", fa: "نقشه بازار",     icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="4"/><line x1="12" y1="2" x2="12" y2="8"/><line x1="12" y1="16" x2="12" y2="22"/><line x1="2" y1="12" x2="8" y2="12"/><line x1="16" y1="12" x2="22" y2="12"/></svg> },
  { id: "trends",    fa: "روندها",         icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg> },
  { id: "leads",     fa: "لیدها",          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg> },
  { id: "industry",  fa: "صنایع",          icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 00-2-2h-4a2 2 0 00-2 2v2"/></svg> },
  { id: "team",      fa: "تیم",            icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg> },
  { id: "alerts",    fa: "هشدارها",        icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg> },
  { id: "brief",     fa: "خلاصه هفتگی",   icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg> },
  { id: "explorer",  fa: "دیتا",           icon: <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/></svg> },
];

const PAGE_NAMES: Record<string, string> = {
  upload: "آنالیز جدید", dashboard: "داشبورد", results: "نتایج آنالیز",
  competitor: "بررسی رقبا", marketmap: "نقشه بازار", trends: "تحلیل روند",
  leads: "لید پایپلاین", profile: "پروفایل", industry: "تحلیل صنایع",
  team: "عملکرد تیم", alerts: "مرکز هشدار", brief: "خلاصه هفتگی",
  explorer: "اکسپلورر داده", reports: "گزارش‌ها", history: "تاریخچه", settings: "تنظیمات",
};

function Sidebar({ screen, setScreen, isDark, setIsDark, hasResult, hasCsv, session, onLogout, alertCount, T }: {
  screen: string; setScreen: (s: string) => void; isDark: boolean; setIsDark: (v: boolean) => void;
  hasResult: boolean; hasCsv: boolean; session: SessionUser; onLogout: () => void; alertCount: number; T: Theme;
}) {
  const D = useD();
  const enabled = (id: string) => {
    if (id === "upload") return true;
    if (id === "results") return hasResult;
    if (["dashboard", "marketmap", "trends", "industry", "team", "explorer",
         "competitor", "leads", "alerts", "brief"].includes(id)) return hasCsv;
    return true;
  };

  return (
    <div style={{ position: "fixed", right: 0, top: 0, bottom: 0, width: 200, background: D.surface, borderLeft: `1px solid ${D.border}`, display: "flex", flexDirection: "column", zIndex: 200, overflowY: "auto" }}>
      {/* Logo */}
      <div style={{ padding: "18px 16px 14px", borderBottom: `1px solid ${D.border}`, flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 34, height: 34, borderRadius: 10, background: `linear-gradient(135deg,${D.accent},${D.accentHov})`, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: `0 4px 14px ${D.accentGlow}`, flexShrink: 0 }}>
            <svg width="17" height="17" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="#fff" strokeWidth="1.6"/><circle cx="12" cy="12" r="5" stroke="#fff" strokeWidth="1.3" opacity=".55"/><circle cx="12" cy="12" r="1.8" fill="#fff"/><line x1="12" y1="3" x2="12" y2="12" stroke="#fff" strokeWidth="1.6"/></svg>
          </div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>Radar</div>
            <div style={{ fontSize: 9.5, color: D.t3, fontFamily: D.mono }}>یکتانت · v2</div>
          </div>
        </div>
      </div>

      {/* Nav */}
      <nav style={{ flex: 1, padding: "8px" }}>
        {NAV_ITEMS.map(item => {
          const active = screen === item.id;
          const on = enabled(item.id);
          return (
            <button key={item.id} onClick={() => on && setScreen(item.id)}
              aria-label={item.fa} aria-current={active ? "page" : undefined}
              style={{ display: "flex", alignItems: "center", gap: 9, width: "100%", padding: "8px 10px", borderRadius: 9, border: "none", cursor: on ? "pointer" : "default", marginBottom: 1,
                background: active ? D.accentDim : "transparent",
                color: active ? D.accent : on ? D.t2 : D.t4,
                fontSize: 12, fontFamily: "Vazirmatn,sans-serif", fontWeight: active ? 600 : 400,
                borderRight: active ? `2.5px solid ${D.accent}` : "2.5px solid transparent",
                transition: "all .12s", position: "relative" }}>
              <span style={{ flexShrink: 0, opacity: active ? 1 : on ? .7 : .35 }}>{item.icon}</span>
              <span style={{ flex: 1, textAlign: "right" }}>{item.fa}</span>
              {item.id === "alerts" && alertCount > 0 && (
                <span style={{ fontSize: 9, fontWeight: 800, background: D.red, color: "#fff", borderRadius: 99, padding: "1px 5px", fontFamily: D.mono, minWidth: 16, textAlign: "center" }}>{alertCount}</span>
              )}
              {active && <div style={{ width: 5, height: 5, borderRadius: "50%", background: D.accent, flexShrink: 0 }} />}
            </button>
          );
        })}
      </nav>

      {/* Bottom: user + controls */}
      <div style={{ padding: "10px 12px 14px", borderTop: `1px solid ${D.border}`, flexShrink: 0 }}>
        {/* User card */}
        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", borderRadius: 9, background: D.accentDim, border: `1px solid ${D.accentBrd}`, marginBottom: 8 }}>
          <div style={{ width: 26, height: 26, borderRadius: "50%", background: D.accent, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 800, color: "#fff", flexShrink: 0 }}>
            {session.managerFa.slice(0, 1)}
          </div>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: D.t1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{session.managerFa}</div>
            <div style={{ fontSize: 9.5, color: D.t3, fontFamily: D.mono, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{session.isAdmin ? "همه تیم‌ها" : session.teamFa}</div>
          </div>
        </div>
        {/* Theme + Logout row */}
        <div style={{ display: "flex", gap: 6 }}>
          <button onClick={() => setIsDark(!isDark)} title={isDark ? "روشن" : "تاریک"} style={{ flex: 1, height: 32, borderRadius: 8, border: `1px solid ${D.border}`, background: "transparent", color: D.t3, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
            {isDark
              ? <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/></svg>
              : <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>}
          </button>
          <button onClick={onLogout} title="خروج" style={{ flex: 1, height: 32, borderRadius: 8, border: `1px solid ${D.redBrd}`, background: D.redDim, color: D.red, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Topbar v2 ─────────────────────────────────────────────────────────────────

function Topbar({ screen, yktShare, alertCount, setScreen }: { screen: string; yktShare: number | null; alertCount: number; setScreen: (s: string) => void }) {
  const D = useD();
  const jalali = toJalali(new Date());
  const pageName = PAGE_NAMES[screen] || screen;
  return (
    <div style={{ position: "fixed", top: 0, right: 200, left: 0, height: 52, display: "flex", alignItems: "center", padding: "0 24px", background: D.surface, borderBottom: `1px solid ${D.border}`, zIndex: 100, gap: 14 }}>
      {/* Breadcrumb */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, flex: 1, minWidth: 0 }}>
        <span style={{ fontSize: 11, color: D.t3, fontFamily: D.mono }}>Radar</span>
        <span style={{ fontSize: 11, color: D.t4 }}>›</span>
        <span style={{ fontSize: 13, fontWeight: 700, color: D.t1 }}>{pageName}</span>
      </div>
      {/* Jalali date */}
      <div style={{ fontSize: 11, color: D.t3, background: D.card, border: `1px solid ${D.border}`, borderRadius: 8, padding: "4px 10px", fontFamily: D.mono, flexShrink: 0 }}>
        {jalali}
      </div>
      {/* Ykt market share pill */}
      {yktShare !== null && (
        <div style={{ display: "flex", alignItems: "center", gap: 6, background: D.accentDim, border: `1px solid ${D.accentBrd}`, borderRadius: 8, padding: "4px 10px", flexShrink: 0 }}>
          <span style={{ fontSize: 10, color: D.t3 }}>سهم یکتانت</span>
          <span style={{ fontSize: 13, fontWeight: 800, color: D.accent, fontFamily: D.mono }}>{yktShare}٪</span>
        </div>
      )}
      {/* Alert bell */}
      <button onClick={() => setScreen("alerts")} style={{ position: "relative", width: 36, height: 36, borderRadius: 9, border: `1px solid ${D.border}`, background: "transparent", color: D.t2, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg>
        {alertCount > 0 && <span style={{ position: "absolute", top: 4, right: 4, width: 8, height: 8, borderRadius: "50%", background: D.red }} />}
      </button>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────

export default function App() {
  const [session, setSession] = useState<SessionUser | null>(() => {
    try {
      const raw = localStorage.getItem(SESSION_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw) as SessionUser & { loginAt?: number };
      // Expire after 7 days
      if (parsed.loginAt && Date.now() - parsed.loginAt > 7 * 86400 * 1000) {
        localStorage.removeItem(SESSION_KEY); return null;
      }
      return parsed;
    } catch { return null; }
  });

  const handleLogout = () => {
    localStorage.removeItem(SESSION_KEY);
    setSession(null);
  };

  const [isDark, setIsDark] = useState(() => {
    try { return localStorage.getItem(THEME_KEY) === "dark"; } catch { return false; }
  });
  const T = useMemo(() => makeTokens(isDark), [isDark]);

  useEffect(() => { try { localStorage.setItem(THEME_KEY, isDark ? "dark" : "light"); } catch { } }, [isDark]);

  const [screen, setScreen] = useState("upload");
  const [step, setStep] = useState<"upload" | "ready" | "loading" | "preview">("upload");
  const [csvData, setCsvData] = useState<{ text: string; rows: Record<string, string>[]; name: string; stats: ReturnType<typeof getStats> } | null>(null);
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [pendingSave, setPendingSave] = useState<{ date: string; dateLabel: string; names: string[] } | null>(null);
  const [saveStatus, setSaveStatus] = useState("");
  const [editableMsg, setEditableMsg] = useState("");
  const [regenLoading, setRegenLoading] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [errMsg, setErrMsg] = useState("");
  const [rawDebug, setRawDebug] = useState("");
  const [showDebug, setShowDebug] = useState(false);
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [savedReports, setSavedReports] = useState<AnalysisResult[]>([]);
  const [searchQuery, setSearchQuery] = useState("");
  const [filterAgency, setFilterAgency] = useState("all");
  const [sortBy, setSortBy] = useState("importance");
  const [selectedTemplate, setSelectedTemplate] = useState("standard");
  const [resultsTab, setResultsTab] = useState<"advertisers" | "leads" | "competitors" | "industry">("advertisers");
  const [modalAdv, setModalAdv] = useState<{ adv: Advertiser; type: string } | null>(null);
  const [toast, setToast] = useState<{ msg: string; type: "success" | "error" } | null>(null);
  const [loadingProgress, setLoadingProgress] = useState(0);
  const loadingTimer = useRef<ReturnType<typeof setInterval> | null>(null);
  const [apiKeyInput, setApiKeyInput] = useState(() => localStorage.getItem("radar:openrouter_key") || "");
  const [showApiKeyInput, setShowApiKeyInput] = useState(false);
  const hasApiKey = !!(OPENROUTER_KEY || localStorage.getItem("radar:openrouter_key"));

  useEffect(() => { loadHistory(); loadReports(); }, []);

  const showToast = (msg: string, type: "success" | "error" = "success") => {
    setToast({ msg, type });
  };

  const loadHistory = () => {
    try { const r = localStorage.getItem(HISTORY_KEY); if (r) setHistory(JSON.parse(r)); } catch { setHistory([]); }
  };
  const loadReports = () => {
    try { const r = localStorage.getItem(REPORTS_KEY); if (r) setSavedReports(JSON.parse(r)); } catch { setSavedReports([]); }
  };
  const persistHistory = (u: HistoryEntry[]) => {
    if (u.length === 0) localStorage.removeItem(HISTORY_KEY);
    else localStorage.setItem(HISTORY_KEY, JSON.stringify(u));
    setHistory(u);
  };
  const confirmSave = () => {
    if (!pendingSave) return;
    try {
      const u = [...history.filter(h => h.date !== pendingSave.date), pendingSave].slice(-5);
      persistHistory(u);
      const rep: AnalysisResult = { ...result!, savedAt: Date.now(), id: pendingSave.date };
      const reps = [...savedReports.filter(r => r.id !== rep.id), rep].slice(-10);
      localStorage.setItem(REPORTS_KEY, JSON.stringify(reps)); setSavedReports(reps);
      setPendingSave(null); setSaveStatus("saved");
      showToast("گزارش ذخیره شد ✓");
    } catch { showToast("خطا در ذخیره", "error"); }
  };
  const deleteHistoryEntry = (date: string) => { try { persistHistory(history.filter(h => h.date !== date)); } catch { } };
  const clearHistory = () => { try { persistHistory([]); } catch { } };
  const uniqueReported = [...new Set(history.flatMap(h => h.names))];

  const handleFile = useCallback((file: File) => {
    if (!file) return;
    readFile(file, (rows, csvText) => { setCsvData({ text: csvText, rows, name: file.name, stats: getStats(rows) }); setStep("ready"); setErrMsg(""); }, () => setErrMsg("خطا در خواندن فایل."));
  }, []);
  const onDrop = useCallback((e: React.DragEvent) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }, [handleFile]);

  const buildManagerMap = (rows: Record<string, string>[]): Record<string, string> => {
    const map: Record<string, string> = {};
    rows.forEach(r => {
      const n = (r.Advertiser_name || r.owner_name || "").toString().trim();
      const m = (r.Account_manager_name || r.account_manager_name || "").toString().trim();
      if (n && m) map[n] = m;
    });
    return map;
  };

  const runAnalysis = async (csvText: string, rows: Record<string, string>[], template = "standard", extra = ""): Promise<AnalysisResult> => {
    const managerMap = buildManagerMap(rows);
    const preview = csvText.length > 15000 ? csvText.slice(0, 15000) + "\n...truncated" : csvText;
    const prevBlock = uniqueReported.length > 0 ? `\n\nPREVIOUSLY REPORTED (skip unless new):\n${uniqueReported.join(", ")}` : "";
    const tmplInstr = TEMPLATE_INSTRUCTIONS[template] ? `\nTEMPLATE: ${TEMPLATE_INSTRUCTIONS[template]}` : "";
    const extraInstr = extra ? `\nADDITIONAL: ${extra}` : "";
    const apiKey = OPENROUTER_KEY || localStorage.getItem("radar:openrouter_key") || "";
    if (!apiKey) throw new Error("کلید API تنظیم نشده. لطفاً کلید OpenRouter را وارد کنید.");
    const resp = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`,
        "HTTP-Referer": "https://raufyektanet-cell.github.io/Radar/",
      },
      body: JSON.stringify({ model: MODEL, max_tokens: 4000, stream: true, messages: [{ role: "system", content: SYSTEM_PROMPT }, { role: "user", content: `Data:\n\n${preview}${prevBlock}${tmplInstr}${extraInstr}\n\nToday is ${todayLabel()}. Start with ##. No preamble.` }] })
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}: ${await resp.text().catch(() => "")}`);
    const reader = resp.body!.getReader();
    const decoder = new TextDecoder();
    let raw = "";
    let buf = "";
    outer: while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });
      const lines = buf.split("\n");
      buf = lines.pop() ?? "";
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith("data: ")) continue;
        const payload = trimmed.slice(6);
        if (payload === "[DONE]") break outer;
        try { const chunk = JSON.parse(payload); const delta = chunk.choices?.[0]?.delta?.content; if (delta) raw += delta; } catch { }
      }
    }
    setRawDebug(raw);
    if (!raw || raw.length < 5) throw new Error("پاسخ خالی از API");
    // Build ref map from uploaded CSV rows — includes industry + all personnel fields
    type LocalRef = { industry1: string; industry2: string; team: string; accountManager: string; performanceManager: string; supervisor: string; };
    const localRef = new Map<string, LocalRef>();
    const nameRef = new Map<string, LocalRef>();
    rows.forEach(r => {
      const id = String(r.Owner_id || r.owner_id || "").trim();
      const entry: LocalRef = {
        industry1: r.Category_level_1 || r.industry_tag_1 || "",
        industry2: r.Category_level_2 || r.industry_tag_2 || "",
        team: r.Team || "",
        accountManager: r.Account_manager_name || r.account_manager_name || "",
        performanceManager: r.Performance_manager_name || "",
        supervisor: r.Supervisor_name || "",
      };
      if (id) localRef.set(id, entry);
      const name = (r.Advertiser_name || r.owner_name || "").trim();
      if (name && !nameRef.has(name)) nameRef.set(name, entry);
    });
    const enrich = (ownerid: string, name?: string): Partial<Advertiser> => {
      const ref = localRef.get(ownerid) || (name ? nameRef.get(name) : undefined);
      return {
        industry1: ref?.industry1 || "",
        industry2: ref?.industry2 || "",
        team: ref?.team || "",
        performanceManager: ref?.performanceManager || "",
        supervisor: ref?.supervisor || "",
      };
    };
    const advertisers = extractBlocks("ADVERTISER", raw).map(b => { const f = parseFields(b); const n = f.NAME || ""; const id = f.OWNERID || ""; const ex = enrich(id, n); return { name: n, ownerid: id, manager: managerMap[n] || ex.team || "", summary: f.SUMMARY || "", agencies: parseAgencies(f.AGENCIES), ...ex }; });
    const leads = extractBlocks("LEAD", raw).map(b => { const f = parseFields(b); const n = f.NAME || ""; const id = f.OWNERID || ""; const ex = enrich(id, n); return { name: n, ownerid: id, manager: managerMap[n] || "", note: f.NOTE || "", agencies: parseAgencies(f.AGENCIES), ...ex }; });
    const competitors = extractBlocks("COMPETITOR", raw).map(b => { const f = parseFields(b); return { platform: f.PLATFORM || "", newclients: f.NEWCLIENTS || "", topclients: f.TOPCLIENTS || "", note: f.NOTE || "" }; });
    const market = (extractBlocks("MARKET", raw)[0] || "").trim();
    const disclaimer = (extractBlocks("DISCLAIMER", raw)[0] || "داده‌ها بر اساس کراول وب هستن و تخمینی‌اند — روندها قابل اعتمادند اما اعداد دقیق نیستن.").trim();
    return { dateLabel: todayLabel(), advertisers, leads, competitors, market, disclaimer };
  };

  const startLoadingProgress = () => {
    setLoadingProgress(0);
    loadingTimer.current = setInterval(() => {
      setLoadingProgress(p => {
        if (p >= 88) { if (loadingTimer.current) clearInterval(loadingTimer.current); return 88; }
        return p + (88 - p) * 0.04 + 0.5;
      });
    }, 300);
  };
  const stopLoadingProgress = () => {
    if (loadingTimer.current) clearInterval(loadingTimer.current);
    setLoadingProgress(100);
  };

  const filterRowsByManager = (rows: Record<string, string>[]): Record<string, string>[] => {
    if (!session || session.isAdmin) return rows;
    const fa = session.managerFa.trim();
    // Primary: match rows where user appears as AM, PM, or Supervisor by Persian name
    const byRole = rows.filter(r => {
      const am = (r.Account_manager_name || "").trim();
      const pm = (r.Performance_manager_name || "").trim();
      const sup = (r.Supervisor_name || "").trim();
      return am === fa || pm === fa || sup === fa;
    });
    if (byRole.length > 0) return byRole;
    // Fallback: English last name partial match across all three role columns
    const lastName = session.managerEn.split(" ").pop()?.toLowerCase() || "";
    if (lastName) {
      const byName = rows.filter(r => {
        const am = (r.Account_manager_name || "").toLowerCase();
        const pm = (r.Performance_manager_name || "").toLowerCase();
        const sup = (r.Supervisor_name || "").toLowerCase();
        return am.includes(lastName) || pm.includes(lastName) || sup.includes(lastName);
      });
      if (byName.length > 0) return byName;
    }
    // Last resort: team-based
    const byTeamFa = rows.filter(r => (r.Team || "").trim() === session.teamFa);
    return byTeamFa.length > 0 ? byTeamFa : rows;
  };

  const analyze = async () => {
    setErrMsg(""); setRawDebug(""); setShowDebug(false); setStep("loading");
    startLoadingProgress();
    try {
      const filteredRows = filterRowsByManager(csvData!.rows);
      const filteredCsv = filteredRows.map(r => Object.values(r).join(",")).join("\n");
      const csvHeader = Object.keys(filteredRows[0] || {}).join(",");
      const parsed = await runAnalysis(csvHeader + "\n" + filteredCsv, filteredRows, selectedTemplate);
      stopLoadingProgress();
      if (!parsed.advertisers.length && !parsed.leads.length && !parsed.competitors.length) {
        setErrMsg("هیچ بلاکی پارس نشد — پاسخ خام در دیباگ"); setShowDebug(true); setStep("ready"); return;
      }
      setResult(parsed); setEditableMsg(buildMessage(parsed));
      setPendingSave({ date: csvData!.stats.lastDate, dateLabel: parsed.dateLabel, names: [...parsed.advertisers.map(a => a.name), ...parsed.leads.map(l => l.name)] });
      setSaveStatus(""); setStep("preview"); setScreen("results"); setResultsTab("advertisers"); setSearchQuery(""); setFilterAgency("all");
    } catch (e) {
      stopLoadingProgress(); setErrMsg("خطا: " + (e as Error).message); setStep("ready");
      showToast("خطا در دریافت پاسخ", "error");
    }
  };

  const regenOne = async (item: Advertiser, type: string) => {
    if (!csvData) return;
    setRegenLoading(item.name);
    try {
      const instr = `Only analyze "${item.name}" (id:${item.ownerid}). Output only one ${type === "lead" ? "##LEAD##" : "##ADVERTISER##"} block and nothing else.`;
      const parsed = await runAnalysis(csvData.text, csvData.rows, selectedTemplate, instr);
      const newItems = type === "lead" ? parsed.leads : parsed.advertisers;
      if (newItems.length > 0) {
        const updated = type === "lead" ? result!.leads.map(l => l.name === item.name ? newItems[0] : l) : result!.advertisers.map(a => a.name === item.name ? newItems[0] : a);
        const nr: AnalysisResult = type === "lead" ? { ...result!, leads: updated } : { ...result!, advertisers: updated };
        setResult(nr); setEditableMsg(buildMessage(nr)); setModalAdv(null);
      }
    } catch (e) { showToast("خطا در بازسازی", "error"); console.error(e); }
    finally { setRegenLoading(null); }
  };

  const reset = () => { setCsvData(null); setResult(null); setPendingSave(null); setStep("upload"); setErrMsg(""); setRawDebug(""); setShowDebug(false); setSaveStatus(""); setSearchQuery(""); setFilterAgency("all"); setScreen("upload"); };

  const saveApiKey = () => {
    const trimmed = apiKeyInput.trim();
    if (trimmed) { localStorage.setItem("radar:openrouter_key", trimmed); showToast("کلید API ذخیره شد ✓"); }
    else { localStorage.removeItem("radar:openrouter_key"); showToast("کلید API حذف شد"); }
    setShowApiKeyInput(false);
  };

  const allAgencies = useMemo(() => { if (!result) return []; const s = new Set<string>(); result.advertisers.forEach(a => a.agencies.forEach(ag => s.add(ag.name))); return [...s]; }, [result]);
  const filteredAdvertisers = useMemo(() => {
    if (!result) return [];
    let items = [...result.advertisers];
    if (searchQuery) items = items.filter(a => a.name.toLowerCase().includes(searchQuery.toLowerCase()) || a.ownerid.includes(searchQuery));
    if (filterAgency !== "all") items = items.filter(a => a.agencies.some(ag => ag.name === filterAgency));
    if (sortBy === "volume") items.sort((a, b) => b.agencies.reduce((s, x) => s + x.value, 0) - a.agencies.reduce((s, x) => s + x.value, 0));
    else if (sortBy === "yektanet") { items.sort((a, b) => { const at = a.agencies.reduce((s, x) => s + x.value, 0) || 1, bt = b.agencies.reduce((s, x) => s + x.value, 0) || 1; const ay = a.agencies.find(x => x.name === "یکتانت"), by = b.agencies.find(x => x.name === "یکتانت"); return ((by?.value || 0) / bt) - ((ay?.value || 0) / at); }); }
    return items;
  }, [result, searchQuery, filterAgency, sortBy]);

  const card = (extra: React.CSSProperties = {}): React.CSSProperties => ({ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16, ...extra });
  const btnStyle = (extra: React.CSSProperties = {}): React.CSSProperties => ({ padding: "8px 16px", borderRadius: 10, border: `1px solid ${T.border}`, background: "transparent", color: T.text2, fontSize: 12, cursor: "pointer", ...extra });

  // ─ Advertiser card ─────────────────────────────────────────────────────────
  const AdvCard = ({ adv, type = "advertiser" }: { adv: Advertiser; type?: string }) => {
    const total = adv.agencies.reduce((s, a) => s + a.value, 0);
    const ykt = adv.agencies.find(a => a.name === "یکتانت");
    const yPct = total > 0 && ykt ? Math.round(ykt.value / total * 100) : 0;
    const accent = type === "lead" ? T.green : T.coral;
    const accentDim = type === "lead" ? T.greenDim : T.coralDim;
    const accentBorder = type === "lead" ? T.greenBorder : T.coralBorder;
    return (
      <div onClick={() => setModalAdv({ adv, type })} role="button" tabIndex={0} aria-label={`مشاهده جزئیات ${adv.name}`}
        onKeyDown={e => (e.key === "Enter" || e.key === " ") && setModalAdv({ adv, type })}
        style={{ ...card({ padding: "14px 16px", cursor: "pointer", display: "flex", alignItems: "center", gap: 16, transition: "border-color 0.15s" }), borderColor: T.border }}
        onMouseEnter={e => (e.currentTarget.style.borderColor = accentBorder)}
        onMouseLeave={e => (e.currentTarget.style.borderColor = T.border)}
        className="fade-up">
        <DonutChart agencies={adv.agencies} size={64} />
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 4 }}>
            <span style={{ fontSize: 14, fontWeight: 600, color: T.text1, direction: "ltr" }}>{adv.name}</span>
            <span style={{ fontSize: 11, color: T.text3 }}>#{adv.ownerid}</span>
            {adv.manager && <span style={{ fontSize: 11, color: accent, background: accentDim, border: `1px solid ${accentBorder}`, padding: "1px 8px", borderRadius: 20 }}>{adv.manager}</span>}
            {adv.industry1 && <span style={{ fontSize: 10, color: T.blue, background: T.blueDim, padding: "1px 7px", borderRadius: 20, border: `1px solid rgba(74,158,232,0.2)` }}>{adv.industry1}</span>}
            {adv.industry2 && <span style={{ fontSize: 10, color: T.text3, background: T.surface2, padding: "1px 7px", borderRadius: 20, border: `1px solid ${T.border}` }}>{adv.industry2}</span>}
          </div>
          {(adv.team || adv.performanceManager || adv.supervisor) && (
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 5 }}>
              {adv.team && <span style={{ fontSize: 10, color: T.text3, background: T.surface2, border: `1px solid ${T.border}`, padding: "1px 8px", borderRadius: 20 }}>تیم: {adv.team}</span>}
              {adv.performanceManager && <span style={{ fontSize: 10, color: T.text3, background: T.surface2, border: `1px solid ${T.border}`, padding: "1px 8px", borderRadius: 20 }}>PM: {adv.performanceManager}</span>}
              {adv.supervisor && <span style={{ fontSize: 10, color: T.text3, background: T.surface2, border: `1px solid ${T.border}`, padding: "1px 8px", borderRadius: 20 }}>سوپروایزر: {adv.supervisor}</span>}
            </div>
          )}
          <p style={{ margin: "0 0 8px", fontSize: 12, color: T.text2, lineHeight: 1.7, display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" }}>
            {type === "lead" ? adv.note : adv.summary}
          </p>
          <AgencyLegend agencies={adv.agencies} />
        </div>
        {yPct > 0 && (
          <div style={{ textAlign: "center", flexShrink: 0 }}>
            <div style={{ fontSize: 20, fontWeight: 700, color: yPct >= 50 ? T.coral : yPct >= 30 ? T.amber : T.danger }}>{yPct}٪</div>
            <div style={{ fontSize: 10, color: T.text3 }}>یکتانت</div>
          </div>
        )}
      </div>
    );
  };

  // ─ Loading screen ───────────────────────────────────────────────────────────
  const LoadingScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "60vh", gap: 28 }}>
      <RadarHero />
      <div style={{ width: "100%", maxWidth: 320 }}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
          <span style={{ fontSize: 13, color: T.text2 }}>در حال آنالیز داده‌ها...</span>
          <span style={{ fontSize: 13, color: T.coral, fontWeight: 600 }}>{Math.round(loadingProgress)}٪</span>
        </div>
        <div style={{ height: 4, background: T.border, borderRadius: 4, overflow: "hidden" }}>
          <div style={{ height: "100%", background: T.coral, borderRadius: 4, width: `${loadingProgress}%`, transition: "width 0.4s ease" }} />
        </div>
      </div>
      <p style={{ margin: 0, fontSize: 12, color: T.text3 }}>Claude در حال بررسی رقابت بازار است</p>
    </div>
  );

  // ─ Upload screen ────────────────────────────────────────────────────────────
  const UploadScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {step === "loading" ? <LoadingScreen /> : step === "upload" ? (
        <>
          <div style={{ textAlign: "center", paddingTop: 20 }}>
            <RadarHero />
            <h1 style={{ margin: "0 0 8px", fontSize: 22, fontWeight: 700, color: T.text1, letterSpacing: "-0.3px" }}>Radar</h1>
            <p style={{ margin: 0, fontSize: 14, color: T.text3 }}>آنالیز رقابتی روزانه یکتانت</p>
          </div>
          <div onDrop={onDrop} onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onClick={() => document.getElementById("fi")!.click()}
            role="button" tabIndex={0} aria-label="بارگذاری فایل CSV یا XLSX"
            onKeyDown={e => (e.key === "Enter" || e.key === " ") && document.getElementById("fi")!.click()}
            style={{ border: `2px dashed ${dragOver ? T.coral : T.border}`, borderRadius: 20, padding: "2.5rem 1.5rem", textAlign: "center", cursor: "pointer", background: dragOver ? T.coralDim : "transparent", transition: "all 0.2s" }}>
            <div style={{ width: 52, height: 52, borderRadius: 14, background: dragOver ? T.coral : T.surface2, border: `1px solid ${T.border}`, margin: "0 auto 14px", display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.2s" }}>
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={dragOver ? "#fff" : T.coral} strokeWidth="1.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>
            </div>
            <p style={{ margin: "0 0 5px", fontWeight: 600, fontSize: 15, color: T.text1 }}>فایل را بکشید یا کلیک کنید</p>
            <p style={{ margin: 0, fontSize: 12, color: T.text3 }}>CSV یا XLSX · داده‌های رقابتی یکتانت</p>
            <input id="fi" type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={e => handleFile(e.target.files![0])} />
          </div>
          {!hasApiKey && !showApiKeyInput && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, padding: "12px 16px", background: T.dangerBg, border: `1px solid ${T.dangerBorder}`, borderRadius: 12 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={T.danger} strokeWidth="2"><circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" /></svg>
                <p style={{ margin: 0, fontSize: 12, color: T.danger }}>کلید OpenRouter تنظیم نشده — بدون آن آنالیز اجرا نمی‌شود</p>
              </div>
              <button onClick={() => setShowApiKeyInput(true)} style={{ padding: "6px 14px", borderRadius: 8, border: `1px solid ${T.dangerBorder}`, background: "transparent", color: T.danger, fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>تنظیم</button>
            </div>
          )}
          {showApiKeyInput && (
            <div style={{ ...card({ padding: "16px 18px" }) }}>
              <p style={{ margin: "0 0 10px", fontSize: 12, fontWeight: 600, color: T.text3 }}>کلید OpenRouter API</p>
              <div style={{ display: "flex", gap: 10 }}>
                <input type="password" value={apiKeyInput} onChange={e => setApiKeyInput(e.target.value)} placeholder="sk-or-..." autoFocus
                  onKeyDown={e => e.key === "Enter" && saveApiKey()}
                  style={{ flex: 1, fontSize: 13, padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: T.surface2, color: T.text1, outline: "none", direction: "ltr" }} />
                <button onClick={saveApiKey} style={{ padding: "10px 18px", borderRadius: 10, border: "none", background: T.coral, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>ذخیره</button>
                <button onClick={() => setShowApiKeyInput(false)} style={{ padding: "10px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: "transparent", color: T.text2, fontSize: 13, cursor: "pointer" }}>لغو</button>
              </div>
              <p style={{ margin: "8px 0 0", fontSize: 11, color: T.text3 }}>کلید از <span style={{ direction: "ltr", display: "inline-block" }}>openrouter.ai/keys</span> دریافت می‌شود و فقط در مرورگر شما ذخیره می‌شود</p>
            </div>
          )}
          {uniqueReported.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", background: T.coralDim, border: `1px solid ${T.coralBorder}`, borderRadius: 12 }}>
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={T.coral} strokeWidth="2"><circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="12" /><line x1="12" y1="16" x2="12.01" y2="16" /></svg>
              <p style={{ margin: 0, fontSize: 12, color: T.coral }}>{uniqueReported.length} تبلیغ‌کننده از ۵ روز اخیر skip می‌شن</p>
            </div>
          )}
        </>
      ) : (
        // Ready step
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div style={card({ padding: "16px 20px" })}>
            <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 16 }}>
              <div style={{ width: 40, height: 40, borderRadius: 11, background: T.coralDim, border: `1px solid ${T.coralBorder}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke={T.coral} strokeWidth="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" /><polyline points="14 2 14 8 20 8" /></svg>
              </div>
              <div style={{ flex: 1 }}>
                <p style={{ margin: 0, fontWeight: 600, fontSize: 14, color: T.text1 }}>{csvData!.name}</p>
                <p style={{ margin: "2px 0 0", fontSize: 12, color: T.text3 }}>{csvData!.rows.length.toLocaleString()} ردیف · {csvData!.stats.dates.length} روز</p>
              </div>
              <button onClick={() => { setCsvData(null); setStep("upload"); }} style={btnStyle()}>تغییر</button>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>
              {[{ l: "تبلیغ‌کننده", v: csvData!.stats.advertisers }, { l: "روزها", v: csvData!.stats.dates.length }, { l: "آخرین تاریخ", v: csvData!.stats.lastDate }].map((c, i) => (
                <div key={i} style={{ background: T.surface2, borderRadius: 10, padding: "10px 12px" }}>
                  <p style={{ margin: "0 0 3px", fontSize: 11, color: T.text3 }}>{c.l}</p>
                  <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: T.text1, direction: "ltr", textAlign: "right" }}>{c.v}</p>
                </div>
              ))}
            </div>
          </div>

          <div style={card({ padding: "14px 18px" })}>
            <p style={{ margin: "0 0 12px", fontSize: 11, fontWeight: 600, color: T.text3, textTransform: "uppercase", letterSpacing: "0.08em" }}>قالب گزارش</p>
            <div style={{ display: "flex", gap: 8 }}>
              {Object.entries(TEMPLATES).map(([key, t]) => (
                <button key={key} onClick={() => setSelectedTemplate(key)} style={{ flex: 1, padding: "10px 8px", borderRadius: 11, border: `1.5px solid ${selectedTemplate === key ? T.coral : T.border}`, background: selectedTemplate === key ? T.coralDim : "transparent", cursor: "pointer", transition: "all 0.18s" }}>
                  <p style={{ margin: "0 0 3px", fontSize: 13, fontWeight: selectedTemplate === key ? 600 : 400, color: selectedTemplate === key ? T.coral : T.text1 }}>{t.label}</p>
                  <p style={{ margin: 0, fontSize: 11, color: T.text3 }}>{t.desc}</p>
                </button>
              ))}
            </div>
          </div>

          <button onClick={analyze} style={{ padding: "14px", borderRadius: 14, border: "none", background: T.coral, color: "#fff", fontSize: 15, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 10, transition: "opacity 0.2s" }}
            onMouseEnter={e => (e.currentTarget.style.opacity = "0.88")}
            onMouseLeave={e => (e.currentTarget.style.opacity = "1")}>
            شروع آنالیز
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><line x1="5" y1="12" x2="19" y2="12" /><polyline points="12 5 19 12 12 19" /></svg>
          </button>

          {errMsg && <div style={{ padding: "12px 16px", borderRadius: 12, background: T.dangerBg, border: `1px solid ${T.dangerBorder}` }}>
            <p style={{ margin: 0, fontSize: 13, color: T.danger }}>{errMsg}</p>
          </div>}
          {showDebug && rawDebug && <pre style={{ margin: 0, fontSize: 11, color: T.text1, background: T.surface2, padding: 12, borderRadius: 10, whiteSpace: "pre-wrap", direction: "ltr", textAlign: "left", maxHeight: 200, overflow: "auto" }}>{rawDebug}</pre>}
        </div>
      )}
    </div>
  );

  // ─ Results screen ───────────────────────────────────────────────────────────
  const ResultsScreen = () => {
    if (!result) return <div style={{ textAlign: "center", padding: "4rem", color: T.text3, fontSize: 14 }}>هنوز آنالیزی انجام نشده</div>;

    // Industry stats from result + refData
    type IndustryStat = { name: string; advertisers: Advertiser[]; leads: Advertiser[]; avgYkt: number; topCompetitor: string };
    const industryMap = new Map<string, { advs: Advertiser[]; leads: Advertiser[] }>();
    [...result.advertisers, ...result.leads].forEach(a => {
      const ind = a.industry1 || "سایر";
      if (!industryMap.has(ind)) industryMap.set(ind, { advs: [], leads: [] });
      if (result.advertisers.includes(a)) industryMap.get(ind)!.advs.push(a);
      else industryMap.get(ind)!.leads.push(a);
    });
    const industryStats: IndustryStat[] = [...industryMap.entries()].map(([name, { advs, leads }]) => {
      const yktPcts = advs.map(a => { const t = a.agencies.reduce((s, x) => s + x.value, 0); const y = a.agencies.find(x => x.name === "یکتانت"); return t > 0 && y ? y.value / t : 0; });
      const avgYkt = yktPcts.length > 0 ? Math.round(yktPcts.reduce((a, b) => a + b, 0) / yktPcts.length * 100) : 0;
      const compCount: Record<string, number> = {};
      advs.forEach(a => a.agencies.filter(x => x.name !== "یکتانت").forEach(x => { compCount[x.name] = (compCount[x.name] || 0) + x.value; }));
      const topCompetitor = Object.entries(compCount).sort((a, b) => b[1] - a[1])[0]?.[0] || "";
      return { name, advertisers: advs, leads, avgYkt, topCompetitor };
    }).sort((a, b) => (b.advertisers.length + b.leads.length) - (a.advertisers.length + a.leads.length));

    const hasIndustry = industryStats.some(s => s.name !== "سایر");
    const tabs: { id: "advertisers" | "leads" | "competitors" | "industry"; label: string; count: number }[] = [
      { id: "advertisers", label: "تبلیغ‌کننده‌ها", count: result.advertisers.length },
      { id: "leads", label: "لیدها", count: result.leads.length },
      { id: "competitors", label: "رقبا", count: result.competitors.length },
      ...(hasIndustry ? [{ id: "industry" as const, label: "صنعت", count: industryStats.length }] : []),
    ];
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        {/* Stats row */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14 }}>
          {[{ l: "تاریخ", v: result.dateLabel, c: T.text1 }, { l: "تبلیغ‌کننده", v: result.advertisers.length, c: T.coral }, { l: "لید", v: result.leads.length, c: T.green }, { l: "رقیب", v: result.competitors.length, c: T.amber }].map((s, i) => (
            <div key={i} style={card({ padding: "16px 20px" })}>
              <p style={{ margin: "0 0 6px", fontSize: 12, color: T.text3 }}>{s.l}</p>
              <p style={{ margin: 0, fontSize: i === 0 ? 14 : 28, fontWeight: 700, color: s.c }}>{s.v}</p>
            </div>
          ))}
        </div>

        {/* Recommended action banner */}
        {result.leads.length > 0 && (
          <div style={{ display: "flex", alignItems: "center", gap: 14, padding: "14px 18px", background: T.greenDim, border: `1px solid ${T.greenBorder}`, borderRadius: 14 }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: T.green, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5"><polyline points="9 11 12 14 22 4" /><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11" /></svg>
            </div>
            <div style={{ flex: 1 }}>
              <p style={{ margin: "0 0 2px", fontSize: 13, fontWeight: 700, color: T.text1 }}>اقدام پیشنهادی</p>
              <p style={{ margin: 0, fontSize: 12, color: T.text2 }}>
                {result.leads.length} لید شناسایی شده — اولویت: <button onClick={() => { setResultsTab("leads"); }} style={{ background: "none", border: "none", color: T.green, fontWeight: 700, fontSize: 12, cursor: "pointer", padding: 0, textDecoration: "underline" }}>{result.leads[0].name}</button>
              </p>
            </div>
            <button onClick={() => setResultsTab("leads")} style={{ padding: "7px 14px", borderRadius: 9, border: `1px solid ${T.greenBorder}`, background: "transparent", color: T.green, fontSize: 12, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>مشاهده لیدها</button>
          </div>
        )}

        {/* Save prompt */}
        {pendingSave && saveStatus !== "saved" && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px 16px", background: T.coralDim, border: `1px solid ${T.coralBorder}`, borderRadius: 14 }}>
            <div>
              <p style={{ margin: "0 0 2px", fontSize: 13, fontWeight: 600, color: T.text1 }}>این گزارش ذخیره نشده</p>
              <p style={{ margin: 0, fontSize: 12, color: T.text3 }}>{pendingSave.names.length} تبلیغ‌کننده</p>
            </div>
            <button onClick={confirmSave} style={{ padding: "8px 18px", borderRadius: 10, border: "none", background: T.coral, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>ذخیره</button>
          </div>
        )}

        {/* Tabs */}
        <div style={{ display: "flex", gap: 4, background: T.surface2, borderRadius: 12, padding: 4 }}>
          {tabs.map(tab => (
            <button key={tab.id} onClick={() => setResultsTab(tab.id)}
              style={{ flex: 1, padding: "9px 8px", borderRadius: 9, border: "none", cursor: "pointer", background: resultsTab === tab.id ? T.surface : "transparent", color: resultsTab === tab.id ? T.text1 : T.text3, fontSize: 13, fontWeight: resultsTab === tab.id ? 600 : 400, transition: "all 0.18s", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
              {tab.label}
              <span style={{ fontSize: 11, background: resultsTab === tab.id ? T.coralDim : "transparent", color: resultsTab === tab.id ? T.coral : T.text3, padding: "1px 7px", borderRadius: 20, fontWeight: 600 }}>{tab.count}</span>
            </button>
          ))}
        </div>

        {/* Tab content */}
        {resultsTab === "advertisers" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <input value={searchQuery} onChange={e => setSearchQuery(e.target.value)} placeholder="جستجو..." style={{ flex: 1, minWidth: 140, fontSize: 14, padding: "10px 14px", borderRadius: 11, border: `1px solid ${T.border}`, background: T.surface, color: T.text1, direction: "rtl", outline: "none" }} />
              <select value={filterAgency} onChange={e => setFilterAgency(e.target.value)} style={{ fontSize: 13, padding: "10px 12px", borderRadius: 11, border: `1px solid ${T.border}`, background: T.surface, color: T.text1, cursor: "pointer" }}>
                <option value="all">همه آژانس‌ها</option>
                {allAgencies.map(a => <option key={a} value={a}>{a}</option>)}
              </select>
              <select value={sortBy} onChange={e => setSortBy(e.target.value)} style={{ fontSize: 13, padding: "10px 12px", borderRadius: 11, border: `1px solid ${T.border}`, background: T.surface, color: T.text1, cursor: "pointer" }}>
                <option value="importance">اهمیت</option>
                <option value="volume">حجم</option>
                <option value="yektanet">سهم یکتانت</option>
              </select>
            </div>
            {filteredAdvertisers.length === 0
              ? <div style={{ textAlign: "center", padding: "3rem", color: T.text3, fontSize: 14, display: "flex", flexDirection: "column", alignItems: "center", gap: 12 }}>
                  <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke={T.text3} strokeWidth="1.5"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>
                  <span>نتیجه‌ای یافت نشد</span>
                  <button onClick={() => { setSearchQuery(""); setFilterAgency("all"); }} style={{ fontSize: 12, padding: "6px 16px", borderRadius: 8, border: `1px solid ${T.border}`, background: "transparent", color: T.text2, cursor: "pointer" }}>پاک کردن فیلتر</button>
                </div>
              : <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(420px, 1fr))", gap: 12 }}>
                  {filteredAdvertisers.map((adv, i) => <AdvCard key={i} adv={adv} type="advertiser" />)}
                </div>}
          </div>
        )}

        {resultsTab === "leads" && (
          <div>
            {result.leads.length === 0
              ? <div style={{ textAlign: "center", padding: "4rem", color: T.text3, fontSize: 14, display: "flex", flexDirection: "column", alignItems: "center", gap: 12 }}>
                  <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke={T.text3} strokeWidth="1.2"><circle cx="12" cy="12" r="10" /><path d="M8 12h8M12 8v8" /></svg>
                  <span>لیدی در این تحلیل یافت نشد</span>
                  <span style={{ fontSize: 12, maxWidth: 280, lineHeight: 1.7 }}>اگر داده کافی در فایل وجود دارد، می‌توانید با قالب تفصیلی مجدداً آنالیز کنید</span>
                </div>
              : <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(420px, 1fr))", gap: 12 }}>
                  {result.leads.map((l, i) => <AdvCard key={i} adv={l} type="lead" />)}
                </div>}
          </div>
        )}

        {resultsTab === "competitors" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))", gap: 12 }}>
            {result.competitors.length === 0
              ? <div style={{ textAlign: "center", padding: "3rem", color: T.text3, fontSize: 13 }}>رقیبی یافت نشد</div>
              : result.competitors.map((c, i) => (
                <div key={i} style={card({ padding: "16px 18px" })} className="fade-up">
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                    <div style={{ width: 36, height: 36, borderRadius: 10, background: T.amberDim, border: `1px solid ${T.amberBorder}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={T.amber} strokeWidth="1.8"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12" /></svg>
                    </div>
                    <span style={{ fontSize: 15, fontWeight: 700, color: T.amber }}>{c.platform}</span>
                  </div>
                  <p style={{ margin: "0 0 8px", fontSize: 13, color: T.text2, lineHeight: 1.8 }}>{c.note}</p>
                  {c.topclients && <p style={{ margin: "4px 0 0", fontSize: 12, color: T.text3 }}>مهم‌ترین: <span style={{ color: T.text2 }}>{c.topclients}</span></p>}
                  {c.newclients && c.newclients !== "ندارد" && <p style={{ margin: "4px 0 0", fontSize: 12, color: T.green }}>جدید: {c.newclients}</p>}
                </div>
              ))}
            {result.market && (
              <div style={{ ...card({ padding: "16px 18px" }), gridColumn: "1 / -1" }}>
                <p style={{ margin: "0 0 8px", fontSize: 12, fontWeight: 600, color: T.text3, textTransform: "uppercase", letterSpacing: "0.08em" }}>تحلیل کلی بازار</p>
                <p style={{ margin: 0, fontSize: 14, color: T.text2, lineHeight: 1.9 }}>{result.market}</p>
              </div>
            )}
          </div>
        )}

        {resultsTab === "industry" && (
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            {industryStats.length === 0
              ? <div style={{ textAlign: "center", padding: "4rem", color: T.text3, fontSize: 14 }}>اطلاعات صنعتی موجود نیست — ابتدا اطلاعات مرجع بارگذاری شود</div>
              : industryStats.map((s, i) => {
                const total = s.advertisers.length + s.leads.length;
                const yktColor = s.avgYkt >= 50 ? T.coral : s.avgYkt >= 30 ? T.amber : T.danger;
                return (
                  <div key={i} style={card({ padding: "16px 20px" })} className="fade-up">
                    <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, marginBottom: 12 }}>
                      <div style={{ flex: 1 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                          <span style={{ fontSize: 15, fontWeight: 700, color: T.text1 }}>{s.name}</span>
                          <span style={{ fontSize: 11, color: T.text3, background: T.surface2, padding: "2px 9px", borderRadius: 20 }}>{total} اکانت</span>
                          {s.leads.length > 0 && <span style={{ fontSize: 11, color: T.green, background: T.greenDim, padding: "2px 9px", borderRadius: 20 }}>{s.leads.length} لید</span>}
                          {s.topCompetitor && <span style={{ fontSize: 11, color: T.amber, background: T.amberDim, padding: "2px 9px", borderRadius: 20 }}>رقیب اصلی: {s.topCompetitor}</span>}
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10 }}>
                          <div style={{ flex: 1, height: 6, borderRadius: 3, background: T.border, overflow: "hidden" }}>
                            <div style={{ width: `${s.avgYkt}%`, height: "100%", background: yktColor, borderRadius: 3, transition: "width 0.5s ease" }} />
                          </div>
                          <span style={{ fontSize: 13, fontWeight: 700, color: yktColor, flexShrink: 0 }}>یکتانت {s.avgYkt}٪</span>
                        </div>
                      </div>
                      <button onClick={() => { setResultsTab("advertisers"); setFilterAgency("all"); setSearchQuery(s.name); }}
                        style={btnStyle({ fontSize: 11, padding: "6px 12px", flexShrink: 0 })}>
                        مشاهده اکانت‌ها
                      </button>
                    </div>
                    {s.advertisers.slice(0, 4).length > 0 && (
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                        {s.advertisers.slice(0, 4).map((a, j) => (
                          <button key={j} onClick={() => setModalAdv({ adv: a, type: "advertiser" })}
                            style={{ fontSize: 11, color: T.text2, background: T.surface2, border: `1px solid ${T.border}`, padding: "3px 10px", borderRadius: 20, cursor: "pointer", direction: "ltr" }}>
                            {a.name}
                          </button>
                        ))}
                        {s.advertisers.length > 4 && <span style={{ fontSize: 11, color: T.text3, padding: "3px 6px" }}>+{s.advertisers.length - 4} دیگر</span>}
                      </div>
                    )}
                  </div>
                );
              })}
          </div>
        )}

        {/* Export + message */}
        <div style={card({ padding: "16px 18px" })}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
            <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: T.text1 }}>پیام نهایی</p>
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => { navigator.clipboard.writeText(editableMsg); showToast("کپی شد ✓"); }}
                style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 16px", borderRadius: 10, border: "none", background: T.coral, color: "#fff", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="9" y="9" width="13" height="13" rx="2" /><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1" /></svg>
                کپی
              </button>
              <button onClick={() => downloadHTML(result!)} style={btnStyle({ display: "flex", alignItems: "center", gap: 6 })}>
                <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>
                HTML
              </button>
              <button onClick={reset} style={btnStyle()}>شروع مجدد</button>
            </div>
          </div>
          <textarea value={editableMsg} onChange={e => setEditableMsg(e.target.value)} rows={10}
            style={{ width: "100%", fontSize: 12, padding: "12px 14px", borderRadius: 10, border: `1px solid ${T.border}`, background: T.surface2, color: T.text1, resize: "vertical", boxSizing: "border-box", direction: "rtl", lineHeight: 2, fontFamily: "'JetBrains Mono', monospace", outline: "none" }} />
          {showDebug && rawDebug && <pre style={{ marginTop: 10, fontSize: 11, color: T.text1, background: T.surface2, padding: 12, borderRadius: 10, whiteSpace: "pre-wrap", direction: "ltr", textAlign: "left", maxHeight: 180, overflow: "auto" }}>{rawDebug}</pre>}
          <button onClick={() => setShowDebug(v => !v)} style={{ ...btnStyle({ marginTop: 8, fontSize: 11, padding: "5px 12px" }), color: T.text3 }}>دیباگ</button>
        </div>
      </div>
    );
  };

  // ─ Reports screen ───────────────────────────────────────────────────────────
  const ReportsScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <p style={{ margin: "0 0 4px", fontSize: 17, fontWeight: 700, color: T.text1 }}>گزارش‌های ذخیره‌شده</p>
      {savedReports.length === 0
        ? <div style={{ textAlign: "center", padding: "4rem", color: T.text3, fontSize: 13 }}>هنوز گزارشی ذخیره نشده</div>
        : [...savedReports].reverse().map((r, i) => (
          <div key={i} style={card({ padding: "16px 20px" })} className="fade-up">
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
              <div>
                <p style={{ margin: "0 0 2px", fontSize: 14, fontWeight: 600, color: T.text1 }}>{r.dateLabel}</p>
                <p style={{ margin: 0, fontSize: 12, color: T.text3 }}>{r.advertisers?.length || 0} تبلیغ‌کننده · {r.leads?.length || 0} لید</p>
              </div>
              <button onClick={() => { setResult(r); setEditableMsg(buildMessage(r)); setScreen("results"); setResultsTab("advertisers"); }}
                style={{ padding: "8px 16px", borderRadius: 10, border: `1px solid ${T.coralBorder}`, background: T.coralDim, color: T.coral, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>مشاهده</button>
            </div>
            {r.advertisers?.slice(0, 3).map((a, j) => {
              const total = a.agencies?.reduce((s, x) => s + x.value, 0) || 1;
              const ykt = a.agencies?.find(x => x.name === "یکتانت");
              return (
                <div key={j} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                  <span style={{ fontSize: 12, color: T.text2, direction: "ltr", minWidth: 100, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</span>
                  <div style={{ flex: 1, height: 4, borderRadius: 2, background: T.border, overflow: "hidden" }}>
                    <div style={{ width: `${Math.round((ykt?.value || 0) / total * 100)}%`, height: "100%", background: T.coral, borderRadius: 2 }} />
                  </div>
                  <span style={{ fontSize: 11, color: T.text3, flexShrink: 0, width: 30, textAlign: "left" }}>{Math.round((ykt?.value || 0) / total * 100)}٪</span>
                </div>
              );
            })}
          </div>
        ))}
    </div>
  );

  // ─ History screen ────────────────────────────────────────────────────────────
  const HistoryScreen = () => (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <p style={{ margin: 0, fontSize: 17, fontWeight: 700, color: T.text1 }}>تبلیغ‌کننده‌های گزارش‌شده</p>
        {history.length > 0 && <button onClick={clearHistory} style={btnStyle({ color: T.danger, borderColor: T.dangerBorder })}>پاک کردن همه</button>}
      </div>
      {history.length === 0
        ? <div style={{ textAlign: "center", padding: "4rem", color: T.text3, fontSize: 13 }}>هنوز گزارشی ذخیره نشده</div>
        : [...history].reverse().map((h, i) => (
          <div key={i} style={card({ padding: "16px 20px" })} className="fade-up">
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ width: 8, height: 8, borderRadius: "50%", background: T.coral, flexShrink: 0 }} />
                <span style={{ fontSize: 13, fontWeight: 600, color: T.text1 }}>{h.dateLabel || h.date}</span>
                <span style={{ fontSize: 11, color: T.text3, background: T.surface2, padding: "2px 9px", borderRadius: 20 }}>{h.names.length} تبلیغ‌کننده</span>
              </div>
              <button onClick={() => deleteHistoryEntry(h.date)} style={btnStyle({ fontSize: 11, padding: "5px 11px", color: T.danger, borderColor: T.dangerBorder })}>حذف</button>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {h.names.map((n, j) => <span key={j} style={{ fontSize: 11, color: T.coral, background: T.coralDim, border: `1px solid ${T.coralBorder}`, padding: "3px 10px", borderRadius: 20, direction: "ltr" }}>{n}</span>)}
            </div>
          </div>
        ))}
    </div>
  );

  // ─ Dashboard screen ─────────────────────────────────────────────────────────
  const DashboardScreen = () => {
    const D = useD();
    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 15, fontWeight: 600, color: D.t1 }}>داشبورد آنالیز رقابتی</p>
        <p style={{ margin: 0, fontSize: 13, color: D.t3 }}>ابتدا یک فایل XLSX آپلود کنید تا داشبورد فعال شود</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>رفتن به آنالیز</button>
      </div>
    );

    const rows = csvData.rows;
    const COMP_COLS = ["Tapsell","Deema","Tavoos","Adexo","Chavosh","Aparat","Daart","Yellowadwise","Najva","Triboon","Jaryan","Telewebion","Adverge","Soroush","Soroush_ny","Bale_ny","Rubika_ny","Eitaa_ny","Bazaar","Myket"];
    const TR_COMP: Record<string, string> = { Tapsell:"تپسل",Deema:"دیما",Tavoos:"طاووس",Adexo:"ادکسو",Chavosh:"چاووش",Aparat:"آپارات",Daart:"دارت",Yellowadwise:"یلو ادوایز",Najva:"نجوا",Triboon:"تریبون",Jaryan:"جریان",Telewebion:"تلوبیون",Adverge:"ادورج",Soroush:"سروش",Soroush_ny:"سروش",Bale_ny:"بله",Rubika_ny:"روبیکا",Eitaa_ny:"ایتا",Bazaar:"بازار",Myket:"مایکت" };

    // Unique advertisers by Owner_id
    const advSet = new Map<string, { name: string; industry: string; team: string; am: string; pm: string; sup: string; sessions: number; ykt: number; spend: number; comps: Record<string,number> }>();
    rows.forEach(r => {
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      if (!id) return;
      const cur = advSet.get(id) || { name: r.Advertiser_name || id, industry: r.Category_level_1 || "", team: r.Team || "", am: r.Account_manager_name || "", pm: r.Performance_manager_name || "", sup: r.Supervisor_name || "", sessions: 0, ykt: 0, spend: 0, comps: {} };
      cur.sessions += Number(r.Total_sessions) || 0;
      cur.ykt += Number(r.Yektanet) || 0;
      cur.spend += Number(r.Daily_spend) || 0;
      COMP_COLS.forEach(c => { cur.comps[c] = (cur.comps[c] || 0) + (Number(r[c]) || 0); });
      advSet.set(id, cur);
    });
    const advList = [...advSet.values()];

    // KPIs
    const totalSessions = advList.reduce((s, a) => s + a.sessions, 0);
    const totalYkt = advList.reduce((s, a) => s + a.ykt, 0);
    const yktShare = totalSessions > 0 ? Math.round(totalYkt / totalSessions * 100) : 0;
    const totalSpend = advList.reduce((s, a) => s + a.spend, 0);
    const dates = [...new Set(rows.map(r => r.Date || r.date).filter(Boolean))].sort();

    // Daily trend (last 7 days)
    const dailyMap = new Map<string, { total: number; ykt: number }>();
    rows.forEach(r => {
      const d = r.Date || r.date || "";
      if (!d) return;
      const cur = dailyMap.get(d) || { total: 0, ykt: 0 };
      cur.total += Number(r.Total_sessions) || 0;
      cur.ykt += Number(r.Yektanet) || 0;
      dailyMap.set(d, cur);
    });
    const dailyTrend = [...dailyMap.entries()].sort((a, b) => a[0].localeCompare(b[0])).slice(-7).map(([date, { total, ykt }]) => ({
      date, total, share: total > 0 ? Math.round(ykt / total * 100) : 0,
    }));

    // Competitor totals
    const compTotals = COMP_COLS.map(col => ({
      col, name: TR_COMP[col] || col,
      total: advList.reduce((s, a) => s + (a.comps[col] || 0), 0),
    })).filter(c => c.total > 0).sort((a, b) => b.total - a.total).slice(0, 8);
    const grandTotal = totalYkt + compTotals.reduce((s, c) => s + c.total, 0);

    // At-risk: ykt share < 35%, sessions > 5000
    const atRisk = advList.filter(a => a.sessions > 5000)
      .map(a => ({ ...a, yktShare: a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0 }))
      .filter(a => a.yktShare > 0 && a.yktShare < 35)
      .sort((a, b) => b.sessions - a.sessions).slice(0, 4);

    // Top advertisers
    const top5 = advList.filter(a => a.ykt > 0)
      .map(a => ({ ...a, yktShare: a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0 }))
      .sort((a, b) => b.sessions - a.sessions).slice(0, 5);

    // Leads: zero ykt, significant sessions
    const leads = advList.filter(a => a.ykt === 0 && a.sessions > 2000)
      .sort((a, b) => b.sessions - a.sessions);

    const yktColor = yktShare >= 50 ? D.accent : yktShare >= 30 ? D.amber : D.red;

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
        {/* KPI row */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12 }}>
          <KpiCard label="کل سشن‌ها" value={formatNumber(totalSessions)} sub={`${advList.length} آگهی‌دهنده`} color={D.accent}
            icon={<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>} />
          <KpiCard label="سهم یکتانت" value={`${yktShare}٪`} sub={`از ${formatNumber(totalYkt)} سشن`} color={yktColor} delay={40}
            icon={<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="4"/></svg>} />
          <KpiCard label="تبلیغ‌کننده فعال" value={advList.length} sub={`${dates.length} روز داده`} color={D.blue} delay={80}
            icon={<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 00-2-2h-4a2 2 0 00-2 2v2"/></svg>} />
          <KpiCard label="لیدهای فعال" value={leads.length} sub="بدون یکتانت" color={D.green} delay={120}
            icon={<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>} />
          <KpiCard label="درآمد تخمینی" value={fmtMoney(totalSpend)} sub="تومان" color={D.amber} delay={160}
            icon={<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>} />
        </div>

        {/* Main 2-col: trend + agency distribution */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 300px", gap: 14 }}>
          <div className="fu" style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "18px 20px" }}>
            <SectionHeader title="روند بازار" sub={`${dailyTrend.length} روز اخیر`} />
            {dailyTrend.length > 1 ? (
              <>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 12 }}>
                  <div>
                    <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 4 }}>مجموع سشن</div>
                    <MiniLineChart data={dailyTrend.map(d => d.total)} color={D.blue} h={80} />
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 4 }}>سهم یکتانت ٪</div>
                    <MiniLineChart data={dailyTrend.map(d => d.share)} color={D.accent} h={80} />
                  </div>
                </div>
                <div style={{ display: "flex", gap: 12, justifyContent: "center", overflowX: "auto" }}>
                  {dailyTrend.map((d, i) => (
                    <div key={i} style={{ textAlign: "center", flexShrink: 0 }}>
                      <div style={{ fontSize: 9, color: D.t3, fontFamily: D.mono }}>{d.date.slice(5)}</div>
                      <div style={{ fontSize: 10, fontWeight: 600, color: D.accent, fontFamily: D.mono }}>{d.share}٪</div>
                      <div style={{ fontSize: 9, color: D.t3, fontFamily: D.mono }}>{formatNumber(d.total)}</div>
                    </div>
                  ))}
                </div>
              </>
            ) : (
              <div style={{ color: D.t3, fontSize: 12, textAlign: "center", padding: "24px 0" }}>داده روزانه کافی نیست</div>
            )}
          </div>

          <div className="fu" style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "18px 20px", animationDelay: "60ms" }}>
            <SectionHeader title="توزیع آژانس‌ها" />
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 9 }}>
              <div style={{ width: 6, height: 6, borderRadius: "50%", background: D.accent, flexShrink: 0 }} />
              <div style={{ fontSize: 12, color: D.t2, width: 80, flexShrink: 0, textAlign: "right" }}>یکتانت</div>
              <div style={{ flex: 1, height: 4, borderRadius: 2, background: D.border2, overflow: "hidden" }}>
                <div style={{ width: `${grandTotal > 0 ? Math.round(totalYkt / grandTotal * 100) : 0}%`, height: "100%", background: D.accent, borderRadius: 2 }} />
              </div>
              <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t2, width: 30, textAlign: "left" }}>{grandTotal > 0 ? Math.round(totalYkt / grandTotal * 100) : 0}٪</div>
            </div>
            {compTotals.map((c, i) => (
              <div key={c.col} style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 9 }}>
                <div style={{ width: 6, height: 6, borderRadius: "50%", background: AGENCY_COLORS[c.name] || "#888", flexShrink: 0 }} />
                <div style={{ fontSize: 12, color: D.t2, width: 80, flexShrink: 0, textAlign: "right" }}>{c.name}</div>
                <div style={{ flex: 1, height: 4, borderRadius: 2, background: D.border2, overflow: "hidden" }}>
                  <div style={{ width: `${grandTotal > 0 ? Math.round(c.total / grandTotal * 100) : 0}%`, height: "100%", background: AGENCY_COLORS[c.name] || "#888", borderRadius: 2, animation: "barGrow .6s ease both", animationDelay: `${i * 60}ms` }} />
                </div>
                <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t2, width: 30, textAlign: "left" }}>{grandTotal > 0 ? Math.round(c.total / grandTotal * 100) : 0}٪</div>
              </div>
            ))}
          </div>
        </div>

        {/* Bottom 2-col: top advertisers + at-risk */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <div className="fu" style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "18px 20px" }}>
            <SectionHeader title="برترین تبلیغ‌کننده‌ها" sub="بر اساس حجم سشن"
              action={<button onClick={() => setScreen("results")} style={{ fontSize: 11, color: D.accent, background: D.accentDim, border: `1px solid ${D.accentBrd}`, borderRadius: 99, padding: "3px 10px", cursor: "pointer", fontFamily: "Vazirmatn,sans-serif" }}>همه ←</button>} />
            {top5.map((a, i) => (
              <div key={a.name} className="fu" style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: i < 4 ? `1px solid ${D.border}` : "none", animationDelay: `${i * 40}ms` }}>
                <div style={{ width: 22, height: 22, borderRadius: 6, background: D.accentDim, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 800, color: D.accent, fontFamily: D.mono, flexShrink: 0 }}>{i + 1}</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, fontWeight: 600, color: D.t1, fontFamily: D.mono, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 3 }}>
                    <div style={{ flex: 1, height: 3, borderRadius: 2, background: D.border, overflow: "hidden" }}>
                      <div style={{ width: `${a.yktShare}%`, height: "100%", background: D.accent, borderRadius: 2 }} />
                    </div>
                    <span style={{ fontSize: 10, color: D.accent, fontFamily: D.mono, flexShrink: 0 }}>{a.yktShare}٪</span>
                  </div>
                </div>
                <div style={{ textAlign: "left", flexShrink: 0 }}>
                  <div style={{ fontSize: 11, fontWeight: 600, color: D.t2, fontFamily: D.mono }}>{formatNumber(a.sessions)}</div>
                </div>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {atRisk.length > 0 && (
              <div className="fu" style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "18px 20px" }}>
                <SectionHeader title="در حال از دست دادن سهم"
                  action={<button onClick={() => setScreen("alerts")} style={{ fontSize: 11, color: D.red, background: D.redDim, border: `1px solid ${D.redBrd}`, borderRadius: 99, padding: "3px 10px", cursor: "pointer", fontFamily: "Vazirmatn,sans-serif" }}>
                    {atRisk.length} هشدار
                  </button>} />
                {atRisk.map((a, i) => (
                  <div key={a.name} className="fu" style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8, animationDelay: `${i * 40}ms` }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 12, fontFamily: D.mono, color: D.t1, marginBottom: 3, overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</div>
                      <Sparkline data={[...Array(7)].map((_, j) => a.yktShare - (6 - j) * 2 + Math.random() * 4)} color={D.red} w={100} h={22} />
                    </div>
                    <div style={{ fontSize: 18, fontWeight: 800, color: D.red, fontFamily: D.mono, letterSpacing: "-1px" }}>{a.yktShare}٪</div>
                  </div>
                ))}
              </div>
            )}

            {leads.length > 0 && (
              <div className="fu" style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "18px 20px", flex: 1 }}>
                <SectionHeader title="فرصت‌های لید"
                  action={<button onClick={() => setScreen("leads")} style={{ fontSize: 11, color: D.green, background: D.greenDim, border: `1px solid ${D.greenBrd}`, borderRadius: 99, padding: "3px 10px", cursor: "pointer", fontFamily: "Vazirmatn,sans-serif" }}>
                    {leads.length} لید ←
                  </button>} />
                {leads.slice(0, 4).map((a, i) => (
                  <div key={a.name} className="fu" style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: i < 3 ? `1px solid ${D.border}` : "none", animationDelay: `${i * 40}ms` }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 12, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</div>
                      <div style={{ fontSize: 10, color: D.t3, marginTop: 2 }}>{a.industry}</div>
                    </div>
                    <span style={{ fontSize: 11, color: D.t2, fontFamily: D.mono }}>{formatNumber(a.sessions)}</span>
                    <span style={{ fontSize: 11, color: D.green, fontWeight: 700, background: D.greenDim, border: `1px solid ${D.greenBrd}`, borderRadius: 6, padding: "2px 7px" }}>لید</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      </div>
    );
  };

  // ─ Settings screen ───────────────────────────────────────────────────────────
  const SettingsScreen = () => {
    const [localKey, setLocalKey] = useState(apiKeyInput);
    const handleSave = () => {
      setApiKeyInput(localKey);
      const trimmed = localKey.trim();
      if (trimmed) { localStorage.setItem("radar:openrouter_key", trimmed); showToast("کلید API ذخیره شد ✓"); }
      else { localStorage.removeItem("radar:openrouter_key"); showToast("کلید API حذف شد"); }
    };
    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        <p style={{ margin: "0 0 4px", fontSize: 17, fontWeight: 700, color: T.text1 }}>تنظیمات</p>

        <div style={card({ padding: "20px 22px" })}>
          <p style={{ margin: "0 0 14px", fontSize: 13, fontWeight: 600, color: T.text2 }}>کلید OpenRouter API</p>
          <div style={{ display: "flex", gap: 10 }}>
            <input type="password" value={localKey} onChange={e => setLocalKey(e.target.value)}
              placeholder="sk-or-..." onKeyDown={e => e.key === "Enter" && handleSave()}
              aria-label="کلید OpenRouter API"
              style={{ flex: 1, fontSize: 13, padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.border}`, background: T.surface2, color: T.text1, outline: "none", direction: "ltr" }} />
            <button onClick={handleSave} style={{ padding: "10px 20px", borderRadius: 10, border: "none", background: T.coral, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>ذخیره</button>
          </div>
          <p style={{ margin: "8px 0 0", fontSize: 11, color: T.text3, lineHeight: 1.6 }}>
            کلید را از سایت openrouter.ai/keys دریافت کنید. این کلید فقط در مرورگر شما ذخیره می‌شود و به هیچ سروری ارسال نمی‌شود.
          </p>
        </div>

        <div style={card({ padding: "18px 22px" })}>
          <p style={{ margin: "0 0 10px", fontSize: 13, fontWeight: 600, color: T.text2 }}>ساختار فایل ورودی</p>
          <p style={{ margin: "0 0 8px", fontSize: 12, color: T.text3, lineHeight: 1.7 }}>ستون‌های مورد انتظار:</p>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {["Date","Owner_id","Advertiser_name","Category_level_1","Category_level_2","Team","Account_manager_name","Daily_spend","Total_sessions"].map(col => (
              <span key={col} style={{ fontSize: 11, color: T.text2, background: T.surface2, border: `1px solid ${T.border}`, padding: "3px 10px", borderRadius: 20, direction: "ltr" }}>{col}</span>
            ))}
          </div>
          <p style={{ margin: "10px 0 0", fontSize: 11, color: T.text3 }}>صنعت و تیم مستقیم از ستون‌های فایل خوانده می‌شود — نیازی به فایل مرجع جداگانه نیست.</p>
        </div>

        <div style={card({ padding: "18px 22px" })}>
          <p style={{ margin: "0 0 10px", fontSize: 13, fontWeight: 600, color: T.text2 }}>اطلاعات کاربر</p>
          <div style={{ display: "flex", gap: 16, fontSize: 13, color: T.text2 }}>
            <span>نام: <strong style={{ color: T.text1 }}>{session!.managerFa}</strong></span>
            {!session!.isAdmin && <span>تیم: <strong style={{ color: T.text1 }}>{session!.teamFa}</strong></span>}
          </div>
        </div>
      </div>
    );
  };

  // ─ Competitor screen ─────────────────────────────────────────────────────────
  const CompetitorScreen = () => {
    const D = useD();
    const [selected, setSelected] = useState<number>(0);

    const COMP_COLS = ["Tapsell","Deema","Tavoos","Adexo","Chavosh","Aparat","Daart","Yellowadwise","Najva","Triboon","Jaryan","Telewebion","Adverge","Soroush","Soroush_ny","Bale_ny","Rubika_ny","Eitaa_ny","Bazaar","Myket"];
    const TR_COMP: Record<string, string> = { Tapsell:"تپسل",Deema:"دیما",Tavoos:"طاووس",Adexo:"ادکسو",Chavosh:"چاووش",Aparat:"آپارات",Daart:"دارت",Yellowadwise:"یلو ادوایز",Najva:"نجوا",Triboon:"تریبون",Jaryan:"جریان",Telewebion:"تلوبیون",Adverge:"ادورج",Soroush:"سروش",Soroush_ny:"سروش",Bale_ny:"بله",Rubika_ny:"روبیکا",Eitaa_ny:"ایتا",Bazaar:"بازار",Myket:"مایکت" };

    if (!csvData && !result) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>برای مشاهده رقبا ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    // Compute competitor stats from csvData
    const compStats = csvData ? COMP_COLS.map(col => {
      const total = csvData.rows.reduce((s, r) => s + (Number(r[col]) || 0), 0);
      const clients = new Set(csvData.rows.filter(r => Number(r[col]) > 0).map(r => r.Advertiser_name)).size;
      return { col, name: TR_COMP[col] || col, total, clients, color: AGENCY_COLORS[TR_COMP[col] || col] || "#888" };
    }).filter(c => c.total > 0).sort((a, b) => b.total - a.total) : [];

    const grandComp = compStats.reduce((s, c) => s + c.total, 0);
    const sel = compStats[selected] || compStats[0];

    // Top advertisers using this competitor
    const topClients = csvData && sel ? csvData.rows
      .reduce((acc: { name: string; val: number }[], r) => {
        const val = Number(r[sel.col]) || 0;
        if (!val) return acc;
        const existing = acc.find(x => x.name === r.Advertiser_name);
        if (existing) existing.val += val; else acc.push({ name: r.Advertiser_name, val });
        return acc;
      }, []).sort((a, b) => b.val - a.val).slice(0, 6) : [];

    // AI competitor notes from result
    const aiComps = result?.competitors || [];

    if (compStats.length === 0) return (
      <div style={{ textAlign: "center", padding: "4rem", color: D.t3, fontSize: 13 }}>داده رقیبی در فایل یافت نشد</div>
    );

    return (
      <div style={{ display: "flex", gap: 16, minHeight: 500 }}>
        {/* Competitor list sidebar */}
        <div style={{ width: 210, flexShrink: 0, background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "14px 10px", overflowY: "auto" }}>
          <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 600, marginBottom: 12, padding: "0 4px" }}>رقبا</div>
          {compStats.map((c, i) => (
            <button key={c.col} onClick={() => setSelected(i)} style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${selected === i ? c.color + "44" : D.border}`, background: selected === i ? c.color + "11" : "transparent", cursor: "pointer", marginBottom: 6, transition: "all .14s", textAlign: "right", fontFamily: "Vazirmatn,sans-serif" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                <div style={{ width: 8, height: 8, borderRadius: "50%", background: c.color, flexShrink: 0 }} />
                <span style={{ fontSize: 13, fontWeight: 600, color: selected === i ? c.color : D.t1 }}>{c.name}</span>
                <span style={{ fontSize: 10, color: D.t3, fontFamily: D.mono, marginRight: "auto" }}>{grandComp > 0 ? Math.round(c.total / grandComp * 100) : 0}٪</span>
              </div>
              <div style={{ fontSize: 10, color: D.t3, fontFamily: D.mono }}>{c.clients} مشتری</div>
            </button>
          ))}
        </div>

        {/* Detail panel */}
        {sel && (
          <div className="fu" style={{ flex: 1, display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <div style={{ width: 44, height: 44, borderRadius: 14, background: sel.color + "22", border: `1.5px solid ${sel.color + "44"}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
                <div style={{ width: 16, height: 16, borderRadius: "50%", background: sel.color }} />
              </div>
              <div>
                <div style={{ fontSize: 22, fontWeight: 800, color: D.t1, letterSpacing: "-.4px" }}>{sel.name}</div>
                <div style={{ fontSize: 12, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>
                  {sel.clients} مشتری · سهم {grandComp > 0 ? Math.round(sel.total / grandComp * 100) : 0}٪ از رقبا
                </div>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
              <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 12, padding: "14px 16px" }}>
                <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>سهم از رقبا</div>
                <div style={{ fontSize: 28, fontWeight: 800, color: sel.color, fontFamily: D.mono, lineHeight: 1 }}>{grandComp > 0 ? Math.round(sel.total / grandComp * 100) : 0}٪</div>
              </div>
              <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 12, padding: "14px 16px" }}>
                <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>کل مشتریان</div>
                <div style={{ fontSize: 28, fontWeight: 800, color: D.t1, fontFamily: D.mono, lineHeight: 1 }}>{sel.clients}</div>
              </div>
              <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 12, padding: "14px 16px" }}>
                <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".08em", marginBottom: 6 }}>حجم سشن</div>
                <div style={{ fontSize: 28, fontWeight: 800, color: D.blue, fontFamily: D.mono, lineHeight: 1 }}>{formatNumber(sel.total)}</div>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 18px" }}>
                <SectionHeader title="مهم‌ترین مشتریان" />
                {topClients.map((c, i) => (
                  <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: i < topClients.length - 1 ? `1px solid ${D.border}` : "none" }}>
                    <span style={{ fontSize: 10, color: D.t3, fontFamily: D.mono, width: 14 }}>{i + 1}</span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 12, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis" }}>{c.name}</div>
                      <div style={{ height: 3, borderRadius: 2, background: D.border, marginTop: 4, overflow: "hidden" }}>
                        <div style={{ width: `${topClients[0] ? Math.round(c.val / topClients[0].val * 100) : 0}%`, height: "100%", background: sel.color, borderRadius: 2 }} />
                      </div>
                    </div>
                    <span style={{ fontSize: 10, fontFamily: D.mono, color: sel.color }}>{formatNumber(c.val)}</span>
                  </div>
                ))}
              </div>

              <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 18px" }}>
                <SectionHeader title="یادداشت AI" />
                {aiComps.filter(c => c.platform.includes(sel.name) || sel.name.includes(c.platform)).slice(0, 1).map((c, i) => (
                  <div key={i}>
                    <p style={{ margin: "0 0 8px", fontSize: 12, color: D.t2, lineHeight: 1.9 }}>{c.note}</p>
                    {c.topclients && <p style={{ margin: "4px 0 0", fontSize: 11, color: D.t3 }}>مهم‌ترین: <span style={{ color: D.t2 }}>{c.topclients}</span></p>}
                    {c.newclients && c.newclients !== "ندارد" && <p style={{ margin: "4px 0 0", fontSize: 11, color: D.green }}>جدید: {c.newclients}</p>}
                  </div>
                ))}
                {aiComps.filter(c => c.platform.includes(sel.name) || sel.name.includes(c.platform)).length === 0 && (
                  <p style={{ fontSize: 12, color: D.t3, marginTop: 8 }}>آنالیز AI موجود نیست — ابتدا آنالیز را اجرا کنید</p>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  // ─ Market map screen ─────────────────────────────────────────────────────────
  const MarketMapScreen = () => {
    const D = useD();
    const [hovered, setHovered] = useState<string | null>(null);
    const [filterCat, setFilterCat] = useState("all");

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>برای مشاهده نقشه بازار ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    const rows = csvData.rows;
    const advSet = new Map<string, { name: string; cat1: string; sessions: number; ykt: number }>();
    rows.forEach(r => {
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      if (!id) return;
      const cur = advSet.get(id) || { name: r.Advertiser_name || id, cat1: r.Category_level_1 || "", sessions: 0, ykt: 0 };
      cur.sessions += Number(r.Total_sessions) || 0;
      cur.ykt += Number(r.Yektanet) || 0;
      advSet.set(id, cur);
    });
    const advList = [...advSet.values()].map(a => ({ ...a, yktShare: a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0 })).filter(a => a.sessions > 0);
    const cats = [...new Set(advList.map(a => a.cat1).filter(Boolean))];
    const items = filterCat === "all" ? advList : advList.filter(a => a.cat1 === filterCat);

    const maxS = Math.max(...items.map(a => a.sessions), 1);
    const W = 900, H = 480, pad = 60;
    const hovItem = items.find(a => a.name === hovered);

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>نقشه بازار</div>
            <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>محور X: حجم سشن · محور Y: سهم یکتانت · اندازه: حجم نسبی</div>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            <Pill label="همه" active={filterCat === "all"} onClick={() => setFilterCat("all")} />
            {cats.slice(0, 5).map(c => <Pill key={c} label={c} active={filterCat === c} onClick={() => setFilterCat(c)} />)}
          </div>
        </div>

        <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, position: "relative", overflow: "hidden", height: H }}>
          {[0, 25, 50, 75, 100].map(y => (
            <div key={y} style={{ position: "absolute", left: pad, right: 20, top: `${(1 - y / 100) * (H - pad * 2) + pad}px`, height: 1, background: y === 50 ? D.border2 : D.border }}>
              <span style={{ position: "absolute", right: "100%", marginRight: 8, fontSize: 9.5, color: D.t3, fontFamily: D.mono, whiteSpace: "nowrap" }}>{y}٪</span>
            </div>
          ))}
          <div style={{ position: "absolute", left: pad + 16, top: 14, fontSize: 10, color: D.green, fontFamily: D.mono, opacity: .7 }}>سهم بالا — محافظت کنید</div>
          <div style={{ position: "absolute", left: pad + 16, bottom: pad + 16, fontSize: 10, color: D.red, fontFamily: D.mono, opacity: .7 }}>سهم پایین — رشد دهید</div>
          <div style={{ position: "absolute", left: pad, right: 20, top: pad, height: (H - pad * 2) * 0.4, background: "rgba(29,184,126,.04)", pointerEvents: "none" }} />
          <div style={{ position: "absolute", left: pad, right: 20, top: pad + (H - pad * 2) * 0.6, height: (H - pad * 2) * 0.4, background: "rgba(224,82,82,.04)", pointerEvents: "none" }} />

          <svg width="100%" height={H} viewBox={`0 0 ${W} ${H}`} style={{ position: "absolute", inset: 0 }}>
            {items.map(a => {
              const x = pad + (Math.log10(a.sessions + 10) / Math.log10(maxS + 10)) * (W - pad * 2);
              const y = pad + (1 - a.yktShare / 100) * (H - pad * 2);
              const r = Math.max(10, Math.sqrt(a.sessions / maxS) * 40);
              const isHov = hovered === a.name;
              const fillColor = a.yktShare >= 60 ? D.green : a.yktShare >= 40 ? D.accent : D.red;
              return (
                <g key={a.name} style={{ cursor: "pointer" }} onMouseEnter={() => setHovered(a.name)} onMouseLeave={() => setHovered(null)}>
                  <circle cx={x} cy={y} r={r + (isHov ? 4 : 0)} fill={fillColor} opacity={isHov ? .8 : .35} style={{ transition: "r .15s, opacity .15s" }} />
                  <circle cx={x} cy={y} r={4} fill={fillColor} opacity={.9} />
                  {(isHov || r > 24) && (
                    <text x={x} y={y - r - 6} textAnchor="middle" fontSize="9.5" fill={D.t2} fontFamily={D.mono} style={{ pointerEvents: "none" }}>
                      {a.name.split(".")[0].slice(0, 20)}
                    </text>
                  )}
                </g>
              );
            })}
          </svg>

          {hovItem && (
            <div style={{ position: "absolute", top: 16, left: 16, background: D.surface, border: `1px solid ${D.border2}`, borderRadius: 12, padding: "12px 14px", pointerEvents: "none", minWidth: 200, zIndex: 10 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: D.t1, fontFamily: D.mono, marginBottom: 6 }}>{hovItem.name}</div>
              <div style={{ display: "flex", gap: 14, fontSize: 11, color: D.t2, fontFamily: D.mono }}>
                <span>سشن: {formatNumber(hovItem.sessions)}</span>
                <span>یکتانت: {hovItem.yktShare}٪</span>
              </div>
              {hovItem.cat1 && <div style={{ fontSize: 10, color: D.t3, marginTop: 4 }}>{hovItem.cat1}</div>}
            </div>
          )}
        </div>

        <div style={{ display: "flex", gap: 16, justifyContent: "center" }}>
          {[{ color: D.green, label: "سهم بالای ۶۰٪" }, { color: D.accent, label: "سهم ۴۰–۶۰٪" }, { color: D.red, label: "سهم زیر ۴۰٪" }].map((l, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: l.color }} />
              <span style={{ fontSize: 11, color: D.t3 }}>{l.label}</span>
            </div>
          ))}
        </div>
      </div>
    );
  };

  // ─ Trends screen ─────────────────────────────────────────────────────────────
  const TrendsScreen = () => {
    const D = useD();
    const [view, setView] = useState<"heatmap" | "sparklines" | "winners">("heatmap");

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>برای مشاهده روندها ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    // Compute per-advertiser per-day yktShare
    const rows = csvData.rows;
    const dates = [...new Set(rows.map(r => r.Date || r.date).filter(Boolean))].sort().slice(-7);
    const advDayMap = new Map<string, { name: string; days: Map<string, { ykt: number; total: number }> }>();
    rows.forEach(r => {
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      const d = r.Date || r.date || "";
      if (!id || !d) return;
      if (!advDayMap.has(id)) advDayMap.set(id, { name: r.Advertiser_name || id, days: new Map() });
      const adv = advDayMap.get(id)!;
      const cur = adv.days.get(d) || { ykt: 0, total: 0 };
      cur.ykt += Number(r.Yektanet) || 0;
      cur.total += Number(r.Total_sessions) || 0;
      adv.days.set(d, cur);
    });

    // Build trend array per advertiser (% ykt per day)
    const advTrends = [...advDayMap.values()].map(a => {
      const trend = dates.map(d => {
        const day = a.days.get(d);
        return day && day.total > 0 ? Math.round(day.ykt / day.total * 100) : 0;
      });
      const totalYkt = [...a.days.values()].reduce((s, x) => s + x.ykt, 0);
      const totalSessions = [...a.days.values()].reduce((s, x) => s + x.total, 0);
      const yktShare = totalSessions > 0 ? Math.round(totalYkt / totalSessions * 100) : 0;
      return { name: a.name, trend, yktShare };
    }).filter(a => a.yktShare > 0).sort((a, b) => b.yktShare - a.yktShare).slice(0, 20);

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>تحلیل روند</div>
            <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>روند سهم یکتانت {dates.length} روز اخیر</div>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <Pill label="هیت‌مپ" active={view === "heatmap"} onClick={() => setView("heatmap")} />
            <Pill label="اسپارک‌لاین" active={view === "sparklines"} onClick={() => setView("sparklines")} />
            <Pill label="برنده/بازنده" active={view === "winners"} onClick={() => setView("winners")} />
          </div>
        </div>

        {view === "heatmap" && (
          <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "16px 20px", overflowX: "auto" }}>
            <div style={{ display: "grid", gridTemplateColumns: `160px repeat(${dates.length}, 1fr) 60px`, gap: 4, marginBottom: 8 }}>
              <div style={{ fontSize: 10, color: D.t3, fontFamily: D.mono }}>تبلیغ‌کننده</div>
              {dates.map(d => <div key={d} style={{ fontSize: 10, color: D.t3, fontFamily: D.mono, textAlign: "center" }}>{d.slice(5)}</div>)}
              <div style={{ fontSize: 10, color: D.t3, fontFamily: D.mono, textAlign: "center" }}>امروز</div>
            </div>
            {advTrends.map((a, i) => (
              <div key={a.name} className="fu" style={{ display: "grid", gridTemplateColumns: `160px repeat(${dates.length}, 1fr) 60px`, gap: 4, marginBottom: 3, animationDelay: `${i * 20}ms` }}>
                <div style={{ fontSize: 10.5, fontFamily: D.mono, color: D.t2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", paddingTop: 4 }}>{a.name}</div>
                {a.trend.map((v, j) => {
                  const intensity = v / 100;
                  const bg = v >= 60 ? `rgba(29,184,126,${intensity * .8})` : v >= 40 ? `rgba(212,98,58,${intensity * .8})` : `rgba(224,82,82,${(1 - intensity) * .8})`;
                  return (
                    <div key={j} style={{ height: 22, borderRadius: 4, background: bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
                      {v > 0 && <span style={{ fontSize: 9, fontFamily: D.mono, color: "rgba(255,255,255,.9)", fontWeight: 600 }}>{v}</span>}
                    </div>
                  );
                })}
                <div style={{ height: 22, borderRadius: 4, background: D.border, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <span style={{ fontSize: 10, fontFamily: D.mono, color: D.t1, fontWeight: 700 }}>{a.yktShare}٪</span>
                </div>
              </div>
            ))}
            <div style={{ display: "flex", gap: 4, marginTop: 12, alignItems: "center" }}>
              <span style={{ fontSize: 10, color: D.t3, marginLeft: 6 }}>کم</span>
              {[0, 20, 40, 60, 80, 100].map(v => (
                <div key={v} style={{ width: 24, height: 12, borderRadius: 3, background: v >= 60 ? `rgba(29,184,126,${v / 100 * .8})` : v >= 40 ? `rgba(212,98,58,${v / 100 * .8})` : `rgba(224,82,82,${(1 - v / 100) * .8})` }} />
              ))}
              <span style={{ fontSize: 10, color: D.t3, marginRight: 6 }}>زیاد</span>
            </div>
          </div>
        )}

        {view === "sparklines" && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
            {advTrends.map((a, i) => (
              <div key={a.name} className="fu" style={{ animationDelay: `${i * 30}ms`, background: D.card, border: `1px solid ${D.border}`, borderRadius: 12, padding: "14px 16px" }}>
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 8 }}>
                  <div style={{ fontSize: 11.5, fontFamily: D.mono, fontWeight: 600, color: D.t1, overflow: "hidden", textOverflow: "ellipsis", maxWidth: "70%" }}>{a.name}</div>
                  <div style={{ fontSize: 18, fontWeight: 800, color: a.yktShare >= 60 ? D.green : a.yktShare >= 40 ? D.accent : D.red, fontFamily: D.mono, lineHeight: 1 }}>{a.yktShare}٪</div>
                </div>
                <Sparkline data={a.trend} color={D.accent} w={180} h={32} filled />
                {a.trend.length > 1 && (
                  <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
                    <span style={{ fontSize: 9, color: D.t3, fontFamily: D.mono }}>{a.trend[0]}٪</span>
                    <span style={{ fontSize: 9, color: a.trend[a.trend.length - 1] >= a.trend[0] ? D.green : D.red, fontFamily: D.mono }}>
                      {a.trend[a.trend.length - 1] >= a.trend[0] ? "▲" : "▼"}{Math.abs(a.trend[a.trend.length - 1] - a.trend[0])}٪
                    </span>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {view === "winners" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <div style={{ background: D.card, border: `1px solid ${D.greenBrd}`, borderRadius: 16, padding: "18px 20px" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: D.green, marginBottom: 14 }}>▲ برنده‌های هفته</div>
              {[...advTrends].filter(a => a.trend.length > 1).sort((a, b) => (b.trend[b.trend.length - 1] - b.trend[0]) - (a.trend[a.trend.length - 1] - a.trend[0])).slice(0, 8).map((a, i) => (
                <div key={a.name} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: i < 7 ? `1px solid ${D.border}` : "none" }}>
                  <Sparkline data={a.trend} color={D.green} w={60} h={20} />
                  <div style={{ flex: 1, minWidth: 0, fontSize: 11, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</div>
                  <span style={{ fontSize: 12, fontFamily: D.mono, color: D.green, fontWeight: 700 }}>+{Math.max(0, a.trend[a.trend.length - 1] - a.trend[0])}٪</span>
                </div>
              ))}
            </div>
            <div style={{ background: D.card, border: `1px solid ${D.redBrd}`, borderRadius: 16, padding: "18px 20px" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: D.red, marginBottom: 14 }}>▼ بازنده‌های هفته</div>
              {[...advTrends].filter(a => a.trend.length > 1).sort((a, b) => (a.trend[a.trend.length - 1] - a.trend[0]) - (b.trend[b.trend.length - 1] - b.trend[0])).slice(0, 8).map((a, i) => (
                <div key={a.name} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: i < 7 ? `1px solid ${D.border}` : "none" }}>
                  <Sparkline data={a.trend} color={D.red} w={60} h={20} />
                  <div style={{ flex: 1, minWidth: 0, fontSize: 11, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</div>
                  <span style={{ fontSize: 12, fontFamily: D.mono, color: D.red, fontWeight: 700 }}>{Math.min(0, a.trend[a.trend.length - 1] - a.trend[0])}٪</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  // ─ Leads screen ───────────────────────────────────────────────────────────────
  const LeadsScreen = () => {
    const D = useD();
    const [filter, setFilter] = useState("all");
    const [sortBy2, setSortBy2] = useState("sessions");

    // Leads from result (AI-identified) or csvData (zero ykt)
    const rawLeads = useMemo(() => {
      const csvLeads: { name: string; id: string; cat1: string; am: string; sessions: number; score: number; priority: "high" | "medium" | "low" }[] = [];
      if (csvData) {
        const advSet2 = new Map<string, { name: string; id: string; cat1: string; am: string; sessions: number; ykt: number }>();
        csvData.rows.forEach(r => {
          const id = (r.Owner_id || r.Advertiser_name || "").trim();
          if (!id) return;
          const cur = advSet2.get(id) || { name: r.Advertiser_name || id, id, cat1: r.Category_level_1 || "", am: r.Account_manager_name || "", sessions: 0, ykt: 0 };
          cur.sessions += Number(r.Total_sessions) || 0;
          cur.ykt += Number(r.Yektanet) || 0;
          advSet2.set(id, cur);
        });
        [...advSet2.values()].filter(a => a.ykt === 0 && a.sessions > 1000).forEach(a => {
          const priority: "high" | "medium" | "low" = a.sessions > 50000 ? "high" : a.sessions > 10000 ? "medium" : "low";
          csvLeads.push({ ...a, score: Math.min(99, Math.round(Math.log10(a.sessions) * 20)), priority });
        });
      }
      return csvLeads;
    }, [csvData]);

    const filtered = filter === "all" ? rawLeads : rawLeads.filter(l => l.priority === filter);
    const sorted = [...filtered].sort((a, b) => sortBy2 === "score" ? b.score - a.score : b.sessions - a.sessions);

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    const counts = { high: rawLeads.filter(l => l.priority === "high").length, medium: rawLeads.filter(l => l.priority === "medium").length, low: rawLeads.filter(l => l.priority === "low").length };

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>لید پایپلاین</div>
            <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>تبلیغ‌کننده‌هایی که یکتانت در آن‌ها حضور ندارد</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            {([["high", "اولویت بالا", D.red], ["medium", "اولویت متوسط", D.amber], ["low", "اولویت پایین", D.t3]] as [string, string, string][]).map(([k, l, c]) => (
              <div key={k} style={{ padding: "8px 14px", borderRadius: 10, background: filter === k ? c + "22" : D.card, border: `1px solid ${filter === k ? c + "44" : D.border}`, cursor: "pointer", transition: "all .14s" }} onClick={() => setFilter(filter === k ? "all" : k)}>
                <div style={{ fontSize: 16, fontWeight: 800, color: c, fontFamily: D.mono, lineHeight: 1 }}>{counts[k as "high" | "medium" | "low"]}</div>
                <div style={{ fontSize: 9.5, color: D.t3, marginTop: 2 }}>{l}</div>
              </div>
            ))}
          </div>
          <select value={sortBy2} onChange={e => setSortBy2(e.target.value)} style={{ padding: "5px 10px", borderRadius: 8, border: `1px solid ${D.border2}`, background: D.card, color: D.t2, fontSize: 11, fontFamily: "Vazirmatn,sans-serif", cursor: "pointer", outline: "none" }}>
            <option value="score">امتیاز</option>
            <option value="sessions">سشن</option>
          </select>
        </div>

        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <Pill label="همه" active={filter === "all"} onClick={() => setFilter("all")} />
          <Pill label="اولویت بالا" active={filter === "high"} onClick={() => setFilter("high")} />
          <Pill label="اولویت متوسط" active={filter === "medium"} onClick={() => setFilter("medium")} />
        </div>

        {sorted.length === 0 ? (
          <div style={{ textAlign: "center", padding: "4rem", color: D.t3, fontSize: 13 }}>لیدی با این فیلتر یافت نشد</div>
        ) : sorted.map((a, i) => (
          <div key={a.name} className="fu" style={{ animationDelay: `${i * 30}ms`, background: D.card, border: `1px solid ${D.border}`, borderRadius: 13, padding: "14px 16px" }}>
            <div style={{ display: "flex", alignItems: "flex-start", gap: 12 }}>
              <div style={{ width: 46, height: 46, borderRadius: 12, background: a.priority === "high" ? D.redDim : a.priority === "medium" ? D.amberDim : D.border, border: `1px solid ${a.priority === "high" ? D.redBrd : a.priority === "medium" ? D.amberBrd : D.border}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <div style={{ textAlign: "center" }}>
                  <div style={{ fontSize: 14, fontWeight: 800, color: a.priority === "high" ? D.red : a.priority === "medium" ? D.amber : D.t3, fontFamily: D.mono, lineHeight: 1 }}>{a.score}</div>
                  <div style={{ fontSize: 8, color: D.t3 }}>امتیاز</div>
                </div>
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 13, fontWeight: 700, color: D.t1, fontFamily: D.mono }}>{a.name}</span>
                  <span style={{ fontSize: 9.5, color: D.t3, fontFamily: D.mono }}>#{a.id}</span>
                  {a.cat1 && <span style={{ fontSize: 9.5, color: D.blue, background: D.blueDim, padding: "1px 6px", borderRadius: 20 }}>{a.cat1}</span>}
                  <Badge label={a.priority === "high" ? "اولویت بالا" : a.priority === "medium" ? "اولویت متوسط" : "اولویت پایین"} color={a.priority === "high" ? D.red : a.priority === "medium" ? D.amber : D.t3} />
                </div>
                <div style={{ height: 4, borderRadius: 2, background: D.redDim, overflow: "hidden" }}>
                  <div style={{ width: "100%", height: "100%", background: D.red, borderRadius: 2, opacity: .3 }} />
                </div>
                <div style={{ fontSize: 10, color: D.t3, marginTop: 4 }}>هیچ سشنی از یکتانت ثبت نشده</div>
              </div>
              <div style={{ textAlign: "center", flexShrink: 0, minWidth: 80 }}>
                <div style={{ fontSize: 16, fontWeight: 800, color: D.t1, fontFamily: D.mono }}>{formatNumber(a.sessions)}</div>
                <div style={{ fontSize: 9.5, color: D.t3 }}>سشن</div>
                {a.am && <div style={{ fontSize: 10, color: D.accent, marginTop: 4, background: D.accentDim, borderRadius: 20, padding: "2px 8px" }}>{a.am.split(" ").slice(0, 2).join(" ")}</div>}
              </div>
            </div>
          </div>
        ))}
      </div>
    );
  };

  // ─ Industry screen ────────────────────────────────────────────────────────────
  const IndustryScreen = () => {
    const D = useD();
    const [selInd, setSelInd] = useState<string | null>(null);

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    const rows = csvData.rows;
    const indMap2 = new Map<string, { sessions: number; ykt: number; advertisers: number; names: string[] }>();
    const advSeen = new Map<string, string>();
    rows.forEach(r => {
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      if (!id) return;
      const ind = r.Category_level_1 || "سایر";
      if (!indMap2.has(ind)) indMap2.set(ind, { sessions: 0, ykt: 0, advertisers: 0, names: [] });
      const cur = indMap2.get(ind)!;
      cur.sessions += Number(r.Total_sessions) || 0;
      cur.ykt += Number(r.Yektanet) || 0;
      if (!advSeen.has(id)) { advSeen.set(id, ind); cur.advertisers++; cur.names.push(r.Advertiser_name || id); }
    });
    const indData = [...indMap2.entries()].map(([name, d], idx) => ({
      name, ...d,
      yktShare: d.sessions > 0 ? Math.round(d.ykt / d.sessions * 100) : 0,
      color: ["#D4623A","#3B82F6","#8B5CF6","#10B981","#F59E0B","#EF4444","#14B8A6","#6366F1"][idx % 8],
      risk: (d.sessions > 0 ? Math.round(d.ykt / d.sessions * 100) : 0) < 35 ? "high" : (d.sessions > 0 ? Math.round(d.ykt / d.sessions * 100) : 0) < 55 ? "medium" : "low",
    })).filter(i => i.name !== "سایر").sort((a, b) => b.sessions - a.sessions);

    const totalInd = indData.reduce((s, i) => s + i.sessions, 0);
    const highRisk = indData.filter(i => i.risk === "high").length;
    const stable = indData.filter(i => i.risk === "low").length;
    const avgShare = indData.length > 0 ? Math.round(indData.reduce((s, i) => s + i.yktShare, 0) / indData.length) : 0;

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        <div style={{ fontSize: 18, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>تحلیل صنایع</div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
          <KpiCard label="صنایع پر ریسک" value={highRisk} color={D.red} sub="سهم یکتانت زیر ۳۵٪" />
          <KpiCard label="صنایع پایدار" value={stable} color={D.green} sub="سهم یکتانت بالای ۵۵٪" delay={40} />
          <KpiCard label="میانگین سهم" value={`${avgShare}٪`} color={D.accent} sub="همه صنایع" delay={80} />
        </div>

        {/* Treemap-style block */}
        <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "20px" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: D.t1, marginBottom: 14 }}>حجم سشن به تفکیک صنعت</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 3 }}>
            {indData.map((ind, i) => {
              const w = Math.max(8, totalInd > 0 ? Math.round(ind.sessions / totalInd * 100) : 0);
              const isHov = selInd === ind.name;
              return (
                <div key={i} onClick={() => setSelInd(isHov ? null : ind.name)}
                  style={{ flex: `0 0 ${w}%`, minWidth: 60, height: 80, borderRadius: 8, background: isHov ? ind.color : ind.color + "33", border: `2px solid ${isHov ? ind.color : ind.color + "55"}`, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 3, transition: "all .15s", padding: "4px 6px" }}>
                  <div style={{ fontSize: isHov ? 11 : 10, fontWeight: 700, color: isHov ? "#fff" : ind.color, textAlign: "center", lineHeight: 1.3 }}>{ind.name}</div>
                  <div style={{ fontSize: 9.5, color: isHov ? "rgba(255,255,255,.8)" : D.t3, fontFamily: D.mono }}>{formatNumber(ind.sessions)}</div>
                  <div style={{ fontSize: 9.5, fontWeight: 700, color: isHov ? "#fff" : ind.yktShare >= 60 ? D.green : ind.yktShare >= 40 ? D.accent : D.red, fontFamily: D.mono }}>{ind.yktShare}٪</div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Detail table */}
        <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, overflow: "hidden" }}>
          <div style={{ padding: "14px 20px", borderBottom: `1px solid ${D.border}`, display: "grid", gridTemplateColumns: "1fr 80px 100px 80px 100px", gap: 12 }}>
            {["صنعت", "تبلیغ‌کننده", "مجموع سشن", "سهم یکتانت", "ریسک"].map(h => (
              <div key={h} style={{ fontSize: 10, fontWeight: 700, color: D.t3, textTransform: "uppercase", letterSpacing: ".06em" }}>{h}</div>
            ))}
          </div>
          {indData.map((ind, i) => (
            <div key={i} onClick={() => setSelInd(selInd === ind.name ? null : ind.name)}
              style={{ padding: "12px 20px", borderBottom: i < indData.length - 1 ? `1px solid ${D.border}` : "none", display: "grid", gridTemplateColumns: "1fr 80px 100px 80px 100px", gap: 12, alignItems: "center", cursor: "pointer", background: selInd === ind.name ? D.cardHov : "transparent", transition: "background .14s" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ width: 10, height: 10, borderRadius: "50%", background: ind.color, flexShrink: 0 }} />
                <span style={{ fontSize: 12, color: D.t1 }}>{ind.name}</span>
              </div>
              <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t2 }}>{ind.advertisers}</div>
              <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t2 }}>{formatNumber(ind.sessions)}</div>
              <div style={{ fontSize: 12, fontWeight: 700, fontFamily: D.mono, color: ind.yktShare >= 60 ? D.green : ind.yktShare >= 40 ? D.accent : D.red }}>{ind.yktShare}٪</div>
              <Badge label={ind.risk === "high" ? "پر ریسک" : ind.risk === "medium" ? "متوسط" : "ایمن"} color={ind.risk === "high" ? D.red : ind.risk === "medium" ? D.amber : D.green} />
            </div>
          ))}
        </div>
      </div>
    );
  };

  // ─ Team screen ────────────────────────────────────────────────────────────────
  const TeamScreen = () => {
    const D = useD();
    const [selTeam, setSelTeam] = useState<string | null>(null);

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    const rows = csvData.rows;
    const teamMap2 = new Map<string, { sessions: number; ykt: number; advertisers: Set<string>; spend: number; members: Set<string> }>();
    rows.forEach(r => {
      const team = r.Team || ""; if (!team) return;
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      const cur = teamMap2.get(team) || { sessions: 0, ykt: 0, advertisers: new Set(), spend: 0, members: new Set() };
      cur.sessions += Number(r.Total_sessions) || 0;
      cur.ykt += Number(r.Yektanet) || 0;
      cur.spend += Number(r.Daily_spend) || 0;
      if (id) cur.advertisers.add(id);
      if (r.Account_manager_name) cur.members.add(r.Account_manager_name);
      teamMap2.set(team, cur);
    });
    const teamColors = [D.accent, D.blue, D.green, D.amber, D.t2];
    const teamsData = [...teamMap2.entries()].map(([name, d], idx) => ({
      name, sessions: d.sessions, ykt: d.ykt, spend: d.spend,
      advertisers: d.advertisers.size, members: [...d.members],
      avgYkt: d.sessions > 0 ? Math.round(d.ykt / d.sessions * 100) : 0,
      color: teamColors[idx % teamColors.length],
    })).sort((a, b) => b.advertisers - a.advertisers);

    const maxAdv2 = Math.max(...teamsData.map(t => t.advertisers), 1);
    const sel2 = teamsData.find(t => t.name === selTeam) || teamsData[0];

    // Advertisers for selected team
    const selTeamAdvs = sel2 ? [...new Map(rows.filter(r => r.Team === sel2.name).map(r => {
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      return [id, { name: r.Advertiser_name || id, cat1: r.Category_level_1 || "" }];
    })).values()].slice(0, 10) : [];

    if (teamsData.length === 0) return <div style={{ textAlign: "center", padding: "4rem", color: D.t3, fontSize: 13 }}>داده تیم در فایل یافت نشد</div>;

    return (
      <div style={{ display: "flex", gap: 16, minHeight: 500 }}>
        {/* Team list sidebar */}
        <div style={{ width: 210, flexShrink: 0, background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "14px 10px", overflowY: "auto" }}>
          <div style={{ fontSize: 10, color: D.t3, textTransform: "uppercase", letterSpacing: ".08em", fontWeight: 600, marginBottom: 12, padding: "0 4px" }}>تیم‌ها</div>
          {teamsData.map((t, i) => (
            <button key={t.name} onClick={() => setSelTeam(t.name)}
              style={{ width: "100%", padding: "10px 12px", borderRadius: 10, border: `1px solid ${sel2?.name === t.name ? t.color + "44" : D.border}`, background: sel2?.name === t.name ? t.color + "11" : "transparent", cursor: "pointer", marginBottom: 6, transition: "all .14s", textAlign: "right", fontFamily: "Vazirmatn,sans-serif" }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: sel2?.name === t.name ? t.color : D.t1, marginBottom: 4 }}>{t.name}</div>
              <div style={{ fontSize: 10, color: D.t3, fontFamily: D.mono }}>{t.advertisers} تبلیغ‌کننده · {t.avgYkt}٪ یکتانت</div>
              <div style={{ height: 3, borderRadius: 2, background: D.border, marginTop: 6, overflow: "hidden" }}>
                <div style={{ width: `${Math.round(t.advertisers / maxAdv2 * 100)}%`, height: "100%", background: t.color, borderRadius: 2 }} />
              </div>
            </button>
          ))}
        </div>

        {/* Team detail */}
        {sel2 && (
          <div className="fu" style={{ flex: 1, display: "flex", flexDirection: "column", gap: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 16, background: D.card, border: `1px solid ${D.border}`, borderRadius: 16, padding: "20px" }}>
              <div style={{ width: 52, height: 52, borderRadius: 14, background: D.accentDim, border: `1px solid ${D.accentBrd}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={D.accent} strokeWidth="1.6"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 20, fontWeight: 800, color: D.t1, marginBottom: 8 }}>{sel2.name}</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {sel2.members.slice(0, 8).map((m, i) => (
                    <span key={i} style={{ fontSize: 11, color: D.accent, background: D.accentDim, border: `1px solid ${D.accentBrd}`, padding: "2px 9px", borderRadius: 20 }}>{m.split(" ").slice(0, 2).join(" ")}</span>
                  ))}
                </div>
              </div>
              <div style={{ textAlign: "center" }}>
                <div style={{ fontSize: 48, fontWeight: 900, color: D.accent, fontFamily: D.mono, lineHeight: 1, letterSpacing: "-2px" }}>{sel2.avgYkt}٪</div>
                <div style={{ fontSize: 10, color: D.t3 }}>میانگین سهم یکتانت</div>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14 }}>
              <KpiCard label="تبلیغ‌کننده فعال" value={sel2.advertisers} color={D.blue} />
              <KpiCard label="مجموع سشن" value={formatNumber(sel2.sessions)} color={D.accent} delay={40} />
              <KpiCard label="درآمد تخمینی" value={fmtMoney(sel2.spend)} color={D.amber} delay={80} />
            </div>

            <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 20px" }}>
              <SectionHeader title="تبلیغ‌کننده‌های این تیم" />
              {selTeamAdvs.map((a, i) => (
                <div key={a.name} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: i < selTeamAdvs.length - 1 ? `1px solid ${D.border}` : "none" }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</div>
                    {a.cat1 && <div style={{ fontSize: 10, color: D.t3, marginTop: 1 }}>{a.cat1}</div>}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  // ─ Alerts screen ─────────────────────────────────────────────────────────────
  const AlertsScreen = () => {
    const D = useD();
    const [filter, setFilter] = useState("all");

    // Generate alerts from csvData
    const generatedAlerts = useMemo(() => {
      if (!csvData) return [];
      const rows = csvData.rows;
      const advSet3 = new Map<string, { name: string; sessions: number; ykt: number; am: string; cat1: string }>();
      rows.forEach(r => {
        const id = (r.Owner_id || r.Advertiser_name || "").trim();
        if (!id) return;
        const cur = advSet3.get(id) || { name: r.Advertiser_name || id, sessions: 0, ykt: 0, am: r.Account_manager_name || "", cat1: r.Category_level_1 || "" };
        cur.sessions += Number(r.Total_sessions) || 0;
        cur.ykt += Number(r.Yektanet) || 0;
        advSet3.set(id, cur);
      });
      const alerts: { id: string; advertiser: string; msg: string; severity: "critical" | "high" | "medium" | "info"; type: "decline" | "lead" | "competitor" | "growth"; read: boolean; time: string }[] = [];
      let idx = 0;
      [...advSet3.values()].forEach(a => {
        const yktPct = a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0;
        if (a.ykt === 0 && a.sessions > 5000) {
          alerts.push({ id: String(idx++), advertiser: a.name, msg: `حجم سشن ${formatNumber(a.sessions)} — بدون حضور یکتانت. فرصت لید.`, severity: a.sessions > 30000 ? "critical" : "high", type: "lead", read: false, time: "امروز" });
        } else if (yktPct < 25 && a.sessions > 10000) {
          alerts.push({ id: String(idx++), advertiser: a.name, msg: `سهم یکتانت ${yktPct}٪ — بسیار پایین. رقبا در حال افزایش سهم.`, severity: "critical", type: "decline", read: false, time: "امروز" });
        } else if (yktPct < 40 && a.sessions > 5000) {
          alerts.push({ id: String(idx++), advertiser: a.name, msg: `سهم یکتانت ${yktPct}٪ — در خطر از دست دادن سهم بیشتر.`, severity: "high", type: "decline", read: false, time: "امروز" });
        }
      });
      return alerts.sort((a, b) => (b.severity === "critical" ? 1 : 0) - (a.severity === "critical" ? 1 : 0)).slice(0, 30);
    }, [csvData]);

    const [localAlerts, setLocalAlerts] = useState(generatedAlerts);
    const markRead = (id: string) => setLocalAlerts(as => as.map(a => a.id === id ? { ...a, read: true } : a));
    const markAll = () => setLocalAlerts(as => as.map(a => ({ ...a, read: true })));
    const filtered = filter === "all" ? localAlerts : filter === "unread" ? localAlerts.filter(a => !a.read) : localAlerts.filter(a => a.type === filter);

    const severityColor: Record<string, string> = { critical: D.red, high: D.amber, medium: D.amber, info: D.blue };

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>ابتدا فایل XLSX آپلود کنید تا هشدارها تولید شوند</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
        <div style={{ paddingBottom: 14, borderBottom: `1px solid ${D.border}`, marginBottom: 14 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 18, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>مرکز هشدار</div>
              <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>{localAlerts.filter(a => !a.read).length} هشدار خوانده‌نشده</div>
            </div>
            <button onClick={markAll} style={{ fontSize: 11, color: D.accent, background: D.accentDim, border: `1px solid ${D.accentBrd}`, borderRadius: 99, padding: "5px 14px", cursor: "pointer", fontFamily: "Vazirmatn,sans-serif" }}>همه را خوانده علامت بزن</button>
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {[["all", "همه"], ["unread", "خوانده‌نشده"], ["decline", "کاهش سهم"], ["lead", "لید"], ["competitor", "رقیب"]].map(([id, l]) => (
              <Pill key={id} label={l} active={filter === id} onClick={() => setFilter(id)} />
            ))}
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {filtered.length === 0 ? (
            <div style={{ textAlign: "center", padding: "4rem", color: D.t3, fontSize: 13 }}>هشداری با این فیلتر یافت نشد</div>
          ) : filtered.map((al, i) => (
            <div key={al.id} className="fu" style={{ animationDelay: `${i * 30}ms`, display: "flex", gap: 12, padding: "14px 16px", borderRadius: 12, transition: "all .14s", background: al.read ? D.card : D.surface, border: `1px solid ${al.read ? D.border : D.border2}`, opacity: al.read ? .75 : 1 }}>
              <div style={{ width: 36, height: 36, borderRadius: 10, flexShrink: 0, background: severityColor[al.severity] + "22", border: `1px solid ${severityColor[al.severity]}44`, display: "flex", alignItems: "center", justifyContent: "center", color: severityColor[al.severity] }}>
                {al.type === "decline" ? (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 18 23 18 23 12"/></svg>
                ) : (
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
                )}
              </div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4, flexWrap: "wrap" }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: D.t1, fontFamily: D.mono }}>{al.advertiser}</span>
                  <Badge label={al.severity === "critical" ? "بحرانی" : al.severity === "high" ? "بالا" : al.severity === "medium" ? "متوسط" : "اطلاع"} color={severityColor[al.severity]} />
                  {!al.read && <div style={{ width: 6, height: 6, borderRadius: "50%", background: D.accent }} />}
                </div>
                <div style={{ fontSize: 12, color: D.t2, lineHeight: 1.7 }}>{al.msg}</div>
                <div style={{ fontSize: 10, color: D.t3, fontFamily: D.mono, marginTop: 4 }}>{al.time}</div>
              </div>
              {!al.read && (
                <button onClick={() => markRead(al.id)} style={{ padding: "5px 12px", borderRadius: 8, border: `1px solid ${D.border}`, background: "transparent", color: D.t3, fontSize: 10, cursor: "pointer", flexShrink: 0, fontFamily: "Vazirmatn,sans-serif", alignSelf: "center" }}>خواندم</button>
              )}
            </div>
          ))}
        </div>
      </div>
    );
  };

  // ─ Brief screen ───────────────────────────────────────────────────────────────
  const BriefScreen = () => {
    const D = useD();
    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", color: D.t3 }}>
        <p style={{ fontSize: 15, marginBottom: 16 }}>ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 10, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "Vazirmatn,sans-serif" }}>رفتن به آپلود</button>
      </div>
    );

    // Compute from csvData
    const rows = csvData?.rows || [];
    const advSet4 = new Map<string, { name: string; sessions: number; ykt: number; cat1: string; am: string }>();
    rows.forEach(r => {
      const id = (r.Owner_id || r.Advertiser_name || "").trim();
      if (!id) return;
      const cur = advSet4.get(id) || { name: r.Advertiser_name || id, sessions: 0, ykt: 0, cat1: r.Category_level_1 || "", am: r.Account_manager_name || "" };
      cur.sessions += Number(r.Total_sessions) || 0;
      cur.ykt += Number(r.Yektanet) || 0;
      advSet4.set(id, cur);
    });
    const advList4 = [...advSet4.values()].map(a => ({ ...a, yktShare: a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0 }));
    const totalSessions4 = advList4.reduce((s, a) => s + a.sessions, 0);
    const totalYkt4 = advList4.reduce((s, a) => s + a.ykt, 0);
    const yktShare4 = totalSessions4 > 0 ? Math.round(totalYkt4 / totalSessions4 * 100) : 0;
    const declining4 = advList4.filter(a => a.yktShare > 0 && a.yktShare < 35).sort((a, b) => b.sessions - a.sessions).slice(0, 3);
    const leads4 = advList4.filter(a => a.ykt === 0 && a.sessions > 2000).sort((a, b) => b.sessions - a.sessions).slice(0, 3);
    const jalali = toJalali(new Date());

    const COMP_COLS4 = ["Tapsell","Deema","Tavoos","Adexo","Chavosh","Aparat","Daart","Yellowadwise","Najva","Triboon","Jaryan","Telewebion","Adverge","Soroush","Soroush_ny","Bale_ny","Rubika_ny","Eitaa_ny","Bazaar","Myket"];
    const TR_COMP4: Record<string, string> = { Tapsell:"تپسل",Deema:"دیما",Tavoos:"طاووس",Adexo:"ادکسو",Chavosh:"چاووش",Aparat:"آپارات",Daart:"دارت",Yellowadwise:"یلو ادوایز",Najva:"نجوا",Triboon:"تریبون",Jaryan:"جریان",Telewebion:"تلوبیون",Adverge:"ادورج",Soroush:"سروش",Soroush_ny:"سروش",Bale_ny:"بله",Rubika_ny:"روبیکا",Eitaa_ny:"ایتا",Bazaar:"بازار",Myket:"مایکت" };
    const topComps = COMP_COLS4.map(col => ({
      name: TR_COMP4[col] || col,
      total: rows.reduce((s, r) => s + (Number(r[col]) || 0), 0),
      color: AGENCY_COLORS[TR_COMP4[col] || col] || "#888",
    })).filter(c => c.total > 0).sort((a, b) => b.total - a.total).slice(0, 4);

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 800, color: D.t1, letterSpacing: "-.3px" }}>خلاصه هفتگی</div>
            <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>{jalali}</div>
          </div>
          <button onClick={() => { navigator.clipboard.writeText(editableMsg || ""); showToast("پیام کپی شد ✓"); }}
            style={{ display: "flex", alignItems: "center", gap: 8, padding: "9px 18px", borderRadius: 10, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "Vazirmatn,sans-serif", boxShadow: `0 2px 12px ${D.accentGlow}` }}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1"/></svg>
            کپی پیام
          </button>
        </div>

        {/* Summary header */}
        <div style={{ background: `linear-gradient(135deg,${D.accentDim},${D.blueDim})`, border: `1px solid ${D.accentBrd}`, borderRadius: 18, padding: "24px 28px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 20 }}>
            <div style={{ width: 44, height: 44, borderRadius: 14, background: D.accent, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: `0 4px 16px ${D.accentGlow}` }}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="9" stroke="#fff" strokeWidth="1.6"/><circle cx="12" cy="12" r="5" stroke="#fff" strokeWidth="1.3" opacity=".55"/><circle cx="12" cy="12" r="1.8" fill="#fff"/><line x1="12" y1="3" x2="12" y2="12" stroke="#fff" strokeWidth="1.6"/></svg>
            </div>
            <div>
              <div style={{ fontSize: 18, fontWeight: 800, color: D.t1 }}>خلاصه اجرایی Radar</div>
              <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono }}>{jalali} — یکتانت</div>
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
            {[
              { l: "سهم بازار", v: `${yktShare4}٪`, dc: D.accent },
              { l: "مجموع سشن", v: formatNumber(totalSessions4), dc: D.green },
              { l: "تبلیغ‌کننده فعال", v: advList4.length, dc: D.blue },
              { l: "لید جدید", v: leads4.length, dc: D.amber },
            ].map((s, i) => (
              <div key={i} style={{ padding: "12px 14px", borderRadius: 12, background: "rgba(0,0,0,.25)" }}>
                <div style={{ fontSize: 9.5, color: D.t3, marginBottom: 4 }}>{s.l}</div>
                <div style={{ fontSize: 22, fontWeight: 800, color: s.dc, fontFamily: D.mono, lineHeight: 1 }}>{s.v}</div>
              </div>
            ))}
          </div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14 }}>
          {/* Declining */}
          <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 18px" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: D.red, marginBottom: 12 }}>▼ کاهش سهم یکتانت</div>
            {declining4.length === 0 ? <div style={{ fontSize: 12, color: D.t3, textAlign: "center", padding: "16px 0" }}>موردی یافت نشد</div> : declining4.map((a, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</div>
                  <Sparkline data={[...Array(7)].map((_, j) => Math.max(0, a.yktShare + (j - 3) * 3))} color={D.red} w={100} h={20} />
                </div>
                <div style={{ fontSize: 14, fontWeight: 800, color: D.red, fontFamily: D.mono }}>{a.yktShare}٪</div>
              </div>
            ))}
          </div>

          {/* Competitors */}
          <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 18px" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: D.amber, marginBottom: 12 }}>⚡ رقبای فعال</div>
            {topComps.map((c, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                <div style={{ width: 8, height: 8, borderRadius: "50%", background: c.color, flexShrink: 0 }} />
                <span style={{ fontSize: 12, color: D.t1, flex: 1 }}>{c.name}</span>
                <span style={{ fontSize: 11, fontFamily: D.mono, color: D.t2 }}>{formatNumber(c.total)}</span>
              </div>
            ))}
          </div>

          {/* Top leads */}
          <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 18px" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: D.green, marginBottom: 12 }}>⭕ لیدهای اولویت‌دار</div>
            {leads4.length === 0 ? <div style={{ fontSize: 12, color: D.t3, textAlign: "center", padding: "16px 0" }}>لیدی یافت نشد</div> : leads4.map((a, i) => (
              <div key={i} style={{ padding: "9px 10px", borderRadius: 9, background: D.greenDim, border: `1px solid ${D.greenBrd}`, marginBottom: 7 }}>
                <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t1 }}>{a.name}</div>
                <div style={{ fontSize: 10, color: D.t3, marginTop: 2 }}>{formatNumber(a.sessions)} سشن — {a.cat1 || "—"}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Market summary */}
        {result?.market && (
          <div style={{ background: D.card, border: `1px solid ${D.border}`, borderRadius: 14, padding: "16px 18px" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: D.t1, marginBottom: 10 }}>تحلیل کلی بازار</div>
            <p style={{ fontSize: 13, color: D.t2, lineHeight: 1.9, margin: 0 }}>{result.market}</p>
          </div>
        )}
      </div>
    );
  };

  // ─ Explorer screen ────────────────────────────────────────────────────────────
  const ExplorerScreen = () => {
    const D = useD();
    const [searchEx, setSearchEx] = useState("");
    const [sortCol, setSortCol] = useState("sessions");
    const [sortDir, setSortDir] = useState<"asc" | "desc">("desc");
    const [page, setPage] = useState(0);
    const PER = 15;

    const allData = useMemo(() => {
      if (!csvData) return [];
      const advSet5 = new Map<string, { name: string; id: string; cat1: string; am: string; team: string; sessions: number; ykt: number; spend: number }>();
      csvData.rows.forEach(r => {
        const id = (r.Owner_id || r.Advertiser_name || "").trim();
        if (!id) return;
        const cur = advSet5.get(id) || { name: r.Advertiser_name || id, id, cat1: r.Category_level_1 || "", am: r.Account_manager_name || "", team: r.Team || "", sessions: 0, ykt: 0, spend: 0 };
        cur.sessions += Number(r.Total_sessions) || 0;
        cur.ykt += Number(r.Yektanet) || 0;
        cur.spend += Number(r.Daily_spend) || 0;
        advSet5.set(id, cur);
      });
      let items = [...advSet5.values()].map(a => ({ ...a, yktShare: a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0 }));
      if (searchEx) items = items.filter(a => a.name.toLowerCase().includes(searchEx.toLowerCase()) || a.id.includes(searchEx) || a.cat1.includes(searchEx) || a.am.includes(searchEx));
      items.sort((a, b) => {
        const av = sortCol === "ykt" ? a.yktShare : sortCol === "spend" ? a.spend : a.sessions;
        const bv = sortCol === "ykt" ? b.yktShare : sortCol === "spend" ? b.spend : b.sessions;
        return sortDir === "desc" ? bv - av : av - bv;
      });
      return items;
    }, [csvData, searchEx, sortCol, sortDir]);

    const pageData = allData.slice(page * PER, (page + 1) * PER);
    const totalPages = Math.ceil(allData.length / PER);
    const toggleSort = (k: string) => { if (sortCol === k) setSortDir(d => d === "desc" ? "asc" : "desc"); else { setSortCol(k); setSortDir("desc"); setPage(0); } };

    const exportCsv = () => {
      const header = "نام,شناسه,صنعت,اکانت منیجر,تیم,سشن,سهم یکتانت,هزینه";
      const rows2 = allData.map(a => `${a.name},${a.id},${a.cat1},${a.am},${a.team},${a.sessions},${a.yktShare}%,${a.spend}`);
      const blob = new Blob([header + "\n" + rows2.join("\n")], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url; link.download = "radar-export.csv"; link.click();
      URL.revokeObjectURL(url);
    };

    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero />
        <p style={{ margin: 0, fontSize: 14, color: D.t3 }}>ابتدا فایل XLSX آپلود کنید</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: D.accent, color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>رفتن به آپلود</button>
      </div>
    );

    const cols = [
      { key: "name", l: "تبلیغ‌کننده", w: "1fr" },
      { key: "id", l: "ID", w: "70px" },
      { key: "cat1", l: "صنعت", w: "130px" },
      { key: "am", l: "اکانت منیجر", w: "120px" },
      { key: "team", l: "تیم", w: "80px" },
      { key: "sessions", l: "سشن", w: "80px", sort: true },
      { key: "ykt", l: "یکتانت٪", w: "90px", sort: true },
      { key: "spend", l: "هزینه", w: "80px", sort: true },
    ];

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
        <div style={{ paddingBottom: 14, borderBottom: `1px solid ${D.border}`, marginBottom: 0, display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: D.t1 }}>اکسپلورر داده</div>
            <div style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, marginTop: 2 }}>{allData.length} تبلیغ‌کننده</div>
          </div>
          <input value={searchEx} onChange={e => { setSearchEx(e.target.value); setPage(0); }} placeholder="جستجو…"
            style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${D.border2}`, background: D.card, color: D.t1, fontSize: 12, fontFamily: "Vazirmatn,sans-serif", width: 200, outline: "none" }} />
          <button onClick={exportCsv} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 9, border: `1px solid ${D.accentBrd}`, background: D.accentDim, color: D.accent, fontSize: 12, cursor: "pointer", fontFamily: "Vazirmatn,sans-serif" }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Export CSV
          </button>
        </div>

        {/* Table header */}
        <div style={{ padding: "0", borderBottom: `1px solid ${D.border2}`, display: "grid", gridTemplateColumns: cols.map(c => c.w).join(" "), gap: 8, alignItems: "center", height: 38, marginTop: 4 }}>
          {cols.map(c => (
            <div key={c.key} onClick={c.sort ? () => toggleSort(c.key) : undefined}
              style={{ fontSize: 10, fontWeight: 700, color: c.sort && sortCol === c.key ? D.accent : D.t3, textTransform: "uppercase", letterSpacing: ".06em", cursor: c.sort ? "pointer" : "default", display: "flex", alignItems: "center", gap: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              {c.l}
              {c.sort && sortCol === c.key && <span style={{ fontSize: 8 }}>{sortDir === "desc" ? "▼" : "▲"}</span>}
            </div>
          ))}
        </div>

        {/* Table body */}
        <div style={{ display: "flex", flexDirection: "column" }}>
          {pageData.map((a, i) => (
            <div key={a.id} style={{ display: "grid", gridTemplateColumns: cols.map(c => c.w).join(" "), gap: 8, alignItems: "center", height: 40, borderBottom: `1px solid ${D.border}`, transition: "background .1s" }}
              onMouseEnter={e => (e.currentTarget as HTMLElement).style.background = D.cardHov}
              onMouseLeave={e => (e.currentTarget as HTMLElement).style.background = "transparent"}>
              <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</div>
              <div style={{ fontSize: 10, fontFamily: D.mono, color: D.t3 }}>{a.id}</div>
              <div style={{ fontSize: 10, color: D.t2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.cat1 || "—"}</div>
              <div style={{ fontSize: 10, color: D.t2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.am.split(" ").slice(0, 2).join(" ")}</div>
              <div style={{ fontSize: 10, fontFamily: D.mono, color: D.t3 }}>{a.team || "—"}</div>
              <div style={{ fontSize: 11, fontFamily: D.mono, color: D.t2, fontWeight: 600 }}>{formatNumber(a.sessions)}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                <div style={{ flex: 1, height: 3, borderRadius: 2, background: D.border, overflow: "hidden" }}>
                  <div style={{ width: `${a.yktShare}%`, height: "100%", background: a.yktShare >= 60 ? D.green : a.yktShare >= 40 ? D.accent : D.red, borderRadius: 2 }} />
                </div>
                <span style={{ fontSize: 10, fontFamily: D.mono, color: a.yktShare >= 60 ? D.green : a.yktShare >= 40 ? D.accent : D.red, width: 28, textAlign: "left" }}>{a.yktShare}٪</span>
              </div>
              <div style={{ fontSize: 10, fontFamily: D.mono, color: D.amber }}>{fmtMoney(a.spend)}</div>
            </div>
          ))}
        </div>

        {/* Pagination */}
        {totalPages > 1 && (
          <div style={{ padding: "10px 0", borderTop: `1px solid ${D.border}`, display: "flex", alignItems: "center", gap: 8, marginTop: 4 }}>
            <span style={{ fontSize: 11, color: D.t3, fontFamily: D.mono, flex: 1 }}>
              {page * PER + 1}–{Math.min((page + 1) * PER, allData.length)} از {allData.length}
            </span>
            {Array.from({ length: Math.min(totalPages, 10) }, (_, i) => (
              <button key={i} onClick={() => setPage(i)} style={{ width: 28, height: 28, borderRadius: 7, border: `1px solid ${page === i ? D.accentBrd : D.border}`, background: page === i ? D.accentDim : "transparent", color: page === i ? D.accent : D.t3, fontSize: 11, cursor: "pointer", fontFamily: D.mono, transition: "all .12s" }}>{i + 1}</button>
            ))}
          </div>
        )}
      </div>
    );
  };

  const tokens = useMemo(() => makeTokens(isDark), [isDark]);

  const topbarYktShare = useMemo<number | null>(() => {
    if (!csvData) return null;
    const rows = csvData.rows;
    const totalSessions = rows.reduce((s, r) => s + (Number(r.Total_sessions) || 0), 0);
    const totalYkt = rows.reduce((s, r) => s + (Number(r.Yektanet) || 0), 0);
    return totalSessions > 0 ? Math.round(totalYkt / totalSessions * 100) : null;
  }, [csvData]);

  const alertCount = useMemo(() => {
    if (!result) return 0;
    return result.advertisers.filter(a => {
      const total = a.agencies.reduce((s, ag) => s + ag.value, 0);
      const ykt = a.agencies.find(ag => ag.name === "یکتانت")?.value ?? 0;
      return total > 0 && ykt / total < 0.35;
    }).length;
  }, [result]);

  if (!session) return (
    <ThemeCtx.Provider value={tokens}>
      <style>{GLOBAL_CSS}</style>
      <LoginPage onLogin={setSession} />
    </ThemeCtx.Provider>
  );

  return (
    <ThemeCtx.Provider value={tokens}>
      <div style={{ minHeight: "100dvh", background: tokens.bg, color: tokens.t1, direction: "rtl", fontFamily: "'Vazirmatn', system-ui, sans-serif", overflowX: "hidden" }}>
        <style>{GLOBAL_CSS}</style>
        <style>{`
          ::-webkit-scrollbar { width: 5px; height: 5px; }
          ::-webkit-scrollbar-track { background: transparent; }
          ::-webkit-scrollbar-thumb { background: ${tokens.border2}; border-radius: 10px; }
          select option { background: ${tokens.card}; color: ${tokens.t1}; }
          button:focus-visible, input:focus-visible, select:focus-visible, textarea:focus-visible { outline: 2px solid ${tokens.accent}; outline-offset: 2px; border-radius: 4px; }
        `}</style>

        <Topbar screen={screen} yktShare={topbarYktShare} alertCount={alertCount} setScreen={setScreen} />
        <Sidebar screen={screen} setScreen={setScreen} isDark={isDark} setIsDark={setIsDark} hasResult={!!result} hasCsv={!!csvData} session={session} onLogout={handleLogout} alertCount={alertCount} T={tokens} />

        <div style={{ marginRight: 200 }}>
          <div style={{ maxWidth: 1200, width: "100%", margin: "0 auto", padding: "92px 32px 80px" }}>
            {screen === "upload" && <UploadScreen />}
            {screen === "dashboard" && <DashboardScreen />}
            {screen === "results" && <ResultsScreen />}
            {screen === "competitor" && <CompetitorScreen />}
            {screen === "marketmap" && <MarketMapScreen />}
            {screen === "trends" && <TrendsScreen />}
            {screen === "leads" && <LeadsScreen />}
            {screen === "industry" && <IndustryScreen />}
            {screen === "team" && <TeamScreen />}
            {screen === "alerts" && <AlertsScreen />}
            {screen === "brief" && <BriefScreen />}
            {screen === "explorer" && <ExplorerScreen />}
            {screen === "reports" && <ReportsScreen />}
            {screen === "history" && <HistoryScreen />}
            {screen === "settings" && <SettingsScreen />}
          </div>
        </div>

        {step === "loading" && (
          <div style={{ position: "fixed", inset: 0, zIndex: 9000, background: "rgba(0,0,0,0.55)", backdropFilter: "blur(4px)", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 24 }}>
            <RadarHero />
            <div style={{ width: 320 }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                <span style={{ fontSize: 13, color: "#fff" }}>در حال آنالیز داده‌ها...</span>
                <span style={{ fontSize: 13, color: tokens.accent, fontWeight: 600 }}>{Math.round(loadingProgress)}٪</span>
              </div>
              <div style={{ height: 4, background: "rgba(255,255,255,0.15)", borderRadius: 4, overflow: "hidden" }}>
                <div style={{ height: "100%", background: tokens.accent, borderRadius: 4, width: `${loadingProgress}%`, transition: "width 0.4s ease" }} />
              </div>
            </div>
            <p style={{ margin: 0, fontSize: 12, color: "rgba(255,255,255,0.6)" }}>Claude در حال بررسی رقابت بازار است</p>
          </div>
        )}
        {modalAdv && (
          <DetailModal adv={modalAdv.adv} type={modalAdv.type} onClose={() => setModalAdv(null)}
            onRegen={() => regenOne(modalAdv.adv, modalAdv.type)} regenLoading={regenLoading === modalAdv.adv.name} />
        )}
        {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} />}
      </div>
    </ThemeCtx.Provider>
  );
}
