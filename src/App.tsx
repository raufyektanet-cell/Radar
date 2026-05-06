import { useState, useCallback, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { TEAMS, ADMIN_USER, type SessionUser } from "./teams";

const MODEL = "anthropic/claude-sonnet-4-5";
const OPENROUTER_KEY = import.meta.env.VITE_OPENROUTER_KEY as string | undefined;
const HISTORY_KEY = "analyzer:reported_advertisers";
const REPORTS_KEY = "analyzer:saved_reports";
const THEME_KEY = "analyzer:theme";
const SESSION_KEY = "radar:session";
const APP_PASSWORD = (import.meta.env.VITE_APP_PASSWORD as string | undefined) || "radar1403";

const DARK = {
  bg: "#0C0C0E", surface: "#141416", surface2: "#1E1E23",
  border: "#26262C", border2: "#30303A",
  text1: "#F0EDE8", text2: "#908E89", text3: "#55524C",
  coral: "#D4623A", coralDim: "rgba(212,98,58,0.14)", coralBorder: "rgba(212,98,58,0.30)",
  green: "#1D9E75", greenDim: "rgba(29,158,117,0.14)", greenBorder: "rgba(29,158,117,0.30)",
  amber: "#C07B28", amberDim: "rgba(192,123,40,0.14)", amberBorder: "rgba(192,123,40,0.30)",
  danger: "#EF4444", dangerBg: "rgba(239,68,68,0.10)", dangerBorder: "rgba(239,68,68,0.28)",
  successBg: "rgba(29,158,117,0.12)", successBorder: "rgba(29,158,117,0.28)", successText: "#34D399",
  blue: "#4A9EE8", blueDim: "rgba(74,158,232,0.13)",
};

const LIGHT = {
  bg: "#F5F0E8", surface: "#FFFFFF", surface2: "#EDE8DF",
  border: "#E0DBD2", border2: "#D3CFC7",
  text1: "#2C2C2A", text2: "#5F5E5A", text3: "#888780",
  coral: "#D4623A", coralDim: "rgba(212,98,58,0.10)", coralBorder: "rgba(212,98,58,0.22)",
  green: "#1D9E75", greenDim: "rgba(29,158,117,0.10)", greenBorder: "rgba(29,158,117,0.22)",
  amber: "#854F0B", amberDim: "rgba(186,117,23,0.10)", amberBorder: "rgba(186,117,23,0.22)",
  danger: "#DC2626", dangerBg: "#FEF2F2", dangerBorder: "#FECACA",
  successBg: "#F0FDF4", successBorder: "#BBF7D0", successText: "#15803D",
  blue: "#185FA5", blueDim: "rgba(55,138,221,0.10)",
};

type Theme = typeof DARK;

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
  brief: "Be very concise — max 1 sentence per advertiser. Only the single most important change.",
  detailed: "Be detailed — up to 4 sentences per advertiser. Include all agencies and full trend.",
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

// ── UI Components ─────────────────────────────────────────────────────────────

function AgencyDonut({ agencies, size = 72 }: { agencies: Agency[]; size?: number }) {
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
      <circle cx={cx} cy={cy} r={r} fill="none" stroke="rgba(128,128,128,0.12)" strokeWidth={size * 0.11} />
      {slices.map((s, i) => (
        <circle key={i} cx={cx} cy={cy} r={r} fill="none" stroke={s.color} strokeWidth={size * 0.11}
          strokeDasharray={`${s.dash} ${s.gap}`} strokeDashoffset={s.dashoffset}
          transform={`rotate(-90 ${cx} ${cy})`} strokeLinecap="round" />
      ))}
      <text x={cx} y={cy} textAnchor="middle" dominantBaseline="middle"
        fontSize={size * 0.17} fontWeight="600" fill={ykt ? "#D4623A" : "#888"} fontFamily="system-ui">
        {yPct}٪
      </text>
    </svg>
  );
}

function AgencyLegend({ agencies, T }: { agencies: Agency[]; T: Theme }) {
  const total = agencies.reduce((s, a) => s + a.value, 0);
  if (!total) return null;
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: "3px 10px", marginTop: 6 }}>
      {agencies.map((a, i) => (
        <span key={i} style={{ fontSize: 11, color: T.text3, display: "flex", alignItems: "center", gap: 4 }}>
          <span style={{ width: 7, height: 7, borderRadius: "50%", background: a.color, display: "inline-block", flexShrink: 0 }} />
          {a.name} {Math.round(a.value / total * 100)}٪
        </span>
      ))}
    </div>
  );
}

function RadarHero({ T }: { T: Theme }) {
  return (
    <div style={{ position: "relative", width: 200, height: 200, margin: "0 auto 28px" }}>
      <svg width="200" height="200" viewBox="0 0 200 200">
        <circle cx="100" cy="100" r="85" fill="none" stroke={T.coral} strokeWidth="0.5" opacity="0.15" />
        <circle cx="100" cy="100" r="60" fill="none" stroke={T.coral} strokeWidth="0.5" opacity="0.22" />
        <circle cx="100" cy="100" r="36" fill="none" stroke={T.coral} strokeWidth="0.5" opacity="0.32" />
        <line x1="100" y1="15" x2="100" y2="185" stroke={T.coral} strokeWidth="0.4" opacity="0.1" />
        <line x1="15" y1="100" x2="185" y2="100" stroke={T.coral} strokeWidth="0.4" opacity="0.1" />
        <g style={{ transformOrigin: "100px 100px", animation: "radarSweep 3s linear infinite" }}>
          <path d="M100,100 L100,15 A85,85 0 0,1 185,100 Z" fill={T.coral} opacity="0.08" />
          <line x1="100" y1="100" x2="100" y2="15" stroke={T.coral} strokeWidth="1.5" opacity="0.6" />
        </g>
        <circle cx="142" cy="68" r="4" fill={T.coral} style={{ animation: "dotPop 3s 0.9s ease-in-out infinite" }} />
        <circle cx="72" cy="140" r="3" fill={T.green} style={{ animation: "dotPop 3s 1.7s ease-in-out infinite" }} />
        <circle cx="158" cy="128" r="2.5" fill={T.amber} style={{ animation: "dotPop 3s 2.3s ease-in-out infinite" }} />
        <circle cx="100" cy="100" r="5" fill={T.coral} />
        <circle cx="100" cy="100" r="5" fill={T.coral} opacity="0.6" style={{ animation: "radarPulse 2.4s ease-out infinite" }} />
      </svg>
    </div>
  );
}

function Toast({ msg, type, onDone, T }: { msg: string; type: "success" | "error"; onDone: () => void; T: Theme }) {
  useEffect(() => { const t = setTimeout(onDone, 3000); return () => clearTimeout(t); }, [onDone]);
  const bg = type === "success" ? T.green : T.danger;
  return (
    <div style={{ position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)", background: bg, color: "#fff", padding: "10px 20px", borderRadius: 12, fontSize: 13, fontWeight: 500, zIndex: 9999, animation: "fadeUp 0.2s ease both", whiteSpace: "nowrap", boxShadow: "0 4px 20px rgba(0,0,0,0.3)" }}>
      {msg}
    </div>
  );
}

function DetailModal({ adv, type, T, onClose, onRegen, regenLoading }: { adv: Advertiser; type: string; T: Theme; onClose: () => void; onRegen: () => void; regenLoading: boolean }) {
  useEffect(() => {
    const esc = (e: KeyboardEvent) => { if (e.key === "Escape") onClose(); };
    document.addEventListener("keydown", esc);
    return () => document.removeEventListener("keydown", esc);
  }, [onClose]);
  const total = adv.agencies.reduce((s, a) => s + a.value, 0);
  return (
    <div role="dialog" aria-modal="true" aria-label={`جزئیات ${adv.name}`} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", zIndex: 500, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }} onClick={onClose}>
      <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 20, padding: "24px 28px", width: "100%", maxWidth: 480, animation: "slideInModal 0.22s ease both", direction: "rtl" }} onClick={e => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 600, color: T.text1, direction: "ltr" }}>{adv.name}</div>
            <div style={{ fontSize: 12, color: T.text3, marginTop: 3, direction: "ltr" }}>#{adv.ownerid}{adv.manager ? ` · ` : ""}{adv.manager && <span style={{ color: T.coral }}>{adv.manager}</span>}</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onRegen} disabled={regenLoading} aria-label="بازسازی تحلیل" style={{ padding: "6px 12px", borderRadius: 8, border: `1px solid ${T.border}`, background: "transparent", color: T.text2, fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 5 }}>
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={regenLoading ? { animation: "spin 0.8s linear infinite" } : {}}><polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10" /></svg>
              بازسازی
            </button>
            <button onClick={onClose} aria-label="بستن" style={{ width: 40, height: 40, borderRadius: 8, border: `1px solid ${T.border}`, background: "transparent", color: T.text3, cursor: "pointer", fontSize: 20, lineHeight: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>×</button>
          </div>
        </div>
        {total > 0 && (
          <div style={{ marginBottom: 18 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 24, marginBottom: 14 }}>
              <AgencyDonut agencies={adv.agencies} size={100} />
              <AgencyLegend agencies={adv.agencies} T={T} />
            </div>
            <div style={{ display: "flex", borderRadius: 6, overflow: "hidden", height: 6 }}>
              {adv.agencies.map((a, i) => <div key={i} style={{ flex: a.value, background: a.color }} />)}
            </div>
          </div>
        )}
        {(adv.industry1 || adv.industry2) && (
          <div style={{ display: "flex", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
            {adv.industry1 && <span style={{ fontSize: 12, color: T.blue, background: T.blueDim, padding: "4px 12px", borderRadius: 20, border: `1px solid rgba(74,158,232,0.2)` }}>{adv.industry1}</span>}
            {adv.industry2 && <span style={{ fontSize: 12, color: T.text3, background: T.surface2, padding: "4px 12px", borderRadius: 20, border: `1px solid ${T.border}` }}>{adv.industry2}</span>}
          </div>
        )}
        {(adv.team || adv.manager || adv.performanceManager || adv.supervisor) && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 14 }}>
            {adv.team && <div style={{ padding: "8px 12px", borderRadius: 10, background: T.surface2, border: `1px solid ${T.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: T.text3 }}>تیم</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: T.text1 }}>{adv.team}</p></div>}
            {adv.manager && <div style={{ padding: "8px 12px", borderRadius: 10, background: T.surface2, border: `1px solid ${T.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: T.text3 }}>اکانت منیجر</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: T.coral }}>{adv.manager}</p></div>}
            {adv.performanceManager && <div style={{ padding: "8px 12px", borderRadius: 10, background: T.surface2, border: `1px solid ${T.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: T.text3 }}>پرفورمنس منیجر</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: T.blue }}>{adv.performanceManager}</p></div>}
            {adv.supervisor && <div style={{ padding: "8px 12px", borderRadius: 10, background: T.surface2, border: `1px solid ${T.border}` }}><p style={{ margin: "0 0 2px", fontSize: 10, color: T.text3 }}>سوپروایزر</p><p style={{ margin: 0, fontSize: 12, fontWeight: 600, color: T.amber }}>{adv.supervisor}</p></div>}
          </div>
        )}
        <div style={{ background: T.surface2, borderRadius: 12, padding: "14px 16px" }}>
          <p style={{ margin: 0, fontSize: 14, color: T.text2, lineHeight: 2, direction: "rtl" }}>
            {type === "lead" ? adv.note : adv.summary}
          </p>
        </div>
      </div>
    </div>
  );
}

// ── Login Page ────────────────────────────────────────────────────────────────

function LoginPage({ T, onLogin }: { T: Theme; onLogin: (u: SessionUser) => void }) {
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

  const deptColors: Record<string, string> = { "بیزینس": T.coral, "محصول": T.blue, "فروش": T.green, "": T.text3 };

  return (
    <div style={{ minHeight: "100dvh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24, direction: "rtl", fontFamily: "'Vazirmatn', system-ui, sans-serif" }}>
      <style>{`@keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}} *{box-sizing:border-box}`}</style>
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

// ── Sidebar ───────────────────────────────────────────────────────────────────

function Sidebar({ screen, setScreen, isDark, setIsDark, hasResult, hasCsv, session, onLogout, T }: {
  screen: string; setScreen: (s: string) => void; isDark: boolean; setIsDark: (v: boolean) => void; hasResult: boolean; hasCsv: boolean; session: SessionUser; onLogout: () => void; T: Theme;
}) {
  const items = [
    { id: "upload", label: "آنالیز", icon: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg> },
    { id: "dashboard", label: "داشبورد", icon: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg>, disabled: !hasCsv },
    { id: "results", label: "نتایج", icon: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><rect x="3" y="3" width="7" height="7" rx="1" /><rect x="14" y="3" width="7" height="7" rx="1" /><rect x="3" y="14" width="7" height="7" rx="1" /><rect x="14" y="14" width="7" height="7" rx="1" /></svg>, disabled: !hasResult },
    { id: "reports", label: "گزارش‌ها", icon: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /></svg> },
    { id: "history", label: "تاریخچه", icon: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 16 14" /></svg> },
    { id: "settings", label: "تنظیمات", icon: <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8"><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z" /></svg> },
  ];
  return (
    <div style={{ position: "fixed", right: 0, top: 0, bottom: 0, width: 80, background: T.surface, borderLeft: `1px solid ${T.border}`, display: "flex", flexDirection: "column", alignItems: "center", paddingTop: 20, paddingBottom: 20, zIndex: 200, gap: 2 }}>
      <div style={{ width: 38, height: 38, borderRadius: 11, background: T.coral, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 16, flexShrink: 0 }}>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="12" r="10" stroke="#fff" strokeWidth="1.5" /><circle cx="12" cy="12" r="6" stroke="#fff" strokeWidth="1" opacity="0.6" /><circle cx="12" cy="12" r="2.5" fill="#fff" /></svg>
      </div>
      {items.map(item => {
        const active = screen === item.id;
        return (
          <button key={item.id} onClick={() => !item.disabled && setScreen(item.id)} title={item.label} aria-label={item.label} aria-current={active ? "page" : undefined}
            style={{ width: 48, height: 48, borderRadius: 13, border: "none", cursor: item.disabled ? "default" : "pointer", background: active ? T.coralDim : "transparent", color: active ? T.coral : item.disabled ? T.text3 : T.text2, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 3, transition: "all 0.18s", opacity: item.disabled ? 0.4 : 1 }}>
            {item.icon}
            <span style={{ fontSize: 8.5, fontWeight: active ? 600 : 400, letterSpacing: "0.01em" }}>{item.label}</span>
          </button>
        );
      })}
      <div style={{ flex: 1 }} />
      {/* User avatar */}
      <div title={`${session.managerFa} — ${session.teamFa}`} aria-label={`کاربر: ${session.managerFa}`} role="img"
        style={{ width: 44, height: 44, borderRadius: "50%", background: T.coralDim, border: `1.5px solid ${T.coralBorder}`, display: "flex", alignItems: "center", justifyContent: "center", color: T.coral, fontWeight: 700, fontSize: 14, marginBottom: 6, flexShrink: 0 }}>
        {session.managerFa.slice(0, 1)}
      </div>
      {/* Theme toggle */}
      <button onClick={() => setIsDark(!isDark)} title={isDark ? "حالت روشن" : "حالت تاریک"} aria-label={isDark ? "تغییر به حالت روشن" : "تغییر به حالت تاریک"}
        style={{ width: 44, height: 44, borderRadius: 10, border: `1px solid ${T.border}`, background: T.surface2, color: T.text2, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.18s", marginBottom: 6 }}>
        {isDark
          ? <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><circle cx="12" cy="12" r="5" /><line x1="12" y1="1" x2="12" y2="3" /><line x1="12" y1="21" x2="12" y2="23" /><line x1="4.22" y1="4.22" x2="5.64" y2="5.64" /><line x1="18.36" y1="18.36" x2="19.78" y2="19.78" /><line x1="1" y1="12" x2="3" y2="12" /><line x1="21" y1="12" x2="23" y2="12" /><line x1="4.22" y1="19.78" x2="5.64" y2="18.36" /><line x1="18.36" y1="5.64" x2="19.78" y2="4.22" /></svg>
          : <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z" /></svg>}
      </button>
      {/* Logout */}
      <button onClick={onLogout} title="خروج" aria-label="خروج از حساب"
        style={{ width: 44, height: 44, borderRadius: 10, border: `1px solid ${T.dangerBorder}`, background: T.dangerBg, color: T.danger, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.18s" }}>
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
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
  const T = isDark ? DARK : LIGHT;

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
    setErrMsg(""); setRawDebug(""); setShowDebug(false); setStep("loading"); setScreen("upload");
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
        <AgencyDonut agencies={adv.agencies} size={64} />
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
          <AgencyLegend agencies={adv.agencies} T={T} />
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
      <RadarHero T={T} />
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
            <RadarHero T={T} />
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
    if (!csvData) return (
      <div style={{ textAlign: "center", padding: "5rem 2rem", display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
        <RadarHero T={T} />
        <p style={{ margin: 0, fontSize: 15, fontWeight: 600, color: T.text1 }}>داشبورد آنالیز رقابتی</p>
        <p style={{ margin: 0, fontSize: 13, color: T.text3 }}>ابتدا یک فایل XLSX آپلود کنید تا داشبورد فعال شود</p>
        <button onClick={() => setScreen("upload")} style={{ padding: "10px 24px", borderRadius: 12, border: "none", background: T.coral, color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer" }}>رفتن به آنالیز</button>
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
    const dates = [...new Set(rows.map(r => r.Date || r.date).filter(Boolean))];

    // Competitor totals
    const compTotals = COMP_COLS.map(col => ({
      col, name: TR_COMP[col] || col,
      total: advList.reduce((s, a) => s + (a.comps[col] || 0), 0),
    })).filter(c => c.total > 0).sort((a, b) => b.total - a.total).slice(0, 10);
    const compMax = compTotals[0]?.total || 1;

    // Industry breakdown (by unique advertiser)
    const indMap = new Map<string, { sessions: number; ykt: number; count: number }>();
    advList.forEach(a => {
      const ind = a.industry || "سایر";
      const cur = indMap.get(ind) || { sessions: 0, ykt: 0, count: 0 };
      cur.sessions += a.sessions; cur.ykt += a.ykt; cur.count++;
      indMap.set(ind, cur);
    });
    const indStats = [...indMap.entries()].map(([name, { sessions, ykt, count }]) => ({
      name, count, sessions,
      yktShare: sessions > 0 ? Math.round(ykt / sessions * 100) : 0,
    })).sort((a, b) => b.sessions - a.sessions);

    // AM performance
    const amMap = new Map<string, { sessions: number; ykt: number; accounts: number; spend: number }>();
    advList.forEach(a => {
      if (!a.am) return;
      const cur = amMap.get(a.am) || { sessions: 0, ykt: 0, accounts: 0, spend: 0 };
      cur.sessions += a.sessions; cur.ykt += a.ykt; cur.accounts++; cur.spend += a.spend;
      amMap.set(a.am, cur);
    });
    const amStats = [...amMap.entries()].map(([name, d]) => ({
      name, accounts: d.accounts,
      yktShare: d.sessions > 0 ? Math.round(d.ykt / d.sessions * 100) : 0,
      sessions: d.sessions, spend: d.spend,
    })).sort((a, b) => b.accounts - a.accounts);

    // Team performance
    const teamMap = new Map<string, { sessions: number; ykt: number; accounts: number }>();
    advList.forEach(a => {
      if (!a.team) return;
      const cur = teamMap.get(a.team) || { sessions: 0, ykt: 0, accounts: 0 };
      cur.sessions += a.sessions; cur.ykt += a.ykt; cur.accounts++;
      teamMap.set(a.team, cur);
    });
    const teamStats = [...teamMap.entries()].map(([name, d]) => ({
      name, accounts: d.accounts,
      yktShare: d.sessions > 0 ? Math.round(d.ykt / d.sessions * 100) : 0,
    })).sort((a, b) => b.accounts - a.accounts);

    // At-risk: ykt share < 35%, sessions > 5000
    const atRisk = advList.filter(a => a.sessions > 5000)
      .map(a => ({ ...a, yktShare: a.sessions > 0 ? Math.round(a.ykt / a.sessions * 100) : 0 }))
      .filter(a => a.yktShare > 0 && a.yktShare < 35)
      .sort((a, b) => b.sessions - a.sessions).slice(0, 8);

    // Opportunities: zero ykt, significant sessions
    const opportunities = advList.filter(a => a.ykt === 0 && a.sessions > 2000)
      .sort((a, b) => b.sessions - a.sessions).slice(0, 8);

    const yktColor = yktShare >= 50 ? T.coral : yktShare >= 30 ? T.amber : T.danger;

    return (
      <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div>
            <p style={{ margin: "0 0 2px", fontSize: 18, fontWeight: 700, color: T.text1 }}>داشبورد رقابتی</p>
            <p style={{ margin: 0, fontSize: 12, color: T.text3 }}>{advList.length} آگهی‌دهنده · {dates.length} روز داده · {csvData.name}</p>
          </div>
          <button onClick={() => setScreen("upload")} style={{ padding: "8px 18px", borderRadius: 10, border: `1px solid ${T.coralBorder}`, background: T.coralDim, color: T.coral, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>شروع آنالیز AI</button>
        </div>

        {/* KPI row */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))", gap: 12 }}>
          {[
            { label: "آگهی‌دهندگان", value: advList.length.toLocaleString(), color: T.text1 },
            { label: "کل سشن‌ها", value: formatNumber(totalSessions), color: T.text1 },
            { label: "سهم یکتانت", value: `${yktShare}٪`, color: yktColor },
            { label: "روزهای داده", value: dates.length, color: T.blue },
            { label: "بودجه یکتانت (M)", value: totalSpend > 0 ? `${(totalSpend / 1000000).toFixed(0)}M` : "—", color: T.green },
            { label: "در معرض خطر", value: atRisk.length, color: T.amber },
          ].map((kpi, i) => (
            <div key={i} style={card({ padding: "16px 18px" })}>
              <p style={{ margin: "0 0 6px", fontSize: 11, color: T.text3 }}>{kpi.label}</p>
              <p style={{ margin: 0, fontSize: 24, fontWeight: 700, color: kpi.color, direction: "ltr", textAlign: "right" }}>{kpi.value}</p>
            </div>
          ))}
        </div>

        {/* Competitive landscape */}
        <div style={card({ padding: "18px 22px" })}>
          <p style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: T.text1 }}>چشم‌انداز رقبا</p>
          <p style={{ margin: "0 0 16px", fontSize: 12, color: T.text3 }}>مقایسه کل سشن‌های جذب‌شده در بازار تبلیغات دیجیتال</p>
          <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontSize: 12, color: T.text1, fontWeight: 600, minWidth: 90, textAlign: "right" }}>یکتانت ●</span>
              <div style={{ flex: 1, height: 24, borderRadius: 7, background: T.border, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${Math.round(totalYkt / compMax * 100)}%`, background: T.coral, borderRadius: 7, minWidth: 4 }} />
              </div>
              <span style={{ fontSize: 11, color: T.text2, minWidth: 60, textAlign: "left", direction: "ltr" }}>{formatNumber(totalYkt)}</span>
            </div>
            {compTotals.map((c, i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontSize: 12, color: T.text2, minWidth: 90, textAlign: "right" }}>{c.name}</span>
                <div style={{ flex: 1, height: 24, borderRadius: 7, background: T.border, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${Math.round(c.total / compMax * 100)}%`, background: AGENCY_COLORS[c.name] || T.text3, borderRadius: 7, minWidth: 2 }} />
                </div>
                <span style={{ fontSize: 11, color: T.text3, minWidth: 60, textAlign: "left", direction: "ltr" }}>{formatNumber(c.total)}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Industry + Team in 2 columns */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))", gap: 16 }}>
          <div style={card({ padding: "18px 22px" })}>
            <p style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: T.text1 }}>تفکیک صنعت</p>
            <p style={{ margin: "0 0 14px", fontSize: 12, color: T.text3 }}>میانگین سهم یکتانت به تفکیک صنعت</p>
            <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
              {indStats.filter(s => s.name !== "سایر").slice(0, 10).map((s, i) => {
                const c = s.yktShare >= 50 ? T.coral : s.yktShare >= 30 ? T.amber : T.danger;
                return (
                  <div key={i}>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4 }}>
                      <span style={{ fontSize: 12, color: T.text2 }}>{s.name} <span style={{ color: T.text3, fontSize: 11 }}>({s.count})</span></span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: c }}>{s.yktShare}٪</span>
                    </div>
                    <div style={{ height: 5, borderRadius: 3, background: T.border, overflow: "hidden" }}>
                      <div style={{ height: "100%", width: `${s.yktShare}%`, background: c, borderRadius: 3 }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {teamStats.length > 0 && (
            <div style={card({ padding: "18px 22px" })}>
              <p style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: T.text1 }}>عملکرد تیم‌ها</p>
              <p style={{ margin: "0 0 14px", fontSize: 12, color: T.text3 }}>سهم یکتانت و تعداد اکانت به تفکیک تیم</p>
              <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                {teamStats.map((t, i) => {
                  const c = t.yktShare >= 50 ? T.coral : t.yktShare >= 30 ? T.amber : T.danger;
                  return (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", borderRadius: 9, background: T.surface2 }}>
                      <div style={{ flex: 1 }}>
                        <p style={{ margin: "0 0 3px", fontSize: 12, color: T.text1, fontWeight: 500 }}>{t.name}</p>
                        <div style={{ height: 4, borderRadius: 2, background: T.border, overflow: "hidden" }}>
                          <div style={{ height: "100%", width: `${t.yktShare}%`, background: c, borderRadius: 2 }} />
                        </div>
                      </div>
                      <span style={{ fontSize: 11, color: T.text3, flexShrink: 0 }}>{t.accounts} اکانت</span>
                      <span style={{ fontSize: 13, fontWeight: 700, color: c, flexShrink: 0 }}>{t.yktShare}٪</span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* AM Leaderboard */}
        {amStats.length > 0 && (
          <div style={card({ padding: "18px 22px" })}>
            <p style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: T.text1 }}>عملکرد اکانت منیجرها</p>
            <p style={{ margin: "0 0 14px", fontSize: 12, color: T.text3 }}>سهم یکتانت، تعداد اکانت و حجم بودجه به تفکیک AM</p>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 8 }}>
              {amStats.slice(0, 12).map((am, i) => {
                const c = am.yktShare >= 50 ? T.coral : am.yktShare >= 30 ? T.amber : T.danger;
                return (
                  <div key={i} style={{ padding: "10px 14px", borderRadius: 10, background: T.surface2, border: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                    <div style={{ width: 32, height: 32, borderRadius: "50%", background: T.surface, border: `2px solid ${c}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, color: c, flexShrink: 0 }}>{am.name.slice(0, 1)}</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <p style={{ margin: "0 0 2px", fontSize: 12, color: T.text1, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{am.name}</p>
                      <p style={{ margin: 0, fontSize: 10, color: T.text3 }}>{am.accounts} اکانت · {formatNumber(am.sessions)} سشن</p>
                    </div>
                    <span style={{ fontSize: 15, fontWeight: 700, color: c, flexShrink: 0 }}>{am.yktShare}٪</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* At-Risk + Opportunities */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(380px, 1fr))", gap: 16 }}>
          {atRisk.length > 0 && (
            <div style={card({ padding: "18px 22px" })}>
              <p style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: T.amber }}>⚠ اکانت‌های در معرض خطر</p>
              <p style={{ margin: "0 0 14px", fontSize: 12, color: T.text3 }}>حجم سشن بالا + سهم یکتانت زیر ۳۵٪</p>
              <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                {atRisk.map((a, i) => (
                  <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", borderRadius: 9, background: T.amberDim, border: `1px solid ${T.amberBorder}` }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <p style={{ margin: "0 0 1px", fontSize: 12, color: T.text1, direction: "ltr", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</p>
                      <p style={{ margin: 0, fontSize: 10, color: T.text3 }}>{a.industry}{a.am ? ` · ${a.am}` : ""}</p>
                    </div>
                    <span style={{ fontSize: 11, color: T.text3, flexShrink: 0 }}>{formatNumber(a.sessions)}</span>
                    <span style={{ fontSize: 14, fontWeight: 700, color: T.amber, flexShrink: 0 }}>{a.yktShare}٪</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {opportunities.length > 0 && (
            <div style={card({ padding: "18px 22px" })}>
              <p style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: T.green }}>◎ فرصت‌های جدید</p>
              <p style={{ margin: "0 0 14px", fontSize: 12, color: T.text3 }}>سهم یکتانت صفر + حجم سشن بالا — احتمال لید</p>
              <div style={{ display: "flex", flexDirection: "column", gap: 7 }}>
                {opportunities.map((a, i) => (
                  <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 12px", borderRadius: 9, background: T.greenDim, border: `1px solid ${T.greenBorder}` }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <p style={{ margin: "0 0 1px", fontSize: 12, color: T.text1, direction: "ltr", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</p>
                      <p style={{ margin: 0, fontSize: 10, color: T.text3 }}>{a.industry}{a.am ? ` · ${a.am}` : ""}</p>
                    </div>
                    <span style={{ fontSize: 11, color: T.text3, flexShrink: 0 }}>{formatNumber(a.sessions)}</span>
                    <span style={{ fontSize: 11, color: T.green, fontWeight: 600, flexShrink: 0 }}>لید</span>
                  </div>
                ))}
              </div>
            </div>
          )}
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

  const globalStyles = `
    @keyframes radarSweep { to { transform: rotate(360deg); } }
    @keyframes radarPulse { 0%{transform:scale(1);opacity:0.6} 100%{transform:scale(3.5);opacity:0} }
    @keyframes dotPop { 0%,100%{opacity:0;transform:scale(0.5)} 30%,70%{opacity:0.9;transform:scale(1)} }
    @keyframes fadeUp { from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:translateY(0)} }
    @keyframes slideInModal { from{opacity:0;transform:scale(0.96) translateY(10px)} to{opacity:1;transform:scale(1) translateY(0)} }
    @keyframes spin { to{transform:rotate(360deg)} }
    .fade-up { animation: fadeUp 0.22s ease both; }
    * { box-sizing: border-box; }
    ::-webkit-scrollbar { width: 5px; height: 5px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: ${T.border}; border-radius: 10px; }
    select option { background: ${T.surface}; color: ${T.text1}; }
    html, body { overflow-y: auto; overflow-x: hidden; }
    button:focus-visible, a:focus-visible, input:focus-visible, select:focus-visible, textarea:focus-visible { outline: 2px solid ${T.coral}; outline-offset: 2px; border-radius: 4px; }
    [role="button"]:focus-visible { outline: 2px solid ${T.coral}; outline-offset: 2px; border-radius: 12px; }
  `;

  if (!session) return <><style>{globalStyles}</style><LoginPage T={T} onLogin={setSession} /></>;

  return (
    <div style={{ minHeight: "100dvh", background: T.bg, color: T.text1, direction: "rtl", fontFamily: "'Vazirmatn', system-ui, sans-serif", overflowX: "hidden" }}>
      <style>{globalStyles}</style>

      <Sidebar screen={screen} setScreen={setScreen} isDark={isDark} setIsDark={setIsDark} hasResult={!!result} hasCsv={!!csvData} session={session} onLogout={handleLogout} T={T} />

      <div style={{ marginRight: 80 }}>
        <div style={{ maxWidth: 1100, width: "100%", margin: "0 auto", padding: "40px 40px 100px" }}>
          {screen === "upload" && <UploadScreen />}
          {screen === "dashboard" && <DashboardScreen />}
          {screen === "results" && <ResultsScreen />}
          {screen === "reports" && <ReportsScreen />}
          {screen === "history" && <HistoryScreen />}
          {screen === "settings" && <SettingsScreen />}
        </div>
      </div>

      {modalAdv && (
        <DetailModal adv={modalAdv.adv} type={modalAdv.type} T={T} onClose={() => setModalAdv(null)}
          onRegen={() => regenOne(modalAdv.adv, modalAdv.type)} regenLoading={regenLoading === modalAdv.adv.name} />
      )}
      {toast && <Toast msg={toast.msg} type={toast.type} onDone={() => setToast(null)} T={T} />}
    </div>
  );
}
