import { useState, useCallback, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { ResponsiveContainer, LineChart, Line } from "recharts";

const MODEL = "anthropic/claude-sonnet-4-5";
const OPENROUTER_KEY = import.meta.env.VITE_OPENROUTER_KEY as string | undefined;
const HISTORY_KEY = "analyzer:reported_advertisers";
const REPORTS_KEY = "analyzer:saved_reports";

const C = {
  coral: "#D4623A", coralDim: "rgba(212,98,58,0.10)", coralBorder: "rgba(212,98,58,0.22)",
  sand: "#D3CFC7", stone: "#5F5E5A", mist: "#888780",
  white: "#FFFFFF", bg: "#F5F0E8",
  border: "#E8E3DA", border2: "#D3CFC7",
  t1: "#2C2C2A", t2: "#5F5E5A", t3: "#888780",
  green: "#1D9E75", greenDim: "rgba(29,158,117,0.10)", amber: "#854F0B", amberDim: "rgba(186,117,23,0.10)",
  blue: "#185FA5", blueDim: "rgba(55,138,221,0.10)",
  danger: "#DC2626", dangerBg: "#FEF2F2", dangerBorder: "#FECACA",
  successBg: "#F0FDF4", successBorder: "#BBF7D0", successText: "#15803D",
};

const AGENCY_COLORS: Record<string, string> = {
  "یکتانت": "#D4623A", "تپسل": "#3B82F6", "ادکسو": "#8B5CF6", "آپارات": "#EF4444",
  "یلو ادوایز": "#F59E0B", "بله": "#10B981", "روبیکا": "#EC4899", "دیما": "#6366F1",
  "طاووس": "#14B8A6", "دارت": "#F97316", "چاووش": "#84CC16", "تلوبیون": "#06B6D4",
  "ایتا": "#A78BFA", "سروش": "#FB923C", "نجوا": "#34D399", "بازار": "#64748B", "مایکت": "#78716C",
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
- date, owner_name (copy EXACTLY), manager_team, account_manager_name, total_sessions
- Agency columns: Yektanet, Tapsell, Adexo, Aparat, Daart, Deema, Tavoos, YellowAdwise, Bale_NonYektanet, Rubika_NonYektanet, Eitaa_NonYektanet, Soroush_NonYektanet, Telewebion, Chavosh, Najva, Triboon, Jaryan, Adverge, Bazaar, Myket
- OUR NETWORK: Yektanet only | COMPETITORS: all others | Empty=0

SPECIAL: همکده،5040،بیکوپلاس،owner_id 9868=SKIP. Aparat/Tavoos/Telewebion=VIDEO only. Aparat-owned=NOT leads.

PRIORITY: +5 reactivated | +4 very high sessions | +3 new/cut agency or Yektanet declining >20pts | +2 3+ agencies/high sessions | +1 dominant changed | -999 id 9868 | -3 همکده/5040/بیکوپلاس | -2 minor. Aim 8-12.

AGENCY TRANSLATIONS: Yektanet→یکتانت|Tapsell→تپسل|Adexo→ادکسو|Aparat→آپارات|Deema→دیما|Tavoos→طاووس|YellowAdwise→یلو ادوایز|Daart→دارت|Chavosh→چاووش|Telewebion→تلوبیون|Bale_NonYektanet→بله (غیریکتانتی)|Rubika_NonYektanet→روبیکا (غیریکتانتی)|Eitaa_NonYektanet→ایتا (غیریکتانتی)|Soroush_NonYektanet→سروش (غیریکتانتی)|Bazaar→بازار|Myket→مایکت

NUMBERS: "حدود ۳۰ هزار سشن" | shares nearest 5٪ | "یکتانت (۴۸٪)|تپسل (۱۸٪)" — NO session counts next to percentages. End each SUMMARY with: حجم کل دیروز: حدود X هزار سشن
DATE: Latest→دیروز|Before→پریروز|Earlier→اوایل هفته. Never English month names.
NO hashtags. NO کمپین. NO translated owner_name. NO English in output. Start with ##

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
AGENCIES: Bale_NonYektanet:1500
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

RULES: No hashtags. Max 2 sentences per advertiser. Only top 2 agencies. LEADS mandatory. No کمپین. No translated owner_name. No English. Start with ##`;

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
  const dates = [...new Set(rows.map(r => normalizeDate(r.date)).filter(Boolean))].sort();
  return { dates, lastDate: dates[dates.length - 1] || "", totalRows: rows.length, advertisers: [...new Set(rows.map(r => r.owner_name).filter(Boolean))].length };
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
  const TR: Record<string, string> = { Yektanet: "یکتانت", Tapsell: "تپسل", Adexo: "ادکسو", Aparat: "آپارات", Deema: "دیما", Tavoos: "طاووس", YellowAdwise: "یلو ادوایز", Daart: "دارت", Chavosh: "چاووش", Telewebion: "تلوبیون", Bale_NonYektanet: "بله", Rubika_NonYektanet: "روبیکا", Eitaa_NonYektanet: "ایتا", Soroush_NonYektanet: "سروش", Najva: "نجوا", Bazaar: "بازار", Myket: "مایکت" };
  return str.split(",").map(s => { const [k, v] = s.trim().split(":"); const name = TR[k?.trim()] || k?.trim() || ""; return { name, value: parseInt(v) || 0, color: AGENCY_COLORS[name] || "#999" }; }).filter(a => a.name && a.value > 0);
}

interface Advertiser { name: string; ownerid: string; manager: string; summary?: string; agencies: Agency[]; note?: string; }
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
    const bars = agencies.map(a => `<div style="flex:${a.value};background:${a.color}"></div>`).join("");
    const legend = agencies.map(a => `<span style="font-size:11px;color:#888;display:inline-flex;align-items:center;gap:4px;margin-left:8px"><span style="width:8px;height:8px;border-radius:50%;background:${a.color};display:inline-block"></span>${a.name} ${Math.round(a.value / total * 100)}٪</span>`).join("");
    return `<div style="display:flex;border-radius:4px;overflow:hidden;height:5px;margin:8px 0 4px">${bars}</div><div>${legend}</div>`;
  };

  const advCards = result.advertisers.map((adv) => {
    const total = adv.agencies.reduce((s, a) => s + a.value, 0);
    const ykt = adv.agencies.find(a => a.name === "یکتانت");
    const yPct = total > 0 && ykt ? Math.round(ykt.value / total * 100) : 0;
    const yColor = yPct >= 60 ? "#D4623A" : yPct >= 40 ? "#854F0B" : "#dc2626";
    return `<div style="display:flex;gap:12px;margin-bottom:10px">
      <div style="width:36px;height:36px;border-radius:9px;background:rgba(212,98,58,0.1);border:1.5px solid rgba(212,98,58,0.22);display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#D4623A" stroke-width="1.5"><path d="M20 7H4a2 2 0 00-2 2v6a2 2 0 002 2h16a2 2 0 002-2V9a2 2 0 00-2-2z"/></svg>
      </div>
      <div onclick="this.querySelector('.body').style.display=this.querySelector('.body').style.display==='block'?'none':'block'" style="flex:1;background:#fff;border:0.5px solid #D3CFC7;border-radius:12px;padding:12px 14px;cursor:pointer">
        <div style="display:flex;justify-content:space-between;align-items:start">
          <div>
            <div style="font-weight:500;font-size:13px;direction:ltr">${adv.name}</div>
            <div style="font-size:11px;color:#888;direction:ltr">#${adv.ownerid}${adv.manager ? ` · <span style="color:#D4623A">${adv.manager}</span>` : ""}</div>
            ${total > 0 ? `<div style="display:flex;align-items:center;gap:8px;margin-top:5px"><div style="flex:1;height:4px;border-radius:2px;background:#EDE8DF;overflow:hidden"><div style="width:${yPct}%;height:100%;background:#D4623A;border-radius:2px"></div></div><span style="font-size:11px;font-weight:500;color:${yColor}">یکتانت ${yPct}٪</span></div>` : ""}
          </div>
          <span style="font-size:11px;color:#888">▼</span>
        </div>
        <div class="body" style="display:none;margin-top:10px;padding-top:10px;border-top:0.5px solid #EDE8DF">
          <p style="font-size:13px;color:#5F5E5A;line-height:1.9;margin:0 0 8px">${adv.summary}</p>
          ${agBar(adv.agencies)}
        </div>
      </div>
    </div>`;
  }).join("");

  const leadCards = result.leads.map(l => {
    const total = l.agencies.reduce((s, a) => s + a.value, 0);
    return `<div style="display:flex;gap:12px;margin-bottom:10px">
      <div style="width:36px;height:36px;border-radius:9px;background:rgba(29,158,117,0.1);border:1.5px solid rgba(29,158,117,0.2);display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#1D9E75" stroke-width="1.5"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
      </div>
      <div onclick="this.querySelector('.body').style.display=this.querySelector('.body').style.display==='block'?'none':'block'" style="flex:1;background:#fff;border:0.5px solid rgba(29,158,117,0.2);border-radius:12px;padding:12px 14px;cursor:pointer">
        <div style="font-weight:500;font-size:13px;direction:ltr">${l.name} <span style="font-size:11px;color:#888;font-weight:400">#${l.ownerid}</span></div>
        ${l.manager ? `<div style="font-size:11px;color:#D4623A;margin-top:2px">${l.manager}</div>` : ""}
        <div class="body" style="display:none;margin-top:8px;padding-top:8px;border-top:0.5px solid #EDE8DF">
          <p style="font-size:13px;color:#5F5E5A;line-height:1.9;margin:0 0 8px">${l.note}</p>
          ${agBar(l.agencies)}
        </div>
      </div>
    </div>`;
  }).join("");

  const compCards = result.competitors.map(c => `
    <div style="display:flex;gap:12px;margin-bottom:10px">
      <div style="width:36px;height:36px;border-radius:9px;background:rgba(186,117,23,0.1);border:1.5px solid rgba(186,117,23,0.2);display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#854F0B" stroke-width="1.5"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
      </div>
      <div style="flex:1;background:#fff;border:0.5px solid #D3CFC7;border-radius:12px;padding:12px 14px">
        <div style="font-weight:500;font-size:13px;color:#854F0B;margin-bottom:4px">${c.platform}</div>
        <p style="font-size:13px;color:#5F5E5A;line-height:1.8;margin:0 0 4px">${c.note}</p>
        ${c.topclients ? `<div style="font-size:12px;color:#888">مهم‌ترین: ${c.topclients}</div>` : ""}
        ${c.newclients && c.newclients !== "ندارد" ? `<div style="font-size:12px;color:#1D9E75">جدید: ${c.newclients}</div>` : ""}
      </div>
    </div>`).join("");

  return `<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>گزارش بازار — ${result.dateLabel}</title>
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:system-ui,sans-serif;background:#F5F0E8;color:#2C2C2A;direction:rtl;padding:1.5rem 1rem}.container{max-width:720px;margin:0 auto}</style>
</head>
<body>
<div class="container">
  <div style="display:flex;align-items:center;gap:14px;margin-bottom:1.5rem;padding-bottom:1.5rem;border-bottom:1px solid #D3CFC7">
    <div style="width:44px;height:44px;border-radius:12px;background:#D4623A;display:flex;align-items:center;justify-content:center;flex-shrink:0">
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="7" height="7" rx="1" fill="#fff"/><rect x="14" y="3" width="7" height="7" rx="1" fill="#fff"/><rect x="3" y="14" width="7" height="7" rx="1" fill="#fff"/><rect x="14" y="14" width="7" height="7" rx="1" fill="#fff"/></svg>
    </div>
    <div>
      <div style="font-size:19px;font-weight:500;letter-spacing:-0.3px">Market Analyzer</div>
      <div style="font-size:12px;color:#888">از مارکت چه خبر؟ ${result.dateLabel}</div>
    </div>
  </div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:20px">
    ${[["تاریخ", result.dateLabel, "#2C2C2A"], ["تبلیغ‌کننده", String(result.advertisers.length), "#D4623A"], ["لید", String(result.leads.length), "#1D9E75"], ["رقیب", String(result.competitors.length), "#854F0B"]].map(([l, v, c]) => `
    <div style="background:#EDE8DF;border-radius:10px;padding:11px 13px">
      <div style="font-size:11px;color:#888;margin-bottom:3px">${l}</div>
      <div style="font-size:19px;font-weight:500;color:${c}">${v}</div>
    </div>`).join("")}
  </div>

  ${result.advertisers.length > 0 ? `<div style="font-size:11px;font-weight:500;color:#888;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:10px;padding-right:48px">تبلیغ‌کننده‌های مهم</div>${advCards}` : ""}
  ${result.leads.length > 0 ? `<div style="font-size:11px;font-weight:500;color:#888;text-transform:uppercase;letter-spacing:0.08em;margin:16px 0 10px;padding-right:48px">فرصت‌های اپروچ</div>${leadCards}` : ""}
  ${result.competitors.length > 0 ? `<div style="font-size:11px;font-weight:500;color:#888;text-transform:uppercase;letter-spacing:0.08em;margin:16px 0 10px;padding-right:48px">پلتفرم‌های رقیب</div>${compCards}` : ""}
  ${result.market ? `<div style="display:flex;gap:12px;margin-top:16px"><div style="width:36px;height:36px;border-radius:9px;background:#EDE8DF;border:1.5px solid #D3CFC7;display:flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:2px"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#888" stroke-width="1.5"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="12" y1="17" x2="12" y2="21"/></svg></div><div style="flex:1;background:#EDE8DF;border-radius:12px;padding:12px 14px"><div style="font-size:11px;color:#888;font-weight:500;text-transform:uppercase;letter-spacing:0.06em;margin-bottom:6px">تحلیل کلی بازار</div><p style="font-size:13px;color:#5F5E5A;line-height:1.9">${result.market}</p></div></div>` : ""}

  <div style="font-size:11px;color:#888;font-style:italic;margin-top:2rem;padding-top:1rem;border-top:1px solid #D3CFC7;text-align:center">${result.disclaimer}</div>
  <div style="font-size:11px;color:#888;text-align:center;margin-top:6px">📍 Powered by Claude</div>
</div>
</body>
</html>`;
}

function downloadHTML(result: AnalysisResult) {
  const html = generateHTML(result);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `market-report-${result.dateLabel.replace(/ /g, "-")}.html`; a.click();
  URL.revokeObjectURL(url);
}

function AgencyBar({ agencies }: { agencies: Agency[] }) {
  if (!agencies || agencies.length === 0) return null;
  const total = agencies.reduce((s, a) => s + a.value, 0);
  return (
    <div style={{ marginTop: 8 }}>
      <div style={{ display: "flex", borderRadius: 5, overflow: "hidden", height: 5 }}>
        {agencies.map((a, i) => <div key={i} title={`${a.name}: ${Math.round(a.value / total * 100)}٪`} style={{ flex: a.value, background: a.color }} />)}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "3px 10px", marginTop: 5 }}>
        {agencies.map((a, i) => (
          <span key={i} style={{ fontSize: 11, color: C.t3, display: "flex", alignItems: "center", gap: 4 }}>
            <span style={{ width: 7, height: 7, borderRadius: "50%", background: a.color, display: "inline-block", flexShrink: 0 }} />
            {a.name} {Math.round(a.value / total * 100)}٪
          </span>
        ))}
      </div>
    </div>
  );
}

export default function App() {
  const [page, setPage] = useState("main");
  const [step, setStep] = useState("upload");
  const [csvData, setCsvData] = useState<{ text: string; rows: Record<string, string>[]; name: string; stats: ReturnType<typeof getStats> } | null>(null);
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [pendingSave, setPendingSave] = useState<{ date: string; dateLabel: string; names: string[] } | null>(null);
  const [saveStatus, setSaveStatus] = useState("");
  const [editableMsg, setEditableMsg] = useState("");
  const [loading, setLoading] = useState(false);
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
  const [expandedCard, setExpandedCard] = useState<string | null>(null);
  const [compareReport, setCompareReport] = useState<AnalysisResult | null>(null);
  const [compareMode, setCompareMode] = useState(false);

  useEffect(() => { loadHistory(); loadReports(); }, []);

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
      setPendingSave(null); setSaveStatus("saved"); setTimeout(() => setSaveStatus(""), 3000);
    } catch { setSaveStatus("error"); }
  };
  const deleteEntry = (date: string) => { try { persistHistory(history.filter(h => h.date !== date)); } catch { } };
  const clearHistory = () => { try { persistHistory([]); } catch { } };
  const uniqueReported = [...new Set(history.flatMap(h => h.names))];

  const handleFile = useCallback((file: File) => {
    if (!file) return;
    readFile(file, (rows, csvText) => { setCsvData({ text: csvText, rows, name: file.name, stats: getStats(rows) }); setStep("ready"); setErrMsg(""); }, () => setErrMsg("خطا در خواندن فایل."));
  }, []);
  const onDrop = useCallback((e: React.DragEvent) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }, [handleFile]);

  const buildManagerMap = (rows: Record<string, string>[]) => {
    const mCol = Object.keys(rows[0] || {}).find(k => k.trim().toLowerCase().replace(/[_ -]/g, "") === "accountmanagername") || Object.keys(rows[0] || {}).find(k => k.trim().toLowerCase().replace(/[_ -]/g, "") === "managerteam") || "";
    const map: Record<string, string> = {};
    rows.forEach(r => { const n = (r.owner_name || "").toString().trim(); const m = mCol ? (r[mCol] || "").toString().trim() : ""; if (!n || !m) return; map[n] = m; const p = n.split(" - "); if (p.length >= 2) { map[p.slice(1).join(" - ").trim()] = m; map[p[0].trim()] = m; } });
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
        try { const chunk = JSON.parse(payload); const delta = chunk.choices?.[0]?.delta?.content; if (delta) raw += delta; } catch { /* partial chunk */ }
      }
    }
    setRawDebug(raw);
    if (!raw || raw.length < 5) throw new Error("پاسخ خالی");
    const advertisers = extractBlocks("ADVERTISER", raw).map(b => { const f = parseFields(b); const n = f.NAME || ""; return { name: n, ownerid: f.OWNERID || "", manager: managerMap[n] || "", summary: f.SUMMARY || "", agencies: parseAgencies(f.AGENCIES) }; });
    const leads = extractBlocks("LEAD", raw).map(b => { const f = parseFields(b); const n = f.NAME || ""; return { name: n, ownerid: f.OWNERID || "", manager: managerMap[n] || "", note: f.NOTE || "", agencies: parseAgencies(f.AGENCIES) }; });
    const competitors = extractBlocks("COMPETITOR", raw).map(b => { const f = parseFields(b); return { platform: f.PLATFORM || "", newclients: f.NEWCLIENTS || "", topclients: f.TOPCLIENTS || "", note: f.NOTE || "" }; });
    const market = (extractBlocks("MARKET", raw)[0] || "").trim();
    const disclaimer = (extractBlocks("DISCLAIMER", raw)[0] || "داده‌ها بر اساس کراول وب هستن و تخمینی‌اند — روندها قابل اعتمادند اما اعداد دقیق نیستن.").trim();
    return { dateLabel: todayLabel(), advertisers, leads, competitors, market, disclaimer };
  };

  const analyze = async () => {
    setLoading(true); setErrMsg(""); setRawDebug(""); setShowDebug(false);
    try {
      const parsed = await runAnalysis(csvData!.text, csvData!.rows, selectedTemplate);
      if (!parsed.advertisers.length && !parsed.leads.length && !parsed.competitors.length) { setErrMsg("هیچ بلاکی پارس نشد"); setShowDebug(true); setLoading(false); return; }
      setResult(parsed); setEditableMsg(buildMessage(parsed));
      setPendingSave({ date: csvData!.stats.lastDate, dateLabel: parsed.dateLabel, names: [...parsed.advertisers.map(a => a.name), ...parsed.leads.map(l => l.name)] });
      setSaveStatus(""); setStep("preview"); setSearchQuery(""); setFilterAgency("all");
    } catch (e) { setErrMsg("خطا: " + (e as Error).message); }
    finally { setLoading(false); }
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
        setResult(nr); setEditableMsg(buildMessage(nr));
      }
    } catch (e) { console.error(e); }
    finally { setRegenLoading(null); }
  };

  const reset = () => { setCsvData(null); setResult(null); setPendingSave(null); setStep("upload"); setErrMsg(""); setRawDebug(""); setShowDebug(false); setSaveStatus(""); setExpandedCard(null); setSearchQuery(""); setFilterAgency("all"); };

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

  const navA: React.CSSProperties = { fontSize: 13, fontWeight: 500, padding: "8px 14px", borderRadius: 8, border: "none", cursor: "pointer", background: C.coral, color: "#fff" };
  const navI: React.CSSProperties = { fontSize: 13, padding: "8px 14px", borderRadius: 8, border: "none", cursor: "pointer", background: "transparent", color: C.t2 };
  const btn = (x: React.CSSProperties = {}): React.CSSProperties => ({ padding: "8px 14px", borderRadius: 10, border: `0.5px solid ${C.border2}`, background: "transparent", color: C.t1, fontSize: 12, cursor: "pointer", ...x });

  const AdvCard = ({ adv, type = "advertiser" }: { adv: Advertiser; type?: string }) => {
    const isExp = expandedCard === adv.name, isRegen = regenLoading === adv.name;
    const total = adv.agencies.reduce((s, a) => s + a.value, 0);
    const ykt = adv.agencies.find(a => a.name === "یکتانت");
    const yPct = total > 0 && ykt ? Math.round(ykt.value / total * 100) : 0;
    const yColor = yPct >= 60 ? C.coral : yPct >= 40 ? C.amber : C.danger;
    const cmpAdv = compareReport?.[type === "lead" ? "leads" : "advertisers"]?.find(a => a.name === adv.name);
    const cmpPct = cmpAdv ? (() => { const ct = cmpAdv.agencies.reduce((s, a) => s + a.value, 0); const cy = cmpAdv.agencies.find(a => a.name === "یکتانت"); return ct > 0 && cy ? Math.round(cy.value / ct * 100) : 0; })() : null;
    return (
      <div style={{ display: "flex", gap: 14, marginBottom: 10 }}>
        <div style={{ flexShrink: 0, zIndex: 1, marginTop: 2 }}>
          <div style={{ width: 36, height: 36, borderRadius: 9, background: type === "lead" ? C.greenDim : C.coralDim, border: `1.5px solid ${type === "lead" ? "rgba(29,158,117,0.2)" : C.coralBorder}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
            {type === "lead"
              ? <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.green} strokeWidth="1.5"><circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="16" /><line x1="8" y1="12" x2="16" y2="12" /></svg>
              : <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.coral} strokeWidth="1.5"><path d="M20 7H4a2 2 0 00-2 2v6a2 2 0 002 2h16a2 2 0 002-2V9a2 2 0 00-2-2z" /></svg>}
          </div>
        </div>
        <div style={{ flex: 1, background: C.white, border: `0.5px solid ${isExp ? C.coralBorder : C.border}`, borderRadius: 13, padding: "0.9rem 1.1rem", cursor: "pointer", transition: "border-color 0.2s" }} onClick={() => setExpandedCard(isExp ? null : adv.name)}>
          <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" }}>
                <span style={{ fontWeight: 500, fontSize: 13, color: C.t1, direction: "ltr" }}>{adv.name}</span>
                <span style={{ fontSize: 11, color: C.t3, direction: "ltr" }}>#{adv.ownerid}</span>
                {adv.manager && <span style={{ fontSize: 11, color: C.coral, background: C.coralDim, padding: "1px 6px", borderRadius: 20 }}>{adv.manager}</span>}
              </div>
              {total > 0 && <div style={{ display: "flex", alignItems: "center", gap: 7, marginTop: 5 }}>
                <div style={{ flex: 1, height: 4, borderRadius: 2, background: C.bg, overflow: "hidden" }}>
                  <div style={{ width: `${yPct}%`, height: "100%", background: C.coral, borderRadius: 2, transition: "width 0.4s" }} />
                </div>
                <span style={{ fontSize: 11, fontWeight: 500, color: yColor, flexShrink: 0 }}>
                  یکتانت {yPct}٪
                  {cmpPct !== null && <span style={{ fontSize: 10, color: yPct > cmpPct ? C.green : C.danger, marginRight: 3 }}>{yPct > cmpPct ? `▲+${yPct - cmpPct}` : `▼${yPct - cmpPct}`}٪</span>}
                </span>
              </div>}
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 5, flexShrink: 0 }}>
              <button onClick={e => { e.stopPropagation(); regenOne(adv, type); }} disabled={!!regenLoading} style={{ padding: "3px 7px", borderRadius: 6, border: `0.5px solid ${C.border}`, background: "transparent", cursor: "pointer", color: C.t3, fontSize: 11, display: "flex", alignItems: "center" }}>
                {isRegen ? <svg width="11" height="11" viewBox="0 0 24 24" fill="none" style={{ animation: "spin 0.8s linear infinite" }}><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" strokeDasharray="31" strokeDashoffset="10" /></svg> : <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10" /></svg>}
              </button>
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke={C.mist} strokeWidth="2" style={{ transform: isExp ? "rotate(180deg)" : "rotate(0)", transition: "transform 0.2s" }}><polyline points="6 9 12 15 18 9" /></svg>
            </div>
          </div>
          {isExp && <div style={{ marginTop: 9, borderTop: `0.5px solid ${C.border}`, paddingTop: 9 }}>
            <p style={{ margin: "0 0 8px", fontSize: 13, color: C.t2, lineHeight: 1.9 }}>{type === "lead" ? adv.note : adv.summary}</p>
            <AgencyBar agencies={adv.agencies} />
            {cmpAdv && <div style={{ marginTop: 10, padding: "8px 10px", background: C.bg, borderRadius: 8 }}>
              <p style={{ margin: "0 0 4px", fontSize: 11, color: C.t3 }}>مقایسه با گزارش قبلی:</p>
              <AgencyBar agencies={cmpAdv.agencies} />
            </div>}
          </div>}
        </div>
      </div>
    );
  };

  // suppress unused import warning — recharts components available for future use
  void ResponsiveContainer; void LineChart; void Line;

  return (
    <div style={{ maxWidth: 760, margin: "0 auto", padding: "2rem 0", direction: "rtl", fontFamily: "system-ui, sans-serif" }}>
      <style>{`@keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}} @keyframes spin{to{transform:rotate(360deg)}} .fe{animation:fadeUp 0.25s ease both}`}</style>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: "2rem", paddingBottom: "1.5rem", borderBottom: `0.5px solid ${C.border}` }}>
        <div style={{ width: 44, height: 44, borderRadius: 12, background: C.coral, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><rect x="3" y="3" width="7" height="7" rx="1" fill="#fff" /><rect x="14" y="3" width="7" height="7" rx="1" fill="#fff" /><rect x="3" y="14" width="7" height="7" rx="1" fill="#fff" /><rect x="14" y="14" width="7" height="7" rx="1" fill="#fff" /></svg>
        </div>
        <div style={{ flex: 1 }}>
          <h2 style={{ margin: 0, fontSize: 19, fontWeight: 500, color: C.t1, letterSpacing: "-0.3px" }}>Market Analyzer</h2>
          <p style={{ margin: 0, fontSize: 12, color: C.t3 }}>آنالیز رقابتی روزانه — یکتانت</p>
        </div>
        <nav style={{ display: "flex", gap: 3, background: C.bg, borderRadius: 10, padding: 3 }}>
          {([["main", "آنالیز"], ["reports", "گزارش‌ها"], ["history", "تاریخچه"]] as [string, string][]).map(([p, l]) => (
            <button key={p} onClick={() => setPage(p)} style={page === p ? navA : navI}>{l}</button>
          ))}
        </nav>
      </div>

      {/* HISTORY */}
      {page === "history" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }} className="fe">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <p style={{ margin: 0, fontSize: 15, fontWeight: 500, color: C.t1 }}>تبلیغ‌کننده‌های گزارش‌شده</p>
            {history.length > 0 && <button onClick={clearHistory} style={{ ...btn(), color: C.danger, border: `0.5px solid ${C.dangerBorder}` }}>پاک کردن همه</button>}
          </div>
          {history.length === 0
            ? <div style={{ textAlign: "center", padding: "3rem", background: C.bg, borderRadius: 16 }}><p style={{ margin: 0, fontSize: 13, color: C.t3 }}>هنوز گزارشی ذخیره نشده.</p></div>
            : [...history].reverse().map((h, i) => (
              <div key={i} style={{ background: C.white, border: `0.5px solid ${C.border}`, borderRadius: 14, padding: "1rem 1.25rem" }} className="fe">
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <div style={{ width: 8, height: 8, borderRadius: "50%", background: C.coral }} />
                    <span style={{ fontSize: 13, fontWeight: 500, color: C.t1 }}>{h.dateLabel || h.date}</span>
                    <span style={{ fontSize: 11, color: C.t3, background: C.bg, padding: "2px 8px", borderRadius: 20 }}>{h.names.length} تبلیغ‌کننده</span>
                  </div>
                  <button onClick={() => deleteEntry(h.date)} style={{ ...btn(), color: C.danger, border: `0.5px solid ${C.dangerBorder}`, fontSize: 11 }}>حذف</button>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {h.names.map((n, j) => <span key={j} style={{ fontSize: 11, color: C.coral, background: C.coralDim, border: `0.5px solid ${C.coralBorder}`, padding: "3px 10px", borderRadius: 20, direction: "ltr" }}>{n}</span>)}
                </div>
              </div>
            ))}
        </div>
      )}

      {/* REPORTS */}
      {page === "reports" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }} className="fe">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <p style={{ margin: 0, fontSize: 15, fontWeight: 500, color: C.t1 }}>گزارش‌های ذخیره‌شده</p>
            <div style={{ display: "flex", gap: 8 }}>
              {compareMode && <button onClick={() => { setCompareMode(false); setCompareReport(null); }} style={{ ...btn(), background: C.coralDim, color: C.coral, border: `0.5px solid ${C.coralBorder}` }}>خروج از مقایسه</button>}
              {!compareMode && savedReports.length > 1 && <button onClick={() => setCompareMode(true)} style={btn()}>مقایسه گزارش‌ها</button>}
            </div>
          </div>
          {savedReports.length === 0
            ? <div style={{ textAlign: "center", padding: "3rem", background: C.bg, borderRadius: 16 }}><p style={{ margin: 0, fontSize: 13, color: C.t3 }}>هنوز گزارشی ذخیره نشده.</p></div>
            : [...savedReports].reverse().map((r, i) => (
              <div key={i} style={{ background: C.white, border: `0.5px solid ${compareReport?.id === r.id ? C.coralBorder : C.border}`, borderRadius: 14, padding: "1rem 1.25rem" }} className="fe">
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                  <div>
                    <p style={{ margin: 0, fontSize: 14, fontWeight: 500, color: C.t1 }}>{r.dateLabel}</p>
                    <p style={{ margin: "2px 0 0", fontSize: 12, color: C.t3 }}>{r.advertisers?.length || 0} تبلیغ‌کننده · {r.leads?.length || 0} لید</p>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    {compareMode && <button onClick={() => setCompareReport(compareReport?.id === r.id ? null : r)} style={{ ...btn(), background: compareReport?.id === r.id ? C.coralDim : "transparent", color: compareReport?.id === r.id ? C.coral : C.t2 }}>{compareReport?.id === r.id ? "انتخاب شده" : "انتخاب"}</button>}
                    <button onClick={() => { setResult(r); setEditableMsg(buildMessage(r)); setPage("main"); setStep("preview"); }} style={btn()}>مشاهده</button>
                  </div>
                </div>
                {r.advertisers?.slice(0, 3).map((a, j) => {
                  const total = a.agencies?.reduce((s, x) => s + x.value, 0) || 1;
                  const ykt = a.agencies?.find(x => x.name === "یکتانت");
                  return (<div key={j} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                    <span style={{ fontSize: 12, color: C.t2, direction: "ltr", minWidth: 110, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.name}</span>
                    <div style={{ flex: 1, height: 4, borderRadius: 2, background: C.bg, overflow: "hidden" }}>
                      <div style={{ width: `${Math.round((ykt?.value || 0) / total * 100)}%`, height: "100%", background: C.coral, borderRadius: 2 }} />
                    </div>
                  </div>);
                })}
              </div>
            ))}
        </div>
      )}

      {/* MAIN */}
      {page === "main" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>

          {/* Steps */}
          <div style={{ display: "flex", alignItems: "center" }}>
            {(["آپلود", "آنالیز", "پیام"] as string[]).map((label, i) => {
              const idx = ["upload", "ready", "preview"].indexOf(step), done = idx > i, active = idx === i;
              return (<div key={i} style={{ display: "flex", alignItems: "center", flex: 1 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 7, flexShrink: 0 }}>
                  <div style={{ width: 26, height: 26, borderRadius: "50%", background: done || active ? C.coral : C.bg, border: `1.5px solid ${done || active ? C.coral : C.border}`, display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.3s" }}>
                    {done ? <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="3"><polyline points="20 6 9 17 4 12" /></svg> : <span style={{ fontSize: 11, fontWeight: 500, color: done || active ? "#fff" : C.mist }}>{i + 1}</span>}
                  </div>
                  <span style={{ fontSize: 12, fontWeight: active ? 500 : 400, color: active ? C.t1 : C.t3 }}>{label}</span>
                </div>
                {i < 2 && <div style={{ flex: 1, height: "1px", background: done ? C.coral : C.border, margin: "0 10px", transition: "background 0.4s" }} />}
              </div>);
            })}
          </div>

          {uniqueReported.length > 0 && step !== "preview" && (
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 13px", background: C.coralDim, border: `0.5px solid ${C.coralBorder}`, borderRadius: 10 }}>
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={C.coral} strokeWidth="2"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" /><line x1="12" y1="9" x2="12" y2="13" /><line x1="12" y1="17" x2="12.01" y2="17" /></svg>
              <p style={{ margin: 0, fontSize: 12, color: C.coral }}>{uniqueReported.length} تبلیغ‌کننده از ۵ روز اخیر skip می‌شن</p>
            </div>
          )}

          {/* UPLOAD */}
          {step === "upload" && (
            <div className="fe" onDrop={onDrop} onDragOver={e => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onClick={() => document.getElementById("fi")!.click()}
              style={{ border: `1.5px dashed ${dragOver ? C.coral : C.sand}`, borderRadius: 20, padding: "3rem 2rem", textAlign: "center", cursor: "pointer", background: dragOver ? C.coralDim : C.bg, transition: "all 0.2s" }}>
              <div style={{ width: 60, height: 60, borderRadius: 16, background: dragOver ? C.coral : "#fff", border: `1.5px solid ${dragOver ? C.coral : C.sand}`, margin: "0 auto 16px", display: "flex", alignItems: "center", justifyContent: "center", transition: "all 0.2s" }}>
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke={dragOver ? "#fff" : C.coral} strokeWidth="1.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>
              </div>
              <p style={{ margin: "0 0 6px", fontWeight: 500, fontSize: 16, color: C.t1 }}>فایل داده‌های رقابتی را آپلود کنید</p>
              <p style={{ margin: "0 0 10px", fontSize: 13, color: C.t3 }}>CSV یا XLSX · بکشید یا کلیک کنید</p>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 5, justifyContent: "center" }}>
                {["date", "owner_name", "account_manager_name", "Yektanet", "Tapsell", "…"].map((c, i) => (
                  <span key={i} style={{ fontSize: 11, color: C.stone, background: "#fff", border: `0.5px solid ${C.sand}`, padding: "3px 9px", borderRadius: 20, direction: "ltr" }}>{c}</span>
                ))}
              </div>
              <input id="fi" type="file" accept=".csv,.xlsx,.xls" style={{ display: "none" }} onChange={e => handleFile(e.target.files![0])} />
            </div>
          )}

          {/* READY */}
          {step === "ready" && csvData && (
            <div className="fe" style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ background: C.white, border: `0.5px solid ${C.border}`, borderRadius: 16, padding: "1.25rem" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14 }}>
                  <div style={{ width: 40, height: 40, borderRadius: 10, background: C.coralDim, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke={C.coral} strokeWidth="1.5"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" /><polyline points="14 2 14 8 20 8" /></svg>
                  </div>
                  <div style={{ flex: 1 }}><p style={{ margin: 0, fontWeight: 500, fontSize: 14, color: C.t1 }}>{csvData.name}</p><p style={{ margin: "2px 0 0", fontSize: 12, color: C.t3 }}>{csvData.rows.length.toLocaleString()} ردیف · {csvData.stats.dates.length} روز</p></div>
                  <button onClick={() => { setCsvData(null); setStep("upload"); }} style={btn()}>تغییر</button>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
                  {([{ l: "تبلیغ‌کننده", v: csvData.stats.advertisers, i: 0 }, { l: "روزها", v: csvData.stats.dates.length, i: 1 }, { l: "گزارش برای", v: todayLabel(), i: 2 }]).map((c) => (
                    <div key={c.i} style={{ background: C.bg, borderRadius: 10, padding: "9px 11px" }}><p style={{ margin: "0 0 2px", fontSize: 11, color: C.t3 }}>{c.l}</p><p style={{ margin: 0, fontSize: 13, fontWeight: 500, color: c.i === 2 ? C.coral : C.t1, direction: "ltr", textAlign: "right" }}>{c.v}</p></div>
                  ))}
                </div>
              </div>

              {/* Template */}
              <div style={{ background: C.white, border: `0.5px solid ${C.border}`, borderRadius: 13, padding: "1rem 1.25rem" }}>
                <p style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 500, color: C.t3, textTransform: "uppercase", letterSpacing: "0.06em" }}>قالب گزارش</p>
                <div style={{ display: "flex", gap: 8 }}>
                  {Object.entries(TEMPLATES).map(([key, t]) => (
                    <button key={key} onClick={() => setSelectedTemplate(key)} style={{ flex: 1, padding: "9px 8px", borderRadius: 10, border: `1.5px solid ${selectedTemplate === key ? C.coral : C.border}`, background: selectedTemplate === key ? C.coralDim : C.bg, cursor: "pointer", transition: "all 0.2s" }}>
                      <p style={{ margin: "0 0 2px", fontSize: 13, fontWeight: selectedTemplate === key ? 500 : 400, color: selectedTemplate === key ? C.coral : C.t1 }}>{t.label}</p>
                      <p style={{ margin: 0, fontSize: 11, color: C.t3 }}>{t.desc}</p>
                    </button>
                  ))}
                </div>
              </div>

              <button onClick={analyze} disabled={loading} style={{ padding: "13px", borderRadius: 13, border: "none", background: loading ? C.bg : C.coral, color: loading ? C.t3 : "#fff", fontSize: 15, fontWeight: 500, cursor: loading ? "not-allowed" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 10, transition: "all 0.2s" }}>
                {loading ? <><svg width="16" height="16" viewBox="0 0 24 24" fill="none" style={{ animation: "spin 0.8s linear infinite" }}><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="2" strokeDasharray="31" strokeDashoffset="10" /></svg>در حال آنالیز...</> : <>شروع آنالیز <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="5" y1="12" x2="19" y2="12" /><polyline points="12 5 19 12 12 19" /></svg></>}
              </button>
              {errMsg && <div style={{ padding: "10px 14px", borderRadius: 10, background: C.dangerBg, border: `0.5px solid ${C.dangerBorder}` }}><p style={{ margin: 0, fontSize: 13, color: C.danger }}>{errMsg}</p></div>}
              {showDebug && rawDebug && <pre style={{ margin: 0, fontSize: 11, color: C.t1, background: C.bg, padding: 12, borderRadius: 10, whiteSpace: "pre-wrap", direction: "ltr", textAlign: "left", maxHeight: 200, overflow: "auto" }}>{rawDebug}</pre>}
            </div>
          )}

          {/* PREVIEW */}
          {step === "preview" && result && (
            <div style={{ display: "flex", flexDirection: "column", gap: 0 }} className="fe">

              {/* Export bar */}
              <div style={{ display: "flex", gap: 8, marginBottom: 14, padding: "10px 14px", background: C.bg, borderRadius: 12, border: `0.5px solid ${C.border}`, alignItems: "center", flexWrap: "wrap" }}>
                <span style={{ fontSize: 12, color: C.t3, flex: 1 }}>خروجی:</span>
                <button onClick={() => navigator.clipboard.writeText(editableMsg)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 16px", borderRadius: 10, border: "none", background: C.coral, color: "#fff", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="9" y="9" width="13" height="13" rx="2" /><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1" /></svg>کپی متن
                </button>
                <button onClick={() => downloadHTML(result)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "8px 16px", borderRadius: 10, border: `0.5px solid ${C.coralBorder}`, background: C.coralDim, color: C.coral, fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" y1="15" x2="12" y2="3" /></svg>دانلود HTML
                </button>
              </div>

              {/* Save */}
              {pendingSave && saveStatus !== "saved" && (
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "11px 15px", background: C.coralDim, border: `1px solid ${C.coralBorder}`, borderRadius: 13, marginBottom: 14 }}>
                  <div><p style={{ margin: "0 0 1px", fontSize: 13, fontWeight: 500, color: C.t1 }}>این نسخه ذخیره نشده</p><p style={{ margin: 0, fontSize: 12, color: C.t3 }}>{pendingSave.dateLabel} · {pendingSave.names.length} تبلیغ‌کننده</p></div>
                  <button onClick={confirmSave} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 9, border: "none", background: C.coral, color: "#fff", fontSize: 13, fontWeight: 500, cursor: "pointer" }}>
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2z" /><polyline points="17 21 17 13 7 13 7 21" /></svg>ذخیره
                  </button>
                </div>
              )}
              {saveStatus === "saved" && <div style={{ padding: "9px 14px", background: C.successBg, border: `0.5px solid ${C.successBorder}`, borderRadius: 9, marginBottom: 14 }}><p style={{ margin: 0, fontSize: 13, color: C.successText }}>✓ ذخیره شد</p></div>}

              {/* Summary */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginBottom: 18 }}>
                {([{ l: "تاریخ", v: result.dateLabel, c: C.t1 }, { l: "تبلیغ‌کننده", v: result.advertisers.length, c: C.coral }, { l: "لید", v: result.leads.length, c: C.green }, { l: "رقیب", v: result.competitors.length, c: C.amber }]).map((c, i) => (
                  <div key={i} style={{ background: C.bg, borderRadius: 10, padding: "10px 12px" }}><p style={{ margin: "0 0 3px", fontSize: 11, color: C.t3 }}>{c.l}</p><p style={{ margin: 0, fontSize: 18, fontWeight: 500, color: c.c }}>{c.v}</p></div>
                ))}
              </div>

              {/* Compare banner */}
              {compareReport && <div style={{ padding: "8px 13px", background: C.blueDim, border: "0.5px solid rgba(55,138,221,0.25)", borderRadius: 9, marginBottom: 12, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <p style={{ margin: 0, fontSize: 12, color: C.blue }}>مقایسه با گزارش {compareReport.dateLabel} فعاله</p>
                <button onClick={() => setCompareReport(null)} style={{ ...btn(), fontSize: 11, padding: "3px 9px", color: C.blue }}>حذف</button>
              </div>}

              {/* Filters */}
              <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
                <input value={searchQuery} onChange={e => setSearchQuery(e.target.value)} placeholder="جستجوی تبلیغ‌کننده..." style={{ flex: 1, minWidth: 130, fontSize: 13, padding: "7px 11px", borderRadius: 9, border: `0.5px solid ${C.border}`, background: C.bg, color: C.t1, direction: "rtl" }} />
                <select value={filterAgency} onChange={e => setFilterAgency(e.target.value)} style={{ fontSize: 12, padding: "7px 9px", borderRadius: 9, border: `0.5px solid ${C.border}`, background: C.bg, color: C.t1 }}>
                  <option value="all">همه آژانس‌ها</option>
                  {allAgencies.map(a => <option key={a} value={a}>{a}</option>)}
                </select>
                <select value={sortBy} onChange={e => setSortBy(e.target.value)} style={{ fontSize: 12, padding: "7px 9px", borderRadius: 9, border: `0.5px solid ${C.border}`, background: C.bg, color: C.t1 }}>
                  <option value="importance">اهمیت</option>
                  <option value="volume">حجم</option>
                  <option value="yektanet">سهم یکتانت ↑</option>
                </select>
              </div>

              {/* Timeline */}
              <div style={{ position: "relative" }}>
                <div style={{ position: "absolute", right: 18, top: 0, bottom: 0, width: "1px", background: C.border, zIndex: 0 }} />

                {filteredAdvertisers.length > 0 && <div style={{ marginBottom: 18 }}>
                  <p style={{ fontSize: 11, fontWeight: 500, color: C.t3, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 10, paddingRight: 40 }}>تبلیغ‌کننده‌های مهم {(searchQuery || filterAgency !== "all") ? `(${filteredAdvertisers.length})` : ""}</p>
                  {filteredAdvertisers.map((adv, i) => <AdvCard key={i} adv={adv} type="advertiser" />)}
                </div>}

                {result.leads.length > 0 && <div style={{ marginBottom: 18 }}>
                  <p style={{ fontSize: 11, fontWeight: 500, color: C.t3, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 10, paddingRight: 40 }}>فرصت‌های اپروچ</p>
                  {result.leads.map((l, i) => <AdvCard key={i} adv={l} type="lead" />)}
                </div>}

                {result.competitors.length > 0 && <div style={{ marginBottom: 18 }}>
                  <p style={{ fontSize: 11, fontWeight: 500, color: C.t3, textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 10, paddingRight: 40 }}>پلتفرم‌های رقیب</p>
                  {result.competitors.map((c, i) => (
                    <div key={i} style={{ display: "flex", gap: 14, marginBottom: 10 }}>
                      <div style={{ width: 36, height: 36, borderRadius: 9, background: C.amberDim, border: "1.5px solid rgba(186,117,23,0.2)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 2 }}>
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.amber} strokeWidth="1.5"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12" /></svg>
                      </div>
                      <div style={{ flex: 1, background: C.white, border: `0.5px solid ${C.border}`, borderRadius: 13, padding: "0.9rem 1.1rem" }}>
                        <p style={{ margin: "0 0 4px", fontWeight: 500, fontSize: 13, color: C.amber }}>{c.platform}</p>
                        <p style={{ margin: "0 0 5px", fontSize: 13, color: C.t2, lineHeight: 1.8 }}>{c.note}</p>
                        {c.topclients && <p style={{ margin: "3px 0 0", fontSize: 12, color: C.t3 }}>مهم‌ترین: <span style={{ color: C.t2 }}>{c.topclients}</span></p>}
                        {c.newclients && c.newclients !== "ندارد" && <p style={{ margin: "3px 0 0", fontSize: 12, color: C.green }}>جدید: {c.newclients}</p>}
                      </div>
                    </div>
                  ))}
                </div>}

                {result.market && <div style={{ display: "flex", gap: 14, marginBottom: 18 }}>
                  <div style={{ width: 36, height: 36, borderRadius: 9, background: C.bg, border: `1.5px solid ${C.border}`, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 2 }}>
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.mist} strokeWidth="1.5"><rect x="2" y="3" width="20" height="14" rx="2" /><line x1="12" y1="17" x2="12" y2="21" /></svg>
                  </div>
                  <div style={{ flex: 1, background: C.bg, borderRadius: 13, padding: "0.9rem 1.1rem" }}>
                    <p style={{ margin: "0 0 5px", fontSize: 11, fontWeight: 500, color: C.t3, textTransform: "uppercase", letterSpacing: "0.06em" }}>تحلیل کلی بازار</p>
                    <p style={{ margin: 0, fontSize: 13, color: C.t2, lineHeight: 1.9 }}>{result.market}</p>
                  </div>
                </div>}
              </div>

              {/* Message */}
              <div style={{ background: C.white, border: `0.5px solid ${C.border}`, borderRadius: 15, padding: "1.1rem 1.25rem", marginTop: 4 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                  <p style={{ margin: 0, fontSize: 13, fontWeight: 500, color: C.t1 }}>پیام نهایی</p>
                  <span style={{ fontSize: 11, color: C.t3, background: C.bg, padding: "2px 8px", borderRadius: 20 }}>قالب: {TEMPLATES[selectedTemplate]?.label}</span>
                </div>
                <textarea value={editableMsg} onChange={e => setEditableMsg(e.target.value)} rows={14}
                  style={{ width: "100%", fontSize: 12, padding: 11, borderRadius: 9, border: `0.5px solid ${C.border}`, background: C.bg, color: C.t1, resize: "vertical", boxSizing: "border-box", direction: "rtl", lineHeight: 2, fontFamily: "monospace" }} />
                <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                  <button onClick={reset} style={btn()}>شروع مجدد</button>
                  <button onClick={() => setShowDebug(v => !v)} style={{ ...btn(), color: C.t3 }}>دیباگ</button>
                </div>
                {showDebug && rawDebug && <pre style={{ marginTop: 10, fontSize: 11, color: C.t1, background: C.bg, padding: 11, borderRadius: 9, whiteSpace: "pre-wrap", direction: "ltr", textAlign: "left", maxHeight: 180, overflow: "auto" }}>{rawDebug}</pre>}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
