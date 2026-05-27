import React, { useEffect, useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  PenLine, BookOpen, Brain, BarChart3, CalendarCheck,
  Mic, Delete, X, Check, TrendingUp, TrendingDown, Trash2,
  Sun, Moon, Type, Settings as SettingsIcon, Download, Square, CheckSquare,
} from "lucide-react";

/* ============================================================
   FX DAYLIGHT JOURNAL  ·  accessibility-first, full-feature
   Tabs: Log · Journal · Psych · Stats · Review
   - No system keyboard needed for numbers (big keypad + voice)
   - High-contrast Daylight / Dark themes, adjustable text size
   - Persists across sessions via window.storage (memory fallback)
   ============================================================ */

const TRADES_KEY = "fxjournal:trades:v2";
const REVIEWS_KEY = "fxjournal:reviews:v1";
const SETTINGS_KEY = "fxjournal:settings:v2";

const INSTRUMENTS = {
  Forex: ["EUR/USD", "GBP/USD", "USD/JPY", "GBP/JPY", "AUD/USD", "USD/CAD", "EUR/JPY", "USD/CHF", "NZD/USD"],
  "US Indices": ["US30", "US100", "US500", "US2000"],
  "Metals & Oil": ["XAU/USD", "XTI/USD"],
};
const ALL_INSTRUMENTS = Object.values(INSTRUMENTS).flat();

const SESSIONS = ["Sydney", "Tokyo", "Frankfurt", "London", "New York"];
const PROP_FIRMS = ["FTMO", "MyForexFunds", "The5ers", "FundedNext", "E8 Funding", "True Forex Funds", "Topstep", "Apex Trader", "My Funded Futures", "BluSky"];
const SETUPS = ["Breakout", "Pullback", "Trend", "Reversal", "Range", "Supply/Demand", "Order block", "News", "Scalp"];
const EMOTIONS = ["Calm", "Confident", "Neutral", "Anxious", "FOMO", "Greedy", "Fearful", "Revenge", "Impatient"];
const MISTAKES = ["No stop loss", "Moved stop", "Over-leveraged", "FOMO entry", "Revenge trade", "Closed early", "Chased price", "No clear setup", "Ignored plan", "Overtrading"];
const MOODS = ["Satisfied", "Calm", "Relieved", "Neutral", "Frustrated", "Angry", "Anxious", "Regretful"];
const POSITIVE_MOODS = ["Satisfied", "Calm", "Relieved"];
const CURRENCIES = ["$", "£", "€", "¥"];

const FONT_BODY = "'Atkinson Hyperlegible', system-ui, sans-serif";
const FONT_DISPLAY = "'Archivo', system-ui, sans-serif";

/* ---------- persistence ---------- */
async function loadKey(key, fallback) {
  try {
    if (typeof window !== "undefined" && window.localStorage) {
      const val = window.localStorage.getItem(key);
      if (val) return JSON.parse(val);
    }
  } catch (e) {}
  return fallback;
}
async function saveKey(key, val) {
  try {
    if (typeof window !== "undefined" && window.localStorage) {
      window.localStorage.setItem(key, JSON.stringify(val));
    }
  } catch (e) {}
}

/* ---------- theme ---------- */
function getTheme(name) {
  if (name === "dark") {
    return {
      name: "dark", bg: "#0A0A0A", surface: "#161616", surfaceAlt: "#202020",
      text: "#FFFFFF", muted: "#BFBFBF", border: "#3A3A3A", borderStrong: "#7A7A7A",
      primary: "#5B9DFF", primaryText: "#06122B", buy: "#2EE56B", buyText: "#04210F",
      sell: "#FF6B6B", sellText: "#2A0606", neutral: "#9A9A9A", neutralText: "#0A0A0A",
      keyBtn: "#222", keyText: "#FFF", keyBorder: "#4A4A4A", chipOff: "#1F1F1F", chipOffText: "#FFF",
      shadow: "0 10px 30px rgba(0,0,0,0.65)",
    };
  }
  return {
    name: "daylight", bg: "#FFFFFF", surface: "#FFFFFF", surfaceAlt: "#F1F1F1",
    text: "#0A0A0A", muted: "#3A3A3A", border: "#111111", borderStrong: "#000000",
    primary: "#0B3D91", primaryText: "#FFFFFF", buy: "#0B7A2F", buyText: "#FFFFFF",
    sell: "#BE1212", sellText: "#FFFFFF", neutral: "#555555", neutralText: "#FFFFFF",
    keyBtn: "#FFFFFF", keyText: "#0A0A0A", keyBorder: "#000000", chipOff: "#FFFFFF", chipOffText: "#0A0A0A",
    shadow: "0 10px 30px rgba(0,0,0,0.20)",
  };
}

const SCALE_LABELS = { "1": "Normal", "1.2": "Large", "1.4": "Huge" };
function makeFS(scale) {
  const f = (n) => Math.round(n * scale);
  return { xs: f(12), sm: f(14), body: f(16), md: f(18), lg: f(21), xl: f(25), display: f(34), huge: f(44) };
}

/* ---------- misc helpers ---------- */
const todayISO = () => new Date().toISOString().slice(0, 10);
const num = (v) => (v === "" || v == null ? NaN : parseFloat(v));
function computeRR(entry, sl, tp) {
  const e = num(entry), s = num(sl), t = num(tp);
  if ([e, s, t].some(isNaN)) return null;
  const risk = Math.abs(e - s), reward = Math.abs(t - e);
  if (risk === 0) return null;
  return reward / risk;
}
const fmtMoney = (n, cur) => `${n > 0 ? "+" : n < 0 ? "−" : ""}${cur}${Math.abs(n).toLocaleString(undefined, { maximumFractionDigits: 2 })}`;

function parseSpokenNumber(text) {
  if (!text) return null;
  let t = (" " + text.toLowerCase() + " ")
    .replace(/\b(negative|minus)\b/g, " - ").replace(/\b(plus|positive)\b/g, " ")
    .replace(/\b(point|dot|decimal)\b/g, " . ");
  const w = { zero: "0", oh: "0", one: "1", two: "2", three: "3", four: "4", five: "5", six: "6", seven: "7", eight: "8", nine: "9" };
  t = t.replace(/[a-z]+/g, (m) => (w[m] !== undefined ? w[m] : " ")).replace(/\s+/g, "");
  const m = t.match(/-?\d*\.?\d+/);
  return m ? m[0] : null;
}

// Combine the multi-select setups + custom setup into one display/export string.
// Falls back to the old single `setup` field for any trades logged before this update.
function setupText(trade) {
  const list = Array.isArray(trade.setups) ? [...trade.setups] : (trade.setup ? [trade.setup] : []);
  if (trade.customSetup && trade.customSetup.trim()) list.push(trade.customSetup.trim());
  return list.join(", ");
}
// Pre-trade emotions / post-trade moods as arrays, with fallback to old single-value fields.
function preEmotionsOf(trade) {
  return Array.isArray(trade.preEmotions) ? trade.preEmotions : (trade.preEmotion ? [trade.preEmotion] : []);
}
function postMoodsOf(trade) {
  return Array.isArray(trade.postMoods) ? trade.postMoods : (trade.postMood ? [trade.postMood] : []);
}

const emptyDraft = () => ({
  date: todayISO(), pnl: "", instrument: "", direction: "",
  propFirm: "", customFirm: "",
  session: "", setups: [], customSetup: "", entry: "", stopLoss: "", takeProfit: "", lots: "",
  entryQuality: 5, outcome: "", hitTP: "", movedSL: "",
  preEmotions: [], confidence: 5, mistakeTags: [], followedPlan: "", reflection: "", postMoods: [],
});
const emptyReview = () => ({ date: todayISO(), rating: 5, conditions: "", wentWell: "", wentWrong: "", lesson: "" });

/* ===================================================================== */
export default function App() {
  const [view, setView] = useState("log");
  const [trades, setTrades] = useState([]);
  const [reviews, setReviews] = useState([]);
  const [settings, setSettings] = useState({ theme: "daylight", scale: "1", haptics: true, currency: "$" });
  const [loaded, setLoaded] = useState(false);
  const [draft, setDraft] = useState(emptyDraft());
  const [editingId, setEditingId] = useState(null);
  const [reviewDraft, setReviewDraft] = useState(emptyReview());
  const [flashMsg, setFlashMsg] = useState("");
  const [keypad, setKeypad] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [selectedFirm, setSelectedFirm] = useState("All");
  const flashTimer = useRef(null);

  const T = getTheme(settings.theme);
  const fs = makeFS(parseFloat(settings.scale));
  const cur = settings.currency || "$";

  useEffect(() => {
    (async () => {
      const t = await loadKey(TRADES_KEY, []);
      const r = await loadKey(REVIEWS_KEY, []);
      const s = await loadKey(SETTINGS_KEY, null);
      if (Array.isArray(t)) setTrades(t);
      if (Array.isArray(r)) setReviews(r);
      if (s) setSettings((p) => ({ ...p, ...s }));
      setLoaded(true);
    })();
  }, []);

  const buzz = useCallback((ms = 10) => {
    if (settings.haptics && typeof navigator !== "undefined" && navigator.vibrate) navigator.vibrate(ms);
  }, [settings.haptics]);

  const flash = useCallback((msg) => {
    setFlashMsg(msg);
    if (flashTimer.current) clearTimeout(flashTimer.current);
    flashTimer.current = setTimeout(() => setFlashMsg(""), 2200);
  }, []);

  const updateSettings = (patch) => { const next = { ...settings, ...patch }; setSettings(next); saveKey(SETTINGS_KEY, next); };

  const openKeypad = (field, label, opts = {}) => { buzz(); setKeypad({ field, label, allowNeg: !!opts.allowNeg, allowDec: opts.allowDec !== false, value: draft[field] || "" }); };
  const commitKeypad = () => {
    if (!keypad) return;
    let v = keypad.value;
    if (v === "-" || v === "." || v === "-.") v = "";
    setDraft((d) => {
      const next = { ...d, [keypad.field]: v };
      if (keypad.field === "pnl") { const n = parseFloat(v); if (!isNaN(n)) next.outcome = n > 0 ? "win" : n < 0 ? "loss" : "be"; }
      return next;
    });
    setKeypad(null);
  };

  const addTrade = () => {
    buzz(20);
    if (!draft.instrument) { flash("Pick an instrument"); return; }
    if (!draft.direction) { flash("Pick Long or Short"); return; }
    if (!draft.outcome) { flash("Pick an outcome (Win / Loss / B-E)"); return; }
    const pnlNum = draft.pnl === "" ? 0 : parseFloat(draft.pnl);
    const rr = computeRR(draft.entry, draft.stopLoss, draft.takeProfit);
    if (editingId) {
      // update in place, preserving id, createdAt and list position
      const next = trades.map((t) => t.id === editingId
        ? { ...t, ...draft, id: t.id, createdAt: t.createdAt, editedAt: new Date().toISOString(), pnl: isNaN(pnlNum) ? 0 : pnlNum, rr }
        : t);
      setTrades(next); saveKey(TRADES_KEY, next);
      setDraft(emptyDraft()); setEditingId(null); flash("Trade updated ✓"); setView("journal");
      return;
    }
    const trade = { ...draft, id: Date.now() + "-" + Math.floor(Math.random() * 1e6), createdAt: new Date().toISOString(), pnl: isNaN(pnlNum) ? 0 : pnlNum, rr };
    const next = [trade, ...trades];
    setTrades(next); saveKey(TRADES_KEY, next);
    setDraft(emptyDraft()); flash("Trade logged ✓"); setView("journal");
  };

  const editTrade = (trade) => {
    buzz();
    // hydrate the form with the existing trade, normalising number fields to strings
    setDraft({
      ...emptyDraft(), ...trade,
      pnl: trade.pnl == null ? "" : String(trade.pnl),
      entry: trade.entry == null ? "" : String(trade.entry),
      stopLoss: trade.stopLoss == null ? "" : String(trade.stopLoss),
      takeProfit: trade.takeProfit == null ? "" : String(trade.takeProfit),
      lots: trade.lots == null ? "" : String(trade.lots),
      mistakeTags: Array.isArray(trade.mistakeTags) ? trade.mistakeTags : [],
      setups: Array.isArray(trade.setups) ? trade.setups : (trade.setup ? [trade.setup] : []),
      customSetup: trade.customSetup || "",
      preEmotions: Array.isArray(trade.preEmotions) ? trade.preEmotions : (trade.preEmotion ? [trade.preEmotion] : []),
      postMoods: Array.isArray(trade.postMoods) ? trade.postMoods : (trade.postMood ? [trade.postMood] : []),
      propFirm: trade.propFirm || "",
      customFirm: trade.customFirm || "",
    });
    setEditingId(trade.id);
    setView("log");
    flash("Editing trade — make your changes");
  };

  const cancelEdit = () => { setDraft(emptyDraft()); setEditingId(null); flash("Edit cancelled"); setView("journal"); };

  const deleteTrade = (id) => { const next = trades.filter((t) => t.id !== id); setTrades(next); saveKey(TRADES_KEY, next); if (editingId === id) { setEditingId(null); setDraft(emptyDraft()); } flash("Trade deleted"); buzz(20); };

  const saveReview = () => {
    buzz(20);
    const rev = { ...reviewDraft, id: Date.now() + "-" + Math.floor(Math.random() * 1e6) };
    const next = [rev, ...reviews]; setReviews(next); saveKey(REVIEWS_KEY, next);
    setReviewDraft(emptyReview()); flash("Session review saved ✓");
  };
  const deleteReview = (id) => { const next = reviews.filter((r) => r.id !== id); setReviews(next); saveKey(REVIEWS_KEY, next); buzz(20); };

  const resetAll = () => { setTrades([]); setReviews([]); saveKey(TRADES_KEY, []); saveKey(REVIEWS_KEY, []); setShowSettings(false); flash("All data cleared"); };

  const screen = { minHeight: "100vh", background: T.bg, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, lineHeight: 1.4, WebkitTapHighlightColor: "transparent" };

  return (
    <div style={screen}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Atkinson+Hyperlegible:ital,wght@0,400;0,700;1,400;1,700&family=Archivo:wght@700;800;900&display=swap');
        * { box-sizing: border-box; }
        button { font-family: inherit; cursor: pointer; }
        ::-webkit-scrollbar { width:0; height:0; }
        .fxbtn { transition: transform .06s ease; }
        .fxbtn:active { transform: scale(.97); }
        textarea:focus, button:focus-visible, input:focus-visible { outline: 3px solid ${T.primary}; outline-offset: 2px; }
        input[type=date]{ color:${T.text}; }
        @keyframes pop { from{transform:scale(.96);opacity:.4} to{transform:scale(1);opacity:1} }
        @keyframes sheetUp { from{transform:translateY(100%)} to{transform:translateY(0)} }
        @media (prefers-reduced-motion: reduce){ .fxbtn,[data-anim]{animation:none!important;transition:none!important} }
      `}</style>

      {/* Header */}
      <header style={{ position: "sticky", top: 0, zIndex: 20, background: T.surface, borderBottom: `2px solid ${T.border}`, padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
          <span aria-hidden style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.md, color: T.primary }}>FX</span>
          <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.body, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>Daylight Journal</span>
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          <IconBtn T={T} fs={fs} label="Change text size" onClick={() => { buzz(); const o = ["1", "1.2", "1.4"]; updateSettings({ scale: o[(o.indexOf(settings.scale) + 1) % o.length] }); }}><Type size={Math.round(fs.md)} strokeWidth={2.5} /></IconBtn>
          <IconBtn T={T} fs={fs} label="Toggle theme" onClick={() => { buzz(); updateSettings({ theme: settings.theme === "daylight" ? "dark" : "daylight" }); }}>{settings.theme === "daylight" ? <Moon size={Math.round(fs.md)} strokeWidth={2.5} /> : <Sun size={Math.round(fs.md)} strokeWidth={2.5} />}</IconBtn>
          <IconBtn T={T} fs={fs} label="Settings" onClick={() => { buzz(); setShowSettings(true); }}><SettingsIcon size={Math.round(fs.md)} strokeWidth={2.5} /></IconBtn>
        </div>
      </header>

      {flashMsg && (
        <div role="status" aria-live="polite" data-anim style={{ position: "fixed", top: 64, left: "50%", transform: "translateX(-50%)", zIndex: 60, background: T.primary, color: T.primaryText, fontWeight: 800, fontSize: fs.sm, padding: "10px 18px", borderRadius: 10, boxShadow: T.shadow, animation: "pop .15s ease", maxWidth: "92%", textAlign: "center" }}>{flashMsg}</div>
      )}

      {/* Firm filter bar — always visible so analytics stay firm-specific */}
      {(() => {
        const firmNames = (trade) => trade.customFirm && trade.customFirm.trim() ? trade.customFirm.trim() : (trade.propFirm || "Unassigned");
        const allFirms = ["All", ...Array.from(new Set(trades.map(firmNames))).sort()];
        const filteredTrades = selectedFirm === "All" ? trades : trades.filter((t) => firmNames(t) === selectedFirm);
        return (
          <>
            <div style={{ background: T.surfaceAlt, borderBottom: `2px solid ${T.border}`, padding: "8px 12px", overflowX: "auto", display: "flex", gap: 8, alignItems: "center" }}>
              <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.xs, color: T.muted, whiteSpace: "nowrap" }}>FIRM:</span>
              {allFirms.map((f) => {
                const on = selectedFirm === f;
                return <button key={f} className="fxbtn" onClick={() => { buzz(); setSelectedFirm(f); }} aria-pressed={on} style={{ flex: "0 0 auto", minHeight: 36, padding: "0 12px", borderRadius: 8, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.xs, background: on ? T.primary : T.surface, color: on ? T.primaryText : T.text, border: `2px solid ${on ? T.primary : T.border}`, whiteSpace: "nowrap" }}>{f}</button>;
              })}
            </div>

            <main style={{ maxWidth: 560, margin: "0 auto", padding: 12, paddingBottom: 96 }}>
              {view === "log" && <LogTrade {...{ T, fs, cur, draft, setDraft, openKeypad, buzz, flash, addTrade, editingId, cancelEdit }} />}
              {view === "journal" && <JournalView {...{ T, fs, cur, trades: filteredTrades, allTrades: trades, deleteTrade, editTrade, buzz, loaded, goLog: () => setView("log"), selectedFirm }} />}
              {view === "psych" && <PsychView {...{ T, fs, cur, trades: filteredTrades, selectedFirm }} />}
              {view === "stats" && <StatsView {...{ T, fs, cur, trades: filteredTrades, allTrades: trades, flash, selectedFirm }} />}
              {view === "review" && <ReviewView {...{ T, fs, reviewDraft, setReviewDraft, saveReview, reviews, deleteReview, buzz, flash }} />}
            </main>
          </>
        );
      })()}

      {/* Bottom nav */}
      <nav style={{ position: "fixed", bottom: 0, left: 0, right: 0, zIndex: 30, background: T.surface, borderTop: `2px solid ${T.border}`, display: "flex", padding: "6px 6px calc(6px + env(safe-area-inset-bottom))", gap: 5, maxWidth: 560, margin: "0 auto" }}>
        <NavTab T={T} fs={fs} active={view === "log"} icon={<PenLine size={Math.round(fs.md)} strokeWidth={2.5} />} label="Log" onClick={() => { buzz(); setView("log"); }} />
        <NavTab T={T} fs={fs} active={view === "journal"} icon={<BookOpen size={Math.round(fs.md)} strokeWidth={2.3} />} label="Journal" onClick={() => { buzz(); setView("journal"); }} />
        <NavTab T={T} fs={fs} active={view === "psych"} icon={<Brain size={Math.round(fs.md)} strokeWidth={2.3} />} label="Psych" onClick={() => { buzz(); setView("psych"); }} />
        <NavTab T={T} fs={fs} active={view === "stats"} icon={<BarChart3 size={Math.round(fs.md)} strokeWidth={2.3} />} label="Stats" onClick={() => { buzz(); setView("stats"); }} />
        <NavTab T={T} fs={fs} active={view === "review"} icon={<CalendarCheck size={Math.round(fs.md)} strokeWidth={2.3} />} label="Review" onClick={() => { buzz(); setView("review"); }} />
      </nav>

      {keypad && <Keypad {...{ T, fs, cur, keypad, setKeypad, commit: commitKeypad, buzz }} />}
      {showSettings && <SettingsSheet {...{ T, fs, settings, updateSettings, buzz, close: () => setShowSettings(false), resetAll, count: trades.length + reviews.length }} />}
    </div>
  );
}

/* ===================== shared atoms ===================== */
function IconBtn({ children, onClick, label, T, fs }) {
  return <button className="fxbtn" onClick={onClick} aria-label={label} title={label} style={{ width: 42, height: 42, minWidth: 42, minHeight: 42, borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, display: "grid", placeItems: "center" }}>{children}</button>;
}
function NavTab({ active, icon, label, onClick, T, fs }) {
  return <button className="fxbtn" onClick={onClick} aria-label={label} aria-current={active} style={{ flex: 1, minHeight: 54, borderRadius: 12, background: active ? T.primary : "transparent", color: active ? T.primaryText : T.text, border: active ? `2px solid ${T.primary}` : `2px solid ${T.border}`, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 1, fontWeight: 800, fontSize: fs.xs, fontFamily: FONT_DISPLAY }}>{icon}<span>{label}</span></button>;
}
function Section({ title, T, fs, children, sub, accent }) {
  return (
    <div style={{ background: T.surface, border: `2px solid ${accent || T.border}`, borderRadius: 12, padding: 13, marginBottom: 12 }}>
      {title && <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, letterSpacing: ".04em", textTransform: "uppercase", color: accent || T.muted, marginBottom: sub ? 2 : 10 }}>{title}</div>}
      {sub && <div style={{ color: T.muted, fontSize: fs.xs, marginBottom: 10 }}>{sub}</div>}
      {children}
    </div>
  );
}
function Label({ children, T, fs }) {
  return <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.xs, letterSpacing: ".04em", textTransform: "uppercase", color: T.muted, margin: "4px 2px 8px" }}>{children}</div>;
}
function Chips({ options, value, onChange, multi, columns = 2, T, fs, buzz }) {
  const isOn = (o) => (multi ? value.includes(o) : value === o);
  const toggle = (o) => { buzz && buzz(); if (multi) onChange(value.includes(o) ? value.filter((x) => x !== o) : [...value, o]); else onChange(value === o ? "" : o); };
  return (
    <div style={{ display: "grid", gridTemplateColumns: `repeat(${columns},1fr)`, gap: 8 }}>
      {options.map((o) => { const on = isOn(o); return (
        <button key={o} className="fxbtn" aria-pressed={on} onClick={() => toggle(o)} style={{ minHeight: 46, borderRadius: 10, fontFamily: FONT_DISPLAY, fontWeight: 700, fontSize: fs.sm, background: on ? T.primary : T.chipOff, color: on ? T.primaryText : T.chipOffText, border: `2px solid ${on ? T.primary : T.border}`, padding: "0 6px", display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
          {on && <Check size={Math.round(fs.sm)} strokeWidth={3} />}{o}
        </button>); })}
    </div>
  );
}
function Seg({ options, value, onChange, T, fs, buzz }) {
  return (
    <div style={{ display: "flex", gap: 8 }}>
      {options.map((o) => { const on = value === o.k; const c = o.color || T.primary; const tc = o.textColor || T.primaryText; return (
        <button key={o.k} className="fxbtn" aria-pressed={on} onClick={() => { buzz && buzz(); onChange(on ? "" : o.k); }} style={{ flex: 1, minHeight: 48, borderRadius: 10, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, background: on ? c : T.chipOff, color: on ? tc : T.chipOffText, border: `2px solid ${on ? c : T.border}`, display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
          {o.icon}{o.label}
        </button>); })}
    </div>
  );
}
function RatingStepper({ value, onChange, labels, color, T, fs, buzz }) {
  const c = color || T.primary;
  const mini = { width: 46, minHeight: 44, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surfaceAlt, color: T.text, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg, lineHeight: 1 };
  const set = (v) => { buzz && buzz(); onChange(Math.max(1, Math.min(10, v))); };
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <button className="fxbtn" aria-label="decrease" onClick={() => set(value - 1)} style={mini}>−</button>
        <div style={{ flex: 1, height: 44, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surfaceAlt, position: "relative", overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: `${(value / 10) * 100}%`, background: c, opacity: 0.22 }} />
          <span style={{ position: "relative", fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg, color: T.text }}>{value}/10</span>
        </div>
        <button className="fxbtn" aria-label="increase" onClick={() => set(value + 1)} style={mini}>+</button>
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: fs.xs, color: T.muted, marginTop: 4, padding: "0 2px" }}>
        <span>{labels[0]}</span><span>{labels[1]}</span><span>{labels[2]}</span>
      </div>
    </div>
  );
}
function NumField({ label, value, onTap, T, fs, prefix, color }) {
  return (
    <div style={{ flex: 1, minWidth: 0 }}>
      <Label T={T} fs={fs}>{label}</Label>
      <button className="fxbtn" onClick={onTap} aria-label={`${label}. ${value ? "Value " + value : "empty"}. Tap to enter.`} style={{ width: "100%", minHeight: 50, textAlign: "left", padding: "0 12px", borderRadius: 10, border: `2px solid ${T.border}`, background: T.surface, color: value ? (color || T.text) : T.muted, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.lg, display: "flex", alignItems: "center", gap: 4 }}>
        {prefix && value ? <span style={{ color: T.muted }}>{prefix}</span> : null}{value || "—"}
      </button>
    </div>
  );
}

/* ===================== Voice ===================== */
function VoiceBtn({ T, fs, onResult, mode, buzz, flash, full }) {
  const [listening, setListening] = useState(false);
  const recRef = useRef(null);
  const start = () => {
    const SR = typeof window !== "undefined" && (window.SpeechRecognition || window.webkitSpeechRecognition);
    if (!SR) { flash && flash("Voice input isn't supported here"); return; }
    buzz(20);
    try {
      const rec = new SR(); rec.lang = "en-GB"; rec.interimResults = false; rec.maxAlternatives = 3;
      rec.onresult = (e) => {
        let best = null;
        for (let i = 0; i < e.results[0].length; i++) { const txt = e.results[0][i].transcript; if (mode === "number") { const n = parseSpokenNumber(txt); if (n !== null) { best = n; break; } } else if (!best) best = txt; }
        if (best) onResult(best); else flash && flash("Didn't catch that — try again");
      };
      rec.onerror = () => { setListening(false); flash && flash("Couldn't hear you — check mic permission"); };
      rec.onend = () => setListening(false);
      recRef.current = rec; setListening(true); rec.start();
    } catch (e) { setListening(false); }
  };
  const stop = () => { try { recRef.current && recRef.current.stop(); } catch (e) {} setListening(false); };
  return (
    <button className="fxbtn" onClick={() => (listening ? stop() : start())} aria-label={listening ? "Stop listening" : "Speak instead of typing"} style={{ width: full ? "100%" : 56, minWidth: 56, minHeight: 50, borderRadius: 10, background: listening ? T.sell : T.surfaceAlt, color: listening ? T.sellText : T.text, border: `2px solid ${listening ? T.sell : T.border}`, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>
      <Mic size={Math.round(fs.md)} strokeWidth={2.5} />{full && <span>{listening ? "Listening…" : "Speak"}</span>}
    </button>
  );
}

/* ===================== Keypad ===================== */
function Keypad({ T, fs, cur, keypad, setKeypad, commit, buzz }) {
  const set = (v) => setKeypad((k) => ({ ...k, value: v }));
  const v = keypad.value;
  const press = (ch) => { buzz(); if (ch === ".") { if (!keypad.allowDec || v.includes(".")) return; set((v || "0") + "."); return; } set(v + ch); };
  const display = v === "" ? "0" : v;
  const dColor = keypad.field === "pnl" ? (parseFloat(v) > 0 ? T.buy : parseFloat(v) < 0 ? T.sell : T.text) : T.text;
  const key = { minHeight: 58, borderRadius: 12, background: T.keyBtn, color: T.keyText, border: `2px solid ${T.keyBorder}`, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.xl, display: "grid", placeItems: "center" };
  return (
    <div role="dialog" aria-modal="true" aria-label={keypad.label} style={{ position: "fixed", inset: 0, zIndex: 80, background: "rgba(0,0,0,.55)", display: "flex", flexDirection: "column", justifyContent: "flex-end" }} onClick={(e) => { if (e.target === e.currentTarget) setKeypad(null); }}>
      <div data-anim style={{ background: T.surface, borderTop: `3px solid ${T.border}`, borderTopLeftRadius: 22, borderTopRightRadius: 22, padding: "14px 14px calc(14px + env(safe-area-inset-bottom))", maxWidth: 560, width: "100%", margin: "0 auto", boxShadow: T.shadow, animation: "sheetUp .18s ease" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, textTransform: "uppercase", letterSpacing: ".04em", color: T.muted }}>{keypad.label}</span>
          <button className="fxbtn" onClick={() => setKeypad(null)} aria-label="Cancel" style={{ width: 48, height: 48, borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, display: "grid", placeItems: "center" }}><X size={Math.round(fs.lg)} strokeWidth={2.6} /></button>
        </div>
        <div style={{ minHeight: 64, borderRadius: 12, border: `2px solid ${T.border}`, background: T.surfaceAlt, display: "flex", alignItems: "center", justifyContent: "flex-end", padding: "0 16px", marginBottom: 12, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.huge, color: dColor, overflow: "hidden" }}>{keypad.field === "pnl" && cur}{display}</div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
          {["1", "2", "3", "4", "5", "6", "7", "8", "9"].map((n) => <button key={n} className="fxbtn" style={key} onClick={() => press(n)} aria-label={n}>{n}</button>)}
          {keypad.allowDec ? <button className="fxbtn" style={key} onClick={() => press(".")} aria-label="point">.</button> : <div />}
          <button className="fxbtn" style={key} onClick={() => press("0")} aria-label="0">0</button>
          <button className="fxbtn" style={{ ...key, background: T.surfaceAlt }} onClick={() => { buzz(); set(v.slice(0, -1)); }} aria-label="Delete digit"><Delete size={Math.round(fs.xl)} strokeWidth={2.3} /></button>
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
          {keypad.allowNeg && <button className="fxbtn" onClick={() => { buzz(); set(v.startsWith("-") ? v.slice(1) : "-" + v); }} aria-label="Toggle sign" style={{ flex: 1, minHeight: 52, borderRadius: 10, border: `2px solid ${T.border}`, background: v.startsWith("-") ? T.sell : T.surfaceAlt, color: v.startsWith("-") ? T.sellText : T.text, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>± Sign</button>}
          <button className="fxbtn" onClick={() => { buzz(20); set(""); }} aria-label="Clear" style={{ flex: 1, minHeight: 52, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surfaceAlt, color: T.text, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>Clear</button>
          <VoiceBtn T={T} fs={fs} mode="number" buzz={buzz} flash={() => {}} onResult={(n) => set(n)} />
        </div>
        <button className="fxbtn" onClick={commit} aria-label="Confirm" style={{ width: "100%", minHeight: 60, borderRadius: 14, border: "none", marginTop: 10, background: T.primary, color: T.primaryText, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><Check size={Math.round(fs.lg)} strokeWidth={3} /> Done</button>
      </div>
    </div>
  );
}

/* ===================== LOG ===================== */
function LogTrade({ T, fs, cur, draft, setDraft, openKeypad, buzz, flash, addTrade, editingId, cancelEdit }) {
  const set = (patch) => setDraft((d) => ({ ...d, ...patch }));
  const rr = computeRR(draft.entry, draft.stopLoss, draft.takeProfit);
  const inputStyle = { width: "100%", minHeight: 50, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surface, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, padding: "0 12px" };

  return (
    <div>
      {/* P&L hero */}
      <Section T={T} fs={fs} title={editingId ? "Edit trade" : "Log trade"} sub={editingId ? "Fix any details, then save your changes" : "Record every trade, every time"} accent={editingId ? T.primary : undefined}>
        {editingId && (
          <div style={{ marginBottom: 12, padding: "8px 12px", borderRadius: 10, background: T.surfaceAlt, border: `2px solid ${T.primary}`, color: T.text, fontSize: fs.sm, fontWeight: 700, display: "flex", alignItems: "center", gap: 8 }}>
            <PenLine size={Math.round(fs.md)} strokeWidth={2.5} /> You're editing an existing trade.
          </div>
        )}
        <button className="fxbtn" onClick={() => openKeypad("pnl", `Profit / Loss (${cur})`, { allowNeg: true })} aria-label={`Profit or loss ${draft.pnl ? cur + draft.pnl : "not set"}. Tap to enter.`} style={{ width: "100%", minHeight: 70, borderRadius: 12, border: `2px solid ${T.border}`, background: T.surfaceAlt, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.huge, color: draft.pnl === "" ? T.muted : parseFloat(draft.pnl) > 0 ? T.buy : parseFloat(draft.pnl) < 0 ? T.sell : T.text, display: "flex", alignItems: "center", justifyContent: "center" }}>
          {draft.pnl === "" ? `${cur}0.00 — tap to enter P&L` : `${parseFloat(draft.pnl) > 0 ? "+" : ""}${cur}${draft.pnl}`}
        </button>
        <div style={{ marginTop: 12 }}>
          <Label T={T} fs={fs}>Date</Label>
          <input type="date" value={draft.date} onChange={(e) => set({ date: e.target.value })} style={inputStyle} aria-label="Trade date" />
        </div>
      </Section>

      {/* Prop Firm */}
      <Section T={T} fs={fs} title="Prop Firm" accent={draft.propFirm || draft.customFirm ? T.primary : undefined}>
        <Chips options={PROP_FIRMS} value={draft.propFirm} onChange={(v) => set({ propFirm: v, customFirm: "" })} columns={2} T={T} fs={fs} buzz={buzz} />
        <input value={draft.customFirm} onChange={(e) => set({ customFirm: e.target.value, propFirm: "" })} placeholder="+ Type your own firm name" aria-label="Custom prop firm name" style={{ width: "100%", minHeight: 48, marginTop: 8, borderRadius: 10, border: `2px solid ${draft.customFirm ? T.primary : T.border}`, background: T.surface, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, padding: "0 12px" }} />
      </Section>

      {/* Instrument */}
      <Section T={T} fs={fs} title="Instrument">
        {Object.entries(INSTRUMENTS).map(([group, list]) => (
          <div key={group} style={{ marginBottom: 10 }}>
            <div style={{ fontSize: fs.xs, fontWeight: 700, color: T.muted, marginBottom: 6 }}>{group}</div>
            <Chips options={list} value={draft.instrument} onChange={(v) => set({ instrument: v })} columns={group === "Metals & Oil" ? 2 : 3} T={T} fs={fs} buzz={buzz} />
          </div>
        ))}
      </Section>

      {/* Direction */}
      <Section T={T} fs={fs} title="Direction">
        <Seg T={T} fs={fs} buzz={buzz} value={draft.direction} onChange={(v) => set({ direction: v })} options={[
          { k: "long", label: "LONG", color: T.buy, textColor: T.buyText, icon: <TrendingUp size={Math.round(fs.md)} strokeWidth={2.8} /> },
          { k: "short", label: "SHORT", color: T.sell, textColor: T.sellText, icon: <TrendingDown size={Math.round(fs.md)} strokeWidth={2.8} /> },
        ]} />
      </Section>

      {/* Execution */}
      <Section T={T} fs={fs} title="Execution">
        <Label T={T} fs={fs}>Session</Label>
        <Chips options={SESSIONS} value={draft.session} onChange={(v) => set({ session: v })} columns={3} T={T} fs={fs} buzz={buzz} />
        <Label T={T} fs={fs}>Setup / Strategy (select all that apply)</Label>
        <Chips options={SETUPS} value={draft.setups} onChange={(v) => set({ setups: v })} multi columns={3} T={T} fs={fs} buzz={buzz} />
        <input value={draft.customSetup} onChange={(e) => set({ customSetup: e.target.value })} placeholder="+ Add your own setup (optional)" aria-label="Custom setup" style={{ width: "100%", minHeight: 48, marginTop: 8, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surface, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, padding: "0 12px" }} />
        <div style={{ display: "flex", gap: 10, marginTop: 4 }}>
          <NumField label="Entry" value={draft.entry} onTap={() => openKeypad("entry", "Entry price")} T={T} fs={fs} />
          <NumField label="Stop Loss" value={draft.stopLoss} onTap={() => openKeypad("stopLoss", "Stop Loss")} T={T} fs={fs} color={T.sell} />
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 4 }}>
          <NumField label="Take Profit" value={draft.takeProfit} onTap={() => openKeypad("takeProfit", "Take Profit")} T={T} fs={fs} color={T.buy} />
          <NumField label="Lot Size" value={draft.lots} onTap={() => openKeypad("lots", "Lot size")} T={T} fs={fs} />
        </div>
        <div style={{ marginTop: 10, padding: "10px 12px", borderRadius: 10, border: `2px dashed ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontWeight: 700, color: T.muted, fontSize: fs.sm }}>Risk : Reward</span>
          <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg, color: rr == null ? T.muted : rr >= 1 ? T.buy : T.sell }}>{rr == null ? "—" : `1 : ${rr.toFixed(2)}`}</span>
        </div>
        <div style={{ marginTop: 12 }}>
          <Label T={T} fs={fs}>Entry quality</Label>
          <RatingStepper value={draft.entryQuality} onChange={(v) => set({ entryQuality: v })} labels={["Poor", "Average", "Perfect"]} color={T.primary} T={T} fs={fs} buzz={buzz} />
        </div>
      </Section>

      {/* Outcome */}
      <Section T={T} fs={fs} title="Outcome">
        <Seg T={T} fs={fs} buzz={buzz} value={draft.outcome} onChange={(v) => set({ outcome: v })} options={[
          { k: "win", label: "✓ WIN", color: T.buy, textColor: T.buyText },
          { k: "loss", label: "✗ LOSS", color: T.sell, textColor: T.sellText },
          { k: "be", label: "B/E", color: T.neutral, textColor: T.neutralText },
        ]} />
        <Label T={T} fs={fs}>Hit your original TP?</Label>
        <Seg T={T} fs={fs} buzz={buzz} value={draft.hitTP} onChange={(v) => set({ hitTP: v })} options={[{ k: "hit", label: "✓ Hit TP" }, { k: "early", label: "Exited early" }]} />
        <Label T={T} fs={fs}>Move your stop loss?</Label>
        <Seg T={T} fs={fs} buzz={buzz} value={draft.movedSL} onChange={(v) => set({ movedSL: v })} options={[{ k: "trailed", label: "Trailed" }, { k: "widened", label: "Widened" }, { k: "kept", label: "Kept it" }]} />
      </Section>

      {/* Psychology */}
      <Section T={T} fs={fs} title="Psychology">
        <Label T={T} fs={fs}>Pre-trade emotion (select all that apply)</Label>
        <Chips options={EMOTIONS} value={draft.preEmotions} onChange={(v) => set({ preEmotions: v })} multi columns={3} T={T} fs={fs} buzz={buzz} />
        <Label T={T} fs={fs}>Confidence in setup</Label>
        <RatingStepper value={draft.confidence} onChange={(v) => set({ confidence: v })} labels={["No edge", "Moderate", "A+ setup"]} color={T.primary} T={T} fs={fs} buzz={buzz} />
        <Label T={T} fs={fs}>Mistake tags (select all that apply)</Label>
        <Chips options={MISTAKES} value={draft.mistakeTags} onChange={(v) => set({ mistakeTags: v })} multi columns={2} T={T} fs={fs} buzz={buzz} />
        <Label T={T} fs={fs}>Followed your plan?</Label>
        <Seg T={T} fs={fs} buzz={buzz} value={draft.followedPlan} onChange={(v) => set({ followedPlan: v })} options={[{ k: "yes", label: "✓ Yes", color: T.buy, textColor: T.buyText }, { k: "no", label: "✗ No", color: T.sell, textColor: T.sellText }]} />
        <Label T={T} fs={fs}>Post-trade reflection</Label>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-start" }}>
          <textarea value={draft.reflection} onChange={(e) => set({ reflection: e.target.value })} rows={3} placeholder="What did you learn?" style={{ flex: 1, resize: "vertical", minHeight: 56, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surface, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, padding: 10 }} />
          <VoiceBtn T={T} fs={fs} mode="text" buzz={buzz} flash={flash} onResult={(txt) => set({ reflection: (draft.reflection ? draft.reflection + " " : "") + txt })} />
        </div>
        <Label T={T} fs={fs}>How do you feel after this trade? (select all that apply)</Label>
        <Chips options={MOODS} value={draft.postMoods} onChange={(v) => set({ postMoods: v })} multi columns={4} T={T} fs={fs} buzz={buzz} />
      </Section>

      <button className="fxbtn" onClick={addTrade} style={{ width: "100%", minHeight: 60, borderRadius: 14, border: "none", background: T.primary, color: T.primaryText, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg, display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}><Check size={Math.round(fs.lg)} strokeWidth={3} /> {editingId ? "Save changes" : "Log this trade"}</button>
      {editingId && (
        <button className="fxbtn" onClick={cancelEdit} style={{ width: "100%", minHeight: 52, borderRadius: 12, marginTop: 10, background: T.surface, color: T.text, border: `2px solid ${T.border}`, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.md, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><X size={Math.round(fs.md)} strokeWidth={2.6} /> Cancel edit</button>
      )}
      <p style={{ textAlign: "center", color: T.muted, fontSize: fs.xs, marginTop: 8 }}>Instrument · Direction · Outcome required</p>
    </div>
  );
}

/* ===================== JOURNAL ===================== */
function JournalView({ T, fs, cur, trades, allTrades, deleteTrade, editTrade, buzz, loaded, goLog, selectedFirm }) {
  const [pairFilter, setPairFilter] = useState("All");
  const [outFilter, setOutFilter] = useState("all");
  const [confirm, setConfirm] = useState(null);
  const pairs = ["All", ...Array.from(new Set(trades.map((t) => t.instrument)))];
  const filtered = trades.filter((t) => (pairFilter === "All" || t.instrument === pairFilter) && (outFilter === "all" || t.outcome === outFilter));

  return (
    <div>
      <Section T={T} fs={fs} title="Journal" sub={`${trades.length} ${trades.length === 1 ? "entry" : "entries"}`}>
        <Label T={T} fs={fs}>Filter pair</Label>
        <div style={{ display: "flex", gap: 8, overflowX: "auto", paddingBottom: 4 }}>
          {pairs.map((p) => { const on = pairFilter === p; return <button key={p} className="fxbtn" onClick={() => { buzz(); setPairFilter(p); }} style={{ flex: "0 0 auto", minHeight: 42, padding: "0 14px", borderRadius: 10, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, background: on ? T.primary : T.chipOff, color: on ? T.primaryText : T.chipOffText, border: `2px solid ${on ? T.primary : T.border}` }}>{p}</button>; })}
        </div>
        <Label T={T} fs={fs}>Filter outcome</Label>
        <Seg T={T} fs={fs} buzz={buzz} value={outFilter === "all" ? "" : outFilter} onChange={(v) => setOutFilter(v || "all")} options={[{ k: "win", label: "Win", color: T.buy, textColor: T.buyText }, { k: "loss", label: "Loss", color: T.sell, textColor: T.sellText }, { k: "be", label: "B/E", color: T.neutral, textColor: T.neutralText }]} />
      </Section>

      {!loaded ? <p style={{ textAlign: "center", color: T.muted }}>Loading…</p>
        : filtered.length === 0 ? (
          <Section T={T} fs={fs}><div style={{ textAlign: "center", padding: 14 }}><p style={{ fontWeight: 700, marginBottom: 12 }}>{trades.length === 0 ? "No trades logged yet." : "No trades match these filters."}</p>{trades.length === 0 && <button className="fxbtn" onClick={() => { buzz(); goLog(); }} style={{ minHeight: 52, padding: "0 20px", borderRadius: 12, border: "none", background: T.primary, color: T.primaryText, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.md }}>+ Log your first trade</button>}</div></Section>
        ) : filtered.map((t) => <TradeRow key={t.id} {...{ T, fs, cur, trade: t, confirm, setConfirm, deleteTrade, editTrade, buzz }} />)}
    </div>
  );
}
function TradeRow({ T, fs, cur, trade, confirm, setConfirm, deleteTrade, editTrade, buzz }) {
  const long = trade.direction === "long";
  const dc = long ? T.buy : T.sell, dtc = long ? T.buyText : T.sellText;
  const pc = trade.pnl > 0 ? T.buy : trade.pnl < 0 ? T.sell : T.text;
  const confirming = confirm === trade.id;
  const tags = trade.mistakeTags || [];
  return (
    <div style={{ background: T.surface, border: `2px solid ${T.border}`, borderRadius: 12, padding: 12, marginBottom: 10 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.md }}>{trade.instrument}</div>
          <div style={{ color: T.muted, fontSize: fs.xs, fontWeight: 700 }}>{trade.date}{trade.session ? " · " + trade.session : ""}{setupText(trade) ? " · " + setupText(trade) : ""}</div>
          {(trade.propFirm || trade.customFirm) && <div style={{ marginTop: 2, display: "inline-block", padding: "1px 8px", borderRadius: 6, background: T.primary, color: T.primaryText, fontSize: fs.xs, fontWeight: 800 }}>{trade.customFirm || trade.propFirm}</div>}
        </div>
        <span style={{ display: "flex", alignItems: "center", gap: 4, padding: "4px 9px", borderRadius: 8, background: dc, color: dtc, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.xs }}>{long ? <TrendingUp size={Math.round(fs.xs)} strokeWidth={3} /> : <TrendingDown size={Math.round(fs.xs)} strokeWidth={3} />}{long ? "LONG" : "SHORT"}</span>
      </div>
      <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", marginTop: 8, gap: 10 }}>
        <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.display, color: pc, lineHeight: 1 }}>{fmtMoney(trade.pnl, cur)}</div>
        <div style={{ textAlign: "right", color: T.muted, fontSize: fs.xs, fontWeight: 700 }}>
          {trade.rr != null && <div>R:R 1:{Number(trade.rr).toFixed(2)}</div>}
          {trade.lots && <div>{trade.lots} lots</div>}
          <span style={{ display: "inline-block", marginTop: 3, padding: "1px 7px", borderRadius: 7, border: `2px solid ${T.border}`, color: T.text }}>{trade.outcome === "win" ? "Win" : trade.outcome === "loss" ? "Loss" : "B/E"}</span>
        </div>
      </div>
      {(preEmotionsOf(trade).length > 0 || trade.followedPlan) && <div style={{ marginTop: 8, fontSize: fs.xs, color: T.muted }}>{preEmotionsOf(trade).length > 0 && <>Felt: {preEmotionsOf(trade).join(", ")}</>}{preEmotionsOf(trade).length > 0 && trade.followedPlan && " · "}{trade.followedPlan && <>Plan: {trade.followedPlan === "yes" ? "followed" : "broke"}</>}</div>}
      {tags.length > 0 && <div style={{ marginTop: 6, display: "flex", flexWrap: "wrap", gap: 5 }}>{tags.map((tg) => <span key={tg} style={{ fontSize: fs.xs, padding: "1px 7px", borderRadius: 7, background: T.surfaceAlt, border: `1px solid ${T.sell}`, color: T.text }}>{tg}</span>)}</div>}
      {trade.reflection && <p style={{ marginTop: 8, marginBottom: 0, fontSize: fs.sm }}>{trade.reflection}</p>}
      <div style={{ marginTop: 10, display: "flex", justifyContent: "flex-end", gap: 8 }}>
        {confirming ? (<>
          <button className="fxbtn" onClick={() => { deleteTrade(trade.id); setConfirm(null); }} style={{ minHeight: 44, padding: "0 14px", borderRadius: 10, border: "none", background: T.sell, color: T.sellText, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>Yes, delete</button>
          <button className="fxbtn" onClick={() => setConfirm(null)} style={{ minHeight: 44, padding: "0 14px", borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>Keep</button>
        </>) : (
          <>
            <button className="fxbtn" onClick={() => editTrade(trade)} aria-label={`Edit ${trade.instrument} trade`} style={{ minHeight: 44, padding: "0 16px", borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, display: "flex", alignItems: "center", gap: 6 }}><PenLine size={Math.round(fs.sm)} strokeWidth={2.5} /> Edit</button>
            <button className="fxbtn" onClick={() => { buzz(); setConfirm(trade.id); }} aria-label={`Delete ${trade.instrument} trade`} style={{ minHeight: 44, width: 48, borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, display: "grid", placeItems: "center" }}><Trash2 size={Math.round(fs.md)} strokeWidth={2.3} /></button>
          </>
        )}
      </div>
    </div>
  );
}

/* ===================== analytics helpers ===================== */
function analyse(trades) {
  const wins = trades.filter((t) => t.outcome === "win");
  const losses = trades.filter((t) => t.outcome === "loss");
  const be = trades.filter((t) => t.outcome === "be");
  const decided = wins.length + losses.length;
  const net = trades.reduce((a, t) => a + (t.pnl || 0), 0);
  const grossWin = wins.reduce((a, t) => a + (t.pnl || 0), 0);
  const grossLoss = Math.abs(losses.reduce((a, t) => a + (t.pnl || 0), 0));
  const wr = decided ? (wins.length / decided) * 100 : null;
  const pf = grossLoss ? grossWin / grossLoss : grossWin > 0 ? Infinity : null;
  const avgWin = wins.length ? grossWin / wins.length : null;
  const avgLoss = losses.length ? -grossLoss / losses.length : null;
  const expectancy = trades.length ? net / trades.length : null;
  // best win streak (chronological)
  const chrono = [...trades].reverse();
  let streak = 0, best = 0;
  chrono.forEach((t) => { if (t.outcome === "win") { streak++; best = Math.max(best, streak); } else streak = 0; });
  return { wins, losses, be, decided, net, grossWin, grossLoss, wr, pf, avgWin, avgLoss, expectancy, best };
}

/* ===================== insight engine ===================== */
// Tags trades with streak context: was this trade entered after N losses/wins in a row?
function withStreakContext(trades) {
  const chrono = [...trades].reverse(); // oldest first
  let lossRun = 0, winRun = 0;
  const out = chrono.map((t) => {
    const enriched = { ...t, afterLossStreak: lossRun, afterWinStreak: winRun };
    if (t.outcome === "win") { winRun++; lossRun = 0; }
    else if (t.outcome === "loss") { lossRun++; winRun = 0; }
    else { /* B/E doesn't reset */ }
    return enriched;
  });
  return out; // oldest first
}

// Builds a ranked list of plain-language insights, each with the chart data behind it.
function buildInsights(trades, cur) {
  const insights = [];
  const enriched = withStreakContext(trades);
  const sum = (arr) => arr.reduce((a, t) => a + (t.pnl || 0), 0);
  const wrOf = (arr) => { const d = arr.filter((t) => t.outcome === "win" || t.outcome === "loss"); return d.length ? (arr.filter((t) => t.outcome === "win").length / d.length) * 100 : null; };

  // 1. Mistake tags by total $ cost
  MISTAKES.forEach((m) => {
    const ts = trades.filter((t) => (t.mistakeTags || []).includes(m));
    if (ts.length >= 1) {
      const cost = sum(ts);
      if (cost < 0) insights.push({ id: "mis-" + m, severity: -cost, type: "mistake",
        headline: `"${m}" cost you ${fmtMoney(cost, cur)}`,
        detail: `Tagged on ${ts.length} ${ts.length === 1 ? "trade" : "trades"}. Cutting this out is your clearest path to a higher P&L.`,
        chart: { kind: "bars", data: [{ label: "With " + m, value: cost / ts.length, color: "sell" }, { label: "Your average", value: trades.length ? sum(trades) / trades.length : 0, color: "neutral" }], format: "money" } });
    }
  });

  // 2. Pre-trade emotions by total $ impact (both leaks and edges)
  EMOTIONS.forEach((e) => {
    const ts = trades.filter((t) => preEmotionsOf(t).includes(e));
    if (ts.length >= 2) {
      const pnl = sum(ts); const wr = wrOf(ts);
      if (pnl < 0) insights.push({ id: "emo-" + e, severity: -pnl, type: "emotion",
        headline: `Trading while ${e} cost you ${fmtMoney(pnl, cur)}`,
        detail: `${ts.length} trades, ${wr == null ? "—" : Math.round(wr) + "%"} win rate. Consider skipping entries when you feel ${e.toLowerCase()}.`,
        chart: { kind: "bars", data: emotionPnlData(trades), format: "money", highlight: e } });
      else if (pnl > 0 && wr != null && wr >= 55) insights.push({ id: "emo-edge-" + e, severity: pnl * 0.5, type: "edge",
        headline: `${e} is your edge: ${fmtMoney(pnl, cur)} across ${ts.length} trades`,
        detail: `${Math.round(wr)}% win rate when you feel ${e.toLowerCase()}. This is your A-game state — trade more like this.`,
        chart: { kind: "bars", data: emotionPnlData(trades), format: "money", highlight: e } });
    }
  });

  // 3. Plan adherence
  const on = trades.filter((t) => t.followedPlan === "yes");
  const off = trades.filter((t) => t.followedPlan === "no");
  if (on.length >= 1 && off.length >= 1) {
    const wrOn = wrOf(on), wrOff = wrOf(off);
    const gap = (wrOn || 0) - (wrOff || 0);
    insights.push({ id: "plan", severity: Math.abs(sum(off)) + Math.abs(gap) * 5, type: gap >= 0 ? "edge" : "warn",
      headline: gap >= 0 ? `Discipline is your edge: ${Math.round(wrOn)}% on-plan vs ${Math.round(wrOff)}% off-plan` : `Off-plan trades win more (${Math.round(wrOff)}%) — review your plan`,
      detail: `Off-plan trading has produced ${fmtMoney(sum(off), cur)} across ${off.length} trades.`,
      chart: { kind: "bars", data: [{ label: "On-plan", value: wrOn || 0, color: "buy" }, { label: "Off-plan", value: wrOff || 0, color: "sell" }], format: "pct" } });
  }

  // 4. Revenge after losing streaks (streak emotional impact)
  const afterLosses = enriched.filter((t) => t.afterLossStreak >= 2);
  if (afterLosses.length >= 2) {
    const pnl = sum(afterLosses); const wr = wrOf(afterLosses); const baseWr = wrOf(trades);
    insights.push({ id: "streak-loss", severity: Math.abs(pnl) + (baseWr && wr != null ? Math.max(0, baseWr - wr) * 5 : 0), type: pnl < 0 ? "warn" : "info",
      headline: pnl < 0 ? `Trading after 2+ losses cost you ${fmtMoney(pnl, cur)}` : `You hold up after losing streaks (${fmtMoney(pnl, cur)})`,
      detail: `${afterLosses.length} trades taken on a losing streak, ${wr == null ? "—" : Math.round(wr) + "%"} win rate vs ${baseWr == null ? "—" : Math.round(baseWr) + "%"} overall. ${pnl < 0 ? "A short break after 2 losses may protect your account." : ""}`,
      chart: { kind: "bars", data: [{ label: "After 2+ losses", value: wr || 0, color: "sell" }, { label: "Overall", value: baseWr || 0, color: "neutral" }], format: "pct" } });
  }

  // 5. Post-trade mood vs P&L (confirms emotional regulation)
  const negMoodTrades = trades.filter((t) => postMoodsOf(t).some((m) => !POSITIVE_MOODS.includes(m)));
  if (negMoodTrades.length >= 3) {
    const angerTrades = trades.filter((t) => postMoodsOf(t).some((m) => m === "Angry" || m === "Frustrated"));
    if (angerTrades.length >= 2) insights.push({ id: "mood", severity: Math.abs(sum(angerTrades)) + 1, type: "info",
      headline: `You felt angry or frustrated after ${angerTrades.length} trades`,
      detail: `Those trades netted ${fmtMoney(sum(angerTrades), cur)}. Strong negative mood often precedes revenge trading — watch the next entry.`,
      chart: { kind: "bars", data: moodPnlData(trades), format: "money", highlight: "Angry" } });
  }

  insights.sort((a, b) => b.severity - a.severity);
  return insights;
}
function emotionPnlData(trades) {
  return EMOTIONS.map((e) => { const ts = trades.filter((t) => preEmotionsOf(t).includes(e)); return { label: e, value: ts.reduce((a, t) => a + (t.pnl || 0), 0), n: ts.length }; }).filter((x) => x.n > 0);
}
function moodPnlData(trades) {
  return MOODS.map((m) => { const ts = trades.filter((t) => postMoodsOf(t).includes(m)); return { label: m, value: ts.reduce((a, t) => a + (t.pnl || 0), 0), n: ts.length }; }).filter((x) => x.n > 0);
}

/* Horizontal bar chart (diverging for money, simple for pct) */
function MiniBars({ data, format, highlight, T, fs, cur }) {
  if (!data || data.length === 0) return <span style={{ color: T.muted, fontSize: fs.xs }}>Not enough data yet.</span>;
  const vals = data.map((d) => d.value);
  const maxAbs = Math.max(1, ...vals.map((v) => Math.abs(v)));
  const colorOf = (d) => {
    if (d.color === "buy") return T.buy; if (d.color === "sell") return T.sell; if (d.color === "neutral") return T.neutral;
    if (format === "pct") return d.value >= 50 ? T.buy : T.sell;
    return d.value >= 0 ? T.buy : T.sell;
  };
  const fmt = (v) => format === "pct" ? Math.round(v) + "%" : fmtMoney(v, cur);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 7, marginTop: 4 }}>
      {data.map((d) => {
        const pct = (Math.abs(d.value) / maxAbs) * 100;
        const hot = highlight && d.label === highlight;
        return (
          <div key={d.label} style={{ display: "grid", gridTemplateColumns: "92px 1fr auto", gap: 8, alignItems: "center" }}>
            <span style={{ fontSize: fs.xs, fontWeight: hot ? 800 : 600, color: hot ? T.text : T.muted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{d.label}</span>
            <div style={{ height: 18, background: T.surfaceAlt, borderRadius: 5, overflow: "hidden", border: hot ? `2px solid ${colorOf(d)}` : "none" }}>
              <div style={{ width: `${pct}%`, height: "100%", background: colorOf(d), opacity: hot ? 1 : 0.85 }} />
            </div>
            <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.xs, color: colorOf(d), textAlign: "right" }}>{fmt(d.value)}</span>
          </div>
        );
      })}
    </div>
  );
}

/* ===================== PSYCH (analysis centrepiece) ===================== */
function PsychView({ T, fs, cur, trades, selectedFirm }) {
  const A = analyse(trades);
  const insights = buildInsights(trades, cur);
  const avg = (arr, f) => { const xs = arr.map(f).filter((x) => typeof x === "number" && !isNaN(x)); return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null; };
  const discTrades = trades.filter((t) => t.followedPlan);
  const discipline = discTrades.length ? (trades.filter((t) => t.followedPlan === "yes").length / discTrades.length) * 10 : null;
  const confAvg = avg(trades, (t) => t.confidence);
  const eqAvg = avg(trades, (t) => t.entryQuality);

  if (trades.length === 0) return <EmptyNote T={T} fs={fs} text="Log trades to unlock your psychological analysis. The more you log, the sharper the insights." />;

  const typeStyle = (type) => {
    if (type === "edge") return { accent: T.buy, tag: "STRENGTH" };
    if (type === "warn") return { accent: T.sell, tag: "FIX THIS" };
    if (type === "mistake") return { accent: T.sell, tag: "COSTLY LEAK" };
    if (type === "emotion") return { accent: T.sell, tag: "EMOTIONAL LEAK" };
    return { accent: T.primary, tag: "WATCH" };
  };

  const headlineLeak = insights.find((i) => i.type === "mistake" || i.type === "emotion" || i.type === "warn");

  return (
    <div>
      <Section T={T} fs={fs} title="Psychology analysis" sub="Ranked by how much each pattern costs or makes you" accent={T.primary}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
          <Mini T={T} fs={fs} label="Discipline" v={discipline == null ? "—" : discipline.toFixed(1)} suffix="/10" />
          <Mini T={T} fs={fs} label="Avg confidence" v={confAvg == null ? "—" : confAvg.toFixed(1)} suffix="/10" />
          <Mini T={T} fs={fs} label="Entry quality" v={eqAvg == null ? "—" : eqAvg.toFixed(1)} suffix="/10" />
        </div>
        {headlineLeak && (
          <div style={{ marginTop: 12, padding: 12, borderRadius: 10, background: T.surfaceAlt, border: `2px solid ${T.sell}` }}>
            <div style={{ fontSize: fs.xs, fontWeight: 800, color: T.sell, letterSpacing: ".05em" }}>👉 YOUR #1 THING TO FIX</div>
            <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.md, marginTop: 4 }}>{headlineLeak.headline}</div>
          </div>
        )}
      </Section>

      {insights.length === 0 ? (
        <Section T={T} fs={fs}><p style={{ color: T.muted, textAlign: "center", margin: 0, padding: 8 }}>No strong patterns yet — keep logging emotions, mistake tags and whether you followed your plan. Insights appear as the data builds.</p></Section>
      ) : insights.map((ins, i) => {
        const st = typeStyle(ins.type);
        return (
          <div key={ins.id} style={{ background: T.surface, border: `2px solid ${T.border}`, borderLeft: `6px solid ${st.accent}`, borderRadius: 12, padding: 13, marginBottom: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.xs, color: st.accent, letterSpacing: ".05em" }}>{st.tag}</span>
              {i === 0 && <span style={{ fontSize: fs.xs, fontWeight: 700, color: T.muted }}>· highest impact</span>}
            </div>
            <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.md, lineHeight: 1.25 }}>{ins.headline}</div>
            <p style={{ margin: "6px 0 10px", fontSize: fs.sm, color: T.muted }}>{ins.detail}</p>
            <MiniBars data={ins.chart.data} format={ins.chart.format} highlight={ins.chart.highlight} T={T} fs={fs} cur={cur} />
          </div>
        );
      })}

      {/* Reference charts always available */}
      <Section T={T} fs={fs} title="P&L by pre-trade emotion">
        <MiniBars data={emotionPnlData(trades)} format="money" T={T} fs={fs} cur={cur} />
      </Section>
      {moodPnlData(trades).length > 0 && (
        <Section T={T} fs={fs} title="P&L by post-trade mood">
          <MiniBars data={moodPnlData(trades)} format="money" T={T} fs={fs} cur={cur} />
        </Section>
      )}
      <Section T={T} fs={fs} title="Mistake tag cost">
        {(() => {
          const rows = MISTAKES.map((m) => { const ts = trades.filter((t) => (t.mistakeTags || []).includes(m)); return { label: m, value: ts.reduce((a, t) => a + (t.pnl || 0), 0), n: ts.length }; }).filter((r) => r.n > 0).sort((a, b) => a.value - b.value);
          return rows.length ? <MiniBars data={rows} format="money" T={T} fs={fs} cur={cur} /> : <span style={{ color: T.muted, fontSize: fs.sm }}>No mistakes tagged — nice.</span>;
        })()}
      </Section>
    </div>
  );
}
function Mini({ T, fs, label, v, suffix }) {
  return (
    <div style={{ background: T.surfaceAlt, border: `2px solid ${T.border}`, borderRadius: 10, padding: "10px 8px", textAlign: "center" }}>
      <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg, color: T.text }}>{v}<span style={{ fontSize: fs.xs, color: T.muted }}>{v === "—" ? "" : suffix}</span></div>
      <div style={{ fontSize: fs.xs, color: T.muted, fontWeight: 700, marginTop: 2 }}>{label}</div>
    </div>
  );
}
/* ===================== STATS ===================== */
function StatsView({ T, fs, cur, trades, allTrades, flash, selectedFirm }) {
  const A = analyse(trades);
  const fileInputRef = React.useRef(null);

  const exportCSV = () => {
    const cols = ["date", "propFirm", "instrument", "direction", "session", "setup", "entry", "stopLoss", "takeProfit", "lots", "rr", "pnl", "outcome", "hitTP", "movedSL", "preEmotion", "confidence", "entryQuality", "mistakeTags", "followedPlan", "postMood", "reflection"];
    const firmName = (t) => t.customFirm && t.customFirm.trim() ? t.customFirm.trim() : (t.propFirm || "");
    const valOf = (t, c) => (c === "setup" ? setupText(t) : c === "preEmotion" ? preEmotionsOf(t) : c === "postMood" ? postMoodsOf(t) : c === "propFirm" ? firmName(t) : t[c]);
    const cell = (v) => { if (v == null) return ""; const s = Array.isArray(v) ? v.join("; ") : String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const csv = [cols.join(","), ...trades.map((t) => cols.map((c) => cell(valOf(t, c))).join(","))].join("\n");
    download("fx-journal.csv", csv, "text/csv;charset=utf-8;");
    flash("CSV downloaded");
  };
  const exportXLSX = () => {
    try {
      const rows = trades.map((t) => ({ ...t, setup: setupText(t), setups: (t.setups || []).join("; "), mistakeTags: (t.mistakeTags || []).join("; "), preEmotion: preEmotionsOf(t).join("; "), postMood: postMoodsOf(t).join("; "), preEmotions: undefined, postMoods: undefined }));
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "Trades");
      XLSX.writeFile(wb, "fx-journal.xlsx"); flash("Excel downloaded");
    } catch (e) { exportCSV(); }
  };

  const importFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        let rows = [];
        const name = file.name.toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
          const wb = XLSX.read(evt.target.result, { type: "binary" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        } else {
          // CSV parsing
          const text = evt.target.result;
          const lines = text.split(/\r?\n/).filter(Boolean);
          const headers = lines[0].split(",").map(h => h.trim().replace(/^"|"$/g, ""));
          rows = lines.slice(1).map(line => {
            const vals = [];
            let cur = "", inQ = false;
            for (let c of line) {
              if (c === '"') { inQ = !inQ; }
              else if (c === "," && !inQ) { vals.push(cur); cur = ""; }
              else cur += c;
            }
            vals.push(cur);
            const obj = {};
            headers.forEach((h, i) => { obj[h] = (vals[i] || "").replace(/^"|"$/g, "").trim(); });
            return obj;
          });
        }
        // Convert rows to trade objects
        const imported = rows.map(r => ({
          id: Date.now() + "-" + Math.random().toString(36).slice(2),
          createdAt: new Date().toISOString(),
          date: r.date || todayISO(),
          propFirm: r.propFirm || "",
          customFirm: "",
          instrument: r.instrument || "",
          direction: r.direction || "",
          session: r.session || "",
          setups: r.setup ? r.setup.split(";").map(s => s.trim()).filter(Boolean) : [],
          entry: r.entry || "",
          stopLoss: r.stopLoss || "",
          takeProfit: r.takeProfit || "",
          lots: r.lots || "",
          rr: r.rr ? parseFloat(r.rr) : null,
          pnl: r.pnl ? parseFloat(r.pnl) : 0,
          outcome: r.outcome || (r.pnl > 0 ? "win" : r.pnl < 0 ? "loss" : "be"),
          hitTP: r.hitTP || "",
          movedSL: r.movedSL || "",
          preEmotions: r.preEmotion ? r.preEmotion.split(";").map(s => s.trim()).filter(Boolean) : [],
          confidence: r.confidence ? parseInt(r.confidence) : 5,
          entryQuality: r.entryQuality ? parseInt(r.entryQuality) : 5,
          mistakeTags: r.mistakeTags ? r.mistakeTags.split(";").map(s => s.trim()).filter(Boolean) : [],
          followedPlan: r.followedPlan || "",
          postMoods: r.postMood ? r.postMood.split(";").map(s => s.trim()).filter(Boolean) : [],
          reflection: r.reflection || "",
        })).filter(r => r.instrument || r.pnl);

        if (imported.length === 0) { flash("No valid trades found in file"); return; }

        // Merge: skip duplicates based on date+instrument+pnl
        const existing = allTrades || trades;
        const isDup = (t) => existing.some(e => e.date === t.date && e.instrument === t.instrument && String(e.pnl) === String(t.pnl));
        const newOnes = imported.filter(t => !isDup(t));
        const merged = [...newOnes, ...existing].sort((a, b) => b.date.localeCompare(a.date));
        saveKey(TRADES_KEY, merged);
        flash(`Imported ${newOnes.length} trades (${imported.length - newOnes.length} duplicates skipped) — refresh to see`);
        setTimeout(() => window.location.reload(), 1800);
      } catch (err) {
        flash("Import failed — check file format");
        console.error(err);
      }
    };
    if (file.name.toLowerCase().endsWith(".csv")) reader.readAsText(file);
    else reader.readAsBinaryString(file);
    e.target.value = "";
  };
  const stat = (label, val, color, sub) => (
    <div style={{ background: T.surface, border: `2px solid ${T.border}`, borderRadius: 12, padding: 12 }}>
      <div style={{ fontSize: fs.xs, color: T.muted, fontWeight: 800, textTransform: "uppercase", letterSpacing: ".04em" }}>{label}</div>
      <div style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.xl, color: color || T.text, lineHeight: 1.1, marginTop: 2 }}>{val}</div>
      {sub && <div style={{ fontSize: fs.xs, color: T.muted, marginTop: 2 }}>{sub}</div>}
    </div>
  );
  const pairRows = Array.from(new Set(trades.map((t) => t.instrument))).map((p) => { const ts = trades.filter((t) => t.instrument === p); const d = ts.filter((t) => t.outcome === "win" || t.outcome === "loss"); const wr = d.length ? Math.round((ts.filter((t) => t.outcome === "win").length / d.length) * 100) : null; const pnl = ts.reduce((a, t) => a + (t.pnl || 0), 0); return { p, n: ts.length, wr, pnl }; }).sort((a, b) => b.pnl - a.pnl);

  if (trades.length === 0) return (
    <div style={{ padding: 12 }}>
      <Section T={T} fs={fs} title="Import Trades" sub="Upload a previously exported CSV or Excel file">
        <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={importFile} style={{ display: "none" }} aria-label="Import trades file" />
        <button className="fxbtn" onClick={() => fileInputRef.current && fileInputRef.current.click()} style={{ width: "100%", minHeight: 56, borderRadius: 10, border: `2px solid ${T.primary}`, background: T.surface, color: T.primary, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.md, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
          ⬆ Import Trades (CSV or Excel)
        </button>
        <p style={{ fontSize: fs.xs, color: T.muted, marginTop: 8, textAlign: "center" }}>Import a file you previously exported — duplicates are skipped automatically</p>
      </Section>
      <EmptyNote T={T} fs={fs} text="No trades yet — log your first trade or import a file above." />
    </div>
  );
  return (
    <div>
      <Section T={T} fs={fs} title="Analytics" sub="Performance at a glance">
        <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
          <button className="fxbtn" onClick={exportCSV} style={{ flex: 1, minHeight: 48, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surfaceAlt, color: T.text, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Download size={Math.round(fs.md)} strokeWidth={2.4} /> CSV</button>
          <button className="fxbtn" onClick={exportXLSX} style={{ flex: 1, minHeight: 48, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surfaceAlt, color: T.text, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}><Download size={Math.round(fs.md)} strokeWidth={2.4} /> Excel</button>
        </div>
        <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={importFile} style={{ display: "none" }} aria-label="Import trades file" />
        <button className="fxbtn" onClick={() => fileInputRef.current && fileInputRef.current.click()} style={{ width: "100%", minHeight: 48, borderRadius: 10, border: `2px solid ${T.primary}`, background: T.surface, color: T.primary, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
          ⬆ Import Trades (CSV or Excel)
        </button>
        <p style={{ fontSize: fs.xs, color: T.muted, marginTop: 6, textAlign: "center" }}>Import a file you previously exported — duplicates are skipped automatically</p>
      </Section>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
        {stat("Total trades", String(trades.length))}
        {stat("Win rate", A.wr == null ? "—" : Math.round(A.wr) + "%")}
        {stat("Net P&L", fmtMoney(A.net, cur), A.net > 0 ? T.buy : A.net < 0 ? T.sell : T.text)}
        {stat("Profit factor", A.pf == null ? "—" : A.pf === Infinity ? "∞" : A.pf.toFixed(2), undefined, "gross win / gross loss")}
        {stat("Avg win", A.avgWin == null ? "—" : fmtMoney(A.avgWin, cur), T.buy)}
        {stat("Avg loss", A.avgLoss == null ? "—" : fmtMoney(A.avgLoss, cur), T.sell)}
        {stat("Expectancy", A.expectancy == null ? "—" : fmtMoney(A.expectancy, cur), A.expectancy >= 0 ? T.buy : T.sell, "avg $ per trade")}
        {stat("Best streak", A.best ? A.best + " W" : "—", T.buy, "consecutive wins")}
      </div>

      <Section T={T} fs={fs} title="Outcome breakdown">
        <div style={{ display: "flex", height: 16, borderRadius: 8, overflow: "hidden", border: `2px solid ${T.border}` }}>
          <div style={{ width: `${(A.wins.length / trades.length) * 100}%`, background: T.buy }} />
          <div style={{ width: `${(A.losses.length / trades.length) * 100}%`, background: T.sell }} />
          <div style={{ width: `${(A.be.length / trades.length) * 100}%`, background: T.neutral }} />
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8, fontSize: fs.sm, fontWeight: 700 }}>
          <span style={{ color: T.buy }}>{A.wins.length} Wins</span><span style={{ color: T.sell }}>{A.losses.length} Losses</span><span style={{ color: T.muted }}>{A.be.length} B/E</span>
        </div>
      </Section>

      <Section T={T} fs={fs} title="By instrument">
        <div style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 8, fontSize: fs.xs, color: T.muted, fontWeight: 800, paddingBottom: 6, borderBottom: `2px solid ${T.border}` }}><span>Pair</span><span>Trades</span><span>WR</span><span>P&L</span></div>
        {pairRows.map((r) => <div key={r.p} style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 8, alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${T.border}`, fontSize: fs.sm }}><span style={{ fontWeight: 700 }}>{r.p}</span><span style={{ textAlign: "right" }}>{r.n}</span><span style={{ textAlign: "right" }}>{r.wr == null ? "—" : r.wr + "%"}</span><span style={{ textAlign: "right", fontFamily: FONT_DISPLAY, fontWeight: 800, color: r.pnl > 0 ? T.buy : r.pnl < 0 ? T.sell : T.text }}>{fmtMoney(r.pnl, cur)}</span></div>)}
      </Section>

      {selectedFirm === "All" && (() => {
        const firmName = (t) => t.customFirm && t.customFirm.trim() ? t.customFirm.trim() : (t.propFirm || "Unassigned");
        const firms = Array.from(new Set((allTrades || trades).map(firmName))).sort();
        const firmRows = firms.map((f) => {
          const ts = (allTrades || trades).filter((t) => firmName(t) === f);
          const d = ts.filter((t) => t.outcome === "win" || t.outcome === "loss");
          const wr = d.length ? Math.round((ts.filter((t) => t.outcome === "win").length / d.length) * 100) : null;
          const pnl = ts.reduce((a, t) => a + (t.pnl || 0), 0);
          return { f, n: ts.length, wr, pnl };
        }).sort((a, b) => b.pnl - a.pnl);
        return (
          <Section T={T} fs={fs} title="By prop firm" sub="Tap a firm in the bar above to filter everything">
            <div style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 8, fontSize: fs.xs, color: T.muted, fontWeight: 800, paddingBottom: 6, borderBottom: `2px solid ${T.border}` }}><span>Firm</span><span>Trades</span><span>WR</span><span>P&L</span></div>
            {firmRows.map((r) => <div key={r.f} style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 8, alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${T.border}`, fontSize: fs.sm }}><span style={{ fontWeight: 700 }}>{r.f}</span><span style={{ textAlign: "right" }}>{r.n}</span><span style={{ textAlign: "right" }}>{r.wr == null ? "—" : r.wr + "%"}</span><span style={{ textAlign: "right", fontFamily: FONT_DISPLAY, fontWeight: 800, color: r.pnl > 0 ? T.buy : r.pnl < 0 ? T.sell : T.text }}>{fmtMoney(r.pnl, cur)}</span></div>)}
          </Section>
        );
      })()}
    </div>
  );
}

/* ===================== REVIEW ===================== */
function ReviewView({ T, fs, reviewDraft, setReviewDraft, saveReview, reviews, deleteReview, buzz, flash }) {
  const set = (patch) => setReviewDraft((d) => ({ ...d, ...patch }));
  const inputStyle = { width: "100%", minHeight: 50, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surface, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, padding: "0 12px" };
  const ta = (key, ph) => (
    <div style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 4 }}>
      <textarea value={reviewDraft[key]} onChange={(e) => set({ [key]: e.target.value })} rows={2} placeholder={ph} style={{ flex: 1, resize: "vertical", minHeight: 52, borderRadius: 10, border: `2px solid ${T.border}`, background: T.surface, color: T.text, fontFamily: FONT_BODY, fontSize: fs.body, padding: 10 }} />
      <VoiceBtn T={T} fs={fs} mode="text" buzz={buzz} flash={flash} onResult={(txt) => set({ [key]: (reviewDraft[key] ? reviewDraft[key] + " " : "") + txt })} />
    </div>
  );
  return (
    <div>
      <Section T={T} fs={fs} title="Review" sub="Daily session debrief">
        <Label T={T} fs={fs}>Date</Label>
        <input type="date" value={reviewDraft.date} onChange={(e) => set({ date: e.target.value })} style={inputStyle} aria-label="Review date" />
        <Label T={T} fs={fs}>Session rating</Label>
        <RatingStepper value={reviewDraft.rating} onChange={(v) => set({ rating: v })} labels={["Terrible", "Okay", "Perfect"]} color={T.primary} T={T} fs={fs} buzz={buzz} />
        <Label T={T} fs={fs}>Market conditions</Label>{ta("conditions", "Trending? Choppy? News?")}
        <Label T={T} fs={fs}>What went well?</Label>{ta("wentWell", "")}
        <Label T={T} fs={fs}>What went wrong?</Label>{ta("wentWrong", "")}
        <Label T={T} fs={fs}>Key lesson for tomorrow</Label>{ta("lesson", "")}
        <button className="fxbtn" onClick={saveReview} style={{ width: "100%", minHeight: 56, borderRadius: 12, border: "none", marginTop: 8, background: T.primary, color: T.primaryText, fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.md, display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}><Check size={Math.round(fs.md)} strokeWidth={3} /> Save session review</button>
      </Section>

      {reviews.length > 0 && <Label T={T} fs={fs}>Past reviews</Label>}
      {reviews.map((r) => (
        <div key={r.id} style={{ background: T.surface, border: `2px solid ${T.border}`, borderRadius: 12, padding: 12, marginBottom: 10 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.md }}>{r.date}</span>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, color: T.primary }}>{r.rating}/10</span>
              <button className="fxbtn" onClick={() => deleteReview(r.id)} aria-label="Delete review" style={{ width: 40, height: 40, borderRadius: 9, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, display: "grid", placeItems: "center" }}><Trash2 size={Math.round(fs.sm)} strokeWidth={2.3} /></button>
            </div>
          </div>
          {r.conditions && <p style={{ margin: "6px 0 0", fontSize: fs.sm }}><b>Conditions:</b> {r.conditions}</p>}
          {r.wentWell && <p style={{ margin: "4px 0 0", fontSize: fs.sm, color: T.buy }}><b>Well:</b> {r.wentWell}</p>}
          {r.wentWrong && <p style={{ margin: "4px 0 0", fontSize: fs.sm, color: T.sell }}><b>Wrong:</b> {r.wentWrong}</p>}
          {r.lesson && <p style={{ margin: "4px 0 0", fontSize: fs.sm }}><b>Lesson:</b> {r.lesson}</p>}
        </div>
      ))}
    </div>
  );
}

function EmptyNote({ T, fs, text }) {
  return <Section T={T} fs={fs}><p style={{ textAlign: "center", color: T.muted, padding: 16, margin: 0 }}>{text}</p></Section>;
}

/* ===================== Settings sheet ===================== */
function SettingsSheet({ T, fs, settings, updateSettings, buzz, close, resetAll, count }) {
  const [confirmReset, setConfirmReset] = useState(false);
  const Toggle = ({ active, onClick, children }) => (
    <button className="fxbtn" onClick={() => { buzz(); onClick(); }} aria-pressed={active} style={{ flex: 1, minHeight: 50, borderRadius: 10, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, background: active ? T.primary : T.chipOff, color: active ? T.primaryText : T.chipOffText, border: `2px solid ${active ? T.primary : T.border}`, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>{children}</button>
  );
  return (
    <div role="dialog" aria-modal="true" aria-label="Settings" style={{ position: "fixed", inset: 0, zIndex: 80, background: "rgba(0,0,0,.55)", display: "flex", flexDirection: "column", justifyContent: "flex-end" }} onClick={(e) => { if (e.target === e.currentTarget) close(); }}>
      <div data-anim style={{ background: T.surface, borderTop: `3px solid ${T.border}`, borderTopLeftRadius: 22, borderTopRightRadius: 22, padding: "14px 14px calc(14px + env(safe-area-inset-bottom))", maxWidth: 560, width: "100%", margin: "0 auto", maxHeight: "88vh", overflowY: "auto", boxShadow: T.shadow, animation: "sheetUp .18s ease" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <span style={{ fontFamily: FONT_DISPLAY, fontWeight: 900, fontSize: fs.lg }}>Settings & accessibility</span>
          <button className="fxbtn" onClick={close} aria-label="Close" style={{ width: 44, height: 44, borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, display: "grid", placeItems: "center" }}><X size={Math.round(fs.lg)} strokeWidth={2.6} /></button>
        </div>

        <Label T={T} fs={fs}>Theme</Label>
        <div style={{ display: "flex", gap: 8 }}>
          <Toggle active={settings.theme === "daylight"} onClick={() => updateSettings({ theme: "daylight" })}><Sun size={Math.round(fs.md)} strokeWidth={2.5} /> Daylight</Toggle>
          <Toggle active={settings.theme === "dark"} onClick={() => updateSettings({ theme: "dark" })}><Moon size={Math.round(fs.md)} strokeWidth={2.5} /> Dark</Toggle>
        </div>
        <p style={{ color: T.muted, fontSize: fs.xs, margin: "8px 2px 0" }}>Daylight = maximum contrast for using your phone outdoors in bright light.</p>

        <Label T={T} fs={fs}>Text size</Label>
        <div style={{ display: "flex", gap: 8 }}>{["1", "1.2", "1.4"].map((s) => <Toggle key={s} active={settings.scale === s} onClick={() => updateSettings({ scale: s })}>{SCALE_LABELS[s]}</Toggle>)}</div>

        <Label T={T} fs={fs}>Currency</Label>
        <div style={{ display: "flex", gap: 8 }}>{CURRENCIES.map((c) => <Toggle key={c} active={settings.currency === c} onClick={() => updateSettings({ currency: c })}>{c}</Toggle>)}</div>

        <Label T={T} fs={fs}>Vibration feedback</Label>
        <button className="fxbtn" onClick={() => { buzz(30); updateSettings({ haptics: !settings.haptics }); }} aria-pressed={settings.haptics} style={{ width: "100%", minHeight: 50, borderRadius: 10, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm, display: "flex", alignItems: "center", justifyContent: "center", gap: 8, background: settings.haptics ? T.primary : T.chipOff, color: settings.haptics ? T.primaryText : T.chipOffText, border: `2px solid ${settings.haptics ? T.primary : T.border}` }}>{settings.haptics ? <CheckSquare size={Math.round(fs.md)} strokeWidth={2.4} /> : <Square size={Math.round(fs.md)} strokeWidth={2.4} />} Buzz on tap {settings.haptics ? "ON" : "OFF"}</button>

        <Label T={T} fs={fs}>Danger zone</Label>
        {confirmReset ? (
          <div style={{ display: "flex", gap: 8 }}>
            <button className="fxbtn" onClick={resetAll} style={{ flex: 1, minHeight: 50, borderRadius: 10, border: "none", background: T.sell, color: T.sellText, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>Delete all {count} items</button>
            <button className="fxbtn" onClick={() => setConfirmReset(false)} style={{ flex: 1, minHeight: 50, borderRadius: 10, background: T.surfaceAlt, color: T.text, border: `2px solid ${T.border}`, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>Cancel</button>
          </div>
        ) : (
          <button className="fxbtn" onClick={() => { buzz(); setConfirmReset(true); }} style={{ width: "100%", minHeight: 50, borderRadius: 10, background: T.surface, color: T.sell, border: `2px solid ${T.sell}`, fontFamily: FONT_DISPLAY, fontWeight: 800, fontSize: fs.sm }}>Clear all data</button>
        )}
        <p style={{ textAlign: "center", color: T.muted, fontSize: fs.xs, marginTop: 12 }}>Saved privately on this device. Not financial advice.</p>
      </div>
    </div>
  );
}

/* ---- download helper ---- */
function download(filename, text, mime) {
  try {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  } catch (e) {}
}
