import { useState, useMemo, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend, LineChart, Line, CartesianGrid } from "recharts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// ─── Supabase config — replace with your values ────────────────────────────
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
// ───────────────────────────────────────────────────────────────────────────

// ── Shared styles ────────────────────────────────────────────────────────────
const authPageStyle = {
  minHeight: "100vh", background: "#f1f5f9",
  display: "flex", alignItems: "center", justifyContent: "center",
  fontFamily: "'DM Sans', 'Segoe UI', sans-serif", padding: 16,
};
const authCardStyle = {
  background: "#fff", borderRadius: 20, padding: "40px 36px",
  boxShadow: "0 4px 32px rgba(0,0,0,0.08)", width: "100%", maxWidth: 420,
};
const inputStyle = {
  width: "100%", padding: "10px 14px", borderRadius: 10, fontSize: 14,
  border: "1px solid #cbd5e1", outline: "none", boxSizing: "border-box",
  fontFamily: "inherit", marginTop: 6,
};
const btnPrimaryStyle = {
  width: "100%", padding: "11px", borderRadius: 10, fontSize: 14,
  fontWeight: 600, background: "#2563eb", color: "#fff", border: "none",
  cursor: "pointer", fontFamily: "inherit", marginTop: 8,
};
const labelStyle = { fontSize: 13, fontWeight: 500, color: "#475569", display: "block", marginTop: 16 };
const errStyle = { fontSize: 13, color: "#ef4444", marginTop: 10, textAlign: "center" };
const okStyle  = { fontSize: 13, color: "#10b981", marginTop: 10, textAlign: "center" };

// ── Login Page ───────────────────────────────────────────────────────────────
function LoginPage() {
  const [email, setEmail]       = useState("");
  const [password, setPassword] = useState("");
  const [error, setError]       = useState("");
  const [loading, setLoading]   = useState(false);

  async function handleLogin() {
    setError(""); setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) setError(error.message);
    setLoading(false);
  }

  return (
    <div style={authPageStyle}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@600;700&display=swap');`}</style>
      <div style={authCardStyle}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 28 }}>
          <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg, #1d4ed8, #7c3aed)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>⚡</div>
          <div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 17, color: "#0f172a" }}>PC Sales Pipeline</div>
            <div style={{ fontSize: 12, color: "#64748b" }}>Sign in to your account</div>
          </div>
        </div>
        <label style={labelStyle}>Email</label>
        <input style={inputStyle} type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="you@company.com" onKeyDown={e => e.key === "Enter" && handleLogin()} />
        <label style={labelStyle}>Password</label>
        <input style={inputStyle} type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" onKeyDown={e => e.key === "Enter" && handleLogin()} />
        {error && <div style={errStyle}>{error}</div>}
        <button style={{ ...btnPrimaryStyle, opacity: loading ? 0.7 : 1 }} onClick={handleLogin} disabled={loading}>
          {loading ? "Signing in…" : "Sign in"}
        </button>
        <div style={{ fontSize: 12, color: "#94a3b8", textAlign: "center", marginTop: 20 }}>
          No account? Ask your admin to invite you.
        </div>
      </div>
    </div>
  );
}

// ── Set Password Page (invite redirect lands here) ───────────────────────────
function SetPasswordPage({ onDone }) {
  const [password, setPassword]   = useState("");
  const [confirm, setConfirm]     = useState("");
  const [error, setError]         = useState("");
  const [success, setSuccess]     = useState(false);
  const [loading, setLoading]     = useState(false);

  async function handleSet() {
    setError("");
    if (password.length < 8) { setError("Password must be at least 8 characters."); return; }
    if (password !== confirm)  { setError("Passwords do not match."); return; }
    setLoading(true);
    const { error } = await supabase.auth.updateUser({ password });
    if (error) { setError(error.message); setLoading(false); return; }
    setSuccess(true);
    setLoading(false);
    setTimeout(onDone, 1500);
  }

  return (
    <div style={authPageStyle}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@600;700&display=swap');`}</style>
      <div style={authCardStyle}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 20, color: "#0f172a", marginBottom: 8 }}>Set your password</div>
        <div style={{ fontSize: 13, color: "#64748b", marginBottom: 24 }}>Welcome! Create a password to activate your account.</div>
        <label style={labelStyle}>New Password</label>
        <input style={inputStyle} type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="Min. 8 characters" />
        <label style={labelStyle}>Confirm Password</label>
        <input style={inputStyle} type="password" value={confirm} onChange={e => setConfirm(e.target.value)} placeholder="Repeat password" onKeyDown={e => e.key === "Enter" && handleSet()} />
        {error   && <div style={errStyle}>{error}</div>}
        {success && <div style={okStyle}>Password set! Redirecting…</div>}
        <button style={{ ...btnPrimaryStyle, opacity: loading ? 0.7 : 1 }} onClick={handleSet} disabled={loading || success}>
          {loading ? "Saving…" : "Set Password & Sign In"}
        </button>
      </div>
    </div>
  );
}

// ── Root App — handles routing between auth states ────────────────────────────
export default function App() {
  const [session, setSession]           = useState(undefined); // undefined = loading
  const [isSettingPassword, setIsSettingPassword] = useState(false);

  useEffect(() => {
    // Check for invite token in URL hash (#access_token + type=invite)
    const hash = window.location.hash;
    if (hash.includes("type=invite") || hash.includes("type=recovery")) {
      setIsSettingPassword(true);
    }

    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (_event === "PASSWORD_RECOVERY") setIsSettingPassword(true);
    });
    return () => subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return <div style={{ minHeight: "100vh", background: "#f1f5f9", display: "flex", alignItems: "center", justifyContent: "center", color: "#64748b", fontFamily: "sans-serif" }}>Loading…</div>;
  }

  if (isSettingPassword) {
    return <SetPasswordPage onDone={() => { setIsSettingPassword(false); window.location.hash = ""; }} />;
  }

  if (!session) return <LoginPage />;

  return <Dashboard session={session} />;
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function parseDate(dateStr) {
  if (!dateStr || !String(dateStr).trim()) return null;
  const s = String(dateStr).trim();
  // ISO format from Supabase date column: YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const [y, m, d] = s.substring(0, 10).split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  // Legacy DD/MM/YYYY or DD/MM/YY formats
  let norm = s.replace(/-/g, '/');
  const parts = norm.split('/');
  if (parts.length !== 3) return null;
  let [d, m, y] = parts.map(Number);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  if (y < 100) y += 2000;
  return new Date(y, m - 1, d);
}
function getQuarter(date) {
  const mo = date.getMonth() + 1;
  if (mo <= 3) return 1;
  if (mo <= 6) return 2;
  if (mo <= 9) return 3;
  return 4;
}
function getYear(date) { return date.getFullYear(); }

const GOAL = 30000000;
const STAGE_COLORS = { Order: "#10b981", "On track": "#f59e0b", Fail: "#ef4444" };
const AGENT_COLORS = { PATHINYA: "#6366f1", LALITA: "#ec4899", SARUN: "#14b8a6" };

// ── Dashboard (renamed from default export) ──────────────────────────────────
function Dashboard({ session }) {

  // ── All hooks first (Rules of Hooks) ─────────────────────────────────────
  const [RAW_DATA, setRawData] = useState(null);
  const [dataError, setDataError] = useState(null);
  const [selectedYear, setSelectedYear] = useState(2026);
  const [selectedQ, setSelectedQ] = useState(1);
  const [activeTab, setActiveTab] = useState("overview");

  useEffect(() => {
    async function fetchData() {
      try {
        const { data, error } = await supabase
          .from('sales_pipeline')
          .select('"QO. Number","Company Name","Contact Person","Project Name","Total Price","Sales Agent","Stage","Create Date","Reason"')
          .order('id', { ascending: true });
        if (error) throw error;
        const mapped = data.map(r => ({
          qo:      r['QO. Number']      ?? '',
          company: r['Company Name']    ?? '',
          contact: r['Contact Person']  ?? '',
          project: r['Project Name']    ?? '',
          price:   parseFloat(r['Total Price']) || 0,
          agent:   r['Sales Agent']     ?? '',
          stage:   r['Stage']           ?? '',
          date:    r['Create Date']     ?? '',
          reason:  r['Reason']          ?? '',
        }));
        setRawData(mapped);
      } catch (err) {
        console.error('Supabase fetch error:', err);
        setDataError(err.message || 'Failed to load data');
      }
    }
    fetchData();
  }, []);

  const enriched = useMemo(() => (RAW_DATA ?? []).map(r => {
    const d = parseDate(r.date);
    return { ...r, parsedDate: d, quarter: d ? getQuarter(d) : null, year: d ? getYear(d) : null };
  }), [RAW_DATA]);

  const years = useMemo(() => {
    const ys = new Set(enriched.map(r => r.year).filter(Boolean));
    return [...ys].sort();
  }, [enriched]);

  const allForQ = useMemo(() =>
    enriched.filter(r => r.year === selectedYear && r.quarter === selectedQ),
  [enriched, selectedYear, selectedQ]);

  const stageCounts = useMemo(() => {
    const counts = { Order: 0, "On track": 0, Fail: 0 };
    allForQ.forEach(r => { if (counts[r.stage] !== undefined) counts[r.stage]++; });
    return counts;
  }, [allForQ]);

  const orderTotal = useMemo(() =>
    allForQ.filter(r => r.stage === "Order").reduce((s, r) => s + r.price, 0),
  [allForQ]);

  const onTrackTotal = useMemo(() =>
    allForQ.filter(r => r.stage === "On track").reduce((s, r) => s + r.price, 0),
  [allForQ]);

  const combinedTotal = orderTotal + onTrackTotal;
  const progress = Math.min((combinedTotal / GOAL) * 100, 100);

  const stagePieData = [
    { name: "Order", value: stageCounts.Order },
    { name: "On track", value: stageCounts["On track"] },
    { name: "Fail", value: stageCounts.Fail },
  ];

  // Agent performance
  const agentData = useMemo(() => {
    const agents = ["PATHINYA", "LALITA", "SARUN"];
    return agents.map(agent => {
      const rows = allForQ.filter(r => r.agent === agent);
      return {
        agent,
        Order: rows.filter(r => r.stage === "Order").length,
        "On track": rows.filter(r => r.stage === "On track").length,
        Fail: rows.filter(r => r.stage === "Fail").length,
        orderRevenue: rows.filter(r => r.stage === "Order").reduce((s, r) => s + r.price, 0),
      };
    });
  }, [allForQ]);

  // Top 5 on-track by price
  const top5OnTrack = useMemo(() => {
    return [...allForQ.filter(r => r.stage === "On track")]
      .sort((a, b) => b.price - a.price)
      .slice(0, 5);
  }, [allForQ]);

  // Monthly revenue trend (for current quarter)
  const monthlyData = useMemo(() => {
    const months = { 1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec" };
    const qMonths = { 1: [1,2,3], 2: [4,5,6], 3: [7,8,9], 4: [10,11,12] };
    const relevantMonths = qMonths[selectedQ];
    return relevantMonths.map(m => {
      const rows = enriched.filter(r => r.year === selectedYear && r.parsedDate && (r.parsedDate.getMonth()+1) === m);
      return {
        month: months[m],
        Order: rows.filter(r => r.stage === "Order").reduce((s, r) => s + r.price, 0) / 1000000,
        "On track": rows.filter(r => r.stage === "On track").reduce((s, r) => s + r.price, 0) / 1000000,
        Fail: rows.filter(r => r.stage === "Fail").reduce((s, r) => s + r.price, 0) / 1000000,
      };
    });
  }, [enriched, selectedYear, selectedQ]);

  const fmt = (n) => n >= 1000000 ? `฿${(n/1000000).toFixed(2)}M` : n >= 1000 ? `฿${(n/1000).toFixed(0)}K` : `฿${n.toFixed(0)}`;

  const tabs = [
    { id: "overview", label: "Overview" },
    { id: "agents", label: "Sales Agents" },
    { id: "pipeline", label: "Top Pipeline" },
    { id: "trend", label: "Monthly Trend" },
  ];

  const quarterLabel = `Q${selectedQ} ${selectedYear}`;

  // ── Early returns after all hooks ────────────────────────────────────────
  if (dataError) return (
    <div style={{ display:'flex', alignItems:'center', justifyContent:'center', height:'100vh', background:'#f1f5f9' }}>
      <div style={{ background:'#fff', border:'1px solid #e2e8f0', borderRadius:12, padding:32, maxWidth:400, textAlign:'center' }}>
        <div style={{ fontSize:32, marginBottom:12 }}>⚠️</div>
        <div style={{ fontWeight:700, color:'#0f172a', marginBottom:8 }}>Failed to load data</div>
        <div style={{ color:'#64748b', fontSize:14 }}>{dataError}</div>
      </div>
    </div>
  );

  if (!RAW_DATA) return (
    <div style={{ display:'flex', alignItems:'center', justifyContent:'center', height:'100vh', background:'#f1f5f9' }}>
      <div style={{ textAlign:'center' }}>
        <div style={{ width:40, height:40, border:'3px solid #e2e8f0', borderTop:'3px solid #6366f1', borderRadius:'50%', animation:'spin 0.8s linear infinite', margin:'0 auto 16px' }} />
        <div style={{ color:'#64748b', fontSize:14 }}>Loading pipeline data…</div>
      </div>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );

  return (
    <div style={{ fontFamily: "'DM Sans', 'Segoe UI', sans-serif", background: "#f1f5f9", minHeight: "100vh", color: "#0f172a", padding: "0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=Space+Grotesk:wght@600;700&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #e2e8f0; }
        ::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 3px; }

        .tab-btn { background: none; border: none; cursor: pointer; padding: 8px 14px; border-radius: 8px; font-family: inherit; font-size: 13px; font-weight: 500; transition: all 0.2s; white-space: nowrap; }
        .tab-btn.active { background: #2563eb; color: white; }
        .tab-btn:not(.active) { color: #64748b; }
        .tab-btn:not(.active):hover { background: #e2e8f0; color: #0f172a; }

        .tab-bar { display: flex; gap: 4px; overflow-x: auto; -webkit-overflow-scrolling: touch; scrollbar-width: none; }
        .tab-bar::-webkit-scrollbar { display: none; }

        .filter-select { background: #ffffff; border: 1px solid #cbd5e1; color: #0f172a; padding: 6px 12px; border-radius: 8px; font-family: inherit; font-size: 13px; cursor: pointer; }
        .filter-select:focus { outline: none; border-color: #3b82f6; }

        .card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; padding: 20px; }
        .kpi-card { background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%); border: 1px solid #e2e8f0; border-radius: 16px; padding: 16px; }
        .progress-bar-bg { background: #e2e8f0; border-radius: 99px; height: 12px; overflow: hidden; }
        .progress-bar-fill { height: 100%; border-radius: 99px; transition: width 1s ease; background: linear-gradient(90deg, #1d4ed8, #3b82f6, #60a5fa); }
        .badge { display: inline-block; padding: 2px 10px; border-radius: 99px; font-size: 11px; font-weight: 600; }
        .badge-order { background: rgba(16,185,129,0.15); color: #10b981; }
        .badge-ontrack { background: rgba(245,158,11,0.15); color: #f59e0b; }
        .badge-fail { background: rgba(239,68,68,0.15); color: #ef4444; }
        .table-row:hover { background: #f8fafc; }
        .glow-green { box-shadow: 0 2px 12px rgba(16,185,129,0.15); }
        .glow-amber { box-shadow: 0 0 20px rgba(245,158,11,0.1); }

        .quarter-btn { background: none; border: 1px solid #cbd5e1; color: #64748b; padding: 5px 12px; border-radius: 8px; font-family: inherit; font-size: 13px; cursor: pointer; transition: all 0.2s; }
        .quarter-btn.active { background: #2563eb; border-color: #2563eb; color: white; }
        .quarter-btn:hover:not(.active) { border-color: #94a3b8; color: #0f172a; }

        .header-inner { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 12px; padding: 20px 32px; }
        .header-controls { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
        .page-pad { padding: 24px 32px; }
        .chart-grid-2 { display: grid; grid-template-columns: 1fr 2fr; gap: 16px; }
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 20px; }

        @media (max-width: 640px) {
          .header-inner { padding: 16px; flex-direction: column; align-items: flex-start; }
          .header-controls { width: 100%; justify-content: flex-start; }
          .page-pad { padding: 16px; }
          .chart-grid-2 { grid-template-columns: 1fr; }
          .kpi-grid { grid-template-columns: repeat(2, 1fr); }
          .card { padding: 16px; border-radius: 12px; }
          .tab-bar { padding: 12px 16px 0; }
          .goal-amount { font-size: 22px !important; }
          .goal-pct { font-size: 24px !important; }
          .top5-price { font-size: 13px !important; }
          .agent-revenue { font-size: 13px !important; }
        }
      `}</style>

      {/* Header */}
      <div style={{ background: "linear-gradient(180deg, #ffffff 0%, #f8fafc 100%)", borderBottom: "1px solid #e2e8f0" }}>
        <div className="header-inner">
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 2 }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: "linear-gradient(135deg, #1d4ed8, #7c3aed)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>⚡</div>
              <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 20, fontWeight: 700, color: "#0f172a", letterSpacing: "-0.3px" }}>PC Sales Pipeline</h1>
            </div>
            <p style={{ fontSize: 12, color: "#64748b" }}>Revenue tracking dashboard · Quarterly target: ฿30,000,000</p>
          </div>
          <div className="header-controls">
            <select className="filter-select" value={selectedYear} onChange={e => setSelectedYear(Number(e.target.value))}>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            {[1, 2, 3, 4].map(q => (
              <button key={q} className={`quarter-btn ${selectedQ === q ? "active" : ""}`} onClick={() => setSelectedQ(q)}>Q{q}</button>
            ))}
            <button onClick={() => supabase.auth.signOut()} style={{ background: "none", border: "1px solid #e2e8f0", borderRadius: 8, padding: "5px 12px", fontSize: 12, color: "#94a3b8", cursor: "pointer", fontFamily: "inherit" }}>Sign out</button>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div style={{ borderBottom: "1px solid #e2e8f0", background: "#ffffff" }}>
        <div className="tab-bar" style={{ padding: "12px 32px 0" }}>
          {tabs.map(t => (
            <button key={t.id} className={`tab-btn ${activeTab === t.id ? "active" : ""}`} onClick={() => setActiveTab(t.id)}>{t.label}</button>
          ))}
        </div>
      </div>

      <div className="page-pad">

        {/* OVERVIEW TAB */}
        {activeTab === "overview" && (
          <div>
            {/* Goal Progress */}
            <div className="card glow-green" style={{ marginBottom: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
                <div style={{marginBottom: "auto"}}>
                  <div style={{ fontSize: 12, color: "#475569", fontWeight: 500, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.5px"}}>Quarterly Revenue Goal · {quarterLabel}</div>
                  <div style={{ fontSize: 14, color: "#64748b" }}>Target: {fmt(GOAL)}</div>
                </div>
                <div style={{ display: "flex", gap: 24, flexWrap: "wrap", justifyContent: "flex-end" }}>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 11, color: "#475569", fontWeight: 500, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 3 }}>
                      <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: 2, background: "#3b82f6", marginRight: 5, verticalAlign: "middle" }} />
                      Confirmed Orders
                    </div>
                    <div style={{ fontSize: 22, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", color: "#1d4ed8" }}>{((orderTotal / GOAL) * 100).toFixed(1)}%</div>
                    <div style={{ fontSize: 13, color: "#475569", marginTop: 1 }}>{fmt(orderTotal)} <span style={{ color: "#94a3b8" }}>of {fmt(GOAL)}</span></div>
                  </div>
                  <div style={{ borderLeft: "1px solid #e2e8f0", paddingLeft: 24, textAlign: "right" }}>
                    <div style={{ fontSize: 11, color: "#475569", fontWeight: 500, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 3 }}>
                      <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: 2, background: "#93c5fd", marginRight: 5, verticalAlign: "middle" }} />
                      Pending Orders
                    </div>
                    <div style={{ fontSize: 22, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", color: "#3b82f6" }}>{Math.min((combinedTotal / GOAL) * 100, 100).toFixed(1)}%</div>
                    <div style={{ fontSize: 13, color: "#475569", marginTop: 1 }}>{fmt(combinedTotal)} <span style={{ color: "#94a3b8" }}>of {fmt(GOAL)}</span></div>
                  </div>
                </div>
              </div>

              {/* Dual-layer progress bar */}
              <div style={{ background: "#e2e8f0", borderRadius: 99, height: 14, overflow: "hidden", position: "relative" }}>
                <div style={{ position: "absolute", left: 0, top: 0, height: "100%", width: `${Math.min((combinedTotal / GOAL) * 100, 100)}%`, background: "#93c5fd", borderRadius: 99, transition: "width 1s ease" }} />
                <div style={{ position: "absolute", left: 0, top: 0, height: "100%", width: `${Math.min((orderTotal / GOAL) * 100, 100)}%`, background: "linear-gradient(90deg, #1d4ed8, #3b82f6)", borderRadius: 99, transition: "width 1s ease" }} />
              </div>
              <div style={{ fontSize: 12, color: "#64748b", marginTop: 8, textAlign: "right" }}>
                {combinedTotal >= GOAL ? "🎯 Goal achieved!" : `${fmt(GOAL - combinedTotal)} remaining`}
              </div>
            </div>

            {/* KPI Cards */}
            <div className="kpi-grid">
              {[
                { label: "Total Quotations", value: allForQ.length, sub: "All statuses", icon: "📋", color: "#3b82f6" },
                { label: "Orders Confirmed", value: stageCounts.Order, sub: `${fmt(orderTotal)}`, icon: "✅", color: "#10b981" },
                { label: "In Progress", value: stageCounts["On track"], sub: `${fmt(onTrackTotal)} potential`, icon: "🔄", color: "#f59e0b" },
                { label: "Failed", value: stageCounts.Fail, sub: "Excluded from total", icon: "❌", color: "#ef4444" },
                { label: "Win Rate", value: `${stageCounts.Order + stageCounts["On track"] + stageCounts.Fail > 0 ? ((stageCounts.Order / (stageCounts.Order + stageCounts.Fail)) * 100).toFixed(0) : 0}%`, sub: "Orders vs closed deals", icon: "🎯", color: "#8b5cf6" },
              ].map((kpi, i) => (
                <div key={i} className="kpi-card" style={{ borderLeft: `3px solid ${kpi.color}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontSize: 11, color: "#475569", fontWeight: 500, textTransform: "uppercase", letterSpacing: "0.5px", marginBottom: 6 }}>{kpi.label}</div>
                      <div style={{ fontSize: 26, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", color: kpi.color }}>{kpi.value}</div>
                      <div style={{ fontSize: 11, color: "#64748b", marginTop: 2 }}>{kpi.sub}</div>
                    </div>
                    <div style={{ fontSize: 22 }}>{kpi.icon}</div>
                  </div>
                </div>
              ))}
            </div>

            {/* Stage Pie + Agent Overview */}
            <div className="chart-grid-2">
              <div className="card">
                <div style={{ fontSize: 13, fontWeight: 600, color: "#64748b", marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>Status Distribution</div>
                <ResponsiveContainer width="100%" height={180}>
                  <PieChart>
                    <Pie data={stagePieData} cx="50%" cy="50%" innerRadius={45} outerRadius={70} dataKey="value" paddingAngle={3}>
                      {stagePieData.map((entry, index) => (
                        <Cell key={index} fill={STAGE_COLORS[entry.name]} />
                      ))}
                    </Pie>
                    <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12 }} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
              <div className="card">
                <div style={{ fontSize: 13, fontWeight: 600, color: "#64748b", marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>Agent Revenue Breakdown</div>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={agentData} barCategoryGap="30%">
                    <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                    <XAxis dataKey="agent" tick={{ fill: "#64748b", fontSize: 12 }} axisLine={false} tickLine={false} />
                    <YAxis tick={{ fill: "#64748b", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={v => `${(v/1000000).toFixed(1)}M`} width={40} />
                    <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12 }} formatter={v => fmt(v)} />
                    <Bar dataKey="orderRevenue" name="Confirmed Revenue" radius={[4,4,0,0]}>
                      {agentData.map((entry, i) => <Cell key={i} fill={AGENT_COLORS[entry.agent]} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        )}

        {/* AGENTS TAB */}
        {activeTab === "agents" && (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 16, marginBottom: 24 }}>
              {agentData.map(agent => (
                <div key={agent.agent} className="card" style={{ borderTop: `3px solid ${AGENT_COLORS[agent.agent]}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                    <div>
                      <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", color: AGENT_COLORS[agent.agent] }}>{agent.agent}</div>
                      <div style={{ fontSize: 12, color: "#475569" }}>{agent.Order + agent["On track"] + agent.Fail} total quotations</div>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <div style={{ fontSize: 14, fontWeight: 600, color: "#0f172a" }}>{fmt(agent.orderRevenue)}</div>
                      <div style={{ fontSize: 11, color: "#64748b" }}>confirmed revenue</div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
                    <div style={{ flex: 1, background: "rgba(16,185,129,0.1)", borderRadius: 10, padding: "10px", textAlign: "center" }}>
                      <div style={{ fontSize: 22, fontWeight: 700, color: "#10b981" }}>{agent.Order}</div>
                      <div style={{ fontSize: 11, color: "#475569" }}>Orders</div>
                    </div>
                    <div style={{ flex: 1, background: "rgba(245,158,11,0.1)", borderRadius: 10, padding: "10px", textAlign: "center" }}>
                      <div style={{ fontSize: 22, fontWeight: 700, color: "#f59e0b" }}>{agent["On track"]}</div>
                      <div style={{ fontSize: 11, color: "#475569" }}>On Track</div>
                    </div>
                    <div style={{ flex: 1, background: "rgba(239,68,68,0.1)", borderRadius: 10, padding: "10px", textAlign: "center" }}>
                      <div style={{ fontSize: 22, fontWeight: 700, color: "#ef4444" }}>{agent.Fail}</div>
                      <div style={{ fontSize: 11, color: "#475569" }}>Failed</div>
                    </div>
                  </div>
                  <div style={{ height: 8, borderRadius: 99, overflow: "hidden", display: "flex", gap: 2 }}>
                    {agent.Order > 0 && <div style={{ flex: agent.Order, background: "#10b981", borderRadius: "99px 0 0 99px" }}></div>}
                    {agent["On track"] > 0 && <div style={{ flex: agent["On track"], background: "#f59e0b" }}></div>}
                    {agent.Fail > 0 && <div style={{ flex: agent.Fail, background: "#ef4444", borderRadius: "0 99px 99px 0" }}></div>}
                  </div>
                </div>
              ))}
            </div>

            <div className="card">
              <div style={{ fontSize: 13, fontWeight: 600, color: "#64748b", marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>Orders by Agent — Count Comparison</div>
              <ResponsiveContainer width="100%" height={220}>
                <BarChart data={agentData} barCategoryGap="35%">
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="agent" tick={{ fill: "#64748b", fontSize: 12 }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fill: "#64748b", fontSize: 11 }} axisLine={false} tickLine={false} width={30} />
                  <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12 }} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="Order" fill="#10b981" radius={[3,3,0,0]} />
                  <Bar dataKey="On track" fill="#f59e0b" radius={[3,3,0,0]} />
                  <Bar dataKey="Fail" fill="#ef4444" radius={[3,3,0,0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* PIPELINE TAB */}
        {activeTab === "pipeline" && (
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 20 }}>
              <div style={{ background: "rgba(245,158,11,0.1)", border: "1px solid rgba(245,158,11,0.2)", borderRadius: 10, padding: "10px 16px" }}>
                <div style={{ fontSize: 11, color: "#f59e0b", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.5px" }}>Top 5 On-Track · {quarterLabel}</div>
                <div style={{ fontSize: 20, fontWeight: 700, color: "#0f172a", marginTop: 2 }}>
                  {fmt(top5OnTrack.reduce((s, r) => s + r.price, 0))} <span style={{ fontSize: 13, color: "#475569" }}>combined potential</span>
                </div>
              </div>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 24 }}>
              {top5OnTrack.map((row, i) => (
                <div key={row.qo} className="card" style={{ display: "flex", alignItems: "center", gap: 16, padding: "16px 20px" }}>
                  <div style={{ width: 36, height: 36, borderRadius: 10, background: `rgba(245,158,11,${0.3 - i*0.04})`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", color: "#f59e0b", fontSize: 16, flexShrink: 0 }}>#{i+1}</div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, color: "#0f172a", fontSize: 14, marginBottom: 3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.company}</div>
                    <div style={{ fontSize: 12, color: "#64748b", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.project || "—"}</div>
                  </div>
                  <div style={{ textAlign: "right", flexShrink: 0 }}>
                    <div style={{ fontSize: 16, fontWeight: 700, color: "#f59e0b", fontFamily: "'Space Grotesk', sans-serif" }}>{fmt(row.price)}</div>
                    <div style={{ fontSize: 11, color: "#475569", marginTop: 2 }}>Agent: {row.agent}</div>
                  </div>
                </div>
              ))}
            </div>

            <div className="card">
              <div style={{ fontSize: 13, fontWeight: 600, color: "#64748b", marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>All On-Track Quotations · {quarterLabel}</div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: "1px solid #e2e8f0" }}>
                      {["QO Number", "Company", "Project", "Agent", "Price", "Date"].map(h => (
                        <th key={h} style={{ textAlign: "left", padding: "8px 12px", color: "#475569", fontWeight: 500, fontSize: 12, textTransform: "uppercase", letterSpacing: "0.5px" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {allForQ.filter(r => r.stage === "On track").sort((a,b) => b.price - a.price).map((row, i) => (
                      <tr key={i} className="table-row" style={{ borderBottom: "1px solid #e2e8f0" }}>
                        <td style={{ padding: "10px 12px", color: "#64748b", fontFamily: "monospace", fontSize: 12 }}>{row.qo}</td>
                        <td style={{ padding: "10px 12px", color: "#0f172a", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.company}</td>
                        <td style={{ padding: "10px 12px", color: "#64748b", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.project || "—"}</td>
                        <td style={{ padding: "10px 12px" }}><span style={{ color: AGENT_COLORS[row.agent], fontWeight: 500 }}>{row.agent}</span></td>
                        <td style={{ padding: "10px 12px", fontWeight: 600, color: "#f59e0b" }}>{fmt(row.price)}</td>
                        <td style={{ padding: "10px 12px", color: "#94a3b8", fontSize: 12 }}>{row.date || "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TREND TAB */}
        {activeTab === "trend" && (
          <div>
            <div className="card" style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 13, fontWeight: 600, color: "#64748b", marginBottom: 16, textTransform: "uppercase", letterSpacing: "0.5px" }}>Monthly Revenue Breakdown · {quarterLabel} (฿M)</div>
              <ResponsiveContainer width="100%" height={250}>
                <BarChart data={monthlyData} barCategoryGap="35%">
                  <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                  <XAxis dataKey="month" tick={{ fill: "#64748b", fontSize: 13 }} axisLine={false} tickLine={false} />
                  <YAxis tick={{ fill: "#64748b", fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={v => `฿${v.toFixed(1)}M`} width={50} />
                  <Tooltip contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12 }} formatter={(v) => [`฿${(v).toFixed(2)}M`]} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Bar dataKey="Order" fill="#10b981" radius={[4,4,0,0]} name="Orders" />
                  <Bar dataKey="On track" fill="#f59e0b" radius={[4,4,0,0]} name="On Track" />
                  <Bar dataKey="Fail" fill="#ef4444" radius={[4,4,0,0]} name="Failed" />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 16 }}>
              {monthlyData.map(m => (
                <div key={m.month} className="card">
                  <div style={{ fontSize: 16, fontWeight: 700, color: "#0f172a", fontFamily: "'Space Grotesk', sans-serif", marginBottom: 12 }}>{m.month}</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 12, color: "#475569" }}>Orders</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: "#10b981" }}>฿{m.Order.toFixed(2)}M</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 12, color: "#475569" }}>On Track</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: "#f59e0b" }}>฿{m["On track"].toFixed(2)}M</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 12, color: "#475569" }}>Failed</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: "#ef4444" }}>฿{m.Fail.toFixed(2)}M</span>
                    </div>
                    <div style={{ borderTop: "1px solid #1e293b", paddingTop: 8, display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 12, color: "#475569", fontWeight: 600 }}>Total Active</span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: "#3b82f6" }}>฿{(m.Order + m["On track"]).toFixed(2)}M</span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      <div style={{ textAlign: "center", padding: "16px", color: "#94a3b8", fontSize: 11 }}>
        PC Sales Pipeline Dashboard · Data from Q1–Q3 2026 · {RAW_DATA.length} records (live from Supabase)
      </div>
    </div>
  );
}