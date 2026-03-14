import { useState, useMemo, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend, LineChart, Line, CartesianGrid } from "recharts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";
import ExcelJS from "https://esm.sh/exceljs@4.4.0";

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

// Format ISO date string (YYYY-MM-DD) to display format (DD/MM/YYYY)
function isoToDisplay(iso) {
  if (!iso) return "";
  const s = String(iso).trim().substring(0, 10);
  const [y, m, d] = s.split("-");
  if (!y || !m || !d) return iso;
  return `${d}/${m}/${y}`;
}

// Parse DD/MM/YYYY or YYYY-MM-DD input back to ISO YYYY-MM-DD for Supabase
function displayToISO(val) {
  if (!val) return null;
  const s = String(val).trim();
  // Already ISO
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // DD/MM/YYYY
  const parts = s.split("/");
  if (parts.length === 3) {
    let [d, m, y] = parts.map(Number);
    if (y < 100) y += 2000;
    if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
    return `${String(y).padStart(4,"0")}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
  }
  return null;
}

const GOAL = 30000000;
const STAGE_COLORS = { Order: "#10b981", "On track": "#f59e0b", Fail: "#ef4444" };
const AGENT_PALETTE = ["#6366f1","#ec4899","#14b8a6","#f59e0b","#10b981","#3b82f6","#8b5cf6","#ef4444","#06b6d4","#84cc16"];
function getAgentColor(agent, agentList) {
  const i = agentList.indexOf(agent);
  return AGENT_PALETTE[i % AGENT_PALETTE.length];
}

// ── Review Modal ──────────────────────────────────────────────────────────────
const FIELDS = [
  { key: "qo_number",      label: "QO Number",      type: "text" },
  { key: "company_name",   label: "Company Name",   type: "text" },
  { key: "contact_person", label: "Contact Person", type: "text" },
  { key: "project_name",   label: "Project Name",   type: "text" },
  { key: "total_price",    label: "Total Price",    type: "number" },
  { key: "revision",       label: "Revision",       type: "number" },
  { key: "validity",       label: "Validity (days)",type: "number" },
  { key: "sales_agent",    label: "Sales Agent",    type: "select", options: [] }, // options populated from agents prop
  { key: "stage",          label: "Stage",          type: "select", options: ["On track", "Order", "Fail"] },
  { key: "create_date",    label: "Create Date",    type: "date" },
  { key: "reason",         label: "Reason",         type: "text" },
  { key: "po_qt",          label: "PO / QT",        type: "text" },
  { key: "follow_up_1",    label: "Follow Up 1",    type: "text" },
  { key: "follow_up_2",    label: "Follow Up 2",    type: "text" },
  { key: "follow_up_3",    label: "Follow Up 3",    type: "text" }
];

function ReviewModal({ row, agents, onSave, onCancel }) {
  const [form, setForm] = useState({ ...row });
  const set = (key, val) => setForm(f => ({ ...f, [key]: val }));

  const overlayStyle = {
    position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)",
    display: "flex", alignItems: "center", justifyContent: "center",
    zIndex: 1000, padding: 16,
  };
  const modalStyle = {
    background: "#fff", borderRadius: 16, width: "100%", maxWidth: 560,
    maxHeight: "90vh", display: "flex", flexDirection: "column",
    boxShadow: "0 24px 64px rgba(0,0,0,0.18)",
  };
  const headerStyle = {
    padding: "20px 24px 16px", borderBottom: "1px solid #e2e8f0",
    display: "flex", alignItems: "center", justifyContent: "space-between",
  };
  const bodyStyle = { padding: "20px 24px", overflowY: "auto", flex: 1 };
  const footerStyle = {
    padding: "16px 24px", borderTop: "1px solid #e2e8f0",
    display: "flex", gap: 10, justifyContent: "flex-end",
  };
  const fieldStyle = { marginBottom: 14 };
  const labelStyle2 = { fontSize: 11, fontWeight: 600, color: "#475569", textTransform: "uppercase", letterSpacing: "0.4px", display: "block", marginBottom: 4 };
  const inputStyle2 = { width: "100%", padding: "8px 10px", borderRadius: 8, border: "1px solid #cbd5e1", fontSize: 13, fontFamily: "inherit", boxSizing: "border-box", color: "#0f172a" };

  return (
    <div style={overlayStyle} onClick={e => e.target === e.currentTarget && onCancel()}>
      <div style={modalStyle}>
        <div style={headerStyle}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 16, color: "#0f172a" }}>Review Quotation</div>
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>Check and edit before saving to Supabase</div>
          </div>
          <button onClick={onCancel} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: "#94a3b8", lineHeight: 1 }}>✕</button>
        </div>

        <div style={bodyStyle}>
          {FIELDS.map(({ key, label, type, options }) => (
            <div key={key} style={fieldStyle}>
              <label style={labelStyle2}>{label}</label>
              {key === "sales_agent" ? (
                <>
                  <input
                    list="agent-options"
                    value={form[key] ?? ""}
                    onChange={e => set(key, e.target.value)}
                    placeholder="Select or type new agent…"
                    style={inputStyle2}
                  />
                  <datalist id="agent-options">
                    {agents.map(a => <option key={a} value={a} />)}
                  </datalist>
                </>
              ) : type === "select" ? (
                <select value={form[key] ?? ""} onChange={e => set(key, e.target.value)} style={inputStyle2}>
                  {options.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              ) : (
                <input
                  type={type}
                  value={form[key] ?? ""}
                  onChange={e => set(key, type === "number" ? parseFloat(e.target.value) || 0 : e.target.value)}
                  style={inputStyle2}
                />
              )}
            </div>
          ))}
        </div>

        <div style={footerStyle}>
          <button onClick={onCancel} style={{ padding: "9px 18px", borderRadius: 8, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 500, cursor: "pointer", color: "#475569" }}>
            Cancel
          </button>
          <button onClick={() => onSave(form)} style={{ padding: "9px 18px", borderRadius: 8, border: "none", background: "#2563eb", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
            Save to Supabase
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Dashboard (renamed from default export) ──────────────────────────────────
function Dashboard({ session }) {

  // ── All hooks first (Rules of Hooks) ─────────────────────────────────────
  const [RAW_DATA, setRawData] = useState(null);
  const [dataError, setDataError] = useState(null);

  // ── Set default quarter/year to current date on first load ─────────────────────
  const today = new Date();
  const defaultYear = today.getFullYear();
  const defaultMonth = today.getMonth();
  const defaultQ = Math.floor((defaultMonth) / 3) + 1;

  // ── Table-tab server-side state ──────────────────────────────────────────
  const [tableRows, setTableRows]     = useState([]);
  const [tableTotal, setTableTotal]   = useState(0);
  const [tableLoading, setTableLoading] = useState(false);
  const [tableRefreshKey, setTableRefreshKey] = useState(0);
  const [selectedYear, setSelectedYear] = useState(defaultYear);
  const [selectedQ, setSelectedQ] = useState(defaultQ);
  const [activeTab, setActiveTab] = useState("overview");
  const [uploadState, setUploadState] = useState("idle"); // idle | loading | success | error
  const [uploadMessage, setUploadMessage] = useState("");
  const [reviewRow, setReviewRow] = useState(null); // null = modal closed
  const [tablePage, setTablePage] = useState(1);
  const [tableSearch, setTableSearch] = useState("");
  const [tableSort, setTableSort] = useState([]); // array of { key, dir: 'asc'|'desc' }
  const [editingCell, setEditingCell] = useState(null); // { id, field }
  const [editingValue, setEditingValue] = useState("");
  const [savingCell, setSavingCell] = useState(null);
  const [deleteConfirm, setDeleteConfirm] = useState(null); // { qo_number, company_name } | null
  const [deletingRow, setDeletingRow] = useState(null); // qo_number string | null

  // ── Helper: quarter → ISO date range ────────────────────────────────────
  function quarterDateRange(year, q) {
    const startMonth = (q - 1) * 3 + 1;
    const endMonth   = q * 3;
    const pad = n => String(n).padStart(2, "0");
    const lastDay = new Date(year, endMonth, 0).getDate();
    return {
      from: `${year}-${pad(startMonth)}-01`,
      to:   `${year}-${pad(endMonth)}-${lastDay}`,
    };
  }

  // ── Fetch stats data filtered by selected quarter/year ───────────────────
  useEffect(() => {
    async function fetchStats() {
      setRawData(null);
      setDataError(null);
      try {
        const { from, to } = quarterDateRange(selectedYear, selectedQ);
        const { data, error } = await supabase
          .from('sales_pipeline')
          .select('qo_number, company_name, contact_person, project_name, total_price, validity, sales_agent, stage, create_date, reason, po_qt, follow_up_1, follow_up_2, follow_up_3, revision')
          .gte('create_date', from)
          .lte('create_date', to)
          .order('id', { ascending: true });
        if (error) throw error;
        const mapped = data.map(r => ({
          qo:      r.qo_number      ?? '',
          company: r.company_name   ?? '',
          contact: r.contact_person ?? '',
          project: r.project_name   ?? '',
          price:   parseFloat(r.total_price) || 0,
          agent:   r.sales_agent    ?? '',
          stage:   r.stage          ?? '',
          date:    r.create_date    ?? '',
          reason:     r.reason      ?? '',
          validity:   r.validity    ?? null,
          po_qt:      r.po_qt       ?? '',
          follow_up_1: r.follow_up_1 ?? '',
          follow_up_2: r.follow_up_2 ?? '',
          follow_up_3: r.follow_up_3 ?? '',
          revision:    r.revision    ?? 0,
        }));
        setRawData(mapped);
      } catch (err) {
        console.error('Supabase fetch error:', err);
        setDataError(err.message || 'Failed to load data');
      }
    }
    fetchStats();
  }, [selectedYear, selectedQ]);

  // ── Fetch table-tab data server-side (search / sort / paginate) ─────────
  const TABLE_PAGE_SIZE = 20;
  // Supabase column name map (display col.key → db column)
  const TABLE_COL_TO_DB = {
    qo_number:      "qo_number",
    create_date:    "create_date",
    company_name:   "company_name",
    contact_person: "contact_person",
    project_name:   "project_name",
    total_price:    "total_price",
    revision:       "revision",
    sales_agent:    "sales_agent",
    stage:          "stage",
    po_qt:          "po_qt",
    validity:       "validity",
    reason:         "reason",
    follow_up_1:    "follow_up_1",
    follow_up_2:    "follow_up_2",
    follow_up_3:    "follow_up_3",
  };

  useEffect(() => {
    if (activeTab !== "table") return;
    async function fetchTable() {
      setTableLoading(true);
      try {
        const from = (tablePage - 1) * TABLE_PAGE_SIZE;
        const to   = from + TABLE_PAGE_SIZE - 1;

        let query = supabase
          .from("sales_pipeline")
          .select(
            "qo_number, company_name, contact_person, project_name, total_price, validity, sales_agent, stage, create_date, reason, po_qt, follow_up_1, follow_up_2, follow_up_3, revision",
            { count: "exact" }
          );

        // Full-text search via ilike on multiple columns
        if (tableSearch.trim()) {
          const q = `%${tableSearch.trim()}%`;
          query = query.or(
            `qo_number.ilike.${q},company_name.ilike.${q},contact_person.ilike.${q},project_name.ilike.${q},sales_agent.ilike.${q},stage.ilike.${q}`
          );
        }

        // Multi-column sort
        if (tableSort.length > 0) {
          tableSort.forEach(({ key, dir }) => {
            const col = TABLE_COL_TO_DB[key] ?? key;
            query = query.order(col, { ascending: dir === "asc", nullsFirst: false });
          });
        } else {
          query = query.order("id", { ascending: true });
        }

        query = query.range(from, to);

        const { data, error, count } = await query;
        if (error) throw error;
        setTableRows(data ?? []);
        setTableTotal(count ?? 0);
      } catch (err) {
        console.error("Table fetch error:", err);
      } finally {
        setTableLoading(false);
      }
    }
    fetchTable();
  }, [activeTab, tablePage, tableSearch, tableSort, tableRefreshKey]);

  const enriched = useMemo(() => (RAW_DATA ?? []).map(r => {
    const d = parseDate(r.date);
    return { ...r, parsedDate: d, quarter: d ? getQuarter(d) : null, year: d ? getYear(d) : null };
  }), [RAW_DATA]);

  const [availableYears, setAvailableYears] = useState([2026]);

  useEffect(() => {
    async function fetchYears() {
      try {
        // Fetch only create_date to derive distinct years without the 1000-row limit issue
        const { data, error } = await supabase
          .from("sales_pipeline")
          .select("create_date")
          .not("create_date", "is", null);
        if (error) throw error;
        const ys = new Set((data ?? []).map(r => {
          const d = parseDate(r.create_date);
          return d ? getYear(d) : null;
        }).filter(Boolean));
        const sorted = [...ys].sort();
        if (sorted.length > 0) setAvailableYears(sorted);
      } catch (err) {
        console.error("fetchYears error:", err);
      }
    }
    fetchYears();
  }, []);

  const years = availableYears;

  // Data is already filtered by quarter/year from Supabase
  const allForQ = enriched;

  const stageCounts = useMemo(() => {
    const counts = { Order: 0, "On track": 0, Fail: 0 };
    allForQ.forEach(r => { if (counts[r.stage] !== undefined) counts[r.stage]++; });
    return counts;
  }, [allForQ]);

  const allStatusTotal = useMemo(() =>
    allForQ.reduce((s, r) => s + r.price, 0),
  [allForQ]);

  const orderTotal = useMemo(() =>
    allForQ.filter(r => r.stage === "Order").reduce((s, r) => s + r.price, 0),
  [allForQ]);

  const onTrackTotal = useMemo(() =>
    allForQ.filter(r => r.stage === "On track").reduce((s, r) => s + r.price, 0),
  [allForQ]);

  const failedTotal = useMemo(() =>
    allForQ.filter(r => r.stage === "Fail").reduce((s, r) => s + r.price, 0),
  [allForQ]);

  const combinedTotal = orderTotal + onTrackTotal;
  const progress = Math.min((combinedTotal / GOAL) * 100, 100);

  const stagePieData = [
    { name: "Order", value: stageCounts.Order },
    { name: "On track", value: stageCounts["On track"] },
    { name: "Fail", value: stageCounts.Fail },
  ];

  const [allAgents, setAllAgents] = useState([]);

  useEffect(() => {
    async function fetchAgents() {
      try {
        const { data, error } = await supabase
          .from("sales_pipeline")
          .select("sales_agent")
          .not("sales_agent", "is", null)
          .neq("sales_agent", "");
        if (error) throw error;
        const set = new Set((data ?? []).map(r => r.sales_agent).filter(Boolean));
        setAllAgents([...set].sort());
      } catch (err) {
        console.error("fetchAgents error:", err);
      }
    }
    fetchAgents();
  }, []);

  // Derive unique agents from full dataset, sorted alphabetically
  const agents = allAgents;

  // Agent performance — only include agents with activity this quarter
  const agentData = useMemo(() => {
    return agents.map(agent => {
      const rows = allForQ.filter(r => r.agent === agent);
      return {
        agent,
        Order: rows.filter(r => r.stage === "Order").length,
        "On track": rows.filter(r => r.stage === "On track").length,
        Fail: rows.filter(r => r.stage === "Fail").length,
        orderRevenue: rows.filter(r => r.stage === "Order").reduce((s, r) => s + r.price, 0),
      };
    }).filter(a => a.Order + a["On track"] + a.Fail > 0);
  }, [agents, allForQ]);

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
      const rows = enriched.filter(r => r.parsedDate && (r.parsedDate.getMonth()+1) === m);
      const orderRows    = rows.filter(r => r.stage === "Order");
      const onTrackRows  = rows.filter(r => r.stage === "On track");
      const failRows     = rows.filter(r => r.stage === "Fail");
      return {
        month: months[m],
        Order:          orderRows.reduce((s, r) => s + r.price, 0) / 1000000,
        "On track":     onTrackRows.reduce((s, r) => s + r.price, 0) / 1000000,
        Fail:           failRows.reduce((s, r) => s + r.price, 0) / 1000000,
        OrderCount:     orderRows.length,
        OnTrackCount:   onTrackRows.length,
        FailCount:      failRows.length,
      };
    });
  }, [enriched, selectedQ]);

  const fmt = (n) => n >= 1000000 ? `฿${(n/1000000).toFixed(2)}M` : n >= 1000 ? `฿${(n/1000).toFixed(0)}K` : `฿${n.toFixed(0)}`;

  const tabs = [
    { id: "overview", label: "Overview" },
    { id: "agents", label: "Sales Agents" },
    { id: "pipeline", label: "Top Pipeline" },
    { id: "trend", label: "Monthly Trend" },
    { id: "table", label: "Data Table" },
  ];

  // ── Excel upload ─────────────────────────────────────────────────────────
  async function extractQuotationData(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[0];

    function getCell(addr) {
      let value = sheet.getCell(addr).value;
      // ExcelJS returns formula cells as { formula, result } — use the result
      if (value !== null && typeof value === "object" && "result" in value) value = value.result;
      if (value === undefined || value === null || value === "")
        throw new Error(`Missing required cell: ${addr}`);
      return value;
    }
    function toISO(value) {
      if (value instanceof Date) {
        if (isNaN(value.getTime())) throw new Error(`Invalid date: ${value}`);
        return value.toISOString().split("T")[0];
      }
      if (typeof value === "string") {
        const parsed = new Date(value);
        if (isNaN(parsed.getTime())) throw new Error(`Invalid date format: ${value}`);
        return parsed.toISOString().split("T")[0];
      }
      if (typeof value === "number") {
        // Excel serial date fallback
        const epoch = new Date(1899, 11, 30);
        epoch.setDate(epoch.getDate() + value);
        return epoch.toISOString().split("T")[0];
      }
      throw new Error(`Unsupported date type: ${value}`);
    }
    function parseCurrency(value) {
      // ExcelJS returns formula cells as { formula, result }
      if (value !== null && typeof value === "object" && "result" in value) value = value.result;
      if (typeof value === "number") return value;
      if (typeof value === "string") {
        const n = Number(value.replace(/,/g, "").trim());
        if (isNaN(n)) throw new Error(`Invalid currency value: ${value}`);
        return n;
      }
      throw new Error(`Unsupported currency format: ${value}`);
    }

    return {
      qo_number:      String(getCell("J4")).trim(),
      company_name:   String(getCell("B7")).trim(),
      contact_person: String(getCell("B6")).trim(),
      project_name:   String(getCell("I6")).trim(),
      validity:       parseCurrency(getCell("I9")),
      sales_agent:    String(getCell("F42")).trim(),
      total_price:    parseCurrency(getCell("J35")),
      create_date:    toISO(getCell("K5")),
      stage:          "On track",
      reason:         "",
    };
  }

  async function refreshData() {
    // Re-fetch stats (quarter-filtered)
    try {
      const { from, to } = quarterDateRange(selectedYear, selectedQ);
      const { data, error: fetchErr } = await supabase
        .from("sales_pipeline")
        .select('qo_number, company_name, contact_person, project_name, total_price, validity, sales_agent, stage, create_date, reason, po_qt, follow_up_1, follow_up_2, follow_up_3, revision')
        .gte('create_date', from)
        .lte('create_date', to)
        .order("id", { ascending: true });
      if (!fetchErr) {
        setRawData(data.map(r => ({
          qo:          r.qo_number      ?? "",
          company:     r.company_name   ?? "",
          contact:     r.contact_person ?? "",
          project:     r.project_name   ?? "",
          price:       parseFloat(r.total_price) || 0,
          agent:       r.sales_agent    ?? "",
          stage:       r.stage          ?? "",
          date:        r.create_date    ?? "",
          reason:      r.reason         ?? "",
          validity:    r.validity       ?? null,
          po_qt:       r.po_qt          ?? "",
          follow_up_1: r.follow_up_1    ?? "",
          follow_up_2: r.follow_up_2    ?? "",
          follow_up_3: r.follow_up_3    ?? "",
          revision:    r.revision       ?? 0,
        })));
      }
    } catch (err) {
      console.error("refreshData stats error:", err);
    }
    // Trigger table tab re-fetch via key bump
    setTableRefreshKey(k => k + 1);
  }

  async function handleUpload(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setUploadState("loading");
    setUploadMessage("");
    try {
      const buffer = await file.arrayBuffer();
      const row = await extractQuotationData(buffer);
      // Add extra fields with defaults so user can fill them in
      const fullRow = { ...row, po_qt: "", follow_up_1: "", follow_up_2: "", follow_up_3: "", revision: 0 };
      setReviewRow(fullRow);
      setUploadState("idle");
    } catch (err) {
      console.error("Upload error:", err);
      setUploadState("error");
      setUploadMessage(err.message);
      setTimeout(() => setUploadState("idle"), 5000);
    }
  }

  async function handleConfirmSave(row) {
    setReviewRow(null);
    setUploadState("loading");
    try {
      const { error } = await supabase
        .from("sales_pipeline")
        .upsert(row, { onConflict: "qo_number" });
      if (error) throw error;
      setUploadState("success");
      setUploadMessage(`✓ ${row.qo_number} saved`);
      await refreshData();
      setTimeout(() => setUploadState("idle"), 3000);
    } catch (err) {
      console.error("Save error:", err);
      setUploadState("error");
      setUploadMessage(err.message);
      setTimeout(() => setUploadState("idle"), 5000);
    }
  }

  async function handleCreateRow() {
    const today = new Date().toISOString().split("T")[0];
    const newRow = {
      qo_number:      `NEW-${Date.now()}`,
      company_name:   "",
      contact_person: "",
      project_name:   "",
      total_price:    0,
      validity:       30,
      sales_agent:    agents[0] ?? "",
      stage:          "On track",
      create_date:    today,
      reason:         "",
      po_qt:          "",
      follow_up_1:    "",
      follow_up_2:    "",
      follow_up_3:    "",
      revision:       0,
    };
    setReviewRow(newRow);
  }

  async function handleCellSave(qoNumber, field, value) {
    setSavingCell(`${qoNumber}-${field}`);
    try {
      const { error } = await supabase
        .from("sales_pipeline")
        .update({ [field]: value })
        .eq("qo_number", qoNumber);
      if (error) throw error;
      await refreshData();
    } catch (err) {
      console.error("Cell save error:", err);
    } finally {
      setSavingCell(null);
      setEditingCell(null);
    }
  }

  async function handleDeleteRow(qoNumber) {
    setDeletingRow(qoNumber);
    try {
      const { error } = await supabase
        .from("sales_pipeline")
        .delete()
        .eq("qo_number", qoNumber);
      if (error) throw error;
      await refreshData();
    } catch (err) {
      console.error("Delete error:", err);
    } finally {
      setDeletingRow(null);
      setDeleteConfirm(null);
    }
  }

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

        @keyframes spin { to { transform: rotate(360deg); } }

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
            <label style={{
              background: uploadState === "success" ? "#10b981" : uploadState === "error" ? "#ef4444" : uploadState === "loading" ? "#94a3b8" : "#2563eb",
              color: "#fff", border: "none", borderRadius: 8, padding: "5px 12px",
              fontSize: 12, fontWeight: 600, cursor: uploadState === "loading" ? "wait" : "pointer",
              fontFamily: "inherit", display: "inline-flex", alignItems: "center", gap: 5, transition: "background 0.2s",
            }}>
              {uploadState === "loading" ? "⏳ Uploading…" : uploadState === "success" ? uploadMessage : uploadState === "error" ? "✕ Error" : "⬆ Upload QO"}
              <input type="file" accept=".xlsx" onChange={handleUpload} style={{ display: "none" }} disabled={uploadState === "loading"} />
            </label>
            {uploadState === "error" && (
              <span style={{ fontSize: 11, color: "#ef4444", maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={uploadMessage}>{uploadMessage}</span>
            )}
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
                <div>
                  <div style={{ fontSize: 12, color: "#475569", fontWeight: 500, marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.5px" }}>Quarterly Revenue Goal · {quarterLabel}</div>
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
                { label: "Total Quotations", value: allForQ.length, sub: `All statuses ${fmt(allStatusTotal)}`, icon: "📋", color: "#3b82f6" },
                { label: "Orders Confirmed", value: stageCounts.Order, sub: `${fmt(orderTotal)}`, icon: "✅", color: "#10b981" },
                { label: "In Progress", value: stageCounts["On track"], sub: `${fmt(onTrackTotal)} potential`, icon: "🔄", color: "#f59e0b" },
                { label: "Failed", value: stageCounts.Fail, sub: `${fmt(failedTotal)} lost`, icon: "❌", color: "#ef4444" },
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
                      {agentData.map((entry, i) => <Cell key={i} fill={getAgentColor(entry.agent, agents)} />)}
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
                <div key={agent.agent} className="card" style={{ borderTop: `3px solid ${getAgentColor(agent.agent, agents)}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                    <div>
                      <div style={{ fontSize: 16, fontWeight: 700, fontFamily: "'Space Grotesk', sans-serif", color: getAgentColor(agent.agent, agents) }}>{agent.agent}</div>
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
                        <td style={{ padding: "10px 12px" }}><span style={{ color: getAgentColor(row.agent, agents), fontWeight: 500 }}>{row.agent}</span></td>
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
                  <Tooltip
                    contentStyle={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 8, fontSize: 12 }}
                    content={({ active, payload, label }) => {
                      if (!active || !payload?.length) return null;
                      const d = monthlyData.find(m => m.month === label);
                      const entries = [
                        { name: "Orders",   value: payload.find(p => p.dataKey === "Order")?.value ?? 0,        count: d?.OrderCount ?? 0,   color: "#10b981" },
                        { name: "On Track", value: payload.find(p => p.dataKey === "On track")?.value ?? 0,     count: d?.OnTrackCount ?? 0, color: "#f59e0b" },
                        { name: "Failed",   value: payload.find(p => p.dataKey === "Fail")?.value ?? 0,         count: d?.FailCount ?? 0,    color: "#ef4444" },
                      ];
                      return (
                        <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, padding: "10px 14px", fontSize: 12 }}>
                          <div style={{ fontWeight: 700, color: "#0f172a", marginBottom: 8 }}>{label}</div>
                          {entries.map(e => (
                            <div key={e.name} style={{ display: "flex", justifyContent: "space-between", gap: 16, marginBottom: 4 }}>
                              <span style={{ color: e.color, fontWeight: 600 }}>● {e.name}</span>
                              <span style={{ color: "#0f172a" }}>฿{e.value.toFixed(2)}M <span style={{ color: "#94a3b8" }}>({e.count} QO)</span></span>
                            </div>
                          ))}
                        </div>
                      );
                    }}
                  />
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
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 12, color: "#475569" }}>Orders</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: "#10b981" }}>
                        ฿{m.Order.toFixed(2)}M <span style={{ fontSize: 11, fontWeight: 400, color: "#94a3b8" }}>({m.OrderCount} QO)</span>
                      </span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 12, color: "#475569" }}>On Track</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: "#f59e0b" }}>
                        ฿{m["On track"].toFixed(2)}M <span style={{ fontSize: 11, fontWeight: 400, color: "#94a3b8" }}>({m.OnTrackCount} QO)</span>
                      </span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 12, color: "#475569" }}>Failed</span>
                      <span style={{ fontSize: 12, fontWeight: 600, color: "#ef4444" }}>
                        ฿{m.Fail.toFixed(2)}M <span style={{ fontSize: 11, fontWeight: 400, color: "#94a3b8" }}>({m.FailCount} QO)</span>
                      </span>
                    </div>
                    <div style={{ borderTop: "1px solid #e2e8f0", paddingTop: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <span style={{ fontSize: 12, color: "#475569", fontWeight: 600 }}>Total Active</span>
                      <span style={{ fontSize: 12, fontWeight: 700, color: "#3b82f6" }}>
                        ฿{(m.Order + m["On track"]).toFixed(2)}M <span style={{ fontSize: 11, fontWeight: 400, color: "#94a3b8" }}>({m.OrderCount + m.OnTrackCount} QO)</span>
                      </span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* DATA TABLE TAB */}
        {activeTab === "table" && (() => {
          const TABLE_COLS = [
            { key: "qo_number",      label: "QO Number",    width: 110, editable: false },
            { key: "create_date",    label: "Date",         width: 110, editable: true,  type: "date" },
            { key: "company_name",   label: "Company",      width: 200, editable: true,  type: "text" },
            { key: "contact_person", label: "Contact",      width: 150, editable: true,  type: "text" },
            { key: "project_name",   label: "Project",      width: 200, editable: true,  type: "text" },
            { key: "total_price",    label: "Price (฿)",    width: 110, editable: true,  type: "number" },
            { key: "revision",       label: "Revision",     width: 80,  editable: true,  type: "number" },
            { key: "sales_agent",    label: "Agent",        width: 110, editable: true,  type: "agent" },
            { key: "stage",          label: "Stage",        width: 100, editable: true,  type: "stage" },
            { key: "po_qt",          label: "PO/QT",        width: 90,  editable: true,  type: "text" },
            { key: "validity",       label: "Validity",     width: 80,  editable: true,  type: "number" },
            { key: "reason",         label: "Reason",       width: 160, editable: true,  type: "text" },
            { key: "follow_up_1",    label: "Follow Up 1",  width: 130, editable: true,  type: "text" },
            { key: "follow_up_2",    label: "Follow Up 2",  width: 130, editable: true,  type: "text" },
            { key: "follow_up_3",    label: "Follow Up 3",  width: 130, editable: true,  type: "text" },
          ];

          // Server-side data — tableRows already paginated/filtered/sorted by Supabase
          const totalPages = Math.max(1, Math.ceil(tableTotal / TABLE_PAGE_SIZE));
          const page = Math.min(tablePage, totalPages);

          // Cycle: none → asc → desc → none
          const toggleSort = key => {
            setTableSort(prev => {
              const existing = prev.find(s => s.key === key);
              if (!existing) return [...prev, { key, dir: "asc" }];
              if (existing.dir === "asc") return prev.map(s => s.key === key ? { key, dir: "desc" } : s);
              return prev.filter(s => s.key !== key);
            });
            setTablePage(1);
          };

          const SortIcon = ({ colKey }) => {
            const entry = tableSort.find(s => s.key === colKey);
            const priority = tableSort.findIndex(s => s.key === colKey) + 1;
            if (!entry) return <span style={{ opacity: 0.2, marginLeft: 4, fontSize: 10 }}>↕</span>;
            return (
              <span style={{ marginLeft: 4, color: "#2563eb", fontSize: 10, fontWeight: 700 }}>
                {entry.dir === "asc" ? "↑" : "↓"}
                {tableSort.length > 1 && <sup style={{ fontSize: 8 }}>{priority}</sup>}
              </span>
            );
          };

          const stageColor = s => s === "Order" ? "#10b981" : s === "On track" ? "#f59e0b" : s === "Fail" ? "#ef4444" : "#94a3b8";
          
          function Paging() {
            return (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 12, color: "#64748b" }}>
                  {tableLoading ? "Loading…" : `${tableTotal} rows · Page ${page} of ${totalPages}`}
                </span>
                <button onClick={() => setTablePage(p => Math.max(1, p - 1))} disabled={page === 1 || tableLoading}
                  style={{ padding: "5px 10px", borderRadius: 6, border: "1px solid #e2e8f0", background: "#fff", cursor: (page === 1 || tableLoading) ? "not-allowed" : "pointer", fontSize: 12, color: (page === 1 || tableLoading) ? "#cbd5e1" : "#475569" }}>‹ Prev</button>
                <button onClick={() => setTablePage(p => Math.min(totalPages, p + 1))} disabled={page === totalPages || tableLoading}
                  style={{ padding: "5px 10px", borderRadius: 6, border: "1px solid #e2e8f0", background: "#fff", cursor: (page === totalPages || tableLoading) ? "not-allowed" : "pointer", fontSize: 12, color: (page === totalPages || tableLoading) ? "#cbd5e1" : "#475569" }}>Next ›</button>
              </div>
            )
          }

          function CellContent({ row, col }) {
            const cellId = `${row.qo_number}-${col.key}`;
            const isEditing = editingCell?.id === row.qo_number && editingCell?.field === col.key;
            const isSaving = savingCell === cellId;
            const val = row[col.key];

            if (!col.editable) return (
              <span style={{ fontSize: 12, color: "#475569" }}>{val ?? "—"}</span>
            );

            if (isEditing) {
              const commitEdit = () => {
                let saveVal = col.type === "number" ? parseFloat(editingValue) || 0 : editingValue;
                if (col.type === "date") saveVal = displayToISO(editingValue);
                if (String(saveVal) !== String(val ?? "")) {
                  handleCellSave(row.qo_number, col.key, saveVal);
                } else {
                  setEditingCell(null);
                }
              };
              if (col.type === "date") return (
                <input autoFocus type="date" value={editingValue}
                  onChange={e => setEditingValue(e.target.value)}
                  onBlur={commitEdit}
                  onKeyDown={e => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") setEditingCell(null); }}
                  style={{ fontSize: 12, border: "1px solid #6366f1", borderRadius: 4, padding: "2px 4px", width: "100%", boxSizing: "border-box" }} />
              );
              if (col.type === "stage") return (
                <select autoFocus value={editingValue} onBlur={commitEdit}
                  onChange={e => { setEditingValue(e.target.value); }}
                  onKeyDown={e => e.key === "Escape" && setEditingCell(null)}
                  style={{ fontSize: 12, border: "1px solid #6366f1", borderRadius: 4, padding: "2px 4px", width: "100%", background: "#fff" }}>
                  {["On track","Order","Fail"].map(o => <option key={o}>{o}</option>)}
                </select>
              );
              if (col.type === "agent") return (
                <>
                  <input autoFocus list="agent-options-table" value={editingValue}
                    onChange={e => setEditingValue(e.target.value)}
                    onBlur={commitEdit}
                    onKeyDown={e => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") setEditingCell(null); }}
                    style={{ fontSize: 12, border: "1px solid #6366f1", borderRadius: 4, padding: "2px 4px", width: "100%", boxSizing: "border-box" }} />
                  <datalist id="agent-options-table">{agents.map(a => <option key={a} value={a} />)}</datalist>
                </>
              );
              return (
                <input autoFocus type={col.type === "number" ? "number" : "text"} value={editingValue}
                  onChange={e => setEditingValue(e.target.value)}
                  onBlur={commitEdit}
                  onKeyDown={e => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") setEditingCell(null); }}
                  style={{ fontSize: 12, border: "1px solid #6366f1", borderRadius: 4, padding: "2px 4px", width: "100%", boxSizing: "border-box" }} />
              );
            }

            const displayVal = col.key === "total_price"
              ? (val ? `฿${Number(val).toLocaleString()}` : "—")
              : col.type === "date"
              ? (isoToDisplay(val) || "—")
              : (val ?? "—");

            return (
              <span
                onClick={() => {
                  if (isSaving) return;
                  setEditingCell({ id: row.qo_number, field: col.key });
                  setEditingValue(col.type === "date" ? (String(val ?? "").substring(0, 10)) : (val ?? ""));
                }}
                title="Click to edit"
                style={{
                  fontSize: 12, display: "block", cursor: "text", borderRadius: 4,
                  padding: "2px 4px", margin: "-2px -4px",
                  color: col.key === "stage" ? stageColor(val) : col.key === "sales_agent" ? getAgentColor(val, agents) : "#0f172a",
                  fontWeight: col.key === "stage" || col.key === "sales_agent" ? 600 : 400,
                  background: isSaving ? "#f1f5f9" : "transparent",
                  transition: "background 0.15s",
                }}
                onMouseEnter={e => { if (!isSaving) e.currentTarget.style.background = "#f8fafc"; }}
                onMouseLeave={e => { e.currentTarget.style.background = isSaving ? "#f1f5f9" : "transparent"; }}
              >
                {isSaving ? "⏳" : displayVal}
              </span>
            );
          }

          return (
            <div>
              {/* Search + pagination controls */}
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <input
                    value={tableSearch} onChange={e => { setTableSearch(e.target.value); setTablePage(1); }}
                    placeholder="Search by QO, company, agent, stage…"
                    style={{ padding: "8px 12px", borderRadius: 8, border: "1px solid #e2e8f0", fontSize: 13, fontFamily: "inherit", width: 260, color: "#0f172a", outline: "none" }}
                  />
                  <button onClick={handleCreateRow}
                    style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: "#2563eb", color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>
                    + New Row
                  </button>
                  {tableSort.length > 0 && (
                    <button onClick={() => setTableSort([])}
                      style={{ padding: "8px 12px", borderRadius: 8, border: "1px solid #e2e8f0", background: "#fff", fontSize: 12, color: "#64748b", cursor: "pointer", fontFamily: "inherit", whiteSpace: "nowrap" }}>
                      ✕ Clear sort ({tableSort.length})
                    </button>
                  )}
                </div>
                <Paging />
              </div>

              {/* Table */}
              <div style={{ overflowX: "auto", borderRadius: 12, border: "1px solid #e2e8f0", background: "#fff", position: "relative" }}>
                {tableLoading && (
                  <div style={{ position: "absolute", inset: 0, background: "rgba(255,255,255,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 10, borderRadius: 12 }}>
                    <div style={{ width: 28, height: 28, border: "3px solid #e2e8f0", borderTop: "3px solid #6366f1", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
                  </div>
                )}
                <table style={{ borderCollapse: "collapse", width: "100%", minWidth: TABLE_COLS.reduce((s, c) => s + c.width, 0) + 48 }}>
                  <thead>
                    <tr style={{ background: "#f8fafc", borderBottom: "1px solid #e2e8f0" }}>
                      {TABLE_COLS.map(col => (
                        <th key={col.key} onClick={() => toggleSort(col.key)}
                          style={{ padding: "10px 12px", textAlign: "left", fontSize: 11, fontWeight: 600, color: tableSort.some(s => s.key === col.key) ? "#2563eb" : "#475569", textTransform: "uppercase", letterSpacing: "0.4px", whiteSpace: "nowrap", width: col.width, cursor: "pointer", userSelect: "none" }}>
                          {col.label}<SortIcon colKey={col.key} />
                        </th>
                      ))}
                      <th style={{ width: 48 }} />
                    </tr>
                  </thead>
                  <tbody>
                    {tableRows.length === 0 && !tableLoading ? (
                      <tr><td colSpan={TABLE_COLS.length + 1} style={{ padding: 32, textAlign: "center", color: "#94a3b8", fontSize: 13 }}>No results found</td></tr>
                    ) : tableRows.map((row, i) => (
                      <tr key={row.qo_number} style={{ borderBottom: "1px solid #f1f5f9", background: i % 2 === 0 ? "#fff" : "#fafbfc" }}>
                        {TABLE_COLS.map(col => (
                          <td key={col.key} style={{ padding: "8px 12px", maxWidth: col.width, overflow: "hidden" }}>
                            <CellContent row={row} col={col} />
                          </td>
                        ))}
                        <td style={{ padding: "4px 8px", textAlign: "center" }}>
                          {deletingRow === row.qo_number ? (
                            <span style={{ display: "inline-block", width: 14, height: 14, border: "2px solid #fca5a5", borderTop: "2px solid #ef4444", borderRadius: "50%", animation: "spin 0.6s linear infinite", verticalAlign: "middle" }} />
                          ) : (
                            <button
                              onClick={() => setDeleteConfirm({ qo_number: row.qo_number, company_name: row.company_name })}
                              title="Delete row"
                              style={{ background: "none", border: "none", cursor: "pointer", padding: "4px 6px", borderRadius: 6, color: "#cbd5e1", fontSize: 14, lineHeight: 1 }}
                              onMouseEnter={e => { e.currentTarget.style.color = "#ef4444"; e.currentTarget.style.background = "#fef2f2"; }}
                              onMouseLeave={e => { e.currentTarget.style.color = "#cbd5e1"; e.currentTarget.style.background = "none"; }}
                            >🗑</button>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div style={{ marginTop: 8, fontSize: 11, color: "#94a3b8" }}>💡 Click any cell to edit · Changes save instantly</div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "right", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
                <Paging />
              </div>
            </div>
          );
        })()}
      </div>

      <div style={{ textAlign: "center", padding: "16px", color: "#94a3b8", fontSize: 11 }}>
        PC Sales Pipeline Dashboard · {tableTotal > 0 ? `${tableTotal} total records` : `${(RAW_DATA ?? []).length} records this quarter`}
      </div>

      {/* ── Delete Confirmation Modal ───────────────────────────────────── */}
      {deleteConfirm && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1100, padding: 16 }}
          onClick={e => { if (e.target === e.currentTarget && !deletingRow) setDeleteConfirm(null); }}>
          <div style={{ background: "#fff", borderRadius: 16, padding: "28px 32px", maxWidth: 400, width: "100%", boxShadow: "0 24px 64px rgba(0,0,0,0.18)", textAlign: "center" }}>
            <div style={{ width: 48, height: 48, borderRadius: "50%", background: "#fef2f2", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, margin: "0 auto 16px" }}>🗑</div>
            <div style={{ fontWeight: 700, fontSize: 17, color: "#0f172a", marginBottom: 8 }}>Delete Row?</div>
            <div style={{ fontSize: 13, color: "#64748b", marginBottom: 6 }}>You are about to permanently delete:</div>
            <div style={{ fontSize: 13, fontWeight: 600, color: "#0f172a", marginBottom: 4 }}>{deleteConfirm.company_name || "—"}</div>
            <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 20, fontFamily: "monospace" }}>{deleteConfirm.qo_number}</div>
            <div style={{ fontSize: 12, color: "#ef4444", marginBottom: 24 }}>This action cannot be undone.</div>
            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              <button onClick={() => setDeleteConfirm(null)} disabled={!!deletingRow}
                style={{ padding: "9px 20px", borderRadius: 8, border: "1px solid #e2e8f0", background: "#fff", fontSize: 13, fontWeight: 500, cursor: deletingRow ? "not-allowed" : "pointer", color: "#475569", opacity: deletingRow ? 0.5 : 1 }}>
                Cancel
              </button>
              <button onClick={() => handleDeleteRow(deleteConfirm.qo_number)} disabled={!!deletingRow}
                style={{ padding: "9px 20px", borderRadius: 8, border: "none", background: "#ef4444", color: "#fff", fontSize: 13, fontWeight: 600, cursor: deletingRow ? "not-allowed" : "pointer", opacity: deletingRow ? 0.8 : 1, display: "inline-flex", alignItems: "center", gap: 8 }}>
                {deletingRow ? (
                  <>
                    <span style={{ display: "inline-block", width: 13, height: 13, border: "2px solid rgba(255,255,255,0.4)", borderTop: "2px solid #fff", borderRadius: "50%", animation: "spin 0.6s linear infinite" }} />
                    Deleting…
                  </>
                ) : "Delete"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Review & Edit Modal ─────────────────────────────────────────── */}
      {reviewRow && (
        <ReviewModal
          row={reviewRow}
          agents={agents}
          onSave={handleConfirmSave}
          onCancel={() => { setReviewRow(null); setUploadState("idle"); }}
        />
      )}
    </div>
  );
}