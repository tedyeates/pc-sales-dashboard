import { useState } from "react";
import { FIELDS } from "../constants";

export default function ReviewModal({ row, agents, onSave, onCancel }) {
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
