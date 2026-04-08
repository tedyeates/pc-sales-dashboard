import { useState } from "react";
import { supabase } from "../supabaseClient";
import { authPageStyle, authCardStyle, inputStyle, btnPrimaryStyle, labelStyle, errStyle, okStyle } from "../styles";

export default function SetPasswordPage({ onDone }) {
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
