import { useState, useEffect } from "react";
import { supabase } from "./supabaseClient";
import LoginPage from "./components/LoginPage";
import SetPasswordPage from "./components/SetPasswordPage";
import Dashboard from "./components/Dashboard";

export default function App() {
  const [session, setSession] = useState(undefined);
  const [isSettingPassword, setIsSettingPassword] = useState(false);

  useEffect(() => {
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
