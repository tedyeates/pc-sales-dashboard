export function parseDate(dateStr) {
  if (!dateStr || !String(dateStr).trim()) return null;
  const s = String(dateStr).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const [y, m, d] = s.substring(0, 10).split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  let norm = s.replace(/-/g, '/');
  const parts = norm.split('/');
  if (parts.length !== 3) return null;
  let [d, m, y] = parts.map(Number);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  if (y < 100) y += 2000;
  return new Date(y, m - 1, d);
}

export function getQuarter(date) {
  const mo = date.getMonth() + 1;
  if (mo <= 3) return 1;
  if (mo <= 6) return 2;
  if (mo <= 9) return 3;
  return 4;
}

export function getYear(date) { return date.getFullYear(); }

export function isoToDisplay(iso) {
  if (!iso) return "";
  const s = String(iso).trim().substring(0, 10);
  const [y, m, d] = s.split("-");
  if (!y || !m || !d) return iso;
  return `${d}/${m}/${y}`;
}

export function displayToISO(val) {
  if (!val) return null;
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parts = s.split("/");
  if (parts.length === 3) {
    let [d, m, y] = parts.map(Number);
    if (y < 100) y += 2000;
    if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
    return `${String(y).padStart(4,"0")}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
  }
  return null;
}

export function getAgentColor(agent, agentList) {
  const AGENT_PALETTE = ["#6366f1","#ec4899","#14b8a6","#f59e0b","#10b981","#3b82f6","#8b5cf6","#ef4444","#06b6d4","#84cc16"];
  const i = agentList.indexOf(agent);
  return AGENT_PALETTE[i % AGENT_PALETTE.length];
}

export const fmt = (n) => n >= 1000000 ? `฿${(n/1000000).toFixed(2)}M` : n >= 1000 ? `฿${(n/1000).toFixed(0)}K` : `฿${n.toFixed(0)}`;
