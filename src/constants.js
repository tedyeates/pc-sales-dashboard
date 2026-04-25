export const GOAL = 30000000;

export const STAGE_COLORS = { Order: "#10b981", "On track": "#f59e0b", Fail: "#ef4444" };

export const AGENT_PALETTE = ["#6366f1","#ec4899","#14b8a6","#f59e0b","#10b981","#3b82f6","#8b5cf6","#ef4444","#06b6d4","#84cc16"];

export const FIELDS = [
  { key: "qo_number",      label: "QO Number",      type: "text" },
  { key: "company_name",   label: "Company Name",   type: "text" },
  { key: "contact_person", label: "Contact Person", type: "text" },
  { key: "project_name",   label: "Project Name",   type: "text" },
  { key: "total_price",    label: "Total Price",    type: "number" },
  { key: "price_cabinet",  label: "Price Cabinet",  type: "number" },
  { key: "price_wire_busbar", label: "Price Wire/Busbar", type: "number" },
  { key: "price_equipment", label: "Price Equipment", type: "number" },
  { key: "revision",       label: "Revision",       type: "number" },
  { key: "validity",       label: "Validity (days)",type: "number" },
  { key: "sales_agent",    label: "Sales Agent",    type: "select", options: [] },
  { key: "stage",          label: "Stage",          type: "select", options: ["On track", "Order", "Fail"] },
  { key: "create_date",    label: "Create Date",    type: "date" },
  { key: "reason",         label: "Reason",         type: "text" },
  { key: "po_qt",          label: "PO / QT",        type: "text" },
  { key: "follow_up_1",    label: "Follow Up 1",    type: "text" },
  { key: "follow_up_2",    label: "Follow Up 2",    type: "text" },
  { key: "follow_up_3",    label: "Follow Up 3",    type: "text" },
];

export const TABLE_COLS = [
  { key: "qo_number",      label: "QO Number",    width: 110, editable: false },
  { key: "create_date",    label: "Date",         width: 110, editable: true,  type: "date" },
  { key: "company_name",   label: "Company",      width: 200, editable: true,  type: "text" },
  { key: "contact_person", label: "Contact",      width: 150, editable: true,  type: "text" },
  { key: "project_name",   label: "Project",      width: 200, editable: true,  type: "text" },
  { key: "total_price",    label: "Price (฿)",    width: 110, editable: true,  type: "number" },
  { key: "price_cabinet",  label: "Cabinet",      width: 100, editable: true,  type: "number" },
  { key: "price_wire_busbar", label: "Wire/Busbar", width: 100, editable: true, type: "number" },
  { key: "price_equipment", label: "Equipment",    width: 100, editable: true,  type: "number" },
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

export const TABLE_PAGE_SIZE = 20;

// Derived: Supabase select string from FIELDS keys
export const SELECT_COLUMNS = FIELDS.map(f => f.key).join(", ");

// Derived: DB column mapping (keys map to themselves)
export const TABLE_COL_TO_DB = Object.fromEntries(FIELDS.map(f => [f.key, f.key]));

// Derived: default empty row from FIELDS
export const DEFAULT_ROW = Object.fromEntries(FIELDS.map(f => [
  f.key,
  f.type === "number" ? 0 : f.type === "date" ? "" : f.key === "stage" ? "On track" : "",
]));

// Derived: keys that should display as currency in the table
export const CURRENCY_KEYS = new Set(
  FIELDS.filter(f => f.key.startsWith("price_") || f.key === "total_price").map(f => f.key)
);
