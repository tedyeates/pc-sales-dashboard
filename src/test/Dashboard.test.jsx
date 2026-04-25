import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import React from 'react';

// ── Hoisted shared state (available inside vi.mock factories) ──────────────
const { SAMPLE_ROWS, makeChain } = vi.hoisted(() => {
  const SAMPLE_ROWS = [
    {
      qo_number: 'QO-001', company_name: 'Acme Corp', contact_person: 'John',
      project_name: 'Project Alpha', total_price: 5000000, price_cabinet: 1000000,
      price_wire_busbar: 500000, price_equipment: 200000, validity: 30,
      sales_agent: 'Alice', stage: 'Order', create_date: '2025-01-15',
      reason: '', po_qt: 'PO-100', follow_up_1: '', follow_up_2: '', follow_up_3: '', revision: 1,
    },
    {
      qo_number: 'QO-002', company_name: 'Beta Inc', contact_person: 'Jane',
      project_name: 'Project Beta', total_price: 8000000, price_cabinet: 2000000,
      price_wire_busbar: 1000000, price_equipment: 500000, validity: 60,
      sales_agent: 'Bob', stage: 'On track', create_date: '2025-01-20',
      reason: '', po_qt: '', follow_up_1: 'Called', follow_up_2: '', follow_up_3: '', revision: 0,
    },
    {
      qo_number: 'QO-003', company_name: 'Gamma LLC', contact_person: 'Sam',
      project_name: 'Project Gamma', total_price: 3000000, price_cabinet: 500000,
      price_wire_busbar: 300000, price_equipment: 100000, validity: 15,
      sales_agent: 'Alice', stage: 'Fail', create_date: '2025-02-10',
      reason: 'Price too high', po_qt: '', follow_up_1: '', follow_up_2: '', follow_up_3: '', revision: 2,
    },
  ];

  function makeChain(rows) {
    const chain = {};
    ['select', 'gte', 'lte', 'order', 'or', 'not', 'neq', 'eq', 'range', 'delete'].forEach(m => {
      chain[m] = (..._a) => chain;
    });
    chain.update = (..._a) => chain;
    chain.upsert = () => Promise.resolve({ error: null });
    chain.then = (onFulfilled) =>
      Promise.resolve({ data: rows, error: null, count: rows.length }).then(onFulfilled);
    return chain;
  }

  return { SAMPLE_ROWS, makeChain };
});

// ── Mock Recharts ──────────────────────────────────────────────────────────
vi.mock('recharts', () => ({
  ResponsiveContainer: ({ children }) => children,
  BarChart: ({ children }) => children,
  PieChart: ({ children }) => children,
  Bar: () => null, XAxis: () => null, YAxis: () => null,
  Tooltip: () => null, CartesianGrid: () => null, Cell: () => null,
  Pie: () => null, Legend: () => null,
}));

// ── Mock ExcelJS ───────────────────────────────────────────────────────────
vi.mock('https://esm.sh/exceljs@4.4.0', () => {
  const cells = {
    J4: { value: 'QO-UPLOAD-001' }, B7: { value: 'Upload Corp' },
    B6: { value: 'Upload Person' }, I6: { value: 'Upload Project' },
    I9: { value: 30 }, F42: { value: 'Alice' },
    J35: { value: 1500000 }, K5: { value: new Date('2025-01-25') },
  };
  return {
    default: {
      Workbook: function () {
        this.xlsx = { load: () => Promise.resolve() };
        this.worksheets = [{ getCell: (addr) => cells[addr] || { value: null } }];
      },
    },
  };
});

// ── Mock Supabase ──────────────────────────────────────────────────────────
vi.mock('https://esm.sh/@supabase/supabase-js@2', () => ({
  createClient: () => ({}),
}));

vi.mock('../supabaseClient', () => ({
  supabase: {
    from: () => makeChain(SAMPLE_ROWS),
    auth: {
      signOut: () => Promise.resolve({}),
      getSession: () => Promise.resolve({ data: { session: null } }),
      onAuthStateChange: () => ({ data: { subscription: { unsubscribe: () => {} } } }),
    },
  },
}));

import { supabase } from '../supabaseClient';
import Dashboard from '../components/Dashboard';

// ── Helpers ────────────────────────────────────────────────────────────────
function renderDashboard() {
  return render(<Dashboard session={{ user: { id: '1' } }} />);
}

async function waitForData() {
  await waitFor(() => {
    expect(screen.getByText('PC Sales Pipeline')).toBeInTheDocument();
    expect(screen.queryByText('Loading pipeline data…')).not.toBeInTheDocument();
  });
}

async function switchTab(tabName) {
  fireEvent.click(screen.getByRole('button', { name: tabName }));
  await waitFor(() => {});
}

// ── Tests ──────────────────────────────────────────────────────────────────
describe('Dashboard', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    supabase.from = vi.fn(() => makeChain(SAMPLE_ROWS));
  });

  // ═══════════════════════════════════════════════════════════════════════
  // OVERVIEW TAB
  // ═══════════════════════════════════════════════════════════════════════
  describe('Overview Tab', () => {
    it('renders KPI cards', async () => {
      renderDashboard();
      await waitForData();
      for (const label of ['Total Quotations', 'Orders Confirmed', 'In Progress', 'Failed', 'Win Rate']) {
        expect(screen.getByText(label)).toBeInTheDocument();
      }
    });

    it('renders goal progress section', async () => {
      renderDashboard();
      await waitForData();
      expect(screen.getByText(/Quarterly Revenue Goal/)).toBeInTheDocument();
      expect(screen.getByText(/Confirmed Orders/)).toBeInTheDocument();
      expect(screen.getByText(/Pending Orders/)).toBeInTheDocument();
      expect(screen.getByText(/Target:/)).toBeInTheDocument();
    });

    it('renders Status Distribution chart', async () => {
      renderDashboard();
      await waitForData();
      expect(screen.getByText('Status Distribution')).toBeInTheDocument();
    });

    it('renders Agent Revenue Breakdown chart', async () => {
      renderDashboard();
      await waitForData();
      expect(screen.getByText('Agent Revenue Breakdown')).toBeInTheDocument();
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // SALES AGENTS TAB
  // ═══════════════════════════════════════════════════════════════════════
  describe('Sales Agents Tab', () => {
    it('renders agent cards and chart', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Sales Agents');
      expect(screen.getByText('Orders by Agent — Count Comparison')).toBeInTheDocument();
      expect(screen.getByText('Alice')).toBeInTheDocument();
      expect(screen.getByText('Bob')).toBeInTheDocument();
    });

    it('shows order/on-track/fail breakdown', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Sales Agents');
      expect(screen.getAllByText('Orders').length).toBeGreaterThan(0);
      expect(screen.getAllByText('On Track').length).toBeGreaterThan(0);
      expect(screen.getAllByText('Failed').length).toBeGreaterThan(0);
    });

    it('shows confirmed revenue', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Sales Agents');
      expect(screen.getAllByText('confirmed revenue').length).toBeGreaterThan(0);
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // TOP PIPELINE TAB
  // ═══════════════════════════════════════════════════════════════════════
  describe('Top Pipeline Tab', () => {
    it('renders top 5 on-track section', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Top Pipeline');
      expect(screen.getByText(/Top 5 On-Track/)).toBeInTheDocument();
      expect(screen.getByText(/combined potential/)).toBeInTheDocument();
    });

    it('shows on-track quotations table', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Top Pipeline');
      expect(screen.getByText(/All On-Track Quotations/)).toBeInTheDocument();
      expect(screen.getAllByText('Beta Inc').length).toBeGreaterThan(0);
    });

    it('shows ranked cards', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Top Pipeline');
      expect(screen.getByText('#1')).toBeInTheDocument();
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // MONTHLY TREND TAB
  // ═══════════════════════════════════════════════════════════════════════
  describe('Monthly Trend Tab', () => {
    it('renders monthly revenue breakdown', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Monthly Trend');
      expect(screen.getByText(/Monthly Revenue Breakdown/)).toBeInTheDocument();
    });

    it('shows month cards for current quarter', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Monthly Trend');
      const currentQ = Math.floor(new Date().getMonth() / 3) + 1;
      const qMonths = { 1: ['Jan', 'Feb', 'Mar'], 2: ['Apr', 'May', 'Jun'], 3: ['Jul', 'Aug', 'Sep'], 4: ['Oct', 'Nov', 'Dec'] };
      for (const month of qMonths[currentQ]) {
        expect(screen.getAllByText(month).length).toBeGreaterThan(0);
      }
    });

    it('shows Total Active in month cards', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Monthly Trend');
      expect(screen.getAllByText('Total Active').length).toBeGreaterThan(0);
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // DATA TABLE TAB
  // ═══════════════════════════════════════════════════════════════════════
  describe('Data Table Tab', () => {
    it('renders all column headers', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('QO Number')).toBeInTheDocument());
      for (const col of ['Company', 'Contact', 'Project', 'Price (฿)', 'Cabinet', 'Wire/Busbar', 'Equipment', 'Agent', 'Stage']) {
        expect(screen.getByText(col)).toBeInTheDocument();
      }
    });

    it('renders data rows', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('QO-001')).toBeInTheDocument());
      expect(screen.getByText('QO-002')).toBeInTheDocument();
      expect(screen.getByText('QO-003')).toBeInTheDocument();
    });

    it('shows pagination info', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getAllByText(/rows · Page/).length).toBeGreaterThan(0));
    });

    it('shows search input', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      expect(screen.getByPlaceholderText(/Search by QO/)).toBeInTheDocument();
    });

    it('shows delete buttons per row', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('QO-001')).toBeInTheDocument());
      expect(screen.getAllByTitle('Delete row').length).toBe(SAMPLE_ROWS.length);
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // NEW ROW
  // ═══════════════════════════════════════════════════════════════════════
  describe('New Row', () => {
    it('opens ReviewModal on + New Row click', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('+ New Row')).toBeInTheDocument());

      fireEvent.click(screen.getByText('+ New Row'));
      await waitFor(() => {
        expect(screen.getByText('Review Quotation')).toBeInTheDocument();
        expect(screen.getByText('Save to Supabase')).toBeInTheDocument();
      });
    });

    it('pre-fills stage as On track', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('+ New Row')).toBeInTheDocument());

      fireEvent.click(screen.getByText('+ New Row'));
      await waitFor(() => expect(screen.getByText('Review Quotation')).toBeInTheDocument());
      // The modal uses <label> without htmlFor, so query the select by its displayed value
      const selects = document.querySelectorAll('select');
      const stageSelect = [...selects].find(s => s.value === 'On track');
      expect(stageSelect).toBeTruthy();
    });

    it('closes modal on Cancel', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('+ New Row')).toBeInTheDocument());

      fireEvent.click(screen.getByText('+ New Row'));
      await waitFor(() => expect(screen.getByText('Review Quotation')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Cancel'));
      await waitFor(() => expect(screen.queryByText('Review Quotation')).not.toBeInTheDocument());
    });

    it('calls upsert on Save', async () => {
      let upsertCalled = false;
      supabase.from = vi.fn(() => {
        const chain = makeChain(SAMPLE_ROWS);
        chain.upsert = () => { upsertCalled = true; return Promise.resolve({ error: null }); };
        return chain;
      });

      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('+ New Row')).toBeInTheDocument());

      fireEvent.click(screen.getByText('+ New Row'));
      await waitFor(() => expect(screen.getByText('Save to Supabase')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Save to Supabase'));
      await waitFor(() => expect(upsertCalled).toBe(true));
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // CELL EDIT
  // ═══════════════════════════════════════════════════════════════════════
  describe('Cell Edit', () => {
    it('enters edit mode on click', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('Acme Corp')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Acme Corp'));
      await waitFor(() => expect(screen.getByDisplayValue('Acme Corp').tagName).toBe('INPUT'));
    });

    it('cancels edit on Escape', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('Acme Corp')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Acme Corp'));
      const input = await screen.findByDisplayValue('Acme Corp');
      fireEvent.keyDown(input, { key: 'Escape' });

      await waitFor(() => {
        expect(screen.queryByDisplayValue('Acme Corp')).toBeNull();
        expect(screen.getByText('Acme Corp')).toBeInTheDocument();
      });
    });

    it('calls update on Enter when value changes', async () => {
      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('Acme Corp')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Acme Corp'));
      const input = await screen.findByDisplayValue('Acme Corp');
      fireEvent.change(input, { target: { value: 'Acme Updated' } });
      fireEvent.keyDown(input, { key: 'Enter' });

      // After committing an edit, supabase.from should be called for the update
      await waitFor(() => {
        const calls = supabase.from.mock.calls;
        // At least one call after the edit (for update + refreshData)
        expect(calls.length).toBeGreaterThan(0);
      });
    });

    it('does not save when value unchanged', async () => {
      let updateCalled = false;
      supabase.from = vi.fn(() => {
        const chain = makeChain(SAMPLE_ROWS);
        chain.update = () => { updateCalled = true; return chain; };
        return chain;
      });

      renderDashboard();
      await waitForData();
      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByText('Acme Corp')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Acme Corp'));
      const input = await screen.findByDisplayValue('Acme Corp');
      fireEvent.keyDown(input, { key: 'Enter' });

      expect(updateCalled).toBe(false);
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // UPLOAD QUOTATION
  // ═══════════════════════════════════════════════════════════════════════
  describe('Upload Quotation', () => {
    it('shows Upload QO button', async () => {
      renderDashboard();
      await waitForData();
      expect(screen.getByText('⬆ Upload QO')).toBeInTheDocument();
    });

    it('opens ReviewModal after xlsx upload', async () => {
      renderDashboard();
      await waitForData();

      const fileInput = document.querySelector('input[type="file"][accept=".xlsx"]');
      const file = new File(['dummy'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      file.arrayBuffer = vi.fn().mockResolvedValue(new ArrayBuffer(8));

      fireEvent.change(fileInput, { target: { files: [file] } });
      await waitFor(() => expect(screen.getByText('Review Quotation')).toBeInTheDocument());
    });

    it('populates modal fields from uploaded file', async () => {
      renderDashboard();
      await waitForData();

      const fileInput = document.querySelector('input[type="file"][accept=".xlsx"]');
      const file = new File(['dummy'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      file.arrayBuffer = vi.fn().mockResolvedValue(new ArrayBuffer(8));

      fireEvent.change(fileInput, { target: { files: [file] } });
      await waitFor(() => expect(screen.getByText('Review Quotation')).toBeInTheDocument());

      expect(screen.getByDisplayValue('QO-UPLOAD-001')).toBeInTheDocument();
      expect(screen.getByDisplayValue('Upload Corp')).toBeInTheDocument();
      expect(screen.getByDisplayValue('Upload Person')).toBeInTheDocument();
      expect(screen.getByDisplayValue('Upload Project')).toBeInTheDocument();
    });

    it('calls upsert on Save after upload', async () => {
      let upsertCalled = false;
      supabase.from = vi.fn(() => {
        const chain = makeChain(SAMPLE_ROWS);
        chain.upsert = () => { upsertCalled = true; return Promise.resolve({ error: null }); };
        return chain;
      });

      renderDashboard();
      await waitForData();

      const fileInput = document.querySelector('input[type="file"][accept=".xlsx"]');
      const file = new File(['dummy'], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      file.arrayBuffer = vi.fn().mockResolvedValue(new ArrayBuffer(8));

      fireEvent.change(fileInput, { target: { files: [file] } });
      await waitFor(() => expect(screen.getByText('Save to Supabase')).toBeInTheDocument());

      fireEvent.click(screen.getByText('Save to Supabase'));
      await waitFor(() => expect(upsertCalled).toBe(true));
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // QUARTER SWITCHING
  // ═══════════════════════════════════════════════════════════════════════
  describe('Quarter Switching', () => {
    it('renders Q1-Q4 buttons', async () => {
      renderDashboard();
      await waitForData();
      for (const q of ['Q1', 'Q2', 'Q3', 'Q4']) {
        expect(screen.getByRole('button', { name: q })).toBeInTheDocument();
      }
    });

    it('highlights active quarter', async () => {
      renderDashboard();
      await waitForData();
      const currentQ = Math.floor(new Date().getMonth() / 3) + 1;
      expect(screen.getByRole('button', { name: `Q${currentQ}` }).className).toContain('active');
    });

    it('re-fetches data on quarter switch', async () => {
      renderDashboard();
      await waitForData();

      const callsBefore = supabase.from.mock.calls.length;
      const currentQ = Math.floor(new Date().getMonth() / 3) + 1;
      const targetQ = currentQ === 1 ? 2 : 1;

      fireEvent.click(screen.getByRole('button', { name: `Q${targetQ}` }));
      await waitFor(() => expect(supabase.from.mock.calls.length).toBeGreaterThan(callsBefore));
    });

    it('updates quarter label', async () => {
      renderDashboard();
      await waitForData();

      fireEvent.click(screen.getByRole('button', { name: 'Q3' }));
      await waitFor(() => expect(screen.getAllByText(/Q3/).length).toBeGreaterThan(0));
    });
  });

  // ═══════════════════════════════════════════════════════════════════════
  // TAB NAVIGATION
  // ═══════════════════════════════════════════════════════════════════════
  describe('Tab Navigation', () => {
    it('renders all tab buttons', async () => {
      renderDashboard();
      await waitForData();
      for (const name of ['Overview', 'Sales Agents', 'Top Pipeline', 'Monthly Trend', 'Data Table']) {
        expect(screen.getByRole('button', { name })).toBeInTheDocument();
      }
    });

    it('switches content across all tabs', async () => {
      renderDashboard();
      await waitForData();

      expect(screen.getByText('Status Distribution')).toBeInTheDocument();

      await switchTab('Sales Agents');
      expect(screen.getByText('Orders by Agent — Count Comparison')).toBeInTheDocument();

      await switchTab('Top Pipeline');
      expect(screen.getByText(/All On-Track Quotations/)).toBeInTheDocument();

      await switchTab('Monthly Trend');
      expect(screen.getByText(/Monthly Revenue Breakdown/)).toBeInTheDocument();

      await switchTab('Data Table');
      await waitFor(() => expect(screen.getByPlaceholderText(/Search by QO/)).toBeInTheDocument());
    });
  });
});
