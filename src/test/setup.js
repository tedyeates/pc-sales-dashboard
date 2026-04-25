import '@testing-library/jest-dom';

// Recharts uses ResizeObserver internally
class ResizeObserver {
  observe() {}
  unobserve() {}
  disconnect() {}
}
window.ResizeObserver = ResizeObserver;
