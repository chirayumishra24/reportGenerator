export const API = import.meta.env.PROD
  ? '/api'
  : 'http://localhost:5000/api';

export const SESSION_STORAGE_KEY = 'report-generator-session';
export const SESSION_VERSION = 5;
export const CUMULATIVE_SHEET_NAME = 'Cumulative Class 10';
export const CLASS10_SECTIONS = ['A', 'B', 'C', 'D', 'E'];
export const PHASE_ONE_CUMULATIVE_ONLY = false;

export const RANGES = [
  { label: '95-100', min: 95, max: 100, color: '#22c55e' },
  { label: '90-94', min: 90, max: 94.999, color: '#3b82f6' },
  { label: '80-89', min: 80, max: 89.999, color: '#8b5cf6' },
  { label: '60-79', min: 60, max: 79.999, color: '#f59e0b' },
  { label: '50-59', min: 50, max: 59.999, color: '#f97316' },
  { label: 'below 50', min: 0, max: 49.999, color: '#ef4444' },
];

export const CHART_COLORS = ['#6366f1', '#ef4444', '#22c55e', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6'];
