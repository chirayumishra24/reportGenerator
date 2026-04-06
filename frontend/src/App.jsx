import React, { useState, useRef, useMemo } from 'react';
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import {
  Upload, FileSpreadsheet, CheckCircle, Download, BarChart3, ArrowLeft, X,
  Loader2, Info, Plus, Trash2, UserCheck, FileDown, RotateCcw, Eye, Edit3, Save
} from 'lucide-react';
import './index.css';

const API = 'http://localhost:5000';

const RANGES = [
  { label: '95-100', min: 95, max: 100, color: '#22c55e' },
  { label: '90-94', min: 90, max: 94.999, color: '#3b82f6' },
  { label: '80-89', min: 80, max: 89.999, color: '#8b5cf6' },
  { label: '60-79', min: 60, max: 79.999, color: '#f59e0b' },
  { label: '50-59', min: 50, max: 59.999, color: '#f97316' },
  { label: 'below 50', min: 0, max: 49.999, color: '#ef4444' },
];

const CHART_COLORS = ['#6366f1', '#ef4444', '#22c55e', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899', '#14b8a6'];

function findTargetColumn(headers) {
  const priorities = ['% in IX+30', 'Grand Total', '% in IX'];
  for (const col of priorities) {
    const found = headers.find(h => h && h.toString().trim().toLowerCase() === col.toLowerCase());
    if (found) return found;
  }
  return null;
}

function getScoreColor(val) {
  const v = parseFloat(val);
  if (isNaN(v)) return '';
  if (v >= 95) return 'score-excellent';
  if (v >= 90) return 'score-great';
  if (v >= 80) return 'score-good';
  if (v >= 60) return 'score-average';
  if (v >= 50) return 'score-low';
  return 'score-poor';
}

// Compute analysis for a single sheet
function computeSheetAnalysis(sheet) {
  if (!sheet || !sheet.rows || sheet.rows.length === 0) return null;
  const targetCol = findTargetColumn(sheet.headers);
  if (!targetCol) return null;

  const rows = [];
  let totalStudents = 0;
  for (const range of RANGES) {
    let count = 0;
    for (const student of sheet.rows) {
      const v = parseFloat(student[targetCol]);
      if (!isNaN(v) && v >= range.min && v <= range.max) count++;
    }
    rows.push({ Range: range.label, Count: count, color: range.color });
    totalStudents += count;
  }

  // Add total row
  rows.push({ Range: 'Total', Count: totalStudents, color: '#1e293b' });
  return { rows, totalStudents, targetCol };
}

export default function App() {
  const [screen, setScreen] = useState('upload');
  const [fileName, setFileName] = useState('');
  const [sheetNames, setSheetNames] = useState([]);
  const [sheets, setSheets] = useState({});
  const [activeSheet, setActiveSheet] = useState('');
  const [loading, setLoading] = useState(false);
  const [showAnalysis, setShowAnalysis] = useState(false);
  const [analysisSheets, setAnalysisSheets] = useState([]);
  const [studentReport, setStudentReport] = useState(null);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [showCreateModal, setShowCreateModal] = useState(false);
  const fileRef = useRef(null);

  // ---------- UPLOAD ----------
  const handleFile = async (file) => {
    if (!file || !file.name.endsWith('.xlsx')) {
      alert('Please upload a valid .xlsx file');
      return;
    }
    setLoading(true);
    setFileName(file.name);
    const formData = new FormData();
    formData.append('file', file);

    try {
      const res = await fetch(`${API}/parse`, { method: 'POST', body: formData });
      const data = await res.json();
      if (data.sheetNames) {
        // Smart filter: remove summary/analysis rows from the imported sheets
        const cleanedSheets = {};
        for (const [sName, sData] of Object.entries(data.sheets)) {
          const validRows = [];
          for (const r of sData.rows) {
            const values = Object.values(r).map(v => String(v).trim().toLowerCase());
            // Filter out rows that are actually embedded analysis tables
            const isSummary = values.some(v => 
              v === '95-100' || v === '90-94' || v === '80-89' || 
              v === '60-79' || v === '50-59' || v === 'below 50'
            );
            // Also filter out completely empty rows
            const isEmpty = values.every(v => v === '' || v === 'null' || v === 'undefined');
            if (!isSummary && !isEmpty) {
              validRows.push(r);
            }
          }
          cleanedSheets[sName] = { headers: sData.headers, rows: validRows };
        }

        setSheetNames(data.sheetNames);
        setSheets(cleanedSheets);
        setActiveSheet(data.sheetNames[0]);
        setAnalysisSheets(data.sheetNames.slice());
        setScreen('dashboard');
      }
    } catch {
      alert('Cannot connect to backend. Is the Node server running on port 5000?');
    }
    setLoading(false);
  };

  // ---------- EDIT CELLS ----------
  const startEdit = (sheetName, rowIdx, col) => {
    setEditingCell({ sheetName, rowIdx, col });
    setEditValue(sheets[sheetName].rows[rowIdx][col] ?? '');
  };

  const saveEdit = () => {
    if (!editingCell) return;
    const { sheetName, rowIdx, col } = editingCell;
    const updated = { ...sheets };
    const val = editValue;
    // Try to parse as number
    const num = parseFloat(val);
    const finalVal = isNaN(num) || val === '' ? val : num;
    updated[sheetName].rows[rowIdx][col] = finalVal;

    // Auto-calculate Grand Total if needed
    const row = updated[sheetName].rows[rowIdx];
    const headers = updated[sheetName].headers;
    const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
    
    if (totalCol && col !== totalCol) {
      let sum = 0;
      // Sum all numeric columns that are NOT S.No, Name, Grand Total, or % columns
      headers.forEach(h => {
        const lowerH = h.toLowerCase();
        if (h !== totalCol && !lowerH.includes('s.no') && !lowerH.includes('sr.') && 
            !lowerH.includes('name') && !lowerH.includes('%') &&
            !lowerH.includes('admn') && !lowerH.includes('admin') &&
            !lowerH.includes('roll') && !lowerH.includes('rank')) {
          const sv = parseFloat(row[h]);
          if (!isNaN(sv)) sum += sv;
        }
      });
      row[totalCol] = sum;
    }

    setSheets(updated);
    setEditingCell(null);
  };

  const cancelEdit = () => setEditingCell(null);

  // ---------- CREATE NEW WORKSPACE ----------
  const handleCreateTemplate = (config) => {
    const { className, numSections, subjects } = config;
    const newSheetNames = [];
    const newSheets = {};
    const sectionLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    const headers = ['S.No', 'Name', ...subjects, 'Grand Total', '% in IX+30'];

    for (let i = 0; i < numSections; i++) {
       const sec = sectionLetters[i] || `${i+1}`;
       const sheetName = `${className} ${sec}`;
       newSheetNames.push(sheetName);
       newSheets[sheetName] = { headers, rows: [] };
    }

    setFileName(`NEW_${className}_DATA.xlsx`);
    setSheetNames(newSheetNames);
    setSheets(newSheets);
    setActiveSheet(newSheetNames[0]);
    setAnalysisSheets(newSheetNames.slice());
    setShowCreateModal(false);
    setScreen('dashboard');
  };

  // ---------- ADD / DELETE STUDENT ----------
  const addStudent = () => {
    if (!activeSheet || !sheets[activeSheet]) return;
    const updated = { ...sheets };
    const headers = updated[activeSheet].headers;
    const newRow = {};
    headers.forEach(h => { newRow[h] = ''; });
    // Auto-fill S.No
    const snCol = headers.find(h => h.toLowerCase().includes('s.no') || h.toLowerCase() === 'sr. no.');
    if (snCol) newRow[snCol] = updated[activeSheet].rows.length + 1;
    updated[activeSheet].rows.push(newRow);
    setSheets(updated);
  };

  const deleteStudent = (rowIdx) => {
    if (!confirm('Delete this student?')) return;
    const updated = { ...sheets };
    updated[activeSheet].rows.splice(rowIdx, 1);
    setSheets(updated);
  };

  // ---------- ADD NEW SHEET ----------
  const addSheet = () => {
    const name = prompt('Enter new sheet name:');
    if (!name || sheetNames.includes(name)) return;
    const headers = ['S.No', 'Name', 'Score'];
    setSheetNames([...sheetNames, name]);
    setSheets({
      ...sheets,
      [name]: { headers, rows: [] }
    });
    setActiveSheet(name);
  };

  // ---------- ANALYSIS ----------
  const analysisData = useMemo(() => {
    if (!showAnalysis) return null;
    const selected = analysisSheets.filter(s => sheets[s]);
    const sums = {};
    selected.forEach(s => sums[s] = 0);
    const rows = [];

    for (const range of RANGES) {
      const row = { Range: range.label };
      let total = 0;
      for (const sn of selected) {
        const sheet = sheets[sn];
        const targetCol = findTargetColumn(sheet.headers);
        if (!targetCol) { row[sn] = 0; continue; }
        let count = 0;
        for (const student of sheet.rows) {
          const v = parseFloat(student[targetCol]);
          if (!isNaN(v) && v >= range.min && v <= range.max) count++;
        }
        row[sn] = count;
        total += count;
        sums[sn] += count;
      }
      row.students = total;
      rows.push(row);
    }

    const grandTotal = Object.values(sums).reduce((a, b) => a + b, 0);
    rows.forEach(r => {
      r['per%'] = grandTotal > 0 ? parseFloat(((r.students / grandTotal) * 100).toFixed(2)) : 0;
    });

    const totalRow = { Range: 'Total' };
    selected.forEach(s => totalRow[s] = sums[s]);
    totalRow.students = grandTotal;
    totalRow['per%'] = 100;
    rows.push(totalRow);

    return { sections: selected, rows, headers: ['Range', ...selected, 'students', 'per%'] };
  }, [showAnalysis, analysisSheets, sheets]);

  // ---------- STUDENT REPORT ----------
  const openStudentReport = (row) => {
    setStudentReport(row);
  };

  // ---------- EXPORT ----------
  const exportExcel = async () => {
    setLoading(true);
    try {
      // Build analysis sheet data
      let analysisSheet = null;
      if (analysisData) {
        analysisSheet = {
          headers: analysisData.headers,
          rows: analysisData.rows,
        };
      }

      const res = await fetch(`${API}/export`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetNames, sheets, analysisSheet })
      });

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Report_${fileName}`;
      a.click();
    } catch {
      alert('Export failed.');
    }
    setLoading(false);
  };

  // ---------- RESET ----------
  const resetAll = () => {
    if (!confirm('Reset everything? All changes will be lost.')) return;
    setScreen('upload');
    setSheets({});
    setSheetNames([]);
    setActiveSheet('');
    setShowAnalysis(false);
    setStudentReport(null);
  };

  const currentSheet = sheets[activeSheet];
  const totalStudents = currentSheet ? currentSheet.rows.length : 0;

  // ========== RENDER ==========
  return (
    <div className="app-wrapper">
      <div className="app-bg" />

      {/* ===== UPLOAD SCREEN ===== */}
      {screen === 'upload' && (
        <div className="main-card fade-in upload-card">
          <div className="card-icon"><FileSpreadsheet size={48} strokeWidth={1.5} /></div>
          <h1>Target Analysis Report Generator</h1>
          <p className="subtitle">Upload your Excel file with section-wise student data to generate comprehensive analysis reports with charts</p>
          <div className="landing-actions">
            <div className="drop-zone" onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }} onClick={() => fileRef.current?.click()}>
              <input ref={fileRef} type="file" accept=".xlsx" style={{ display: 'none' }} onChange={e => handleFile(e.target.files[0])} />
              {loading ? (
                <div className="drop-content"><Loader2 size={40} className="spin" /><p>Parsing file...</p></div>
              ) : (
                <div className="drop-content"><Upload size={40} /><p><strong>Click to upload</strong> or drag &amp; drop</p><span className="drop-hint">.xlsx files only</span></div>
              )}
            </div>
            
            <div className="divider"><span>OR</span></div>
            
            <button className="btn-primary create-btn" onClick={() => setShowCreateModal(true)}>
              <Plus size={20} /> Create New Template
            </button>
          </div>

          <div className="info-bar"><Info size={16} /><span>Your data stays local — processed on your machine only.</span></div>
        </div>
      )}

      {/* ===== CREATE MODAL ===== */}
      {showCreateModal && (
        <CreateTemplateModal 
          onClose={() => setShowCreateModal(false)}
          onCreate={handleCreateTemplate}
        />
      )}

      {/* ===== DASHBOARD ===== */}
      {screen === 'dashboard' && (
        <div className="main-card fade-in dashboard-card">
          {/* Header */}
          <div className="dash-header">
            <div className="dash-header-left">
              <FileSpreadsheet size={28} className="header-icon" />
              <div>
                <h2>Student Data Manager</h2>
                <p className="header-sub">View, edit, and analyze student performance data</p>
              </div>
            </div>
          </div>

          {/* Info Bar */}
          <div className="file-info-bar">
            <div className="info-chip"><strong>File:</strong> {fileName}</div>
            <div className="info-chip"><strong>Sheets:</strong> {sheetNames.length}</div>
            <div className="info-chip"><strong>Active:</strong> {activeSheet}</div>
            <div className="info-chip highlight"><strong>Students:</strong> {totalStudents}</div>
          </div>

          {/* Toolbar */}
          <div className="toolbar">
            <button className="tool-btn primary" onClick={addStudent}><Plus size={16} /> Add Student</button>
            <button className="tool-btn success" onClick={() => { setShowAnalysis(true); }}><BarChart3 size={16} /> Class Analysis</button>
            <button className="tool-btn outline" onClick={exportExcel} disabled={loading}>
              {loading ? <Loader2 size={16} className="spin" /> : <FileDown size={16} />} Export Excel
            </button>
            <button className="tool-btn outline" onClick={addSheet}><Plus size={16} /> New Sheet</button>
            <button className="tool-btn danger-outline" onClick={resetAll}><RotateCcw size={16} /> Reset All</button>
          </div>

          {/* Sheet Tabs */}
          <div className="sheet-tabs">
            {sheetNames.map(name => (
              <button key={name} className={`sheet-tab ${activeSheet === name ? 'active' : ''}`} onClick={() => { setActiveSheet(name); setShowAnalysis(false); }}>
                {name}
              </button>
            ))}
          </div>

          {/* Data Table or Analysis */}
          {showAnalysis ? (
            <AnalysisPanel
              data={analysisData}
              sheets={sheetNames}
              selected={analysisSheets}
              setSelected={setAnalysisSheets}
              onClose={() => setShowAnalysis(false)}
            />
          ) : currentSheet ? (
            <>
              <div className="table-container">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th className="row-num">#</th>
                      {currentSheet.headers.map(h => <th key={h}>{h}</th>)}
                      <th className="actions-col">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentSheet.rows.length === 0 ? (
                      <tr><td colSpan={currentSheet.headers.length + 2} className="empty-msg">No students added yet. Click "Add Student" to begin.</td></tr>
                    ) : currentSheet.rows.map((row, ri) => (
                      <tr key={ri}>
                        <td className="row-num">{ri + 1}</td>
                        {currentSheet.headers.map(h => {
                          const isEditing = editingCell && editingCell.sheetName === activeSheet && editingCell.rowIdx === ri && editingCell.col === h;
                          const val = row[h];
                          const scoreClass = (h.toLowerCase().includes('%') || h.toLowerCase().includes('total')) ? getScoreColor(val) : '';
                          return (
                            <td key={h} className={`data-cell ${scoreClass}`} onDoubleClick={() => startEdit(activeSheet, ri, h)}>
                              {isEditing ? (
                                <input
                                  className="cell-input"
                                  value={editValue}
                                  onChange={e => setEditValue(e.target.value)}
                                  onBlur={saveEdit}
                                  onKeyDown={e => { if (e.key === 'Enter') saveEdit(); if (e.key === 'Escape') cancelEdit(); }}
                                  autoFocus
                                />
                              ) : (
                                <span>{val !== '' && val !== null && val !== undefined ? String(val) : '—'}</span>
                              )}
                            </td>
                          );
                        })}
                        <td className="actions-col">
                          <button className="icon-btn view" title="View Report" onClick={() => openStudentReport(row)}>
                            <Eye size={15} />
                          </button>
                          <button className="icon-btn delete" title="Delete" onClick={() => deleteStudent(ri)}>
                            <Trash2 size={15} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Per-Sheet Analysis */}
              <SheetAnalysis sheetName={activeSheet} sheet={currentSheet} />
            </>
          ) : null}

          <div className="footer-info">
            Target Analysis Report Generator — Double-click any cell to edit • Data auto-saved in session
          </div>
        </div>
      )}

      {/* ===== STUDENT REPORT MODAL ===== */}
      {studentReport && (
        <StudentReportModal student={studentReport} headers={currentSheet?.headers || []} onClose={() => setStudentReport(null)} />
      )}
    </div>
  );
}

// ========== ANALYSIS PANEL ==========
function AnalysisPanel({ data, sheets, selected, setSelected, onClose }) {
  const toggleSheet = (name) => {
    setSelected(prev => prev.includes(name) ? prev.filter(s => s !== name) : [...prev, name]);
  };

  if (!data) return null;

  return (
    <div className="analysis-panel fade-in">
      <div className="analysis-header">
        <h3><BarChart3 size={20} /> Section-wise Target Analysis</h3>
        <button className="icon-btn" onClick={onClose}><X size={18} /></button>
      </div>

      {/* Sheet selector */}
      <div className="analysis-sheets">
        <span className="label">Include sheets:</span>
        {sheets.map(s => (
          <button key={s} className={`sheet-pill ${selected.includes(s) ? 'active' : ''}`} onClick={() => toggleSheet(s)}>
            {selected.includes(s) && <CheckCircle size={14} />} {s}
          </button>
        ))}
      </div>

      {/* Charts Box */}
      <div className="chart-box" style={{ padding: '2rem 1.5rem', background: '#fff' }}>
        <h4 style={{ marginBottom: '1.5rem', color: 'var(--text)', borderBottom: '1px solid var(--border)', paddingBottom: '0.75rem' }}>
          {data.sections.map(s => s.replace(/^X\s*/, 'X-')).join(', ')} — Performance Visuals
        </h4>
        
        <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)', gap: '2rem', height: '350px' }}>
          {/* Bar Chart */}
          <div>
            <h5 style={{ textAlign: 'center', marginBottom: '1rem', color: 'var(--text-secondary)' }}>Bar Analysis</h5>
            <ResponsiveContainer width="100%" height="85%">
              <BarChart data={data.rows.slice(0, -1)} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip contentStyle={{ borderRadius: 10, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)' }} />
                <Legend />
                {data.sections.map((s, i) => (
                  <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[4, 4, 0, 0]} />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </div>

          {/* Line Chart */}
          <div>
            <h5 style={{ textAlign: 'center', marginBottom: '1rem', color: 'var(--text-secondary)' }}>Trend Line Analysis</h5>
            <ResponsiveContainer width="100%" height="85%">
              <LineChart data={data.rows.slice(0, -1)} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} />
                <Tooltip contentStyle={{ borderRadius: 10, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)' }} />
                <Legend />
                {data.sections.map((s, i) => (
                  <Line key={s} type="monotone" dataKey={s} stroke={CHART_COLORS[i % CHART_COLORS.length]} strokeWidth={3} activeDot={{ r: 6 }} />
                ))}
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Table */}
      <div className="table-container">
        <table className="data-table analysis-tbl">
          <thead>
            <tr>
              {data.headers.map(h => <th key={h}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {data.rows.map((row, i) => (
              <tr key={i} className={i === data.rows.length - 1 ? 'total-row' : ''}>
                {data.headers.map(h => (
                  <td key={h} className={h === 'Range' ? `range-cell range-${row[h]?.replace(/[^a-zA-Z0-9]/g, '')}` : ''}>
                    {h === 'per%' ? `${row[h]}%` : row[h]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ========== PER-SHEET ANALYSIS ==========
function SheetAnalysis({ sheetName, sheet }) {
  const analysis = useMemo(() => computeSheetAnalysis(sheet), [sheet]);
  if (!analysis) return null;

  const chartData = analysis.rows.slice(0, -1); // exclude total row

  return (
    <div className="sheet-analysis fade-in">
      <h3 className="sheet-analysis-title">
        <BarChart3 size={18} /> {sheetName} — Score Distribution
      </h3>
      <div className="sheet-analysis-body">
        {/* Table */}
        <div className="sheet-analysis-table">
          <table className="data-table analysis-tbl mini-tbl">
            <thead>
              <tr>
                <th>Range</th>
                <th>Students</th>
              </tr>
            </thead>
            <tbody>
              {analysis.rows.map((row, i) => (
                <tr key={i} className={i === analysis.rows.length - 1 ? 'total-row' : ''}>
                  <td className={`range-cell range-${row.Range?.replace(/[^a-zA-Z0-9]/g, '')}`}>
                    {row.Range}
                  </td>
                  <td style={{ textAlign: 'center', fontWeight: 600 }}>{row.Count}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Chart */}
        <div className="sheet-analysis-chart">
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={chartData} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
              <XAxis dataKey="Range" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} />
              <Tooltip />
              <Legend />
              <Bar dataKey="Count" name={sheetName} radius={[4, 4, 0, 0]}>
                {chartData.map((entry, i) => (
                  <Cell key={i} fill={entry.color} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

// ========== STUDENT REPORT MODAL ==========
function StudentReportModal({ student, headers, onClose }) {
  // Find numeric columns for the score breakdown
  const scoreFields = headers.filter(h => {
    const val = student[h];
    return typeof val === 'number' && !h.toLowerCase().includes('s.no') && !h.toLowerCase().includes('sr.');
  });

  const nameCol = headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
  const name = nameCol ? student[nameCol] : 'Student';

  const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase().includes('total'));
  const total = totalCol ? student[totalCol] : null;

  const percentCol = headers.find(h => h.toLowerCase().includes('% in ix+30') || h.toLowerCase().includes('% in ix'));
  const percent = percentCol ? student[percentCol] : null;

  // Pie chart data for subject scores
  const subjectCols = headers.filter(h => {
    const lowerH = h.toLowerCase();
    return typeof student[h] === 'number' && 
      !lowerH.includes('s.no') && !lowerH.includes('sr.') && 
      !lowerH.includes('name') && !lowerH.includes('%') &&
      !lowerH.includes('admn') && !lowerH.includes('admin') &&
      !lowerH.includes('roll') && !lowerH.includes('rank') &&
      !lowerH.includes('dob') && !lowerH.includes('date') &&
      !lowerH.includes('grand total') && !lowerH.includes('total');
  });

  const pieData = subjectCols.map((h, i) => ({
    name: h.replace(/\s*\+\s*30/g, '').trim(),
    value: student[h] || 0,
    color: CHART_COLORS[i % CHART_COLORS.length]
  }));

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card fade-in print-region" onClick={e => e.stopPropagation()}>
        <div className="modal-header print-hide">
          <h3><UserCheck size={20} /> Student Report</h3>
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <button className="tool-btn outline" onClick={() => window.print()}>
              <FileDown size={14} /> Download PDF
            </button>
            <button className="icon-btn" onClick={onClose}><X size={18} /></button>
          </div>
        </div>

        <div className="report-name-bar">
          <div className="student-avatar">{name?.charAt(0)}</div>
          <div>
            <h2>{name}</h2>
            {percent !== null && (
              <span className={`score-badge ${getScoreColor(percent)}`}>
                {percent}% ({percent >= 90 ? 'Excellent' : percent >= 80 ? 'Very Good' : percent >= 60 ? 'Good' : percent >= 50 ? 'Average' : 'Needs Improvement'})
              </span>
            )}
          </div>
        </div>

        {/* Details Grid */}
        <div className="report-grid">
          {headers.map(h => (
            <div key={h} className="report-field">
              <label>{h}</label>
              <span className={typeof student[h] === 'number' ? getScoreColor(student[h]) : ''}>
                {student[h] !== '' && student[h] !== null && student[h] !== undefined ? String(student[h]) : '—'}
              </span>
            </div>
          ))}
        </div>

        {/* Subject Chart */}
        {pieData.length > 0 && (
          <div className="chart-box report-chart">
            <h4>Subject-wise Score Breakdown</h4>
            <ResponsiveContainer width="100%" height={250}>
              <BarChart data={pieData} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-30} textAnchor="end" height={60} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip />
                <Bar dataKey="value" radius={[4, 4, 0, 0]}>
                  {pieData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
      {/* CSS injected directly for print layout of this modal */}
      <style dangerouslySetInnerHTML={{__html: `
        @media print {
          body * { visibility: hidden; }
          .print-region, .print-region * { visibility: visible; }
          .print-region {
            position: absolute !important;
            left: 0 !important; top: 0 !important;
            width: 100% !important; margin: 0 !important;
            box-shadow: none !important; border: none !important;
          }
          .print-hide { display: none !important; }
          .modal-overlay { background: transparent !important; }
          /* Fix chart rendering on print */
          .recharts-responsive-container { width: 100% !important; height: 350px !important; }
        }
      `}} />
    </div>
  );
}

// ========== CREATE TEMPLATE MODAL ==========
function CreateTemplateModal({ onClose, onCreate }) {
  const [className, setClassName] = useState('Class X');
  const [numSections, setNumSections] = useState(5);
  const [subjectInput, setSubjectInput] = useState('');
  const [subjects, setSubjects] = useState(['English', 'Hindi', 'Maths', 'Science', 'Social Science', 'IT / AI']);

  const addSubject = (e) => {
    e.preventDefault();
    if (subjectInput.trim() && !subjects.includes(subjectInput.trim())) {
      setSubjects([...subjects, subjectInput.trim()]);
      setSubjectInput('');
    }
  };

  const removeSubject = (subj) => {
    setSubjects(subjects.filter(s => s !== subj));
  };

  const handleCreate = () => {
    if (!className.trim()) { alert('Class name is required'); return; }
    if (numSections < 1) { alert('At least 1 section is required'); return; }
    if (subjects.length === 0) { alert('Add at least 1 subject'); return; }
    
    // Use exact expected target sheet headers
    const newHeaders = ['S.No', 'Admn. No.', 'Name', ...subjects, 'Grand Total', '% in IX'];
    
    onCreate({ className: className.trim(), numSections, subjects: newHeaders.slice(3, -2) });
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card fade-in" style={{ maxWidth: '500px' }} onClick={e => e.stopPropagation()}>
        <div className="modal-header">
          <h3><Plus size={20} /> Create Blank Class Template</h3>
          <button className="icon-btn" onClick={onClose}><X size={18} /></button>
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Class Level / Name</label>
          <input className="cell-input" value={className} onChange={e => setClassName(e.target.value)} placeholder="e.g. Class X" style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Number of Sections (A, B, C...)</label>
          <input type="number" min="1" max="20" className="cell-input" value={numSections} onChange={e => setNumSections(parseInt(e.target.value) || 1)} style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
        </div>

        <div style={{ marginBottom: '1.5rem' }}>
          <label style={{ display: 'block', fontSize: '0.82rem', fontWeight: 600, color: 'var(--text-secondary)', marginBottom: '0.4rem' }}>Subjects</label>
          <form onSubmit={addSubject} style={{ display: 'flex', gap: '0.5rem', marginBottom: '0.75rem' }}>
            <input className="cell-input" value={subjectInput} onChange={e => setSubjectInput(e.target.value)} placeholder="Type new subject" style={{ padding: '0.6rem', fontSize: '0.9rem' }} />
            <button type="submit" className="tool-btn outline" style={{ minWidth: "60px" }}>Add</button>
          </form>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '0.4rem' }}>
            {subjects.map(s => (
              <span key={s} style={{ display: 'inline-flex', alignItems: 'center', gap: '0.3rem', padding: '0.3rem 0.6rem', background: '#f1f5f9', borderRadius: '6px', fontSize: '0.8rem', fontWeight: 500, border: '1px solid #e2e8f0' }}>
                {s} <X size={14} style={{ cursor: 'pointer', color: '#64748b' }} onClick={() => removeSubject(s)} />
              </span>
            ))}
            {subjects.length === 0 && <span style={{fontSize: '0.8rem', color: '#94a3b8'}}>No subjects added</span>}
          </div>
        </div>

        <div style={{ marginTop: '2rem', display: 'flex', gap: '0.75rem', justifyContent: 'flex-end', paddingTop: '1rem', borderTop: '1px solid var(--border)' }}>
          <button className="tool-btn outline" onClick={onClose}>Cancel</button>
          <button className="tool-btn primary" onClick={handleCreate}>Create Template</button>
        </div>
      </div>
    </div>
  );
}
