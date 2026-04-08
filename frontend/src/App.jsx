import React, { useState, useRef, useMemo, useCallback } from 'react';
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis
} from 'recharts';
import {
  Upload, FileSpreadsheet, CheckCircle, Download, BarChart3, ArrowLeft, X,
  Loader2, Info, Plus, Trash2, UserCheck, FileDown, RotateCcw, Eye, Edit3, Save,
  Printer, Search, TrendingUp, TrendingDown, Award, Users
} from 'lucide-react';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import './index.css';

const API = import.meta.env.PROD 
  ? '/api' 
  : 'http://localhost:5000/api';

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

// Check if a subject value means "not opted" (dash, empty, etc.)
function isNotOpted(val) {
  if (val === null || val === undefined || val === '') return true;
  const s = String(val).trim();
  return s === '-' || s === '—' || s === '–' || s === 'N/A' || s === 'NA' || s === '';
}

// Determine which columns are "subject" columns (not metadata, not totals, not %)
function getSubjectColumns(headers) {
  return headers.filter(h => {
    const lowerH = h.toLowerCase();
    return !lowerH.includes('s.no') && !lowerH.includes('sr.') &&
      !lowerH.includes('name') && !lowerH.includes('%') &&
      !lowerH.includes('admn') && !lowerH.includes('admin') &&
      !lowerH.includes('roll') && !lowerH.includes('rank') &&
      !lowerH.includes('dob') && !lowerH.includes('date') &&
      !lowerH.includes('father') && !lowerH.includes('mother') &&
      !lowerH.includes('gender') && !lowerH.includes('enrollment') &&
      !lowerH.includes('grand total') && !lowerH.includes('total') &&
      !lowerH.includes('column') &&
      !lowerH.includes('+30') && !lowerH.includes('+ 30') && // Exclude +30 target columns
      !lowerH.includes('ix 100') && !lowerH.includes('x target'); // Exclude derived columns
  });
}

// Get max marks for a subject from the header (e.g. "English  80" -> 80)
function getMaxMarksFromHeader(header) {
  const match = header.match(/(\d+)\s*$/);
  return match ? parseInt(match[1]) : null;
}

// Recalculate Grand Total and % excluding not-opted subjects (dash/empty)
function recalcGrandTotal(row, headers) {
  const subjectCols = getSubjectColumns(headers);
  const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
  if (!totalCol) return row;

  let sum = 0;
  let maxMarks = 0;
  let hasAny = false;
  subjectCols.forEach(h => {
    const val = row[h];
    if (!isNotOpted(val)) {
      const sv = parseFloat(val);
      if (!isNaN(sv)) {
        sum += sv;
        hasAny = true;
        const headerMax = getMaxMarksFromHeader(h);
        maxMarks += headerMax || 100;
      }
    }
  });

  row[totalCol] = hasAny ? sum : '';

  // Also recalculate % in IX if it exists
  const pctCol = headers.find(h => h.toLowerCase().includes('% in ix') && !h.toLowerCase().includes('+30'));
  if (pctCol && hasAny && maxMarks > 0) {
    row[pctCol] = parseFloat(((sum / maxMarks) * 100).toFixed(2));
  }

  return row;
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
  const [searchQuery, setSearchQuery] = useState('');
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
        const cleanedSheets = {};
        for (const [sName, sData] of Object.entries(data.sheets)) {
          const validRows = [];
          for (const r of sData.rows) {
            const values = Object.values(r).map(v => String(v).trim().toLowerCase());
            const isSummary = values.some(v => 
              v === '95-100' || v === '90-94' || v === '80-89' || 
              v === '60-79' || v === '50-59' || v === 'below 50'
            );
            const isEmpty = values.every(v => v === '' || v === 'null' || v === 'undefined');
            if (!isSummary && !isEmpty) {
              // Recalculate Grand Total excluding not-opted subjects
              recalcGrandTotal(r, sData.headers);
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
    } catch (err) {
      console.error('Upload error:', err);
      alert('Cannot connect to backend. Please check the server is running. Error: ' + err.message);
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
    const num = parseFloat(val);
    const finalVal = isNaN(num) || val === '' ? val : num;
    updated[sheetName].rows[rowIdx][col] = finalVal;

    // Recalculate Grand Total excluding not-opted subjects
    const row = updated[sheetName].rows[rowIdx];
    const headers = updated[sheetName].headers;
    recalcGrandTotal(row, headers);

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
  }, [analysisSheets, sheets]);

  // Subject-wise section comparison data
  const subjectComparison = useMemo(() => {
    const selected = analysisSheets.filter(s => sheets[s]);
    if (selected.length === 0) return null;

    const firstSheet = sheets[selected[0]];
    if (!firstSheet) return null;
    const subjectCols = getSubjectColumns(firstSheet.headers);
    if (subjectCols.length === 0) return null;

    const data = subjectCols.map(subject => {
      const entry = { subject: subject.replace(/\s*\+\s*30/g, '').replace(/\s+\d+$/, '').trim() };
      for (const sn of selected) {
        const sheet = sheets[sn];
        let sum = 0, count = 0;
        for (const student of sheet.rows) {
          const val = student[subject];
          // Skip not-opted subjects
          if (isNotOpted(val)) continue;
          const v = parseFloat(val);
          if (!isNaN(v)) { sum += v; count++; }
        }
        entry[sn] = count > 0 ? parseFloat((sum / count).toFixed(1)) : 0;
      }
      return entry;
    });

    return { data, sections: selected, subjects: subjectCols };
  }, [analysisSheets, sheets]);

  // ---------- STUDENT REPORT ----------
  const openStudentReport = (row) => {
    setStudentReport(row);
  };

  // ---------- EXPORT EXCEL ----------
  // Charts are generated automatically on the server — no need to capture from UI
  const exportExcel = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API}/export`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sheetNames, sheets })
      });

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        throw new Error(errData.error || 'Export failed');
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Report_${fileName}`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Export error:', err);
      alert('Export failed: ' + err.message);
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

  // ---------- SEARCH & FILTER ----------
  const filteredRows = useMemo(() => {
    if (!currentSheet) return [];
    if (!searchQuery.trim()) return currentSheet.rows;
    const q = searchQuery.toLowerCase();
    const nameCol = currentSheet.headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
    return currentSheet.rows.filter(row => {
      if (nameCol && String(row[nameCol] || '').toLowerCase().includes(q)) return true;
      return Object.values(row).some(v => String(v).toLowerCase().includes(q));
    });
  }, [currentSheet, searchQuery]);

  // ---------- CLASS STATS ----------
  const classStats = useMemo(() => {
    if (!currentSheet || currentSheet.rows.length === 0) return null;
    const totalCol = currentSheet.headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
    const percentCol = findTargetColumn(currentSheet.headers); // % in IX+30, Grand Total, or % in IX
    const nameCol = currentSheet.headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
    
    // Use percentage column for average; fall back to totalCol
    const avgCol = percentCol || totalCol;
    if (!avgCol) return null;

    // For ranking/top/bottom, prefer totalCol (raw marks), fall back to avgCol
    const rankCol = totalCol || avgCol;

    const scored = currentSheet.rows.filter(r => !isNaN(parseFloat(r[rankCol])) && parseFloat(r[rankCol]) > 0);
    if (scored.length === 0) return null;

    // Average uses percentage column
    const avgScores = scored.filter(r => !isNaN(parseFloat(r[avgCol]))).map(r => parseFloat(r[avgCol]));
    const avg = avgScores.length > 0 ? avgScores.reduce((a, b) => a + b, 0) / avgScores.length : 0;

    // Top/bottom uses rank column (Grand Total)
    const rankScores = scored.map(r => parseFloat(r[rankCol]));
    const maxScore = Math.max(...rankScores);
    const minScore = Math.min(...rankScores);
    const topStudent = scored.find(r => parseFloat(r[rankCol]) === maxScore);
    const bottomStudent = scored.find(r => parseFloat(r[rankCol]) === minScore);

    // Rank students by total (descending)
    const ranked = [...scored].sort((a, b) => parseFloat(b[rankCol]) - parseFloat(a[rankCol]));
    const rankMap = new Map();
    ranked.forEach((r, i) => rankMap.set(r, i + 1));

    return {
      avg: avg.toFixed(1),
      topName: nameCol ? (topStudent?.[nameCol] || '—') : '—',
      topScore: maxScore,
      bottomName: nameCol ? (bottomStudent?.[nameCol] || '—') : '—',
      bottomScore: minScore,
      totalScored: scored.length,
      rankMap,
    };
  }, [currentSheet]);

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
          <div className="info-bar"><Info size={16} /><span>Your data stays local — processed securely and not stored.</span></div>
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
          <div className="dash-header">
            <div className="dash-header-left">
              <FileSpreadsheet size={28} className="header-icon" />
              <div>
                <h2>Student Data Manager</h2>
                <p className="header-sub">View, edit, and analyze student performance data</p>
              </div>
            </div>
          </div>

          <div className="file-info-bar">
            <div className="info-chip"><strong>File:</strong> {fileName}</div>
            <div className="info-chip"><strong>Sheets:</strong> {sheetNames.length}</div>
            <div className="info-chip"><strong>Active:</strong> {activeSheet}</div>
            <div className="info-chip highlight"><strong>Students:</strong> {totalStudents}</div>
          </div>

          <div className="toolbar">
            <button className="tool-btn primary" onClick={addStudent}><Plus size={16} /> Add Student</button>
            <button className="tool-btn success" onClick={() => { setShowAnalysis(true); }}><BarChart3 size={16} /> Class Analysis</button>
            <button className="tool-btn outline" onClick={exportExcel} disabled={loading}>
              {loading ? <Loader2 size={16} className="spin" /> : <FileDown size={16} />} Export Excel
            </button>
            <button className="tool-btn outline" onClick={addSheet}><Plus size={16} /> New Sheet</button>
            <button className="tool-btn danger-outline" onClick={resetAll}><RotateCcw size={16} /> Reset All</button>
            <div className="search-box">
              <Search size={15} className="search-icon" />
              <input className="search-input" placeholder="Search students..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} />
              {searchQuery && <button className="search-clear" onClick={() => setSearchQuery('')}><X size={14} /></button>}
            </div>
          </div>

          {/* Quick Stats Bar */}
          {classStats && (
            <div className="stats-bar fade-in">
              <div className="stat-card">
                <div className="stat-icon"><Users size={18} /></div>
                <div><span className="stat-label">Class Average</span><span className="stat-value">{classStats.avg}</span></div>
              </div>
              <div className="stat-card stat-top">
                <div className="stat-icon"><TrendingUp size={18} /></div>
                <div><span className="stat-label">Top Scorer</span><span className="stat-value">{classStats.topName} <small>({classStats.topScore})</small></span></div>
              </div>
              <div className="stat-card stat-bottom">
                <div className="stat-icon"><TrendingDown size={18} /></div>
                <div><span className="stat-label">Needs Attention</span><span className="stat-value">{classStats.bottomName} <small>({classStats.bottomScore})</small></span></div>
              </div>
              <div className="stat-card">
                <div className="stat-icon"><Award size={18} /></div>
                <div><span className="stat-label">Scored Students</span><span className="stat-value">{classStats.totalScored} / {totalStudents}</span></div>
              </div>
            </div>
          )}

          <div className="sheet-tabs">
            {sheetNames.map(name => (
              <button key={name} className={`sheet-tab ${activeSheet === name ? 'active' : ''}`} onClick={() => { setActiveSheet(name); setShowAnalysis(false); }}>
                {name}
              </button>
            ))}
          </div>

          {showAnalysis ? (
            <AnalysisPanel
              data={analysisData}
              subjectComparison={subjectComparison}
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
                      {classStats && <th className="rank-col">Rank</th>}
                      {currentSheet.headers.map(h => <th key={h}>{h}</th>)}
                      <th className="actions-col">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRows.length === 0 ? (
                      <tr><td colSpan={currentSheet.headers.length + (classStats ? 3 : 2)} className="empty-msg">
                        {searchQuery ? `No students matching "${searchQuery}"` : 'No students added yet. Click "Add Student" to begin.'}
                      </td></tr>
                    ) : filteredRows.map((row, fi) => {
                      const ri = currentSheet.rows.indexOf(row);
                      const rank = classStats?.rankMap?.get(row);
                      return (
                      <tr key={ri}>
                        <td className="row-num">{ri + 1}</td>
                        {classStats && (
                          <td className="rank-col">
                            {rank ? (
                              <span className={`rank-badge ${rank <= 3 ? 'rank-top' : rank <= 10 ? 'rank-mid' : ''}`}>
                                {rank <= 3 ? ['🥇','🥈','🥉'][rank-1] : `#${rank}`}
                              </span>
                            ) : '—'}
                          </td>
                        )}
                        {currentSheet.headers.map(h => {
                          const isEditing = editingCell && editingCell.sheetName === activeSheet && editingCell.rowIdx === ri && editingCell.col === h;
                          const val = row[h];
                          const isSubjectCol = getSubjectColumns(currentSheet.headers).includes(h);
                          const notOpted = isSubjectCol && isNotOpted(val);
                          const scoreClass = (h.toLowerCase().includes('%') || h.toLowerCase().includes('total')) ? getScoreColor(val) : '';
                          return (
                            <td key={h} className={`data-cell ${scoreClass} ${notOpted ? 'not-opted-cell' : ''}`} onDoubleClick={() => startEdit(activeSheet, ri, h)}>
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
                                <span>{notOpted ? '—' : (val !== '' && val !== null && val !== undefined ? String(val) : '—')}</span>
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
                    );})}
                    {searchQuery && filteredRows.length > 0 && (
                      <tr><td colSpan={currentSheet.headers.length + (classStats ? 3 : 2)} className="search-info">Showing {filteredRows.length} of {currentSheet.rows.length} students</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
              <SheetAnalysis sheetName={activeSheet} sheet={currentSheet} />
            </>
          ) : null}

          <div className="footer-info">
            Target Analysis Report Generator — Double-click any cell to edit • Subjects marked with "-" are excluded from totals & graphs
          </div>
        </div>
      )}

      {/* ===== STUDENT REPORT MODAL ===== */}
      {studentReport && (
        <StudentReportModal 
          student={studentReport} 
          headers={currentSheet?.headers || []} 
          sheetName={activeSheet} 
          onClose={() => setStudentReport(null)} 
        />
      )}
    </div>
  );
}

// ========== ANALYSIS PANEL ==========
function AnalysisPanel({ data, subjectComparison, sheets, selected, setSelected, onClose }) {
  const toggleSheet = (name) => {
    setSelected(prev => prev.includes(name) ? prev.filter(s => s !== name) : [...prev, name]);
  };

  if (!data) return null;

  const pieData = data.rows.slice(0, -1).map(r => ({
    name: r.Range,
    value: r.students,
  }));
  const PIE_COLORS = ['#22c55e', '#3b82f6', '#8b5cf6', '#f59e0b', '#f97316', '#ef4444'];

  return (
    <div className="analysis-panel fade-in">
      <div className="analysis-header">
        <h3><BarChart3 size={20} /> Section-wise Target Analysis</h3>
        <button className="icon-btn" onClick={onClose}><X size={18} /></button>
      </div>

      <div className="analysis-sheets">
        <span className="label">Include sheets:</span>
        {sheets.map(s => (
          <button key={s} className={`sheet-pill ${selected.includes(s) ? 'active' : ''}`} onClick={() => toggleSheet(s)}>
            {selected.includes(s) && <CheckCircle size={14} />} {s}
          </button>
        ))}
      </div>

      {/* Charts */}
      <div className="chart-box" id="section-analysis-charts" style={{ padding: '2rem 1.5rem', background: '#fff' }}>
        <h4 style={{ marginBottom: '1.5rem', color: 'var(--text)', borderBottom: '1px solid var(--border)', paddingBottom: '0.75rem', fontSize: '1rem' }}>
          {data.sections.map(s => s.replace(/^X\s*/, 'X-')).join(', ')} — Performance Visuals
        </h4>
        
        <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)', gap: '2rem', marginBottom: '2rem' }}>
          <div className="chart-inner-box">
            <h5 className="chart-subtitle">Section-wise Bar Analysis</h5>
            <ResponsiveContainer width="100%" height={300}>
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

          <div className="chart-inner-box">
            <h5 className="chart-subtitle">Trend Line Analysis</h5>
            <ResponsiveContainer width="100%" height={300}>
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

        <div style={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1fr)', gap: '2rem' }}>
          <div className="chart-inner-box">
            <h5 className="chart-subtitle">Overall Distribution</h5>
            <ResponsiveContainer width="100%" height={300}>
              <PieChart>
                <Pie
                  data={pieData.filter(d => d.value > 0)}
                  cx="50%" cy="50%"
                  innerRadius={60} outerRadius={110}
                  paddingAngle={3}
                  dataKey="value"
                  label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                >
                  {pieData.map((entry, i) => (
                    <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>

          {subjectComparison && subjectComparison.data.length > 0 && (
            <div className="chart-inner-box">
              <h5 className="chart-subtitle">Subject-wise Average (Section Comparison)</h5>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={subjectComparison.data} margin={{ top: 10, right: 10, left: 0, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                  <XAxis dataKey="subject" tick={{ fontSize: 10 }} interval={0} angle={-25} textAnchor="end" height={60} />
                  <YAxis tick={{ fontSize: 11 }} />
                  <Tooltip contentStyle={{ borderRadius: 10, border: 'none', boxShadow: '0 4px 20px rgba(0,0,0,0.1)' }} />
                  <Legend />
                  {subjectComparison.sections.map((s, i) => (
                    <Bar key={s} dataKey={s} fill={CHART_COLORS[i % CHART_COLORS.length]} radius={[4, 4, 0, 0]} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
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

  const chartData = analysis.rows.slice(0, -1);
  const chartId = `sheet-chart-${sheetName.replace(/\s+/g, '_')}`;

  return (
    <div className="sheet-analysis fade-in">
      <h3 className="sheet-analysis-title">
        <BarChart3 size={18} /> {sheetName} — Score Distribution
      </h3>
      <div className="sheet-analysis-body">
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
        <div className="sheet-analysis-chart" id={chartId}>
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
function StudentReportModal({ student, headers, sheetName, onClose }) {
  const [downloading, setDownloading] = useState(false);
  const reportRef = useRef(null);

  const subjectCols = getSubjectColumns(headers);

  // Only opted subjects — exclude "-" / empty / not-opted
  const optedSubjects = subjectCols.filter(h => !isNotOpted(student[h]));
  
  // Not-opted subjects
  const notOptedSubjects = subjectCols.filter(h => isNotOpted(student[h]));

  const nameCol = headers.find(h => h.toLowerCase().includes('name') && !h.toLowerCase().includes('father') && !h.toLowerCase().includes('mother'));
  const name = nameCol ? student[nameCol] : 'Student';

  const totalCol = headers.find(h => h.toLowerCase().includes('grand total') || h.toLowerCase() === 'total');
  const total = totalCol ? student[totalCol] : null;

  const percentCol = headers.find(h => h.toLowerCase().includes('% in ix+30') || h.toLowerCase().includes('% in ix'));
  const percent = percentCol ? student[percentCol] : null;

  // Calculate obtained and max marks from opted subjects only
  let obtainedMarks = 0;
  let maxMarks = 0;
  optedSubjects.forEach(h => {
    const val = parseFloat(student[h]);
    if (!isNaN(val)) {
      obtainedMarks += val;
      const headerMax = getMaxMarksFromHeader(h);
      maxMarks += headerMax || 100; // default to 100 if max marks not in header
    }
  });
  const calculatedPercent = maxMarks > 0 ? parseFloat(((obtainedMarks / maxMarks) * 100).toFixed(2)) : null;

  // Chart data — only opted subjects
  const barData = optedSubjects.map((h, i) => ({
    name: h.replace(/\s*\+\s*30/g, '').replace(/\s+\d+$/, '').trim(),
    value: parseFloat(student[h]) || 0,
    maxMarks: getMaxMarksFromHeader(h) || 100,
    color: CHART_COLORS[i % CHART_COLORS.length]
  }));

  // Radar chart data
  const radarData = optedSubjects.map(h => {
    const maxM = getMaxMarksFromHeader(h) || 100;
    const score = parseFloat(student[h]) || 0;
    return {
      subject: h.replace(/\s*\+\s*30/g, '').replace(/\s+\d+$/, '').trim(),
      score,
      percentage: parseFloat(((score / maxM) * 100).toFixed(1)),
    };
  });

  // Metadata fields (non-subject, non-total)
  const metaFields = headers.filter(h => {
    return !subjectCols.includes(h) && 
      !h.toLowerCase().includes('grand total') && 
      !h.toLowerCase().includes('total') && 
      !h.toLowerCase().includes('%');
  });

  // PDF Download
  const downloadPDF = async () => {
    if (!reportRef.current) return;
    setDownloading(true);

    try {
      await new Promise(r => setTimeout(r, 500));

      const canvas = await html2canvas(reportRef.current, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        logging: false,
        windowWidth: 800,
      });

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');

      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 10;
      const usableWidth = pageWidth - (margin * 2);

      const imgWidth = canvas.width;
      const imgHeight = canvas.height;
      const ratio = usableWidth / imgWidth;
      const scaledHeight = imgHeight * ratio;

      if (scaledHeight <= pageHeight - (margin * 2)) {
        pdf.addImage(imgData, 'PNG', margin, margin, usableWidth, scaledHeight);
      } else {
        let remainingHeight = imgHeight;
        let sourceY = 0;
        const pageContentHeight = (pageHeight - (margin * 2)) / ratio;

        while (remainingHeight > 0) {
          const sliceHeight = Math.min(pageContentHeight, remainingHeight);
          const tempCanvas = document.createElement('canvas');
          tempCanvas.width = imgWidth;
          tempCanvas.height = sliceHeight;
          const ctx = tempCanvas.getContext('2d');
          ctx.drawImage(canvas, 0, sourceY, imgWidth, sliceHeight, 0, 0, imgWidth, sliceHeight);

          const sliceData = tempCanvas.toDataURL('image/png');
          const scaledSliceHeight = sliceHeight * ratio;

          if (sourceY > 0) pdf.addPage();
          pdf.addImage(sliceData, 'PNG', margin, margin, usableWidth, scaledSliceHeight);

          sourceY += sliceHeight;
          remainingHeight -= sliceHeight;
        }
      }

      pdf.save(`${name}_Report.pdf`);
    } catch (err) {
      console.error('PDF generation failed:', err);
      alert('PDF download failed. Try again.');
    }
    setDownloading(false);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-card fade-in" onClick={e => e.stopPropagation()} style={{ maxWidth: '850px' }}>
        <div className="modal-header">
          <h3><UserCheck size={20} /> Student Report</h3>
          <div style={{ display: 'flex', gap: '0.5rem' }}>
            <button className="tool-btn primary" onClick={downloadPDF} disabled={downloading}>
              {downloading ? <Loader2 size={14} className="spin" /> : <Download size={14} />}
              {downloading ? ' Generating...' : ' Download PDF'}
            </button>
            <button className="icon-btn" onClick={onClose}><X size={18} /></button>
          </div>
        </div>

        <div ref={reportRef} className="pdf-content">
          {/* Report Header */}
          <div className="pdf-header">
            <h2 className="pdf-title">Student Performance Report</h2>
            <p className="pdf-subtitle">{sheetName} • Generated on {new Date().toLocaleDateString('en-IN', { year: 'numeric', month: 'long', day: 'numeric' })}</p>
          </div>

          {/* Student Info Bar */}
          <div className="report-name-bar">
            <div className="student-avatar">{name?.charAt(0)}</div>
            <div>
              <h2>{name}</h2>
              {calculatedPercent !== null && (
                <span className={`score-badge ${getScoreColor(calculatedPercent)}`}>
                  {calculatedPercent}% ({calculatedPercent >= 90 ? 'Excellent' : calculatedPercent >= 80 ? 'Very Good' : calculatedPercent >= 60 ? 'Good' : calculatedPercent >= 50 ? 'Average' : 'Needs Improvement'})
                </span>
              )}
            </div>
            <div className="total-badge-group">
              {total !== null && total !== '' && (
                <div className="total-badge">
                  <span className="total-label">Obtained</span>
                  <span className="total-value">{obtainedMarks}</span>
                </div>
              )}
              {maxMarks > 0 && (
                <div className="total-badge">
                  <span className="total-label">Max Marks</span>
                  <span className="total-value total-max">{maxMarks}</span>
                </div>
              )}
            </div>
          </div>

          {/* Student Details */}
          {metaFields.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">📋 Student Details</h4>
              <div className="report-grid">
                {metaFields.map(h => {
                  const val = student[h];
                  return (
                    <div key={h} className="report-field">
                      <label>{h}</label>
                      <span>{val !== '' && val !== null && val !== undefined ? String(val) : '—'}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Opted Subjects - Score Matrix */}
          {optedSubjects.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">📊 Subject Scores (Opted Subjects Only)</h4>
              <div className="report-grid">
                {optedSubjects.map(h => {
                  const val = student[h];
                  const maxM = getMaxMarksFromHeader(h);
                  return (
                    <div key={h} className="report-field">
                      <label>{h}</label>
                      <span className={typeof val === 'number' ? getScoreColor((val / (maxM || 100)) * 100) : ''}>
                        {String(val)} {maxM ? `/ ${maxM}` : ''}
                      </span>
                    </div>
                  );
                })}
                {totalCol && (
                  <div className="report-field highlight-field">
                    <label>Grand Total</label>
                    <span className="score-excellent">{obtainedMarks} / {maxMarks}</span>
                  </div>
                )}
                {calculatedPercent !== null && (
                  <div className="report-field highlight-field">
                    <label>Percentage</label>
                    <span className={getScoreColor(calculatedPercent)}>{calculatedPercent}%</span>
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Not-opted Subjects */}
          {notOptedSubjects.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title" style={{ color: 'var(--text-secondary)', fontSize: '0.85rem' }}>
                🚫 Not Opted Subjects
              </h4>
              <div className="not-opted-list">
                {notOptedSubjects.map(h => (
                  <span key={h} className="not-opted-tag">{h.replace(/\s+\d+$/, '').trim()}</span>
                ))}
              </div>
            </div>
          )}

          {/* Subject Bar Chart */}
          {barData.length > 0 && (
            <div className="report-section">
              <h4 className="report-section-title">📈 Subject-wise Score Chart</h4>
              <div className="chart-box report-chart">
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={barData} margin={{ top: 10, right: 20, left: 0, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="rgba(0,0,0,0.05)" />
                    <XAxis dataKey="name" tick={{ fontSize: 10 }} interval={0} angle={-30} textAnchor="end" height={70} />
                    <YAxis tick={{ fontSize: 12 }} />
                    <Tooltip />
                    <Bar dataKey="value" name="Score" radius={[4, 4, 0, 0]}>
                      {barData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          {/* Radar Chart */}
          {radarData.length >= 3 && (
            <div className="report-section">
              <h4 className="report-section-title">🎯 Performance Radar</h4>
              <div className="chart-box report-chart">
                <ResponsiveContainer width="100%" height={280}>
                  <RadarChart data={radarData} cx="50%" cy="50%" outerRadius="75%">
                    <PolarGrid stroke="rgba(0,0,0,0.08)" />
                    <PolarAngleAxis dataKey="subject" tick={{ fontSize: 10 }} />
                    <PolarRadiusAxis tick={{ fontSize: 10 }} domain={[0, 100]} />
                    <Radar name="Score %" dataKey="percentage" stroke="#6366f1" fill="#6366f1" fillOpacity={0.25} strokeWidth={2} />
                    <Tooltip />
                  </RadarChart>
                </ResponsiveContainer>
              </div>
            </div>
          )}

          {/* Footer */}
          <div className="pdf-footer">
            <p>This report was auto-generated by the Target Analysis Report Generator.</p>
            <p>Subjects marked with "-" are not opted and excluded from Grand Total & percentage calculations.</p>
          </div>
        </div>
      </div>
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
