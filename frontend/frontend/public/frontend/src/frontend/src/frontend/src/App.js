import React, { useState, useRef, useCallback } from 'react';
import axios from 'axios';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  Cell, Legend
} from 'recharts';

const API = 'https://carbon-canary.onrender.com';
const STATUS_COLOR = { above: '#4ade80', average: '#fbbf24', below: '#f87171' };

function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: '#182118', border: '1px solid #1e2b1e', borderRadius: 8, padding: '10px 14px', fontSize: 12, fontFamily: 'IBM Plex Mono' }}>
      <div style={{ color: '#e2f0e4', fontWeight: 600, marginBottom: 6 }}>{label}</div>
      {payload.map(p => <div key={p.name} style={{ color: p.color, marginBottom: 2 }}>{p.name}: <strong>{p.value}</strong></div>)}
    </div>
  );
}

function ScoreBar({ score, status }) {
  return (
    <div style={{ width: 120, display: 'inline-block' }}>
      <div style={{ height: 6, background: '#1e2b1e', borderRadius: 3, overflow: 'hidden' }}>
        <div style={{ width: `${score}%`, height: '100%', background: STATUS_COLOR[status], borderRadius: 3 }} />
      </div>
    </div>
  );
}

// ── COMPETITOR VIEW (simple) ──────────────────────────────────────────────────
function CompetitorView({ competitorView }) {
  if (!competitorView) return null;
  const { leaderboard, yourRank, totalCompanies } = competitorView;
  const chartData = leaderboard.map(c => ({ name: c.isYou ? 'You' : c.name, score: c.overallScore, isYou: c.isYou }));

  return (
    <div className="panel full" style={{ marginBottom: 20 }}>
      <div className="panel-title">How You Compare</div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 16, marginBottom: 20 }}>
        <div style={{
          width: 64, height: 64, borderRadius: '50%', background: 'rgba(74,222,128,0.12)',
          border: '2px solid #4ade80', display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontFamily: 'IBM Plex Mono', fontSize: 22, fontWeight: 700, color: '#4ade80', flexShrink: 0,
        }}>#{yourRank}</div>
        <div>
          <div style={{ fontSize: 15, fontWeight: 600 }}>You rank #{yourRank} of {totalCompanies}</div>
          <div style={{ fontSize: 12, color: '#6b8f72' }}>Compared to similar companies (sample data)</div>
        </div>
      </div>

      <ResponsiveContainer width="100%" height={Math.max(180, chartData.length * 46)}>
        <BarChart data={chartData} layout="vertical" margin={{ top: 4, right: 30, bottom: 4, left: 8 }}>
          <CartesianGrid strokeDasharray="3 3" stroke="#1e2b1e" horizontal={false} />
          <XAxis type="number" domain={[0, 100]} tick={{ fill: '#6b8f72', fontSize: 10 }} />
          <YAxis type="category" dataKey="name" width={110} tick={{ fill: '#e2f0e4', fontSize: 11, fontFamily: 'IBM Plex Mono' }} />
          <Tooltip content={<CustomTooltip />} />
          <Bar dataKey="score" name="Score" radius={[0, 4, 4, 0]} barSize={20}>
            {chartData.map((c, i) => <Cell key={i} fill={c.isYou ? '#4ade80' : '#374a37'} />)}
          </Bar>
        </BarChart>
      </ResponsiveContainer>

      <div style={{ fontSize: 11, color: '#6b8f72', marginTop: 12 }}>Sample data for illustration only.</div>
    </div>
  );
}

// ── RECOMMENDATIONS PANEL (simple) ────────────────────────────────────────────
function RecommendationsPanel({ recommendations }) {
  if (!recommendations || recommendations.length === 0) return null;

  return (
    <div className="panel full" style={{ marginBottom: 20 }}>
      <div className="panel-title">What To Improve</div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
        {recommendations.map((rec) => (
          <div key={rec.category} style={{
            display: 'flex', alignItems: 'flex-start', gap: 14, padding: '14px 16px',
            background: '#182118', border: '1px solid #1e2b1e', borderRadius: 10,
          }}>
            <div style={{ fontSize: 20, lineHeight: 1, flexShrink: 0 }}>
              {rec.priority === 'high' ? '🔴' : '🟡'}
            </div>
            <div>
              <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 4 }}>{rec.category}</div>
              <div style={{ fontSize: 13, color: '#a8c4ab', lineHeight: 1.5 }}>{rec.suggestions[0]}</div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── DOWNLOAD EXCEL REPORT ──────────────────────────────────────────────────────
function downloadExcel(data) {
  const isInvoice = data.documentType === 'invoice';
  const wb = XLSX.utils.book_new();

  if (isInvoice) {
    const summary = [
      ['Carbon Canary - Invoice CO2 Emission Report'],
      ['File', data.fileName],
      ['Supplier', data.supplierName || 'Unknown'],
      ['Invoice Number', data.invoiceNumber || 'Unknown'],
      ['Invoice Date', data.invoiceDate || 'Unknown'],
      ['Total CO2 (kg)', data.totalCO2kg],
      ['Total CO2 (tonnes)', data.totalCO2t],
      ['Emission Rating', data.co2Rating],
      ['Confidence Score', `${data.confidenceScore}/10`],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summary);
    wsSummary['!cols'] = [{ wch: 22 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    const breakdownHeader = ['Item', 'Category', 'Quantity', 'Unit', 'EU Factor (kg CO2/unit)', 'CO2 (kg)', 'Estimated'];
    const breakdownRows = data.emissionBreakdown.map(e => [e.item, e.category, e.quantity, e.unit, e.factor, e.co2kg, e.estimated ? 'Yes' : 'No']);
    const wsBreakdown = XLSX.utils.aoa_to_sheet([breakdownHeader, ...breakdownRows, [], ['TOTAL', '', '', '', '', data.totalCO2kg]]);
    wsBreakdown['!cols'] = [{ wch: 24 }, { wch: 16 }, { wch: 10 }, { wch: 8 }, { wch: 20 }, { wch: 12 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, wsBreakdown, 'Emission Breakdown');
  } else {
    const summary = [
      ['Carbon Canary - CSRD Sustainability Report Scoring'],
      ['File', data.fileName],
      ['Overall Score', data.overallScore],
      ['Industry Benchmark', data.overallBenchmark],
      ['Gap', data.overallScore - data.overallBenchmark],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summary);
    wsSummary['!cols'] = [{ wch: 22 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');
  }

  const catHeader = ['Category', 'Score', 'Benchmark', 'Gap', 'Status', 'Insight'];
  const catRows = data.results.map(r => [r.category, r.score, r.benchmark, r.gap, r.status, r.insight]);
  const wsCat = XLSX.utils.aoa_to_sheet([catHeader, ...catRows]);
  wsCat['!cols'] = [{ wch: 22 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 50 }];
  XLSX.utils.book_append_sheet(wb, wsCat, 'Category Performance');

  if (!isInvoice) {
    const wsStrongWeak = XLSX.utils.aoa_to_sheet([
      ['STRONG AREAS'], ...data.strong.map(s => [s]), [],
      ['WEAK AREAS'], ...data.weak.map(w => [w]),
    ]);
    wsStrongWeak['!cols'] = [{ wch: 30 }];
    XLSX.utils.book_append_sheet(wb, wsStrongWeak, 'Highlights');
  }

  XLSX.writeFile(wb, `carbon-canary-report-${data.fileName.replace(/\.[^.]+$/, '')}.xlsx`);
}

// ── DOWNLOAD PDF REPORT ────────────────────────────────────────────────────────
function downloadPDF(data) {
  const isInvoice = data.documentType === 'invoice';
  const date = new Date().toLocaleDateString('en-GB');
  const doc = new jsPDF();
  const pageW = doc.internal.pageSize.getWidth();
  const GREEN = [74, 222, 128];
  const FOREST = [10, 15, 13];
  const INK = [30, 30, 30];
  const MUTED = [110, 120, 115];

  // Header band
  doc.setFillColor(...FOREST);
  doc.rect(0, 0, pageW, 32, 'F');
  doc.setTextColor(...GREEN);
  doc.setFontSize(16);
  doc.setFont(undefined, 'bold');
  doc.text('Carbon Canary', 14, 14);
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(12);
  doc.text(isInvoice ? 'Invoice CO2 Emission Report' : 'CSRD Sustainability Report Scoring', 14, 22);
  doc.setFontSize(8);
  doc.setTextColor(200, 220, 205);
  doc.setFont(undefined, 'normal');
  doc.text(`${data.fileName}  ·  Generated ${date}  ·  Climathon 2026`, 14, 28);

  let y = 42;

  // KPI summary line
  doc.setTextColor(...INK);
  doc.setFontSize(10);
  doc.setFont(undefined, 'bold');
  if (isInvoice) {
    doc.text(`Total CO2: ${data.totalCO2kg} kg  (${data.totalCO2t} t)`, 14, y);
    doc.text(`Rating: ${data.co2Rating}`, 100, y);
    doc.text(`Confidence: ${data.confidenceScore}/10`, 150, y);
    y += 6;
    doc.setFont(undefined, 'normal');
    doc.setFontSize(9);
    doc.text(`Supplier: ${data.supplierName || 'Unknown'}   Invoice: ${data.invoiceNumber || 'Unknown'}   Date: ${data.invoiceDate || 'Unknown'}`, 14, y);
    y += 8;
    if (data.notes) {
      doc.setTextColor(...MUTED);
      doc.setFontSize(8);
      const noteLines = doc.splitTextToSize(data.notes, pageW - 28);
      doc.text(noteLines, 14, y);
      y += noteLines.length * 4 + 4;
    }
  } else {
    doc.text(`Overall Score: ${data.overallScore} / 100`, 14, y);
    doc.text(`Benchmark: ${data.overallBenchmark}`, 90, y);
    doc.text(`Gap: ${data.overallScore - data.overallBenchmark > 0 ? '+' : ''}${data.overallScore - data.overallBenchmark}`, 150, y);
    y += 8;
  }

  // Emission breakdown table (invoice only)
  if (isInvoice) {
    doc.setTextColor(...INK);
    doc.setFontSize(11);
    doc.setFont(undefined, 'bold');
    doc.text('Emission Breakdown', 14, y);
    y += 4;
    autoTable(doc, {
      startY: y,
      head: [['Item', 'Category', 'Quantity', 'EU Factor', 'CO2 (kg)']],
      body: data.emissionBreakdown.map(e => [
        e.item + (e.estimated ? ' (est.)' : ''), e.category, `${e.quantity} ${e.unit}`, e.factor, e.co2kg,
      ]),
      foot: [['TOTAL', '', '', '', `${data.totalCO2kg} kg`]],
      theme: 'striped',
      headStyles: { fillColor: FOREST, textColor: 255, fontSize: 8 },
      footStyles: { fillColor: [230, 240, 232], textColor: INK, fontStyle: 'bold', fontSize: 8 },
      bodyStyles: { fontSize: 8 },
      margin: { left: 14, right: 14 },
    });
    y = doc.lastAutoTable.finalY + 10;
  }

  // Category performance table
  if (y > 250) { doc.addPage(); y = 16; }
  doc.setTextColor(...INK);
  doc.setFontSize(11);
  doc.setFont(undefined, 'bold');
  doc.text('Category Performance vs Benchmark', 14, y);
  y += 4;
  autoTable(doc, {
    startY: y,
    head: [['Category', 'Score', 'Benchmark', 'Gap', 'Status', 'Insight']],
    body: data.results.map(r => [r.category, r.score, r.benchmark, (r.gap > 0 ? '+' : '') + r.gap, r.status, r.insight]),
    theme: 'striped',
    headStyles: { fillColor: FOREST, textColor: 255, fontSize: 8 },
    bodyStyles: { fontSize: 7.5 },
    columnStyles: { 5: { cellWidth: 60 } },
    margin: { left: 14, right: 14 },
  });
  y = doc.lastAutoTable.finalY + 10;

  // Strong / Weak (report mode)
  if (!isInvoice) {
    if (y > 260) { doc.addPage(); y = 16; }
    doc.setFontSize(9);
    doc.setFont(undefined, 'bold');
    doc.setTextColor(22, 163, 74);
    doc.text(`Strong Areas: ${data.strong.join(', ') || 'None'}`, 14, y);
    y += 6;
    doc.setTextColor(220, 38, 38);
    doc.text(`Weak Areas: ${data.weak.join(', ') || 'None'}`, 14, y);
    y += 10;
  }

  // Footer
  doc.setFontSize(7.5);
  doc.setFont(undefined, 'italic');
  doc.setTextColor(...MUTED);
  doc.text('Carbon Canary · Scope 3 Emission Analytics · DE/PL Corridor · Climathon 2026', 14, 290);

  doc.save(`carbon-canary-report-${data.fileName.replace(/\.[^.]+$/, '')}.pdf`);
}

// ── DOWNLOAD BUTTONS ──────────────────────────────────────────────────────────
function DownloadButtons({ data }) {
  return (
    <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
      <button onClick={() => downloadPDF(data)} style={{
        padding: '10px 20px', background: '#4ade80', color: '#050a06', border: 'none',
        borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: 'pointer', fontFamily: 'inherit'
      }}>⬇ Download PDF Report</button>
      <button onClick={() => downloadExcel(data)} style={{
        padding: '10px 20px', background: 'transparent', color: '#4ade80',
        border: '1px solid #4ade80', borderRadius: 8, fontWeight: 600, fontSize: 13,
        cursor: 'pointer', fontFamily: 'inherit'
      }}>⬇ Download Excel Data</button>
    </div>
  );
}

// ── UPLOAD VIEW ───────────────────────────────────────────────────────────────
function UploadView({ onResult }) {
  const [dragging, setDragging] = useState(false);
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const inputRef = useRef();

  const handleFile = (f) => { setError(''); setFile(f); };
  const handleDrop = useCallback((e) => { e.preventDefault(); setDragging(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }, []);

  const handleSubmit = async () => {
    if (!file) return;
    setLoading(true); setError('');
    const fd = new FormData();
    fd.append('file', file);
    try {
      const { data } = await axios.post(`${API}/api/upload`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      onResult(data);
    } catch (err) {
      setError(err.response?.data?.error || 'Upload failed. Please try again.');
    } finally { setLoading(false); }
  };

  if (loading) return (
    <div className="upload-section">
      <div className="loading">
        <div className="spinner" />
        <div className="loading-text">Analyzing document & calculating CO₂…</div>
      </div>
    </div>
  );

  return (
    <div className="upload-section">
      <div className="upload-eyebrow">🐦 Carbon Canary</div>
      <h1 className="upload-headline">Scope 3 <em>Carbon</em><br />Intelligence Dashboard</h1>
      <p className="upload-sub">Upload a <strong style={{ color: '#4ade80' }}>supplier invoice</strong> to calculate CO₂, or a <strong style={{ color: '#fbbf24' }}>sustainability report</strong> to score CSRD compliance.</p>
      <div className={`dropzone${dragging ? ' active' : ''}`}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current.click()}>
        <input ref={inputRef} type="file" accept=".pdf,.xlsx,.xls,.csv,.txt"
          onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
          onClick={(e) => e.stopPropagation()} />
        <div className="dropzone-icon">📄</div>
        <div className="dropzone-label">Drop your file here or click to browse</div>
        <div className="dropzone-types">PDF · XLSX · CSV · TXT</div>
      </div>
      {file && <div className="file-selected"><span>✓</span><span>{file.name}</span><span style={{ color: '#6b8f72' }}>({(file.size / 1024).toFixed(1)} KB)</span></div>}
      {error && <div className="error-box">{error}</div>}
      <div style={{ display: 'flex', gap: 12, marginTop: 20, flexWrap: 'wrap', justifyContent: 'center' }}>
        <div style={{ background: 'rgba(74,222,128,0.08)', border: '1px solid rgba(74,222,128,0.2)', borderRadius: 8, padding: '8px 16px', fontSize: 12, color: '#4ade80', fontFamily: 'IBM Plex Mono' }}>📦 Invoice → CO₂ Calculation</div>
        <div style={{ background: 'rgba(251,191,36,0.08)', border: '1px solid rgba(251,191,36,0.2)', borderRadius: 8, padding: '8px 16px', fontSize: 12, color: '#fbbf24', fontFamily: 'IBM Plex Mono' }}>📊 Report → CSRD Scoring</div>
      </div>
      <button className="upload-btn" disabled={!file} onClick={handleSubmit} style={{ marginTop: 20 }}>Analyze Document →</button>
    </div>
  );
}

// ── INVOICE DASHBOARD ─────────────────────────────────────────────────────────
function InvoiceDashboard({ data, onReset }) {
  const { fileName, totalCO2kg, totalCO2t, co2Rating, co2Color, emissionBreakdown, results, strong, weak, overallScore, supplierName, invoiceNumber, invoiceDate, confidenceScore, notes, competitorView, recommendations } = data;
  const barData = emissionBreakdown.map(e => ({ name: e.item.length > 14 ? e.item.slice(0, 14) + '…' : e.item, co2: e.co2kg }));
  const radarData = results.map(r => ({ category: r.category.split(' ')[0], Score: r.score, Benchmark: r.benchmark }));
  const confColor = confidenceScore >= 7 ? '#4ade80' : confidenceScore >= 4 ? '#fbbf24' : '#f87171';

  return (
    <div className="dashboard">
      <div className="dash-header">
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <div className="dash-title">Invoice CO₂ Emission Report</div>
            <span style={{ background: 'rgba(74,222,128,0.12)', color: '#4ade80', padding: '2px 10px', borderRadius: 4, fontSize: 11, fontFamily: 'IBM Plex Mono' }}>INVOICE</span>
            <span style={{ background: `${confColor}1F`, color: confColor, padding: '2px 10px', borderRadius: 4, fontSize: 11, fontFamily: 'IBM Plex Mono' }}>CONFIDENCE {confidenceScore}/10</span>
          </div>
          <div className="dash-filename">📄 {fileName} · {supplierName} · {invoiceNumber} · {invoiceDate}</div>
        </div>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <DownloadButtons data={data} />
          <button className="reset-btn" onClick={onReset}>← New File</button>
        </div>
      </div>

      <div className="kpi-grid">
        <div className={`kpi-card ${co2Color}`} style={{ gridColumn: 'span 2' }}>
          <div className="kpi-label">Total CO₂ Emissions</div>
          <div className="kpi-value" style={{ fontSize: 44 }}>{totalCO2kg} <span style={{ fontSize: 18 }}>kg</span></div>
          <div className="kpi-sub">{totalCO2t} tonnes CO₂eq · {co2Rating} emission level</div>
        </div>
        <div className="kpi-card"><div className="kpi-label">Line Items</div><div className="kpi-value" style={{ color: '#e2f0e4' }}>{emissionBreakdown.length}</div><div className="kpi-sub">Analyzed</div></div>
        <div className={`kpi-card ${co2Color}`}><div className="kpi-label">Rating</div><div className="kpi-value">{co2Rating}</div><div className="kpi-sub">{totalCO2kg < 500 ? '< 500kg Low' : totalCO2kg < 2000 ? '500–2000kg Medium' : '> 2000kg High'}</div></div>
        <div className="kpi-card"><div className="kpi-label">Overall Score</div><div className="kpi-value" style={{ color: '#e2f0e4' }}>{overallScore}</div><div className="kpi-sub">vs benchmark 65</div></div>
      </div>

      {notes && <div style={{ background: 'rgba(74,222,128,0.06)', border: '1px solid rgba(74,222,128,0.15)', borderRadius: 10, padding: '12px 20px', marginBottom: 20, fontSize: 13, color: '#6b8f72' }}>⚠ {notes}</div>}

      <div className="dash-grid" style={{ marginBottom: 20 }}>
        <div className="panel">
          <div className="panel-title">CO₂ by Line Item (kg)</div>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={barData} margin={{ top: 4, right: 16, bottom: 60, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e2b1e" />
              <XAxis dataKey="name" tick={{ fill: '#6b8f72', fontSize: 10, fontFamily: 'IBM Plex Mono' }} angle={-35} textAnchor="end" interval={0} />
              <YAxis tick={{ fill: '#6b8f72', fontSize: 10 }} />
              <Tooltip content={<CustomTooltip />} />
              <Bar dataKey="co2" name="CO₂ (kg)" radius={[4, 4, 0, 0]}>
                {barData.map((_, i) => <Cell key={i} fill={i % 2 === 0 ? '#4ade80' : '#22c55e'} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="panel">
          <div className="panel-title">Score vs Benchmark by Category</div>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={radarData} layout="vertical" margin={{ top: 4, right: 24, bottom: 4, left: 8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e2b1e" horizontal={false} />
              <XAxis type="number" domain={[0, 100]} tick={{ fill: '#6b8f72', fontSize: 10 }} />
              <YAxis type="category" dataKey="category" width={90} tick={{ fill: '#e2f0e4', fontSize: 11, fontFamily: 'IBM Plex Mono' }} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, fontFamily: 'IBM Plex Mono', color: '#6b8f72' }} />
              <Bar dataKey="Score" fill="#4ade80" radius={[0, 4, 4, 0]} barSize={14} />
              <Bar dataKey="Benchmark" fill="#fbbf24" radius={[0, 4, 4, 0]} barSize={14} fillOpacity={0.5} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="panel full" style={{ marginBottom: 20 }}>
        <div className="panel-title">Detailed Emission Breakdown</div>
        <div style={{ overflowX: 'auto' }}>
          <table className="cat-table">
            <thead><tr><th>Item</th><th>Category</th><th>Quantity</th><th>EU Factor (kg/unit)</th><th style={{ textAlign: 'right' }}>CO₂ (kg)</th></tr></thead>
            <tbody>
              {emissionBreakdown.map((e, i) => (
                <tr key={i}>
                  <td><div className="cat-name" style={{ fontSize: 12 }}>{e.item}{e.estimated && <span style={{ color: '#fbbf24', fontSize: 10, marginLeft: 6 }}>⚠ estimated</span>}</div></td>
                  <td style={{ fontSize: 12, color: '#6b8f72' }}>{e.category}</td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 12 }}>{e.quantity} {e.unit}</td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 12, color: '#6b8f72' }}>{e.factor}</td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 14, fontWeight: 700, color: '#4ade80', textAlign: 'right' }}>{e.co2kg}</td>
                </tr>
              ))}
              <tr style={{ borderTop: '2px solid #1e2b1e' }}>
                <td colSpan={3} style={{ fontWeight: 700, fontSize: 14 }}>TOTAL</td>
                <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 18, fontWeight: 700, color: co2Color === 'above' ? '#4ade80' : co2Color === 'below' ? '#f87171' : '#fbbf24', textAlign: 'right' }} colSpan={2}>{totalCO2kg} kg</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      <div className="panel full" style={{ marginBottom: 20 }}>
        <div className="panel-title">Category Performance vs Benchmark</div>
        <div style={{ overflowX: 'auto' }}>
          <table className="cat-table">
            <thead><tr><th>Category</th><th>Score</th><th>Progress</th><th>Benchmark</th><th>Status</th><th style={{ textAlign: 'right' }}>Gap</th></tr></thead>
            <tbody>
              {results.map(r => (
                <tr key={r.category}>
                  <td><div className="cat-name">{r.category}</div><div className="cat-desc">{r.description}</div></td>
                  <td><span className="score-num" style={{ color: STATUS_COLOR[r.status] }}>{r.score}</span></td>
                  <td><ScoreBar score={r.score} status={r.status} /></td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 13, color: '#6b8f72' }}>{r.benchmark}</td>
                  <td><span className={`badge ${r.status}`}>{r.status}</span></td>
                  <td><div className={`gap-num ${r.gap >= 0 ? 'gap-pos' : 'gap-neg'}`}>{r.gap > 0 ? '+' : ''}{r.gap}</div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <CompetitorView competitorView={competitorView} />
      <RecommendationsPanel recommendations={recommendations} />

      <div className="panel full">
        <div className="panel-title">Category Highlights</div>
        <div className="sw-grid">
          <div>
            <div className="sw-title" style={{ color: '#4ade80' }}>↑ Strong ({strong.length})</div>
            {strong.length === 0 ? <div className="sw-empty">No categories above benchmark</div> : strong.map(c => <div key={c} className="sw-item strong">✓ {c}</div>)}
          </div>
          <div>
            <div className="sw-title" style={{ color: '#f87171' }}>↓ Weak ({weak.length})</div>
            {weak.length === 0 ? <div className="sw-empty">No categories below benchmark</div> : weak.map(c => <div key={c} className="sw-item weak">✗ {c}</div>)}
          </div>
        </div>
      </div>
    </div>
  );
}

// ── REPORT DASHBOARD ──────────────────────────────────────────────────────────
function ReportDashboard({ data, onReset }) {
  const { fileName, overallScore, overallBenchmark, results, strong, weak, competitorView, recommendations } = data;
  const overallStatus = overallScore >= overallBenchmark + 10 ? 'above' : overallScore <= overallBenchmark - 10 ? 'below' : 'average';
  const barData = results.map(r => ({ name: r.category.length > 14 ? r.category.slice(0, 14) + '…' : r.category, Score: r.score, Benchmark: r.benchmark, status: r.status }));
  const radarData = results.map(r => ({ category: r.category.split(' ')[0], Score: r.score, Benchmark: r.benchmark }));

  return (
    <div className="dashboard">
      <div className="dash-header">
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div className="dash-title">CSRD Sustainability Scoring</div>
            <span style={{ background: 'rgba(251,191,36,0.12)', color: '#fbbf24', padding: '2px 10px', borderRadius: 4, fontSize: 11, fontFamily: 'IBM Plex Mono' }}>REPORT</span>
          </div>
          <div className="dash-filename">📄 {fileName}</div>
        </div>
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <DownloadButtons data={data} />
          <button className="reset-btn" onClick={onReset}>← New File</button>
        </div>
      </div>

      <div className="kpi-grid">
        <div className={`kpi-card ${overallStatus}`}><div className="kpi-label">Overall Score</div><div className="kpi-value">{overallScore}</div><div className="kpi-sub">Benchmark: {overallBenchmark}/100</div></div>
        <div className="kpi-card"><div className="kpi-label">Categories</div><div className="kpi-value" style={{ color: '#e2f0e4' }}>{results.length}</div><div className="kpi-sub">Assessed</div></div>
        <div className="kpi-card above"><div className="kpi-label">Strong</div><div className="kpi-value">{strong.length}</div><div className="kpi-sub">Above benchmark</div></div>
        <div className="kpi-card below"><div className="kpi-label">Weak</div><div className="kpi-value">{weak.length}</div><div className="kpi-sub">Below benchmark</div></div>
        <div className={`kpi-card ${overallStatus}`}><div className="kpi-label">vs Industry</div><div className="kpi-value">{overallScore - overallBenchmark > 0 ? '+' : ''}{overallScore - overallBenchmark}</div><div className="kpi-sub">Gap</div></div>
      </div>

      <div className="dash-grid" style={{ marginBottom: 20 }}>
        <div className="panel">
          <div className="panel-title">Score vs Benchmark</div>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={barData} margin={{ top: 4, right: 16, bottom: 60, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e2b1e" />
              <XAxis dataKey="name" tick={{ fill: '#6b8f72', fontSize: 10, fontFamily: 'IBM Plex Mono' }} angle={-40} textAnchor="end" interval={0} />
              <YAxis tick={{ fill: '#6b8f72', fontSize: 10 }} domain={[0, 100]} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, fontFamily: 'IBM Plex Mono', color: '#6b8f72', paddingTop: 8 }} />
              <Bar dataKey="Score" radius={[4, 4, 0, 0]}>{barData.map((e, i) => <Cell key={i} fill={STATUS_COLOR[e.status]} />)}</Bar>
              <Bar dataKey="Benchmark" fill="#1e2b1e" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="panel">
          <div className="panel-title">Score vs Benchmark by Category</div>
          <ResponsiveContainer width="100%" height={280}>
            <BarChart data={radarData} layout="vertical" margin={{ top: 4, right: 24, bottom: 4, left: 8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#1e2b1e" horizontal={false} />
              <XAxis type="number" domain={[0, 100]} tick={{ fill: '#6b8f72', fontSize: 10 }} />
              <YAxis type="category" dataKey="category" width={90} tick={{ fill: '#e2f0e4', fontSize: 11, fontFamily: 'IBM Plex Mono' }} />
              <Tooltip content={<CustomTooltip />} />
              <Legend wrapperStyle={{ fontSize: 11, fontFamily: 'IBM Plex Mono', color: '#6b8f72' }} />
              <Bar dataKey="Score" fill="#4ade80" radius={[0, 4, 4, 0]} barSize={14} />
              <Bar dataKey="Benchmark" fill="#fbbf24" radius={[0, 4, 4, 0]} barSize={14} fillOpacity={0.5} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="panel full" style={{ marginBottom: 20 }}>
        <div className="panel-title">Gap Heatmap</div>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(160px,1fr))', gap: 8 }}>
          {[...results].sort((a, b) => a.gap - b.gap).map(r => {
            const intensity = Math.min(Math.abs(r.gap) / 40, 1);
            const bg = r.gap >= 0 ? `rgba(74,222,128,${0.08 + intensity * 0.3})` : `rgba(248,113,113,${0.08 + intensity * 0.3})`;
            return (
              <div key={r.category} style={{ background: bg, border: `1px solid ${r.gap >= 0 ? 'rgba(74,222,128,0.2)' : 'rgba(248,113,113,0.2)'}`, borderRadius: 8, padding: 12 }}>
                <div style={{ fontSize: 10, fontFamily: 'IBM Plex Mono', color: '#6b8f72', marginBottom: 4 }}>{r.category}</div>
                <div style={{ fontSize: 20, fontFamily: 'IBM Plex Mono', fontWeight: 700, color: r.gap >= 0 ? '#4ade80' : '#f87171' }}>{r.gap > 0 ? '+' : ''}{r.gap}</div>
                <div style={{ fontSize: 10, color: '#6b8f72', marginTop: 2 }}>vs benchmark</div>
              </div>
            );
          })}
        </div>
      </div>

      <div className="panel full" style={{ marginBottom: 20 }}>
        <div className="panel-title">Category Breakdown</div>
        <div style={{ overflowX: 'auto' }}>
          <table className="cat-table">
            <thead><tr><th>Category</th><th>Score</th><th>Progress</th><th>Benchmark</th><th>Status</th><th style={{ textAlign: 'right' }}>Gap</th></tr></thead>
            <tbody>
              {results.map(r => (
                <tr key={r.category}>
                  <td><div className="cat-name">{r.category}</div><div className="cat-desc">{r.description}</div></td>
                  <td><span className="score-num" style={{ color: STATUS_COLOR[r.status] }}>{r.score}</span></td>
                  <td><ScoreBar score={r.score} status={r.status} /></td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 13, color: '#6b8f72' }}>{r.benchmark}</td>
                  <td><span className={`badge ${r.status}`}>{r.status}</span></td>
                  <td><div className={`gap-num ${r.gap >= 0 ? 'gap-pos' : 'gap-neg'}`}>{r.gap > 0 ? '+' : ''}{r.gap}</div></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <CompetitorView competitorView={competitorView} />
      <RecommendationsPanel recommendations={recommendations} />

      <div className="panel full">
        <div className="panel-title">Category Highlights</div>
        <div className="sw-grid">
          <div>
            <div className="sw-title" style={{ color: '#4ade80' }}>↑ Strong ({strong.length})</div>
            {strong.length === 0 ? <div className="sw-empty">No categories above benchmark</div> : strong.map(c => <div key={c} className="sw-item strong">✓ {c}</div>)}
          </div>
          <div>
            <div className="sw-title" style={{ color: '#f87171' }}>↓ Weak ({weak.length})</div>
            {weak.length === 0 ? <div className="sw-empty">No categories below benchmark</div> : weak.map(c => <div key={c} className="sw-item weak">✗ {c}</div>)}
          </div>
        </div>
      </div>
    </div>
  );
}

// ── ROOT ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [result, setResult] = useState(null);
  return (
    <div className="app">
      <header className="header">
        <span className="header-logo">🐦 Carbon Canary</span>
        <span className="header-sub">Scope 3 Analytics · DE/PL Corridor · Climathon 2026</span>
      </header>
      {!result && <UploadView onResult={setResult} />}
      {result?.documentType === 'invoice' && <InvoiceDashboard data={result} onReset={() => setResult(null)} />}
      {result?.documentType === 'report' && <ReportDashboard data={result} onReset={() => setResult(null)} />}
    </div>
  );
}
