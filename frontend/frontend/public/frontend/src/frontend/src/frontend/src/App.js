import React, { useState, useRef, useCallback } from 'react';
import axios from 'axios';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
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

      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {leaderboard.map((c, i) => (
          <div key={c.name} style={{
            display: 'flex', alignItems: 'center', gap: 12, padding: '10px 14px', borderRadius: 8,
            background: c.isYou ? 'rgba(74,222,128,0.08)' : 'transparent',
            border: c.isYou ? '1px solid rgba(74,222,128,0.25)' : '1px solid transparent',
          }}>
            <div style={{ fontFamily: 'IBM Plex Mono', fontSize: 12, color: '#6b8f72', width: 18 }}>{i + 1}</div>
            <div style={{
              fontSize: 13, fontWeight: c.isYou ? 700 : 400, color: c.isYou ? '#4ade80' : '#e2f0e4', flex: 1,
            }}>{c.isYou ? 'You' : c.name}</div>
            <div style={{
              fontFamily: 'IBM Plex Mono', fontSize: 14, fontWeight: 700,
              color: c.isYou ? '#4ade80' : '#a8c4ab',
            }}>{c.overallScore}</div>
          </div>
        ))}
      </div>
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

// ── DOWNLOAD CSV REPORT ───────────────────────────────────────────────────────
function downloadCSV(data) {
  const isInvoice = data.documentType === 'invoice';
  let csv = '';

  if (isInvoice) {
    csv += `Carbon Canary - Invoice CO2 Emission Report\n`;
    csv += `File,${data.fileName}\n`;
    csv += `Supplier,${data.supplierName || 'Unknown'}\n`;
    csv += `Invoice Number,${data.invoiceNumber || 'Unknown'}\n`;
    csv += `Invoice Date,${data.invoiceDate || 'Unknown'}\n`;
    csv += `Total CO2 (kg),${data.totalCO2kg}\n`;
    csv += `Total CO2 (tonnes),${data.totalCO2t}\n`;
    csv += `Emission Rating,${data.co2Rating}\n`;
    csv += `Confidence,${data.confidence}\n\n`;
    csv += `EMISSION BREAKDOWN\n`;
    csv += `Item,Category,Quantity,Unit,EU Factor (kg CO2/unit),CO2 (kg),Estimated,Source\n`;
    data.emissionBreakdown.forEach(e => {
      csv += `${e.item},${e.category},${e.quantity},${e.unit},${e.factor},${e.co2kg},${e.estimated ? 'Yes' : 'No'},${e.source}\n`;
    });
    csv += `\nTOTAL,,,,, ${data.totalCO2kg} kg,,\n\n`;
    csv += `CATEGORY PERFORMANCE\n`;
    csv += `Category,Score,Benchmark,Gap,Status,Insight\n`;
    data.results.forEach(r => {
      csv += `${r.category},${r.score},${r.benchmark},${r.gap},${r.status},${r.insight}\n`;
    });
  } else {
    csv += `Carbon Canary - CSRD Sustainability Report Scoring\n`;
    csv += `File,${data.fileName}\n`;
    csv += `Overall Score,${data.overallScore}\n`;
    csv += `Industry Benchmark,${data.overallBenchmark}\n`;
    csv += `Gap,${data.overallScore - data.overallBenchmark}\n\n`;
    csv += `CATEGORY SCORES\n`;
    csv += `Category,Score,Benchmark,Gap,Status,Insight\n`;
    data.results.forEach(r => {
      csv += `${r.category},${r.score},${r.benchmark},${r.gap},${r.status},${r.insight}\n`;
    });
    csv += `\nSTRONG AREAS\n`;
    data.strong.forEach(s => { csv += `${s}\n`; });
    csv += `\nWEAK AREAS\n`;
    data.weak.forEach(w => { csv += `${w}\n`; });
  }

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `carbon-canary-report-${data.fileName.replace(/\.[^.]+$/, '')}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── DOWNLOAD HTML REPORT ──────────────────────────────────────────────────────
function downloadHTML(data) {
  const isInvoice = data.documentType === 'invoice';
  const date = new Date().toLocaleDateString('en-GB');

  const rowsHTML = isInvoice
    ? (data.emissionBreakdown || []).map(e => `
      <tr>
        <td>${e.item}${e.estimated ? ' <span style="color:#f59e0b;font-size:11px">⚠ est.</span>' : ''}</td>
        <td>${e.category}</td>
        <td>${e.quantity} ${e.unit}</td>
        <td>${e.factor}</td>
        <td><strong>${e.co2kg}</strong></td>
        <td style="font-size:11px;color:#6b7280">${e.source}</td>
      </tr>`).join('')
    : '';

  const catRows = (data.results || []).map(r => `
    <tr>
      <td><strong>${r.category}</strong><br><span style="font-size:11px;color:#6b7280">${r.description}</span></td>
      <td style="color:${r.status === 'above' ? '#16a34a' : r.status === 'below' ? '#dc2626' : '#d97706'};font-weight:700">${r.score}</td>
      <td>${r.benchmark}</td>
      <td style="color:${r.gap >= 0 ? '#16a34a' : '#dc2626'}">${r.gap > 0 ? '+' : ''}${r.gap}</td>
      <td><span style="padding:2px 8px;border-radius:4px;font-size:11px;background:${r.status === 'above' ? '#dcfce7' : r.status === 'below' ? '#fee2e2' : '#fef9c3'};color:${r.status === 'above' ? '#15803d' : r.status === 'below' ? '#b91c1c' : '#a16207'}">${r.status}</span></td>
      <td style="font-size:12px;color:#374151">${r.insight}</td>
    </tr>`).join('');

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Carbon Canary Report - ${data.fileName}</title>
<style>
  body { font-family: 'Segoe UI', sans-serif; margin: 0; padding: 40px; background: #f9fafb; color: #111827; }
  .header { background: #0a0f0d; color: white; padding: 32px 40px; border-radius: 12px; margin-bottom: 32px; }
  .logo { color: #4ade80; font-size: 22px; font-weight: 700; margin-bottom: 8px; }
  .title { font-size: 28px; font-weight: 700; margin-bottom: 4px; }
  .meta { color: #6b7280; font-size: 13px; }
  .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; margin-bottom: 32px; }
  .kpi { background: white; border: 1px solid #e5e7eb; border-radius: 10px; padding: 20px; }
  .kpi-label { font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: #6b7280; margin-bottom: 8px; }
  .kpi-value { font-size: 32px; font-weight: 700; color: #111827; }
  .kpi-sub { font-size: 12px; color: #6b7280; margin-top: 4px; }
  .section { background: white; border: 1px solid #e5e7eb; border-radius: 10px; padding: 24px; margin-bottom: 24px; }
  .section-title { font-size: 13px; text-transform: uppercase; letter-spacing: 2px; color: #6b7280; margin-bottom: 16px; border-bottom: 1px solid #e5e7eb; padding-bottom: 8px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th { text-align: left; padding: 8px 12px; border-bottom: 2px solid #e5e7eb; font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: #6b7280; }
  td { padding: 10px 12px; border-bottom: 1px solid #f3f4f6; vertical-align: top; }
  .total-row td { font-weight: 700; font-size: 15px; border-top: 2px solid #e5e7eb; background: #f9fafb; }
  .co2-big { font-size: 48px; font-weight: 700; color: ${data.co2Color === 'above' ? '#16a34a' : data.co2Color === 'below' ? '#dc2626' : '#d97706'}; }
  .note { background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 12px 16px; font-size: 12px; color: #15803d; margin-bottom: 24px; }
  .footer { text-align: center; color: #9ca3af; font-size: 12px; margin-top: 40px; }
  @media print { body { padding: 20px; } }
</style>
</head>
<body>
<div class="header">
  <div class="logo">🐦 Carbon Canary</div>
  <div class="title">${isInvoice ? 'Invoice CO₂ Emission Report' : 'CSRD Sustainability Report Scoring'}</div>
  <div class="meta">📄 ${data.fileName} · Generated ${date} · Climathon 2026</div>
</div>

${isInvoice ? `
<div class="kpi-grid">
  <div class="kpi" style="grid-column:span 2">
    <div class="kpi-label">Total CO₂ Emissions</div>
    <div class="co2-big">${data.totalCO2kg} <span style="font-size:20px;color:#6b7280">kg</span></div>
    <div class="kpi-sub">${data.totalCO2t} tonnes CO₂eq · Rating: <strong>${data.co2Rating}</strong></div>
  </div>
  <div class="kpi"><div class="kpi-label">Supplier</div><div style="font-size:16px;font-weight:600">${data.supplierName || 'Unknown'}</div></div>
  <div class="kpi"><div class="kpi-label">Invoice</div><div style="font-size:16px;font-weight:600">${data.invoiceNumber || 'Unknown'}</div></div>
  <div class="kpi"><div class="kpi-label">Date</div><div style="font-size:16px;font-weight:600">${data.invoiceDate || 'Unknown'}</div></div>
  <div class="kpi"><div class="kpi-label">Confidence</div><div style="font-size:16px;font-weight:600;color:${data.confidence === 'high' ? '#16a34a' : data.confidence === 'medium' ? '#d97706' : '#dc2626'}">${(data.confidence || '').toUpperCase()}</div></div>
</div>
${data.notes ? `<div class="note">🤖 ${data.notes}</div>` : ''}
<div class="section">
  <div class="section-title">Emission Breakdown by Line Item</div>
  <table>
    <thead><tr><th>Item</th><th>Category</th><th>Quantity</th><th>EU Factor</th><th>CO₂ (kg)</th><th>Source</th></tr></thead>
    <tbody>${rowsHTML}</tbody>
    <tfoot><tr class="total-row"><td colspan="4">TOTAL CO₂ EMISSIONS</td><td>${data.totalCO2kg} kg</td><td></td></tr></tfoot>
  </table>
</div>` : `
<div class="kpi-grid">
  <div class="kpi"><div class="kpi-label">Overall Score</div><div class="kpi-value">${data.overallScore}</div><div class="kpi-sub">Benchmark: ${data.overallBenchmark}/100</div></div>
  <div class="kpi"><div class="kpi-label">Strong Areas</div><div class="kpi-value" style="color:#16a34a">${data.strong.length}</div></div>
  <div class="kpi"><div class="kpi-label">Weak Areas</div><div class="kpi-value" style="color:#dc2626">${data.weak.length}</div></div>
  <div class="kpi"><div class="kpi-label">vs Benchmark</div><div class="kpi-value" style="color:${data.overallScore >= data.overallBenchmark ? '#16a34a' : '#dc2626'}">${data.overallScore - data.overallBenchmark > 0 ? '+' : ''}${data.overallScore - data.overallBenchmark}</div></div>
</div>`}

<div class="section">
  <div class="section-title">Category Performance vs Benchmark</div>
  <table>
    <thead><tr><th>Category</th><th>Score</th><th>Benchmark</th><th>Gap</th><th>Status</th><th>Insight</th></tr></thead>
    <tbody>${catRows}</tbody>
  </table>
</div>

<div class="footer">
  Carbon Canary · Scope 3 Emission Analytics · DE/PL Corridor · Climathon 2026<br>
  Data sources: EU JRC Emission Factor Database 2023, IPCC AR6
</div>
</body>
</html>`;

  const blob = new Blob([html], { type: 'text/html' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `carbon-canary-report-${data.fileName.replace(/\.[^.]+$/, '')}.html`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── DOWNLOAD BUTTONS ──────────────────────────────────────────────────────────
function DownloadButtons({ data }) {
  return (
    <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
      <button onClick={() => downloadHTML(data)} style={{
        padding: '10px 20px', background: '#4ade80', color: '#050a06', border: 'none',
        borderRadius: 8, fontWeight: 700, fontSize: 13, cursor: 'pointer', fontFamily: 'inherit'
      }}>⬇ Download HTML Report</button>
      <button onClick={() => downloadCSV(data)} style={{
        padding: '10px 20px', background: 'transparent', color: '#4ade80',
        border: '1px solid #4ade80', borderRadius: 8, fontWeight: 600, fontSize: 13,
        cursor: 'pointer', fontFamily: 'inherit'
      }}>⬇ Download CSV Data</button>
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
  const { fileName, totalCO2kg, totalCO2t, co2Rating, co2Color, emissionBreakdown, results, strong, weak, overallScore, supplierName, invoiceNumber, invoiceDate, confidence, notes, competitorView, recommendations } = data;
  const barData = emissionBreakdown.map(e => ({ name: e.item.length > 14 ? e.item.slice(0, 14) + '…' : e.item, co2: e.co2kg }));
  const radarData = results.map(r => ({ category: r.category.split(' ')[0], Score: r.score, Benchmark: r.benchmark }));
  const confColor = confidence === 'high' ? '#4ade80' : confidence === 'medium' ? '#fbbf24' : '#f87171';

  return (
    <div className="dashboard">
      <div className="dash-header">
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <div className="dash-title">Invoice CO₂ Emission Report</div>
            <span style={{ background: 'rgba(74,222,128,0.12)', color: '#4ade80', padding: '2px 10px', borderRadius: 4, fontSize: 11, fontFamily: 'IBM Plex Mono' }}>INVOICE</span>
            <span style={{ background: `rgba(${confidence === 'high' ? '74,222,128' : confidence === 'medium' ? '251,191,36' : '248,113,113'},0.12)`, color: confColor, padding: '2px 10px', borderRadius: 4, fontSize: 11, fontFamily: 'IBM Plex Mono' }}>{(confidence || '').toUpperCase()} CONFIDENCE</span>
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
          <div className="panel-title">Radar Overview</div>
          <ResponsiveContainer width="100%" height={260}>
            <RadarChart data={radarData}>
              <PolarGrid stroke="#1e2b1e" />
              <PolarAngleAxis dataKey="category" tick={{ fill: '#6b8f72', fontSize: 10, fontFamily: 'IBM Plex Mono' }} />
              <PolarRadiusAxis domain={[0, 100]} tick={false} axisLine={false} />
              <Radar name="Score" dataKey="Score" stroke="#4ade80" fill="#4ade80" fillOpacity={0.15} />
              <Radar name="Benchmark" dataKey="Benchmark" stroke="#fbbf24" fill="#fbbf24" fillOpacity={0.08} strokeDasharray="4 2" />
              <Legend wrapperStyle={{ fontSize: 11, fontFamily: 'IBM Plex Mono', color: '#6b8f72' }} />
            </RadarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="panel full" style={{ marginBottom: 20 }}>
        <div className="panel-title">Detailed Emission Breakdown</div>
        <div style={{ overflowX: 'auto' }}>
          <table className="cat-table">
            <thead><tr><th>Item</th><th>Category</th><th>Quantity</th><th>EU Factor (kg/unit)</th><th style={{ textAlign: 'right' }}>CO₂ (kg)</th><th>Source</th></tr></thead>
            <tbody>
              {emissionBreakdown.map((e, i) => (
                <tr key={i}>
                  <td><div className="cat-name" style={{ fontSize: 12 }}>{e.item}{e.estimated && <span style={{ color: '#fbbf24', fontSize: 10, marginLeft: 6 }}>⚠ estimated</span>}</div></td>
                  <td style={{ fontSize: 12, color: '#6b8f72' }}>{e.category}</td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 12 }}>{e.quantity} {e.unit}</td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 12, color: '#6b8f72' }}>{e.factor}</td>
                  <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 14, fontWeight: 700, color: '#4ade80', textAlign: 'right' }}>{e.co2kg}</td>
                  <td style={{ fontSize: 10, color: '#6b8f72' }}>{e.source}</td>
                </tr>
              ))}
              <tr style={{ borderTop: '2px solid #1e2b1e' }}>
                <td colSpan={4} style={{ fontWeight: 700, fontSize: 14 }}>TOTAL</td>
                <td style={{ fontFamily: 'IBM Plex Mono', fontSize: 18, fontWeight: 700, color: co2Color === 'above' ? '#4ade80' : co2Color === 'below' ? '#f87171' : '#fbbf24', textAlign: 'right' }}>{totalCO2kg} kg</td>
                <td />
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
          <div className="panel-title">Radar Overview</div>
          <ResponsiveContainer width="100%" height={280}>
            <RadarChart data={radarData}>
              <PolarGrid stroke="#1e2b1e" />
              <PolarAngleAxis dataKey="category" tick={{ fill: '#6b8f72', fontSize: 10, fontFamily: 'IBM Plex Mono' }} />
              <PolarRadiusAxis domain={[0, 100]} tick={false} axisLine={false} />
              <Radar name="Score" dataKey="Score" stroke="#4ade80" fill="#4ade80" fillOpacity={0.15} />
              <Radar name="Benchmark" dataKey="Benchmark" stroke="#fbbf24" fill="#fbbf24" fillOpacity={0.08} strokeDasharray="4 2" />
              <Legend wrapperStyle={{ fontSize: 11, fontFamily: 'IBM Plex Mono', color: '#6b8f72' }} />
            </RadarChart>
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
