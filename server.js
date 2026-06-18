const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
app.use(cors());
app.use(express.json());

const upload = multer({ dest: 'uploads/', limits: { fileSize: 20 * 1024 * 1024 } });
if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');

// ─── EU EMISSION FACTORS (kg CO2eq per unit) ──────────────────────────────────
const MATERIAL_FACTORS = [
  { keywords: ['steel','stahl','stal'], label: 'Steel', factor: 1.85, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['iron','eisen','żelazo'], label: 'Iron', factor: 1.91, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['alumin'], label: 'Aluminium', factor: 8.24, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['copper','kupfer','miedź'], label: 'Copper', factor: 3.49, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['plastic','kunststoff','tworzywo'], label: 'Plastic', factor: 3.14, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['rubber','gummi','guma'], label: 'Rubber', factor: 3.18, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['glass','glas','szkło','szklo'], label: 'Glass', factor: 0.85, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['paper','papier','papier'], label: 'Paper', factor: 1.29, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['cardboard','karton','karton'], label: 'Cardboard', factor: 0.89, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['wood','holz','drewno'], label: 'Wood', factor: 0.46, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['cement','beton','concrete'], label: 'Cement/Concrete', factor: 0.83, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['chemical','chemik','chemii','substancj','kwas','acid','solvent','lösungsmittel'], label: 'Chemicals', factor: 2.10, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['fertilizer','dünger','nawóz'], label: 'Fertilizer', factor: 4.20, unit: 'kg', source: 'IPCC AR6' },
  { keywords: ['textile','textil','tkanina','fabric','stoff'], label: 'Textile', factor: 15.0, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['diesel','kraftstoff'], label: 'Diesel Fuel', factor: 2.68, unit: 'litre', source: 'EU JRC 2023' },
  { keywords: ['petrol','benzin','benzyna'], label: 'Petrol', factor: 2.31, unit: 'litre', source: 'EU JRC 2023' },
  { keywords: ['electricity','strom','elektryczność','prąd','energia elektryczna'], label: 'Electricity', factor: 0.276, unit: 'kWh', source: 'EU avg 2023' },
  { keywords: ['transport','truck','lkw','freight','spedition','logistics','logistik','przewóz','dostawa'], label: 'Road Transport', factor: 0.096, unit: 'tkm', source: 'EU JRC 2023' },
  { keywords: ['train','bahn','rail','kolej'], label: 'Rail Transport', factor: 0.028, unit: 'tkm', source: 'EU JRC 2023' },
  { keywords: ['air freight','luftfracht','lotniczy'], label: 'Air Freight', factor: 0.602, unit: 'tkm', source: 'IPCC AR6' },
  { keywords: ['oil','öl','olej'], label: 'Oil/Lubricant', factor: 3.15, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['paint','farbe','farba','coating'], label: 'Paint/Coating', factor: 2.80, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['packaging','verpackung','opakowanie'], label: 'Packaging', factor: 1.10, unit: 'kg', source: 'EU JRC 2023' },
  { keywords: ['food','lebensmittel','żywność','meat','fleisch'], label: 'Food Products', factor: 5.50, unit: 'kg', source: 'IPCC AR6' },
  { keywords: ['machinery','maschine','maszyna','equipment'], label: 'Machinery', factor: 3.20, unit: 'kg', source: 'EU JRC 2023' },
];

const BENCHMARKS = {
  'Material Emissions':   { benchmark: 60, description: 'CO₂ intensity of materials in this invoice' },
  'Transport Emissions':  { benchmark: 68, description: 'Logistics and delivery carbon footprint' },
  'Data Completeness':    { benchmark: 75, description: 'Quality and completeness of emission data' },
  'Supplier Coverage':    { benchmark: 60, description: 'Number of emission sources identified' },
  'Emission Intensity':   { benchmark: 65, description: 'CO₂ per unit value of goods' },
  'Audit Readiness':      { benchmark: 72, description: 'CSRD compliance readiness of this document' },
};

const REPORT_BENCHMARKS = {
  'Scope 3 Coverage':    { benchmark: 80, description: 'Percentage of Scope 3 categories tracked' },
  'Data Completeness':   { benchmark: 75, description: 'Completeness of emission data fields' },
  'Supplier Engagement': { benchmark: 60, description: 'Suppliers with verified emission data' },
  'Reporting Frequency': { benchmark: 70, description: 'Regularity and timeliness of reporting' },
  'Emission Intensity':  { benchmark: 65, description: 'CO2eq per unit of production/revenue' },
  'Reduction Target':    { benchmark: 50, description: 'Progress toward stated reduction goals' },
  'Audit Readiness':     { benchmark: 72, description: 'Readiness for CSRD/third-party audit' },
  'Transport Emissions': { benchmark: 68, description: 'Logistics and transport carbon tracking' },
  'Material Sourcing':   { benchmark: 62, description: 'Sustainability of raw material sourcing' },
  'Waste & Circularity': { benchmark: 45, description: 'Waste reduction and circular practices' },
};

// ─── PDF TEXT EXTRACTION ──────────────────────────────────────────────────────
function extractTextFromPDF(buffer) {
  try {
    const str = buffer.toString('latin1');
    const blocks = [];
    const btRegex = /BT([\s\S]*?)ET/g;
    let m;
    while ((m = btRegex.exec(str)) !== null) {
      const b = m[1];
      const tj = /\((.*?)\)\s*Tj/g;
      let t;
      while ((t = tj.exec(b)) !== null) blocks.push(t[1]);
      const hex = /<([0-9A-Fa-f]+)>\s*Tj/g;
      let h;
      while ((h = hex.exec(b)) !== null) {
        let d = '';
        for (let i = 0; i < h[1].length; i += 2)
          d += String.fromCharCode(parseInt(h[1].substr(i, 2), 16));
        blocks.push(d);
      }
    }
    const readable = str.replace(/[^\x20-\x7E\n\r]/g, ' ').replace(/\s+/g, ' ');
    return (blocks.join(' ') + ' ' + readable).slice(0, 10000);
  } catch (e) { return ''; }
}

// ─── SMART NUMBER EXTRACTION ──────────────────────────────────────────────────
function extractNumbers(text) {
  const numbers = [];
  // Match patterns like: 100 kg, 50,5 kg, 1.500 t, 200 litre, 50 kWh
  const regex = /(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)\s*(kg|t\b|ton|tonne|litre|liter|ltr|l\b|kwh|km|pcs|szt|stk|stück|m²|m2|m³|m3)/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    let val = match[1].replace(/\./g, '').replace(',', '.');
    val = parseFloat(val);
    const unit = match[2].toLowerCase();
    // Convert tonnes to kg
    let valueInBase = val;
    let baseUnit = unit;
    if (unit === 't' || unit === 'ton' || unit === 'tonne') { valueInBase = val * 1000; baseUnit = 'kg'; }
    if (!isNaN(val) && val > 0 && val < 1000000) {
      numbers.push({ raw: match[0], value: valueInBase, unit: baseUnit, position: match.index });
    }
  }
  return numbers;
}

// ─── EXTRACT LINE ITEMS FROM TEXT ─────────────────────────────────────────────
function extractLineItems(text) {
  const lower = text.toLowerCase();
  const numbers = extractNumbers(text);
  const lineItems = [];
  const usedFactors = new Set();

  for (const mf of MATERIAL_FACTORS) {
    for (const kw of mf.keywords) {
      if (lower.includes(kw) && !usedFactors.has(mf.label)) {
        // Find the closest quantity to this keyword occurrence
        const kwIdx = lower.indexOf(kw);
        let bestNum = null;
        let bestDist = 999999;

        for (const num of numbers) {
          const dist = Math.abs(num.position - kwIdx);
          if (dist < bestDist && dist < 500) { // within 500 chars
            bestDist = dist;
            bestNum = num;
          }
        }

        // Use found quantity or make a context-based estimate
        let quantity, unit, estimated;
        if (bestNum) {
          quantity = bestNum.value;
          unit = bestNum.unit;
          estimated = false;
        } else {
          // Estimate based on typical invoice values per material
          const estimates = {
            'Steel': 500, 'Iron': 300, 'Aluminium': 100, 'Copper': 50,
            'Plastic': 200, 'Rubber': 100, 'Glass': 150, 'Paper': 80,
            'Cardboard': 60, 'Wood': 200, 'Cement/Concrete': 1000,
            'Chemicals': 150, 'Fertilizer': 500, 'Textile': 50,
            'Diesel Fuel': 100, 'Petrol': 80, 'Electricity': 500,
            'Road Transport': 100, 'Rail Transport': 200, 'Air Freight': 50,
            'Oil/Lubricant': 20, 'Paint/Coating': 30, 'Packaging': 40,
            'Food Products': 100, 'Machinery': 200,
          };
          quantity = estimates[mf.label] || 100;
          unit = mf.unit;
          estimated = true;
        }

        const co2kg = Math.round(quantity * mf.factor * 100) / 100;
        lineItems.push({
          item: kw.charAt(0).toUpperCase() + kw.slice(1),
          category: mf.label,
          quantity,
          unit,
          factor: mf.factor,
          co2kg,
          estimated,
          source: mf.source,
        });
        usedFactors.add(mf.label);
        break;
      }
    }
  }

  // If nothing found, add generic goods
  if (lineItems.length === 0) {
    const anyNumber = numbers[0];
    lineItems.push({
      item: 'General Goods',
      category: 'General',
      quantity: anyNumber ? anyNumber.value : 100,
      unit: anyNumber ? anyNumber.unit : 'kg',
      factor: 1.5,
      co2kg: Math.round((anyNumber ? anyNumber.value : 100) * 1.5 * 100) / 100,
      estimated: true,
      source: 'IPCC AR6 Default',
    });
  }

  return lineItems;
}

// ─── EXTRACT INVOICE METADATA ─────────────────────────────────────────────────
function extractMetadata(text) {
  const lower = text.toLowerCase();

  // Supplier name — look for common patterns
  let supplierName = 'Unknown Supplier';
  const supplierPatterns = [
    /(?:from|von|od|supplier|lieferant|sprzedawca)[:\s]+([A-Z][A-Za-z\s&.,]+?)(?:\n|ltd|gmbh|sp\.z|s\.a|inc)/i,
    /^([A-Z][A-Za-z\s&.,]{3,40}(?:GmbH|Ltd|S\.A\.|Sp\.z\.o\.o\.|Inc|AG))/m,
  ];
  for (const p of supplierPatterns) {
    const m = text.match(p);
    if (m) { supplierName = m[1].trim(); break; }
  }

  // Invoice number
  let invoiceNumber = 'Unknown';
  const invPatterns = [/(?:invoice|rechnung|faktura|nr\.?|no\.?)[:\s#]*([A-Z0-9\-\/]{3,20})/i];
  for (const p of invPatterns) {
    const m = text.match(p);
    if (m) { invoiceNumber = m[1].trim(); break; }
  }

  // Date
  let invoiceDate = 'Unknown';
  const datePatterns = [/(\d{1,2}[.\/\-]\d{1,2}[.\/\-]\d{2,4})/];
  for (const p of datePatterns) {
    const m = text.match(p);
    if (m) { invoiceDate = m[1]; break; }
  }

  // Currency
  let currency = 'EUR';
  if (lower.includes('pln') || lower.includes('zł') || lower.includes('zloty')) currency = 'PLN';
  else if (lower.includes('usd') || lower.includes('$')) currency = 'USD';
  else if (lower.includes('gbp') || lower.includes('£')) currency = 'GBP';

  return { supplierName, invoiceNumber, invoiceDate, currency };
}

// ─── BUILD CATEGORY SCORES ────────────────────────────────────────────────────
function buildCategoryScores(lineItems, totalCO2kg, text) {
  const lower = text.toLowerCase();
  const hasTransport = lineItems.some(i => i.category.includes('Transport'));
  const itemCount = lineItems.length;
  const hasRealQty = lineItems.filter(i => !i.estimated).length;

  const scores = {
    'Material Emissions': Math.min(100, Math.max(10, 100 - Math.round(totalCO2kg / 100))),
    'Transport Emissions': hasTransport ? 55 : (lower.includes('delivery') || lower.includes('lieferung') || lower.includes('dostawa') ? 40 : 25),
    'Data Completeness': Math.min(90, 30 + hasRealQty * 15 + itemCount * 5),
    'Supplier Coverage': Math.min(90, 20 + itemCount * 12),
    'Emission Intensity': Math.min(100, Math.max(10, 90 - Math.round(totalCO2kg / 80))),
    'Audit Readiness': Math.min(85, 35 + itemCount * 8 + (hasRealQty > 0 ? 15 : 0)),
  };

  return Object.entries(scores).map(([category, score]) => {
    const bm = BENCHMARKS[category];
    const gap = score - bm.benchmark;
    const status = gap >= 10 ? 'above' : gap <= -10 ? 'below' : 'average';
    return {
      category, score, benchmark: bm.benchmark,
      description: bm.description, gap, status,
      insight: gap >= 10 ? `Good — ${Math.abs(gap)} pts above benchmark.`
        : gap >= 0 ? `Within benchmark range. Minor improvements possible.`
        : `Below benchmark by ${Math.abs(gap)} pts. Action recommended.`,
    };
  });
}

// ─── DETECT DOCUMENT TYPE ─────────────────────────────────────────────────────
function detectDocumentType(text) {
  const lower = text.toLowerCase();
  let inv = 0, rep = 0;
  ['invoice','rechnung','faktura','total','quantity','menge','price','preis','vat','mwst','netto','brutto'].forEach(k => { if (lower.includes(k)) inv++; });
  ['scope 3','csrd','sustainability report','emission factor','net zero','carbon neutral','ghg inventory'].forEach(k => { if (lower.includes(k)) rep++; });
  return inv >= rep ? 'invoice' : 'report';
}

// ─── REPORT SCORING ───────────────────────────────────────────────────────────
function analyzeReport(text) {
  const lower = text.toLowerCase();
  const rules = {
    'Scope 3 Coverage': [['scope 3','upstream','downstream','value chain'],['category 1','category 4','purchased goods'],['all categories','full scope 3','comprehensive']],
    'Data Completeness': [['invoice','document','data extraction'],['confirmed','verified','validated'],['complete data','full records']],
    'Supplier Engagement': [['supplier','vendor','procurement'],['supplier portal','questionnaire'],['zero supplier friction','passive','automatic']],
    'Reporting Frequency': [['annual','yearly','report'],['quarterly','monthly','regular'],['real-time','continuous','automated']],
    'Emission Intensity': [['co2','carbon','ghg','emission factor'],['co2eq','tco2','scope'],['intensity','per unit']],
    'Reduction Target': [['reduction','target','goal','commitment'],['net zero','carbon neutral'],['sbti','science-based','paris']],
    'Audit Readiness': [['csrd','audit','compliance'],['cited','source','traceable','transparent'],['eu jrc','ipcc','emission factor database']],
    'Transport Emissions': [['transport','logistics','freight','delivery'],['shipping','distance','mode'],['last mile','carrier']],
    'Material Sourcing': [['material','raw material','sourcing'],['sustainable','recycled','low-carbon'],['supply chain','goods']],
    'Waste & Circularity': [['waste','circular','recycling'],['landfill','end-of-life'],['zero waste','circularity']],
  };

  const results = Object.entries(rules).map(([category, ruleset]) => {
    let score = 30;
    for (const kwList of ruleset)
      for (const kw of kwList)
        if (lower.includes(kw)) { score += 20; break; }
    score = Math.min(100, score);
    const bm = REPORT_BENCHMARKS[category];
    const gap = score - bm.benchmark;
    const status = gap >= 10 ? 'above' : gap <= -10 ? 'below' : 'average';
    return { category, score, benchmark: bm.benchmark, description: bm.description, gap, status,
      insight: gap >= 15 ? `Strong — ${Math.abs(gap)} pts above benchmark.`
        : gap >= 0 ? `Meeting standard.`
        : gap >= -15 ? `Slightly below by ${Math.abs(gap)} pts.`
        : `Significant gap of ${Math.abs(gap)} pts. Prioritize this.` };
  });

  return {
    mode: 'report',
    overallScore: Math.round(results.reduce((s, r) => s + r.score, 0) / results.length),
    overallBenchmark: 65,
    results,
    strong: results.filter(r => r.status === 'above').map(r => r.category),
    weak: results.filter(r => r.status === 'below').map(r => r.category),
  };
}

// ─── UPLOAD ROUTE ─────────────────────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const filePath = req.file.path;
  const ext = path.extname(req.file.originalname).toLowerCase();
  let text = '';

  try {
    if (ext === '.pdf') {
      text = extractTextFromPDF(fs.readFileSync(filePath));
    } else if (ext === '.xlsx' || ext === '.xls') {
      const wb = xlsx.readFile(filePath);
      wb.SheetNames.forEach(n => { text += xlsx.utils.sheet_to_csv(wb.Sheets[n]) + '\n'; });
    } else if (ext === '.csv' || ext === '.txt') {
      text = fs.readFileSync(filePath, 'utf-8');
    } else {
      fs.unlinkSync(filePath);
      return res.status(400).json({ error: 'Unsupported file type.' });
    }
    fs.unlinkSync(filePath);

    const docType = detectDocumentType(text);

    if (docType === 'report') {
      const analysis = analyzeReport(text);
      return res.json({ fileName: req.file.originalname, documentType: 'report', wordCount: text.split(/\s+/).filter(Boolean).length, ...analysis });
    }

    // INVOICE MODE
    const metadata = extractMetadata(text);
    const lineItems = extractLineItems(text);
    const totalCO2kg = Math.round(lineItems.reduce((s, i) => s + i.co2kg, 0) * 100) / 100;
    const totalCO2t = Math.round(totalCO2kg / 10) / 100;
    const co2Rating = totalCO2kg < 500 ? 'Low' : totalCO2kg < 2000 ? 'Medium' : 'High';
    const co2Color = totalCO2kg < 500 ? 'above' : totalCO2kg < 2000 ? 'average' : 'below';
    const results = buildCategoryScores(lineItems, totalCO2kg, text);
    const overallScore = Math.round(results.reduce((s, r) => s + r.score, 0) / results.length);
    const estimatedCount = lineItems.filter(i => i.estimated).length;
    const confidence = estimatedCount === 0 ? 'high' : estimatedCount < lineItems.length / 2 ? 'medium' : 'low';

    res.json({
      fileName: req.file.originalname,
      documentType: 'invoice',
      wordCount: text.split(/\s+/).filter(Boolean).length,
      ...metadata,
      totalCO2kg,
      totalCO2t,
      co2Rating,
      co2Color,
      confidence,
      notes: estimatedCount > 0
        ? `${estimatedCount} of ${lineItems.length} quantities were estimated because exact values were not found in the document. Upload a structured CSV or Excel for more accurate results.`
        : `All quantities extracted directly from the document.`,
      emissionBreakdown: lineItems,
      overallScore,
      overallBenchmark: 65,
      results,
      strong: results.filter(r => r.status === 'above').map(r => r.category),
      weak: results.filter(r => r.status === 'below').map(r => r.category),
    });
  } catch (err) {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    console.error(err);
    res.status(500).json({ error: 'Failed to process: ' + err.message });
  }
});

app.get('/api/benchmarks', (req, res) => res.json({ ...BENCHMARKS, ...REPORT_BENCHMARKS }));
app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

const PORT = 3001;
app.listen(PORT, () => console.log(`Carbon Canary backend running on http://localhost:${PORT}`));
