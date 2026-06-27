# 🐦 Carbon Canary — Scope 3 Analytics Dashboard

**Automated Scope 3 Carbon Emission Reporting for Cross-Border DE/PL Manufacturers**  
Built for Climathon 2026 | Lusatia-Silesia Region

---

## System Architecture

```
carbon-canary/
├── backend/          ← Node.js + Express API (port 3001)
│   └── server.js     ← Upload, extract, score, benchmark
├── frontend/         ← React dashboard (port 3000)
│   └── src/
│       ├── App.js    ← Full dashboard UI
│       └── index.css ← Design system
└── start.sh          ← One-command launcher
```

---

## Quick Start

### Prerequisites
- Node.js 16+ and npm

### Step 1 — Install Dependencies

```bash
# Backend
cd backend && npm install

# Frontend
cd ../frontend && npm install
```

### Step 2 — Start Everything

```bash
# From the carbon-canary/ root:
bash start.sh
```

Or start manually in two terminals:

```bash
# Terminal 1 — Backend
cd backend && node server.js

# Terminal 2 — Frontend
cd frontend && npm start
```

### Step 3 — Open Dashboard

Visit **http://localhost:3000**

---

## Features

| Module | What it does |
|--------|-------------|
| **Upload** | PDF, Excel (.xlsx/.xls), CSV, TXT files up to 20MB |
| **Extraction** | LLM keyword-based extraction of emission-relevant data |
| **Scoring Engine** | 10 dynamic Scope 3 categories, scored 0–100 |
| **Benchmarking** | EU/industry benchmark comparison per category |
| **Gap Analysis** | Variance from benchmark with above/average/below labels |
| **Visualizations** | Bar chart, Radar chart, Gap heatmap |
| **Insights** | Auto-generated performance insights per category |

---

## Scoring Categories

1. Scope 3 Coverage
2. Data Completeness
3. Supplier Engagement
4. Reporting Frequency
5. Emission Intensity
6. Reduction Target
7. Audit Readiness
8. Transport Emissions
9. Material Sourcing
10. Waste & Circularity

---

## API Endpoints

| Method | Route | Description |
|--------|-------|-------------|
| POST | `/api/upload` | Upload and analyze a file |
| GET | `/api/benchmarks` | Get all benchmark values |
| GET | `/api/health` | Health check |

### Example Upload Response

```json
{
  "fileName": "supplier_invoice.pdf",
  "overallScore": 68,
  "overallBenchmark": 65,
  "results": [
    {
      "category": "Scope 3 Coverage",
      "score": 75,
      "benchmark": 80,
      "gap": -5,
      "status": "average",
      "insight": "Meeting standard — within benchmark range..."
    }
  ],
  "strong": ["Audit Readiness", "Transport Emissions"],
  "weak": ["Waste & Circularity"]
}
```

---

## Test with Carbon Canary PDF

The included `Carbon_Canary_Project.pdf` scores well on Audit Readiness, 
Transport Emissions, and Scope 3 Coverage — try uploading it to see the 
dashboard in action.

---

*Climathon 2026 · DE/PL Corridor · Prototype*
