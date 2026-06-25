# INFINITE — Facility Assessment Snapshot
### Managed by MEDELITE

A web application for evaluating skilled nursing facilities before partnership or outreach discussions. Built for the Medelite Director workflow — enter any CMS CCN to instantly pull star ratings, hospitalization metrics, and generate polished PDF and Word reports.

---

## Live Features

- **Dynamic CCN Lookup** — Enter any 6-digit CMS Certification Number; CCN `686123` loads instantly from verified reference data, all others query Medicare Care Compare via AI web search
- **All CMS Star Ratings** — Overall, Health Inspection, Staffing, Quality of Resident Care (1–5 stars)
- **All 12 Hospitalization & ED Metrics** — Short-stay and long-stay metrics with facility vs. national vs. state benchmark charts
- **Manual Operational Inputs** — EMR, Current Census, Patient Type, Medelite Coverage History, Medical Coverage
- **Download as PDF** — Downloads a print-ready HTML file; open in browser and click "Print / Save as PDF"
- **Download as DOCX** — Generates a fully editable Word document (.docx) in the browser with no server required; includes a clickable Medicare.gov hyperlink
- **Assessment Summary** — AI-generated flags (✅ ok / ⚠️ warn / 🔴 critical / ℹ️ info) based on real metric analysis

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| UI Framework | React 18 + Vite |
| Charts | Recharts (radar + bar) |
| DOCX Generation | Pure browser JavaScript (ZIP + OOXML, no npm) |
| PDF Export | HTML download → browser print |
| CMS Data | Anthropic API (`claude-sonnet-4-6`) + web_search tool |
| Hosting | GitHub Pages / Netlify |

---

## Deploy in 3 Minutes

### Option A — Netlify Drop (30 seconds, no account needed)
1. Unzip `medelite-deploy.zip`
2. Go to [app.netlify.com/drop](https://app.netlify.com/drop)
3. Drag the `dist/` folder onto the page
4. Get a live URL instantly

### Option B — GitHub Pages

```bash
# 1. Unzip and enter the project
unzip medelite-deploy.zip
cd medelite-deploy

# 2. Install dependencies
npm install

# 3. Create a GitHub repo at github.com/new
#    Name it: medelite-facility-assessment
#    Set to Public

# 4. Push to GitHub
git remote add origin https://github.com/YOUR_USERNAME/medelite-facility-assessment.git
git push -u origin master

# 5. Deploy to GitHub Pages (builds + pushes gh-pages branch)
npm run deploy
```

**6.** Go to your repo → Settings → Pages → Source: `gh-pages` branch → Save

**Live URL:** `https://YOUR_USERNAME.github.io/medelite-facility-assessment/`

---

## Test Case

| Field | Value |
|-------|-------|
| **CCN** | `686123` |
| **Facility** | Kendall Lakes Healthcare and Rehab Center |
| **Location** | 5280 SW 157th Ave, Miami, FL 33193 |
| **Overall Star** | 1 |
| **STR Hospitalization** | 18.7% (below national avg of 21.5%) |
| **LT ED Visits** | 6.94/1,000 days (4.2× national avg) |

Medicare source: [care-compare/details/nursing-home/686123](https://www.medicare.gov/care-compare/details/nursing-home/686123)

---

## CMS Data Dictionary Field Mappings

Per `NH_Data_Dictionary` (Feb 2026):

| Report Label | CMS Field Name | Data File |
|-------------|---------------|-----------|
| Census Capacity | Number of Certified Beds | NH_ProviderInfo |
| Overall Star Rating | Overall Rating | NH_ProviderInfo |
| Health Inspection | Health Inspection Rating | NH_ProviderInfo |
| Staffing | Staffing Rating | NH_ProviderInfo |
| Quality of Resident Care | QM Rating | NH_ProviderInfo |
| Short Term Hospitalization | % short-stay rehospitalized (Claims QM) | NH_QualityMsr_Claims |
| STR ED Visit | % short-stay ED visit | NH_QualityMsr_Claims |
| LT Hospitalization | Hospitalizations per 1000 long-stay resident days | NH_QualityMsr_Claims |
| ED Visit (LT) | ED visits per 1000 long-stay resident days | NH_QualityMsr_Claims |

---

## Project Structure

```
medelite-deploy/
├── src/
│   └── App.jsx          # Complete application (single file)
├── dist/                # Production build (deploy this folder)
├── index.html
├── package.json
├── vite.config.js
└── README.md
```

---

## Download Note

**PDF:** Clicking "Download as PDF" saves a `.html` file. Open it in your browser and press **Ctrl+P** (or Cmd+P on Mac), then choose **"Save as PDF"** as the destination. The layout is Letter-sized and print-ready.

**DOCX:** Clicking "Download as DOCX" downloads a fully editable `.docx` file directly. The Medicare.gov URL in the footer is a live hyperlink in Word.

---

*Built by Medelite Engineering · Data sourced from CMS Provider Data Catalog · Updated Feb 2026*
>>>>>>> 7555dfb (Initial commit)
