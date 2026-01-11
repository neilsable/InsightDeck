InsightDeck is a lightweight web app that turns structured operational data into an **executive-ready PowerPoint deck** in seconds.

Upload a CSV → InsightDeck automatically:
- computes KPIs
- generates a trend chart (auto-fitted to slide bounds)
- produces a **2-slide deck**:
  1) Executive One-Pager (KPIs + chart)
  2) Appendix (insights, risks, actions, drivers, method)

Created by **Neil Sable**.

---

## Why this exists

Teams already have dashboards and data. The hard part is the last mile:  
**turning metrics into a clean, leadership-ready slide** every week.

InsightDeck automates that last mile.

---

## Input format

Upload a **CSV** (max **10 MB**) with these required columns:

| Column | Type | Example |
|---|---|---|
| `day` | date (`YYYY-MM-DD`) | `2025-10-01` |
| `service` | string | `CorePlatform` |
| `usage_units` | int | `1900` |
| `cost_gbp` | float | `410.5` |
| `incidents` | int | `2` |
| `sla_pct` | float | `99.92` |

---

## Run locally

### 1) Setup
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
2) Start the app
bash
Copy code
export PYTHONPATH=.
python -m uvicorn app.main:app --reload
Open:

http://127.0.0.1:8000

API
Generate deck
POST /generate-deck

Request: multipart form upload file (CSV)

Response: .pptx file download

Deployment
This repo is prepared for Vercel import:

requirements.txt for dependencies

vercel.json for routing FastAPI correctly

Serverless compatibility note
Matplotlib caching is redirected to /tmp (recommended for serverless environments).

Roadmap (optional)
SQL upload + SELECT-only validation

Multiple deck templates (Ops / Finance / Retail)

Branding controls (palette/logo)

LLM narrative plug-in (optional)
MD

yaml
Copy code

---

# 3) Ensure requirements and vercel routing exist (quick check)

```bash
test -f requirements.txt || cat > requirements.txt <<'REQ'
fastapi
uvicorn
python-multipart
pandas
matplotlib
python-pptx
pillow
REQ
bash
Copy code
test -f vercel.json || cat > vercel.json <<'JSON'
{
  "builds": [
    { "src": "app/main.py", "use": "@vercel/python" }
  ],
  "routes": [
    { "src": "/(.*)", "dest": "app/main.py" }
  ]
}
JSON
