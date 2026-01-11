<div align="center">

# 🚀 InsightDeck

### From raw data to an executive-ready deck — in seconds.

<p align="center">
  <a href="https://insight-deck-sandy.vercel.app/" target="_blank">
    <img src="https://img.shields.io/badge/🚀%20Live%20App-Open%20InsightDeck-blue?style=for-the-badge" />
  </a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/FastAPI-Backend-success" />
  <img src="https://img.shields.io/badge/Python-Automation-blue" />
  <img src="https://img.shields.io/badge/PowerPoint-Automated-orange" />
  <img src="https://img.shields.io/badge/Deployed-Vercel-black" />
</p>

</div>

---

## ✨ What is InsightDeck?

**InsightDeck** is a lightweight analytics-to-presentation engine.

It takes **structured operational data (CSV)** and automatically generates an  
**executive-ready PowerPoint deck** — complete with:

- KPI tiles
- Trend charts (auto-fitted, no overflow)
- Clear insights, risks, and actions
- Consistent layout rules (no overlaps, no broken slides)

This eliminates the most painful part of reporting:
> _Turning data into leadership-ready slides every week._

---

## 🎯 Why this matters

Dashboards already exist.  
Executives still ask for **slides**.

InsightDeck bridges that gap by automating the **last mile of analytics**:
- No manual formatting
- No copy-pasting charts
- No layout fixing at midnight

Just **upload → generate → present**.

---

## 🧠 What happens under the hood

```text
CSV Upload
   ↓
Data Validation & KPI Computation (Python + Pandas)
   ↓
Trend Analysis & Chart Generation (Matplotlib)
   ↓
Slide Layout Engine (python-pptx)
   ↓
Executive PowerPoint (.pptx)
All charts and files are generated safely using serverless-compatible paths.

📊 Input format
Upload a CSV (≤ 10 MB) with the following required columns:

Column	Type	Example
day	date	2025-10-01
service	text	CorePlatform
usage_units	int	1900
cost_gbp	float	410.50
incidents	int	2
sla_pct	float	99.92

🖥️ Run locally
bash
Copy code
# 1. Create environment
python3 -m venv .venv
source .venv/bin/activate

# 2. Install dependencies
python -m pip install -r requirements.txt

# 3. Run the app
export PYTHONPATH=.
python -m uvicorn app.main:app --reload
Open:

cpp
Copy code
http://127.0.0.1:8000
☁️ Live deployment
The app is deployed on Vercel:

👉 https://insight-deck-sandy.vercel.app/

The repository is prepared for serverless execution:

Clean dependency management

Deterministic routing

/tmp-based file generation

No environment-specific assumptions

🧩 Tech stack
FastAPI — API & UI routing

Python — data processing & automation

Pandas — KPI computation

Matplotlib — chart rendering

python-pptx — slide generation

Vercel — serverless deployment

🛣️ Roadmap
SQL upload (read-only SELECT validation)

Multiple deck templates (Ops / Finance / Retail)

Branding controls (logo, palette, typography)

Optional AI narrative layer (pluggable)

<div align="center">
Built by Neil Sable
Data → Insight → Decision

</div> ```
