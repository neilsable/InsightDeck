#!/bin/bash
set -e
cd "$(dirname "$0")"

source ".venv/bin/activate"

echo "Starting InsightDeck on http://127.0.0.1:8000"
echo "Docs UI: http://127.0.0.1:8000/docs"

python3 -m uvicorn app.main:app --reload --port 8000
