import uuid
import shutil

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from pathlib import Path
from fastapi.staticfiles import StaticFiles
import os

app = FastAPI(title="InsightDeck")


from starlette.responses import JSONResponse
import traceback

@app.middleware("http")
async def catch_exceptions_middleware(request, call_next):
    try:
        return await call_next(request)
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(
            {"error": str(e), "type": e.__class__.__name__},
            status_code=500
        )

# Serve static UI
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/", response_class=HTMLResponse)
def home():
    with open(os.path.join(STATIC_DIR, "index.html")) as f:
        return f.read()

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOADS = Path("/tmp/uploads")
OUTPUTS = Path("/tmp/outputs")
STATIC = BASE_DIR / "app" / "static"

UPLOADS.mkdir(exist_ok=True)
OUTPUTS.mkdir(exist_ok=True)

MAX_MB = 10
MAX_BYTES = MAX_MB * 1024 * 1024

@app.get("/", response_class=HTMLResponse)
def home():
    index_path = STATIC / "index.html"
    if not index_path.exists():
        return HTMLResponse("<h1>InsightDeck</h1><p>UI missing: app/static/index.html</p>", status_code=500)
    return HTMLResponse(index_path.read_text(encoding="utf-8"))

@app.get("/health")
def health():
    return {"status": "ok", "app": "InsightDeck"}

@app.post("/generate-deck")
async def generate_deck(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".csv"):
        return JSONResponse(
            {"error": "Invalid input. Please upload a .csv file."},
            status_code=400
        )

    job_id = str(uuid.uuid4())[:8]
    csv_path = Path("/tmp") / f"{job_id}_{file.filename}"
    out_path = Path("/tmp") / f"InsightDeck_{job_id}.pptx"
    # Save upload to disk
    with csv_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)

    # Server-side max size check
    size = os.path.getsize(csv_path)
    if size > MAX_BYTES:
        csv_path.unlink(missing_ok=True)
        return JSONResponse(
            {"error": f"File too large. Max size is {MAX_MB} MB."},
            status_code=400
        )

    try:
        generate_ppt_from_csv(csv_path=csv_path, out_pptx=out_path)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)

    return FileResponse(
        path=str(out_path),
        filename=out_path.name,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
@app.get("/")
def root():
    return {"status": "ok", "app": "InsightDeck"}

