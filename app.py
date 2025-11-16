from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import tempfile
import uvicorn
import urllib.parse
import os

from engine_adapter import run_rebalance

# =====================================================
# 📘 Rebalance Engine v1.5 — FastAPI Web Service
# -----------------------------------------------------
# Supports drag-and-drop UI, bilingual output, and
# dynamic Excel download.
# =====================================================

app = FastAPI(title="Rebalance Engine v1.5", version="1.5")

# ✅ Mount static assets (CSS + JS)
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.get("/health")
def health():
    """Simple health check endpoint for monitoring."""
    return {"status": "ok"}


@app.post("/api/rebalance")
async def api_rebalance(file: UploadFile = File(...)):
    """
    Main API endpoint.
    Accepts Excel input, runs the v1.5 engine, and streams the output file.
    """
    try:
        # ✅ 1. Save uploaded file to temporary directory
        tmpdir = Path(tempfile.mkdtemp(prefix="upload_"))
        in_path = tmpdir / file.filename
        with open(in_path, "wb") as f:
            f.write(await file.read())

        # ✅ 2. Run engine (returns full output path)
        out_path = Path(run_rebalance(str(in_path)))

        if not out_path.exists():
            raise FileNotFoundError("Engine did not produce an output file.")

        # ✅ 3. Encode filename for correct download in browsers
        encoded_filename = urllib.parse.quote(out_path.name)
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }

        # ✅ 4. Stream file back to client
        return StreamingResponse(
            open(out_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    except Exception as e:
        # ✅ 5. Catch and return all errors in JSON form
        return JSONResponse(status_code=500, content={"error": str(e)})


@app.get("/", response_class=HTMLResponse)
def index():
    """Serve main frontend (index.html)."""
    index_path = Path("index.html")
    if not index_path.exists():
        return HTMLResponse("<h2>index.html not found.</h2>", status_code=404)
    with open(index_path, encoding="utf-8") as f:
        return f.read()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
