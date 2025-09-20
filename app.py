from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from pathlib import Path
import tempfile
import uvicorn

from engine_adapter import run_rebalance

app = FastAPI(title="Rebalance Engine v1.4 API", version="1.0")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/api/rebalance")
async def api_rebalance(file: UploadFile = File(...)):
    try:
        tmpdir = Path(tempfile.mkdtemp(prefix="upload_"))
        in_path = tmpdir / file.filename
        with open(in_path, "wb") as f:
            f.write(await file.read())

        out_path = run_rebalance(str(in_path))  # auto-generate filename

        return FileResponse(
            path=out_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=Path(out_path).name,
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.get("/", response_class=HTMLResponse)
def index():
    with open("index.html") as f:
        return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
