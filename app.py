# app.py
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from pathlib import Path
import tempfile
import uvicorn

from engine_adapter import run_rebalance

from fastapi.responses import HTMLResponse

app = FastAPI(title="Rebalance Engine v1.4 API", version="1.0")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/api/rebalance")
async def api_rebalance(
    file: UploadFile = File(..., description="Input Excel file"),
    output_name: str | None = Form(None),
):
    try:
        tmpdir = Path(tempfile.mkdtemp(prefix="upload_"))
        in_path = tmpdir / file.filename
        with open(in_path, "wb") as f:
            f.write(await file.read())

        out_path = run_rebalance(str(in_path), output_filename=output_name)

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
    
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

# serve static files (e.g., index.html, CSS, JS)
app.mount("/static", StaticFiles(directory="."), name="static")

@app.get("/")
async def root():
    return FileResponse("index.html")


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
