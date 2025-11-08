from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse, HTMLResponse
from pathlib import Path
import tempfile
import uvicorn
import urllib.parse
from fastapi.staticfiles import StaticFiles


from engine_adapter import run_rebalance

app = FastAPI()

# ✅ Mount static files (CSS + JS)
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/api/rebalance")
async def api_rebalance(file: UploadFile = File(...)):
    try:
        # Save uploaded file to tmpdir
        tmpdir = Path(tempfile.mkdtemp(prefix="upload_"))
        in_path = tmpdir / file.filename
        with open(in_path, "wb") as f:
            f.write(await file.read())

        # Run engine — it auto-generates the dynamic output filename
        out_path = Path(run_rebalance(str(in_path)))

        # Encode filename for Content-Disposition
        encoded_filename = urllib.parse.quote(out_path.name)

        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }

        return StreamingResponse(
            open(out_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

@app.get("/", response_class=HTMLResponse)
def index():
    with open("index.html") as f:
        return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
