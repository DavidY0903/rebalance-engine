from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import shutil
import tempfile
from rebalance_engine_v1_4 import run_engine

app = FastAPI()

# Allow frontend ↔ backend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # ⚠️ for open use; restrict later if needed
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve index.html at root
@app.get("/", response_class=HTMLResponse)
async def read_index():
    here = os.path.dirname(__file__)
    with open(os.path.join(here, "index.html"), "r", encoding="utf-8") as f:
        return f.read()

# API endpoint for rebalancing
@app.post("/api/rebalance")
async def rebalance(file: UploadFile, output_name: str = Form(None)):
    try:
        # Save uploaded file temporarily
        tmpdir = tempfile.mkdtemp(prefix="rebalance_")
        input_path = os.path.join(tmpdir, file.filename)
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Decide output filename
        if not output_name:
            output_name = "rebalance_result.xlsx"
        if not output_name.endswith(".xlsx"):
            output_name += ".xlsx"
        output_path = os.path.join(tmpdir, output_name)

        # Run engine
        run_engine(input_path, output_path)

        # Return file for download
        return FileResponse(
            path=output_path,
            filename=output_name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
