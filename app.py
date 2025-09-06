from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
import os
import shutil
import tempfile
from rebalance_engine_v1_4 import run_engine

app = FastAPI()

@app.get("/")
def root():
    return {"message": "Rebalance Engine API is running"}

# ✅ Health check for Render
@app.get("/healthz")
def health_check():
    return {"status": "ok"}

@app.post("/api/rebalance")
async def rebalance(file: UploadFile = File(...), output_name: str = Form(None)):
    try:
        # Save upload to temp dir
        tmpdir = tempfile.mkdtemp(prefix="rebalance_")
        input_path = os.path.join(tmpdir, file.filename)
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Pick output filename
        if not output_name:
            base, _ = os.path.splitext(file.filename)
            output_name = f"{base}_rebalance.xlsx"
        output_path = os.path.join(tmpdir, output_name)

        # Run engine
        run_engine(input_path, output_path)

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=output_name
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
