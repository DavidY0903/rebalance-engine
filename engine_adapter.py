# engine_adapter.py
import tempfile
from pathlib import Path
from rebalance_engine_v1_4 import run_engine
from datetime import datetime, timedelta
import re
import os
import pandas as pd

def run_rebalance(input_path: str, *, output_filename: str | None = None) -> str:
    tmpdir = Path(tempfile.mkdtemp(prefix="rebalance_"))

    if output_filename:
        out_path = tmpdir / output_filename
    else:
        hk_now = datetime.utcnow() + timedelta(hours=8)
        timestamp_str = hk_now.strftime("%Y-%m-%d %H %M %S")

        # -------------------------------
        # ✅ Improved user tag extraction
        # -------------------------------
        try:
            user_df__tmp = pd.read_excel(input_path, sheet_name="Input File", header=None)
            _a1 = user_df__tmp.iloc[0, 0] if user_df__tmp.shape[0] > 0 and user_df__tmp.shape[1] > 0 else ""
            if isinstance(_a1, str) and _a1.strip():
                user_name = _a1.strip()
            else:
                base = os.path.basename(input_path)
                parts = re.findall(r"\(([^)]+)\)", base)
                user_name = ", ".join(parts) if parts else "User"
        except Exception:
            base = os.path.basename(input_path)
            parts = re.findall(r"\(([^)]+)\)", base)
            user_name = ", ".join(parts) if parts else "User"

        output_filename = f"rebalance recommendation ({user_name})({timestamp_str}).xlsx"
        out_path = tmpdir / output_filename

    return run_engine(str(input_path), str(out_path))
