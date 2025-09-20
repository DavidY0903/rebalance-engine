# engine_adapter.py
import tempfile
from pathlib import Path
from rebalance_engine_v1_4 import run_engine
from datetime import datetime, timedelta
import re
import os
import pandas as pd

def run_rebalance(input_path: str, *, output_filename: str | None = None) -> str:
    """
    Adapter for HTTP service.
    Produces Excel in a temp dir and returns the path.
    Ensures consistent naming:
      rebalance recommendation (<user tag>)(YYYY-MM-DD HH MM SS).xlsx
    """
    tmpdir = Path(tempfile.mkdtemp(prefix="rebalance_"))

    if output_filename:
        out_path = tmpdir / output_filename
    else:
        # generate HK timestamp
        hk_now = datetime.utcnow() + timedelta(hours=8)
        timestamp_str = hk_now.strftime("%Y-%m-%d %H %M %S")

        # try to extract <user tag> from Input File sheet A1
        try:
            user_df__tmp = pd.read_excel(input_path, sheet_name="Input File", header=None)
            _a1 = user_df__tmp.iloc[0, 0] if user_df__tmp.shape[0] > 0 and user_df__tmp.shape[1] > 0 else ""
            if isinstance(_a1, str) and _a1.strip():
                user_name = _a1.strip()
            else:
                m = re.search(r"\(([^)]+)\)", os.path.basename(input_path))
                user_name = m.group(1) if m else "User"
        except Exception:
            m = re.search(r"\(([^)]+)\)", os.path.basename(input_path))
            user_name = m.group(1) if m else "User"

        output_filename = f"rebalance recommendation ({user_name})({timestamp_str}).xlsx"
        out_path = tmpdir / output_filename

    return run_engine(str(input_path), str(out_path))
