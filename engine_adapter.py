import tempfile
from pathlib import Path
from datetime import datetime, timedelta
import re
import os
import pandas as pd

# ✅ Import v1.5 main() function
import rebalance_engine_v1_5


def run_rebalance(input_path: str, *, output_filename: str | None = None) -> str:
    tmpdir = Path(tempfile.mkdtemp(prefix="rebalance_"))

    if output_filename:
        out_path = tmpdir / output_filename
    else:
        hk_now = datetime.utcnow() + timedelta(hours=8)
        timestamp_str = hk_now.strftime("%Y-%m-%d %H %M %S")

        # ✅ Improved user tag extraction (optional A1 detection)
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

    # ✅ Patch: override tkinter-based file picker
    rebalance_engine_v1_5._pick_input_filename = lambda: str(input_path)

    # ✅ Patch: override output filename behavior (inside v1.5 engine)
    # This ensures output goes exactly to out_path instead of same folder as input
    original_save = rebalance_engine_v1_5.Workbook.save

    def patched_save(self, filename):
        original_save(self, str(out_path))

    original_save = rebalance_engine_v1_5.Workbook.save

    try:
        rebalance_engine_v1_5.Workbook.save = patched_save
        rebalance_engine_v1_5.main()
    finally:
        rebalance_engine_v1_5.Workbook.save = original_save

    return str(out_path)
