# engine_adapter.py
import tempfile
from pathlib import Path
from rebalance_engine_v1_4 import run_engine
from datetime import datetime, timedelta

def run_rebalance(input_path: str, *, output_filename: str | None = None) -> str:
    """
    Adapter for HTTP service.
    Produces Excel in a temp dir and returns the path.
    """
    tmpdir = Path(tempfile.mkdtemp(prefix="rebalance_"))

    if output_filename:
        # Use custom filename if explicitly given
        out_path = tmpdir / output_filename
    else:
        # Auto-generate timestamped filename based on input file stem
        hk_now = datetime.utcnow() + timedelta(hours=8)
        timestamp_str = hk_now.strftime("%Y-%m-%d %H %M %S")
        stem = Path(input_path).stem
        out_path = tmpdir / f"rebalance_{stem}({timestamp_str}).xlsx"

    return run_engine(str(input_path), str(out_path))
