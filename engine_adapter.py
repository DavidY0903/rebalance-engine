# engine_adapter.py
import tempfile
from pathlib import Path
from rebalance_engine_v1_4 import run_engine

def run_rebalance(input_path: str, *, output_filename: str | None = None) -> str:
    """
    Adapter for HTTP service.
    Produces Excel in a temp dir and returns the path.
    """
    tmpdir = Path(tempfile.mkdtemp(prefix="rebalance_"))
    out_path = tmpdir / (output_filename or f"rebalance_{Path(input_path).stem}.xlsx")
    return run_engine(str(input_path), str(out_path))
