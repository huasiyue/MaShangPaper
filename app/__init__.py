from pathlib import Path


# Allow `uvicorn app.main:app` to run from the repository root by pointing the
# top-level `app` package to the real backend package directory.
__path__ = [str(Path(__file__).resolve().parent.parent / "backend" / "app")]
