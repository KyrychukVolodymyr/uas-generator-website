from __future__ import annotations

import subprocess
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
RUNTIME_DIR = Path.home() / "Library" / "Application Support" / "UASGenerator2026V2"
VENV_DIR = RUNTIME_DIR / "runtime_venv"
REQUIRED_PACKAGES = ["pandas", "openpyxl", "pypdf", "reportlab", "requests", "certifi"]


def venv_python() -> Path:
    return VENV_DIR / "bin" / "python3"


def ensure_runtime() -> Path:
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)
    if not venv_python().exists():
        subprocess.run([sys.executable, "-m", "venv", str(VENV_DIR)], check=True)
    py = venv_python()
    subprocess.run([str(py), "-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel"], check=True)
    subprocess.run([str(py), "-m", "pip", "install", *REQUIRED_PACKAGES], check=True)
    return py


def verify_imports(py: Path) -> None:
    code = "import pandas, openpyxl, pypdf, reportlab, requests, certifi; print('IMPORTS_OK=True')"
    subprocess.run([str(py), "-c", code], check=True)


def main() -> int:
    try:
        py = ensure_runtime()
        verify_imports(py)
        cmd = [str(py), str(APP_DIR / "app_mac.py"), *sys.argv[1:]]
        return subprocess.run(cmd, cwd=str(APP_DIR)).returncode
    except Exception as exc:
        print(f"BOOTSTRAP_RUNTIME_FAILED={exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
