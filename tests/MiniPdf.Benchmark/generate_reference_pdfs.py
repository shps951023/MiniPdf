"""
Generate reference PDFs from .xlsx files using LibreOffice (soffice).

Prerequisites:
    1. Install LibreOffice: https://www.libreoffice.org/download/
    2. Ensure 'soffice' is on PATH, or set LIBREOFFICE_PATH env var.
       - Windows default: C:\\Program Files\\LibreOffice\\program\\soffice.exe
       - macOS:  /Applications/LibreOffice.app/Contents/MacOS/soffice
       - Linux:  /usr/bin/soffice

Usage:
    python generate_reference_pdfs.py [--xlsx-dir ../MiniPdf.Scripts/output] [--pdf-dir ./reference_pdfs]

This converts every .xlsx in the input directory to PDF using LibreOffice,
producing the "ground truth" reference that MiniPdf output is compared against.
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def find_libreoffice() -> str:
    """Locate the LibreOffice soffice executable."""
    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path and os.path.isfile(env_path):
        return env_path

    # Common locations
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c

    # Try PATH
    which = shutil.which("soffice") or shutil.which("libreoffice")
    if which:
        return which

    print("ERROR: LibreOffice not found. Install it or set LIBREOFFICE_PATH env var.")
    sys.exit(1)


def convert_xlsx_to_pdf(soffice: str, xlsx_path: str, output_dir: str) -> bool:
    """Convert a single .xlsx to PDF via LibreOffice."""
    try:
        # Use a unique user profile to avoid lock conflicts
        with tempfile.TemporaryDirectory() as tmp_profile:
            cmd = [
                soffice,
                "--headless",
                "--norestore",
                f"-env:UserInstallation=file:///{tmp_profile.replace(os.sep, '/')}",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                xlsx_path,
            ]
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )
            if result.returncode != 0:
                print(f"  ERR {Path(xlsx_path).name}: {result.stderr.strip()}")
                return False
            return True
    except subprocess.TimeoutExpired:
        print(f"  TIMEOUT {Path(xlsx_path).name}")
        return False
    except Exception as e:
        print(f"  ERR {Path(xlsx_path).name}: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Generate reference PDFs via LibreOffice")
    parser.add_argument("--xlsx-dir", default=os.path.join("..", "MiniPdf.Scripts", "output"),
                        help="Directory containing .xlsx files")
    parser.add_argument("--pdf-dir", default="reference_pdfs",
                        help="Output directory for reference PDFs")
    args = parser.parse_args()

    xlsx_dir = os.path.abspath(args.xlsx_dir)
    pdf_dir = os.path.abspath(args.pdf_dir)

    if not os.path.isdir(xlsx_dir):
        print(f"ERROR: xlsx directory not found: {xlsx_dir}")
        print("Run generate_classic_xlsx.py first to create test Excel files.")
        sys.exit(1)

    os.makedirs(pdf_dir, exist_ok=True)

    soffice = find_libreoffice()
    print(f"LibreOffice: {soffice}")
    print(f"Input:  {xlsx_dir}")
    print(f"Output: {pdf_dir}")
    print()

    xlsx_files = sorted(Path(xlsx_dir).glob("*.xlsx"))
    if not xlsx_files:
        print("No .xlsx files found.")
        sys.exit(1)

    passed = 0
    failed = 0
    for xlsx in xlsx_files:
        ok = convert_xlsx_to_pdf(soffice, str(xlsx), pdf_dir)
        if ok:
            pdf_name = xlsx.stem + ".pdf"
            print(f"  OK  {pdf_name}")
            passed += 1
        else:
            failed += 1

    print(f"\nDone! Passed: {passed}, Failed: {failed}, Total: {len(xlsx_files)}")


if __name__ == "__main__":
    main()
