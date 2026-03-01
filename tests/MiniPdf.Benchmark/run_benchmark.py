"""
Automated benchmark test: generates Excel files, converts them to PDF
via MiniPdf and LibreOffice, compares the results, and produces a report.

This is the single entry point for the full "self-evolution" pipeline.

Prerequisites:
    pip install openpyxl pymupdf
    LibreOffice installed (for reference PDF generation)
    .NET 9 SDK (for MiniPdf)

Usage:
    python run_benchmark.py                   # full pipeline
    python run_benchmark.py --skip-generate   # skip Excel generation
    python run_benchmark.py --skip-reference   # skip LibreOffice conversion
    python run_benchmark.py --skip-minipdf     # skip MiniPdf conversion
    python run_benchmark.py --compare-only     # only run comparison (assumes PDFs exist)
"""

import argparse
import os
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
XLSX_DIR = SCRIPT_DIR / ".." / "MiniPdf.Scripts" / "output"
MINIPDF_PDF_DIR = SCRIPT_DIR / ".." / "MiniPdf.Scripts" / "pdf_output"
REFERENCE_PDF_DIR = SCRIPT_DIR / "reference_pdfs"
REPORT_DIR = SCRIPT_DIR / "reports"


def banner(msg: str):
    print(f"\n{'='*60}")
    print(f"  {msg}")
    print(f"{'='*60}\n")


def run(cmd: list[str], cwd: str = None, check: bool = True) -> int:
    """Run a command and return exit code."""
    print(f"  > {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=cwd)
    if check and result.returncode != 0:
        print(f"  ⚠ Command exited with code {result.returncode}")
    return result.returncode


def step_generate_xlsx():
    """Step 1: Generate test Excel files using openpyxl."""
    banner("Step 1: Generate Test Excel Files")
    scripts_dir = SCRIPT_DIR / ".." / "MiniPdf.Scripts"
    return run(
        [sys.executable, "generate_classic_xlsx.py"],
        cwd=str(scripts_dir),
    )


def step_generate_minipdf_pdfs():
    """Step 2: Convert Excel files to PDF using MiniPdf."""
    banner("Step 2: Convert Excel → PDF (MiniPdf)")
    scripts_dir = SCRIPT_DIR / ".." / "MiniPdf.Scripts"

    # Use dotnet run with the .cs script
    return run(
        ["dotnet", "run", "--project", "convert_xlsx_to_pdf.cs"],
        cwd=str(scripts_dir),
    )


def step_generate_reference_pdfs():
    """Step 3: Convert Excel files to PDF using LibreOffice (reference)."""
    banner("Step 3: Convert Excel → PDF (LibreOffice Reference)")
    return run(
        [sys.executable, "generate_reference_pdfs.py",
         "--xlsx-dir", str(XLSX_DIR.resolve()),
         "--pdf-dir", str(REFERENCE_PDF_DIR.resolve())],
        cwd=str(SCRIPT_DIR),
        check=False,
    )


def step_compare():
    """Step 4: Compare MiniPdf PDFs against reference PDFs."""
    banner("Step 4: Compare MiniPdf vs Reference")
    return run(
        [sys.executable, "compare_pdfs.py",
         "--minipdf-dir", str(MINIPDF_PDF_DIR.resolve()),
         "--reference-dir", str(REFERENCE_PDF_DIR.resolve()),
         "--report-dir", str(REPORT_DIR.resolve())],
        cwd=str(SCRIPT_DIR),
    )


def step_analyze_report():
    """Step 5: Print key findings from the report."""
    banner("Step 5: Analysis Summary")
    json_path = REPORT_DIR / "comparison_report.json"
    md_path = REPORT_DIR / "comparison_report.md"

    if json_path.exists():
        import json
        with open(json_path, "r", encoding="utf-8") as f:
            results = json.load(f)

        total = len(results)
        scores = [r.get("overall_score", 0) for r in results]
        avg = sum(scores) / total if total else 0
        excellent = sum(1 for s in scores if s >= 0.9)
        good = sum(1 for s in scores if 0.7 <= s < 0.9)
        poor = sum(1 for s in scores if s < 0.7)

        print(f"  Total test cases: {total}")
        print(f"  Average score:    {avg:.4f}")
        print(f"  Excellent (≥0.9): {excellent}")
        print(f"  Good (0.7-0.9):   {good}")
        print(f"  Poor (<0.7):      {poor}")
        print()

        if poor > 0:
            print("  ⚠ Cases needing improvement:")
            for r in sorted(results, key=lambda x: x.get("overall_score", 0)):
                score = r.get("overall_score", 0)
                if score < 0.7:
                    print(f"    - {r['name']}: {score}")
            print()

        print(f"  Full report: {md_path}")
        print(f"  JSON data:   {json_path}")
    else:
        print("  No report found. Run the full pipeline first.")


def main():
    parser = argparse.ArgumentParser(description="MiniPdf Benchmark Pipeline")
    parser.add_argument("--skip-generate", action="store_true", help="Skip Excel generation")
    parser.add_argument("--skip-minipdf", action="store_true", help="Skip MiniPdf PDF conversion")
    parser.add_argument("--skip-reference", action="store_true", help="Skip LibreOffice reference conversion")
    parser.add_argument("--compare-only", action="store_true", help="Only run comparison step")
    args = parser.parse_args()

    banner("MiniPdf Self-Evolution Benchmark Pipeline")
    print(f"  XLSX dir:      {XLSX_DIR.resolve()}")
    print(f"  MiniPdf PDFs:  {MINIPDF_PDF_DIR.resolve()}")
    print(f"  Reference PDFs:{REFERENCE_PDF_DIR.resolve()}")
    print(f"  Reports:       {REPORT_DIR.resolve()}")

    if args.compare_only:
        step_compare()
        step_analyze_report()
        return

    if not args.skip_generate:
        step_generate_xlsx()

    if not args.skip_minipdf:
        step_generate_minipdf_pdfs()

    if not args.skip_reference:
        step_generate_reference_pdfs()

    step_compare()
    step_analyze_report()

    banner("Pipeline Complete")


if __name__ == "__main__":
    main()
