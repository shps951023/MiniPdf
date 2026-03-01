"""
Compare MiniPdf-generated PDFs against LibreOffice reference PDFs.

Produces a detailed comparison report including:
  - Text content diff (extracted text comparison)
  - Page count comparison
  - File size comparison
  - Visual pixel diff (if pdf2image / Poppler is available)

Prerequisites:
    pip install pymupdf   # for text extraction + rendering

Usage:
    python compare_pdfs.py [--minipdf-dir ./minipdf_pdfs] [--reference-dir ./reference_pdfs] [--report-dir ./reports]
"""

import argparse
import difflib
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# Try to import fitz (PyMuPDF) for text extraction and visual comparison
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False
    print("WARNING: PyMuPDF not installed. Install with: pip install pymupdf")
    print("         Text extraction and visual comparison will be disabled.\n")


def extract_text_pymupdf(pdf_path: str) -> list[str]:
    """Extract text from each page using PyMuPDF."""
    pages = []
    doc = fitz.open(pdf_path)
    for page in doc:
        pages.append(page.get_text("text"))
    doc.close()
    return pages


def extract_text_fallback(pdf_path: str) -> list[str]:
    """Fallback: extract raw ASCII strings from PDF binary."""
    with open(pdf_path, "rb") as f:
        data = f.read()
    # Very rough extraction: find text between BT..ET operators
    text = data.decode("latin-1", errors="replace")
    # Extract parenthesized strings (PDF text objects)
    import re
    strings = re.findall(r"\(([^)]*)\)", text)
    return ["\n".join(strings)]


def render_page_to_pixels(pdf_path: str, page_num: int, dpi: int = 150):
    """Render a PDF page to a pixel map using PyMuPDF. Returns (width, height, samples)."""
    doc = fitz.open(pdf_path)
    if page_num >= len(doc):
        doc.close()
        return None
    page = doc[page_num]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    result = (pix.width, pix.height, pix.samples)
    doc.close()
    return result


def pixel_diff_score(pix1, pix2) -> float:
    """
    Compare two pixmaps and return a similarity score 0.0-1.0.
    1.0 = identical, 0.0 = completely different.
    """
    if pix1 is None or pix2 is None:
        return 0.0

    w1, h1, s1 = pix1
    w2, h2, s2 = pix2

    if w1 != w2 or h1 != h2:
        # Different dimensions - compare what we can
        min_len = min(len(s1), len(s2))
        if min_len == 0:
            return 0.0
        matching = sum(1 for a, b in zip(s1[:min_len], s2[:min_len]) if a == b)
        return matching / min_len

    total = len(s1)
    if total == 0:
        return 1.0
    matching = sum(1 for a, b in zip(s1, s2) if a == b)
    return matching / total


def save_visual_diff(pdf1_path: str, pdf2_path: str, output_dir: str, name: str, dpi: int = 150):
    """Save visual diff images for each page."""
    if not HAS_FITZ:
        return []

    diff_images = []
    doc1 = fitz.open(pdf1_path)
    doc2 = fitz.open(pdf2_path)
    max_pages = max(len(doc1), len(doc2))

    for i in range(max_pages):
        mat = fitz.Matrix(dpi / 72, dpi / 72)

        if i < len(doc1):
            pix1 = doc1[i].get_pixmap(matrix=mat, alpha=False)
        else:
            pix1 = None

        if i < len(doc2):
            pix2 = doc2[i].get_pixmap(matrix=mat, alpha=False)
        else:
            pix2 = None

        # Save individual renderings
        if pix1:
            path1 = os.path.join(output_dir, f"{name}_p{i+1}_minipdf.png")
            pix1.save(path1)

        if pix2:
            path2 = os.path.join(output_dir, f"{name}_p{i+1}_reference.png")
            pix2.save(path2)

        diff_images.append({
            "page": i + 1,
            "minipdf_img": f"{name}_p{i+1}_minipdf.png" if pix1 else None,
            "reference_img": f"{name}_p{i+1}_reference.png" if pix2 else None,
        })

    doc1.close()
    doc2.close()
    return diff_images


def compare_single(minipdf_path: str, reference_path: str, report_images_dir: str, name: str) -> dict:
    """Compare a single pair of PDFs and return a detailed result."""
    result = {
        "name": name,
        "minipdf_exists": os.path.isfile(minipdf_path),
        "reference_exists": os.path.isfile(reference_path),
    }

    if not result["minipdf_exists"]:
        result["error"] = "MiniPdf PDF not found"
        result["score"] = 0.0
        return result

    if not result["reference_exists"]:
        result["error"] = "Reference PDF not found"
        result["score"] = 0.0
        return result

    # File sizes
    result["minipdf_size"] = os.path.getsize(minipdf_path)
    result["reference_size"] = os.path.getsize(reference_path)

    # Page counts
    if HAS_FITZ:
        doc_m = fitz.open(minipdf_path)
        doc_r = fitz.open(reference_path)
        result["minipdf_pages"] = len(doc_m)
        result["reference_pages"] = len(doc_r)
        doc_m.close()
        doc_r.close()
    else:
        result["minipdf_pages"] = "?"
        result["reference_pages"] = "?"

    # Text extraction and comparison
    if HAS_FITZ:
        try:
            text_m = extract_text_pymupdf(minipdf_path)
            text_r = extract_text_pymupdf(reference_path)
        except Exception as e:
            text_m = extract_text_fallback(minipdf_path)
            text_r = extract_text_fallback(reference_path)
            result["text_extract_warning"] = str(e)
    else:
        text_m = extract_text_fallback(minipdf_path)
        text_r = extract_text_fallback(reference_path)

    # Flatten text for comparison
    flat_m = "\n---PAGE---\n".join(text_m).strip()
    flat_r = "\n---PAGE---\n".join(text_r).strip()

    # Text similarity (SequenceMatcher) — page-aware
    if len(flat_m) == 0 and len(flat_r) == 0:
        # Both empty — treat as identical
        result["text_similarity"] = 1.0
    else:
        sm = difflib.SequenceMatcher(None, flat_m, flat_r)
        result["text_similarity"] = round(sm.ratio(), 4)

    # Also compute flat text similarity (ignoring page breaks)
    # This is fairer when page break positions differ but content is the same
    flat_m_no_page = flat_m.replace("\n---PAGE---\n", "\n")
    flat_r_no_page = flat_r.replace("\n---PAGE---\n", "\n")
    if len(flat_m_no_page) == 0 and len(flat_r_no_page) == 0:
        result["flat_text_similarity"] = 1.0
    else:
        sm_flat = difflib.SequenceMatcher(None, flat_m_no_page, flat_r_no_page)
        result["flat_text_similarity"] = round(sm_flat.ratio(), 4)

    # Use the higher of page-aware and flat text similarity
    result["text_similarity"] = max(result["text_similarity"], result["flat_text_similarity"])

    # Unified diff
    diff_lines = list(difflib.unified_diff(
        flat_m.splitlines(keepends=True),
        flat_r.splitlines(keepends=True),
        fromfile=f"minipdf/{name}.pdf",
        tofile=f"reference/{name}.pdf",
        lineterm="",
    ))
    result["text_diff"] = "\n".join(diff_lines) if diff_lines else "(identical)"

    # Visual comparison
    visual_scores = []
    if HAS_FITZ:
        max_pages = max(result["minipdf_pages"], result["reference_pages"])
        for p in range(max_pages):
            pix_m = render_page_to_pixels(minipdf_path, p)
            pix_r = render_page_to_pixels(reference_path, p)
            score = pixel_diff_score(pix_m, pix_r)
            visual_scores.append(round(score, 4))

        result["visual_scores"] = visual_scores
        result["visual_avg"] = round(sum(visual_scores) / len(visual_scores), 4) if visual_scores else 0.0

        # Save diff images
        os.makedirs(report_images_dir, exist_ok=True)
        result["diff_images"] = save_visual_diff(minipdf_path, reference_path, report_images_dir, name)

    # Overall score: weighted average (text 40%, visual 40%, page-count match 20%)
    page_score = 1.0 if result.get("minipdf_pages") == result.get("reference_pages") else 0.5
    text_score = result["text_similarity"]
    vis_score = result.get("visual_avg", text_score)  # fallback to text if no visual

    result["overall_score"] = round(text_score * 0.4 + vis_score * 0.4 + page_score * 0.2, 4)

    return result


def generate_report(results: list[dict], report_dir: str):
    """Generate a markdown + JSON comparison report."""
    # JSON dump
    json_path = os.path.join(report_dir, "comparison_report.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False, default=str)

    # Markdown report
    md_path = os.path.join(report_dir, "comparison_report.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("# MiniPdf vs Reference PDF Comparison Report\n\n")
        f.write(f"Generated: {datetime.now().isoformat()}\n\n")

        # Summary table
        f.write("## Summary\n\n")
        f.write("| # | Test Case | Text Sim | Visual Avg | Pages (M/R) | Overall |\n")
        f.write("|---|-----------|----------|------------|-------------|--------|\n")

        for i, r in enumerate(results, 1):
            name = r["name"]
            text_sim = r.get("text_similarity", "N/A")
            vis_avg = r.get("visual_avg", "N/A")
            mp = r.get("minipdf_pages", "?")
            rp = r.get("reference_pages", "?")
            overall = r.get("overall_score", "N/A")

            # Color coding via emoji
            if isinstance(overall, (int, float)):
                if overall >= 0.9:
                    emoji = "🟢"
                elif overall >= 0.7:
                    emoji = "🟡"
                else:
                    emoji = "🔴"
            else:
                emoji = "⚪"

            f.write(f"| {i} | {emoji} {name} | {text_sim} | {vis_avg} | {mp}/{rp} | **{overall}** |\n")

        avg_overall = sum(r.get("overall_score", 0) for r in results) / len(results) if results else 0
        f.write(f"\n**Average Overall Score: {avg_overall:.4f}**\n\n")

        # Detailed sections
        f.write("## Detailed Results\n\n")
        for r in results:
            name = r["name"]
            f.write(f"### {name}\n\n")

            if "error" in r:
                f.write(f"**Error:** {r['error']}\n\n")
                continue

            f.write(f"- **Text Similarity:** {r.get('text_similarity', 'N/A')}\n")
            f.write(f"- **Visual Average:** {r.get('visual_avg', 'N/A')}\n")
            f.write(f"- **Overall Score:** {r.get('overall_score', 'N/A')}\n")
            f.write(f"- **Pages:** MiniPdf={r.get('minipdf_pages', '?')}, Reference={r.get('reference_pages', '?')}\n")
            f.write(f"- **File Size:** MiniPdf={r.get('minipdf_size', '?')} bytes, Reference={r.get('reference_size', '?')} bytes\n\n")

            diff = r.get("text_diff", "")
            if diff and diff != "(identical)":
                f.write("<details><summary>Text Diff</summary>\n\n```diff\n")
                # Truncate very long diffs
                if len(diff) > 3000:
                    f.write(diff[:3000])
                    f.write(f"\n... ({len(diff) - 3000} more characters)\n")
                else:
                    f.write(diff)
                f.write("\n```\n</details>\n\n")
            else:
                f.write("Text content: ✅ Identical\n\n")

        # Improvement suggestions
        f.write("## Improvement Suggestions\n\n")
        low_scores = [(r["name"], r.get("overall_score", 0)) for r in results if r.get("overall_score", 1) < 0.8]
        if low_scores:
            low_scores.sort(key=lambda x: x[1])
            f.write("The following test cases scored below 0.8 and need attention:\n\n")
            for name, score in low_scores:
                f.write(f"1. **{name}** (score: {score})\n")
            f.write("\nReview the text diffs and visual comparisons above to identify specific rendering issues.\n")
        else:
            f.write("All test cases scored 0.8 or above. 🎉\n")

    print(f"\nReports saved:")
    print(f"  Markdown: {md_path}")
    print(f"  JSON:     {json_path}")


def main():
    parser = argparse.ArgumentParser(description="Compare MiniPdf PDFs against reference PDFs")
    parser.add_argument("--minipdf-dir", default=os.path.join("..", "MiniPdf.Scripts", "pdf_output"),
                        help="Directory containing MiniPdf-generated PDFs")
    parser.add_argument("--reference-dir", default="reference_pdfs",
                        help="Directory containing reference PDFs (from LibreOffice)")
    parser.add_argument("--report-dir", default="reports",
                        help="Output directory for comparison reports")
    args = parser.parse_args()

    minipdf_dir = os.path.abspath(args.minipdf_dir)
    reference_dir = os.path.abspath(args.reference_dir)
    report_dir = os.path.abspath(args.report_dir)
    images_dir = os.path.join(report_dir, "images")

    os.makedirs(report_dir, exist_ok=True)

    print(f"MiniPdf PDFs:    {minipdf_dir}")
    print(f"Reference PDFs:  {reference_dir}")
    print(f"Report output:   {report_dir}")
    print()

    # Collect all test names from both directories
    names = set()
    for d in [minipdf_dir, reference_dir]:
        if os.path.isdir(d):
            for f in Path(d).glob("*.pdf"):
                names.add(f.stem)

    if not names:
        print("No PDF files found in either directory.")
        print("Run the following first:")
        print("  1. python generate_classic_xlsx.py       (generate test Excel files)")
        print("  2. dotnet run convert_xlsx_to_pdf.cs      (generate MiniPdf PDFs)")
        print("  3. python generate_reference_pdfs.py      (generate reference PDFs)")
        sys.exit(1)

    results = []
    for name in sorted(names):
        mp = os.path.join(minipdf_dir, f"{name}.pdf")
        rp = os.path.join(reference_dir, f"{name}.pdf")
        print(f"Comparing: {name} ...", end=" ")
        result = compare_single(mp, rp, images_dir, name)
        score = result.get("overall_score", "N/A")
        print(f"score={score}")
        results.append(result)

    generate_report(results, report_dir)

    # Print summary
    avg = sum(r.get("overall_score", 0) for r in results) / len(results) if results else 0
    print(f"\n{'='*60}")
    print(f"Overall Average Score: {avg:.4f}")
    print(f"{'='*60}")

    if avg < 0.7:
        print("⚠ Many test cases are significantly different from the reference.")
        print("  Check the report for details and improvement suggestions.")


if __name__ == "__main__":
    main()
