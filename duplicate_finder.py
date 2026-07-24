r"""
Duplicate Finder
----------------
Scans a file, a folder, or a mixed list of files and finds duplicates among:
  - Images        (exact hash + perceptual/"AI similarity" hash)
  - PDFs          (exact hash + text-similarity for near-duplicates)
  - Excel rows    (exact + fuzzy row matching, across one or many workbooks)
  - Documents     (.docx/.txt - text hash + similarity)

Usage (CLI):
    python duplicate_finder.py --path "D:\Data" --out duplicates_report.xlsx
    python duplicate_finder.py --path "D:\Data\file1.pdf" "D:\Data\file2.pdf"

Usage (as a library, e.g. from your PDF Extractor tool):
    from duplicate_finder import DuplicateFinder

    finder = DuplicateFinder(similarity_threshold=90)
    report = finder.scan(["D:/Data"])           # folder(s) and/or file(s)
    finder.export_report(report, "duplicates_report.xlsx")

Install:
    pip install pillow imagehash pdfplumber openpyxl pandas python-docx
"""

import os
import sys
import hashlib
import argparse
from collections import defaultdict

import pandas as pd

# Optional deps - imported lazily / guarded so the tool degrades gracefully
try:
    from PIL import Image
    import imagehash
    HAS_IMAGE_LIBS = True
except ImportError:
    HAS_IMAGE_LIBS = False

try:
    import pdfplumber
    HAS_PDF_LIB = True
except ImportError:
    HAS_PDF_LIB = False

try:
    import docx  # python-docx
    HAS_DOCX_LIB = True
except ImportError:
    HAS_DOCX_LIB = False


IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff"}
PDF_EXT = {".pdf"}
EXCEL_EXT = {".xlsx", ".xls", ".csv"}
DOC_EXT = {".docx", ".txt"}


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _sha256_of_file(path, chunk_size=1 << 20):
    """Exact-duplicate hash for any binary file."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(chunk_size):
            h.update(chunk)
    return h.hexdigest()


def _collect_files(paths):
    """Expand a mix of files/folders into a flat file list."""
    all_files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, files in os.walk(p):
                for name in files:
                    all_files.append(os.path.join(root, name))
        elif os.path.isfile(p):
            all_files.append(p)
    return all_files


def _bucket_by_type(files):
    buckets = defaultdict(list)
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext in IMAGE_EXT:
            buckets["images"].append(f)
        elif ext in PDF_EXT:
            buckets["pdfs"].append(f)
        elif ext in EXCEL_EXT:
            buckets["excel"].append(f)
        elif ext in DOC_EXT:
            buckets["docs"].append(f)
    return buckets


def _text_similarity(a, b):
    """Cheap similarity score (0-100) between two text blobs."""
    import difflib
    if not a or not b:
        return 0
    return round(difflib.SequenceMatcher(None, a, b).ratio() * 100, 1)


def _extract_pdf_text(path, max_pages=5):
    if not HAS_PDF_LIB:
        return ""
    try:
        text = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages[:max_pages]:
                text.append(page.extract_text() or "")
        return "\n".join(text)
    except Exception:
        return ""


def _extract_docx_text(path):
    if path.lower().endswith(".txt"):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
        except Exception:
            return ""
    if not HAS_DOCX_LIB:
        return ""
    try:
        d = docx.Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    except Exception:
        return ""


# --------------------------------------------------------------------------
# Main class
# --------------------------------------------------------------------------

class DuplicateFinder:
    def __init__(self, similarity_threshold=90, image_hash_distance=5):
        """
        similarity_threshold: 0-100, minimum % match to flag documents/rows as near-duplicate
        image_hash_distance:  max perceptual-hash distance to flag images as near-duplicate
                              (0 = identical, higher = more different; 5 is a good default)
        """
        self.similarity_threshold = similarity_threshold
        self.image_hash_distance = image_hash_distance

    # ---- public API -----------------------------------------------------

    def scan(self, paths):
        """
        paths: list of file paths and/or folder paths.
        Returns a dict report: {"images": [...], "pdfs": [...], "excel_rows": [...], "docs": [...]}
        """
        files = _collect_files(paths)
        buckets = _bucket_by_type(files)

        report = {
            "images": self._find_image_duplicates(buckets["images"]),
            "pdfs": self._find_pdf_duplicates(buckets["pdfs"]),
            "excel_rows": self._find_excel_row_duplicates(buckets["excel"]),
            "docs": self._find_doc_duplicates(buckets["docs"]),
        }
        return report

    def export_report(self, report, out_path="duplicates_report.xlsx"):
        """Writes each duplicate category to its own sheet in one Excel file."""
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            for sheet_name, rows in report.items():
                df = pd.DataFrame(rows) if rows else pd.DataFrame(
                    columns=["info"], data=[["No duplicates found"]]
                )
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        return out_path

    # ---- images -----------------------------------------------------------

    def _find_image_duplicates(self, image_files):
        results = []
        if not image_files:
            return results

        exact = defaultdict(list)
        for f in image_files:
            try:
                exact[_sha256_of_file(f)].append(f)
            except Exception:
                continue
        for hash_val, group in exact.items():
            if len(group) > 1:
                for f in group:
                    results.append({"type": "exact", "hash": hash_val, "file": f})

        if HAS_IMAGE_LIBS:
            phashes = {}
            for f in image_files:
                try:
                    phashes[f] = imagehash.average_hash(Image.open(f))
                except Exception:
                    continue

            seen = set()
            files_list = list(phashes.keys())
            for i in range(len(files_list)):
                for j in range(i + 1, len(files_list)):
                    f1, f2 = files_list[i], files_list[j]
                    if f1 in seen and f2 in seen:
                        continue
                    distance = phashes[f1] - phashes[f2]
                    if distance <= self.image_hash_distance:
                        results.append({
                            "type": "similar",
                            "distance": distance,
                            "file_a": f1,
                            "file_b": f2,
                        })
                        seen.update([f1, f2])
        else:
            print("[duplicate_finder] pillow/imagehash not installed - "
                  "skipping perceptual similarity check for images.")

        return results

    # ---- pdfs ---------------------------------------------------------------

    def _find_pdf_duplicates(self, pdf_files):
        results = []
        if not pdf_files:
            return results

        exact = defaultdict(list)
        for f in pdf_files:
            try:
                exact[_sha256_of_file(f)].append(f)
            except Exception:
                continue
        for hash_val, group in exact.items():
            if len(group) > 1:
                for f in group:
                    results.append({"type": "exact", "hash": hash_val, "file": f})

        texts = {f: _extract_pdf_text(f) for f in pdf_files}
        checked = set()
        files_list = list(texts.keys())
        for i in range(len(files_list)):
            for j in range(i + 1, len(files_list)):
                f1, f2 = files_list[i], files_list[j]
                if (f1, f2) in checked:
                    continue
                score = _text_similarity(texts[f1], texts[f2])
                if score >= self.similarity_threshold:
                    results.append({
                        "type": "similar",
                        "similarity_pct": score,
                        "file_a": f1,
                        "file_b": f2,
                    })
                checked.add((f1, f2))

        return results

    # ---- excel rows -----------------------------------------------------

    def _find_excel_row_duplicates(self, excel_files):
        results = []
        if not excel_files:
            return results

        row_map = defaultdict(list)  # normalized row string -> [(file, sheet, row_idx)]

        for f in excel_files:
            try:
                if f.lower().endswith(".csv"):
                    sheets = {"Sheet1": pd.read_csv(f, dtype=str)}
                else:
                    sheets = pd.read_excel(f, sheet_name=None, dtype=str)
            except Exception:
                continue

            for sheet_name, df in sheets.items():
                df = df.fillna("")
                for idx, row in df.iterrows():
                    key = "|".join(str(v).strip().lower() for v in row.values)
                    row_map[key].append((f, sheet_name, idx + 2))  # +2 ~ human row incl header

        for key, locations in row_map.items():
            if len(locations) > 1:
                for f, sheet, row_idx in locations:
                    results.append({
                        "file": f,
                        "sheet": sheet,
                        "row": row_idx,
                        "duplicate_count": len(locations),
                    })

        return results

    # ---- documents --------------------------------------------------------

    def _find_doc_duplicates(self, doc_files):
        results = []
        if not doc_files:
            return results

        exact = defaultdict(list)
        for f in doc_files:
            try:
                exact[_sha256_of_file(f)].append(f)
            except Exception:
                continue
        for hash_val, group in exact.items():
            if len(group) > 1:
                for f in group:
                    results.append({"type": "exact", "hash": hash_val, "file": f})

        texts = {f: _extract_docx_text(f) for f in doc_files}
        checked = set()
        files_list = list(texts.keys())
        for i in range(len(files_list)):
            for j in range(i + 1, len(files_list)):
                f1, f2 = files_list[i], files_list[j]
                if (f1, f2) in checked:
                    continue
                score = _text_similarity(texts[f1], texts[f2])
                if score >= self.similarity_threshold:
                    results.append({
                        "type": "similar",
                        "similarity_pct": score,
                        "file_a": f1,
                        "file_b": f2,
                    })
                checked.add((f1, f2))

        return results


# --------------------------------------------------------------------------
# CLI entry point
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Find duplicate images, PDFs, Excel rows, and documents.")
    parser.add_argument("--path", nargs="+", required=True,
                         help="One or more file paths and/or folder paths to scan.")
    parser.add_argument("--out", default="duplicates_report.xlsx",
                         help="Output Excel report path.")
    parser.add_argument("--threshold", type=int, default=90,
                         help="Similarity %% threshold for near-duplicate text/docs (default 90).")
    args = parser.parse_args()

    finder = DuplicateFinder(similarity_threshold=args.threshold)
    report = finder.scan(args.path)
    out_path = finder.export_report(report, args.out)

    total = sum(len(v) for v in report.values())
    print(f"Done. Found {total} duplicate entries across categories.")
    print(f"Report saved to: {out_path}")


if __name__ == "__main__":
    main()
