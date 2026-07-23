r"""
PDF Extractor - Feature Router
------------------------------
Single entry point that wires the two new features into your existing
PDF Extractor tool. Accepts any mix of:
  --files   individual file(s): .pdf, .xlsx/.xls/.csv, images, .docx/.txt
  --folder  a folder to scan recursively
  --url     a webpage to scrape (Web Table Scraper feature)

Examples:
    # Find duplicates across a folder (images, pdfs, excel, docs)
    python app.py duplicates --folder "D:\Data"

    # Find duplicates across specific files
    python app.py duplicates --files "D:\Data\a.pdf" "D:\Data\b.pdf"

    # Scrape a URL for tables/products/prices/contacts
    python app.py scrape --url "https://example.com/products"

Install everything at once:
    pip install pillow imagehash pdfplumber openpyxl pandas python-docx requests beautifulsoup4 lxml
"""

import argparse

from duplicate_finder import DuplicateFinder
from web_table_scraper import WebTableScraper


def run_duplicates(args):
    inputs = []
    if args.folder:
        inputs.append(args.folder)
    if args.files:
        inputs.extend(args.files)

    if not inputs:
        print("Provide --folder and/or --files to scan.")
        return

    finder = DuplicateFinder(similarity_threshold=args.threshold)
    report = finder.scan(inputs)
    out_path = finder.export_report(report, args.out)

    total = sum(len(v) for v in report.values())
    print(f"Duplicate scan complete. {total} duplicate entries found.")
    print(f"Report: {out_path}")


def run_scrape(args):
    if not args.url:
        print("Provide --url to scrape.")
        return

    scraper = WebTableScraper()
    data = scraper.scrape(args.url)
    out_path = scraper.export_report(data, args.out)

    print(
        f"Scrape complete: {len(data['tables'])} table(s), "
        f"{len(data['products'])} product(s), {len(data['prices'])} price(s), "
        f"{len(data['contacts']['emails'])} email(s), {len(data['contacts']['phones'])} phone(s)."
    )
    print(f"Report: {out_path}")


def main():
    parser = argparse.ArgumentParser(description="PDF Extractor - Duplicate Finder & Web Table Scraper")
    sub = parser.add_subparsers(dest="command", required=True)

    dup = sub.add_parser("duplicates", help="Find duplicate images, PDFs, Excel rows, and documents.")
    dup.add_argument("--folder", help="Folder to scan recursively.")
    dup.add_argument("--files", nargs="+", help="Specific files to include in the scan.")
    dup.add_argument("--threshold", type=int, default=90, help="Similarity %% for near-duplicate text (default 90).")
    dup.add_argument("--out", default="duplicates_report.xlsx", help="Output Excel report path.")
    dup.set_defaults(func=run_duplicates)

    scrape = sub.add_parser("scrape", help="Scrape a URL for tables, products, prices, and contacts.")
    scrape.add_argument("--url", required=True, help="URL to scrape.")
    scrape.add_argument("--out", default="scraped_data.xlsx", help="Output Excel report path.")
    scrape.set_defaults(func=run_scrape)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
