"""
Web Table Scraper
------------------
Paste a URL and automatically extract:
  - HTML tables
  - Product lists (name + price, best-effort heuristics)
  - Prices found anywhere on the page
  - Contacts (emails / phone numbers)

Exports everything to a single Excel workbook (one sheet per category).

Usage (CLI):
    python web_table_scraper.py --url "https://example.com/products" --out scraped_data.xlsx

Usage (as a library, e.g. from your PDF Extractor tool):
    from web_table_scraper import WebTableScraper

    scraper = WebTableScraper()
    data = scraper.scrape("https://example.com/products")
    scraper.export_report(data, "scraped_data.xlsx")

Install:
    pip install requests beautifulsoup4 pandas lxml openpyxl
"""

import re
import argparse
from io import StringIO

import requests
import pandas as pd
from bs4 import BeautifulSoup


EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(\+?\d{1,3}[\s.-]?)?(\(?\d{2,4}\)?[\s.-]?)?\d{3,4}[\s.-]?\d{3,4}")
PRICE_RE = re.compile(r"(?:[$€£₹]\s?\d[\d,]*\.?\d{0,2})|(?:\d[\d,]*\.?\d{0,2}\s?(?:USD|EUR|GBP|INR))")


class WebTableScraper:
    def __init__(self, timeout=15, user_agent=None):
        self.timeout = timeout
        self.headers = {
            "User-Agent": user_agent or (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            )
        }

    # ---- public API -----------------------------------------------------

    def scrape(self, url):
        """
        Returns a dict:
          {
            "tables": [DataFrame, DataFrame, ...],
            "products": [{"name":..., "price":...}, ...],
            "prices": ["$19.99", ...],
            "contacts": {"emails": [...], "phones": [...]},
          }
        """
        html = self._fetch(url)
        soup = BeautifulSoup(html, "lxml")

        return {
            "tables": self._extract_tables(html),
            "products": self._extract_products(soup),
            "prices": self._extract_prices(soup),
            "contacts": self._extract_contacts(soup),
        }

    def export_report(self, data, out_path="scraped_data.xlsx"):
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            # HTML tables - one sheet per table
            if data["tables"]:
                for i, df in enumerate(data["tables"], start=1):
                    df.to_excel(writer, sheet_name=f"Table_{i}"[:31], index=False)
            else:
                pd.DataFrame([["No tables found"]], columns=["info"]).to_excel(
                    writer, sheet_name="Tables", index=False
                )

            # Products
            products_df = pd.DataFrame(data["products"]) if data["products"] else pd.DataFrame(
                columns=["name", "price"]
            )
            products_df.to_excel(writer, sheet_name="Products", index=False)

            # Prices
            prices_df = pd.DataFrame({"price": data["prices"]}) if data["prices"] else pd.DataFrame(
                columns=["price"]
            )
            prices_df.to_excel(writer, sheet_name="Prices", index=False)

            # Contacts
            emails = data["contacts"]["emails"]
            phones = data["contacts"]["phones"]
            max_len = max(len(emails), len(phones), 1)
            contacts_df = pd.DataFrame({
                "email": emails + [""] * (max_len - len(emails)),
                "phone": phones + [""] * (max_len - len(phones)),
            })
            contacts_df.to_excel(writer, sheet_name="Contacts", index=False)

        return out_path

    # ---- internals --------------------------------------------------------

    def _fetch(self, url):
        resp = requests.get(url, headers=self.headers, timeout=self.timeout)
        resp.raise_for_status()
        return resp.text

    def _extract_tables(self, html):
        try:
            return pd.read_html(StringIO(html), flavor="lxml")
        except ValueError:
            # No <table> elements found on the page
            return []
        except ImportError:
            # lxml missing for some reason - fall back to default parser
            try:
                return pd.read_html(StringIO(html))
            except ValueError:
                return []

    def _extract_products(self, soup):
        """
        Heuristic product extraction: looks for common e-commerce markup patterns
        (class names containing 'product', 'item', 'card') and pairs a name with
        a nearby price. Works on many storefronts out of the box; site-specific
        selectors can be added for higher accuracy.
        """
        products = []
        candidates = soup.find_all(
            lambda tag: tag.name in ("div", "li", "article")
            and tag.get("class")
            and any(
                key in " ".join(tag.get("class")).lower()
                for key in ("product", "item", "card")
            )
        )

        for tag in candidates:
            text = tag.get_text(" ", strip=True)
            price_match = PRICE_RE.search(text)
            if not price_match:
                continue

            name_tag = tag.find(["h1", "h2", "h3", "h4", "a", "span"])
            name = name_tag.get_text(strip=True) if name_tag else text[:60]

            products.append({
                "name": name,
                "price": price_match.group(0),
            })

        # de-duplicate
        seen = set()
        unique_products = []
        for p in products:
            key = (p["name"], p["price"])
            if key not in seen:
                seen.add(key)
                unique_products.append(p)

        return unique_products

    def _extract_prices(self, soup):
        text = soup.get_text(" ", strip=True)
        return list(dict.fromkeys(PRICE_RE.findall(text)))  # de-dup, keep order

    def _extract_contacts(self, soup):
        text = soup.get_text(" ", strip=True)
        html_str = str(soup)

        emails = list(dict.fromkeys(EMAIL_RE.findall(text)))

        # also check mailto: / tel: links, which are more reliable than free text
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if href.startswith("mailto:"):
                emails.append(href.replace("mailto:", "").split("?")[0])

        phones_raw = PHONE_RE.findall(text)
        phones = list(dict.fromkeys(
            "".join(p) for p in phones_raw if len("".join(p).strip()) >= 7
        ))

        return {
            "emails": list(dict.fromkeys(emails)),
            "phones": phones,
        }


# --------------------------------------------------------------------------
# CLI entry point
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Scrape tables, products, prices, and contacts from a URL.")
    parser.add_argument("--url", required=True, help="URL to scrape.")
    parser.add_argument("--out", default="scraped_data.xlsx", help="Output Excel file path.")
    args = parser.parse_args()

    scraper = WebTableScraper()
    data = scraper.scrape(args.url)
    out_path = scraper.export_report(data, args.out)

    print(f"Found {len(data['tables'])} table(s), {len(data['products'])} product(s), "
          f"{len(data['prices'])} price(s), {len(data['contacts']['emails'])} email(s), "
          f"{len(data['contacts']['phones'])} phone(s).")
    print(f"Report saved to: {out_path}")


if __name__ == "__main__":
    main()
