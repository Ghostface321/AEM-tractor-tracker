import re
import io
from datetime import datetime
from urllib.parse import urljoin

import requests
import pdfplumber
import pandas as pd
from bs4 import BeautifulSoup

PAGE_URL = "https://www.aem.org/market-share-statistics/us-ag-tractor-and-combine-reports"
OUTPUT_FILE = "data.xlsx"


def get_latest_report_info():
    response = requests.get(PAGE_URL, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    title_el = soup.find(["h3", "h4"], string=re.compile(r"Tractor and Combine Report", re.I))
    if not title_el:
        raise RuntimeError("Report title not found on the AEM page.")

    link_el = soup.find("a", string=re.compile(r"Download Report", re.I))
    if not link_el or not link_el.get("href"):
        raise RuntimeError("Download link not found on the AEM page.")

    title = title_el.get_text(" ", strip=True)
    pdf_url = urljoin(PAGE_URL, link_el["href"])

    return title, pdf_url


def download_pdf(pdf_url: str) -> bytes:
    response = requests.get(pdf_url, timeout=60)
    response.raise_for_status()
    return response.content


def extract_text_from_pdf(pdf_bytes: bytes) -> str:
    text_parts = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")

    full_text = "\n".join(text_parts).strip()

    if not full_text:
        raise RuntimeError("No text could be extracted from the PDF.")

    return full_text


def parse_report(text: str, source_title: str, source_pdf_url: str) -> pd.DataFrame:
    month_match = re.search(
        r"AEM United States Ag Tractor and Combine Report\s+([A-Za-z]+\s+\d{4})",
        text,
    )
    report_month = month_match.group(1) if month_match else None

    release_match = re.search(r"Report Released\s+(\d{1,2}/\d{1,2}/\d{4})", text)
    release_date = release_match.group(1) if release_match else None

    lines = [line.strip() for line in text.splitlines() if line.strip()]

    valid_categories = {
        "2WD Farm Tractors",
        "< 40 HP",
        "40 < 100 HP",
        "100+ HP",
        "Total 2WD Farm Tractors",
        "4WD Farm Tractors",
        "Total Farm Tractors",
        "Self-Prop Combines",
    }

    row_pattern = re.compile(
        r"^(?P<category>.+?)\s+"
        r"(?P<current_month>[\d,]+)\s+"
        r"(?P<prior_year_month>[\d,]+)\s+"
        r"(?P<month_pct>-?[\d.]+)\s+"
        r"(?P<ytd_current>[\d,]+)\s+"
        r"(?P<ytd_prior>[\d,]+)\s+"
        r"(?P<ytd_pct>-?[\d.]+)\s+"
        r"(?P<inventory>[\d,]+)$"
    )

    rows = []

    for line in lines:
        match = row_pattern.match(line)
        if not match:
            continue

        category = match.group("category").strip()
        if category not in valid_categories:
            continue

        rows.append(
            {
                "report_month": report_month,
                "release_date": release_date,
                "category": category,
                "current_month_units": int(match.group("current_month").replace(",", "")),
                "prior_year_month_units": int(match.group("prior_year_month").replace(",", "")),
                "month_pct_change": float(match.group("month_pct")),
                "ytd_current_units": int(match.group("ytd_current").replace(",", "")),
                "ytd_prior_units": int(match.group("ytd_prior").replace(",", "")),
                "ytd_pct_change": float(match.group("ytd_pct")),
                "beginning_inventory": int(match.group("inventory").replace(",", "")),
                "source_title": source_title,
                "source_pdf_url": source_pdf_url,
                "scraped_at_utc": datetime.utcnow().isoformat(timespec="seconds"),
            }
        )

    if not rows:
        raise RuntimeError("No data rows could be parsed from the PDF.")

    return pd.DataFrame(rows)


def save_to_excel(df: pd.DataFrame):
    try:
        existing = pd.read_excel(OUTPUT_FILE)

        combined = pd.concat([existing, df], ignore_index=True)
        combined = combined.drop_duplicates(subset=["report_month", "category"], keep="last")
    except FileNotFoundError:
        combined = df.copy()
    except Exception:
        combined = df.copy()

    combined = combined.sort_values(["report_month", "category"]).reset_index(drop=True)
    combined.to_excel(OUTPUT_FILE, index=False)


def main():
    source_title, pdf_url = get_latest_report_info()
    pdf_bytes = download_pdf(pdf_url)
    text = extract_text_from_pdf(pdf_bytes)
    df = parse_report(text, source_title, pdf_url)
    save_to_excel(df)

    print("Done.")
    print(df)


if __name__ == "__main__":
    main()
