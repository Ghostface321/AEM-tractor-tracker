import os
import re
import io
import json
from datetime import datetime
from urllib.parse import urljoin

import requests
import pdfplumber
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2.service_account import Credentials

PAGE_URL = "https://www.aem.org/market-share-statistics/us-ag-tractor-and-combine-reports"

def get_latest_report_info():
    r = requests.get(PAGE_URL, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    title_el = soup.find(["h3", "h4"], string=re.compile(r"Tractor and Combine Report", re.I))
    if not title_el:
        raise RuntimeError("Report title not found.")

    link_el = soup.find("a", string=re.compile(r"Download Report", re.I))
    if not link_el or not link_el.get("href"):
        raise RuntimeError("Download link not found.")

    title = title_el.get_text(" ", strip=True)
    pdf_url = urljoin(PAGE_URL, link_el["href"])
    return title, pdf_url

def download_pdf(pdf_url: str) -> bytes:
    r = requests.get(pdf_url, timeout=60)
    r.raise_for_status()
    return r.content

def extract_text(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
    return "\n".join(parts)

def parse_report(text: str, source_title: str, source_pdf_url: str) -> pd.DataFrame:
    month_match = re.search(
        r"AEM United States Ag Tractor and Combine Report\s+([A-Za-z]+\s+\d{4})",
        text
    )
    report_month = month_match.group(1) if month_match else None

    release_match = re.search(r"Report Released\s+(\d{1,2}/\d{1,2}/\d{4})", text)
    release_date = release_match.group(1) if release_match else None

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

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

    pattern = re.compile(
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
        m = pattern.match(line)
        if not m:
            continue

        category = m.group("category").strip()
        if category not in valid_categories:
            continue

        rows.append({
            "report_month": report_month,
            "release_date": release_date,
            "category": category,
            "current_month_units": int(m.group("current_month").replace(",", "")),
            "prior_year_month_units": int(m.group("prior_year_month").replace(",", "")),
            "month_pct_change": float(m.group("month_pct")),
            "ytd_current_units": int(m.group("ytd_current").replace(",", "")),
            "ytd_prior_units": int(m.group("ytd_prior").replace(",", "")),
            "ytd_pct_change": float(m.group("ytd_pct")),
            "beginning_inventory": int(m.group("inventory").replace(",", "")),
            "source_title": source_title,
            "source_pdf_url": source_pdf_url,
            "scraped_at_utc": datetime.utcnow().isoformat(timespec="seconds"),
        })

    if not rows:
        raise RuntimeError("No rows parsed from PDF.")

    return pd.DataFrame(rows)

def get_gspread_client():
    raw_json = os.environ["GCP_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw_json)

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

def write_to_google_sheets(df: pd.DataFrame):
    sheet_name = os.environ["GOOGLE_SHEET_NAME"]
    worksheet_name = os.environ.get("WORKSHEET_NAME", "data")

    gc = get_gspread_client()
    sh = gc.open(sheet_name)

    try:
        ws = sh.worksheet(worksheet_name)
    except:
        ws = sh.add_worksheet(title=worksheet_name, rows=1000, cols=30)

    existing = pd.DataFrame(ws.get_all_records())
    if not existing.empty:
        combined = pd.concat([existing, df], ignore_index=True)
        combined = combined.drop_duplicates(subset=["report_month", "category"], keep="last")
    else:
        combined = df.copy()

    ws.clear()
    set_with_dataframe(ws, combined.sort_values(["report_month", "category"]).reset_index(drop=True))

def main():
    title, pdf_url = get_latest_report_info()
    pdf_bytes = download_pdf(pdf_url)
    text = extract_text(pdf_bytes)
    df = parse_report(text, title, pdf_url)
    write_to_google_sheets(df)
    print(df)

if __name__ == "__main__":
    main()
