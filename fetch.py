"""
fetch.py  (Stage 1)
Downloads new-issue bond pricing term sheets from SEC EDGAR into data/raw.

FINANCE CONTEXT
When a company sells bonds publicly in the US, the syndicate desk sends out
a "pricing term sheet" the moment the deal prices: size, coupon, maturity,
spread to Treasuries, ratings, bookrunners. The issuer files it with the SEC
as an "FWP" (free writing prospectus), which makes it public. This script
finds those filings for a date window and saves each one as a text file.

HOW TO RUN
    python fetch.py             -> last 7 days (the weekly run)
    python fetch.py --days 30   -> last 30 days (good for the first build)

Running it twice is safe: filings already seen are skipped.
"""
import argparse
import hashlib
import re
import sys
import time
from datetime import date, timedelta

import requests
from bs4 import BeautifulSoup

import config

SEPARATOR = "=" * 40  # splits our header lines from the document text


def search_edgar(start, end, headers):
    """Ask EDGAR for every FWP filed between two dates that contains our
    search phrase. Results come back 100 at a time, so we keep asking
    for the next page until we have them all ("pagination")."""
    hits = []
    offset = 0
    while True:
        params = {
            "q": config.SEARCH_PHRASE,
            "forms": "FWP",
            "dateRange": "custom",
            "startdt": start.isoformat(),
            "enddt": end.isoformat(),
            "from": offset,
        }
        response = requests.get(config.EDGAR_SEARCH_URL, params=params,
                                headers=headers, timeout=30)
        response.raise_for_status()  # stop with a clear error if the SEC refuses
        data = response.json()

        page = data["hits"]["hits"]
        hits.extend(page)
        total = data["hits"]["total"]["value"]
        offset += len(page)
        if not page or offset >= total:
            break
        time.sleep(config.SEC_PAUSE_SECONDS)
    return hits


def document_url(hit):
    """Build the web address of the filed document.
    EDGAR stores files at /data/<company id>/<filing id without dashes>/<file>.
    The search result's _id looks like '0001193125-26-386604:d131263dfwp.htm'
    i.e. '<filing id>:<file name>'."""
    filing_id, filename = hit["_id"].split(":", 1)
    cik = hit["_source"]["ciks"][0].lstrip("0")  # company id, leading zeros dropped
    return f"{config.EDGAR_ARCHIVE_URL}/{cik}/{filing_id.replace('-', '')}/{filename}"


def html_to_text(html_bytes):
    """Turn the filed web page into clean plain text.
    Term sheets are mostly tables ('Coupon: | 4.625%'), so we rebuild each
    table row as one line with ' | ' between cells. That keeps each label
    next to its value, which makes extraction far more reliable."""
    soup = BeautifulSoup(html_bytes, "html.parser")
    for tag in soup(["script", "style"]):
        tag.decompose()  # remove code/styling, keep only readable text

    for row in soup.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"])]
        cells = [c for c in cells if c]
        row.replace_with(soup.new_string(" | ".join(cells) + "\n"))

    lines = [re.sub(r"\s+", " ", line).strip() for line in soup.get_text("\n").splitlines()]
    return "\n".join(line for line in lines if line)


def issuer_name(hit):
    """'Parker-Hannifin Corp  (PH)  (CIK 0000076334)' -> 'Parker-Hannifin Corp'"""
    return re.split(r"\s+\(", hit["_source"]["display_names"][0])[0].strip()


def slugify(text):
    """'Parker-Hannifin Corp' -> 'parker-hannifin-corp' (safe for file names)"""
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")[:40]


def text_hash(text):
    """A fingerprint of the document text. Identical text = identical hash."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def load_existing_hashes():
    """Fingerprint every file already in data/raw, so we never save the
    same term sheet twice (see the co-issuer note in main)."""
    hashes = set()
    for path in config.RAW_DIR.glob("*.txt"):
        content = path.read_text(encoding="utf-8", errors="ignore")
        body = content.split(SEPARATOR, 1)[-1].strip()  # ignore our header lines
        hashes.add(text_hash(body))
    return hashes


def main():
    parser = argparse.ArgumentParser(description="Fetch bond pricing term sheets from SEC EDGAR")
    parser.add_argument("--days", type=int, default=7, help="how many days back to search")
    args = parser.parse_args()

    if not config.SEC_USER_AGENT:
        sys.exit("SEC_USER_AGENT is missing from .env (it should be your name and email).")
    headers = {"User-Agent": config.SEC_USER_AGENT}

    config.RAW_DIR.mkdir(parents=True, exist_ok=True)
    seen_ids = set()
    if config.FETCH_LOG.exists():
        seen_ids = set(config.FETCH_LOG.read_text().split())
    seen_hashes = load_existing_hashes()

    end = date.today()
    start = end - timedelta(days=args.days)
    print(f"Searching EDGAR for pricing term sheets filed {start} to {end}...")
    hits = search_edgar(start, end, headers)
    print(f"Found {len(hits)} filings.")

    saved = skipped = duplicates = failed = 0
    with open(config.FETCH_LOG, "a") as log:
        for hit in hits:
            filing_id = hit["_source"]["adsh"]
            if filing_id in seen_ids:
                skipped += 1
                continue

            url = document_url(hit)
            try:
                time.sleep(config.SEC_PAUSE_SECONDS)
                response = requests.get(url, headers=headers, timeout=30)
                response.raise_for_status()
            except requests.RequestException as err:
                # Not logged as seen, so it will be retried on the next run
                print(f"  FAILED  {issuer_name(hit)}: {err}")
                failed += 1
                continue

            body = html_to_text(response.content)
            fingerprint = text_hash(body)

            # FINANCE NOTE: a bond is often guaranteed by other companies in
            # the same group (e.g. a parent and its subsidiaries). Each of
            # those "co-registrants" files the SAME term sheet under its own
            # name, so one deal can appear 5 times. Same text = same deal,
            # so we keep the first copy only, or issuance volume would be
            # counted several times over.
            if fingerprint in seen_hashes:
                duplicates += 1
            else:
                name = issuer_name(hit)
                filed = hit["_source"]["file_date"]
                header = (f"SOURCE: {url}\n"
                          f"ISSUER (SEC filer): {name}\n"
                          f"FILED: {filed}\n"
                          f"{SEPARATOR}\n")
                out_path = config.RAW_DIR / f"{filed}_{slugify(name)}_{filing_id}.txt"
                out_path.write_text(header + body, encoding="utf-8")
                seen_hashes.add(fingerprint)
                saved += 1
                print(f"  saved   {out_path.name}")

            log.write(filing_id + "\n")  # remember it, so it's never downloaded again
            seen_ids.add(filing_id)

    print(f"\nDone. New: {saved} | Duplicate copies skipped: {duplicates} | "
          f"Already had: {skipped} | Failed: {failed}")


if __name__ == "__main__":
    main()
