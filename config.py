"""
config.py
Central settings for the bond tracker. Every other script imports from here,
so if a folder or the model name changes, it only changes in one place.
"""
import os
from pathlib import Path

from dotenv import load_dotenv

# Read .env so secrets (API key, SEC contact details) are available
# without ever being written into the code
load_dotenv()

# ---------- Folders ----------
# Path(__file__).parent is the folder this file lives in (bond-tracker),
# so the paths work no matter where you run a script from.
ROOT = Path(__file__).parent
RAW_DIR = ROOT / "data" / "raw"                # downloaded announcements
EXTRACTED_DIR = ROOT / "data" / "extracted"    # one JSON per announcement
VALIDATION_DIR = ROOT / "data" / "validation"  # hand-checked answers
OUTPUT_DIR = ROOT / "output"
CHARTS_DIR = OUTPUT_DIR / "charts"
FETCH_LOG = ROOT / "data" / "fetched_ids.txt"  # every SEC filing already seen

# ---------- Claude ----------
MODEL = "claude-haiku-4-5-20251001"  # cheapest current Claude model
COST_LOG = ROOT / "data" / "cost_log.csv"      # tokens and cost per file

# Haiku 4.5 list prices in US dollars per million tokens.
# Used only to report what the project cost; check the Anthropic
# pricing page if you want to confirm they haven't changed.
PRICE_INPUT_PER_M = 1.00
PRICE_OUTPUT_PER_M = 5.00

# Term sheets are 1 to 3 pages. Some FWPs are 40-page structured-note
# brochures that aren't bonds at all. Sending only the first ~40,000
# characters (about 10,000 tokens) caps the cost of any one file while
# still covering every real term sheet in full.
MAX_CHARS = 40_000

# ---------- Dataset (Stage 3) ----------
DATASET_CSV = OUTPUT_DIR / "bond_dataset.csv"
DATASET_XLSX = OUTPUT_DIR / "bond_dataset.xlsx"

# FX rates to convert every deal into USD, so volumes in different
# currencies can be added up. Source: Federal Reserve H.10 release,
# noon buying rates for 18 September 2026 (USD per 1 unit of currency).
# A fixed snapshot is a simplification: good enough for comparing sector
# volumes, but it is not the exact USD value on each deal's own date.
FX_DATE = "2026-09-18"
FX_TO_USD = {
    "USD": 1.0,
    "EUR": 1.1464,
    "GBP": 1.3372,
    "CAD": 1 / 1.4008,   # H.10 quotes CAD per USD, so we invert it
    "JPY": 1 / 156.87,   # same for JPY
    "CHF": 1 / 0.8242,   # same for CHF
    "AUD": 0.7111,
}

# ---------- SEC EDGAR ----------
# The SEC requires automated requests to say who is making them
# (name + email), so they can contact you if a script misbehaves.
SEC_USER_AGENT = os.getenv("SEC_USER_AGENT")

# EDGAR full-text search: the same engine behind the search box on sec.gov
EDGAR_SEARCH_URL = "https://efts.sec.gov/LATEST/search-index"
# Where the actual filed documents live
EDGAR_ARCHIVE_URL = "https://www.sec.gov/Archives/edgar/data"

# The exact phrase we search for. Syndicate desks title the final deal
# terms a "Pricing Term Sheet", so this targets new bond deals and cuts
# out most of the other FWP noise (fund marketing, structured notes).
SEARCH_PHRASE = '"pricing term sheet"'

# SEC allows up to 10 requests per second; pausing 0.2s keeps us well under
SEC_PAUSE_SECONDS = 0.2
