"""
build_dataset.py  (Stage 3)
Turns the JSON extractions into one clean table: one row per bond (tranche),
saved as output/bond_dataset.csv and a formatted output/bond_dataset.xlsx.

THE KEY IDEA
Claude extracted what each document SAYS. This script works out everything
that needs a RULE: tenor, composite rating, investment grade vs high yield,
USD-equivalent size, and removal of duplicate bonds. Rules in code give the
same answer every time and can be audited line by line.

HOW TO RUN
    python build_dataset.py
It costs nothing (no API calls) and can be re-run as often as you like.
"""
import json
import re
from datetime import date

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config

# ----------------------------------------------------------------------
# RATINGS
# FINANCE: the three agencies use different labels for the same level of
# credit risk (Moody's "Baa3" = S&P/Fitch "BBB-"). Mapping them all onto
# one numeric scale (1 = AAA, best) lets us compare and sort them.
# The investment grade line sits between 10 (BBB-/Baa3) and 11 (BB+/Ba1).
# Many investors (pension funds, insurers) can only hold investment grade,
# which is why that line matters so much for pricing.
# ----------------------------------------------------------------------
SP_SCALE = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
            "BB+", "BB", "BB-", "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"]
MOODYS_SCALE = ["Aaa", "Aa1", "Aa2", "Aa3", "A1", "A2", "A3", "Baa1", "Baa2", "Baa3",
                "Ba1", "Ba2", "Ba3", "B1", "B2", "B3", "Caa1", "Caa2", "Caa3", "Ca", "C"]
SP_SCORE = {r: i + 1 for i, r in enumerate(SP_SCALE)}          # S&P and Fitch share labels
MOODYS_SCORE = {r: i + 1 for i, r in enumerate(MOODYS_SCALE)}
IG_CUTOFF = 10  # BBB- / Baa3


def rating_score(raw, scale):
    """'Baa2 (stable)' -> 9. Returns None if blank or not a real rating."""
    if not raw:
        return None
    text = str(raw).replace("(P)", "").strip()   # "(P)" = provisional rating
    match = re.match(r"([A-C][a-cA-C]{0,2}[1-3+-]?)", text)
    return scale.get(match.group(1)) if match else None


def composite_score(moodys, sp, fitch):
    """FINANCE: the bond index convention (used by Bloomberg's indices):
    take the MIDDLE of the three ratings; if only two exist, take the LOWER
    (more cautious); if only one, use it. Note a higher score = worse credit."""
    scores = sorted(s for s in (rating_score(moodys, MOODYS_SCORE),
                                rating_score(sp, SP_SCORE),
                                rating_score(fitch, SP_SCORE)) if s is not None)
    if not scores:
        return None
    if len(scores) == 3:
        return scores[1]
    return scores[-1]


# ----------------------------------------------------------------------
# TENOR
# FINANCE: tenor = years until the bond is repaid. For callable bank bonds
# (e.g. a 4NC3) the market prices to the CALL date, because the issuer is
# expected to repay then. Most corporate bonds also have a par call 1 to 6
# months before maturity; that barely changes the tenor, so we only use the
# call date when it's at least a year before maturity.
# ----------------------------------------------------------------------
TENOR_BUCKETS = [(0, 3.5, "1-3y"), (3.5, 6, "5y"), (6, 8.5, "7y"),
                 (8.5, 15, "10y"), (15, 25, "20y"), (25, 100, "30y+")]


def to_date(value):
    try:
        return date.fromisoformat(value) if value else None
    except ValueError:
        return None


def years_between(start, end):
    return round((end - start).days / 365.25, 2) if start and end else None


def bucket(years):
    if years is None:
        return None
    return next((label for lo, hi, label in TENOR_BUCKETS if lo <= years < hi), None)


# ----------------------------------------------------------------------
# BENCHMARK
# FINANCE: a spread only means something relative to its benchmark. USD
# deals price over US Treasuries, EUR deals usually over mid-swaps or
# German Bunds, GBP over gilts. We tag each so charts never mix them.
# ----------------------------------------------------------------------
def benchmark_type(text):
    """Tag the benchmark. Uses whole-word matching (the \\b in the patterns),
    because a plain substring check for "ust" also matches "AUGUST", which
    once tagged euro bonds priced over an August Bund as Treasury deals."""
    if not text:
        return None
    t = text.lower()
    if "swap" in t:
        return "Mid-swaps"
    if "sofr" in t:
        return "SOFR"
    if re.search(r"\b(bund|dbr|obl|bobl|german)", t):
        return "Bund"
    if re.search(r"\b(gilt|ukt)\b", t):
        return "Gilt"
    if re.search(r"\b(ust|treasury|treasuries|t)\b", t):
        return "UST"
    # Some USD term sheets name the benchmark bond without saying
    # "Treasury", e.g. "4.250% due August 15, 2029". A USD fixed-rate
    # benchmark written like that is a Treasury (checked by main() below,
    # which only applies this to USD bonds).
    if re.search(r"\d%\s+due\b", t):
        return "UST?"
    return "Other"


def usd_fix(tag, currency):
    """Resolve the 'UST?' case: USD -> UST, anything else -> Other."""
    if tag == "UST?":
        return "UST" if currency == "USD" else "Other"
    return tag


# ----------------------------------------------------------------------
# SANITY CHECKS: flag, don't delete. A flagged row stays in the dataset so
# a human can look at it; silently dropping data hides problems.
# ----------------------------------------------------------------------
def quality_flags(row):
    flags = []
    if row["coupon_pct"] is not None and not 0 <= row["coupon_pct"] <= 20:
        flags.append("coupon out of range")
    if row["spread_bps"] is not None and not 0 < row["spread_bps"] <= 1500:
        flags.append("spread out of range")
    if row["yield_pct"] is not None and not 0 < row["yield_pct"] <= 20:
        flags.append("yield out of range")
    if row["tenor_years"] is not None and row["tenor_years"] <= 0:
        flags.append("maturity before issue")
    if row["currency"] and row["currency"] not in config.FX_TO_USD:
        flags.append("no FX rate for currency")
    if row["size_m"] is None:
        flags.append("size missing")
    return "; ".join(flags) or None


def source_url(source_file):
    """Read the SEC link fetch.py wrote on the first line of the raw file."""
    path = config.RAW_DIR / source_file
    if path.suffix == ".txt" and path.exists():
        first = path.read_text(encoding="utf-8", errors="ignore").splitlines()[0]
        if first.startswith("SOURCE:"):
            return first.replace("SOURCE:", "").strip()
    return None


# ----------------------------------------------------------------------
# BUILD
# ----------------------------------------------------------------------
def load_rows():
    """Flatten every JSON into rows: one per tranche, with the deal-level
    fields (issuer, ratings, banks) copied onto each of its tranches."""
    rows, excluded = [], []
    for path in sorted(config.EXTRACTED_DIR.glob("*.json")):
        record = json.loads(path.read_text(encoding="utf-8"))
        d = record["data"]
        if not d.get("is_bond_pricing"):
            excluded.append({"source_file": record["source_file"],
                             "reason": d.get("not_bond_reason")})
            continue

        comp = composite_score(d.get("rating_moodys"), d.get("rating_sp"), d.get("rating_fitch"))
        for t in d.get("tranches", []):
            trade = to_date(d.get("trade_date"))
            start = to_date(d.get("settlement_date")) or trade
            maturity = to_date(t.get("maturity_date"))
            call = to_date(t.get("first_call_date"))

            tenor = years_between(start, maturity)
            to_call = years_between(start, call)
            use_call = (to_call is not None and tenor is not None and tenor - to_call >= 1)
            pricing_tenor = to_call if use_call else tenor
            # FINANCE: a perpetual (e.g. a bank AT1) has no maturity at all,
            # so the market can only price it to its first call date.
            if maturity is None and to_call is not None:
                pricing_tenor = to_call

            ccy = t.get("currency")
            size_m = t["principal_amount"] / 1e6 if t.get("principal_amount") else None
            fx = config.FX_TO_USD.get(ccy)

            row = {
                "issuer": d.get("issuer"),
                "sector": d.get("sector"),
                "trade_date": trade,
                "settlement_date": to_date(d.get("settlement_date")),
                "currency": ccy,
                "size_m": size_m,
                "size_usd_m": round(size_m * fx, 1) if size_m and fx else None,
                "coupon_pct": t.get("coupon_pct"),
                "coupon_type": t.get("coupon_type"),
                "maturity_date": maturity,
                "first_call_date": call,
                "tenor_years": tenor,
                "pricing_tenor_years": pricing_tenor,
                "tenor_bucket": bucket(pricing_tenor),
                "benchmark": t.get("benchmark"),
                "benchmark_type": usd_fix(benchmark_type(t.get("benchmark")), ccy),
                "spread_bps": t.get("spread_bps"),
                "yield_pct": t.get("yield_pct"),
                "issue_price_pct": t.get("issue_price_pct"),
                "rating_moodys": d.get("rating_moodys"),
                "rating_sp": d.get("rating_sp"),
                "rating_fitch": d.get("rating_fitch"),
                "composite_rating": SP_SCALE[comp - 1] if comp else None,
                "rating_score": comp,
                # FINANCE: unrated is its own bucket. Guessing IG or HY for
                # an unrated bond would be exactly the kind of made-up
                # number this project is designed to avoid.
                "grade": ("IG" if comp <= IG_CUTOFF else "HY") if comp else "Unrated",
                "seniority": t.get("seniority"),
                "bookrunners": "; ".join(d.get("bookrunners") or []) or None,
                "use_of_proceeds": d.get("use_of_proceeds"),
                "pricing_quote": t.get("pricing_quote"),
                "source_file": record["source_file"],
                "source_url": source_url(record["source_file"]),
            }
            row["flags"] = quality_flags(row)
            rows.append(row)
    return pd.DataFrame(rows), pd.DataFrame(excluded)


def remove_duplicates(df):
    """FINANCE: the same bond can reach us twice, e.g. filed by a parent and
    a finance subsidiary with slightly different wording (so fetch.py's
    exact-text check can't catch it). A bond is uniquely identified by its
    currency, size, coupon and maturity, so matching rows are one bond."""
    key = ["currency", "size_m", "coupon_pct", "maturity_date"]
    has_key = df[key].notna().all(axis=1)
    dupes = df[has_key].duplicated(subset=key, keep="first")
    dupe_index = dupes[dupes].index
    return df.drop(index=dupe_index), len(dupe_index)


# ----------------------------------------------------------------------
# EXCEL FORMATTING (openpyxl)
# ----------------------------------------------------------------------
NUMBER_FORMATS = {
    "size_m": "#,##0", "size_usd_m": "#,##0", "coupon_pct": "0.000",
    "tenor_years": "0.0", "pricing_tenor_years": "0.0", "spread_bps": "0.0",
    "yield_pct": "0.000", "issue_price_pct": "0.000",
    "trade_date": "dd-mmm-yy", "settlement_date": "dd-mmm-yy",
    "maturity_date": "dd-mmm-yy", "first_call_date": "dd-mmm-yy",
}


def format_sheet(ws, df):
    header_fill = PatternFill("solid", fgColor="1F3864")  # dark navy
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "B2"            # keep issuer column and header visible
    ws.auto_filter.ref = ws.dimensions

    for i, col in enumerate(df.columns, start=1):
        letter = get_column_letter(i)
        longest = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
        ws.column_dimensions[letter].width = min(max(longest + 2, 9), 45)
        fmt = NUMBER_FORMATS.get(col)
        if fmt:
            for cell in ws[letter][1:]:
                cell.number_format = fmt


def main():
    df, excluded = load_rows()
    if df.empty:
        print("No bond data found. Run extract.py first.")
        return
    df, n_dupes = remove_duplicates(df)
    df = df.sort_values(["trade_date", "issuer", "maturity_date"]).reset_index(drop=True)

    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(config.DATASET_CSV, index=False, encoding="utf-8-sig")  # -sig so Excel reads it correctly

    with pd.ExcelWriter(config.DATASET_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Bonds", index=False)
        format_sheet(writer.sheets["Bonds"], df)
        if not excluded.empty:
            excluded.to_excel(writer, sheet_name="Excluded", index=False)
            format_sheet(writer.sheets["Excluded"], excluded)
        notes = pd.DataFrame({"Note": [
            f"FX: converted to USD at Federal Reserve H.10 rates for {config.FX_DATE}.",
            "Composite rating: middle of Moody's/S&P/Fitch; lower of two; else the one available.",
            "IG = composite BBB-/Baa3 or better. Unrated bonds are not classified.",
            "Pricing tenor uses the par call date where it is 1+ year before maturity.",
            "Sector is classified by the model from the issuer's business; all other fields are copied from the filing.",
        ]})
        notes.to_excel(writer, sheet_name="Notes", index=False)
        writer.sheets["Notes"].column_dimensions["A"].width = 100

    flagged = df["flags"].notna().sum()
    print(f"Bonds: {len(df)} | Deals: {df['source_file'].nunique()} | "
          f"Duplicate bonds removed: {n_dupes} | Excluded documents: {len(excluded)} | "
          f"Rows flagged for review: {flagged}")
    print(f"Grade split: {df['grade'].value_counts().to_dict()}")
    print(f"Currencies: {df['currency'].value_counts().to_dict()}")
    print(f"Spread populated: {df['spread_bps'].notna().sum()} of {len(df)} bonds")
    if flagged:
        print("\nFlagged rows:")
        print(df.loc[df["flags"].notna(), ["issuer", "currency", "size_m", "flags"]].to_string())
    print(f"\nSaved {config.DATASET_CSV.name} and {config.DATASET_XLSX.name} in output/")


if __name__ == "__main__":
    main()
