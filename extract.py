"""
extract.py  (Stage 2)
Reads every announcement in data/raw, asks Claude to pull out the deal
terms, and saves one JSON file per announcement in data/extracted.

THE KEY IDEA
Claude only READS: it copies out what the document states, and returns
null for anything that isn't there. Anything that needs calculating
(tenor, IG vs HY, size in millions) is done later in Python, in Stage 3,
because a rule in code is consistent and auditable and a model is not.

HOW TO RUN
    python extract.py --limit 3   -> try it on 3 files first
    python extract.py             -> process everything not yet done

Files that already have a JSON in data/extracted are skipped, so
re-running never pays for the same file twice.
"""
import argparse
import csv
import json
from datetime import datetime

import pdfplumber
from anthropic import Anthropic

import config

# ----------------------------------------------------------------------
# 1. The prompt: tells Claude what world it's in and the rules to follow.
#    (Remember the test call that defined "credit spread" as an options
#    strategy? This is the fix: explicit context and definitions.)
# ----------------------------------------------------------------------
SYSTEM_PROMPT = """You are a capital markets analyst on a debt syndicate desk.
You are reading a document filed with the SEC, which is usually the final
pricing term sheet for a new bond issue. Extract the deal terms exactly as
the document states them.

RULES
1. Only record what the document states. If a field is not stated, use null.
   Never guess, estimate, or calculate a value that is not written down.
2. Some documents are not new bond pricings (for example preferred stock,
   equity, structured or index-linked notes, fund marketing). For those, set
   is_bond_pricing to false, explain why in one short sentence, and return
   an empty tranches list.
3. One term sheet often prices several bonds at once ("tranches"), e.g.
   a 5-year and a 10-year. Return one tranche object per bond.
4. Amounts: give the full number in the stated currency, e.g.
   "$500,000,000" -> 500000000 and "EUR 700 million" -> 700000000.
5. Coupon, yield and price are percentages as plain numbers:
   "4.625%" -> 4.625, "99.812% of principal" -> 99.812.
6. Spread is in basis points (1bp = 0.01%) over a benchmark government bond
   or swap rate. "T+85", "+85 bps over UST" and "85 basis points" -> 85.
   Record the benchmark separately (e.g. "UST 4.25% due 2035",
   "German Bund", "Mid-swaps"). If the deal is quoted only as a yield
   (common for high yield bonds), spread_bps is null.
7. Dates as YYYY-MM-DD. The trade date is when the bond was priced;
   the settlement date is when it is issued and paid for.
8. Ratings exactly as written, e.g. "Baa2", "BBB", "A-". Ignore outlook
   notes such as "(stable)" or "(S)".
9. For sector, pick the closest option from the allowed list based on the
   issuer's business. This is the only field that is a classification
   rather than a copied value.
10. Many bank bonds are callable, e.g. a "4NC3" can be repaid at par after
   3 years and matures after 4. Record that par call date as first_call_date.
   Almost every corporate bond also has a "make-whole" call; that does NOT
   count, so leave first_call_date null unless there is a par call date.
11. pricing_quote: copy the exact short line of text you took the spread or
   yield from, so a human can check it against the source."""

# ----------------------------------------------------------------------
# 2. The schema: the exact shape of JSON Claude must return.
#    We pass it as a "tool". Forcing Claude to call this tool means the
#    reply is always valid JSON with exactly these fields, never free text.
# ----------------------------------------------------------------------
def nullable(json_type, description):
    """Helper: a field that can hold a value OR null (for 'not stated')."""
    return {"type": [json_type, "null"], "description": description}


TRANCHE_SCHEMA = {
    "type": "object",
    "properties": {
        "currency": nullable("string", "ISO code: USD, EUR, GBP, CAD..."),
        "principal_amount": nullable("number", "Full face amount in the stated currency"),
        "coupon_pct": nullable("number", "Annual coupon rate in %"),
        "coupon_type": {"type": ["string", "null"],
                        "enum": ["fixed", "floating", "fixed-to-floating", "zero", None]},
        "maturity_date": nullable("string", "YYYY-MM-DD"),
        "first_call_date": nullable("string", "First date the issuer can redeem early "
                                    "at par (the 'par call' or fixed-to-floating reset "
                                    "date), YYYY-MM-DD. Null if no such call. Ignore "
                                    "make-whole calls."),
        "benchmark": nullable("string", "What the spread is measured against"),
        "spread_bps": nullable("number", "Spread over benchmark in basis points"),
        "yield_pct": nullable("number", "Re-offer yield / yield to maturity in %"),
        "issue_price_pct": nullable("number", "Price to public as % of face value"),
        "seniority": {"type": ["string", "null"],
                      "enum": ["senior secured", "senior unsecured", "senior non-preferred",
                               "subordinated", "junior subordinated", "covered", None]},
        "pricing_quote": nullable("string", "Exact text the spread/yield was read from"),
    },
    "required": ["currency", "principal_amount", "coupon_pct", "coupon_type",
                 "maturity_date", "first_call_date", "benchmark", "spread_bps", "yield_pct",
                 "issue_price_pct", "seniority", "pricing_quote"],
}

EXTRACTION_TOOL = {
    "name": "record_bond_deal",
    "description": "Record the terms of a new bond issue from a pricing document.",
    "input_schema": {
        "type": "object",
        "properties": {
            "is_bond_pricing": {"type": "boolean",
                                "description": "True only if this prices a new bond issue"},
            "not_bond_reason": nullable("string", "If false, one sentence why"),
            "issuer": nullable("string", "Legal name of the issuing entity"),
            "sector": {"type": ["string", "null"],
                       "enum": ["Financials", "Utilities", "Industrials", "Technology",
                                "Healthcare", "Consumer", "Energy", "Real Estate",
                                "Telecoms & Media", "Materials", "Government & Agency",
                                None]},
            "trade_date": nullable("string", "Pricing date, YYYY-MM-DD"),
            "settlement_date": nullable("string", "Issue/settlement date, YYYY-MM-DD"),
            "rating_moodys": nullable("string", "e.g. Baa2"),
            "rating_sp": nullable("string", "e.g. BBB"),
            "rating_fitch": nullable("string", "e.g. BBB+"),
            "bookrunners": {"type": "array", "items": {"type": "string"},
                            "description": "Joint book-running managers / lead banks. Empty if not stated."},
            "use_of_proceeds": nullable("string", "Short summary, only if stated"),
            "tranches": {"type": "array", "items": TRANCHE_SCHEMA},
        },
        "required": ["is_bond_pricing", "not_bond_reason", "issuer", "sector",
                     "trade_date", "settlement_date", "rating_moodys", "rating_sp",
                     "rating_fitch", "bookrunners", "use_of_proceeds", "tranches"],
    },
}


# ----------------------------------------------------------------------
# 3. Reading files
# ----------------------------------------------------------------------
def read_document(path):
    """Return the text of a .txt or .pdf file."""
    if path.suffix.lower() == ".pdf":
        with pdfplumber.open(path) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    return path.read_text(encoding="utf-8", errors="ignore")


def files_to_process():
    """Every announcement in data/raw that doesn't have a JSON yet."""
    done = {p.stem for p in config.EXTRACTED_DIR.glob("*.json")}
    candidates = sorted(list(config.RAW_DIR.glob("*.txt")) + list(config.RAW_DIR.glob("*.pdf")))
    return [p for p in candidates if p.stem not in done]


# ----------------------------------------------------------------------
# 4. Calling Claude
# ----------------------------------------------------------------------
def extract_one(client, text):
    """Send one document to Claude and return (extracted_data, usage)."""
    response = client.messages.create(
        model=config.MODEL,
        max_tokens=4000,              # room for several tranches; caps cost
        system=SYSTEM_PROMPT,
        tools=[EXTRACTION_TOOL],
        tool_choice={"type": "tool", "name": "record_bond_deal"},  # must use the schema
        messages=[{"role": "user",
                   "content": f"<document>\n{text[:config.MAX_CHARS]}\n</document>"}],
    )
    tool_call = next(block for block in response.content if block.type == "tool_use")
    return tool_call.input, response.usage


def cost_usd(usage):
    return (usage.input_tokens * config.PRICE_INPUT_PER_M
            + usage.output_tokens * config.PRICE_OUTPUT_PER_M) / 1_000_000


def log_cost(filename, usage):
    """Append one line per file to data/cost_log.csv, so the total project
    cost is a real measured number, not an estimate."""
    new_file = not config.COST_LOG.exists()
    with open(config.COST_LOG, "a", newline="") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["timestamp", "file", "model", "input_tokens",
                             "output_tokens", "cost_usd"])
        writer.writerow([datetime.now().isoformat(timespec="seconds"), filename,
                         config.MODEL, usage.input_tokens, usage.output_tokens,
                         f"{cost_usd(usage):.5f}"])


def main():
    parser = argparse.ArgumentParser(description="Extract bond terms with Claude")
    parser.add_argument("--limit", type=int, help="only process this many files")
    args = parser.parse_args()

    config.EXTRACTED_DIR.mkdir(parents=True, exist_ok=True)
    todo = files_to_process()
    if args.limit:
        todo = todo[:args.limit]
    print(f"{len(todo)} file(s) to process.")

    client = Anthropic()  # reads ANTHROPIC_API_KEY from the environment (.env)
    total_cost = 0.0
    bonds = not_bonds = errors = 0

    for path in todo:
        try:
            data, usage = extract_one(client, read_document(path))
        except Exception as err:
            # One bad file shouldn't stop the run. No JSON is written,
            # so it will be retried next time.
            print(f"  ERROR    {path.name}: {err}")
            errors += 1
            continue

        record = {
            "source_file": path.name,
            "model": config.MODEL,
            "extracted_at": datetime.now().isoformat(timespec="seconds"),
            "usage": {"input_tokens": usage.input_tokens,
                      "output_tokens": usage.output_tokens},
            "data": data,
        }
        out = config.EXTRACTED_DIR / f"{path.stem}.json"
        out.write_text(json.dumps(record, indent=2), encoding="utf-8")
        log_cost(path.name, usage)
        total_cost += cost_usd(usage)

        if data.get("is_bond_pricing"):
            bonds += 1
            print(f"  BOND     {path.name}: {len(data.get('tranches', []))} tranche(s)")
        else:
            not_bonds += 1
            print(f"  SKIPPED  {path.name}: {data.get('not_bond_reason')}")

    print(f"\nDone. Bond deals: {bonds} | Not bonds: {not_bonds} | Errors: {errors} | "
          f"Cost this run: ${total_cost:.4f}")


if __name__ == "__main__":
    main()
