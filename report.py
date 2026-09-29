"""
report.py  (Stage 6)
Builds the one-page "Weekly New Issue Market Update" as Word and PDF.

THE KEY IDEA
Every number in the note comes from Python (the dataset). Claude only
writes the words around them, from a fact sheet we give it. Afterwards a
check confirms every number Claude wrote appears in that fact sheet, so
the model can't slip an invented spread into a document styled as a desk
note.

HOW TO RUN
    python report.py                          -> week ending on the latest deal
    python report.py --week-ending 2026-09-18 -> a specific week
Output: output/reports/weekly_update_<week-end date>.docx and .pdf
(The PDF step uses Microsoft Word, via docx2pdf.)
"""
import argparse
import json
import re
from datetime import timedelta

import pandas as pd
from anthropic import Anthropic
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import config
from analyse import comparable_spreads, load

AUTHOR = "Nick Dugdale"
NAVY, GREY, LIGHT = RGBColor(0x1F, 0x38, 0x64), RGBColor(0x52, 0x51, 0x4E), "EEF2F8"
# Each week gets its own dated file, so older updates aren't overwritten
REPORTS_DIR = config.OUTPUT_DIR / "reports"


# ======================================================================
# 1. THE FACTS (all computed in Python)
# ======================================================================
def senior_ig(d):
    """Senior investment grade bonds over Treasuries: the cleanest read of
    where the market is pricing, without hybrids or other benchmarks."""
    c = comparable_spreads(d)
    return c[(c["grade"] == "IG") & ~c["seniority"].fillna("").str.contains("subordinated")]


def deal_table(d):
    """Collapse bonds into deals: one line per issuer per pricing day."""
    rows = []
    for _, g in d.groupby("deal_id"):
        tenors = sorted(g["pricing_tenor_years"].dropna().round().astype(int).unique())
        spreads = g.loc[g["benchmark_type"].isin(["UST", "Bund", "Gilt", "Mid-swaps"]), "spread_bps"].dropna()
        rows.append({
            "issuer": short_name(g["issuer"].iloc[0]),
            "trade_date": g["trade_date"].iloc[0].strftime("%d %b"),
            "rating": g["composite_rating"].iloc[0] if pd.notna(g["composite_rating"].iloc[0]) else "NR",
            "currency": "/".join(sorted(g["currency"].dropna().unique())),
            "size_usd_m": round(g["size_usd_m"].sum()),
            "bonds": len(g),
            "tenors": ", ".join(f"{t}y" for t in tenors) or "-",
            "spread_range": (f"+{spreads.min():.0f}" if spreads.min() == spreads.max()
                             else f"+{spreads.min():.0f} to +{spreads.max():.0f}") if len(spreads) else "-",
            "hybrid_or_sub": bool(g["seniority"].fillna("").str.contains("subordinated").any()),
            "sector": g["sector"].iloc[0],
        })
    return pd.DataFrame(rows).sort_values("size_usd_m", ascending=False)


def short_name(name):
    """'Aon North America, Inc. and Aon Global Holdings plc' -> 'Aon North America'"""
    name = re.split(r"\s+and\s+|,", name)[0]
    return re.sub(r"\b(Inc|Corporation|Corp|plc|LLC|L\.L\.C|Ltd|Company|Co|S\.A|N\.V)\b\.?", "", name).strip(" .,")


MIN_BONDS = 3   # fewer bonds than this in a tenor bucket = no meaningful median


def tenor_comparison(week, before):
    """FINANCE: compare this week's new-issue spreads with the prior period,
    tenor by tenor, but ONLY where both periods have enough bonds. One BBB
    30-year bond is a single print, not "the 30y part of the curve".
    The direction (tighter/wider) is worked out here in Python, so the
    model never has to do the comparison itself."""
    out = {}
    order = ["1-3y", "5y", "7y", "10y", "20y", "30y+"]
    for b in order:
        w = week.loc[week["tenor_bucket"] == b, "spread_bps"]
        p = before.loc[before["tenor_bucket"] == b, "spread_bps"]
        if w.empty:
            continue
        row = {"bonds_this_week": int(len(w)), "median_this_week_bps": round(float(w.median())),
               "bonds_prior": int(len(p)),
               "median_prior_bps": round(float(p.median())) if len(p) else None}
        if len(w) >= MIN_BONDS and len(p) >= MIN_BONDS:
            diff = row["median_this_week_bps"] - row["median_prior_bps"]
            row["vs_prior"] = "tighter" if diff < -2 else "wider" if diff > 2 else "unchanged"
        else:
            row["vs_prior"] = "too few bonds to compare"
        out[b] = row
    return out


def build_facts(df, week_end):
    week_start = week_end - timedelta(days=6)
    week = df[(df["trade_date"] >= week_start) & (df["trade_date"] <= week_end)]
    before = df[df["trade_date"] < week_start]
    if week.empty:
        raise SystemExit(f"No deals priced between {week_start.date()} and {week_end.date()}.")

    n_prior_weeks = max(1, round((week_start - df["trade_date"].min()).days / 7))
    deals = deal_table(week)
    ig_w, ig_b = senior_ig(week), senior_ig(before)
    curve = tenor_comparison(ig_w, ig_b)

    return {
        "week_start": week_start.strftime("%d %b %Y"),
        "week_end": week_end.strftime("%d %b %Y"),
        "week": {
            "volume_usd_bn": round(week["size_usd_m"].sum() / 1000, 1),
            "deals": int(deals.shape[0]),
            "bonds": int(len(week)),
            "ig_share_pct": round(100 * week.loc[week["grade"] == "IG", "size_usd_m"].sum()
                                  / week["size_usd_m"].sum()),
            "currencies": sorted(week["currency"].dropna().unique().tolist()),
            "hybrid_or_subordinated_deals": int(deals["hybrid_or_sub"].sum()),
            "sector_volume_usd_bn": (week.groupby("sector")["size_usd_m"].sum() / 1000)
                                    .round(1).sort_values(ascending=False).to_dict(),
        },
        "senior_ig_spreads_by_tenor": curve,
        "prior_period": {
            "from": df["trade_date"].min().strftime("%d %b %Y"),
            "weeks": n_prior_weeks,
            "avg_weekly_volume_usd_bn": round(before["size_usd_m"].sum() / 1000 / n_prior_weeks, 1)
                                        if len(before) else None,
        },
        "notable_deals": deals.head(6).to_dict("records"),
    }


# ======================================================================
# 2. THE WORDS (Claude, constrained to the facts)
# ======================================================================
COMMENTARY_TOOL = {
    "name": "write_update",
    "description": "Commentary for a weekly new-issue market update.",
    "input_schema": {
        "type": "object",
        "properties": {
            "summary": {"type": "string",
                        "description": "2-3 sentences, max 70 words: the week in one paragraph."},
            "themes": {"type": "array", "items": {"type": "string"}, "minItems": 3, "maxItems": 3,
                       "description": "Three market themes, max 30 words each."},
        },
        "required": ["summary", "themes"],
    },
}

SYSTEM = """You are a debt capital markets analyst writing the weekly new-issue
update that goes to the syndicate desk. Style: concise, factual, market
shorthand (e.g. "IG", "bps", "10y", "priced at T+85"). No hype, no advice.

STRICT RULES
1. Use ONLY numbers that appear in the fact sheet, exactly as given. Do not
   calculate new numbers (no differences, percentages or totals of your own).
   If something isn't in the fact sheet, don't mention it.
2. Spread moves: describe a tenor as tighter/wider ONLY using its "vs_prior"
   value. If it says "too few bonds to compare", don't claim a move for that
   tenor. Never generalise to "across the curve" unless every comparable
   tenor moved the same way.
3. These are CREDIT SPREADS, not interest rates. Never use yield-curve terms
   such as steepening, flattening, bull or bear.
4. Don't contradict yourself between the summary and the themes.
5. Write amounts as $7.1bn, spreads as +90bp."""


def numbers_in(text):
    return set(re.findall(r"\d+(?:\.\d+)?", text.replace(",", "")))


def allowed_numbers(facts):
    """Every number in the fact sheet, plus rounded versions of it, plus the
    same amount in billions (sizes are in $m, so 3,750 may be written as
    $3.75bn: a unit change, not a new number)."""
    allowed = set()
    for n in numbers_in(json.dumps(facts)):
        f = float(n)
        for v in (f, f / 1000):
            allowed.update({f"{v:.0f}", f"{v:.1f}", f"{v:.2f}", f"{v:g}"})
        allowed.add(n)
    return allowed


def write_commentary(facts):
    client = Anthropic()
    allowed = allowed_numbers(facts)
    for attempt in (1, 2):   # one retry if it uses a number it wasn't given
        response = client.messages.create(
            model=config.MODEL, max_tokens=800, system=SYSTEM,
            tools=[COMMENTARY_TOOL], tool_choice={"type": "tool", "name": "write_update"},
            messages=[{"role": "user", "content": "Fact sheet:\n" + json.dumps(facts, indent=1)}],
        )
        out = next(b for b in response.content if b.type == "tool_use").input
        text = out["summary"] + " " + " ".join(out["themes"])
        unverified = sorted(numbers_in(text) - allowed)
        if not unverified:
            print(f"Commentary check: every number traced to the fact sheet (attempt {attempt}).")
            return out
        print(f"Attempt {attempt}: numbers not in the fact sheet: {unverified}. "
              + ("Retrying..." if attempt == 1 else "Using it, but CHECK these before sending."))
    return out


# ======================================================================
# 3. THE DOCUMENT (python-docx)
# ======================================================================
def shade(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tc_pr.append(shd)


def no_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tbl_pr.append(borders)


def para(container, text="", size=9, bold=False, color=None, space_after=2, align=None):
    p = container.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    if text:
        run = p.add_run(text)
        run.font.size, run.bold = Pt(size), bold
        if color:
            run.font.color.rgb = color
    return p


def heading(doc, text):
    p = para(doc, text.upper(), size=8.5, bold=True, color=NAVY, space_after=2)
    p.paragraph_format.space_before = Pt(5)


def build_docx(facts, words, results, docx_out):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)          # A4
    sec.top_margin = sec.bottom_margin = Cm(1.3)
    sec.left_margin = sec.right_margin = Cm(1.6)
    style = doc.styles["Normal"]
    style.font.name, style.font.size = "Arial", Pt(9)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")

    # --- Title block
    para(doc, "Weekly New Issue Market Update", size=17, bold=True, color=NAVY, space_after=0)
    para(doc, f"Week ending {facts['week_end']}  |  SEC-registered corporate bond issuance  |  "
              f"Prepared by {AUTHOR}", size=8.5, color=GREY, space_after=6)

    # --- KPI strip
    w, p = facts["week"], facts["prior_period"]
    curve = facts["senior_ig_spreads_by_tenor"]
    ten = curve.get("10y")
    kpis = [
        (f"${w['volume_usd_bn']}bn", "Volume this week"
         + (f" (avg ${p['avg_weekly_volume_usd_bn']}bn)" if p["avg_weekly_volume_usd_bn"] else "")),
        (f"{w['deals']} / {w['bonds']}", "Deals / bonds priced"),
        (f"{w['ig_share_pct']}%", "Investment grade share"),
        (f"+{ten['median_this_week_bps']}bp" if ten else "-",
         (f"Median senior IG 10y spread, {ten['bonds_this_week']} bond(s)"
          + (f" (prior +{ten['median_prior_bps']}bp)" if ten["median_prior_bps"] else ""))
         if ten else "No senior IG 10y bonds this week"),
    ]
    t = doc.add_table(rows=2, cols=4)
    no_borders(t)
    for i, (big, small) in enumerate(kpis):
        for r in (0, 1):
            shade(t.cell(r, i), LIGHT)
        c0 = t.cell(0, i).paragraphs[0]
        run = c0.add_run(big)
        run.font.size, run.bold, run.font.color.rgb = Pt(15), True, NAVY
        c1 = t.cell(1, i).paragraphs[0]
        run = c1.add_run(small)
        run.font.size, run.font.color.rgb = Pt(7.5), GREY
        for c in (c0, c1):
            c.alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.paragraph_format.space_after = Pt(1)

    # --- Summary
    heading(doc, "Summary")
    para(doc, words["summary"], space_after=2)

    # --- Notable deals
    heading(doc, "Notable deals")
    cols = ["Issuer", "Priced", "Rating", "Ccy", "Size ($m eq.)", "Bonds", "Tenors", "Spread (bps)"]
    keys = ["issuer", "trade_date", "rating", "currency", "size_usd_m", "bonds", "tenors", "spread_range"]
    t = doc.add_table(rows=1, cols=len(cols))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False                      # use our column widths, not Word's guesses
    widths = [4.6, 1.4, 1.3, 1.1, 2.2, 1.3, 2.9, 2.6]   # cm, fits the 17.8cm text width
    for i, h in enumerate(cols):
        cell = t.rows[0].cells[i]
        shade(cell, "1F3864")
        run = cell.paragraphs[0].add_run(h)
        run.font.size, run.bold, run.font.color.rgb = Pt(7.5), True, RGBColor(255, 255, 255)
    for n, deal in enumerate(facts["notable_deals"]):
        row = t.add_row().cells
        for i, k in enumerate(keys):
            v = deal[k]
            text = f"{v:,.0f}" if k == "size_usd_m" else str(v)
            if k == "issuer" and deal["hybrid_or_sub"]:
                text += " *"
            run = row[i].paragraphs[0].add_run(text)
            run.font.size = Pt(7.5)
            if n % 2:
                shade(row[i], "F6F6F4")
    for i, col in enumerate(t.columns):
        col.width = Cm(widths[i])
    for row in t.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Cm(widths[i])
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    para(doc, "* includes subordinated or hybrid bonds. Spreads over each bond's benchmark "
              "(UST for USD, Bunds/mid-swaps for EUR, gilts for GBP).", size=6.5, color=GREY, space_after=0)

    # --- Charts (last 30 days of context)
    heading(doc, f"Pricing context: {p['from']} to {facts['week_end']}")
    t = doc.add_table(rows=1, cols=2)
    no_borders(t)
    for i, name in enumerate(("2_spread_by_tenor.png", "1_spread_by_rating.png")):
        path = config.CHARTS_DIR / name
        if path.exists():
            t.cell(0, i).paragraphs[0].add_run().add_picture(str(path), width=Cm(8.4))

    # --- Themes
    heading(doc, "Key themes")
    for theme in words["themes"]:
        bp = doc.add_paragraph(style="List Bullet")
        bp.paragraph_format.space_after = Pt(1)
        run = bp.add_run(theme)
        run.font.size = Pt(9)

    # --- Small print: sources, method, limitations
    accuracy = ""
    if results:
        g, s = results.get("grounding") or {}, results.get("spot_check") or {}
        if g and s:
            accuracy = (f" Extraction QC: {g['grounded_pct']}% of {g['values_checked']} extracted "
                        f"values traced to source filings; {s['accuracy_pct']}% correct on a manual "
                        f"spot check of {s['values_checked']} values.")
    small = (
        f"Source: SEC EDGAR pricing term sheets (form FWP), {p['from']} to {facts['week_end']}. "
        f"Terms extracted from filings with Claude ({config.MODEL}); all figures computed in Python "
        f"from the extracted data. Commentary drafted by Claude from those figures and checked "
        f"number-by-number against them.{accuracy} "
        "LIMITATIONS: covers SEC-registered deals only, so excludes 144A/Reg S issuance (most US high "
        "yield and many European deals); volumes are a sample, not the market. Non-USD sizes "
        f"converted at Federal Reserve H.10 rates for {config.FX_DATE}. Composite rating = middle of "
        "Moody's/S&P/Fitch; sector is a model classification. Tenor measured to par call where "
        "earlier. Automated draft for discussion; not investment advice."
    )
    p_small = para(doc, "", space_after=0)
    p_small.paragraph_format.space_before = Pt(6)
    run = p_small.add_run(small)
    run.font.size, run.font.color.rgb = Pt(6.5), GREY

    doc.save(docx_out)
    print(f"Saved {docx_out.name}")


def to_pdf(docx_out):
    pdf_out = docx_out.with_suffix(".pdf")
    try:
        from docx2pdf import convert
        convert(str(docx_out), str(pdf_out))
        print(f"Saved {pdf_out}")
    except Exception as err:
        print(f"PDF step skipped ({err}). Open the .docx in Word and use File > Save As > PDF.")


def main():
    parser = argparse.ArgumentParser(description="Build the weekly new-issue update")
    parser.add_argument("--week-ending", help="YYYY-MM-DD (default: latest trade date)")
    args = parser.parse_args()

    df = load()
    week_end = pd.Timestamp(args.week_ending) if args.week_ending else df["trade_date"].max()
    facts = build_facts(df, week_end)
    (config.OUTPUT_DIR / "weekly_facts.json").write_text(json.dumps(facts, indent=2, default=str))

    results_path = config.VALIDATION_DIR / "results.json"
    results = json.loads(results_path.read_text()) if results_path.exists() else None

    words = write_commentary(facts)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    docx_out = REPORTS_DIR / f"weekly_update_{week_end.date()}.docx"
    build_docx(facts, words, results, docx_out)
    to_pdf(docx_out)


if __name__ == "__main__":
    main()
