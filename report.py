"""
report.py  (Stage 6)
Builds the one-page "New Issue Market Update" (weekly or monthly) as Word and PDF.

THE KEY IDEA
Every number in the note comes from Python (the dataset). Claude only
writes the words around them, from a fact sheet we give it. Afterwards a
check confirms every number Claude wrote appears in that fact sheet, so
the model can't slip an invented spread into a document styled as a desk
note.

HOW TO RUN
    python report.py                          -> week ending on the latest deal
    python report.py --week-ending 2026-09-18 -> a specific week
    python report.py --period month           -> last full calendar month
    python report.py --period month --month 2026-09
Output: output/reports/<weekly|monthly>_update_<date>.docx and .pdf, plus a
copy of the newest PDF as output/reports/latest.pdf (the README links to it).
PDF conversion uses Microsoft Word on Windows, or LibreOffice elsewhere
(e.g. when GitHub runs it automatically).
"""
import argparse
import json
import re
import shutil
import subprocess
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
# Each report gets its own dated file, so older updates aren't overwritten
REPORTS_DIR = config.OUTPUT_DIR / "reports"

# How far back the "prior period" comparison looks: the previous 4 weeks
# for a weekly note, the previous 3 months for a monthly one.
PRIOR_WINDOW = {"week": pd.Timedelta(weeks=4), "month": pd.DateOffset(months=3)}


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
    """FINANCE: compare this period's new-issue spreads with the prior period,
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
        row = {"bonds_this_period": int(len(w)), "median_this_period_bps": round(float(w.median())),
               "bonds_prior": int(len(p)),
               "median_prior_bps": round(float(p.median())) if len(p) else None}
        if len(w) >= MIN_BONDS and len(p) >= MIN_BONDS:
            diff = row["median_this_period_bps"] - row["median_prior_bps"]
            row["vs_prior"] = "tighter" if diff < -2 else "wider" if diff > 2 else "unchanged"
        else:
            row["vs_prior"] = "too few bonds to compare"
        out[b] = row
    return out


def period_bounds(period, df, week_ending=None, month=None):
    """Work out the start and end dates of the report period."""
    if period == "month":
        if month:
            start = pd.Timestamp(month + "-01")
        else:   # default: the last COMPLETE calendar month before today
            start = (pd.Timestamp.today().normalize().replace(day=1) - pd.DateOffset(months=1))
        end = start + pd.DateOffset(months=1) - pd.Timedelta(days=1)
        return start, end, start.strftime("%B %Y")
    end = pd.Timestamp(week_ending) if week_ending else df["trade_date"].max()
    start = end - timedelta(days=6)
    return start, end, f"Week ending {end.strftime('%d %b %Y')}"


def build_facts(df, period, start, end, label):
    this = df[(df["trade_date"] >= start) & (df["trade_date"] <= end)]
    prior_start = max(start - PRIOR_WINDOW[period], df["trade_date"].min())
    before = df[(df["trade_date"] >= prior_start) & (df["trade_date"] < start)]
    if this.empty:
        raise SystemExit(f"No deals priced between {start.date()} and {end.date()}.")

    # FINANCE: average volume per week (or month) in the prior window, so
    # "a quiet week" means quiet relative to recent activity.
    days_per_period = 7 if period == "week" else 30.4
    covered = (start - prior_start).days
    n_prior = max(1, round(covered / days_per_period))
    # Only quote an average if the data covers at least one full prior
    # period; otherwise "avg" would really be a few days' worth.
    has_full_prior = covered >= 0.9 * days_per_period
    deals = deal_table(this)
    curve = tenor_comparison(senior_ig(this), senior_ig(before))
    chart_from = max(df["trade_date"].min(), end - pd.Timedelta(days=config.LOOKBACK_DAYS))

    return {
        "period": period,
        "label": label,
        "start": start.strftime("%d %b %Y"),
        "end": end.strftime("%d %b %Y"),
        "this_period": {
            "volume_usd_bn": round(this["size_usd_m"].sum() / 1000, 1),
            "deals": int(deals.shape[0]),
            "bonds": int(len(this)),
            "ig_share_pct": round(100 * this.loc[this["grade"] == "IG", "size_usd_m"].sum()
                                  / this["size_usd_m"].sum()),
            "currencies": sorted(this["currency"].dropna().unique().tolist()),
            "hybrid_or_subordinated_deals": int(deals["hybrid_or_sub"].sum()),
            "sector_volume_usd_bn": (this.groupby("sector")["size_usd_m"].sum() / 1000)
                                    .round(1).sort_values(ascending=False).to_dict(),
        },
        "senior_ig_spreads_by_tenor": curve,
        "prior_period": {
            "from": prior_start.strftime("%d %b %Y"),
            f"{period}s": n_prior,
            f"avg_{period}ly_volume_usd_bn": round(before["size_usd_m"].sum() / 1000 / n_prior, 1)
                                             if len(before) and has_full_prior else None,
        },
        "chart_window_from": chart_from.strftime("%d %b %Y"),
        "notable_deals": deals.head(6).to_dict("records"),
    }


# ======================================================================
# 2. THE WORDS (Claude, constrained to the facts)
# ======================================================================
COMMENTARY_TOOL = {
    "name": "write_update",
    "description": "Commentary for a new-issue market update.",
    "input_schema": {
        "type": "object",
        "properties": {
            "summary": {"type": "string",
                        "description": "2-3 sentences, max 70 words: the period in one paragraph."},
            "themes": {"type": "array", "items": {"type": "string"}, "minItems": 3, "maxItems": 3,
                       "description": "Three market themes, max 30 words each."},
        },
        "required": ["summary", "themes"],
    },
}

SYSTEM = """You are a debt capital markets analyst writing the {period}ly new-issue
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
            model=config.MODEL, max_tokens=800, system=SYSTEM.replace("{period}", facts["period"]),
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
    title = "Weekly" if facts["period"] == "week" else "Monthly"
    para(doc, f"{title} New Issue Market Update", size=17, bold=True, color=NAVY, space_after=0)
    para(doc, f"{facts['label']}  |  SEC-registered corporate bond issuance  |  "
              f"Prepared by {AUTHOR}", size=8.5, color=GREY, space_after=6)

    # --- KPI strip
    w, p, per = facts["this_period"], facts["prior_period"], facts["period"]
    curve = facts["senior_ig_spreads_by_tenor"]
    ten = curve.get("10y")
    avg = p[f"avg_{per}ly_volume_usd_bn"]
    kpis = [
        (f"${w['volume_usd_bn']}bn", f"Volume this {per}" + (f" (avg ${avg}bn)" if avg else "")),
        (f"{w['deals']} / {w['bonds']}", "Deals / bonds priced"),
        (f"{w['ig_share_pct']}%", "Investment grade share"),
        (f"+{ten['median_this_period_bps']}bp" if ten else "-",
         (f"Median senior IG 10y spread, {ten['bonds_this_period']} bond(s)"
          + (f" (prior +{ten['median_prior_bps']}bp)" if ten["median_prior_bps"] else ""))
         if ten else f"No senior IG 10y bonds this {per}"),
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
    heading(doc, f"Pricing context: {facts['chart_window_from']} to {facts['end']}")
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
        f"Source: SEC EDGAR pricing term sheets (form FWP), {p['from']} to {facts['end']}. "
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
    """Word -> PDF. On Windows this uses Microsoft Word (docx2pdf); on
    Linux (e.g. GitHub's servers) it uses LibreOffice instead."""
    pdf_out = docx_out.with_suffix(".pdf")
    try:
        from docx2pdf import convert
        convert(str(docx_out), str(pdf_out))
    except Exception as err:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            print(f"PDF step skipped ({err}). Open the .docx in Word and use File > Save As > PDF.")
            return None
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir",
                        str(docx_out.parent), str(docx_out)], check=True, capture_output=True)
    print(f"Saved {pdf_out}")
    return pdf_out


def main():
    parser = argparse.ArgumentParser(description="Build the new-issue market update")
    parser.add_argument("--period", choices=["week", "month"], default="week")
    parser.add_argument("--week-ending", help="YYYY-MM-DD (weekly; default: latest trade date)")
    parser.add_argument("--month", help="YYYY-MM (monthly; default: last full month)")
    args = parser.parse_args()

    df = load()
    start, end, label = period_bounds(args.period, df, args.week_ending, args.month)
    facts = build_facts(df, args.period, start, end, label)
    (config.OUTPUT_DIR / f"{args.period}ly_facts.json").write_text(
        json.dumps(facts, indent=2, default=str))

    # Accuracy line for the small print: the spot check from results.json,
    # plus the newest automatic grounding check if it has been re-run.
    results_path = config.VALIDATION_DIR / "results.json"
    grounding_path = config.VALIDATION_DIR / "grounding.json"
    results = json.loads(results_path.read_text()) if results_path.exists() else None
    if results and grounding_path.exists():
        results["grounding"] = json.loads(grounding_path.read_text())

    words = write_commentary(facts)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    stamp = end.strftime("%Y-%m") if args.period == "month" else str(end.date())
    docx_out = REPORTS_DIR / f"{args.period}ly_update_{stamp}.docx"
    build_docx(facts, words, results, docx_out)
    pdf = to_pdf(docx_out)
    if pdf and pdf.exists():
        shutil.copyfile(pdf, REPORTS_DIR / "latest.pdf")   # stable link for the README
        print("Copied to latest.pdf")


if __name__ == "__main__":
    main()
