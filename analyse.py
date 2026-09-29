"""
analyse.py  (Stage 5)
Summary statistics and four charts from output/bond_dataset.csv.

OUTPUTS
    output/summary.json            key numbers (the weekly report reads these)
    output/league_table.csv        bookrunner league table
    output/charts/*.png            four charts, sized for the one-page report

HOW TO RUN
    python analyse.py
No API calls: free, and safe to re-run whenever the dataset changes.
"""
import json
import re

import matplotlib
matplotlib.use("Agg")            # draw straight to files, no pop-up windows
import matplotlib.pyplot as plt
from matplotlib import font_manager
import numpy as np
import pandas as pd

import config

# ----------------------------------------------------------------------
# STYLE: one consistent, restrained look for every chart.
# Colours are a colour-blind-safe pair (checked with a CVD validator):
# blue for the main series, orange only when a second series is needed.
# ----------------------------------------------------------------------
BLUE, ORANGE = "#2a78d6", "#eb6834"
INK, INK_2, GRID = "#0b0b0b", "#52514e", "#e4e3df"
# Arial on Windows; fall back quietly to matplotlib's built-in font elsewhere
FONT = "Arial" if "Arial" in {f.name for f in font_manager.fontManager.ttflist} else "DejaVu Sans"
plt.rcParams.update({
    "font.family": FONT, "font.size": 9,
    "axes.edgecolor": GRID, "axes.labelcolor": INK_2, "axes.titleweight": "bold",
    "axes.titlesize": 10.5, "axes.titlelocation": "left", "axes.titlecolor": INK,
    "xtick.color": INK_2, "ytick.color": INK_2, "axes.spines.top": False,
    "axes.spines.right": False, "figure.dpi": 200, "savefig.bbox": "tight",
    "legend.frameon": False,
})
FIGSIZE = (6.2, 3.3)


def style_axes(ax, grid_axis="y"):
    ax.grid(axis=grid_axis, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)                 # gridlines behind the data
    ax.tick_params(length=0)


def save(fig, name):
    config.CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(config.CHARTS_DIR / name, facecolor="white")
    plt.close(fig)


# ----------------------------------------------------------------------
# DATA
# ----------------------------------------------------------------------
def load():
    df = pd.read_csv(config.DATASET_CSV, parse_dates=["trade_date", "maturity_date"])
    # FINANCE: one "deal" can be filed as several documents (TD filed each
    # bond separately), so a deal = same issuer pricing on the same day.
    df["deal_id"] = df["issuer"].str.lower().str[:25] + "|" + df["trade_date"].astype(str)
    return df


def comparable_spreads(df):
    """FINANCE: only compare like with like. Spreads here are all over US
    Treasuries (the USD benchmark), on fixed-rate or fixed-to-floating bonds.
    Floating-rate notes quote a spread over SOFR, a different benchmark."""
    return df[(df["benchmark_type"] == "UST") & df["spread_bps"].notna()
              & df["pricing_tenor_years"].notna()]


# ----------------------------------------------------------------------
# BOOKRUNNER LEAGUE TABLE
# FINANCE: league tables rank banks by deals led. Term sheets name the same
# bank many ways ("BofA Securities, Inc.", "Merrill Lynch International"),
# so we map names to one parent bank before counting.
# ----------------------------------------------------------------------
BANK_ALIASES = [
    (r"morgan stanley", "Morgan Stanley"), (r"goldman", "Goldman Sachs"),
    (r"j\.?\s?p\.?\s?morgan", "J.P. Morgan"), (r"bofa|merrill", "BofA Securities"),
    (r"citi", "Citigroup"), (r"barclays", "Barclays"), (r"wells fargo", "Wells Fargo"),
    (r"hsbc", "HSBC"), (r"deutsche", "Deutsche Bank"), (r"bnp", "BNP Paribas"),
    (r"mizuho", "Mizuho"), (r"mufg|mitsubishi", "MUFG"), (r"smbc|sumitomo", "SMBC Nikko"),
    (r"\brbc\b|royal bank of canada", "RBC"), (r"\btd\b|toronto", "TD Securities"),
    (r"\bbmo\b|montreal", "BMO"), (r"scotia", "Scotiabank"), (r"santander", "Santander"),
    (r"bbva|bilbao", "BBVA"), (r"agricole", "Credit Agricole"), (r"\bsg\b|soci[eé]t[eé] g", "Societe Generale"),
    (r"natixis", "Natixis"), (r"\bubs\b", "UBS"), (r"jefferies", "Jefferies"),
    (r"u\.?s\.? banc", "U.S. Bancorp"), (r"\bpnc\b", "PNC"), (r"truist", "Truist"),
    (r"ing\b", "ING"), (r"lloyds", "Lloyds"), (r"natwest", "NatWest"), (r"cibc", "CIBC"),
    (r"keybanc", "KeyBanc"), (r"regions", "Regions"), (r"nomura", "Nomura"),
]


def parent_bank(name):
    low = name.lower()
    for pattern, parent in BANK_ALIASES:
        if re.search(pattern, low):
            return parent
    return re.sub(r"[,.].*$", "", name).strip()      # fall back to the name before any comma


def league_table(df):
    deals = df.groupby("deal_id").agg(bookrunners=("bookrunners", "first"),
                                     usd_m=("size_usd_m", "sum"))
    rows = []
    for _, deal in deals.dropna(subset=["bookrunners"]).iterrows():
        banks = sorted({parent_bank(b) for b in deal["bookrunners"].split(";") if b.strip()})
        for bank in banks:
            # FINANCE: standard league-table convention gives each joint
            # bookrunner an equal share of the deal's volume.
            rows.append({"bank": bank, "usd_m": deal["usd_m"] / len(banks)})
    lt = (pd.DataFrame(rows).groupby("bank")
          .agg(deals=("usd_m", "size"), credited_usd_m=("usd_m", "sum"))
          .sort_values(["credited_usd_m", "deals"], ascending=False).round(0))
    lt.insert(0, "rank", range(1, len(lt) + 1))
    return lt


# ----------------------------------------------------------------------
# CHART 1: spread by rating (and the cost of subordination)
# ----------------------------------------------------------------------
def chart_spread_by_rating(df):
    d = comparable_spreads(df).dropna(subset=["rating_score"]).copy()
    d["sub"] = d["seniority"].fillna("").str.contains("subordinated")
    ratings = (d.sort_values("rating_score")["composite_rating"].drop_duplicates().tolist())
    xpos = {r: i for i, r in enumerate(ratings)}
    rng = np.random.default_rng(0)         # fixed jitter so the chart is identical each run

    fig, ax = plt.subplots(figsize=FIGSIZE)
    for is_sub, colour, label in ((False, BLUE, "Senior"), (True, ORANGE, "Subordinated / hybrid")):
        part = d[d["sub"] == is_sub]
        if part.empty:
            continue
        x = part["composite_rating"].map(xpos) + rng.uniform(-0.18, 0.18, len(part))
        ax.scatter(x, part["spread_bps"], s=34, color=colour, alpha=0.85,
                   edgecolor="white", linewidth=0.8, label=label, zorder=3)
    # median of SENIOR bonds per rating, drawn as a short dark line
    for r, grp in d[~d["sub"]].groupby("composite_rating"):
        m = grp["spread_bps"].median()
        ax.plot([xpos[r] - 0.3, xpos[r] + 0.3], [m, m], color=INK, linewidth=2, zorder=4)
    ax.set_xticks(range(len(ratings)), ratings)
    ax.set_ylabel("Spread over Treasuries (bps)")
    ax.set_title("New-issue spread by composite rating")
    ax.legend(loc="upper left", fontsize=8)
    ax.text(1, -0.16, "Line = median senior spread. USD Treasury-benchmarked bonds only.",
            transform=ax.transAxes, ha="right", fontsize=7, color=INK_2)
    style_axes(ax)
    save(fig, "1_spread_by_rating.png")


# ----------------------------------------------------------------------
# CHART 2: the credit curve (spread by tenor)
# ----------------------------------------------------------------------
def chart_spread_by_tenor(df):
    d = comparable_spreads(df)
    # FINANCE: isolate the curve by holding credit quality and seniority
    # roughly constant: senior investment grade only, split A vs BBB.
    d = d[(d["grade"] == "IG") & ~d["seniority"].fillna("").str.contains("subordinated")]
    groups = ((d["rating_score"] <= 7, BLUE, "A- or better"),
              (d["rating_score"] >= 8, ORANGE, "BBB+ to BBB-"))

    fig, ax = plt.subplots(figsize=FIGSIZE)
    for mask, colour, label in groups:
        part = d[mask].sort_values("pricing_tenor_years")
        if part.empty:
            continue
        ax.scatter(part["pricing_tenor_years"], part["spread_bps"], s=34, color=colour,
                   alpha=0.85, edgecolor="white", linewidth=0.8, label=label, zorder=3)
        # median per tenor bucket joined up = a simple "fair value" curve
        curve = part.groupby("tenor_bucket").agg(t=("pricing_tenor_years", "median"),
                                                 s=("spread_bps", "median")).sort_values("t")
        if len(curve) > 1:
            ax.plot(curve["t"], curve["s"], color=colour, linewidth=2, zorder=2)
    ax.set_xlabel("Years to maturity (or to par call)")
    ax.set_ylabel("Spread over Treasuries (bps)")
    ax.set_title("IG credit curve: longer bonds pay more spread")
    ax.legend(loc="lower right", fontsize=8)
    style_axes(ax, "both")
    save(fig, "2_spread_by_tenor.png")


# ----------------------------------------------------------------------
# CHART 3: issuance volume by sector
# ----------------------------------------------------------------------
def chart_volume_by_sector(df):
    vol = (df.groupby("sector")["size_usd_m"].sum() / 1000).sort_values()
    fig, ax = plt.subplots(figsize=FIGSIZE)
    bars = ax.barh(vol.index, vol.values, color=BLUE, height=0.62)
    for bar, v in zip(bars, vol.values):
        ax.text(v + vol.max() * 0.01, bar.get_y() + bar.get_height() / 2, f"${v:,.1f}bn",
                va="center", fontsize=8, color=INK_2)
    ax.set_xlabel("Volume (USD bn equivalent)")
    ax.set_title("Issuance volume by sector")
    ax.set_xlim(0, vol.max() * 1.15)
    style_axes(ax, "x")
    save(fig, "3_volume_by_sector.png")


# ----------------------------------------------------------------------
# CHART 4: IG vs HY split
# ----------------------------------------------------------------------
def chart_grade_split(df):
    order = [g for g in ("IG", "HY", "Unrated") if g in set(df["grade"])]
    vol = df.groupby("grade")["size_usd_m"].sum().reindex(order) / 1000
    count = df.groupby("grade").size().reindex(order)
    share = 100 * vol / vol.sum()
    labels = {"IG": "Investment grade", "HY": "High yield", "Unrated": "Unrated"}

    fig, ax = plt.subplots(figsize=(FIGSIZE[0], 2.1))
    bars = ax.barh([labels[g] for g in order][::-1], vol.values[::-1], color=BLUE, height=0.55)
    for bar, g in zip(bars, order[::-1]):
        ax.text(bar.get_width() + vol.max() * 0.01, bar.get_y() + bar.get_height() / 2,
                f"${vol[g]:,.1f}bn  ·  {share[g]:.0f}%  ·  {count[g]} bonds",
                va="center", fontsize=8, color=INK_2)
    ax.set_xlim(0, vol.max() * 1.45)
    ax.set_xlabel("Volume (USD bn equivalent)")
    ax.set_title("Investment grade vs high yield")
    style_axes(ax, "x")
    save(fig, "4_ig_vs_hy.png")


# ----------------------------------------------------------------------
# SUMMARY STATS
# ----------------------------------------------------------------------
def summary(df, lt):
    comp = comparable_spreads(df)
    senior_ig = comp[(comp["grade"] == "IG") & ~comp["seniority"].fillna("").str.contains("subordinated")]
    deals = (df.groupby("deal_id").agg(issuer=("issuer", "first"), trade_date=("trade_date", "first"),
                                       bonds=("issuer", "size"), usd_m=("size_usd_m", "sum"),
                                       currency=("currency", "first"))
             .sort_values("usd_m", ascending=False))
    return {
        "window_start": df["trade_date"].min().date().isoformat(),
        "window_end": df["trade_date"].max().date().isoformat(),
        "deals": int(deals.shape[0]),
        "bonds": int(len(df)),
        "total_usd_bn": round(df["size_usd_m"].sum() / 1000, 1),
        "avg_deal_usd_m": round(deals["usd_m"].mean(), 0),
        "avg_bonds_per_deal": round(len(df) / deals.shape[0], 1),
        "volume_by_currency_usd_bn": (df.groupby("currency")["size_usd_m"].sum() / 1000).round(1).to_dict(),
        "volume_by_grade_usd_bn": (df.groupby("grade")["size_usd_m"].sum() / 1000).round(1).to_dict(),
        "volume_by_sector_usd_bn": (df.groupby("sector")["size_usd_m"].sum() / 1000)
                                   .round(1).sort_values(ascending=False).to_dict(),
        "median_senior_ig_spread_by_tenor_bps": senior_ig.groupby("tenor_bucket")["spread_bps"]
                                                .median().round(0).to_dict(),
        "median_spread_by_rating_bps": comp.groupby("composite_rating")["spread_bps"].median().round(0).to_dict(),
        "largest_deals": [{"issuer": r.issuer, "trade_date": r.trade_date.date().isoformat(),
                           "bonds": int(r.bonds), "usd_bn": round(r.usd_m / 1000, 2)}
                          for r in deals.head(5).itertuples()],
        "top_bookrunners": lt.head(5).reset_index()[["bank", "deals", "credited_usd_m"]].to_dict("records"),
    }


def main():
    df = load()
    lt = league_table(df)
    lt.to_csv(config.OUTPUT_DIR / "league_table.csv", encoding="utf-8-sig")
    chart_spread_by_rating(df)
    chart_spread_by_tenor(df)
    chart_volume_by_sector(df)
    chart_grade_split(df)

    s = summary(df, lt)
    (config.OUTPUT_DIR / "summary.json").write_text(json.dumps(s, indent=2, default=float),
                                                    encoding="utf-8")
    print(f"Window: {s['window_start']} to {s['window_end']}")
    print(f"{s['deals']} deals | {s['bonds']} bonds | ${s['total_usd_bn']}bn total "
          f"| avg deal ${s['avg_deal_usd_m']:,.0f}m")
    print(f"By grade ($bn): {s['volume_by_grade_usd_bn']}")
    print(f"Median senior IG spread by tenor (bps): {s['median_senior_ig_spread_by_tenor_bps']}")
    print("\nLargest deals:")
    for d in s["largest_deals"]:
        print(f"  {d['issuer']} ({d['trade_date']}): ${d['usd_bn']}bn across {d['bonds']} bonds")
    print("\nBookrunner league table (top 10):")
    print(lt.head(10).to_string())
    print(f"\nCharts saved in {config.CHARTS_DIR}")


if __name__ == "__main__":
    main()
