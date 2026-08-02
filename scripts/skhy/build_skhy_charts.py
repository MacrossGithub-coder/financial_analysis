"""SK hynix (SKHY) FY2026 Q2 earnings-update chart builder."""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/SKHY")
OUT.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(ROOT / "tmp/pydeps"))

NAVY = "#102A43"
BLUE = "#2563A6"
ORANGE = "#F47C20"
RED = "#C62828"
GREEN = "#2E7D32"
GOLD = "#B7791F"
GREY = "#6B7280"
LIGHT = "#EAF0F6"
PALE_ORANGE = "#FFF1E6"

plt.rcParams.update(
    {
        "font.family": "serif",
        "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
        "axes.unicode_minus": False,
        "axes.titleweight": "bold",
        "axes.titlesize": 12,
        "axes.labelsize": 9.5,
        "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.5,
        "legend.fontsize": 8.2,
        "figure.dpi": 150,
    }
)


QUARTERS = ["3Q24", "4Q24", "1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26"]
REVENUE = np.array([17.573, 19.767, 17.639, 22.232, 24.449, 32.827, 52.576, 79.319])
GROSS_PROFIT = np.array([9.171, 10.366, 10.102, 11.983, 14.029, 22.576, 41.679, 65.991])
OP = np.array([7.030, 8.083, 7.441, 9.213, 11.383, 19.170, 37.610, 60.543])
EBITDA = np.array([10.100, 11.249, 10.769, 12.645, 14.949, 22.732, 41.336, 64.563])
NET = np.array([5.753, 8.006, 8.108, 6.996, 12.598, 15.246, 40.346, 93.923])
GM = np.array([52, 52, 57, 54, 57, 69, 79, 83])
OM = np.array([40, 41, 42, 41, 47, 58, 72, 76])


def get_market_data(ticker: str) -> dict:
    """Fetch live market data dynamically from yfinance."""
    try:
        import yfinance as yf

        t = yf.Ticker(ticker)
        info = t.fast_info
        hist = t.history(start="2026-07-10", auto_adjust=False)
        return {
            "price": round(float(info.last_price), 2),
            "market_cap": float(info.market_cap),
            "52w_high": round(float(info.year_high), 2),
            "52w_low": round(float(info.year_low), 2),
            "currency": str(info.currency),
            "history_dates": [d.strftime("%Y-%m-%d") for d in hist.index],
            "history_close": [round(float(x), 4) for x in hist["Close"]],
        }
    except Exception as exc:
        print(f"WARNING: yfinance market-data retrieval failed for {ticker}: {exc}")
        return {
            "price": "N/A",
            "market_cap": "N/A",
            "52w_high": "N/A",
            "52w_low": "N/A",
            "currency": "USD",
            "history_dates": [],
            "history_close": [],
        }


def get_fx_data() -> float | str:
    try:
        import yfinance as yf

        return round(float(yf.Ticker("KRW=X").fast_info.last_price), 4)
    except Exception as exc:
        print(f"WARNING: yfinance FX retrieval failed: {exc}")
        return "N/A"


def style_ax(ax, grid=True):
    ax.spines[["top", "right"]].set_visible(False)
    if grid:
        ax.yaxis.grid(True, color="#D8DEE7", linewidth=0.7, alpha=0.8)
        ax.set_axisbelow(True)
    ax.tick_params(colors="#3F4A5A")


def save(fig, stem):
    fig.tight_layout()
    path = OUT / f"skhy_chart{stem}.png"
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(path.name)


def chart1_revenue():
    x = np.arange(len(QUARTERS))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    colors = [BLUE] * 7 + [ORANGE]
    bars = ax.bar(x, REVENUE, color=colors, width=0.62)
    for b, v in zip(bars, REVENUE):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.1, f"{v:.1f}", ha="center", fontsize=8.2, color=NAVY, fontweight="bold")
    ax.set_xticks(x, QUARTERS)
    ax.set_ylabel("Revenue (KRW tn)")
    ax.set_title("Quarterly revenue: 2Q26 set another record")
    ax.annotate("+51% QoQ / +257% YoY", xy=(7, REVENUE[-1]), xytext=(5.15, 72), arrowprops=dict(arrowstyle="->", color=ORANGE), color=ORANGE, fontweight="bold")
    style_ax(ax)
    save(fig, "1_revenue")


def chart2_operating_profit():
    x = np.arange(len(QUARTERS))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    bars = ax.bar(x, OP, color=[NAVY] * 7 + [ORANGE], width=0.62)
    for b, v in zip(bars, OP):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.9, f"{v:.1f}", ha="center", fontsize=8.2, color=NAVY, fontweight="bold")
    ax.set_xticks(x, QUARTERS)
    ax.set_ylabel("Operating profit (KRW tn)")
    ax.set_title("Operating profit: fifth consecutive quarterly record")
    ax.annotate("76% operating margin", xy=(7, OP[-1]), xytext=(5.0, 55), arrowprops=dict(arrowstyle="->", color=ORANGE), color=ORANGE, fontweight="bold")
    style_ax(ax)
    save(fig, "2_operating_profit")


def chart3_margins():
    x = np.arange(len(QUARTERS))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    ax.plot(x, GM, marker="o", linewidth=2.3, color=BLUE, label="Gross margin")
    ax.plot(x, OM, marker="o", linewidth=2.3, color=ORANGE, label="Operating margin")
    for i in [0, 3, 5, 6, 7]:
        ax.text(i, GM[i] + 2.1, f"{GM[i]}%", ha="center", fontsize=8, color=BLUE)
        ax.text(i, OM[i] - 4.1, f"{OM[i]}%", ha="center", fontsize=8, color=ORANGE)
    ax.set_xticks(x, QUARTERS)
    ax.set_ylim(30, 90)
    ax.set_ylabel("Margin")
    ax.set_title("Margins expanded as ASP gains outpaced cost growth")
    ax.legend(frameon=False, ncol=2, loc="upper left")
    style_ax(ax)
    save(fig, "3_margins")


def chart4_normalized_net():
    reported = NET.copy()
    normalized = NET.copy()
    normalized[-2] = 40.346 - 9.94 * (1 - 0.235)
    normalized[-1] = 93.923 - 63.27 * (1 - 0.235)
    x = np.arange(len(QUARTERS))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    w = 0.35
    ax.bar(x - w / 2, reported, w, color=BLUE, label="Reported net profit")
    ax.bar(x + w / 2, normalized, w, color=ORANGE, label="Ex-investment valuation gains*")
    ax.set_xticks(x, QUARTERS)
    ax.set_ylabel("Net profit (KRW tn)")
    ax.set_title("Headline net profit was amplified by investment-asset gains")
    ax.legend(frameon=False, ncol=2, loc="upper left")
    ax.annotate("KRW63.27tn investment-asset gains", xy=(7 - w / 2, reported[-1]), xytext=(4.55, 82), arrowprops=dict(arrowstyle="->", color=RED), color=RED, fontsize=8.4, fontweight="bold")
    ax.text(0.01, -0.22, "*Illustrative normalization applies the quarter's effective tax rate; not a company-reported non-GAAP measure.", transform=ax.transAxes, fontsize=7.2, color=GREY)
    style_ax(ax)
    save(fig, "4_normalized_net")


def chart5_beat_miss():
    metrics = ["Revenue", "Operating\nprofit"]
    actual = np.array([79.3187, 60.5426])
    consensus = np.array([84.1, 64.1])
    miss_pct = (actual / consensus - 1) * 100
    x = np.arange(2)
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    w = 0.34
    ax.bar(x - w / 2, consensus, w, color="#BFC7D1", label="Pre-results consensus")
    ax.bar(x + w / 2, actual, w, color=ORANGE, label="Actual")
    for i in range(2):
        ax.text(i - w / 2, consensus[i] + 1.2, f"{consensus[i]:.1f}", ha="center", fontsize=8.5)
        ax.text(i + w / 2, actual[i] + 1.2, f"{actual[i]:.1f}", ha="center", fontsize=8.5, color=NAVY, fontweight="bold")
        ax.text(i, 18, f"{miss_pct[i]:.1f}% miss", ha="center", color=RED, fontweight="bold", fontsize=10)
    ax.set_xticks(x, metrics)
    ax.set_ylabel("KRW tn")
    ax.set_ylim(0, 95)
    ax.set_title("2Q26 actuals missed elevated pre-results expectations")
    ax.legend(frameon=False, ncol=2, loc="upper right")
    style_ax(ax)
    save(fig, "5_beat_miss")


def chart6_product_mix():
    qs = ["1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26"]
    rev = np.array([17.639, 22.232, 24.449, 32.827, 52.576, 79.319])
    dram_share = np.array([0.80, 0.77, 0.78, 0.76, 0.78, 0.73])
    nand_share = np.array([0.18, 0.21, 0.20, 0.23, 0.21, 0.27])
    other_share = 1 - dram_share - nand_share
    x = np.arange(len(qs))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    dram = rev * dram_share
    nand = rev * nand_share
    other = rev * other_share
    ax.bar(x, dram, color=BLUE, label="DRAM")
    ax.bar(x, nand, bottom=dram, color=ORANGE, label="NAND")
    ax.bar(x, other, bottom=dram + nand, color="#BFC7D1", label="Other")
    for i, total in enumerate(rev):
        ax.text(i, total + 1, f"{total:.1f}", ha="center", fontsize=8.1, fontweight="bold", color=NAVY)
    ax.text(5, dram[-1] + nand[-1] / 2, "27%", ha="center", va="center", color="white", fontweight="bold")
    ax.set_xticks(x, qs)
    ax.set_ylabel("Revenue (KRW tn)")
    ax.set_title("Product mix: NAND contribution increased on pricing and eSSD")
    ax.legend(frameon=False, ncol=3, loc="upper left")
    style_ax(ax)
    save(fig, "6_product_mix")


def chart7_cash_debt():
    qs = ["2Q25", "3Q25", "4Q25", "1Q26", "2Q26"]
    cash = np.array([16.962, 24.4, 34.942, 54.330, 87.958])
    debt = np.array([21.840, 21.9, 22.248, 19.318, 18.587])
    x = np.arange(len(qs))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    w = 0.34
    ax.bar(x - w / 2, cash, w, color=BLUE, label="Cash & short-term investments")
    ax.bar(x + w / 2, debt, w, color=GREY, label="Interest-bearing debt")
    for i in range(len(qs)):
        ax.text(i - w / 2, cash[i] + 1.1, f"{cash[i]:.1f}", ha="center", fontsize=8)
        ax.text(i + w / 2, debt[i] + 1.1, f"{debt[i]:.1f}", ha="center", fontsize=8)
    ax.set_xticks(x, qs)
    ax.set_ylabel("KRW tn")
    ax.set_title("Balance sheet moved rapidly into net cash")
    ax.legend(frameon=False, ncol=2, loc="upper left")
    ax.text(2.55, 82, "2Q26 net cash: KRW69.4tn", color=GREEN, fontweight="bold")
    style_ax(ax)
    save(fig, "7_cash_debt")


def chart8_cash_flow():
    labels = ["Operating\nCF", "PP&E\ncapex", "Illustrative\nFCF", "Net cash\nchange"]
    vals = [65.710, -10.671, 55.039, 33.628]
    colors = [BLUE, RED, GREEN, ORANGE]
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    bars = ax.bar(np.arange(4), vals, color=colors, width=0.62)
    ax.axhline(0, color="#6B7280", linewidth=0.8)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + (1.4 if v >= 0 else -2.8), f"{v:.1f}", ha="center", va="bottom" if v >= 0 else "top", fontsize=8.5, color=NAVY, fontweight="bold")
    ax.set_xticks(np.arange(4), labels)
    ax.set_ylabel("KRW tn")
    ax.set_title("2Q26 cash conversion remained strong despite working-capital build")
    ax.text(0.01, -0.22, "Illustrative FCF = cash from operations less acquisition of PP&E; company does not report this as a non-GAAP metric.", transform=ax.transAxes, fontsize=7.2, color=GREY)
    style_ax(ax)
    save(fig, "8_cash_flow")


def chart9_outlook():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(9.2, 3.7))
    categories = ["2026 DRAM\ndemand", "2026 NAND\ndemand", "3Q DRAM\nshipments", "3Q NAND\nshipments"]
    low = np.array([23, 17, 9, 1])
    high = np.array([27, 19, 11, 4])
    mid = (low + high) / 2
    err = np.vstack([mid - low, high - mid])
    ax1.bar(np.arange(4), mid, color=[BLUE, ORANGE, BLUE, ORANGE], width=0.62, yerr=err, capsize=5)
    ax1.set_xticks(np.arange(4), categories)
    ax1.set_ylabel("Bit growth")
    ax1.set_title("Demand and 3Q shipment guidance")
    ax1.set_ylim(0, 32)
    style_ax(ax1)
    milestones = ["2Q26", "1H26", "2H26", "YE26", "Early 2027"]
    y = np.arange(len(milestones))[::-1]
    ax2.scatter([0] * len(y), y, s=65, color=ORANGE, zorder=3)
    ax2.plot([0] * len(y), y, color="#F3B27A", linewidth=2)
    texts = ["HBM4 shipments began", "HBM4E samples delivered", "HBM4 volume ramp", "321L ~50% of domestic NAND capacity", "M15X/Yongin capacity acceleration"]
    for yi, m, txt in zip(y, milestones, texts):
        ax2.text(-0.08, yi, m, ha="right", va="center", fontsize=8, color=GREY)
        ax2.text(0.08, yi, txt, ha="left", va="center", fontsize=7.9, color=NAVY)
    ax2.set_xlim(-0.5, 2.5)
    ax2.set_ylim(-0.7, len(y) - 0.3)
    ax2.axis("off")
    ax2.set_title("Technology and capacity milestones", pad=10)
    save(fig, "9_outlook")


def chart10_estimate_revisions():
    labels = ["FY26 Revenue", "FY26 OP", "FY26 Norm. NI", "FY27 Revenue", "FY27 OP", "FY27 Norm. NI"]
    old = np.array([339.433, 266.650, 217.318, 506.832, 419.522, 319.022])
    new = np.array([324.895, 248.153, 190.2, 420.0, 315.0, 242.0])
    x = np.arange(len(labels))
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    w = 0.36
    ax.bar(x - w / 2, old, w, color="#BFC7D1", label="Pre-results base case")
    ax.bar(x + w / 2, new, w, color=ORANGE, label="Post-results estimates")
    for i, pct in enumerate((new / old - 1) * 100):
        ax.text(i, max(old[i], new[i]) + 13, f"{pct:.1f}%", ha="center", color=RED, fontsize=8.2, fontweight="bold")
    ax.set_xticks(x, labels, rotation=18, ha="right")
    ax.set_ylabel("KRW tn")
    ax.set_title("Estimate revisions: prudence on supply, mix and cycle duration")
    ax.legend(frameon=False, ncol=2, loc="upper left")
    style_ax(ax)
    save(fig, "10_estimates")


def chart11_price(market):
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    if market["history_close"]:
        dates = market["history_dates"]
        close = np.array(market["history_close"])
        x = np.arange(len(close))
        ax.plot(x, close, color=BLUE, linewidth=2.2)
        ax.fill_between(x, close, close.min() - 3, color=LIGHT, alpha=0.8)
        tick_idx = np.unique(np.linspace(0, len(x) - 1, min(6, len(x))).astype(int))
        ax.set_xticks(tick_idx, [dates[i][5:] for i in tick_idx])
        if "2026-07-29" in dates:
            i = dates.index("2026-07-29")
            ax.scatter(i, close[i], color=ORANGE, s=55, zorder=4)
            ax.annotate("2Q26 release", xy=(i, close[i]), xytext=(max(0, i - 6), close[i] - 15), arrowprops=dict(arrowstyle="->", color=ORANGE), color=ORANGE, fontweight="bold")
        ax.axhline(149, color=GREY, linestyle="--", linewidth=1, label="IPO price: $149")
        ax.set_ylabel("SKHY close (US$)")
        ax.legend(frameon=False, loc="upper right")
    else:
        ax.text(0.5, 0.5, "Market data unavailable", transform=ax.transAxes, ha="center", va="center", color=GREY)
    ax.set_title("SKHY trading since the July 2026 Nasdaq listing")
    style_ax(ax)
    save(fig, "11_price")


def chart12_valuation():
    eps = np.array([20.5, 22.0, 23.1, 24.5, 26.0])
    multiples = np.array([6.0, 7.0, 7.5, 8.0, 9.0])
    values = np.outer(multiples, eps)
    fig, ax = plt.subplots(figsize=(9.2, 3.7))
    im = ax.imshow(values, cmap="Blues", aspect="auto")
    for i in range(len(multiples)):
        for j in range(len(eps)):
            ax.text(j, i, f"${values[i, j]:.0f}", ha="center", va="center", color="white" if values[i, j] > 175 else NAVY, fontsize=8.4, fontweight="bold")
    ax.set_xticks(np.arange(len(eps)), [f"${x:.1f}" for x in eps])
    ax.set_yticks(np.arange(len(multiples)), [f"{x:.1f}x" for x in multiples])
    ax.set_xlabel("FY27 normalized ADR EPS")
    ax.set_ylabel("Target P/E")
    ax.set_title("ADR fair-value sensitivity")
    ax.scatter(2, 2, marker="s", s=420, facecolors="none", edgecolors=ORANGE, linewidths=2.4)
    ax.text(2, 2.43, "Base case", ha="center", color=ORANGE, fontsize=8, fontweight="bold")
    fig.colorbar(im, ax=ax, fraction=0.025, pad=0.03, label="US$ / ADS")
    save(fig, "12_valuation")


def main():
    market = get_market_data("SKHY")
    market["usdkrw"] = get_fx_data()
    (OUT / "skhy_market_data.json").write_text(json.dumps(market, ensure_ascii=False, indent=2), encoding="utf-8")
    chart1_revenue()
    chart2_operating_profit()
    chart3_margins()
    chart4_normalized_net()
    chart5_beat_miss()
    chart6_product_mix()
    chart7_cash_debt()
    chart8_cash_flow()
    chart9_outlook()
    chart10_estimate_revisions()
    chart11_price(market)
    chart12_valuation()


if __name__ == "__main__":
    main()
