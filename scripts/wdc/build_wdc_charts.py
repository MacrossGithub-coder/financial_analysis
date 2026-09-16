#!/usr/bin/env python3
"""Generate charts for the Western Digital Q4 FY2026 earnings update."""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import yfinance as yf


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "WDC"
OUT.mkdir(parents=True, exist_ok=True)

plt.rcParams.update(
    {
        "font.family": "serif",
        "font.serif": ["Times New Roman"],
        "axes.unicode_minus": False,
        "figure.dpi": 150,
        "savefig.dpi": 150,
        "savefig.bbox": "tight",
        "axes.grid": True,
        "grid.alpha": 0.22,
        "axes.spines.top": False,
        "axes.spines.right": False,
    }
)

NAVY = "#17365D"
BLUE = "#2F5597"
CYAN = "#37A6C8"
TEAL = "#168A8C"
GREEN = "#416900"
ORANGE = "#C65D17"
RED = "#A61B1B"
GREY = "#7B8494"
LIGHT = "#DCE6F1"

quarters = ["Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26"]
revenue = np.array([2.605, 2.818, 3.017, 3.337, 3.747])
nongaap_eps = np.array([1.70, 1.78, 2.13, 2.72, 3.56])
gaap_eps = np.array([0.67, 3.07, 4.73, 8.20, 8.21])
gross_margin = np.array([41.3, 43.9, 46.1, 50.5, 54.4])
operating_margin = np.array([28.1, 30.4, 33.8, 38.6, 44.2])
nearline_eb = np.array([170, 183, 192, 199, 209])
non_nearline_eb = np.array([20, 21, 23, 23, 22])
fcf = np.array([0.675, 0.599, 0.653, 0.978, 1.281])
cloud_mix = np.array([90, 89, 89, 89, 89])
client_mix = np.array([5, 5, 6, 5, 6])
consumer_mix = np.array([5, 6, 5, 6, 5])


def save(fig, filename):
    fig.tight_layout()
    fig.savefig(OUT / filename)
    plt.close(fig)


def label_bars(ax, bars, labels, offset=0.03, color="#182230", size=8.5):
    for bar, label in zip(bars, labels):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + offset,
            label,
            ha="center",
            va="bottom",
            fontsize=size,
            fontweight="bold",
            color=color,
        )


# 1. Revenue progression
fig, ax = plt.subplots(figsize=(9.2, 4.7))
bars = ax.bar(quarters, revenue, color=[BLUE] * 4 + [NAVY], width=0.62)
label_bars(ax, bars, [f"${x:.3f}B" for x in revenue], offset=0.045)
ax.axhline(3.699, color=RED, linestyle="--", linewidth=1.2, label="Q4 consensus: $3.699B")
ax.set_ylim(0, 4.25)
ax.set_ylabel("Revenue ($bn)")
ax.set_title("Quarterly Revenue Progression", fontweight="bold")
ax.legend(loc="upper left", frameon=False)
save(fig, "wdc_chart1_quarterly_revenue.png")


# 2. EPS progression
fig, ax = plt.subplots(figsize=(9.2, 4.7))
x = np.arange(len(quarters))
w = 0.34
b1 = ax.bar(x - w / 2, nongaap_eps, w, color=BLUE, label="Non-GAAP EPS")
b2 = ax.bar(x + w / 2, gaap_eps, w, color=LIGHT, edgecolor=NAVY, label="GAAP EPS")
for bars, values in ((b1, nongaap_eps), (b2, gaap_eps)):
    label_bars(ax, bars, [f"${v:.2f}" for v in values], offset=0.10, size=7.7)
ax.axhline(3.3002, color=RED, linestyle="--", linewidth=1.1, label="Q4 adjusted consensus: $3.30")
ax.set_xticks(x, quarters)
ax.set_ylim(0, 9.4)
ax.set_ylabel("Diluted EPS ($)")
ax.set_title("GAAP and Non-GAAP EPS", fontweight="bold")
ax.legend(loc="upper left", frameon=False, ncol=2, fontsize=8.5)
save(fig, "wdc_chart2_eps.png")


# 3. Margin progression
fig, ax = plt.subplots(figsize=(9.2, 4.7))
ax.plot(quarters, gross_margin, marker="o", linewidth=2.4, color=BLUE, label="Non-GAAP gross margin")
ax.plot(quarters, operating_margin, marker="s", linewidth=2.4, color=TEAL, label="Non-GAAP operating margin")
for i, value in enumerate(gross_margin):
    ax.annotate(f"{value:.1f}%", (i, value), xytext=(0, 9), textcoords="offset points", ha="center", fontsize=8, color=NAVY)
for i, value in enumerate(operating_margin):
    ax.annotate(f"{value:.1f}%", (i, value), xytext=(0, -14), textcoords="offset points", ha="center", fontsize=8, color=TEAL)
ax.set_ylim(20, 60)
ax.set_ylabel("Margin")
ax.set_title("Margin Expansion Accelerated Through FY2026", fontweight="bold")
ax.legend(loc="upper left", frameon=False)
save(fig, "wdc_chart3_margins.png")


# 4. Exabyte shipments
fig, ax = plt.subplots(figsize=(9.2, 4.7))
ax.bar(quarters, nearline_eb, color=NAVY, width=0.62, label="Nearline")
ax.bar(quarters, non_nearline_eb, bottom=nearline_eb, color=CYAN, width=0.62, label="Non-nearline")
for i, total in enumerate(nearline_eb + non_nearline_eb):
    ax.text(i, total + 4, f"{total} EB", ha="center", fontweight="bold", fontsize=8.5)
ax.set_ylim(0, 260)
ax.set_ylabel("Exabytes shipped")
ax.set_title("Quarterly Exabyte Shipments", fontweight="bold")
ax.legend(loc="upper left", frameon=False)
save(fig, "wdc_chart4_exabytes.png")


# 5. End-market revenue mix
fig, ax = plt.subplots(figsize=(9.2, 4.7))
ax.bar(quarters, cloud_mix, color=NAVY, width=0.62, label="Cloud")
ax.bar(quarters, client_mix, bottom=cloud_mix, color=CYAN, width=0.62, label="Client")
ax.bar(quarters, consumer_mix, bottom=cloud_mix + client_mix, color=ORANGE, width=0.62, label="Consumer")
for i in range(len(quarters)):
    ax.text(i, cloud_mix[i] / 2, f"{cloud_mix[i]}%", ha="center", va="center", color="white", fontweight="bold", fontsize=9)
ax.set_ylim(0, 100)
ax.set_ylabel("Revenue mix")
ax.set_title("End-Market Revenue Mix", fontweight="bold")
ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.23), ncol=3, frameon=False)
save(fig, "wdc_chart5_end_market_mix.png")


# 6. Beat/miss dashboard
fig, axes = plt.subplots(1, 2, figsize=(9.2, 4.4))
metrics = [("Revenue", 3.699, 3.747, "$bn"), ("Non-GAAP EPS", 3.3002, 3.56, "$")]
for ax, (name, consensus, actual, unit) in zip(axes, metrics):
    bars = ax.bar(["Consensus", "Reported"], [consensus, actual], color=[GREY, BLUE], width=0.55)
    variance = (actual / consensus - 1) * 100
    labels = [f"{unit}{consensus:.2f}" if unit == "$" else f"${consensus:.3f}B", f"{unit}{actual:.2f}" if unit == "$" else f"${actual:.3f}B"]
    label_bars(ax, bars, labels, offset=max(actual * 0.025, 0.04), size=9)
    ax.text(0.5, 0.08, f"Beat: +{variance:.1f}%", transform=ax.transAxes, ha="center", color=GREEN, fontweight="bold", fontsize=10)
    ax.set_title(name, fontweight="bold")
    ax.set_ylim(0, actual * 1.25)
fig.suptitle("Q4 FY2026 Beat vs. Published Consensus", fontweight="bold", y=1.01)
save(fig, "wdc_chart6_beat_miss.png")


# 7. Free cash flow
fig, ax = plt.subplots(figsize=(9.2, 4.7))
bars = ax.bar(quarters, fcf, color=[CYAN] * 4 + [TEAL], width=0.62)
label_bars(ax, bars, [f"${v:.3f}B" for v in fcf], offset=0.025)
ax.set_ylim(0, 1.5)
ax.set_ylabel("Free cash flow ($bn)")
ax.set_title("Quarterly Free Cash Flow", fontweight="bold")
save(fig, "wdc_chart7_fcf.png")


# 8. FY2026 revenue growth decomposition
fig, ax = plt.subplots(figsize=(8.8, 4.7))
drivers = ["Exabytes\nsold", "ASP per\nexabyte", "Total revenue\ngrowth"]
values = [25, 8, 36]
bars = ax.bar(drivers, values, color=[CYAN, ORANGE, NAVY], width=0.56)
label_bars(ax, bars, [f"+{v}%" for v in values], offset=0.8)
ax.set_ylim(0, 43)
ax.set_ylabel("Year-over-year growth")
ax.set_title("FY2026 Growth: Volume and Pricing Both Contributed", fontweight="bold")
save(fig, "wdc_chart8_growth_drivers.png")


# 9. Q1 FY2027 guidance vs Q4 actual
fig, axes = plt.subplots(1, 3, figsize=(9.4, 4.4))
guidance = [
    ("Revenue ($bn)", 3.747, 4.10),
    ("Non-GAAP gross margin", 54.4, 55.5),
    ("Non-GAAP EPS ($)", 3.56, 4.00),
]
for ax, (label, actual, guide) in zip(axes, guidance):
    bars = ax.bar(["Q4A", "Q1E\nmidpoint"], [actual, guide], color=[GREY, BLUE], width=0.55)
    formats = [f"{actual:.1f}%", f"{guide:.1f}%"] if "margin" in label else [f"${actual:.2f}", f"${guide:.2f}"]
    label_bars(ax, bars, formats, offset=max(guide * 0.03, 0.08), size=8.5)
    ax.set_title(label, fontsize=9.5, fontweight="bold")
    ax.set_ylim(0, guide * 1.28)
fig.suptitle("Q1 FY2027 Guidance Extends the Earnings Ramp", fontweight="bold", y=1.01)
save(fig, "wdc_chart9_guidance.png")


# 10. Estimate revisions
fig, axes = plt.subplots(1, 2, figsize=(9.2, 4.5))
revision_sets = [
    ("FY2027E revenue ($bn)", 18.2, 18.8),
    ("FY2027E non-GAAP EPS ($)", 17.6, 18.9),
]
for ax, (label, old, new) in zip(axes, revision_sets):
    bars = ax.bar(["Pre-print", "Post-print"], [old, new], color=[GREY, BLUE], width=0.55)
    label_bars(ax, bars, [f"${old:.1f}", f"${new:.1f}"], offset=new * 0.025, size=9)
    ax.text(0.5, 0.08, f"+{(new / old - 1) * 100:.1f}%", transform=ax.transAxes, ha="center", color=GREEN, fontweight="bold")
    ax.set_ylim(0, new * 1.25)
    ax.set_title(label, fontsize=10, fontweight="bold")
fig.suptitle("Illustrative Estimate Revisions", fontweight="bold", y=1.01)
save(fig, "wdc_chart10_estimate_revisions.png")


# 11. One-year price history (dynamic Yahoo data)
fig, ax = plt.subplots(figsize=(9.2, 4.7))
try:
    history = yf.download("WDC", period="1y", auto_adjust=True, progress=False)
    close = history["Close"]
    if getattr(close, "ndim", 1) > 1:
        close = close.iloc[:, 0]
    ax.plot(close.index, close.values, color=NAVY, linewidth=2.1)
    ax.axvline(np.datetime64("2026-08-05"), color=RED, linestyle="--", linewidth=1.1, label="Q4 FY26 release")
    ax.scatter(close.index[-1], close.iloc[-1], color=ORANGE, s=35, zorder=4)
    ax.annotate(f"${close.iloc[-1]:.2f}", (close.index[-1], close.iloc[-1]), xytext=(-42, 12), textcoords="offset points", fontsize=8.5, fontweight="bold")
    ax.legend(loc="upper left", frameon=False)
except Exception as exc:
    print(f"WARNING: price-history retrieval failed: {exc}")
    dates = np.array(["2026-08-04", "2026-08-05", "2026-08-06", "2026-09-15"], dtype="datetime64[D]")
    prices = [548.395, 519.014, 451.385, 411.960]
    ax.plot(dates, prices, marker="o", color=NAVY, linewidth=2.1)
ax.set_ylabel("Adjusted share price ($)")
ax.set_title("WDC One-Year Share Price", fontweight="bold")
save(fig, "wdc_chart11_price_history.png")

print(f"Generated 11 charts in {OUT}")
