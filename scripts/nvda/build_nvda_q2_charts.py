#!/usr/bin/env python3
"""Generate the Q2 FY2027 chart pack for NVIDIA's earnings update."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "NVDA"
DATA = ROOT / "data" / "nvda"
OUT.mkdir(parents=True, exist_ok=True)

GREEN = "#76B900"
DARK_GREEN = "#416900"
PALE_GREEN = "#CDE6A0"
NAVY = "#17365D"
BLUE = "#2F5597"
LIGHT_BLUE = "#9EC5E5"
ORANGE = "#E67E22"
RED = "#C0392B"
GREY = "#667085"
LIGHT_GREY = "#D9DEE7"
INK = "#182230"

plt.rcParams.update(
    {
        "font.family": "serif",
        "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
        "axes.unicode_minus": False,
        "figure.dpi": 150,
        "savefig.dpi": 150,
        "savefig.bbox": "tight",
        "axes.grid": True,
        "grid.alpha": 0.22,
        "grid.color": "#AAB2C0",
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.titleweight": "bold",
        "axes.titlesize": 12.5,
        "axes.labelsize": 10,
        "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.5,
    }
)

quarters = ["Q3\nFY25", "Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26", "Q1\nFY27", "Q2\nFY27"]
revenue = np.array([35.082, 39.331, 44.062, 46.743, 57.006, 68.127, 81.615, 96.221])
dc_revenue = np.array([30.771, 35.580, 39.112, 41.096, 51.215, 62.300, 75.246, 89.023])
gaap_eps = np.array([0.78, 0.89, 0.76, 1.08, 1.30, 1.76, 2.39, 2.46])
non_gaap_eps = np.array([0.81, 0.89, 0.81, 1.05, 1.30, 1.62, 1.87, 2.22])
gaap_gm = np.array([74.6, 73.0, 60.5, 72.4, 73.4, 75.0, 74.9, 75.0])
non_gaap_gm = np.array([75.0, 73.5, 61.0, 72.7, 73.6, 75.2, 75.0, 75.0])


def save(fig, filename: str) -> None:
    fig.savefig(OUT / filename, facecolor="white")
    plt.close(fig)


def value_labels(ax, bars, fmt, dy=0.6, color=INK, size=8.0) -> None:
    for bar in bars:
        value = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_y() + value + dy,
            fmt(value),
            ha="center",
            va="bottom",
            fontsize=size,
            color=color,
            fontweight="bold",
        )


# Figure 1: quarterly revenue progression.
fig, ax = plt.subplots(figsize=(9.6, 4.2))
colors = [GREEN] * 7 + [DARK_GREEN]
bars = ax.bar(quarters, revenue, color=colors, width=0.66, edgecolor="white")
value_labels(ax, bars, lambda v: f"${v:.1f}B", dy=1.2)
ax.set_ylim(0, 108)
ax.set_ylabel("Revenue ($bn)")
ax.set_title("Quarterly Revenue Has More Than Doubled Year over Year")
ax.text(7, 88, "+106% YoY", ha="center", va="center", color="white", fontsize=10, fontweight="bold")
save(fig, "nvda_q2_chart1_quarterly_revenue.png")


# Figure 2: Data Center revenue and mix.
fig, ax1 = plt.subplots(figsize=(9.6, 4.2))
bars = ax1.bar(quarters, dc_revenue, color=GREEN, width=0.66, alpha=0.9, edgecolor="white", label="Data Center revenue")
value_labels(ax1, bars, lambda v: f"${v:.1f}B", dy=1.0, size=7.8)
ax1.set_ylim(0, 100)
ax1.set_ylabel("Data Center revenue ($bn)")
mix = dc_revenue / revenue * 100
ax2 = ax1.twinx()
ax2.plot(quarters, mix, color=ORANGE, marker="D", linewidth=2.2, label="Share of total revenue")
ax2.set_ylim(82, 95)
ax2.set_ylabel("Share of total revenue (%)", color=ORANGE)
ax2.spines["right"].set_visible(True)
ax2.grid(False)
for i, value in enumerate(mix):
    ax2.annotate(f"{value:.0f}%", (i, value), xytext=(0, 8), textcoords="offset points", ha="center", fontsize=7.5, color=ORANGE)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", frameon=False, fontsize=8.5)
ax1.set_title("Data Center Reached $89.0B and 92.5% of Revenue")
save(fig, "nvda_q2_chart2_data_center.png")


# Figure 3: EPS progression with methodology marker.
fig, ax = plt.subplots(figsize=(9.6, 4.2))
x = np.arange(len(quarters))
width = 0.35
b1 = ax.bar(x - width / 2, gaap_eps, width, color=BLUE, label="GAAP EPS", edgecolor="white")
b2 = ax.bar(x + width / 2, non_gaap_eps, width, color=GREEN, label="Non-GAAP EPS", edgecolor="white")
for bars in (b1, b2):
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.04, f"${bar.get_height():.2f}", ha="center", fontsize=7.2)
ax.set_xticks(x, quarters)
ax.set_ylim(0, 2.85)
ax.set_ylabel("Diluted EPS ($)")
ax.legend(loc="upper left", frameon=False, ncol=2, fontsize=8.5)
ax.set_title("Q2 Non-GAAP EPS Rose 120% YoY to $2.22")
ax.text(4.7, 0.18, "FY27 non-GAAP includes stock-based compensation", fontsize=7.5, color=GREY)
save(fig, "nvda_q2_chart3_eps.png")


# Figure 4: gross-margin recovery.
fig, ax = plt.subplots(figsize=(9.6, 4.2))
ax.plot(quarters, gaap_gm, "o-", color=BLUE, linewidth=2.2, label="GAAP gross margin")
ax.plot(quarters, non_gaap_gm, "s--", color=GREEN, linewidth=2.2, label="Non-GAAP gross margin")
ax.fill_between(np.arange(len(quarters)), gaap_gm, non_gaap_gm, color=LIGHT_GREY, alpha=0.3)
ax.set_ylim(55, 79)
ax.set_ylabel("Gross margin (%)")
ax.legend(loc="lower right", frameon=False, fontsize=8.5)
ax.annotate("H20 charge", xy=(2, 60.5), xytext=(1.2, 56.8), color=RED, fontsize=8.5, arrowprops={"arrowstyle": "->", "color": RED})
ax.annotate("75.0%", xy=(7, 75.0), xytext=(6.45, 77.0), color=DARK_GREEN, fontsize=9, fontweight="bold", arrowprops={"arrowstyle": "->", "color": DARK_GREEN})
ax.set_title("Gross Margin Recovered to 75% Despite the Product Transition")
save(fig, "nvda_q2_chart4_gross_margin.png")


# Figure 5: reported versus consensus.
fig, axes = plt.subplots(1, 3, figsize=(9.6, 4.2))
metrics = [
    ("Revenue ($bn)", 92.164, 96.221, "+4.4%"),
    ("Non-GAAP EPS ($)", 2.091, 2.220, "+6.2%"),
    ("Data Center ($bn)", 85.860, 89.023, "+3.7%"),
]
for ax, (name, consensus, reported, variance) in zip(axes, metrics):
    bars = ax.bar(["Consensus", "Reported"], [consensus, reported], color=[LIGHT_GREY, GREEN], width=0.62)
    ax.set_title(name, fontsize=9.5)
    ax.grid(axis="x", visible=False)
    ax.set_ylim(0, max(consensus, reported) * 1.22)
    for bar, value in zip(bars, [consensus, reported]):
        label = f"{value:.2f}" if value < 10 else f"{value:.1f}"
        ax.text(bar.get_x() + bar.get_width() / 2, value + max(consensus, reported) * 0.035, label, ha="center", fontsize=8, fontweight="bold")
    ax.text(0.5, max(consensus, reported) * 1.12, variance, ha="center", fontsize=10, color=DARK_GREEN, fontweight="bold")
fig.suptitle("Q2 Beat Across Revenue, EPS and Data Center", fontsize=12.5, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.93))
save(fig, "nvda_q2_chart5_beat_miss.png")


# Figure 6: market-platform mix.
fig, ax = plt.subplots(figsize=(8.8, 4.2))
labels = ["Hyperscale", "AI Clouds, Industrial\n& Enterprise", "Edge Computing"]
values = [48.710, 40.313, 7.198]
colors = [GREEN, BLUE, ORANGE]
bars = ax.barh(labels, values, color=colors, edgecolor="white")
for bar, value in zip(bars, values):
    ax.text(value + 0.8, bar.get_y() + bar.get_height() / 2, f"${value:.1f}B", va="center", fontsize=9, fontweight="bold")
ax.set_xlim(0, 55)
ax.set_xlabel("Revenue ($bn)")
ax.invert_yaxis()
ax.set_title("ACIE Was the Fastest-Growing Market Platform (+25% QoQ)")
save(fig, "nvda_q2_chart6_platform_mix.png")


# Figure 7: guidance versus prior guide and Street.
fig, ax = plt.subplots(figsize=(9.6, 4.2))
labels = ["Q2 prior guide", "Q2 consensus", "Q2 reported", "Q3 pre-result\nconsensus", "Q3 guidance", "Our Q3E"]
values = [91.0, 92.164, 96.221, 105.16, 108.0, 110.0]
colors = [GREY, LIGHT_GREY, GREEN, LIGHT_BLUE, BLUE, DARK_GREEN]
bars = ax.bar(labels, values, color=colors, edgecolor="white", width=0.66)
for bar, value in zip(bars, values):
    ax.text(bar.get_x() + bar.get_width() / 2, value + 1.3, f"${value:.1f}B", ha="center", fontsize=7.8, fontweight="bold")
ax.set_ylim(0, 122)
ax.set_ylabel("Revenue ($bn)")
ax.set_title("Q3 Guidance Again Cleared the Pre-Results Bar")
save(fig, "nvda_q2_chart7_guidance.png")


# Figure 8: working-capital and cash conversion.
fig, ax = plt.subplots(figsize=(9.6, 4.2))
labels = ["Accounts\nreceivable", "Inventory", "Operating cash\nflow", "Free cash\nflow"]
q1 = [40.710, 25.797, 50.344, 48.554]
q2 = [63.059, 31.575, 24.077, 21.341]
x = np.arange(len(labels))
width = 0.34
b1 = ax.bar(x - width / 2, q1, width, color=LIGHT_GREY, label="Q1 FY27")
b2 = ax.bar(x + width / 2, q2, width, color=GREEN, label="Q2 FY27")
for bars in (b1, b2):
    value_labels(ax, bars, lambda v: f"${v:.1f}B", dy=1.0, size=7.5)
ax.set_xticks(x, labels)
ax.set_ylim(0, 72)
ax.set_ylabel("$bn")
ax.legend(loc="upper right", frameon=False, fontsize=8.5)
ax.set_title("Working Capital Absorbed Cash as Receivables and Inventory Rose")
save(fig, "nvda_q2_chart8_working_capital.png")


# Figure 9: strategic commitments and guarantees.
fig, ax = plt.subplots(figsize=(9.6, 4.2))
labels = ["Supply &\ncapacity", "Cloud service\nagreement", "Data-center\nleases", "Equity\ninvestments", "Capital\nexpenditure", "Guarantees\n(max exposure)"]
values = [279, 29, 25, 25, 8, 108.5]
colors = [GREEN, BLUE, LIGHT_BLUE, ORANGE, GREY, RED]
bars = ax.bar(labels, values, color=colors, edgecolor="white", width=0.65)
for bar, value in zip(bars, values):
    ax.text(bar.get_x() + bar.get_width() / 2, value + 6, f"${value:g}B", ha="center", fontsize=7.8, fontweight="bold")
ax.axhline(119, color=DARK_GREEN, linestyle="--", linewidth=1.3, label="Prior-quarter supply commitment: $119B")
ax.set_ylim(0, 320)
ax.set_ylabel("Maximum / future commitment ($bn)")
ax.legend(loc="upper right", frameon=False, fontsize=8.2)
ax.set_title("Scale of Commitments Is Now a Core Part of the Risk/Reward")
save(fig, "nvda_q2_chart9_commitments.png")


# Figure 10: estimate revisions.
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.2))
for ax, title, old, new, unit in [
    (axes[0], "FY2027E revenue", 365.0, 411.8, "revenue"),
    (axes[1], "FY2027E non-GAAP EPS", 8.80, 9.39, "$"),
]:
    bars = ax.bar(["Pre-Q2 baseline", "Post-Q2 estimate"], [old, new], color=[LIGHT_GREY, GREEN], width=0.62)
    ax.set_title(title, fontsize=10.5)
    ax.set_ylim(0, new * 1.28)
    for bar, value in zip(bars, [old, new]):
        text = f"${value:.1f}B" if unit == "revenue" else f"${value:.2f}"
        ax.text(bar.get_x() + bar.get_width() / 2, value + new * 0.04, text, ha="center", fontsize=8.5, fontweight="bold")
    ax.text(0.5, new * 1.18, f"+{(new / old - 1) * 100:.1f}%", ha="center", color=DARK_GREEN, fontsize=10, fontweight="bold")
    ax.grid(axis="x", visible=False)
fig.suptitle("Q2 Beat and FY2028 Framework Drive Estimate Upgrades", fontsize=12.5, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.93))
save(fig, "nvda_q2_chart10_estimate_revisions.png")


# Figure 11: one-year share price, dynamically sourced from Longbridge.
price_path = DATA / "longbridge_kline_1y.json"
if price_path.exists():
    payload = json.loads(price_path.read_text())
    dates = [datetime.fromisoformat(row["time"].replace("Z", "+00:00")) for row in payload]
    closes = [float(row["close"]) for row in payload]
    fig, ax = plt.subplots(figsize=(9.6, 4.2))
    ax.plot(dates, closes, color=NAVY, linewidth=2.0)
    ax.fill_between(dates, closes, min(closes) * 0.92, color=LIGHT_BLUE, alpha=0.24)
    ax.axhline(320, color=GREEN, linestyle="--", linewidth=1.6, label="Price target: $320")
    ax.scatter(dates[-1], closes[-1], color=ORANGE, s=42, zorder=5)
    ax.annotate(f"Aug. 26 close: ${closes[-1]:.2f}", (dates[-1], closes[-1]), xytext=(-110, -28), textcoords="offset points", fontsize=8.5, color=ORANGE, arrowprops={"arrowstyle": "->", "color": ORANGE})
    ax.set_ylim(min(closes) * 0.92, 335)
    ax.set_ylabel("Adjusted close ($)")
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.legend(loc="upper left", frameon=False, fontsize=8.5)
    ax.set_title("One-Year Price History Still Leaves Material Upside to Our Base Case")
    save(fig, "nvda_q2_chart11_price_history.png")
else:
    raise FileNotFoundError(price_path)

print("Generated 11 NVIDIA Q2 FY2027 charts in", OUT)
