#!/usr/bin/env python3
"""Generate the chart pack for Western Digital's Q4 FY2026 earnings update."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "WDC"
DATA = ROOT / "data" / "wdc"
OUT.mkdir(parents=True, exist_ok=True)

NAVY = "#123B5D"
BLUE = "#2F6B9A"
LIGHT_BLUE = "#A9CBE3"
TEAL = "#008A8A"
GREEN = "#2F7D32"
PALE_GREEN = "#B9D9B0"
ORANGE = "#D97706"
RED = "#B42318"
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
        "axes.titlesize": 12.2,
        "axes.labelsize": 9.5,
        "xtick.labelsize": 8.3,
        "ytick.labelsize": 8.3,
    }
)

quarters = ["Q3\nFY25", "Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26"]
revenue = np.array([2.294, 2.605, 2.818, 3.017, 3.337, 3.747])
non_gaap_eps = np.array([1.38, 1.70, 1.78, 2.13, 2.72, 3.56])
non_gaap_gm = np.array([40.1, 41.3, 43.9, 46.1, 50.5, 54.4])
non_gaap_opm = np.array([26.0, 28.1, 30.4, 33.8, 38.6, 44.2])
fcf = np.array([0.436, 0.675, 0.599, 0.653, 0.978, 1.281])


def save(fig, filename: str) -> None:
    fig.savefig(OUT / filename, facecolor="white")
    plt.close(fig)


def labels(ax, bars, formatter, pad=0.04, size=8.0) -> None:
    top = max(bar.get_height() for bar in bars)
    for bar in bars:
        value = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + top * pad,
            formatter(value),
            ha="center",
            va="bottom",
            fontsize=size,
            color=INK,
            fontweight="bold",
        )


# Figure 1: quarterly revenue.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
bars = ax.bar(quarters, revenue, color=[LIGHT_BLUE] * 5 + [NAVY], width=0.66, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.2f}B", pad=0.035)
ax.set_ylim(0, 4.25)
ax.set_ylabel("Revenue ($bn)")
ax.set_title("Revenue Reached $3.75B as Cloud Demand and Pricing Improved")
ax.text(5, 3.38, "+44% YoY", ha="center", color="white", fontsize=9.5, fontweight="bold")
save(fig, "wdc_q4_chart1_quarterly_revenue.png")


# Figure 2: non-GAAP EPS progression.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
bars = ax.bar(quarters, non_gaap_eps, color=[LIGHT_BLUE] * 5 + [TEAL], width=0.66, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.2f}", pad=0.035)
ax.set_ylim(0, 4.05)
ax.set_ylabel("Non-GAAP diluted EPS ($)")
ax.set_title("Non-GAAP EPS More Than Doubled Year over Year")
ax.text(5, 3.15, "+109% YoY", ha="center", color="white", fontsize=9.2, fontweight="bold")
save(fig, "wdc_q4_chart2_eps.png")


# Figure 3: margin progression.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.plot(quarters, non_gaap_gm, marker="o", linewidth=2.4, color=NAVY, label="Gross margin")
ax.plot(quarters, non_gaap_opm, marker="s", linewidth=2.4, color=TEAL, label="Operating margin")
for i, value in enumerate(non_gaap_gm):
    ax.annotate(f"{value:.1f}%", (i, value), xytext=(0, 7), textcoords="offset points", ha="center", fontsize=7.8, color=NAVY)
for i, value in enumerate(non_gaap_opm):
    ax.annotate(f"{value:.1f}%", (i, value), xytext=(0, -13), textcoords="offset points", ha="center", fontsize=7.8, color=TEAL)
ax.set_ylim(20, 60)
ax.set_ylabel("Margin (%)")
ax.legend(frameon=False, loc="upper left", ncol=2)
ax.set_title("Pricing and Mix Expanded Both Gross and Operating Margins")
save(fig, "wdc_q4_chart3_margins.png")


# Figure 4: free cash flow.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
bars = ax.bar(quarters, fcf, color=[LIGHT_BLUE] * 5 + [GREEN], width=0.66, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.2f}B", pad=0.035)
ax.set_ylim(0, 1.48)
ax.set_ylabel("Free cash flow ($bn)")
ax.set_title("Quarterly Free Cash Flow Reached $1.28B and a 34% Margin")
save(fig, "wdc_q4_chart4_free_cash_flow.png")


# Figure 5: end-market mix.
mix_quarters = ["Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26"]
cloud = np.array([90, 89, 89, 89, 89])
client = np.array([5, 5, 6, 5, 6])
consumer = np.array([5, 6, 5, 6, 5])
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.bar(mix_quarters, cloud, color=NAVY, label="Cloud")
ax.bar(mix_quarters, client, bottom=cloud, color=TEAL, label="Client")
ax.bar(mix_quarters, consumer, bottom=cloud + client, color=ORANGE, label="Consumer")
for i in range(len(mix_quarters)):
    ax.text(i, cloud[i] / 2, f"{cloud[i]}%", ha="center", va="center", color="white", fontsize=9, fontweight="bold")
ax.set_ylim(0, 100)
ax.set_ylabel("Revenue mix (%)")
ax.legend(frameon=False, loc="lower center", bbox_to_anchor=(0.5, -0.22), ncol=3)
ax.set_title("Cloud Has Held Near 89% of Revenue for Four Consecutive Quarters")
save(fig, "wdc_q4_chart5_end_market_mix.png")


# Figure 6: exabyte shipments.
metric_quarters = ["Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26"]
nearline = np.array([170, 183, 192, 199, 209])
non_nearline = np.array([20, 21, 23, 23, 22])
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.bar(metric_quarters, nearline, color=NAVY, label="Nearline")
ax.bar(metric_quarters, non_nearline, bottom=nearline, color=LIGHT_BLUE, label="Non-nearline")
for i, total in enumerate(nearline + non_nearline):
    ax.text(i, total + 5, f"{total} EB", ha="center", fontsize=8.2, fontweight="bold")
ax.set_ylim(0, 260)
ax.set_ylabel("Exabytes shipped")
ax.legend(frameon=False, loc="upper left", ncol=2)
ax.set_title("Exabytes Shipped Rose 22% YoY Despite Quarterly Mix Variability")
save(fig, "wdc_q4_chart6_exabytes.png")


# Figure 7: result and guidance comparisons.
fig, axes = plt.subplots(1, 3, figsize=(9.6, 4.0))
comparisons = [
    ("Q4 revenue ($bn)", 3.65, 3.747, "Prior guide", "Actual"),
    ("Q4 gross margin (%)", 51.5, 54.4, "Prior midpoint", "Actual"),
    ("Q4 non-GAAP EPS ($)", 3.25, 3.56, "Prior midpoint", "Actual"),
]
for ax, (title, old, actual, old_label, new_label) in zip(axes, comparisons):
    bars = ax.bar([old_label, new_label], [old, actual], color=[LIGHT_GREY, GREEN], width=0.62)
    ax.set_ylim(0, actual * 1.27)
    ax.set_title(title, fontsize=9.5)
    ax.grid(axis="x", visible=False)
    for bar, value in zip(bars, [old, actual]):
        fmt = f"{value:.3f}" if title.startswith("Q4 revenue") else f"{value:.1f}"
        ax.text(bar.get_x() + bar.get_width() / 2, value + actual * 0.035, fmt, ha="center", fontsize=8.2, fontweight="bold")
fig.suptitle("Q4 Revenue, Margin and EPS All Exceeded Prior Guidance Midpoints", fontsize=12.2, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.92))
save(fig, "wdc_q4_chart7_prior_guidance.png")


# Figure 8: pricing versus cost per terabyte.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
values = [17.5, -8.0]
bars = ax.barh(["Blended price per TB", "Cost per TB"], values, color=[GREEN, BLUE], height=0.55)
ax.axvline(0, color=GREY, linewidth=1)
ax.text(values[0] + 0.7, bars[0].get_y() + bars[0].get_height() / 2, "high-teens YoY", va="center", fontsize=9, fontweight="bold", color=GREEN)
ax.text(values[1] - 0.7, bars[1].get_y() + bars[1].get_height() / 2, "-8% YoY", va="center", ha="right", fontsize=9, fontweight="bold", color=BLUE)
ax.set_xlim(-13, 25)
ax.set_xlabel("Approximate year-over-year change (%)")
ax.set_title("Price per Terabyte Rose While Cost per Terabyte Declined")
save(fig, "wdc_q4_chart8_price_cost.png")


# Figure 9: estimate revisions.
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.0))
for ax, title, old, new, formatter in [
    (axes[0], "FY2027E revenue", 18.1, 19.6, lambda v: f"${v:.1f}B"),
    (axes[1], "FY2027E non-GAAP EPS", 17.50, 20.50, lambda v: f"${v:.2f}"),
]:
    bars = ax.bar(["Pre-Q4 baseline", "Post-Q4 estimate"], [old, new], color=[LIGHT_GREY, TEAL], width=0.62)
    ax.set_ylim(0, new * 1.28)
    ax.set_title(title, fontsize=10.3)
    ax.grid(axis="x", visible=False)
    for bar, value in zip(bars, [old, new]):
        ax.text(bar.get_x() + bar.get_width() / 2, value + new * 0.035, formatter(value), ha="center", fontsize=8.4, fontweight="bold")
    ax.text(0.5, new * 1.18, f"+{(new / old - 1) * 100:.1f}%", ha="center", fontsize=9.5, color=GREEN, fontweight="bold")
fig.suptitle("Stronger Pricing and Q1 Guidance Raise the FY2027 Base Case", fontsize=12.2, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.92))
save(fig, "wdc_q4_chart9_estimate_revisions.png")


# Figure 10: valuation scenarios.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
scenario_names = ["Bear", "Base", "Bull"]
scenario_values = [360, 600, 820]
bars = ax.bar(scenario_names, scenario_values, color=[RED, NAVY, GREEN], width=0.58, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.0f}", pad=0.025, size=9)
ax.axhline(415.82, color=ORANGE, linestyle="--", linewidth=1.6, label="Reference price: $415.82")
ax.set_ylim(0, 930)
ax.set_ylabel("Implied value per share ($)")
ax.legend(frameon=False, loc="upper left")
ax.set_title("Scenario Values Span $360 to $820 per Share")
save(fig, "wdc_q4_chart10_valuation_scenarios.png")


# Figure 11: one-year price history, supplied by Longbridge.
price_path = DATA / "longbridge_kline_1y.json"
if not price_path.exists():
    raise FileNotFoundError(price_path)
rows = json.loads(price_path.read_text())
dates = [datetime.fromisoformat(row["time"].replace("Z", "+00:00")) for row in rows]
closes = [float(row["close"]) for row in rows]
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.plot(dates, closes, color=NAVY, linewidth=2.0)
ax.fill_between(dates, closes, min(closes) * 0.90, color=LIGHT_BLUE, alpha=0.25)
ax.axhline(600, color=GREEN, linestyle="--", linewidth=1.6, label="Price target: $600")
ax.axvline(datetime.fromisoformat("2026-08-05T04:00:00+00:00"), color=ORANGE, linestyle=":", linewidth=1.4, label="Q4 release")
ax.scatter(dates[-1], closes[-1], color=ORANGE, s=42, zorder=5)
ax.annotate(f"15 Sep.: ${closes[-1]:.2f}", (dates[-1], closes[-1]), xytext=(-95, -24), textcoords="offset points", fontsize=8.2, color=ORANGE, arrowprops={"arrowstyle": "->", "color": ORANGE})
ax.set_ylim(min(closes) * 0.90, max(850, max(closes) * 1.08))
ax.set_ylabel("Adjusted close ($)")
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
ax.legend(frameon=False, loc="upper left", fontsize=8.3)
ax.set_title("The Post-Earnings Drawdown Has Reset the Risk-Reward")
save(fig, "wdc_q4_chart11_price_history.png")

print("Generated 11 WDC Q4 FY2026 charts in", OUT)
