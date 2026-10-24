#!/usr/bin/env python3
"""Generate the chart pack for Western Digital's Q1 FY2027 earnings update."""

from __future__ import annotations

import json
import os
import subprocess
from datetime import datetime
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "WDC"
OUT.mkdir(parents=True, exist_ok=True)
os.environ.setdefault("MPLCONFIGDIR", "/tmp/wdc-mpl-cache")

NAVY = "#123B5D"
BLUE = "#2F6B9A"
LIGHT_BLUE = "#A9CBE3"
TEAL = "#008A8A"
GREEN = "#2F7D32"
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
        "axes.titlesize": 12.0,
        "axes.labelsize": 9.3,
        "xtick.labelsize": 8.2,
        "ytick.labelsize": 8.2,
    }
)

quarters = ["Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26", "Q1\nFY27"]
revenue = np.array([2.818, 3.017, 3.337, 3.747, 4.381])
non_gaap_eps = np.array([1.78, 2.13, 2.72, 3.56, 4.64])
gaap_eps = np.array([3.07, 4.70, 8.20, 8.21, 4.63])
non_gaap_gm = np.array([43.9, 46.1, 50.5, 54.4, 58.1])
non_gaap_opm = np.array([30.4, 33.8, 38.6, 44.2, 49.1])
fcf = np.array([0.599, 0.653, 0.978, 1.281, 1.377])
nearline = np.array([183, 192, 199, 209, 223])
non_nearline = np.array([21, 23, 23, 22, 25])


def save(fig, filename: str) -> None:
    fig.savefig(OUT / filename, facecolor="white")
    plt.close(fig)


def labels(ax, bars, formatter, pad=0.035, size=8.0) -> None:
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


def quote() -> dict:
    proc = subprocess.run(
        ["longbridge", "quote", "WDC.US", "--format", "json"],
        check=True,
        capture_output=True,
        text=True,
        timeout=30,
    )
    return json.loads(proc.stdout)[0]


def kline() -> list[dict]:
    proc = subprocess.run(
        [
            "longbridge",
            "kline",
            "WDC.US",
            "--period",
            "day",
            "--count",
            "260",
            "--adjust",
            "forward",
            "--format",
            "json",
        ],
        check=True,
        capture_output=True,
        text=True,
        timeout=45,
    )
    return json.loads(proc.stdout)


# Figure 1: quarterly revenue.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
bars = ax.bar(quarters, revenue, color=[LIGHT_BLUE] * 4 + [NAVY], width=0.66, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.2f}B")
ax.set_ylim(0, 4.95)
ax.set_ylabel("Revenue ($bn)")
ax.set_title("Revenue Reached $4.38B as Cloud Demand and Pricing Accelerated")
ax.text(4, 4.0, "+55% YoY", ha="center", color="white", fontsize=9.2, fontweight="bold")
save(fig, "wdc_q1fy27_chart1_revenue.png")


# Figure 2: EPS progression.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
x = np.arange(len(quarters))
width = 0.36
bars1 = ax.bar(x - width / 2, non_gaap_eps, width, color=TEAL, label="Non-GAAP EPS")
bars2 = ax.bar(x + width / 2, gaap_eps, width, color=LIGHT_BLUE, label="GAAP EPS")
for bars in (bars1, bars2):
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.18, f"${bar.get_height():.2f}", ha="center", fontsize=7.3)
ax.set_xticks(x, quarters)
ax.set_ylim(0, 9.4)
ax.set_ylabel("Diluted EPS ($)")
ax.legend(frameon=False, ncol=2, loc="upper left")
ax.set_title("Q1 Non-GAAP EPS Rose 161% YoY to $4.64")
save(fig, "wdc_q1fy27_chart2_eps.png")


# Figure 3: margins.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.plot(quarters, non_gaap_gm, marker="o", linewidth=2.4, color=NAVY, label="Gross margin")
ax.plot(quarters, non_gaap_opm, marker="s", linewidth=2.4, color=TEAL, label="Operating margin")
for i, value in enumerate(non_gaap_gm):
    ax.annotate(f"{value:.1f}%", (i, value), xytext=(0, 7), textcoords="offset points", ha="center", fontsize=7.8, color=NAVY)
for i, value in enumerate(non_gaap_opm):
    ax.annotate(f"{value:.1f}%", (i, value), xytext=(0, -13), textcoords="offset points", ha="center", fontsize=7.8, color=TEAL)
ax.set_ylim(25, 63)
ax.set_ylabel("Margin (%)")
ax.legend(frameon=False, loc="upper left", ncol=2)
ax.set_title("Pricing and Higher-Capacity Mix Drove Another Margin Step-Up")
save(fig, "wdc_q1fy27_chart3_margins.png")


# Figure 4: free cash flow.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
bars = ax.bar(quarters, fcf, color=[LIGHT_BLUE] * 4 + [GREEN], width=0.66, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.2f}B")
ax.set_ylim(0, 1.6)
ax.set_ylabel("Free cash flow ($bn)")
ax.set_title("Free Cash Flow Rose 130% YoY to $1.38B")
save(fig, "wdc_q1fy27_chart4_fcf.png")


# Figure 5: end-market mix.
cloud = np.array([89, 89, 89, 89, 88])
client = np.array([5, 6, 5, 6, 7])
consumer = np.array([6, 5, 6, 5, 5])
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.bar(quarters, cloud, color=NAVY, label="Cloud")
ax.bar(quarters, client, bottom=cloud, color=TEAL, label="Client")
ax.bar(quarters, consumer, bottom=cloud + client, color=ORANGE, label="Consumer")
for i, value in enumerate(cloud):
    ax.text(i, value / 2, f"{value}%", ha="center", va="center", color="white", fontsize=9, fontweight="bold")
ax.set_ylim(0, 100)
ax.set_ylabel("Revenue mix (%)")
ax.legend(frameon=False, loc="lower center", bbox_to_anchor=(0.5, -0.23), ncol=3)
ax.set_title("Cloud Remained 88% of Revenue Despite Faster Client Growth")
save(fig, "wdc_q1fy27_chart5_mix.png")


# Figure 6: exabytes.
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.bar(quarters, nearline, color=NAVY, label="Nearline")
ax.bar(quarters, non_nearline, bottom=nearline, color=LIGHT_BLUE, label="Non-nearline")
for i, total in enumerate(nearline + non_nearline):
    ax.text(i, total + 5, f"{total} EB", ha="center", fontsize=8.1, fontweight="bold")
ax.set_ylim(0, 275)
ax.set_ylabel("Exabytes shipped")
ax.legend(frameon=False, loc="upper left", ncol=2)
ax.set_title("Exabytes Shipped Increased 22% YoY and 8% Sequentially")
save(fig, "wdc_q1fy27_chart6_exabytes.png")


# Figure 7: actual vs. prior guidance and consensus.
fig, axes = plt.subplots(1, 3, figsize=(9.6, 4.0))
comparisons = [
    ("Revenue ($bn)", [4.10, 4.145, 4.381], ["Guide mid", "Street", "Actual"]),
    ("Gross margin (%)", [55.5, 58.1], ["Guide mid", "Actual"]),
    ("Non-GAAP EPS ($)", [4.00, 4.098, 4.64], ["Guide mid", "Street", "Actual"]),
]
for ax, (title, values, names) in zip(axes, comparisons):
    colors = [LIGHT_GREY] * (len(values) - 1) + [GREEN]
    bars = ax.bar(names, values, color=colors, width=0.62)
    ax.set_ylim(0, max(values) * 1.25)
    ax.set_title(title, fontsize=9.5)
    ax.grid(axis="x", visible=False)
    ax.tick_params(axis="x", labelrotation=15)
    for bar, value in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, value + max(values) * 0.035, f"{value:.2f}", ha="center", fontsize=7.7, fontweight="bold")
fig.suptitle("Q1 Revenue, Margin and EPS Cleared Prior Guidance", fontsize=12.0, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.92))
save(fig, "wdc_q1fy27_chart7_beat.png")


# Figure 8: Q2 management guide vs. current consensus.
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.0))
for ax, title, guide, street, formatter in [
    (axes[0], "Q2 FY2027 revenue", 5.10, 5.127, lambda v: f"${v:.2f}B"),
    (axes[1], "Q2 FY2027 non-GAAP EPS", 5.80, 5.873, lambda v: f"${v:.2f}"),
]:
    bars = ax.bar(["Guide midpoint", "Street (23 Oct.)"], [guide, street], color=[NAVY, LIGHT_BLUE], width=0.62)
    ax.set_ylim(0, street * 1.25)
    ax.set_title(title, fontsize=10.0)
    ax.grid(axis="x", visible=False)
    for bar, value in zip(bars, [guide, street]):
        ax.text(bar.get_x() + bar.get_width() / 2, value + street * 0.035, formatter(value), ha="center", fontsize=8.3, fontweight="bold")
fig.suptitle("Q2 Guidance Anchors a Much Higher Earnings Run-Rate", fontsize=12.0, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.92))
save(fig, "wdc_q1fy27_chart8_guidance.png")


# Figure 9: estimate revisions.
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.0))
for ax, title, old, new, formatter in [
    (axes[0], "FY2027E revenue", 19.6, 21.3, lambda v: f"${v:.1f}B"),
    (axes[1], "FY2027E non-GAAP EPS", 20.50, 25.00, lambda v: f"${v:.2f}"),
]:
    bars = ax.bar(["Prior", "Updated"], [old, new], color=[LIGHT_GREY, TEAL], width=0.62)
    ax.set_ylim(0, new * 1.28)
    ax.set_title(title, fontsize=10.2)
    ax.grid(axis="x", visible=False)
    for bar, value in zip(bars, [old, new]):
        ax.text(bar.get_x() + bar.get_width() / 2, value + new * 0.035, formatter(value), ha="center", fontsize=8.4, fontweight="bold")
    ax.text(0.5, new * 1.18, f"+{(new / old - 1) * 100:.1f}%", ha="center", fontsize=9.3, color=GREEN, fontweight="bold")
fig.suptitle("Q1 Beat and Q2 Guide Lift the FY2027 Base Case", fontsize=12.0, fontweight="bold")
fig.tight_layout(rect=(0, 0, 1, 0.92))
save(fig, "wdc_q1fy27_chart9_revisions.png")


# Figure 10: valuation scenarios.
reference = float(quote()["last"])
fig, ax = plt.subplots(figsize=(9.6, 4.0))
scenario_names = ["Bear", "Base", "Bull"]
scenario_values = [420, 620, 800]
bars = ax.bar(scenario_names, scenario_values, color=[RED, NAVY, GREEN], width=0.58, edgecolor="white")
labels(ax, bars, lambda value: f"${value:.0f}", pad=0.025, size=9)
ax.axhline(reference, color=ORANGE, linestyle="--", linewidth=1.6, label=f"23 Oct. close: ${reference:.2f}")
ax.set_ylim(0, 900)
ax.set_ylabel("Implied value per share ($)")
ax.legend(frameon=False, loc="upper left")
ax.set_title("Scenario Values Span $420 to $800 per Share")
save(fig, "wdc_q1fy27_chart10_valuation.png")


# Figure 11: one-year price history.
rows = kline()
dates = [datetime.fromisoformat(row["time"].replace("Z", "+00:00")) for row in rows]
closes = [float(row["close"]) for row in rows]
fig, ax = plt.subplots(figsize=(9.6, 4.0))
ax.plot(dates, closes, color=NAVY, linewidth=2.0)
ax.fill_between(dates, closes, min(closes) * 0.90, color=LIGHT_BLUE, alpha=0.25)
ax.axhline(620, color=GREEN, linestyle="--", linewidth=1.6, label="Price target: $620")
release_dt = datetime.fromisoformat("2026-10-22T04:00:00+00:00")
ax.axvline(release_dt, color=ORANGE, linestyle=":", linewidth=1.4, label="Q1 FY2027 release")
ax.scatter(dates[-1], closes[-1], color=ORANGE, s=42, zorder=5)
ax.annotate(f"23 Oct.: ${closes[-1]:.2f}", (dates[-1], closes[-1]), xytext=(-95, -24), textcoords="offset points", fontsize=8.2, color=ORANGE, arrowprops={"arrowstyle": "->", "color": ORANGE})
ax.set_ylim(min(closes) * 0.90, max(700, max(closes) * 1.08))
ax.set_ylabel("Adjusted close ($)")
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
ax.legend(frameon=False, loc="upper left", fontsize=8.1)
ax.set_title("The Post-Earnings Selloff Expanded Upside to the Updated Target")
save(fig, "wdc_q1fy27_chart11_price.png")

print("Generated 11 WDC Q1 FY2027 charts in", OUT)
