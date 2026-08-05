#!/usr/bin/env python3
"""Generate the 11 figures used in the SPCX Q2 2026 earnings update."""

from __future__ import annotations

import os
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import yfinance as yf


OUT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/SPCX")
OUT.mkdir(parents=True, exist_ok=True)

plt.rcParams.update(
    {
        "font.family": "serif",
        "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
        "axes.unicode_minus": False,
        "figure.dpi": 150,
        "savefig.dpi": 150,
        "savefig.bbox": "tight",
        "axes.grid": True,
        "grid.alpha": 0.18,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.titleweight": "bold",
    }
)

NAVY = "#0B2545"
BLUE = "#005288"
CYAN = "#00A3E0"
ORANGE = "#E87722"
PURPLE = "#7B2D8E"
GREEN = "#2E7D32"
RED = "#C62828"
GOLD = "#B7791F"
GRAY = "#6B7280"
LIGHT = "#DCE6F1"


def finish(fig: plt.Figure, name: str) -> None:
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(OUT / name)
    plt.close(fig)


def label_bars(ax, bars, fmt="{:.1f}", dy=0.03, color=NAVY):
    for bar in bars:
        value = bar.get_height()
        offset = dy if value >= 0 else -dy
        va = "bottom" if value >= 0 else "top"
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + offset,
            fmt.format(value),
            ha="center",
            va=va,
            fontsize=9,
            fontweight="bold",
            color=color,
        )


# Figure 1: reported results vs pre-earnings consensus
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.4))
revenue_bars = axes[0].bar(["Consensus", "Reported"], [6.90, 7.814], color=[GRAY, BLUE], width=0.55)
label_bars(axes[0], revenue_bars, "${:.2f}B", dy=0.10)
axes[0].set_title("Q2 Revenue: 13.2% Beat")
axes[0].set_ylabel("USD billions")
axes[0].set_ylim(0, 9)
eps_bars = axes[1].bar(["Consensus", "Reported"], [-0.29, -0.09], color=[GRAY, GREEN], width=0.55)
label_bars(axes[1], eps_bars, "${:.2f}", dy=0.01)
axes[1].axhline(0, color=NAVY, linewidth=0.8)
axes[1].set_title("GAAP EPS: $0.20 Better")
axes[1].set_ylabel("Loss per share")
axes[1].set_ylim(-0.35, 0.05)
fig.suptitle("SpaceX Q2 2026 Beat/Miss Snapshot", fontsize=14, fontweight="bold", color=NAVY)
finish(fig, "spcx_q2_chart1_beat_miss.png")


# Figure 2: segment revenue trend
periods = ["Q2'25", "Q1'26", "Q2'26"]
space_rev = [0.746, 0.619, 0.962]
conn_rev = [2.588, 3.257, 4.291]
ai_rev = [0.737, 0.818, 2.561]
x = np.arange(len(periods))
w = 0.23
fig, ax = plt.subplots(figsize=(9.2, 4.8))
b1 = ax.bar(x - w, space_rev, w, label="Space", color=ORANGE)
b2 = ax.bar(x, conn_rev, w, label="Connectivity", color=CYAN)
b3 = ax.bar(x + w, ai_rev, w, label="AI", color=PURPLE)
for bars in (b1, b2, b3):
    label_bars(ax, bars, "${:.1f}", dy=0.07)
ax.set_xticks(x, periods)
ax.set_ylabel("Revenue (USD billions)")
ax.set_title("Segment Revenue: AI Inflects, Connectivity Compounds")
ax.legend(ncol=3, frameon=False, loc="upper left")
ax.set_ylim(0, 5.2)
finish(fig, "spcx_q2_chart2_segment_revenue.png")


# Figure 3: segment operating income/loss
space_op = [-0.369, -0.662, -0.542]
conn_op = [0.923, 1.188, 1.656]
ai_op = [-1.524, -2.469, -1.257]
fig, ax = plt.subplots(figsize=(9.2, 4.8))
b1 = ax.bar(x - w, space_op, w, label="Space", color=ORANGE)
b2 = ax.bar(x, conn_op, w, label="Connectivity", color=CYAN)
b3 = ax.bar(x + w, ai_op, w, label="AI", color=PURPLE)
for bars in (b1, b2, b3):
    label_bars(ax, bars, "${:.1f}", dy=0.08)
ax.axhline(0, color=NAVY, linewidth=0.9)
ax.set_xticks(x, periods)
ax.set_ylabel("Operating income/(loss), USD billions")
ax.set_title("Connectivity Profit Pool Funds Space and AI Losses")
ax.legend(ncol=3, frameon=False, loc="lower left")
ax.set_ylim(-3.0, 2.2)
finish(fig, "spcx_q2_chart3_segment_operating_income.png")


# Figure 4: Starlink subscribers and ARPU
subs = [6.0, 10.3, 12.0]
arpu = [85, 66, 66]
fig, ax1 = plt.subplots(figsize=(9.2, 4.8))
bars = ax1.bar(periods, subs, color=CYAN, width=0.5, alpha=0.9, label="Subscribers")
label_bars(ax1, bars, "{:.1f}M", dy=0.25)
ax1.set_ylabel("Subscribers (millions)", color=BLUE)
ax1.set_ylim(0, 14)
ax2 = ax1.twinx()
ax2.grid(False)
ax2.plot(periods, arpu, color=GOLD, marker="o", linewidth=2.5, markersize=7, label="ARPU")
for i, value in enumerate(arpu):
    ax2.text(i, value + 2.5, f"${value}", ha="center", fontsize=9, fontweight="bold", color=GOLD)
ax2.set_ylabel("Monthly ARPU (USD)", color=GOLD)
ax2.set_ylim(50, 100)
ax1.set_title("Starlink: Subscribers Double YoY; ARPU Stabilizes QoQ")
finish(fig, "spcx_q2_chart4_starlink_metrics.png")


# Figure 5: adjusted EBITDA and margin
ebitda = [1.214, 1.127, 3.538]
revenue = [4.071, 4.694, 7.814]
margin = [a / b * 100 for a, b in zip(ebitda, revenue)]
fig, ax1 = plt.subplots(figsize=(9.2, 4.8))
bars = ax1.bar(periods, ebitda, color=[LIGHT, LIGHT, GREEN], edgecolor=BLUE, width=0.5)
label_bars(ax1, bars, "${:.1f}B", dy=0.09)
ax1.set_ylabel("Adjusted EBITDA (USD billions)")
ax1.set_ylim(0, 4.3)
ax2 = ax1.twinx()
ax2.grid(False)
ax2.plot(periods, margin, color=PURPLE, marker="o", linewidth=2.5)
for i, value in enumerate(margin):
    ax2.text(i, value + 2.2, f"{value:.1f}%", ha="center", fontsize=9, fontweight="bold", color=PURPLE)
ax2.set_ylabel("Adjusted EBITDA margin")
ax2.set_ylim(0, 55)
ax1.set_title("Adjusted EBITDA Nearly Triples as AI Turns Positive")
finish(fig, "spcx_q2_chart5_ebitda_margin.png")


# Figure 6: AI revenue vs capex
ai_capex = [0.749, 7.723, 15.828]
fig, ax = plt.subplots(figsize=(9.2, 4.8))
b1 = ax.bar(x - 0.17, ai_rev, 0.34, label="AI revenue", color=PURPLE)
b2 = ax.bar(x + 0.17, ai_capex, 0.34, label="AI capex", color=RED, alpha=0.82)
label_bars(ax, b1, "${:.1f}B", dy=0.25)
label_bars(ax, b2, "${:.1f}B", dy=0.25)
ax.set_xticks(x, periods)
ax.set_ylabel("USD billions")
ax.set_ylim(0, 18.5)
ax.set_title("AI Monetization Improves, but Capex Runs 6.2x Revenue")
ax.legend(frameon=False)
finish(fig, "spcx_q2_chart6_ai_revenue_capex.png")


# Figure 7: Q2 capex allocation
capex_values = [1.174, 1.367, 15.828]
fig, ax = plt.subplots(figsize=(7.6, 4.8))
wedges, _, autotexts = ax.pie(
    capex_values,
    labels=["Space", "Connectivity", "AI"],
    colors=[ORANGE, CYAN, PURPLE],
    autopct="%1.0f%%",
    startangle=100,
    wedgeprops={"edgecolor": "white", "linewidth": 2},
    textprops={"fontsize": 10},
)
for t in autotexts:
    t.set_color("white")
    t.set_fontweight("bold")
ax.set_title("Q2 2026 Capex Mix: 86% Allocated to AI")
ax.text(0, -1.17, "Total capex: $18.4B", ha="center", color=NAVY, fontweight="bold")
finish(fig, "spcx_q2_chart7_capex_mix.png")


# Figure 8: H1 cash flow
fig, ax = plt.subplots(figsize=(8.8, 4.8))
cf_labels = ["Operating cash flow", "Capital expenditures", "Free cash flow"]
cf_values = [3.466, -28.476, -25.010]
bars = ax.bar(cf_labels, cf_values, color=[GREEN, RED, RED], width=0.5)
label_bars(ax, bars, "${:.1f}B", dy=0.7)
ax.axhline(0, color=NAVY, linewidth=0.9)
ax.set_ylabel("USD billions")
ax.set_ylim(-32, 8)
ax.set_title("H1 2026: Operating Cash Flow Swamped by AI-Led Capex")
finish(fig, "spcx_q2_chart8_cash_flow.png")


# Figure 9: estimate revisions
metrics = ["FY26 Revenue", "FY26 Adj. EBITDA", "FY26 Capex", "FY27 Revenue", "FY27 Adj. EBITDA", "FY27 Capex"]
old = [22.0, 5.5, 35.0, 28.0, 9.0, 30.0]
new = [35.0, 15.5, 65.0, 58.0, 26.0, 75.0]
y = np.arange(len(metrics))
fig, ax = plt.subplots(figsize=(9.4, 5.4))
ax.barh(y + 0.18, old, 0.34, label="Prior estimate", color=GRAY)
ax.barh(y - 0.18, new, 0.34, label="Post-Q2 estimate", color=BLUE)
for i, value in enumerate(new):
    ax.text(value + 1.0, i - 0.18, f"${value:.1f}B", va="center", fontsize=8.5, fontweight="bold")
ax.set_yticks(y, metrics)
ax.invert_yaxis()
ax.set_xlabel("USD billions")
ax.set_title("Estimate Revisions: Growth and Spending Both Reset Higher")
ax.legend(frameon=False, loc="lower right")
ax.set_xlim(0, 84)
finish(fig, "spcx_q2_chart9_estimate_revisions.png")


# Figure 10: public-market price performance, dynamically sourced via yfinance
fig, ax = plt.subplots(figsize=(9.4, 4.8))
try:
    hist = yf.Ticker("SPCX").history(start="2026-06-12", end="2026-08-06", auto_adjust=False)
    if hist.empty:
        raise ValueError("SPCX price history is empty")
    ax.plot(hist.index, hist["Close"], color=BLUE, linewidth=2.2)
    ax.fill_between(hist.index, hist["Close"], 100, color=CYAN, alpha=0.10)
    ax.axhline(135, color=GOLD, linestyle="--", linewidth=1.4, label="IPO price: $135")
    ax.axvline(hist.index[-2], color=PURPLE, linestyle=":", linewidth=1.4, label="Q2 results")
    ax.scatter(hist.index[-1], hist["Close"].iloc[-1], color=RED, s=45, zorder=3)
    ax.text(hist.index[-1], hist["Close"].iloc[-1] + 7, f"${hist['Close'].iloc[-1]:.2f}", ha="right", fontsize=9, fontweight="bold")
    ax.set_ylim(95, max(240, float(hist["High"].max()) * 1.05))
except Exception as exc:
    print(f"Warning: could not fetch SPCX market history: {exc}")
    fallback_dates = np.arange(4)
    ax.plot(fallback_dates, [160.95, 201.80, 108.37, 115.09], color=BLUE, linewidth=2.2)
    ax.set_xticks(fallback_dates, ["IPO", "Peak", "Pre-Q2", "Post-Q2"])
    ax.axhline(135, color=GOLD, linestyle="--", linewidth=1.4, label="IPO price: $135")
ax.set_ylabel("Share price (USD)")
ax.set_title("SPCX Since IPO: Fundamental Beat Meets Capex and Unlock Overhang")
ax.legend(frameon=False)
finish(fig, "spcx_q2_chart10_price_since_ipo.png")


# Figure 11: valuation scenarios
fig, ax = plt.subplots(figsize=(8.6, 4.7))
scenario_names = ["Bear", "Base PT", "Bull"]
scenario_values = [80, 150, 225]
bars = ax.bar(scenario_names, scenario_values, color=[RED, BLUE, GREEN], width=0.5, alpha=0.88)
label_bars(ax, bars, "${:.0f}", dy=5)
ax.axhline(115.75, color=GOLD, linestyle="--", linewidth=1.4, label="Reference price: $115.75")
ax.set_ylabel("Implied value per share (USD)")
ax.set_ylim(0, 260)
ax.set_title("12-Month Valuation Scenarios")
ax.legend(frameon=False)
finish(fig, "spcx_q2_chart11_valuation_scenarios.png")

print(f"Generated 11 SPCX Q2 charts in {OUT}")
