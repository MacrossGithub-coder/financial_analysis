#!/usr/bin/env python3
"""Generate the 11 figures used in the SPCX Q3 2026 earnings update."""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "SPCX"
DATA = ROOT / "data" / "spcx" / "q3_2026_facts.json"
OUT.mkdir(parents=True, exist_ok=True)

with DATA.open(encoding="utf-8") as handle:
    FACTS = json.load(handle)

plt.rcParams.update(
    {
        "font.family": "serif",
        "font.serif": ["Times New Roman", "Times", "DejaVu Serif"],
        "axes.unicode_minus": False,
        "figure.dpi": 150,
        "savefig.dpi": 150,
        "savefig.bbox": "tight",
        "axes.grid": True,
        "grid.alpha": 0.16,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.titleweight": "bold",
    }
)

BLACK = "#111111"
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
    fig.tight_layout(rect=(0, 0.02, 1, 0.98))
    fig.savefig(OUT / name)
    plt.close(fig)


def label_bars(ax, bars, fmt="{:.1f}", dy=0.03, color=BLACK, fontsize=8.5):
    for bar in bars:
        value = bar.get_height()
        offset = dy if value >= 0 else -dy
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + offset,
            fmt.format(value),
            ha="center",
            va="bottom" if value >= 0 else "top",
            fontsize=fontsize,
            fontweight="bold",
            color=color,
        )


periods = FACTS["quarterly"]["periods"]
x = np.arange(len(periods))


# Figure 1: beat/miss snapshot
fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.4))
rev_est = FACTS["consensus"]["revenue"]
rev_act = FACTS["q3_income_statement"]["revenue"]
eps_est = FACTS["consensus"]["gaap_diluted_eps"]
eps_act = FACTS["q3_income_statement"]["diluted_eps"]
bars = axes[0].bar(["Consensus", "Reported"], [rev_est, rev_act], color=[GRAY, BLUE], width=0.55)
label_bars(axes[0], bars, "${:.1f}B", dy=0.25)
axes[0].set_title(f"Revenue: {(rev_act / rev_est - 1) * 100:.1f}% Beat")
axes[0].set_ylabel("USD billions")
axes[0].set_ylim(0, 17)
bars = axes[1].bar(["Consensus", "Reported"], [eps_est, eps_act], color=[GRAY, GREEN], width=0.55)
label_bars(axes[1], bars, "${:.2f}", dy=0.012)
axes[1].set_title(f"GAAP EPS: +${eps_act - eps_est:.2f} vs. Consensus")
axes[1].set_ylabel("USD per diluted share")
axes[1].set_ylim(0, 0.42)
fig.suptitle("SpaceX Q3 2026 Beat/Miss Snapshot", fontsize=14, fontweight="bold", color=BLACK)
finish(fig, "spcx_q3_chart1_beat_miss.png")


# Figure 2: quarterly revenue and EBITDA
revenue = FACTS["quarterly"]["revenue"]
ebitda = FACTS["quarterly"]["adjusted_ebitda"]
fig, ax = plt.subplots(figsize=(9.2, 4.8))
width = 0.34
b1 = ax.bar(x - width / 2, revenue, width, label="Revenue", color=BLUE)
b2 = ax.bar(x + width / 2, ebitda, width, label="Adjusted EBITDA", color=GREEN)
label_bars(ax, b1, "${:.1f}B", dy=0.25)
label_bars(ax, b2, "${:.1f}B", dy=0.25)
ax.set_xticks(x, periods)
ax.set_ylabel("USD billions")
ax.set_ylim(0, 17)
ax.set_title("Revenue and Adjusted EBITDA Inflect in Q3")
ax.legend(frameon=False, ncol=2)
finish(fig, "spcx_q3_chart2_revenue_ebitda.png")


# Figure 3: segment revenue
space = FACTS["segments"]["space"]["revenue"]
connectivity = FACTS["segments"]["connectivity"]["revenue"]
si = FACTS["segments"]["si"]["revenue"]
fig, ax = plt.subplots(figsize=(9.2, 4.8))
w = 0.24
b1 = ax.bar(x - w, space, w, label="Space", color=ORANGE)
b2 = ax.bar(x, connectivity, w, label="Connectivity", color=CYAN)
b3 = ax.bar(x + w, si, w, label="Super Intelligence", color=PURPLE)
for bars in (b1, b2, b3):
    label_bars(ax, bars, "${:.1f}", dy=0.15, fontsize=7.8)
ax.set_xticks(x, periods)
ax.set_ylabel("Revenue (USD billions)")
ax.set_ylim(0, 8.8)
ax.set_title("SI Becomes the Largest Revenue Segment")
ax.legend(frameon=False, ncol=3, loc="upper left")
finish(fig, "spcx_q3_chart3_segment_revenue.png")


# Figure 4: segment operating income
space_op = FACTS["segments"]["space"]["operating_income"]
conn_op = FACTS["segments"]["connectivity"]["operating_income"]
si_op = FACTS["segments"]["si"]["operating_income"]
fig, ax = plt.subplots(figsize=(9.2, 4.8))
b1 = ax.bar(x - w, space_op, w, label="Space", color=ORANGE)
b2 = ax.bar(x, conn_op, w, label="Connectivity", color=CYAN)
b3 = ax.bar(x + w, si_op, w, label="Super Intelligence", color=PURPLE)
for bars in (b1, b2, b3):
    label_bars(ax, bars, "${:.1f}", dy=0.08, fontsize=7.8)
ax.axhline(0, color=BLACK, linewidth=0.9)
ax.set_xticks(x, periods)
ax.set_ylabel("Operating income/(loss), USD billions")
ax.set_ylim(-3.1, 3.1)
ax.set_title("SI Crosses into GAAP Operating Profit")
ax.legend(frameon=False, ncol=3, loc="lower left")
finish(fig, "spcx_q3_chart4_segment_operating_income.png")


# Figure 5: consolidated margin trend
fig, ax = plt.subplots(figsize=(9.2, 4.8))
ax.plot(periods, FACTS["quarterly"]["gross_margin_pct"], marker="o", linewidth=2.4, color=BLUE, label="Gross margin")
ax.plot(periods, FACTS["quarterly"]["operating_margin_pct"], marker="o", linewidth=2.4, color=ORANGE, label="Operating margin")
ax.plot(periods, FACTS["quarterly"]["net_margin_pct"], marker="o", linewidth=2.4, color=GREEN, label="Net margin")
ax.axhline(0, color=BLACK, linewidth=0.8)
for series, color in (
    (FACTS["quarterly"]["gross_margin_pct"], BLUE),
    (FACTS["quarterly"]["operating_margin_pct"], ORANGE),
    (FACTS["quarterly"]["net_margin_pct"], GREEN),
):
    ax.text(3.04, series[-1], f"{series[-1]:.1f}%", va="center", fontsize=8.5, fontweight="bold", color=color)
ax.set_ylabel("Margin")
ax.set_ylim(-105, 72)
ax.set_title("Q3 Operating Leverage Reverses the Margin Profile")
ax.legend(frameon=False, ncol=3, loc="lower center")
finish(fig, "spcx_q3_chart5_margin_trends.png")


# Figure 6: Starlink subscribers and ARPU
subs = FACTS["segments"]["connectivity"]["subscribers_m"]
arpu = FACTS["segments"]["connectivity"]["monthly_arpu"]
fig, ax1 = plt.subplots(figsize=(9.2, 4.8))
bars = ax1.bar(periods, subs, color=CYAN, width=0.5, label="Subscribers")
label_bars(ax1, bars, "{:.1f}M", dy=0.25)
ax1.set_ylabel("Subscribers (millions)", color=BLUE)
ax1.set_ylim(0, 16)
ax2 = ax1.twinx()
ax2.grid(False)
ax2.plot(periods, arpu, color=GOLD, marker="o", linewidth=2.5, markersize=7, label="ARPU")
for i, value in enumerate(arpu):
    ax2.text(i, value + 2.5, f"${value}", ha="center", fontsize=8.5, fontweight="bold", color=GOLD)
ax2.set_ylabel("Monthly consumer ARPU (USD)", color=GOLD)
ax2.set_ylim(50, 95)
ax1.set_title("Starlink Volume Compounds as ARPU Turns Sequentially Higher")
finish(fig, "spcx_q3_chart6_starlink_metrics.png")


# Figure 7: SI monetization and capex intensity
compute = FACTS["segments"]["si"]["compute_gw"]
si_capex = [4.760, 7.723, 15.828, 16.444]
fig, ax1 = plt.subplots(figsize=(9.2, 4.8))
b1 = ax1.bar(x - 0.18, si, 0.36, label="SI revenue", color=PURPLE)
b2 = ax1.bar(x + 0.18, si_capex, 0.36, label="SI capex", color=RED, alpha=0.82)
label_bars(ax1, b1, "${:.1f}B", dy=0.3, fontsize=7.8)
label_bars(ax1, b2, "${:.1f}B", dy=0.3, fontsize=7.8)
ax1.set_xticks(x, periods)
ax1.set_ylabel("USD billions")
ax1.set_ylim(0, 19)
ax2 = ax1.twinx()
ax2.grid(False)
ax2.plot(periods, compute, color=GOLD, marker="D", linewidth=2.2, label="Nameplate compute")
ax2.set_ylabel("Nameplate compute (GW)", color=GOLD)
ax2.set_ylim(0, 2.2)
ax1.set_title("SI Monetization Accelerates while Capex Remains Extreme")
lines, labels = ax1.get_legend_handles_labels()
line2, label2 = ax2.get_legend_handles_labels()
ax1.legend(lines + line2, labels + label2, frameon=False, ncol=3, loc="upper left")
finish(fig, "spcx_q3_chart7_si_revenue_capex_compute.png")


# Figure 8: nine-month cash-flow profile
cf = FACTS["cash_balance_sheet"]
cf_labels = ["Operating cash flow", "Capital expenditures", "Free cash flow"]
cf_values = [cf["nine_month_operating_cash_flow"], -cf["nine_month_capex"], cf["nine_month_free_cash_flow"]]
fig, ax = plt.subplots(figsize=(8.8, 4.8))
bars = ax.bar(cf_labels, cf_values, color=[GREEN, RED, RED], width=0.5)
label_bars(ax, bars, "${:.1f}B", dy=1.2)
ax.axhline(0, color=BLACK, linewidth=0.9)
ax.set_ylabel("USD billions")
ax.set_ylim(-55, 22)
ax.set_title("9M 2026 Cash Flow: Growth Still Requires External Funding")
finish(fig, "spcx_q3_chart8_cash_flow.png")


# Figure 9: management outlook bridge
fig, axes = plt.subplots(1, 2, figsize=(9.8, 4.7))
bars = axes[0].bar(["Q3 revenue", "Q1'27 guide midpoint"], [14.575, 35.0], color=[BLUE, GREEN], width=0.55)
label_bars(axes[0], bars, "${:.1f}B", dy=0.8)
axes[0].set_ylim(0, 41)
axes[0].set_ylabel("Quarterly revenue")
axes[0].set_title("Q1 2027 Revenue Guidance")
bars = axes[1].bar(["Q3 actual", "Q4 guide midpoint"], [19.593, 37.5], color=[GRAY, RED], width=0.55)
label_bars(axes[1], bars, "${:.1f}B", dy=0.8)
axes[1].set_ylim(0, 44)
axes[1].set_ylabel("Quarterly capex")
axes[1].set_title("Q4 2026 Capex Guidance")
fig.suptitle("Management Signals a Faster Revenue and Investment Step-Up", fontsize=13.5, fontweight="bold")
finish(fig, "spcx_q3_chart9_guidance.png")


# Figure 10: estimate revisions
metrics = ["FY26 revenue", "FY26 adj. EBITDA", "FY26 capex", "FY27 revenue", "FY27 adj. EBITDA", "FY27 capex"]
old = [35.0, 15.5, 65.0, 58.0, 26.0, 75.0]
new = [50.5, 29.0, 85.5, 142.0, 72.0, 150.0]
y = np.arange(len(metrics))
fig, ax = plt.subplots(figsize=(9.4, 5.4))
ax.barh(y + 0.18, old, 0.34, label="Post-Q2 estimate", color=GRAY)
ax.barh(y - 0.18, new, 0.34, label="Post-Q3 estimate", color=BLUE)
for i, value in enumerate(new):
    ax.text(value + 2.0, i - 0.18, f"${value:.1f}B", va="center", fontsize=8.0, fontweight="bold")
ax.set_yticks(y, metrics)
ax.invert_yaxis()
ax.set_xlabel("USD billions")
ax.set_title("Estimate Revisions: Revenue Visibility and Capital Intensity Both Rise")
ax.legend(frameon=False, loc="lower right")
ax.set_xlim(0, 170)
finish(fig, "spcx_q3_chart10_estimate_revisions.png")


# Figure 11: valuation scenarios
valuation = FACTS["valuation"]
fig, ax = plt.subplots(figsize=(8.8, 4.8))
names = ["Bear", "Base PT", "Bull"]
values = [valuation["bear"], valuation["base"], valuation["bull"]]
bars = ax.bar(names, values, color=[RED, BLUE, GREEN], width=0.52, alpha=0.9)
label_bars(ax, bars, "${:.0f}", dy=7)
ax.axhline(valuation["reference_price"], color=GOLD, linestyle="--", linewidth=1.6, label=f"10/20 close ${valuation['reference_price']:.2f}")
ax.set_ylabel("Value per share (USD)")
ax.set_ylim(0, 370)
ax.set_title("Twelve-Month Valuation Scenarios")
ax.legend(frameon=False)
finish(fig, "spcx_q3_chart11_valuation_scenarios.png")

print("Generated 11 SPCX Q3 2026 charts")
