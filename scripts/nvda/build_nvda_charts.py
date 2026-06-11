#!/usr/bin/env python3
"""Generate charts for NVIDIA Q1 FY2027 Earnings Update Report."""

import os
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

OUT = os.path.join(os.path.dirname(__file__), "..", "..", "output", "NVDA")
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "font.family":        "serif",
    "font.serif":         ["Times New Roman"],
    "axes.unicode_minus": False,
    "figure.dpi":         150,
    "savefig.dpi":        150,
    "savefig.bbox":       "tight",
    "axes.grid":          True,
    "grid.alpha":         0.3,
    "axes.spines.top":    False,
    "axes.spines.right":  False,
})

NVIDIA_GREEN = "#76B900"
DARK_GREEN   = "#4A7A00"
LIGHT_GREEN  = "#A8D600"
BLUE         = "#1E88E5"
DARK_BLUE    = "#0D47A1"
ORANGE       = "#FF6F00"
RED          = "#D32F2F"
GRAY         = "#757575"
LIGHT_GRAY   = "#BDBDBD"

# ── Historical Quarterly Data (Q2 FY25 through Q1 FY27) ────────────────────
quarters    = ["Q2\nFY25", "Q3\nFY25", "Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26", "Q4\nFY26", "Q1\nFY27"]
revenue     = [30.0, 35.1, 39.3, 44.1, 46.7, 57.0, 68.1, 81.6]
dc_revenue  = [26.3, 30.8, 35.6, 39.1, 42.0, 51.2, 62.3, 75.2]
non_dc_rev  = [r - d for r, d in zip(revenue, dc_revenue)]

gaap_gm     = [75.1, 74.6, 73.0, 60.5, 72.4, 73.4, 75.0, 74.9]
nongaap_gm  = [75.7, 75.0, 73.5, 61.0, 72.6, 73.6, 75.2, 75.0]

gaap_eps    = [0.67, 0.78, 0.89, 0.76, 1.08, 1.30, 1.76, 2.39]
nongaap_eps = [0.68, 0.81, 0.89, 0.81, 1.10, 1.30, 1.62, 1.87]

op_income   = [18.6, 21.9, 24.0, 21.6, 27.1, 36.0, 44.3, 53.5]
op_margin   = [oi / r * 100 for oi, r in zip(op_income, revenue)]

fcf         = [13.5, 16.8, 18.5, 26.1, 24.0, 30.0, 34.9, 49.0]

dc_compute  = [None, None, None, None, None, None, None, 60.4]
dc_network  = [None, None, None, None, None, None, None, 14.8]

# ── Chart 1: Quarterly Revenue Progression ──────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
colors = [NVIDIA_GREEN] * 7 + [DARK_GREEN]
bars = ax.bar(quarters, revenue, color=colors, edgecolor="white", width=0.65)
for b, v in zip(bars, revenue):
    ax.text(b.get_x() + b.get_width()/2, v + 1.0, f"${v:.1f}B", ha="center", va="bottom", fontsize=9, fontweight="bold")
ax.axhline(y=78.8, color=RED, linestyle="--", linewidth=1.2, alpha=0.7, label="Q1 FY27 Consensus: $78.8B")
ax.set_ylabel("Revenue ($ Billion)", fontsize=11)
ax.set_title("NVIDIA Quarterly Revenue Progression", fontsize=13, fontweight="bold", pad=12)
ax.legend(loc="upper left", fontsize=9)
ax.set_ylim(0, 95)
fig.savefig(os.path.join(OUT, "nvda_chart1_quarterly_revenue.png"))
plt.close()

# ── Chart 2: Quarterly EPS Progression ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
x = np.arange(len(quarters))
w = 0.35
b1 = ax.bar(x - w/2, gaap_eps, w, label="GAAP EPS", color=BLUE, edgecolor="white")
b2 = ax.bar(x + w/2, nongaap_eps, w, label="Non-GAAP EPS", color=NVIDIA_GREEN, edgecolor="white")
for b, v in zip(b1, gaap_eps):
    ax.text(b.get_x() + b.get_width()/2, v + 0.03, f"${v:.2f}", ha="center", va="bottom", fontsize=7.5)
for b, v in zip(b2, nongaap_eps):
    ax.text(b.get_x() + b.get_width()/2, v + 0.03, f"${v:.2f}", ha="center", va="bottom", fontsize=7.5)
ax.axhline(y=1.77, color=RED, linestyle="--", linewidth=1.2, alpha=0.7, label="Q1 FY27 Non-GAAP Consensus: $1.77")
ax.set_xticks(x)
ax.set_xticklabels(quarters)
ax.set_ylabel("Diluted EPS ($)", fontsize=11)
ax.set_title("NVIDIA Quarterly EPS Progression", fontsize=13, fontweight="bold", pad=12)
ax.legend(loc="upper left", fontsize=9)
ax.set_ylim(0, 2.8)
fig.savefig(os.path.join(OUT, "nvda_chart2_quarterly_eps.png"))
plt.close()

# ── Chart 3: Quarterly Margin Trends ────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
ax.plot(quarters, nongaap_gm, "o-", color=NVIDIA_GREEN, linewidth=2.5, markersize=7, label="Non-GAAP Gross Margin")
ax.plot(quarters, gaap_gm, "s--", color=BLUE, linewidth=2, markersize=6, label="GAAP Gross Margin")
ax.plot(quarters, op_margin, "^-", color=ORANGE, linewidth=2, markersize=6, label="GAAP Operating Margin")
for i, (g, ng) in enumerate(zip(gaap_gm, nongaap_gm)):
    ax.annotate(f"{ng:.1f}%", (quarters[i], ng), textcoords="offset points", xytext=(0, 10), fontsize=7.5, ha="center", color=DARK_GREEN)
ax.annotate("H20 charge\nimpact", xy=(quarters[3], gaap_gm[3]), xytext=(0, -25),
            textcoords="offset points", fontsize=8, ha="center", color=RED,
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2))
ax.set_ylabel("Margin (%)", fontsize=11)
ax.set_title("NVIDIA Quarterly Margin Trends", fontsize=13, fontweight="bold", pad=12)
ax.legend(loc="lower right", fontsize=9)
ax.set_ylim(45, 82)
fig.savefig(os.path.join(OUT, "nvda_chart3_margin_trends.png"))
plt.close()

# ── Chart 4: Revenue by Segment (Data Center vs. Other) ─────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
ax.bar(quarters, dc_revenue, color=NVIDIA_GREEN, label="Data Center", edgecolor="white", width=0.65)
ax.bar(quarters, non_dc_rev, bottom=dc_revenue, color=LIGHT_GRAY, label="Other (Gaming / Edge / ProViz / Auto)", edgecolor="white", width=0.65)
for i, (d, t) in enumerate(zip(dc_revenue, revenue)):
    pct = d / t * 100
    ax.text(i, d / 2, f"${d:.1f}B\n({pct:.0f}%)", ha="center", va="center", fontsize=8, fontweight="bold", color="white")
ax.set_ylabel("Revenue ($ Billion)", fontsize=11)
ax.set_title("NVIDIA Revenue: Data Center vs. Other Segments", fontsize=13, fontweight="bold", pad=12)
ax.legend(loc="upper left", fontsize=9)
ax.set_ylim(0, 95)
fig.savefig(os.path.join(OUT, "nvda_chart4_segment_revenue.png"))
plt.close()

# ── Chart 5: Data Center Revenue YoY Growth ─────────────────────────────────
dc_yoy = [None, None, None, None,
          (42.0 / 26.3 - 1) * 100,
          (51.2 / 30.8 - 1) * 100,
          (62.3 / 35.6 - 1) * 100,
          (75.2 / 39.1 - 1) * 100]
fig, ax1 = plt.subplots(figsize=(10, 5.5))
ax1.bar(quarters, dc_revenue, color=NVIDIA_GREEN, alpha=0.7, label="DC Revenue ($B)", edgecolor="white", width=0.65)
ax1.set_ylabel("Data Center Revenue ($ Billion)", fontsize=11, color=DARK_GREEN)
ax1.set_ylim(0, 90)
ax2 = ax1.twinx()
valid_q = [q for q, y in zip(quarters, dc_yoy) if y is not None]
valid_y = [y for y in dc_yoy if y is not None]
ax2.plot(valid_q, valid_y, "D-", color=ORANGE, linewidth=2.5, markersize=8, label="YoY Growth (%)")
for q, y in zip(valid_q, valid_y):
    ax2.annotate(f"{y:.0f}%", (q, y), textcoords="offset points", xytext=(0, 10), fontsize=9, fontweight="bold", ha="center", color=ORANGE)
ax2.set_ylabel("YoY Growth (%)", fontsize=11, color=ORANGE)
ax2.set_ylim(0, 120)
ax2.spines["right"].set_visible(True)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=9)
ax1.set_title("NVIDIA Data Center Revenue & YoY Growth", fontsize=13, fontweight="bold", pad=12)
fig.savefig(os.path.join(OUT, "nvda_chart5_dc_revenue_growth.png"))
plt.close()

# ── Chart 6: Q1 FY27 Beat/Miss Waterfall ────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
categories = ["Consensus\nRevenue", "DC Beat", "Edge Beat", "Reported\nRevenue"]
values     = [78.8, 2.0, 0.8, 81.6]
bottoms    = [0, 78.8, 80.8, 0]
colors_wf  = [GRAY, NVIDIA_GREEN, LIGHT_GREEN, DARK_GREEN]
bars = ax.bar(categories, values, bottom=bottoms, color=colors_wf, edgecolor="white", width=0.55)
ax.text(0, 78.8 / 2, "$78.8B", ha="center", va="center", fontsize=11, fontweight="bold", color="white")
ax.text(1, 78.8 + 1.0, "+$2.0B", ha="center", va="bottom", fontsize=10, fontweight="bold", color=DARK_GREEN)
ax.text(2, 80.8 + 0.4, "+$0.8B", ha="center", va="bottom", fontsize=10, fontweight="bold", color=DARK_GREEN)
ax.text(3, 81.6 / 2, "$81.6B", ha="center", va="center", fontsize=11, fontweight="bold", color="white")
ax.set_ylabel("Revenue ($ Billion)", fontsize=11)
ax.set_title("Q1 FY2027 Revenue Beat: Consensus to Reported", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 90)
fig.savefig(os.path.join(OUT, "nvda_chart6_beat_waterfall.png"))
plt.close()

# ── Chart 7: Free Cash Flow Trend ───────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
bars = ax.bar(quarters, fcf, color=BLUE, edgecolor="white", width=0.65, alpha=0.85)
for b, v in zip(bars, fcf):
    ax.text(b.get_x() + b.get_width()/2, v + 0.5, f"${v:.1f}B", ha="center", va="bottom", fontsize=9, fontweight="bold")
ax.set_ylabel("Free Cash Flow ($ Billion)", fontsize=11)
ax.set_title("NVIDIA Quarterly Free Cash Flow", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 58)
fig.savefig(os.path.join(OUT, "nvda_chart7_free_cash_flow.png"))
plt.close()

# ── Chart 8: Data Center Sub-segments (Q1 FY27) ────────────────────────────
fig, ax = plt.subplots(figsize=(7, 5.5))
sub_labels = ["Compute\n$60.4B", "Networking\n$14.8B"]
sub_values = [60.4, 14.8]
sub_colors = [NVIDIA_GREEN, DARK_BLUE]
wedges, texts, autotexts = ax.pie(sub_values, labels=sub_labels, colors=sub_colors,
                                   autopct="%1.0f%%", startangle=90, textprops={"fontsize": 11},
                                   wedgeprops={"edgecolor": "white", "linewidth": 2})
for at in autotexts:
    at.set_fontsize(12)
    at.set_fontweight("bold")
    at.set_color("white")
ax.set_title("Q1 FY2027 Data Center Revenue Mix\n(Compute vs. Networking)", fontsize=13, fontweight="bold", pad=12)
fig.savefig(os.path.join(OUT, "nvda_chart8_dc_mix.png"))
plt.close()

# ── Chart 9: Estimate Revision — FY2027E Updates ───────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
metrics = ["Revenue\n($B)", "Non-GAAP\nEPS ($)", "Gross\nMargin (%)"]
old_est = [340.0, 8.20, 74.5]
new_est = [365.0, 8.80, 75.0]
x = np.arange(len(metrics))
w = 0.3
b1 = ax.bar(x - w/2, old_est, w, label="Pre-Q1 Estimate", color=LIGHT_GRAY, edgecolor="white")
b2 = ax.bar(x + w/2, new_est, w, label="Post-Q1 Estimate", color=NVIDIA_GREEN, edgecolor="white")
for i in range(len(metrics)):
    chg = (new_est[i] / old_est[i] - 1) * 100
    ax.text(x[i] + w/2, new_est[i] + 1, f"+{chg:.1f}%", ha="center", va="bottom", fontsize=9, fontweight="bold", color=DARK_GREEN)
ax.set_xticks(x)
ax.set_xticklabels(metrics)
ax.set_title("FY2027E Estimate Revisions (Post Q1 Results)", fontsize=13, fontweight="bold", pad=12)
ax.legend(loc="upper right", fontsize=9)
fig.savefig(os.path.join(OUT, "nvda_chart9_estimate_revisions.png"))
plt.close()

# ── Chart 10: Quarterly Revenue Sequential Growth ───────────────────────────
seq_growth = [None]
for i in range(1, len(revenue)):
    seq_growth.append((revenue[i] / revenue[i - 1] - 1) * 100)
fig, ax = plt.subplots(figsize=(10, 5.5))
valid_q2 = [q for q, s in zip(quarters, seq_growth) if s is not None]
valid_sg = [s for s in seq_growth if s is not None]
bars = ax.bar(valid_q2, valid_sg, color=[NVIDIA_GREEN if s > 0 else RED for s in valid_sg], edgecolor="white", width=0.55)
for b, v in zip(bars, valid_sg):
    ax.text(b.get_x() + b.get_width()/2, v + 0.3, f"+{v:.1f}%", ha="center", va="bottom", fontsize=9, fontweight="bold")
ax.set_ylabel("Sequential Growth (%)", fontsize=11)
ax.set_title("NVIDIA Quarterly Revenue — Sequential Growth Rate", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 28)
fig.savefig(os.path.join(OUT, "nvda_chart10_sequential_growth.png"))
plt.close()

print(f"All 10 charts saved to {OUT}")
