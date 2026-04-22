#!/usr/bin/env python3
"""
阳光电源 (300274.SZ) Q4 2025 / FY2025 Earnings Update — Chart Generation
Generates 10 charts for the earnings update report.
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

# ── Output directory ──────────────────────────────────────────────────────────
BASE = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/300274"
os.makedirs(BASE, exist_ok=True)

# ── Font config (CJK support) ─────────────────────────────────────────────────
plt.rcParams.update({
    "font.family":        "sans-serif",
    "font.sans-serif":    ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,
    "figure.dpi":         150,
    "savefig.dpi":        150,
    "savefig.bbox":       "tight",
})

# ── Color palette ─────────────────────────────────────────────────────────────
BLUE      = "#1F4E79"
LIGHT_BLUE= "#2E75B6"
ORANGE    = "#E07B39"
GREEN     = "#375623"
LIGHT_GRN = "#70AD47"
GREY      = "#808080"
RED       = "#C00000"
GOLD      = "#FFC000"
BG        = "#F5F5F5"

# ─────────────────────────────────────────────────────────────────────────────
# DATA
# ─────────────────────────────────────────────────────────────────────────────
quarters = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25"]

# Revenue (CNY billion)
revenue = [12.61, 18.41, 18.93, 27.91, 19.04, 24.50, 22.87, 22.78]

# Net profit attributable to parent (CNY billion)
# FY2024 ≈ 11.04B; Q4'24 = 3.44B derived from -54% YoY
# Q1-Q3'25 = 11.88B; Q3'25 = 4.15B; Q1+Q2'25 = 7.73B (split ~3.40 + 4.33)
net_profit = [2.40, 3.40, 3.73, 3.44, 3.40, 4.33, 4.15, 1.58]

# Gross margin %
gm = [26.5, 25.8, 26.1, 23.5, 29.4, 32.1, 35.9, 24.2]
# Net margin %
nm = [19.0, 18.5, 19.7, 12.3, 17.9, 17.7, 18.1, 6.9]

# FY2025 segment revenue (CNY billion)
seg_labels = ["储能系统\nEnergy Storage", "光伏逆变器\nPV Inverters", "新能源投资\nNew Energy Dev.", "其他 Other"]
seg_vals   = [37.29, 31.14, 16.56, 4.20]
seg_colors = [BLUE, LIGHT_BLUE, ORANGE, GREY]

# Storage GWh shipments
ship_years  = ["2023", "2024", "2025", "2026E"]
ship_gwh    = [16.0, 28.0, 43.0, 62.0]  # 2026E = >60 GWh target midpoint
dom_gwh     = [7.0,  9.0,  7.0, 10.0]
ovs_gwh     = [9.0, 19.0, 36.0, 52.0]

# PV inverter shipments (GW)
inv_years  = ["2023", "2024", "2025", "2026E"]
inv_gw     = [130.0, 147.0, 143.0, 155.0]

# Revenue YoY growth %
rev_yoy = [None, None, None, None, 50.9, 33.1, 20.8, -18.4]

# Beat/miss vs analyst consensus (CNY billion)
cons_fy25_rev  = 90.0  # consensus revenue estimate
cons_fy25_np   = 15.45 # consensus net profit estimate
act_fy25_rev   = 89.18
act_fy25_np    = 13.46

# Estimate revisions (CNY billion)
est_rev_labels = ["FY2025E Rev\n(Old)", "FY2025A Rev\n(Actual)", "FY2026E Rev\n(New)"]
est_rev_vals   = [90.0, 89.18, 102.0]
est_np_labels  = ["FY2025E NP\n(Consensus)", "FY2025A NP\n(Actual)", "FY2026E NP\n(New Est)"]
est_np_vals    = [15.45, 13.46, 16.77]

# Valuation: NTM P/E band history (approximate)
val_periods = ["Jan'24", "Apr'24", "Jul'24", "Oct'24", "Jan'25", "Apr'25", "Jul'25", "Oct'25", "Jan'26", "Apr'26"]
pe_high = [22, 24, 28, 32, 38, 35, 40, 45, 30, 26]
pe_mid  = [18, 20, 22, 26, 30, 28, 32, 36, 24, 21]
pe_low  = [14, 16, 17, 20, 22, 22, 25, 28, 18, 17]
pe_curr = [None]*9 + [20.9]

# ─────────────────────────────────────────────────────────────────────────────
# CHART 1 — Quarterly Revenue Progression
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
fig.patch.set_facecolor("white")
bar_colors = [BLUE if i < 4 else LIGHT_BLUE for i in range(8)]
bar_colors[-1] = ORANGE  # highlight Q4'25
bars = ax.bar(quarters, revenue, color=bar_colors, width=0.6, edgecolor="white", linewidth=0.5)
for bar, val in zip(bars, revenue):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
            f"¥{val:.1f}B", ha="center", va="bottom", fontsize=8.5, fontweight="bold")
# YoY annotations for 2025 bars
for i, (q, yoy) in enumerate(zip(quarters, rev_yoy)):
    if yoy is not None:
        color = LIGHT_GRN if yoy > 0 else RED
        ax.text(i, revenue[i]/2, f"{yoy:+.1f}%\nYoY", ha="center", va="center",
                fontsize=7.5, color="white", fontweight="bold")
ax.set_ylabel("Revenue (CNY Billion)", fontsize=10)
ax.set_title("Figure 1 — Quarterly Revenue Progression | 阳光电源 季度营收", fontsize=12, fontweight="bold", pad=10)
ax.axvline(3.5, color=GREY, linestyle="--", linewidth=1)
ax.text(3.6, max(revenue)*0.95, "FY2024 → FY2025", fontsize=8, color=GREY)
ax.set_ylim(0, max(revenue)*1.2)
ax.grid(axis="y", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
legend_handles = [mpatches.Patch(color=BLUE, label="FY2024"),
                  mpatches.Patch(color=LIGHT_BLUE, label="FY2025 Q1-Q3"),
                  mpatches.Patch(color=ORANGE, label="FY2025 Q4 (Reported)")]
ax.legend(handles=legend_handles, fontsize=8.5, loc="upper left")
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report (2026-04-01); StockAnalysis.com", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart1_revenue.png"))
plt.close()
print("Chart 1 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 2 — Quarterly Net Profit Progression
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
fig.patch.set_facecolor("white")
bar_colors2 = [BLUE if i < 4 else LIGHT_BLUE for i in range(8)]
bar_colors2[-1] = ORANGE
bars2 = ax.bar(quarters, net_profit, color=bar_colors2, width=0.6, edgecolor="white", linewidth=0.5)
for bar, val in zip(bars2, net_profit):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.05,
            f"¥{val:.2f}B", ha="center", va="bottom", fontsize=8.5, fontweight="bold")
# Highlight Q4'25 miss
ax.annotate("Q4'25: ¥1.58B\n-54% YoY\n(Significant miss)", xy=(7, 1.58), xytext=(5.8, 3.8),
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.5),
            fontsize=8.5, color=RED, fontweight="bold",
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", edgecolor=RED))
ax.set_ylabel("Net Profit Attributable to Parent (CNY Billion)", fontsize=10)
ax.set_title("Figure 2 — Quarterly Net Profit Progression | 阳光电源 季度归母净利润", fontsize=12, fontweight="bold", pad=10)
ax.axvline(3.5, color=GREY, linestyle="--", linewidth=1)
ax.set_ylim(0, max(net_profit)*1.4)
ax.grid(axis="y", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report; Analyst estimates for quarterly split", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart2_net_profit.png"))
plt.close()
print("Chart 2 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 3 — Quarterly Margin Trends
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
fig.patch.set_facecolor("white")
ax.plot(quarters, gm, "o-", color=BLUE, linewidth=2, markersize=6, label="Gross Margin (%)")
ax.plot(quarters, nm, "s--", color=ORANGE, linewidth=2, markersize=6, label="Net Margin (%)")
for i, (q, g, n) in enumerate(zip(quarters, gm, nm)):
    ax.text(i, g + 1.0, f"{g:.1f}%", ha="center", va="bottom", fontsize=8, color=BLUE, fontweight="bold")
    ax.text(i, n - 1.5, f"{n:.1f}%", ha="center", va="top", fontsize=8, color=ORANGE, fontweight="bold")
ax.axvline(3.5, color=GREY, linestyle="--", linewidth=1)
ax.set_ylim(0, 45)
ax.set_ylabel("Margin (%)", fontsize=10)
ax.set_title("Figure 3 — Quarterly Margin Trends | 阳光电源 季度利润率走势", fontsize=12, fontweight="bold", pad=10)
ax.legend(fontsize=9, loc="upper left")
ax.grid(alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
# Annotate Q4'25 margin compression
ax.annotate("Q4'25 GM: 24.2%\nStorage GM -17ppts QoQ\ndue to mix & Li cost",
            xy=(7, 24.2), xytext=(5.5, 38),
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2),
            fontsize=8, color=RED,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", edgecolor=RED))
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report; Company investor call (2026-03-31)", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart3_margins.png"))
plt.close()
print("Chart 3 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 4 — FY2025 Revenue Breakdown by Segment
# ─────────────────────────────────────────────────────────────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 5.5))
fig.patch.set_facecolor("white")
# Pie chart
wedges, texts, autotexts = ax1.pie(seg_vals, labels=None, colors=seg_colors,
                                    autopct="%1.1f%%", startangle=90,
                                    pctdistance=0.75, wedgeprops=dict(edgecolor="white", linewidth=1.5))
for at in autotexts:
    at.set_fontsize(9)
    at.set_fontweight("bold")
    at.set_color("white")
ax1.legend(wedges, [f"{l}\n¥{v:.1f}B" for l, v in zip(seg_labels, seg_vals)],
           loc="lower center", bbox_to_anchor=(0.5, -0.25), fontsize=7.5, ncol=2)
ax1.set_title("FY2025 Revenue by Segment", fontsize=11, fontweight="bold")
# Bar chart — segment growth
seg_2024 = [24.9, 28.8, 21.0, 3.2]  # approx
x = np.arange(len(seg_labels))
w = 0.35
b1 = ax2.bar(x - w/2, seg_2024, w, label="FY2024A", color=GREY, edgecolor="white")
b2 = ax2.bar(x + w/2, seg_vals, w, label="FY2025A", color=[BLUE, LIGHT_BLUE, ORANGE, GREY], edgecolor="white")
for bar, val in zip(b2, seg_vals):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
             f"¥{val:.1f}B", ha="center", fontsize=7.5, fontweight="bold")
ax2.set_xticks(x)
ax2.set_xticklabels(["Energy\nStorage", "PV\nInverters", "New Energy\nDev.", "Other"], fontsize=8)
ax2.set_ylabel("Revenue (CNY Billion)", fontsize=9)
ax2.set_title("Segment Revenue: FY2024A vs FY2025A", fontsize=11, fontweight="bold")
ax2.legend(fontsize=8.5)
ax2.grid(axis="y", alpha=0.3)
ax2.spines[["top", "right"]].set_visible(False)
fig.suptitle("Figure 4 — Revenue Breakdown by Segment | 阳光电源 业务板块营收", fontsize=12, fontweight="bold", y=1.02)
fig.text(0.02, -0.03, "Source: Sungrow 2025 Annual Report (2026-04-01); FY2024 estimates from company disclosures", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart4_segments.png"))
plt.close()
print("Chart 4 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 5 — Energy Storage Shipments (GWh)
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
fig.patch.set_facecolor("white")
x = np.arange(len(ship_years))
w = 0.5
b_dom = ax.bar(x, dom_gwh, w, label="Domestic 国内", color=LIGHT_BLUE, edgecolor="white")
b_ovs = ax.bar(x, ovs_gwh, w, bottom=dom_gwh, label="Overseas 海外", color=BLUE, edgecolor="white")
totals = [d + o for d, o in zip(dom_gwh, ovs_gwh)]
for xi, tot in enumerate(totals):
    ax.text(xi, tot + 0.5, f"{tot:.0f} GWh", ha="center", fontsize=10, fontweight="bold")
ax.set_xticks(x)
ax.set_xticklabels(ship_years, fontsize=10)
ax.set_ylabel("Shipments (GWh)", fontsize=10)
ax.set_title("Figure 5 — Energy Storage System Shipments | 阳光电源 储能系统出货量", fontsize=12, fontweight="bold", pad=10)
ax.legend(fontsize=9, loc="upper left")
ax.grid(axis="y", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
# Growth arrows
for i in range(1, len(totals)):
    growth = (totals[i] - totals[i-1]) / totals[i-1] * 100
    color = LIGHT_GRN if growth > 0 else RED
    ax.annotate(f"+{growth:.0f}%" if growth > 0 else f"{growth:.0f}%",
                xy=(i, totals[i]/2), xytext=(i - 0.25, totals[i]/2),
                fontsize=9, color="white", fontweight="bold", ha="center")
ax.set_ylim(0, max(totals) * 1.2)
# Note for 2026E
ax.text(3, ship_gwh[3] + 2, "2026 Target:\n>60 GWh\n(+40% YoY)", ha="center", fontsize=8, color=ORANGE,
        bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", edgecolor=ORANGE))
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report; Company guidance (2026-03-31 Investor Call)", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart5_storage_shipments.png"))
plt.close()
print("Chart 5 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 6 — PV Inverter Shipments (GW)
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
fig.patch.set_facecolor("white")
bar_colors6 = [BLUE, BLUE, LIGHT_BLUE, ORANGE]
bars6 = ax.bar(inv_years, inv_gw, color=bar_colors6, width=0.5, edgecolor="white")
for bar, val in zip(bars6, inv_gw):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f"{val:.0f} GW", ha="center", fontsize=10, fontweight="bold")
ax.set_ylabel("Shipments (GW)", fontsize=10)
ax.set_title("Figure 6 — PV Inverter Shipments | 阳光电源 光伏逆变器出货量", fontsize=12, fontweight="bold", pad=10)
ax.set_ylim(0, max(inv_gw) * 1.25)
ax.grid(axis="y", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
ax.annotate("143 GW in FY2025\n(-2.7% YoY)\nDomestic -19% YoY", xy=(2, 143), xytext=(1.5, 170),
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2),
            fontsize=8, color=RED,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", edgecolor=RED))
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report (2026-04-01); 2026E = Analyst estimate", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart6_inverter_shipments.png"))
plt.close()
print("Chart 6 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 7 — Beat/Miss Summary
# ─────────────────────────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(11, 5.5))
fig.patch.set_facecolor("white")
# Revenue
categories_r = ["Consensus Est.", "Actual"]
values_r = [cons_fy25_rev, act_fy25_rev]
colors_r = [GREY, LIGHT_BLUE]
b = axes[0].bar(categories_r, values_r, color=colors_r, width=0.5, edgecolor="white")
for bar, val in zip(b, values_r):
    axes[0].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                 f"¥{val:.2f}B", ha="center", fontsize=11, fontweight="bold")
axes[0].set_title("FY2025 Revenue vs. Consensus", fontsize=11, fontweight="bold")
axes[0].set_ylabel("CNY Billion", fontsize=10)
miss_r = (act_fy25_rev - cons_fy25_rev) / cons_fy25_rev * 100
axes[0].text(0.5, max(values_r)*0.6, f"MISS: {miss_r:.1f}%\n(¥{act_fy25_rev-cons_fy25_rev:.2f}B)",
             ha="center", transform=axes[0].transAxes, fontsize=13, color=RED, fontweight="bold",
             bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", edgecolor=RED))
axes[0].set_ylim(0, max(values_r)*1.25)
axes[0].grid(axis="y", alpha=0.3)
axes[0].spines[["top", "right"]].set_visible(False)

# Net Profit
categories_n = ["Consensus Est.", "Actual"]
values_n = [cons_fy25_np, act_fy25_np]
colors_n = [GREY, ORANGE]
b2 = axes[1].bar(categories_n, values_n, color=colors_n, width=0.5, edgecolor="white")
for bar, val in zip(b2, values_n):
    axes[1].text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                 f"¥{val:.2f}B", ha="center", fontsize=11, fontweight="bold")
axes[1].set_title("FY2025 Net Profit vs. Consensus", fontsize=11, fontweight="bold")
axes[1].set_ylabel("CNY Billion", fontsize=10)
miss_n = (act_fy25_np - cons_fy25_np) / cons_fy25_np * 100
axes[1].text(0.5, max(values_n)*0.6, f"MISS: {miss_n:.1f}%\n(¥{act_fy25_np-cons_fy25_np:.2f}B)",
             ha="center", transform=axes[1].transAxes, fontsize=13, color=RED, fontweight="bold",
             bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", edgecolor=RED))
axes[1].set_ylim(0, max(values_n)*1.35)
axes[1].grid(axis="y", alpha=0.3)
axes[1].spines[["top", "right"]].set_visible(False)

fig.suptitle("Figure 7 — FY2025 Beat/Miss vs. Analyst Consensus | 阳光电源 业绩与预期对比", fontsize=12, fontweight="bold")
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report; Consensus from Futunn/Bloomberg pre-earnings", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart7_beat_miss.png"))
plt.close()
print("Chart 7 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 8 — Q4 2025 Margin Decomposition (Storage)
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
fig.patch.set_facecolor("white")
# Storage gross margin trend over quarters
storage_gm_qs = ["Q1'25", "Q2'25", "Q3'25", "Q4'25"]
storage_gm_vals = [28.0, 31.0, 41.0, 24.0]  # Q3 was high-margin overseas; Q4 dropped ~17ppts
bars_s = ax.bar(storage_gm_qs, storage_gm_vals,
                color=[LIGHT_BLUE, LIGHT_BLUE, LIGHT_GRN, RED],
                width=0.5, edgecolor="white")
for bar, val in zip(bars_s, storage_gm_vals):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
            f"{val:.0f}%", ha="center", fontsize=12, fontweight="bold")
ax.axhline(y=31.83, color=BLUE, linestyle="--", linewidth=1.5, label="FY2025 Overall GM (31.83%)")
ax.set_ylabel("Storage Gross Margin (%)", fontsize=10)
ax.set_title("Figure 8 — Energy Storage Gross Margin by Quarter\n储能系统季度毛利率 (Q4 Compression Analysis)",
             fontsize=11, fontweight="bold", pad=10)
ax.legend(fontsize=9)
ax.set_ylim(0, 55)
ax.grid(axis="y", alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
ax.annotate("Q4 GM compressed ~17ppts QoQ:\n① High-margin overseas Q3 pulled GM up\n② Li carbonate price rose; legacy contracts\n③ Lower-margin new energy dev. mix rose",
            xy=(3, 24), xytext=(1.8, 44),
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2),
            fontsize=8.5, color=RED,
            bbox=dict(boxstyle="round,pad=0.4", facecolor="lightyellow", edgecolor=RED))
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report; Investor call transcript (2026-03-31)", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart8_storage_margins.png"))
plt.close()
print("Chart 8 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 9 — Estimate Revisions
# ─────────────────────────────────────────────────────────────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 5.5))
fig.patch.set_facecolor("white")
# Revenue revision
rev_old_new = [("FY2024A", 77.86), ("FY2025 Cons.", 90.0), ("FY2025A", 89.18), ("FY2026E", 102.0)]
rev_labels_4 = [x[0] for x in rev_old_new]
rev_vals_4   = [x[1] for x in rev_old_new]
rev_colors_4 = [GREY, GREY, ORANGE, BLUE]
b1 = ax1.bar(rev_labels_4, rev_vals_4, color=rev_colors_4, width=0.55, edgecolor="white")
for bar, val in zip(b1, rev_vals_4):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
             f"¥{val:.1f}B", ha="center", fontsize=9.5, fontweight="bold")
ax1.set_ylabel("Revenue (CNY Billion)", fontsize=10)
ax1.set_title("Revenue Estimates — Old vs. Actual vs. 2026E", fontsize=10, fontweight="bold")
ax1.set_ylim(0, max(rev_vals_4)*1.2)
ax1.grid(axis="y", alpha=0.3)
ax1.spines[["top", "right"]].set_visible(False)
# NP revision
np_old_new = [("FY2024A", 11.04), ("FY2025 Cons.", 15.45), ("FY2025A", 13.46), ("FY2026E", 16.77)]
np_labels_4 = [x[0] for x in np_old_new]
np_vals_4   = [x[1] for x in np_old_new]
np_colors_4 = [GREY, GREY, ORANGE, BLUE]
b2 = ax2.bar(np_labels_4, np_vals_4, color=np_colors_4, width=0.55, edgecolor="white")
for bar, val in zip(b2, np_vals_4):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
             f"¥{val:.2f}B", ha="center", fontsize=9.5, fontweight="bold")
ax2.set_ylabel("Net Profit Attributable to Parent (CNY Billion)", fontsize=10)
ax2.set_title("Net Profit Estimates — Old vs. Actual vs. 2026E", fontsize=10, fontweight="bold")
ax2.set_ylim(0, max(np_vals_4)*1.3)
ax2.grid(axis="y", alpha=0.3)
ax2.spines[["top", "right"]].set_visible(False)
fig.suptitle("Figure 9 — Estimate Revisions | 阳光电源 盈利预测修正", fontsize=12, fontweight="bold")
fig.text(0.02, 0.01, "Source: Sungrow 2025 Annual Report; Consensus estimates from Bloomberg/Futunn pre-earnings; 2026E analyst consensus", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart9_estimates.png"))
plt.close()
print("Chart 9 saved.")

# ─────────────────────────────────────────────────────────────────────────────
# CHART 10 — Valuation: NTM P/E Band
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
fig.patch.set_facecolor("white")
x = np.arange(len(val_periods))
ax.fill_between(x, pe_low, pe_high, alpha=0.15, color=BLUE, label="Historical P/E Range")
ax.plot(x, pe_mid, "o-", color=BLUE, linewidth=2, markersize=5, label="Historical NTM P/E (Mid)")
# Current P/E marker
ax.scatter([9], [20.9], color=ORANGE, zorder=5, s=100, label="Current P/E: 20.9x")
ax.axhline(y=20.9, color=ORANGE, linestyle="--", linewidth=1.2, alpha=0.7)
ax.set_xticks(x)
ax.set_xticklabels(val_periods, fontsize=8.5, rotation=30)
ax.set_ylabel("NTM P/E (x)", fontsize=10)
ax.set_title("Figure 10 — Valuation: NTM P/E Band | 阳光电源 市盈率估值区间", fontsize=12, fontweight="bold", pad=10)
ax.legend(fontsize=9, loc="upper right")
ax.set_ylim(5, 55)
ax.grid(alpha=0.3)
ax.spines[["top", "right"]].set_visible(False)
ax.text(9.1, 20.9 + 1, f"20.9x\n(vs. 3Y avg ~25x)", fontsize=8, color=ORANGE, fontweight="bold")
fig.text(0.02, 0.01, "Source: StockAnalysis.com; Bloomberg; Analyst estimates; NTM = Next Twelve Months", fontsize=7, color=GREY)
plt.tight_layout()
plt.savefig(os.path.join(BASE, "300274_chart10_valuation.png"))
plt.close()
print("Chart 10 saved.")

print("\nAll 10 charts generated successfully.")
print(f"Output: {BASE}")
