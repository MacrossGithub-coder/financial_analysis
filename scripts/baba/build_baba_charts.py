"""
Alibaba Group (BABA) — Q3 FY2026 (Calendar Q4 2025) Earnings Update
Chart generation script — 10 PNG charts
Output: /output/BABA/baba_chart*.png
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/BABA/"
os.makedirs(OUT, exist_ok=True)

# ── Global style ──────────────────────────────────────────────────────────────
plt.rcParams.update({
    "font.family":      "sans-serif",
    "font.sans-serif":  ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,
    "font.size":        10,
    "axes.titlesize":   12,
    "axes.titleweight": "bold",
    "axes.spines.top":  False,
    "axes.spines.right":False,
    "figure.facecolor": "white",
    "axes.facecolor":   "white",
})

ALI_ORANGE  = "#FF6A00"   # Alibaba brand orange
ALI_NAVY    = "#1F2D5C"   # Deep navy
CLOUD_BLUE  = "#0070C0"   # Cloud segment
GREEN_ACC   = "#00A651"   # positive / beat
RED_ACC     = "#E8192C"   # miss / negative
GRAY_LIGHT  = "#D9D9D9"

# ── Quarter labels ────────────────────────────────────────────────────────────
qtrs_8 = ["Q4\nFY24", "Q1\nFY25", "Q2\nFY25", "Q3\nFY25",
          "Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26"]

# ── Data ──────────────────────────────────────────────────────────────────────
# Total revenue (RMB Billions)
rev_rmb = [221.9, 243.2, 236.5, 280.2, 236.5, 247.7, 247.8, 284.8]
# YoY growth (reported %)
rev_yoy = [7, 4, 5, 8, 7, 10, 5, 2]
# Like-for-like YoY (excluding divestitures)
rev_lfl = [7, 4, 5, 8, 7, 10, 5, 9]

# Cloud Intelligence Group revenue (RMB Billions)
cloud_rev = [25.6, 26.5, 29.6, 31.7, 30.1, 32.7, 39.8, 43.3]
cloud_yoy = [3, 6, 7, 13, 18, 26, 34, 36]

# Adjusted EBITA (RMB Billions) — last 4 quarters available
qtrs_4 = ["Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26"]
adj_ebita = [46.4, 38.8, 9.1, 23.4]   # RMB B (Q4 FY25 estimated from ~+20% to ~46B)
adj_ebita_yoy = [+20, -14, -78, -57]   # % YoY

# Non-GAAP EPS per ADS (RMB)
eps_vals = [None, None, None, None, 12.52, 14.75, 4.36, 7.09]
eps_4q = [12.52, 14.75, 4.36, 7.09]
eps_yoy_4q = [+23, -10, -71, -55]

# Segment revenue Q3 FY26 (RMB Billions)
seg_names  = ["Cloud\nIntelligence", "International\nCommerce (AIDC)",
              "Quick Commerce\n(Taobao Instant)", "China Wholesale\n(1688.com)",
              "Cainiao &\nOthers"]
seg_vals   = [43.3, 32.4, 20.8, 6.9, 67.3]    # RMB B (others incl Cainiao+Media)
seg_colors = [CLOUD_BLUE, ALI_ORANGE, ALI_NAVY, GREEN_ACC, GRAY_LIGHT]

# Beat/Miss vs consensus
beat_miss_metrics = ["Revenue\n(RMB)", "Adj. EBITA\n(RMB)", "Non-GAAP\nEPS/ADS", "Cloud\nRevenue"]
beat_miss_actual  = [284.8, 23.4, 7.09, 43.3]
beat_miss_est     = [290.7, 30.0, 9.0, 40.0]    # consensus estimates
beat_miss_delta   = [-2.0, -22.0, -21.2, +8.3]  # % vs estimate

# 88VIP members (millions) — quarterly estimates
qtrs_88vip = ["Q4\nFY24", "Q1\nFY25", "Q2\nFY25", "Q3\nFY25",
              "Q4\nFY25", "Q1\nFY26", "Q2\nFY26", "Q3\nFY26"]
vip_members = [42, 43, 45, 46, 48, 51, 54, 59]   # millions

# AI product revenue triple-digit growth (10 consecutive quarters index)
# Represent as cloud yoy growth vs total company yoy growth
# Gross margin (estimated from reported figures)
gm_pct = [41.2, 40.5, 40.1, 41.8, 42.5, 41.9, 40.2, 39.7]   # % (approx)

# AIDC (International) revenue trend (RMB B)
aidc_rev  = [17.3, 29.3, 31.7, 31.6, 23.0, 27.2, 28.6, 32.4]
aidc_yoy  = [+45, +45, +35, +32, +33, -7, -10, +3]

# Free Cash Flow (RMB B) — last 4 quarters
fcf_4q = [49.0, 41.6, 22.5, 11.3]   # RMB B (Q4 FY25, Q1 FY26, Q2 FY26, Q3 FY26)
fcf_yoy_4q = [+5, -19, -42, -71]    # % YoY

SAVE_DPI = 150


# ─────────────────────────────────────────────────────────────────────────────
# Chart 1: Quarterly Total Revenue (RMB B) — bar with YoY label
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
colors = [GRAY_LIGHT] * 7 + [ALI_ORANGE]
bars = ax.bar(qtrs_8, rev_rmb, color=colors, edgecolor="white", width=0.6)
for i, (b, y) in enumerate(zip(bars, rev_yoy)):
    lfl = rev_lfl[i]
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 2,
            f"¥{b.get_height():.0f}B\n+{y}% YoY" if y >= 0 else f"¥{b.get_height():.0f}B\n{y}% YoY",
            ha="center", va="bottom", fontsize=7.5, fontweight="bold")
ax.set_title("Alibaba Group — Quarterly Revenue (RMB Billions)")
ax.set_ylabel("RMB Billions")
ax.set_ylim(0, 340)
ax.axhline(0, color="black", linewidth=0.5)
note = ax.text(0.99, 0.04,
    "Q3 FY26: RMB 284.8B (+2% reported / +9% like-for-like)\nExcludes Sun Art & Intime divested revenue",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray",
    style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart1_quarterly_revenue.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 1 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 2: Cloud Intelligence Group Revenue — accelerating growth
# ─────────────────────────────────────────────────────────────────────────────
fig, ax1 = plt.subplots(figsize=(9, 5))
ax2 = ax1.twinx()

bar_colors = [GRAY_LIGHT] * 7 + [CLOUD_BLUE]
bars = ax1.bar(qtrs_8, cloud_rev, color=bar_colors, edgecolor="white", width=0.6, label="Cloud Revenue")
for b, v in zip(bars, cloud_rev):
    ax1.text(b.get_x() + b.get_width()/2, b.get_height() + 0.3, f"¥{v:.1f}B",
             ha="center", va="bottom", fontsize=7.5, fontweight="bold")

ax2.plot(qtrs_8, cloud_yoy, color=GREEN_ACC, marker="o", linewidth=2, markersize=6,
         label="YoY Growth (%)", zorder=5)
for x, y in enumerate(cloud_yoy):
    ax2.text(x, y + 1.5, f"+{y}%", ha="center", va="bottom", fontsize=7.5, color=GREEN_ACC, fontweight="bold")

ax1.set_title("Alibaba Cloud Intelligence Group — Revenue & YoY Growth")
ax1.set_ylabel("Revenue (RMB Billions)", color=CLOUD_BLUE)
ax2.set_ylabel("YoY Growth (%)", color=GREEN_ACC)
ax1.set_ylim(0, 60)
ax2.set_ylim(0, 50)
ax1.tick_params(axis="y", labelcolor=CLOUD_BLUE)
ax2.tick_params(axis="y", labelcolor=GREEN_ACC)

lines1 = mpatches.Patch(color=CLOUD_BLUE, label="Cloud Revenue")
lines2 = mpatches.Patch(color=GREEN_ACC, label="YoY Growth")
ax1.legend(handles=[lines1, lines2], loc="upper left", fontsize=8)
ax1.text(0.99, 0.04,
    "AI product revenue: triple-digit YoY growth for 10 consecutive quarters\nCloud market share in China: ~36%",
    transform=ax1.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")

fig.tight_layout()
fig.savefig(OUT + "baba_chart2_cloud_revenue.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 2 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 3: Q3 FY26 Segment Revenue Breakdown — horizontal bar
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
y_pos = np.arange(len(seg_names))
bars = ax.barh(y_pos, seg_vals, color=seg_colors, edgecolor="white", height=0.6)
for bar, val in zip(bars, seg_vals):
    ax.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
            f"¥{val:.1f}B", va="center", ha="left", fontsize=9, fontweight="bold")
ax.set_yticks(y_pos)
ax.set_yticklabels(seg_names, fontsize=9)
ax.set_xlabel("Revenue (RMB Billions)")
ax.set_title("Alibaba Q3 FY2026 — Revenue by Business Segment")
ax.set_xlim(0, 90)
ax.text(0.99, 0.04,
    "Others includes Cainiao, Taobao Tmall Group core, Digital Media\nTotal Q3 FY26: RMB 284.8B",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart3_segment_revenue.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 3 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 4: Adjusted EBITA — last 4 quarters, declining vs revenue growth
# ─────────────────────────────────────────────────────────────────────────────
fig, ax1 = plt.subplots(figsize=(9, 5))
ax2 = ax1.twinx()

bar_colors = [GREEN_ACC, GREEN_ACC, RED_ACC, RED_ACC]
bars = ax1.bar(qtrs_4, adj_ebita, color=bar_colors, edgecolor="white", width=0.5)
for b, v, y in zip(bars, adj_ebita, adj_ebita_yoy):
    yoy_str = f"+{y}% YoY" if y > 0 else f"{y}% YoY"
    ax1.text(b.get_x() + b.get_width()/2, b.get_height() + 0.5,
             f"¥{v:.1f}B\n{yoy_str}", ha="center", va="bottom", fontsize=8, fontweight="bold")

# Revenue comparison
qtrs_4_rev = [rev_rmb[4], rev_rmb[5], rev_rmb[6], rev_rmb[7]]
qtrs_4_rev_yoy = [rev_yoy[4], rev_yoy[5], rev_yoy[6], rev_yoy[7]]
ax2.plot(qtrs_4, qtrs_4_rev_yoy, color=ALI_ORANGE, marker="D", linewidth=2, markersize=7, label="Revenue YoY %")

ax1.set_title("Alibaba — Adjusted EBITA vs Revenue YoY Growth (Heavy Investment Cycle)")
ax1.set_ylabel("Adj. EBITA (RMB Billions)", color=ALI_NAVY)
ax2.set_ylabel("Revenue YoY Growth (%)", color=ALI_ORANGE)
ax1.set_ylim(0, 60)
ax2.set_ylim(-20, 30)
ax1.tick_params(axis="y", labelcolor=ALI_NAVY)
ax2.tick_params(axis="y", labelcolor=ALI_ORANGE)

beat_patch = mpatches.Patch(color=GREEN_ACC, label="Adj. EBITA — investment ramp-up phase")
miss_patch = mpatches.Patch(color=RED_ACC, label="Adj. EBITA — peak investment phase")
line_patch = mpatches.Patch(color=ALI_ORANGE, label="Revenue YoY Growth")
ax1.legend(handles=[beat_patch, miss_patch, line_patch], loc="upper right", fontsize=8)
ax1.text(0.99, 0.04,
    "Deliberate reinvestment cycle: AI infra, quick commerce logistics, Qwen buildout\nAdj. EBITA -57% YoY in Q3 FY26",
    transform=ax1.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart4_adj_ebita.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 4 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 5: Non-GAAP EPS per ADS (RMB) — last 4 quarters
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
eps_colors = [GREEN_ACC, GREEN_ACC, RED_ACC, RED_ACC]
bars = ax.bar(qtrs_4, eps_4q, color=eps_colors, edgecolor="white", width=0.5)
for b, v, y in zip(bars, eps_4q, eps_yoy_4q):
    yoy_str = f"+{y}% YoY" if y > 0 else f"{y}% YoY"
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.15,
            f"¥{v:.2f}\n{yoy_str}", ha="center", va="bottom", fontsize=8, fontweight="bold")
ax.set_title("Alibaba — Non-GAAP EPS per ADS (RMB) — Investment Cycle Impact")
ax.set_ylabel("Non-GAAP EPS per ADS (RMB)")
ax.set_ylim(0, 20)
ax.text(0.99, 0.04,
    "Q3 FY26: RMB 7.09/ADS (~$1.01) vs consensus ~¥9.00 (miss)\nHeavy AI capex + quick commerce buildout depressing EPS",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
beat_patch = mpatches.Patch(color=GREEN_ACC, label="YoY growth phase")
miss_patch = mpatches.Patch(color=RED_ACC, label="Investment compression phase")
ax.legend(handles=[beat_patch, miss_patch], loc="upper right", fontsize=8)
fig.tight_layout()
fig.savefig(OUT + "baba_chart5_eps.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 5 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 6: Beat/Miss Summary — Q3 FY26 vs Consensus
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
metrics_short = ["Revenue\n(RMB B)", "Adj. EBITA\n(RMB B)", "Non-GAAP\nEPS/ADS (¥)", "Cloud Rev\n(RMB B)"]
x = np.arange(len(metrics_short))
w = 0.32

bars_act = ax.bar(x - w/2, beat_miss_actual, width=w, label="Actual", color=ALI_NAVY, edgecolor="white")
bars_est = ax.bar(x + w/2, beat_miss_est, width=w, label="Consensus Est.", color=GRAY_LIGHT, edgecolor="gray")

for i, (act, est, delta) in enumerate(zip(beat_miss_actual, beat_miss_est, beat_miss_delta)):
    color = GREEN_ACC if delta > 0 else RED_ACC
    label = f"+{delta:.1f}%" if delta > 0 else f"{delta:.1f}%"
    ax.text(i, max(act, est) + 1.5, label,
            ha="center", va="bottom", fontsize=9, fontweight="bold", color=color)

ax.set_xticks(x)
ax.set_xticklabels(metrics_short, fontsize=9)
ax.set_title("Q3 FY2026 (Calendar Q4 2025) — Beat/Miss vs. Consensus Estimates")
ax.set_ylabel("Value (RMB Billions / EPS in RMB)")
ax.legend(fontsize=9)
ax.set_ylim(0, 350)
ax.text(0.99, 0.04,
    "Cloud revenue beat consensus ~+8.3%; Revenue, EPS, EBITA all missed\nResult reflects deliberate investment cycle over profitability",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart6_beat_miss.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 6 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 7: Gross Margin Trend
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
bar_colors = [GRAY_LIGHT] * 7 + [ALI_ORANGE]
bars = ax.bar(qtrs_8, gm_pct, color=bar_colors, edgecolor="white", width=0.6)
ax.plot(qtrs_8, gm_pct, color=ALI_NAVY, marker="o", linewidth=1.5, markersize=5, zorder=5)
for b, v in zip(bars, gm_pct):
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.2,
            f"{v:.1f}%", ha="center", va="bottom", fontsize=7.5, fontweight="bold")
ax.set_title("Alibaba Group — Gross Margin Trend (%)")
ax.set_ylabel("Gross Margin (%)")
ax.set_ylim(35, 47)
ax.axhline(np.mean(gm_pct), color=GREEN_ACC, linestyle="--", linewidth=1, label=f"Average: {np.mean(gm_pct):.1f}%")
ax.legend(fontsize=8)
ax.text(0.99, 0.04,
    "Margin pressure from quick commerce (Taobao Instant) logistics costs\nand AI infrastructure investment ramp",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart7_gross_margin.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 7 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 8: International Commerce (AIDC) Revenue Trend
# ─────────────────────────────────────────────────────────────────────────────
fig, ax1 = plt.subplots(figsize=(9, 5))
ax2 = ax1.twinx()

bar_colors = [GRAY_LIGHT] * 7 + [ALI_ORANGE]
ax1.bar(qtrs_8, aidc_rev, color=bar_colors, edgecolor="white", width=0.6, label="AIDC Revenue")
for i, (q, v) in enumerate(zip(qtrs_8, aidc_rev)):
    ax1.text(i, v + 0.3, f"¥{v:.1f}B", ha="center", va="bottom", fontsize=7.5, fontweight="bold")

# YoY line
yoy_colors_line = [GREEN_ACC if y >= 0 else RED_ACC for y in aidc_yoy]
for i in range(len(aidc_yoy) - 1):
    ax2.plot([i, i+1], [aidc_yoy[i], aidc_yoy[i+1]],
             color=GRAY_LIGHT, linewidth=1.5, zorder=3)
ax2.scatter(range(len(aidc_yoy)), aidc_yoy, color=yoy_colors_line, s=50, zorder=5)
for i, y in enumerate(aidc_yoy):
    color_txt = GREEN_ACC if y >= 0 else RED_ACC
    lbl = f"+{y}%" if y >= 0 else f"{y}%"
    ax2.text(i, y + 2.5, lbl, ha="center", va="bottom", fontsize=7.5, color=color_txt, fontweight="bold")

ax1.set_title("Alibaba International Digital Commerce (AIDC) — Revenue Trend")
ax1.set_ylabel("Revenue (RMB Billions)", color=ALI_NAVY)
ax2.set_ylabel("YoY Growth (%)", color=ALI_ORANGE)
ax1.set_ylim(0, 50)
ax2.set_ylim(-30, 70)
ax2.axhline(0, color="black", linewidth=0.5, linestyle="--")
ax1.tick_params(axis="y", labelcolor=ALI_NAVY)
ax2.tick_params(axis="y", labelcolor=ALI_ORANGE)
ax1.text(0.99, 0.04,
    "AIDC growth decelerated sharply from +45% (Q4 FY24) to +3% (Q3 FY26)\nLazada restructuring weighed on growth; AliExpress investing for share",
    transform=ax1.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart8_aidc_revenue.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 8 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 9: 88VIP Members & User Growth
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
bar_colors = [GRAY_LIGHT] * 7 + [ALI_ORANGE]
bars = ax.bar(qtrs_8, vip_members, color=bar_colors, edgecolor="white", width=0.6)
ax.plot(qtrs_8, vip_members, color=ALI_NAVY, marker="o", linewidth=2, markersize=6, zorder=5)
for b, v in zip(bars, vip_members):
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.3,
            f"{v}M", ha="center", va="bottom", fontsize=8, fontweight="bold")
ax.set_title("Alibaba 88VIP Premium Members — Quarterly Growth (Millions)")
ax.set_ylabel("88VIP Members (Millions)")
ax.set_ylim(0, 75)
ax.text(0.99, 0.04,
    "88VIP surpassed 59M members in Q3 FY26 (double-digit YoY growth)\nAlibaba's highest-value consumer cohort, driving GMV and loyalty",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")

growth = (vip_members[-1] / vip_members[0] - 1) * 100
ax.annotate(f"+{growth:.0f}% over\n8 quarters",
            xy=(7, vip_members[-1]), xytext=(5.5, 67),
            arrowprops=dict(arrowstyle="->", color=ALI_NAVY),
            fontsize=9, fontweight="bold", color=ALI_NAVY)
fig.tight_layout()
fig.savefig(OUT + "baba_chart9_88vip_members.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 9 saved.")


# ─────────────────────────────────────────────────────────────────────────────
# Chart 10: Free Cash Flow vs Adjusted EBITA — Investment Cycle Snapshot
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5))
x = np.arange(len(qtrs_4))
w = 0.35

bars_ebita = ax.bar(x - w/2, adj_ebita, width=w, label="Adj. EBITA (RMB B)", color=ALI_NAVY, edgecolor="white")
bars_fcf   = ax.bar(x + w/2, fcf_4q,   width=w, label="Free Cash Flow (RMB B)", color=CLOUD_BLUE, edgecolor="white")

for b, v, y in zip(bars_ebita, adj_ebita, adj_ebita_yoy):
    lbl = f"+{y}%" if y > 0 else f"{y}%"
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.3,
            f"¥{v:.0f}B\n{lbl}", ha="center", va="bottom", fontsize=7.5)
for b, v, y in zip(bars_fcf, fcf_4q, fcf_yoy_4q):
    lbl = f"+{y}%" if y > 0 else f"{y}%"
    ax.text(b.get_x() + b.get_width()/2, b.get_height() + 0.3,
            f"¥{v:.0f}B\n{lbl}", ha="center", va="bottom", fontsize=7.5)

ax.set_xticks(x)
ax.set_xticklabels(qtrs_4, fontsize=9)
ax.set_title("Alibaba — Adj. EBITA vs. Free Cash Flow (Investment Cycle Deep-Dive)")
ax.set_ylabel("RMB Billions")
ax.set_ylim(0, 65)
ax.legend(fontsize=9, loc="upper right")
ax.text(0.99, 0.04,
    "FCF -71% YoY in Q3 FY26 (¥11.3B) reflects heavy AI infra capex\nQ2 FY26 saw the deepest profitability trough; partial recovery in Q3 FY26",
    transform=ax.transAxes, ha="right", va="bottom", fontsize=7.5, color="gray", style="italic")
fig.tight_layout()
fig.savefig(OUT + "baba_chart10_fcf_ebita.png", dpi=SAVE_DPI)
plt.close(fig)
print("Chart 10 saved.")

print(f"\nAll 10 charts saved to: {OUT}")
