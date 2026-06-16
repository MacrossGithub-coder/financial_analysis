"""
Build RKLB (Rocket Lab USA) Q1 FY2026 Earnings Update Charts
Generates 10 PNG charts to output/RKLB/
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/RKLB/"
os.makedirs(OUT, exist_ok=True)

BLUE  = "#1B3A6B"
LBLUE = "#4A90D9"
TEAL  = "#2E8B6A"
GRAY  = "#8C8C8C"
RED   = "#C0392B"
GREEN = "#27AE60"
GOLD  = "#F39C12"
BG    = "#F8F9FA"

plt.rcParams.update({
    "font.family": "Times New Roman",
    "font.size": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
})

# ─── DATA ─────────────────────────────────────────────────────────────────────

quarters_9 = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]
revenue_9  = [92.8, 106.0, 105.0, 132.4, 122.6, 144.0, 155.0, 179.7, 200.3]

# Segment breakdown (Q1 2026 vs Q1 2025)
seg_labels_q1 = ["Launch Services", "Space Systems"]
seg_q1_26 = [63.7, 136.7]
seg_q1_25 = [35.6, 86.9]

# Gross margin quarterly (9 quarters)
gm_quarters = quarters_9
gm_gaap     = [20.0, 24.0, 26.0, 29.0, 28.7, 33.0, 37.0, 37.9, 38.2]
gm_nongaap  = [28.0, 32.0, 34.0, 37.0, 33.4, 38.0, 41.6, 44.0, 43.0]

# Adj. EBITDA quarterly (5 quarters)
ebitda_qtrs = ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]
ebitda_vals = [-30.0, -29.0, -26.3, -17.4, -11.8]

# EPS quarterly (GAAP)
eps_qtrs = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]
eps_vals = [-0.12, -0.11, -0.10, -0.09, -0.12, -0.10, -0.08, -0.09, -0.07]

# Backlog progression ($B)
backlog_qtrs = ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]
backlog_vals = [1.06, 1.30, 1.50, 1.85, 2.20]

# Launch cadence
launch_qtrs = ["Q1'24", "Q2'24", "Q3'24", "Q4'24", "Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26"]
launches    = [3, 4, 5, 4, 5, 5, 4, 7, 4]

# Annual revenue
annual_rev_years = ["2021", "2022", "2023", "2024", "2025", "2026E"]
annual_rev_vals  = [29.5, 211.0, 245.0, 436.2, 601.8, 890.0]

# YoY growth by quarter
yoy_growth = [69.0, 71.0, 55.0, 121.0, 32.1, 35.8, 47.6, 35.7, 63.5]


# ─── CHART 1: Quarterly Revenue (9 quarters) ──────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG)
ax.set_facecolor(BG)
colors = [GRAY if q != "Q1'26" else BLUE for q in quarters_9]
bars = ax.bar(quarters_9, revenue_9, color=colors, edgecolor="white", width=0.6, zorder=3)
for bar, val in zip(bars, revenue_9):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
            f"${val:.0f}M" if val >= 100 else f"${val:.1f}M",
            ha="center", va="bottom", fontsize=9,
            fontweight="bold" if val == 200.3 else "normal",
            color=BLUE if val == 200.3 else "black")
ax.set_title("Quarterly Revenue (Q1 2024 – Q1 2026)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Revenue (US$M)", fontsize=10)
ax.set_ylim(0, 250)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:.0f}M"))
ax.axhline(200, color=BLUE, linewidth=0.8, linestyle="--", alpha=0.4, zorder=2)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.text(0.98, 0.97, "First $200M+ Quarter", transform=ax.transAxes, ha="right", va="top",
        fontsize=8, color=BLUE, style="italic")
plt.tight_layout()
plt.savefig(OUT + "rklb_chart1_quarterly_revenue.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 1 saved.")

# ─── CHART 2: YoY Revenue Growth ──────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_colors = [GREEN if g > 0 else RED for g in yoy_growth]
bars = ax.bar(quarters_9, yoy_growth, color=bar_colors, edgecolor="white", width=0.6, zorder=3)
for bar, val in zip(bars, yoy_growth):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
            f"{val:.0f}%", ha="center", va="bottom", fontsize=9)
ax.set_title("Year-over-Year Revenue Growth (%)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("YoY Growth (%)", fontsize=10)
ax.set_ylim(0, 150)
ax.axhline(63.5, color=BLUE, linewidth=1.2, linestyle="--", alpha=0.5, zorder=2)
ax.text(8.4, 66, "Q1'26\n+63.5%", fontsize=7.5, color=BLUE, ha="right")
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart2_yoy_growth.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 2 saved.")

# ─── CHART 3: Q1 2026 Segment Revenue (Pie + YoY Bar) ──────────────────────
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 5), facecolor=BG)
fig.patch.set_facecolor(BG)

ax1.set_facecolor(BG)
wedge_colors = [LBLUE, BLUE]
wedges, texts, autotexts = ax1.pie(
    seg_q1_26, labels=seg_labels_q1, autopct="%1.0f%%",
    colors=wedge_colors, startangle=90,
    wedgeprops=dict(edgecolor="white", linewidth=2),
    textprops=dict(fontsize=10)
)
for at in autotexts:
    at.set_fontsize(10); at.set_fontweight("bold"); at.set_color("white")
ax1.set_title("Q1 2026 Revenue Mix\n($200.3M Total)", fontsize=11, fontweight="bold", color=BLUE)

x = np.arange(2)
width = 0.35
ax2.set_facecolor(BG)
b1 = ax2.bar(x - width/2, seg_q1_25, width, label="Q1 2025", color=GRAY, edgecolor="white")
b2 = ax2.bar(x + width/2, seg_q1_26, width, label="Q1 2026", color=BLUE, edgecolor="white")
ax2.set_xticks(x); ax2.set_xticklabels(["Launch Services", "Space Systems"])
ax2.set_ylabel("Revenue (US$M)"); ax2.set_ylim(0, 180)
ax2.set_title("Segment Revenue YoY\n(Q1 2025 vs Q1 2026)", fontsize=11, fontweight="bold", color=BLUE)
for bar in list(b1) + list(b2):
    ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
             f"${bar.get_height():.1f}M", ha="center", va="bottom", fontsize=8.5)
ax2.text(0, 55, "+79%", ha="center", fontsize=9, color=GREEN, fontweight="bold")
ax2.text(1, 120, "+57%", ha="center", fontsize=9, color=GREEN, fontweight="bold")
ax2.legend(fontsize=9)
ax2.grid(axis="y", linestyle="--", alpha=0.4)
ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)

plt.suptitle("Revenue Segment Analysis", fontsize=13, fontweight="bold", color=BLUE, y=1.01)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart3_segment_revenue.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 3 saved.")

# ─── CHART 4: Beat / Miss Summary Table ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4.5), facecolor=BG)
ax.set_facecolor(BG)
ax.axis("off")

table_data = [
    ["Metric",           "Consensus Est.",   "Actual Q1'26",  "Variance",       "Result"],
    ["Revenue",          "$189.4M",          "$200.3M",       "+$10.9M (+5.8%)", "BEAT"],
    ["GAAP EPS",         "($0.08)",          "($0.07)",       "+$0.01",          "BEAT"],
    ["Adj. EBITDA",      "($26.0M)",         "($11.8M)",      "+$14.2M",         "BEAT"],
    ["GAAP Gross Margin","34–36% (guide)",   "38.2%",         "+220–420bps",     "BEAT"],
    ["Non-GAAP GM",      "39–41% (guide)",   "43.0%",         "+200–400bps",     "BEAT"],
]
col_widths = [0.18, 0.20, 0.18, 0.22, 0.12]
tbl = ax.table(
    cellText=table_data, cellLoc="center", loc="center",
    bbox=[0.0, 0.0, 1.0, 1.0], colWidths=col_widths
)
tbl.auto_set_font_size(False)
tbl.set_fontsize(10)
for (r, c), cell in tbl.get_celld().items():
    cell.set_edgecolor("#CCCCCC")
    if r == 0:
        cell.set_facecolor(BLUE)
        cell.set_text_props(color="white", fontweight="bold")
        cell.set_height(0.18)
    else:
        cell.set_facecolor(BG)
        cell.set_height(0.15)
        if c == 4:
            cell.set_facecolor("#D5F5E3")
            cell.set_text_props(color=GREEN, fontweight="bold")

ax.set_title("Q1 2026 Beat / Miss Summary — All Metrics Beat", fontsize=13, fontweight="bold", color=BLUE, pad=15)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart4_beat_miss.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 4 saved.")

# ─── CHART 5: Gross Margin Trend ─────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG)
ax.set_facecolor(BG)
ax.plot(range(len(gm_quarters)), gm_gaap, color=BLUE, marker="o", linewidth=2, markersize=7,
        label="GAAP Gross Margin", zorder=3)
ax.plot(range(len(gm_quarters)), gm_nongaap, color=TEAL, marker="s", linewidth=2, markersize=7,
        linestyle="--", label="Non-GAAP Gross Margin", zorder=3)
for i, (g, ng) in enumerate(zip(gm_gaap, gm_nongaap)):
    ax.annotate(f"{g:.1f}%", (i, g), textcoords="offset points",
                xytext=(0, 8), ha="center", fontsize=8, color=BLUE)
    ax.annotate(f"{ng:.1f}%", (i, ng), textcoords="offset points",
                xytext=(0, -14), ha="center", fontsize=8, color=TEAL)
ax.fill_between(range(len(gm_quarters)), gm_gaap, alpha=0.1, color=BLUE)
ax.set_xticks(range(len(gm_quarters))); ax.set_xticklabels(gm_quarters)
ax.set_ylabel("Gross Margin (%)", fontsize=10)
ax.set_ylim(10, 55)
ax.set_title("Gross Margin Trend (Q1 2024 – Q1 2026)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.legend(loc="lower right", fontsize=9)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart5_gross_margin.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 5 saved.")

# ─── CHART 6: Adjusted EBITDA Trend ──────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols = [RED if v < 0 else GREEN for v in ebitda_vals]
bars = ax.bar(ebitda_qtrs, ebitda_vals, color=bar_cols, edgecolor="white", width=0.5, zorder=3)
for bar, val in zip(bars, ebitda_vals):
    ax.text(bar.get_x() + bar.get_width()/2, val - 1.5,
            f"(${abs(val):.1f}M)", ha="center", va="top", fontsize=9.5, color="white", fontweight="bold")
ax.axhline(0, color="black", linewidth=0.8)
ax.set_title("Adjusted EBITDA — Narrowing Losses (Q1'25 – Q1'26)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Adj. EBITDA (US$M)", fontsize=10)
ax.set_ylim(-40, 5)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.annotate("61% improvement\nYoY", xy=(4, -11.8), xytext=(4, 2),
            arrowprops=dict(arrowstyle="->", color=GREEN, lw=1.5),
            fontsize=9, color=GREEN, ha="center", fontweight="bold")
plt.tight_layout()
plt.savefig(OUT + "rklb_chart6_adj_ebitda.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 6 saved.")

# ─── CHART 7: EPS Quarterly Trend ──────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols_eps = [GRAY if q != "Q1'26" else BLUE for q in eps_qtrs]
bars = ax.bar(eps_qtrs, eps_vals, color=bar_cols_eps, edgecolor="white", width=0.55, zorder=3)
for bar, val in zip(bars, eps_vals):
    ax.text(bar.get_x() + bar.get_width()/2, val - 0.003,
            f"(${abs(val):.2f})", ha="center", va="top", fontsize=8.5, color="white", fontweight="bold")
ax.axhline(0, color="black", linewidth=0.8)
ax.set_title("GAAP EPS Quarterly Trend (Q1 2024 – Q1 2026)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("GAAP EPS ($)", fontsize=10)
ax.set_ylim(-0.16, 0.02)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.text(0.98, 0.97, "Losses narrowing steadily", transform=ax.transAxes, ha="right", va="top",
        fontsize=8, color=GREEN, style="italic")
plt.tight_layout()
plt.savefig(OUT + "rklb_chart7_eps_trend.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 7 saved.")

# ─── CHART 8: Backlog Growth ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols_bl = [GRAY if q != "Q1'26" else BLUE for q in backlog_qtrs]
bars = ax.bar(backlog_qtrs, backlog_vals, color=bar_cols_bl, edgecolor="white", width=0.55, zorder=3)
for bar, val in zip(bars, backlog_vals):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.03,
            f"${val:.2f}B", ha="center", va="bottom", fontsize=10,
            fontweight="bold" if val == 2.20 else "normal",
            color=BLUE if val == 2.20 else "black")
ax.plot(range(len(backlog_qtrs)), backlog_vals, color=LBLUE, marker="o", linewidth=1.5, markersize=5, zorder=4)
ax.set_title("Backlog Growth (Q1 2025 – Q1 2026)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Backlog (US$B)", fontsize=10)
ax.set_ylim(0, 2.8)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"${x:.1f}B"))
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
ax.text(0.98, 0.85, "+108% YoY", transform=ax.transAxes, ha="right", va="top",
        fontsize=10, color=GREEN, fontweight="bold")
plt.tight_layout()
plt.savefig(OUT + "rklb_chart8_backlog.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 8 saved.")

# ─── CHART 9: Launch Cadence ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG)
ax.set_facecolor(BG)
bar_cols_l = []
for q in launch_qtrs:
    if q == "Q1'26":
        bar_cols_l.append(TEAL)
    elif "'25" in q:
        bar_cols_l.append(BLUE)
    else:
        bar_cols_l.append(LBLUE)
bars = ax.bar(launch_qtrs, launches, color=bar_cols_l, edgecolor="white", width=0.6, zorder=3)
for bar, val in zip(bars, launches):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
            str(val), ha="center", va="bottom", fontsize=10, fontweight="bold")
ax.set_title("Electron Launch Cadence (Q1 2024 – Q1 2026)", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Number of Launches", fontsize=10)
ax.set_ylim(0, 10)
ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
p1 = mpatches.Patch(color=LBLUE, label="FY2024 (16 total)")
p2 = mpatches.Patch(color=BLUE, label="FY2025 (21 total)")
p3 = mpatches.Patch(color=TEAL, label="FY2026 (Q1: 4 missions)")
ax.legend(handles=[p1, p2, p3], fontsize=9)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart9_launch_cadence.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 9 saved.")

# ─── CHART 10: Q2 2026 Guidance vs Street & Revenue Trend ───────────────────
fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
ax.set_facecolor(BG)
guide_qtrs = ["Q3'25", "Q4'25", "Q1'26", "Q2'26E"]
guide_rev  = [155.0, 179.7, 200.3, None]

for i, (q, v) in enumerate(zip(guide_qtrs[:-1], guide_rev[:-1])):
    col = BLUE if q == "Q1'26" else GRAY
    ax.bar(i, v, color=col, edgecolor="white", width=0.5, zorder=3)
    ax.text(i, v + 2, f"${v:.0f}M", ha="center", va="bottom", fontsize=10)

# Guidance range bar
guide_low, guide_high = 225.0, 240.0
ax.bar(3, guide_high - guide_low, bottom=guide_low,
       color=GOLD, edgecolor="white", width=0.5, alpha=0.8, zorder=3, label="Guidance Range")
ax.text(3, (guide_low + guide_high)/2, f"$225–240M", ha="center", va="center",
        fontsize=9, fontweight="bold", color="#7D6608")

# Street estimate line
street_est = 207.5
ax.axhline(street_est, color=RED, linewidth=1.2, linestyle=":", alpha=0.7, zorder=2)
ax.text(3.3, street_est + 1, f"Street: ${street_est:.0f}M", fontsize=8, color=RED, ha="left")

ax.set_xticks(range(4))
ax.set_xticklabels(guide_qtrs)
ax.set_title("Revenue Trend & Q2 2026 Guidance vs. Street", fontsize=13, fontweight="bold", color=BLUE, pad=10)
ax.set_ylabel("Revenue (US$M)", fontsize=10)
ax.set_ylim(120, 260)
ax.legend(fontsize=9)
ax.grid(axis="y", linestyle="--", alpha=0.4, zorder=0)
plt.tight_layout()
plt.savefig(OUT + "rklb_chart10_guidance.png", dpi=150, bbox_inches="tight")
plt.close()
print("Chart 10 saved.")

print("\nAll 10 charts saved to:", OUT)
