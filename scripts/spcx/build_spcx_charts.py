#!/usr/bin/env python3
"""Generate charts for SpaceX (SPCX) Q1 2026 Earnings Update Report."""

import os
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

OUT = os.path.join(os.path.dirname(__file__), "..", "..", "output", "SPCX")
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

SPACEX_BLUE  = "#005288"
DARK_BLUE    = "#003366"
LIGHT_BLUE   = "#4A90D9"
STARLINK_T   = "#00A3E0"
SPACE_ORANGE = "#E87722"
AI_PURPLE    = "#7B2D8E"
GREEN        = "#2E7D32"
RED          = "#D32F2F"
GRAY         = "#757575"
LIGHT_GRAY   = "#BDBDBD"

# ══ Data ════════════════════════════════════════════════════════════════════

# Annual consolidated revenue ($B)
years_annual     = ["FY2023E", "FY2024", "FY2025", "Q1'26\nAnn."]
rev_annual       = [9.5, 14.1, 18.7, 18.8]  # Q1 annualised = 4.694*4

# FY2025 & Q1 2026 segment revenue ($B)
seg_labels       = ["Connectivity\n(Starlink)", "Space\n(Launch/Dragon)", "AI\n(xAI / X)"]
seg_fy25         = [11.39, 4.10, 3.20]
seg_q1_26_ann    = [3.26*4, 0.619*4, 0.815*4]  # annualised

# Segment operating income / (loss) ($B)
seg_opinc_fy25   = [4.40, -0.66, -6.36]
seg_opinc_q1_26  = [1.188, None, -2.50]  # Space Q1 op not separately disclosed; estimate
seg_opinc_q1_est = [1.188, 0.37, -2.50]  # residual for Space: -1.943 - 1.188 - (-2.5) = -0.631? Let me recalculate
# Total Q1 op loss = -1.943; Connectivity = 1.188; AI = -2.5; Space = -1.943 - 1.188 + 2.5 = -0.631
seg_opinc_q1_26  = [1.188, -0.631, -2.500]

# Adjusted EBITDA by segment ($B)
seg_ebitda_fy25  = [4.40, 0.65, -1.24]
ebitda_total     = [None, None, 6.58, 1.13]  # FY23E?, FY24?, FY25, Q1'26

# Starlink subscribers (M)
sub_periods      = ["YE\n2023", "YE\n2024", "YE\n2025", "Q1\n2026"]
subs             = [2.3, 4.4, 8.9, 10.3]

# Starlink ARPU ($/month)
arpu_periods     = ["2023", "2024", "2025", "Q1 2026"]
arpu             = [99, 85, 76, 66]

# CapEx by segment ($B) — FY2025
capex_seg_labels = ["Connectivity", "Space", "AI"]
capex_fy25       = [4.2, 3.8, 12.7]
capex_q1_26      = [None, None, 7.7]  # Only AI disclosed

# ── Chart 1: Annual Revenue Growth ─────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
colors = [LIGHT_GRAY, SPACEX_BLUE, SPACEX_BLUE, LIGHT_BLUE]
bars = ax.bar(years_annual, rev_annual, color=colors, edgecolor="white", width=0.55)
for b, v in zip(bars, rev_annual):
    ax.text(b.get_x() + b.get_width()/2, v + 0.3, f"${v:.1f}B", ha="center",
            va="bottom", fontsize=10, fontweight="bold")
# Growth annotations
growths = [None, 48, 33, None]
for i, g in enumerate(growths):
    if g:
        ax.annotate(f"+{g}% YoY", xy=(i, rev_annual[i]),
                    xytext=(0, -18), textcoords="offset points",
                    ha="center", fontsize=9, color=GREEN, fontweight="bold")
ax.set_ylabel("Revenue ($ Billion)", fontsize=11)
ax.set_title("SpaceX Consolidated Revenue", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 24)
fig.text(0.15, 0.02, "Note: FY2023 estimated; Q1'26 Ann. = Q1 revenue × 4.", fontsize=8, color=GRAY)
fig.savefig(os.path.join(OUT, "spcx_chart1_annual_revenue.png"))
plt.close()

# ── Chart 2: Revenue by Segment — FY2025 vs Q1'26 Annualised ──────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
x = np.arange(len(seg_labels))
w = 0.32
b1 = ax.bar(x - w/2, seg_fy25, w, label="FY2025", color=SPACEX_BLUE, edgecolor="white")
b2 = ax.bar(x + w/2, seg_q1_26_ann, w, label="Q1'26 Annualised", color=LIGHT_BLUE, edgecolor="white")
for b, v in zip(b1, seg_fy25):
    ax.text(b.get_x() + b.get_width()/2, v + 0.15, f"${v:.1f}B", ha="center", fontsize=9, fontweight="bold")
for b, v in zip(b2, seg_q1_26_ann):
    ax.text(b.get_x() + b.get_width()/2, v + 0.15, f"${v:.1f}B", ha="center", fontsize=9, fontweight="bold")
ax.set_xticks(x)
ax.set_xticklabels(seg_labels)
ax.set_ylabel("Revenue ($ Billion)", fontsize=11)
ax.set_title("SpaceX Revenue by Segment", fontsize=13, fontweight="bold", pad=12)
ax.legend(fontsize=9)
ax.set_ylim(0, 16)
fig.savefig(os.path.join(OUT, "spcx_chart2_segment_revenue.png"))
plt.close()

# ── Chart 3: Segment Operating Income / (Loss) ────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
x = np.arange(len(seg_labels))
w = 0.32
colors_fy = [GREEN if v >= 0 else RED for v in seg_opinc_fy25]
colors_q1 = [GREEN if v >= 0 else RED for v in seg_opinc_q1_26]
b1 = ax.bar(x - w/2, seg_opinc_fy25, w, label="FY2025", color=colors_fy, edgecolor="white", alpha=0.7)
b2 = ax.bar(x + w/2, [v*4 for v in seg_opinc_q1_26], w, label="Q1'26 Annualised",
            color=colors_q1, edgecolor="white", alpha=1.0)
for b, v in zip(b1, seg_opinc_fy25):
    ypos = v + 0.15 if v >= 0 else v - 0.4
    ax.text(b.get_x() + b.get_width()/2, ypos, f"${v:.1f}B", ha="center", fontsize=9, fontweight="bold")
for b, v_q in zip(b2, seg_opinc_q1_26):
    v = v_q * 4
    ypos = v + 0.15 if v >= 0 else v - 0.4
    ax.text(b.get_x() + b.get_width()/2, ypos, f"${v:.1f}B", ha="center", fontsize=9, fontweight="bold")
ax.axhline(0, color="black", linewidth=0.8)
ax.set_xticks(x)
ax.set_xticklabels(seg_labels)
ax.set_ylabel("Operating Income / (Loss) ($ Billion)", fontsize=11)
ax.set_title("SpaceX Segment Profitability (Annualised)", fontsize=13, fontweight="bold", pad=12)
ax.legend(fontsize=9)
ax.set_ylim(-12, 7)
fig.savefig(os.path.join(OUT, "spcx_chart3_segment_profitability.png"))
plt.close()

# ── Chart 4: Starlink Subscriber Growth ────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
bars = ax.bar(sub_periods, subs, color=STARLINK_T, edgecolor="white", width=0.5)
for b, v in zip(bars, subs):
    ax.text(b.get_x() + b.get_width()/2, v + 0.15, f"{v:.1f}M", ha="center",
            fontsize=11, fontweight="bold")
# Growth rates
sub_yoy = [None, 91, 102, 49]  # YoY growth %
for i, g in enumerate(sub_yoy):
    if g:
        ax.annotate(f"+{g}%", xy=(i, subs[i]), xytext=(0, -18),
                    textcoords="offset points", ha="center", fontsize=9,
                    color=GREEN, fontweight="bold")
ax.set_ylabel("Subscribers (Millions)", fontsize=11)
ax.set_title("Starlink Subscriber Growth", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 13)
fig.savefig(os.path.join(OUT, "spcx_chart4_starlink_subs.png"))
plt.close()

# ── Chart 5: Starlink ARPU Trend ──────────────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
ax.plot(arpu_periods, arpu, "o-", color=STARLINK_T, linewidth=2.5, markersize=8)
for i, (p, a) in enumerate(zip(arpu_periods, arpu)):
    ax.annotate(f"${a}/mo", (p, a), textcoords="offset points", xytext=(0, 12),
                fontsize=10, fontweight="bold", ha="center", color=DARK_BLUE)
ax.fill_between(arpu_periods, arpu, alpha=0.15, color=STARLINK_T)
ax.set_ylabel("Monthly ARPU ($)", fontsize=11)
ax.set_title("Starlink Average Revenue Per User (ARPU)", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(50, 115)
ax.annotate("ARPU declining as Starlink\nexpands into lower-income markets",
            xy=(3, 66), xytext=(-80, -35), textcoords="offset points",
            fontsize=9, color=RED, fontstyle="italic",
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2))
fig.savefig(os.path.join(OUT, "spcx_chart5_starlink_arpu.png"))
plt.close()

# ── Chart 6: FY2025 Revenue Mix Pie ───────────────────────────────────────
fig, ax = plt.subplots(figsize=(7.5, 5.5))
labels_pie = [f"Connectivity\n(Starlink)\n${seg_fy25[0]:.1f}B",
              f"Space\n(Launch)\n${seg_fy25[1]:.1f}B",
              f"AI\n(xAI / X)\n${seg_fy25[2]:.1f}B"]
colors_pie = [STARLINK_T, SPACE_ORANGE, AI_PURPLE]
wedges, texts, autotexts = ax.pie(seg_fy25, labels=labels_pie, colors=colors_pie,
                                   autopct="%1.0f%%", startangle=90,
                                   textprops={"fontsize": 10},
                                   wedgeprops={"edgecolor": "white", "linewidth": 2})
for at in autotexts:
    at.set_fontsize(12)
    at.set_fontweight("bold")
    at.set_color("white")
ax.set_title("FY2025 Revenue Mix by Segment\n(Total: $18.7B)", fontsize=13, fontweight="bold", pad=12)
fig.savefig(os.path.join(OUT, "spcx_chart6_revenue_mix.png"))
plt.close()

# ── Chart 7: FY2025 CapEx by Segment ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(8, 5.5))
bars = ax.bar(capex_seg_labels, capex_fy25,
              color=[STARLINK_T, SPACE_ORANGE, AI_PURPLE], edgecolor="white", width=0.5)
for b, v in zip(bars, capex_fy25):
    ax.text(b.get_x() + b.get_width()/2, v + 0.2, f"${v:.1f}B", ha="center",
            fontsize=11, fontweight="bold")
    pct = v / sum(capex_fy25) * 100
    ax.text(b.get_x() + b.get_width()/2, v / 2, f"{pct:.0f}%", ha="center",
            fontsize=12, fontweight="bold", color="white")
ax.set_ylabel("Capital Expenditures ($ Billion)", fontsize=11)
ax.set_title("FY2025 CapEx by Segment (Total: $20.7B)", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 16)
fig.savefig(os.path.join(OUT, "spcx_chart7_capex_breakdown.png"))
plt.close()

# ── Chart 8: Q1 2026 Cash Flow Waterfall ──────────────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
cf_labels = ["Adj\nEBITDA", "Working\nCapital", "CapEx", "Other", "Free\nCash Flow"]
cf_vals   = [1.13, 0.7, -10.1, -0.83, -9.1]
cf_bottom = [0, 1.13, 1.83, -8.27, 0]
cf_colors = [GREEN, GREEN, RED, RED, RED]
bars = ax.bar(cf_labels, [abs(v) for v in cf_vals], bottom=[max(0, b) if v >= 0 else b + v if b > 0 else b for b, v in zip(cf_bottom, cf_vals)],
              color=cf_colors, edgecolor="white", width=0.5, alpha=0.85)
# Simpler approach: just use direct bar chart
plt.close()

fig, ax = plt.subplots(figsize=(9, 5.5))
cf_labels2 = ["Adj EBITDA", "CapEx", "FCF"]
cf_vals2   = [1.13, -10.1, -9.1]
colors2    = [GREEN, RED, RED]
bars = ax.bar(cf_labels2, cf_vals2, color=colors2, edgecolor="white", width=0.45, alpha=0.85)
for b, v in zip(bars, cf_vals2):
    ypos = v + 0.15 if v >= 0 else v - 0.35
    ax.text(b.get_x() + b.get_width()/2, ypos, f"${v:.1f}B", ha="center",
            fontsize=11, fontweight="bold", color="white" if v < 0 else "black")
ax.axhline(0, color="black", linewidth=0.8)
ax.set_ylabel("$ Billion", fontsize=11)
ax.set_title("Q1 2026 Cash Generation Overview", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(-12, 3)
ax.annotate("$10.1B CapEx driven by\nAI infrastructure ($7.7B)",
            xy=(1, -10.1), xytext=(50, 30), textcoords="offset points",
            fontsize=9, color=DARK_BLUE, fontstyle="italic",
            arrowprops=dict(arrowstyle="->", color=DARK_BLUE, lw=1.2))
fig.savefig(os.path.join(OUT, "spcx_chart8_cash_flow.png"))
plt.close()

# ── Chart 9: Adjusted EBITDA by Segment (FY2025) ──────────────────────────
fig, ax = plt.subplots(figsize=(9, 5.5))
ebitda_colors = [GREEN if v >= 0 else RED for v in seg_ebitda_fy25]
bars = ax.bar(seg_labels, seg_ebitda_fy25, color=ebitda_colors, edgecolor="white", width=0.5, alpha=0.85)
for b, v in zip(bars, seg_ebitda_fy25):
    ypos = v + 0.1 if v >= 0 else v - 0.3
    ax.text(b.get_x() + b.get_width()/2, ypos, f"${v:.2f}B", ha="center",
            fontsize=10, fontweight="bold")
ax.axhline(0, color="black", linewidth=0.8)
# Add total line
ax.axhline(y=6.58, color=SPACEX_BLUE, linestyle="--", linewidth=1.5, alpha=0.6)
ax.text(2.35, 6.58 + 0.15, "Consolidated: $6.6B", fontsize=9, color=SPACEX_BLUE, fontweight="bold")
ax.set_ylabel("Adjusted EBITDA ($ Billion)", fontsize=11)
ax.set_title("FY2025 Adjusted EBITDA by Segment", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(-3, 8)
fig.savefig(os.path.join(OUT, "spcx_chart9_ebitda_segment.png"))
plt.close()

# ── Chart 10: Post-IPO Valuation Context ──────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
comp_names = ["SpaceX\n(SPCX)", "Palantir\n(PLTR)", "Amazon\n(AMZN)", "Tesla\n(TSLA)", "Meta\n(META)"]
comp_ps    = [94, 55, 4.2, 12, 9.5]  # Price/Sales multiples (approximate)
comp_colors = [AI_PURPLE, LIGHT_BLUE, GRAY, GRAY, GRAY]
bars = ax.bar(comp_names, comp_ps, color=comp_colors, edgecolor="white", width=0.5)
for b, v in zip(bars, comp_ps):
    ax.text(b.get_x() + b.get_width()/2, v + 1, f"{v:.1f}x", ha="center",
            fontsize=10, fontweight="bold")
ax.set_ylabel("Price / Sales (x)", fontsize=11)
ax.set_title("IPO Valuation: SpaceX P/S Multiple vs. Peers", fontsize=13, fontweight="bold", pad=12)
ax.set_ylim(0, 110)
ax.annotate("SpaceX trades at extreme premium\nrelative to large-cap tech peers",
            xy=(0, 94), xytext=(80, -15), textcoords="offset points",
            fontsize=9, color=RED, fontstyle="italic",
            arrowprops=dict(arrowstyle="->", color=RED, lw=1.2))
fig.savefig(os.path.join(OUT, "spcx_chart10_valuation_comps.png"))
plt.close()

print(f"All 10 charts saved to {OUT}")
