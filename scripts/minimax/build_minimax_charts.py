"""
MiniMax Group (0100.HK) FY2025 Annual Results — Chart Generator
Generates 10 institutional-quality charts for the earnings update report
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np

OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/MINIMAX/"

NAVY   = "#003366"
BLUE   = "#0066CC"
RED    = "#CC0000"
GRAY   = "#5E6A71"
LGRAY  = "#D3D3D3"
GREEN  = "#1A7A4A"
AMBER  = "#E87722"
TEAL   = "#00838F"
PURPLE = "#6A1B9A"

plt.rcParams.update({
    "font.family":        "sans-serif",
    "font.sans-serif":    ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,
    "figure.dpi":         150,
})

def style_ax(ax, title, ylabel="", xlabel=""):
    ax.set_title(title, fontsize=11, fontweight='bold', color=NAVY, pad=8)
    if ylabel: ax.set_ylabel(ylabel, fontsize=8, color=GRAY)
    if xlabel: ax.set_xlabel(xlabel, fontsize=8, color=GRAY)
    ax.tick_params(colors=GRAY, labelsize=8)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(LGRAY)
    ax.spines['bottom'].set_color(LGRAY)
    ax.yaxis.grid(True, color=LGRAY, linewidth=0.5, linestyle='--')
    ax.set_axisbelow(True)

# ─────────────────────────────────────────────────
# Chart 1: Annual Revenue Progression
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
years = ['FY2023', 'FY2024', 'FY2025']
revenue = [7.2, 30.5, 79.0]
bars = ax.bar(years, revenue, color=[LGRAY, NAVY, RED], width=0.55, zorder=3)
for bar, val in zip(bars, revenue):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.2,
            f'${val:.1f}M', ha='center', va='bottom', fontsize=9, fontweight='bold', color=NAVY)
ax.annotate('+158.9% YoY', xy=(2, 79.0), xytext=(2, 92),
            fontsize=8, color=RED, ha='center', fontweight='bold')
style_ax(ax, 'MiniMax 年度收入增长趋势', '收入 ($M)')
ax.set_ylim(0, 105)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart1_revenue.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 1 saved — Annual Revenue")

# ─────────────────────────────────────────────────
# Chart 2: Revenue by Segment (AI-native vs Open Platform)
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
years2 = ['FY2024', 'FY2025']
ai_native = [21.8, 53.1]
open_platform = [8.7, 26.0]
x = np.arange(len(years2))
w = 0.35
b1 = ax.bar(x - w/2, ai_native, w, color=NAVY, label='AI原生产品', zorder=3)
b2 = ax.bar(x + w/2, open_platform, w, color=TEAL, label='开放平台及企业服务', zorder=3)
for bar, val in zip(b1, ai_native):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
            f'${val:.1f}M', ha='center', va='bottom', fontsize=8, fontweight='bold', color=NAVY)
for bar, val in zip(b2, open_platform):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
            f'${val:.1f}M', ha='center', va='bottom', fontsize=8, fontweight='bold', color=TEAL)
ax.set_xticks(x); ax.set_xticklabels(years2)
ax.legend(fontsize=8, loc='upper left')
style_ax(ax, 'MiniMax 收入结构：按业务板块', '收入 ($M)')
ax.set_ylim(0, 70)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart2_segment.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 2 saved — Revenue by Segment")

# ─────────────────────────────────────────────────
# Chart 3: Gross Margin Trend
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
years3 = ['FY2023', 'FY2024', 'FY2025']
gm = [-15.0, 12.2, 25.4]
colors3 = [RED, AMBER, GREEN]
bars3 = ax.bar(years3, gm, color=colors3, width=0.55, zorder=3)
for bar, val in zip(bars3, gm):
    offset = 1.0 if val >= 0 else -2.5
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + offset,
            f'{val:.1f}%', ha='center', va='bottom', fontsize=9, fontweight='bold',
            color=GREEN if val > 0 else RED)
ax.axhline(0, color=GRAY, linewidth=0.8)
ax.annotate('+13.2pp YoY', xy=(2, 25.4), xytext=(2, 32),
            fontsize=8, color=GREEN, ha='center', fontweight='bold')
style_ax(ax, 'MiniMax 毛利率趋势', '毛利率 (%)')
ax.set_ylim(-25, 40)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart3_margin.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 3 saved — Gross Margin Trend")

# ─────────────────────────────────────────────────
# Chart 4: Operating Expense Breakdown
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
cats = ['研发', '销售及分销', '行政']
fy24 = [189.0, 87.0, 14.4]
fy25 = [252.8, 51.9, 36.8]
x4 = np.arange(len(cats))
w4 = 0.35
b4a = ax.bar(x4 - w4/2, fy24, w4, color=LGRAY, label='FY2024', zorder=3)
b4b = ax.bar(x4 + w4/2, fy25, w4, color=NAVY, label='FY2025', zorder=3)
for bar, val in zip(b4a, fy24):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 3,
            f'${val:.0f}M', ha='center', va='bottom', fontsize=8, color=GRAY)
for bar, val in zip(b4b, fy25):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 3,
            f'${val:.1f}M', ha='center', va='bottom', fontsize=8, fontweight='bold', color=NAVY)
ax.set_xticks(x4); ax.set_xticklabels(cats)
ax.legend(fontsize=8)
style_ax(ax, 'MiniMax 运营费用结构', '金额 ($M)')
ax.set_ylim(0, 310)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart4_opex.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 4 saved — OpEx Breakdown")

# ─────────────────────────────────────────────────
# Chart 5: Adjusted Net Loss Trend
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
years5 = ['FY2023', 'FY2024', 'FY2025']
adj_loss = [-128.0, -244.2, -250.9]
bars5 = ax.bar(years5, adj_loss, color=[AMBER, RED, RED], width=0.55, zorder=3)
for bar, val in zip(bars5, adj_loss):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() - 8,
            f'${val:.1f}M', ha='center', va='top', fontsize=9, fontweight='bold', color=RED)
ax.axhline(0, color=GRAY, linewidth=0.8)
style_ax(ax, 'MiniMax 调整后净亏损趋势', '调整后净亏损 ($M)')
ax.set_ylim(-310, 20)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart5_adj_loss.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 5 saved — Adjusted Net Loss")

# ─────────────────────────────────────────────────
# Chart 6: Geographic Revenue Split
# ─────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(8, 4))
for i, (yr, dom, intl) in enumerate([('FY2024', 30.2, 69.8), ('FY2025', 27.0, 73.0)]):
    ax = axes[i]
    sizes = [intl, dom]
    colors6 = [NAVY, LGRAY]
    wedges, texts, autotexts = ax.pie(sizes, labels=['海外', '国内'], colors=colors6,
                                       autopct='%1.1f%%', startangle=90, textprops={'fontsize': 9})
    for t in autotexts:
        t.set_fontweight('bold')
    ax.set_title(yr, fontsize=11, fontweight='bold', color=NAVY, pad=8)
fig.suptitle('MiniMax 收入地域分布', fontsize=12, fontweight='bold', color=NAVY, y=1.02)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart6_geo.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 6 saved — Geographic Revenue Split")

# ─────────────────────────────────────────────────
# Chart 7: Cumulative User Growth
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
periods = ['2023', '9M2025\n(招股书)', 'FY2025']
users = [3.1, 212.0, 236.0]
bars7 = ax.bar(periods, users, color=[LGRAY, NAVY, RED], width=0.55, zorder=3)
for bar, val in zip(bars7, users):
    label = f'{val:.0f}M' if val >= 10 else f'{val:.1f}M'
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 4,
            label, ha='center', va='bottom', fontsize=9, fontweight='bold', color=NAVY)
style_ax(ax, 'MiniMax 累计用户数', '累计用户 (百万)')
ax.set_ylim(0, 280)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart7_users.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 7 saved — Cumulative Users")

# ─────────────────────────────────────────────────
# Chart 8: R&D Expense vs Revenue
# ─────────────────────────────────────────────────
fig, ax1 = plt.subplots(figsize=(7, 4))
years8 = ['FY2023', 'FY2024', 'FY2025']
rd = [110.0, 189.0, 252.8]
rev = [7.2, 30.5, 79.0]
x8 = np.arange(len(years8))
ax1.bar(x8, rd, 0.5, color=NAVY, alpha=0.8, label='研发支出', zorder=3)
for i, val in enumerate(rd):
    ax1.text(i, val + 4, f'${val:.0f}M', ha='center', fontsize=8, fontweight='bold', color=NAVY)
ax2 = ax1.twinx()
ax2.plot(x8, rev, color=RED, marker='o', linewidth=2.5, markersize=8, label='收入', zorder=4)
for i, val in enumerate(rev):
    ax2.text(i + 0.1, val + 5, f'${val:.1f}M', fontsize=8, fontweight='bold', color=RED)
ax1.set_xticks(x8); ax1.set_xticklabels(years8)
style_ax(ax1, 'MiniMax 研发投入 vs 收入', '研发支出 ($M)')
ax2.set_ylabel('收入 ($M)', fontsize=8, color=RED)
ax2.tick_params(colors=RED, labelsize=8)
ax2.spines['right'].set_color(RED)
lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=8, loc='upper left')
ax1.set_ylim(0, 320); ax2.set_ylim(0, 100)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart8_rd_vs_rev.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 8 saved — R&D vs Revenue")

# ─────────────────────────────────────────────────
# Chart 9: Cash Position
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
periods9 = ['FY2024末', 'FY2025末', 'ARR\n(2026.2月)']
cash = [880.6, 1050.3, None]
arr_val = 150.0
bars9_cash = ax.bar(['FY2024末', 'FY2025末'], [880.6, 1050.3], color=[LGRAY, NAVY], width=0.45, zorder=3)
for bar, val in zip(bars9_cash, [880.6, 1050.3]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 15,
            f'${val:.1f}M', ha='center', va='bottom', fontsize=9, fontweight='bold', color=NAVY)
ax.annotate('+$169.7M\n(+19.3%)', xy=(1, 1050.3), xytext=(1.4, 1050.3),
            fontsize=8, color=GREEN, fontweight='bold', ha='center',
            arrowprops=dict(arrowstyle='->', color=GREEN, lw=1.2))
style_ax(ax, 'MiniMax 现金储备（含受限现金及金融资产）', '金额 ($M)')
ax.set_ylim(0, 1250)
fig.tight_layout()
fig.savefig(OUT + "minimax_chart9_cash.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 9 saved — Cash Position")

# ─────────────────────────────────────────────────
# Chart 10: Analyst Target Prices
# ─────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 4))
analysts = ['高盛', '中金', '国泰海通', '东吴证券', '一致预期']
targets = [1018, 1109, 1012, 1100, 1114]
current = 768.5
bars10 = ax.barh(analysts, targets, color=[BLUE, NAVY, TEAL, PURPLE, RED], height=0.5, zorder=3)
ax.axvline(current, color=AMBER, linewidth=2, linestyle='--', label=f'当前价 HK${current:.0f}', zorder=4)
for bar, val in zip(bars10, targets):
    upside = (val / current - 1) * 100
    ax.text(val + 10, bar.get_y() + bar.get_height()/2,
            f'HK${val:,.0f} (+{upside:.0f}%)', va='center', fontsize=8, fontweight='bold', color=NAVY)
ax.legend(fontsize=8, loc='lower right')
style_ax(ax, 'MiniMax 分析师目标价 vs 当前股价', '目标价 (HK$)')
ax.set_xlim(0, 1350)
ax.invert_yaxis()
fig.tight_layout()
fig.savefig(OUT + "minimax_chart10_targets.png", dpi=150, bbox_inches='tight')
plt.close()
print("Chart 10 saved — Analyst Targets")

print("\n✅ All 10 charts generated successfully.")
