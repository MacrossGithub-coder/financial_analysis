"""
build_002475_charts.py
生成立讯精密 (002475.SZ) Q4 2025 业绩更新报告所需的 10 张图表
输出目录: output/002475/
"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

# ── 输出目录 ────────────────────────────────────────────────────────────────
OUT = "/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis/output/002475/"
os.makedirs(OUT, exist_ok=True)

# ── 品牌色彩 ─────────────────────────────────────────────────────────────────
BLUE   = "#1A3F7A"   # Luxshare navy blue (primary)
TEAL   = "#0E8A7D"   # teal (secondary / AI & auto)
GOLD   = "#C8960C"   # amber gold (accent)
RED    = "#C41230"   # red (miss / negative)
LGRAY  = "#E8EEF6"   # light blue-gray background
MGRAY  = "#8B8B8B"   # mid gray
WHITE  = "#FFFFFF"
BLUE2  = "#2E70C5"   # medium blue
GREEN  = "#2E7D32"   # green (beat)

# ── 字体 ──────────────────────────────────────────────────────────────────────
plt.rcParams.update({
    "font.family":      "sans-serif",
    "font.sans-serif":  ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,
    "axes.titlesize":   13,
    "axes.labelsize":   10,
    "xtick.labelsize":  9,
    "ytick.labelsize":  9,
    "legend.fontsize":  9,
    "figure.facecolor": WHITE,
    "axes.facecolor":   WHITE,
    "axes.spines.top":  False,
    "axes.spines.right":False,
})

DPI = 150

# ── 核心数据 ──────────────────────────────────────────────────────────────────
QTRS = ["Q1\n2024", "Q2\n2024", "Q3\n2024", "Q4\n2024",
        "Q1\n2025", "Q2\n2025", "Q3\n2025", "Q4\n2025E"]

# 季度营收（亿元）— Q4 2025 为隐含推算值（官方年报未披露）
REV_DATA = [524.07, 511.91, 735.79, 916.18,
            617.88, 627.15, 964.11, 960.0]

# 季度归母净利润（亿元）— Q4 2025 为隐含中间值（官方预告中值）
NP_DATA  = [24.71, 29.25, 36.79, 42.91,
            30.44, 36.01, 48.74, 53.3]

# 2025各季度营收同比（%）
REV_YOY_2025 = [17.9, 22.5, 31.0, 4.8]   # Q1–Q4 2025

# 2025各季度归母净利同比（%）
NP_YOY_2025  = [23.2, 23.1, 32.5, 24.2]   # Q1–Q4 2025

# Q4 2025 是估算值（灰色填充，虚线轮廓）
N = len(QTRS)
ESTIMATED_IDX = N - 1   # 最后一根柱子为估算

# ── 颜色辅助函数 ──────────────────────────────────────────────────────────────
def bar_colors(base, n, estimated_idx):
    """返回 n 个柱子的颜色列表，estimated_idx 位置用浅色"""
    cols = [base] * n
    cols[estimated_idx] = MGRAY
    return cols

def save(fig, name):
    path = os.path.join(OUT, name)
    fig.savefig(path, dpi=DPI, bbox_inches="tight")
    plt.close(fig)
    print(f"  saved → {name}")


# ════════════════════════════════════════════════════════════════════════════
# Chart 1 — 季度营收走势
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(10, 5.5))
colors_rev = bar_colors(BLUE, N, ESTIMATED_IDX)
bars = ax.bar(QTRS, REV_DATA, color=colors_rev, width=0.6, edgecolor=WHITE, linewidth=0.8)
# Q4 2025E 虚线边框
bars[ESTIMATED_IDX].set_edgecolor(BLUE)
bars[ESTIMATED_IDX].set_linewidth(1.5)
bars[ESTIMATED_IDX].set_linestyle("--")

for bar, val in zip(bars, REV_DATA):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 15,
            f"{val:.0f}", ha="center", va="bottom", fontsize=8.5, color="#333333")

# 分隔线：2024 vs 2025
ax.axvline(x=3.5, color=MGRAY, linestyle="--", linewidth=0.8, alpha=0.6)
ax.text(1.5, max(REV_DATA) * 0.95, "FY2024", ha="center", fontsize=9,
        color=MGRAY, fontweight="bold")
ax.text(5.5, max(REV_DATA) * 0.95, "FY2025", ha="center", fontsize=9,
        color=BLUE, fontweight="bold")

ax.set_ylabel("营业收入（亿元人民币）")
ax.set_title("立讯精密 — 季度营业收入走势", fontweight="bold", pad=12)
ax.set_ylim(0, max(REV_DATA) * 1.18)

legend_patches = [
    mpatches.Patch(color=BLUE, label="实际值"),
    mpatches.Patch(color=MGRAY, label="Q4 2025 隐含推算（官方年报待披露）"),
]
ax.legend(handles=legend_patches, loc="upper left", framealpha=0.8)
ax.set_facecolor(LGRAY)
ax.grid(axis="y", color=WHITE, linewidth=0.8)
fig.text(0.5, -0.02, "来源：立讯精密2025三季报（2025-10-30）；Q4 2025为FY2025全年预告中值减前三季度",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart1_quarterly_revenue.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 2 — 季度归母净利润走势
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(10, 5.5))
colors_np = bar_colors(TEAL, N, ESTIMATED_IDX)
bars2 = ax.bar(QTRS, NP_DATA, color=colors_np, width=0.6, edgecolor=WHITE, linewidth=0.8)
bars2[ESTIMATED_IDX].set_edgecolor(TEAL)
bars2[ESTIMATED_IDX].set_linewidth(1.5)
bars2[ESTIMATED_IDX].set_linestyle("--")

for bar, val in zip(bars2, NP_DATA):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.6,
            f"{val:.1f}", ha="center", va="bottom", fontsize=8.5, color="#333333")

# YoY 标签仅 2025 年各季度
yoy_labels = NP_YOY_2025
q25_indices = [4, 5, 6, 7]
for i, yoy in zip(q25_indices, yoy_labels):
    color = GREEN if yoy >= 0 else RED
    suffix = "E" if i == ESTIMATED_IDX else ""
    ax.text(i, NP_DATA[i] / 2, f"+{yoy:.1f}%{suffix}\nYoY",
            ha="center", va="center", fontsize=8, color=WHITE, fontweight="bold")

ax.axvline(x=3.5, color=MGRAY, linestyle="--", linewidth=0.8, alpha=0.6)
ax.text(1.5, max(NP_DATA) * 0.92, "FY2024", ha="center", fontsize=9, color=MGRAY, fontweight="bold")
ax.text(5.5, max(NP_DATA) * 0.92, "FY2025", ha="center", fontsize=9, color=TEAL, fontweight="bold")

ax.set_ylabel("归母净利润（亿元人民币）")
ax.set_title("立讯精密 — 季度归母净利润走势", fontweight="bold", pad=12)
ax.set_ylim(0, max(NP_DATA) * 1.22)

legend_patches2 = [
    mpatches.Patch(color=TEAL, label="实际值"),
    mpatches.Patch(color=MGRAY, label="Q4 2025 隐含推算（官方年报待披露）"),
]
ax.legend(handles=legend_patches2, loc="upper left", framealpha=0.8)
ax.set_facecolor(LGRAY)
ax.grid(axis="y", color=WHITE, linewidth=0.8)
fig.text(0.5, -0.02, "来源：立讯精密季度报告；Q4 2025 基于FY2025全年业绩预告中值（165.18亿–171.86亿）",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart2_quarterly_netprofit.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 3 — 季度营收同比增速（2025全年）
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(8, 5))
QTRS_2025 = ["Q1 2025", "Q2 2025", "Q3 2025", "Q4 2025E"]
rev_colors = [BLUE, BLUE, BLUE, MGRAY]

bars3 = ax.bar(QTRS_2025, REV_YOY_2025, color=rev_colors, width=0.5,
               edgecolor=WHITE, linewidth=0.8)
bars3[-1].set_edgecolor(BLUE)
bars3[-1].set_linestyle("--")
bars3[-1].set_linewidth(1.5)

for bar, val in zip(bars3, REV_YOY_2025):
    ax.text(bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.5,
            f"+{val:.1f}%", ha="center", va="bottom", fontsize=10, fontweight="bold",
            color=BLUE if bar.get_facecolor()[:3] != (0.545, 0.545, 0.545) else MGRAY)

ax.axhline(y=0, color="#333333", linewidth=0.8)
ax.set_ylabel("营收同比增速（%）")
ax.set_title("立讯精密 — 2025年各季度营收同比增速", fontweight="bold", pad=12)
ax.set_ylim(0, max(REV_YOY_2025) * 1.35)
ax.set_facecolor(LGRAY)
ax.grid(axis="y", color=WHITE, linewidth=0.8)
fig.text(0.5, -0.02, "来源：立讯精密季度报告；Q4 2025E为分析师隐含推算值",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart3_revenue_growth.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 4 — 毛利率走势
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(9, 5))
GM_LABELS = ["FY2023", "H1 2024", "H2 2024", "FY2024", "H1 2025", "9M 2025"]
GM_VALUES  = [11.6,    11.71,    9.58,    10.40,    11.61,    12.15]
gm_colors  = [BLUE2, BLUE2, RED, BLUE, TEAL, TEAL]

bars4 = ax.bar(GM_LABELS, GM_VALUES, color=gm_colors, width=0.55,
               edgecolor=WHITE, linewidth=0.8)

for bar, val in zip(bars4, GM_VALUES):
    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.15,
            f"{val:.2f}%", ha="center", va="bottom", fontsize=9, fontweight="bold",
            color="#333333")

# 连接折线
x_pos = [bar.get_x() + bar.get_width() / 2 for bar in bars4]
ax.plot(x_pos, GM_VALUES, color=GOLD, linewidth=1.5, linestyle="--",
        marker="o", markersize=5, markerfacecolor=GOLD, zorder=5)

ax.set_ylabel("毛利率（%）")
ax.set_title("立讯精密 — 综合毛利率走势", fontweight="bold", pad=12)
ax.set_ylim(8, 13.5)
ax.set_facecolor(LGRAY)
ax.grid(axis="y", color=WHITE, linewidth=0.8)

note = "注：H2 2024毛利率受iPhone 16量产爬坡及汇率压力影响；9M 2025含并购负商誉收益~4.79亿元"
fig.text(0.5, -0.04, note + "\n来源：立讯精密年报、中报、三季报",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart4_margin_trend.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 5 — FY2024 分业务收入结构（饼图）
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(8, 6))
SEG_LABELS = ["消费性电子\n2,240.9亿", "通讯互联\n183.6亿",
              "汽车互联\n137.6亿",     "PC及其他\n125.9亿"]
SEG_VALS   = [2240.94, 183.60, 137.58, 125.83]
SEG_COLORS = [BLUE, TEAL, GOLD, BLUE2]
explode    = (0.04, 0.04, 0.04, 0.04)

wedges, texts, autotexts = ax.pie(
    SEG_VALS, labels=SEG_LABELS, colors=SEG_COLORS,
    autopct="%1.1f%%", startangle=140, explode=explode,
    textprops={"fontsize": 9}, pctdistance=0.78,
    wedgeprops={"edgecolor": WHITE, "linewidth": 2},
)
for at in autotexts:
    at.set_fontsize(9)
    at.set_fontweight("bold")
    at.set_color(WHITE)

ax.set_title("立讯精密 FY2024 — 分业务营收结构\n（总营收 2,687.95亿元，+15.9% YoY）",
             fontweight="bold", pad=15)
fig.text(0.5, -0.01,
         "来源：立讯精密2024年年度报告（2025-04-28）",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart5_segment_fy2024.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 6 — H1 2025 vs H1 2024 分业务收入对比
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(9, 5.5))
SEG3 = ["消费性电子", "通讯互联产品", "汽车互联产品"]
H1_2024 = [855.5, 74.66, 47.55]   # 由H1 2025 YoY逆推
H1_2025 = [977.99, 110.98, 86.58]
YOYS_H1  = [14.3, 48.7, 82.1]

x = np.arange(len(SEG3))
w = 0.35
bars_24 = ax.bar(x - w/2, H1_2024, w, label="H1 2024", color=LGRAY,
                 edgecolor=BLUE, linewidth=1.2)
bars_25 = ax.bar(x + w/2, H1_2025, w, label="H1 2025", color=[BLUE, TEAL, GOLD])

for bar, val in zip(bars_24, H1_2024):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 8,
            f"{val:.0f}", ha="center", va="bottom", fontsize=8.5, color="#333333")
for bar, val, yoy in zip(bars_25, H1_2025, YOYS_H1):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 8,
            f"{val:.0f}", ha="center", va="bottom", fontsize=8.5, color="#333333")
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height()/2,
            f"+{yoy:.0f}%\nYoY", ha="center", va="center", fontsize=8,
            color=WHITE, fontweight="bold")

ax.set_xticks(x)
ax.set_xticklabels(SEG3, fontsize=9)
ax.set_ylabel("营业收入（亿元人民币）")
ax.set_title("立讯精密 — H1 2025 vs H1 2024 分业务营收对比", fontweight="bold", pad=12)
ax.legend(loc="upper left", framealpha=0.8)
ax.set_facecolor(LGRAY)
ax.grid(axis="y", color=WHITE, linewidth=0.8)
ax.set_ylim(0, max(H1_2025) * 1.3)
fig.text(0.5, -0.02, "来源：立讯精密2025年中期报告（2025-08）",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart6_segment_h1_2025.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 7 — FY2025 业绩预告 vs 市场预期（Beat/Miss）
# ════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(8, 5))
METRICS   = ["FY2025 归母净利润\n（亿元）", "FY2025 EPS\n（元/股）"]
ACTUAL    = [168.5, 2.32]        # 预告中值
CONSENSUS = [164.3, 2.26]        # 预告前市场一致预期

x = np.arange(len(METRICS))
w = 0.32
b_cons = ax.bar(x - w/2, CONSENSUS, w, label="预告前市场一致预期", color=LGRAY,
                edgecolor=BLUE, linewidth=1.2)
b_act  = ax.bar(x + w/2, ACTUAL,    w, label="FY2025 业绩预告中值", color=[TEAL, TEAL])

for bar, val in zip(b_cons, CONSENSUS):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.5,
            f"{val:.1f}", ha="center", va="bottom", fontsize=9, color="#333333")
for bar, val in zip(b_act, ACTUAL):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1.5,
            f"{val:.2f}", ha="center", va="bottom", fontsize=9, fontweight="bold",
            color=TEAL)

# Beat 百分比标注
beat_pct = [(a - c) / c * 100 for a, c in zip(ACTUAL, CONSENSUS)]
for i, pct in enumerate(beat_pct):
    ax.annotate(f"+{pct:.1f}% 超预期", xy=(i + w/2, ACTUAL[i] + 5),
                xytext=(i + w/2 + 0.25, ACTUAL[i] + 10),
                fontsize=9, color=GREEN, fontweight="bold",
                arrowprops=dict(arrowstyle="->", color=GREEN, lw=1.2))

ax.set_xticks(x)
ax.set_xticklabels(METRICS, fontsize=9)
ax.set_ylabel("值（亿元 / 元）")
ax.set_title("立讯精密 FY2025 — 业绩预告 vs 市场一致预期", fontweight="bold", pad=12)
ax.legend(loc="upper left", framealpha=0.8)
ax.set_facecolor(LGRAY)
ax.grid(axis="y", color=WHITE, linewidth=0.8)
ax.set_ylim(0, max(ACTUAL) * 1.25)
fig.text(0.5, -0.02, "来源：立讯精密2025年度业绩预告（公告2025-136，2025-10-30）；一致预期来自Wind",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart7_beat_miss.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 8 — 年度营收 & 归母净利润对比（FY2023–FY2027E）
# ════════════════════════════════════════════════════════════════════════════
fig, ax1 = plt.subplots(figsize=(10, 5.5))
YEARS = ["FY2023A", "FY2024A", "FY2025E", "FY2026E", "FY2027E"]
ANN_REV = [2319.0, 2687.95, 3169.0, 4081.0, 4722.0]   # 亿元
ANN_NP  = [109.6,  133.66,  168.5,  218.3,  271.4]    # 亿元

x = np.arange(len(YEARS))
w = 0.42
ANN_COLORS = [BLUE, BLUE, TEAL, LGRAY, LGRAY]
BAR_EC     = [WHITE, WHITE, WHITE, TEAL, TEAL]

bars_r = ax1.bar(x - w/2, ANN_REV, w, label="营业收入（亿元，左轴）",
                 color=ANN_COLORS, edgecolor=BAR_EC, linewidth=1.2)
for i in [3, 4]:
    bars_r[i].set_linestyle("--")
    bars_r[i].set_linewidth(1.5)

ax2 = ax1.twinx()
ax2.plot(x, ANN_NP, color=GOLD, linewidth=2.0, marker="o",
         markersize=7, markerfacecolor=GOLD, label="归母净利润（亿元，右轴）")

for xi, val in zip(x, ANN_REV):
    ax1.text(xi - w/2, val + 30, f"{val:.0f}", ha="center", va="bottom",
             fontsize=8, color="#333333")
for xi, val in zip(x, ANN_NP):
    ax2.text(xi + 0.05, val + 3, f"{val:.1f}", ha="center", va="bottom",
             fontsize=8, color=GOLD, fontweight="bold")

ax1.set_ylabel("营业收入（亿元人民币）")
ax2.set_ylabel("归母净利润（亿元人民币）", color=GOLD)
ax2.tick_params(axis="y", labelcolor=GOLD)
ax1.set_xticks(x)
ax1.set_xticklabels(YEARS, fontsize=9)
ax1.set_title("立讯精密 — 年度营收与利润增长趋势（FY2023–FY2027E）",
              fontweight="bold", pad=12)
ax1.set_facecolor(LGRAY)
ax1.grid(axis="y", color=WHITE, linewidth=0.8)
ax1.set_ylim(0, max(ANN_REV) * 1.22)
ax2.set_ylim(0, max(ANN_NP) * 1.35)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", framealpha=0.8)

fig.text(0.5, -0.02,
         "注：FY2025E营收为分析师隐含推算值（非官方披露）；净利润采用业绩预告中值；FY2026E/FY2027E为Wind一致预期\n来源：Wind、同花顺；立讯精密历年年报",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart8_annual_comparison.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 9 — 券商 FY2026E 盈利预测 & 目标价
# ════════════════════════════════════════════════════════════════════════════
fig, ax1 = plt.subplots(figsize=(10, 5.5))
BROKERS   = ["中银证券", "华泰证券", "中泰证券", "浦银国际", "野村证券",\
             "高盛", "Wind\n一致预期"]
FY26_EPS  = [2.79, 3.03, 3.11, 3.04, None, None, 3.00]
TARGETS   = [65.0, 68.0, 74.0, 74.3, 84.3, 82.0, 68.2]
RATINGS   = ["买入", "买入", "买入", "买入", "买入", "买入", "买入"]

x = np.arange(len(BROKERS))
t_colors = [TEAL if t < 75 else BLUE for t in TARGETS]
t_colors[-1] = GOLD  # Wind 一致预期

bars9 = ax1.bar(x, TARGETS, color=t_colors, width=0.55,
                edgecolor=WHITE, linewidth=0.8)

ax2 = ax1.twinx()
eps_vals = [e if e is not None else 0 for e in FY26_EPS]
ax2.plot(x[:5], [e for e in FY26_EPS if e is not None],
         color=GOLD, linewidth=2, marker="D", markersize=7,
         markerfacecolor=GOLD, label="FY2026E EPS（元，右轴）", zorder=5)

for xi, t in enumerate(TARGETS):
    ax1.text(xi, t + 1.2, f"¥{t:.1f}", ha="center", va="bottom",
             fontsize=8.5, color="#333333", fontweight="bold")

for xi, e in enumerate(FY26_EPS):
    if e is not None:
        ax2.text(xi, e + 0.05, f"{e:.2f}元", ha="center", va="bottom",
                 fontsize=8, color=GOLD, fontweight="bold")

ax1.set_ylabel("目标价（人民币元）")
ax2.set_ylabel("FY2026E EPS（元）", color=GOLD)
ax2.tick_params(axis="y", labelcolor=GOLD)
ax1.set_xticks(x)
ax1.set_xticklabels(BROKERS, fontsize=8.5)
ax1.set_title("立讯精密 — 主要券商目标价 & FY2026E EPS 汇总",
              fontweight="bold", pad=12)
ax1.set_facecolor(LGRAY)
ax1.grid(axis="y", color=WHITE, linewidth=0.8)
ax1.set_ylim(0, max(TARGETS) * 1.22)
ax2.set_ylim(2.5, 3.5)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend([mpatches.Patch(color=TEAL, label="目标价 <¥75"),
            mpatches.Patch(color=BLUE, label="目标价 ≥¥75"),
            mpatches.Patch(color=GOLD, label="Wind 一致预期")] + lines2,
           ["目标价 <¥75", "目标价 ≥¥75", "Wind 一致预期"] + labels2,
           loc="upper left", framealpha=0.8, fontsize=8)

fig.text(0.5, -0.02,
         "来源：各券商研究报告（2025H2–2026Q1）；野村、高盛目标价不含EPS详情\n评级均为买入(Buy)/增持(Outperform)",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart9_broker_targets.png")


# ════════════════════════════════════════════════════════════════════════════
# Chart 10 — EPS 趋势与盈利增速预测（FY2023–FY2027E）
# ════════════════════════════════════════════════════════════════════════════
fig, ax1 = plt.subplots(figsize=(9, 5.5))
EPS_VALS  = [1.51, 1.84, 2.32, 3.00, 3.73]   # FY2023–FY2027E
EPS_GROWTH= [None, 21.9, 26.1, 29.3, 24.3]   # YoY growth %
YEARS5    = ["FY2023A", "FY2024A", "FY2025E", "FY2026E", "FY2027E"]

EPS_COLORS = [BLUE, BLUE, TEAL, LGRAY, LGRAY]
EPS_EC     = [WHITE, WHITE, WHITE, TEAL, TEAL]

bars10 = ax1.bar(np.arange(5), EPS_VALS, color=EPS_COLORS, width=0.5,
                 edgecolor=EPS_EC, linewidth=1.2)
for i in [3, 4]:
    bars10[i].set_linestyle("--")
    bars10[i].set_linewidth(1.5)

ax2 = ax1.twinx()
g_x   = [i for i, g in enumerate(EPS_GROWTH) if g is not None]
g_vals= [g for g in EPS_GROWTH if g is not None]
ax2.plot(g_x, g_vals, color=RED, linewidth=2, marker="o",
         markersize=7, markerfacecolor=RED,
         label="归母净利润增速（%，右轴）", zorder=5)
ax2.fill_between(g_x, 0, g_vals, color=RED, alpha=0.08)

for xi, val in enumerate(EPS_VALS):
    ax1.text(xi, val + 0.04, f"{val:.2f}元", ha="center", va="bottom",
             fontsize=9, fontweight="bold", color="#333333")

for xi, g in zip(g_x, g_vals):
    ax2.text(xi, g + 0.5, f"+{g:.1f}%", ha="center", va="bottom",
             fontsize=8, color=RED, fontweight="bold")

ax1.set_ylabel("EPS（元人民币）")
ax2.set_ylabel("归母净利润同比增速（%）", color=RED)
ax2.tick_params(axis="y", labelcolor=RED)
ax1.set_xticks(np.arange(5))
ax1.set_xticklabels(YEARS5, fontsize=9)
ax1.set_title("立讯精密 — EPS 趋势与盈利增速（FY2023–FY2027E）",
              fontweight="bold", pad=12)
ax1.set_facecolor(LGRAY)
ax1.grid(axis="y", color=WHITE, linewidth=0.8)
ax1.set_ylim(0, max(EPS_VALS) * 1.25)
ax2.set_ylim(0, max(g_vals) * 1.5)

lines1, labels1 = ax1.get_legend_handles_labels()
lines2, labels2 = ax2.get_legend_handles_labels()
ax1.legend([mpatches.Patch(color=BLUE, label="EPS 实际值"),
            mpatches.Patch(color=TEAL, label="EPS FY2025E"),
            mpatches.Patch(color=LGRAY, label="EPS 分析师预测")] + lines2,
           ["EPS 实际值", "EPS FY2025E", "EPS 分析师预测"] + labels2,
           loc="upper left", framealpha=0.8, fontsize=8)

fig.text(0.5, -0.02,
         "注：FY2025E EPS基于业绩预告中值（归母净利润168.5亿/总股本72.7亿股）；FY2026/2027E为Wind一致预期\n来源：Wind、同花顺、立讯精密公告",
         ha="center", fontsize=7.5, color=MGRAY)
save(fig, "lxs_chart10_eps_forecast.png")

print("\n✅ 全部 10 张图表生成完毕")
print(f"   输出目录: {OUT}")
