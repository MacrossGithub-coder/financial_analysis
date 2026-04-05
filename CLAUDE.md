# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## 目录结构规则

```
financial_analysis/
├── CLAUDE.md                        # 项目规则（本文件）
├── package.json                     # Node.js 依赖声明（pptxgenjs）
├── package-lock.json                # 锁文件
├── node_modules/                    # Node.js 依赖包（勿移动，.mjs 脚本依赖根目录路径）
│
├── scripts/                         # 所有构建脚本，按标的代码分子目录
│   ├── aapl/                        # Apple 相关脚本
│   │   ├── build_aapl_charts.py     # 生成图表 PNG
│   │   ├── build_aapl_report.py     # 生成英文 DOCX 报告
│   │   └── build_aapl_report_cn.py  # 生成中文 DOCX 报告
│   ├── ibrx/                        # ImmunityBio 相关脚本
│   │   ├── build_ibrx_charts.py
│   │   ├── build_ibrx_report_en.py
│   │   └── build_ibrx_report_cn.py
│   ├── smic/                        # 中芯国际相关脚本
│   │   ├── build_smic_charts.py
│   │   ├── build_smic_report_en.py
│   │   ├── build_smic_report_cn.py
│   │   └── build_smic_onepager.mjs  # One-pager PPTX（Node.js/PptxGenJS）
│   ├── uuuu/                        # Energy Fuels 相关脚本
│   │   ├── build_uuuu_charts.py
│   │   ├── build_uuuu_report.py
│   │   └── build_uuuu_report_cn.py
│   └── utils/                       # 通用工具脚本
│       ├── build_comps.py           # 可比公司分析（Comps）构建器
│       └── preview_pptx.py          # PPTX 预览渲染工具（python-pptx + PIL）
│
└── output/                          # 所有生成文件，按标的代码分子目录
    ├── AAPL/                        # Apple 输出
    │   ├── AAPL_Q4_CY2025_Earnings_Update.docx
    │   ├── AAPL_Q4_CY2025_业绩更新报告_中文版.docx
    │   └── chart*.png               # 图表文件（chart1_revenue.png 等）
    ├── IBRX/                        # ImmunityBio 输出
    │   ├── IBRX_Q4_FY2025_Earnings_Update.docx
    │   ├── IBRX_Q4_FY2025_业绩更新报告_中文版.docx
    │   └── ibrx_chart*.png
    ├── SMIC/                        # 中芯国际输出
    │   ├── SMIC_Q4_2025_Earnings_Update.docx
    │   ├── SMIC_Q4_2025_业绩更新报告_中文版.docx
    │   ├── SMIC_OnePager_StripProfile.pptx
    │   ├── SMIC_Foundry_Comps_FY2024.xlsx
    │   ├── smic_onepager_preview.png
    │   └── smic_chart*.png
    └── UUUU/                        # Energy Fuels 输出
        ├── UUUU_Q4_FY2025_Earnings_Update.docx
        ├── UUUU_Q4_FY2025_业绩更新报告_中文版.docx
        └── uuuu_chart*.png
```

### 新增标的时的规则

1. **脚本目录**：在 `scripts/{TICKER}/` 下创建（小写标的代码）
2. **输出目录**：在 `output/{TICKER}/` 下保存（大写标的代码）
3. **脚本路径变量**：`OUT` 或 `BASE` 必须指向 `output/{TICKER}/` 的绝对路径，不得使用 `__file__` 相对路径或根目录
4. **Node.js 脚本**：`.mjs` 中 `require()` 路径必须使用 `node_modules/` 的绝对路径

### 脚本命名规范

| 用途 | 命名格式 | 示例 |
|------|---------|------|
| 图表生成 | `build_{ticker}_charts.py` | `build_aapl_charts.py` |
| 英文报告 | `build_{ticker}_report.py` 或 `_report_en.py` | `build_smic_report_en.py` |
| 中文报告 | `build_{ticker}_report_cn.py` | `build_uuuu_report_cn.py` |
| PPT one-pager | `build_{ticker}_onepager.mjs` | `build_smic_onepager.mjs` |

### 输出文件命名规范

| 类型 | 命名格式 | 示例 |
|------|---------|------|
| 英文报告 | `{TICKER}_{Quarter}_{Type}.docx` | `IBRX_Q4_FY2025_Earnings_Update.docx` |
| 中文报告 | `{TICKER}_{Quarter}_{中文类型}_中文版.docx` | `SMIC_Q4_2025_业绩更新报告_中文版.docx` |
| PPT | `{TICKER}_{Type}.pptx` | `SMIC_OnePager_StripProfile.pptx` |
| 图表 | `{ticker}_chart{N}_{description}.png` | `ibrx_chart1_revenue.png` |

---

## 报告生成规则

**每次生成报告必须同时输出两个文件：**
1. **原始语言版本**（英文或用户指定语言）
2. **中文版本** — 文件名以 `_中文版` 结尾，内容完整翻译，不得仅做摘要

中文版要求：
- 正文字体使用**宋体**（`eastAsia` font tag），标题使用**黑体**
- 所有章节、表格、图表说明、资料来源均须翻译为中文
- 数据与英文版保持一致，图表可共用（无需重新生成）
- 与英文版同步输出，不得事后补充

---

## 技术栈说明

| 工具 | 用途 |
|------|------|
| `python-docx` | DOCX 报告生成（英文 + 中文） |
| `matplotlib` | 图表生成（PNG，150 dpi） |
| `PptxGenJS` (Node.js) | PPTX one-pager 生成 |
| `python-pptx` + `Pillow` | PPTX 预览渲染 |
| `openpyxl` | Excel 模型（Comps 等） |
| `yfinance` | 实时/历史股价及市场数据获取 |

---

## matplotlib 中文字体规则

**图表脚本（`build_*_charts.py`）中含有中文标题、坐标轴标签或图例时，必须使用以下字体配置，否则中文字符将显示为乱码（方块）。**

`Times New Roman` 不含 CJK 字符，**禁止**单独用作 `font.family`。

### 标准字体配置（所有含中文的图表脚本必须使用）

```python
plt.rcParams.update({
    # Arial Unicode MS 同时支持拉丁字符与 CJK 字符，避免中文乱码
    "font.family":        "sans-serif",
    "font.sans-serif":    ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB", "Times New Roman"],
    "axes.unicode_minus": False,   # 避免负号显示为方块
    # ... 其余参数
})
```

### 本机可用 CJK 字体（按优先级排列）

| 字体名 | 路径 | 说明 |
|--------|------|------|
| `Arial Unicode MS` | `/System/Library/Fonts/Supplemental/Arial Unicode.ttf` | **首选**，同时支持拉丁 + CJK |
| `STHeiti` | `/System/Library/AssetsV2/…/STHEITI.ttf` | 无衬线中文黑体，图表效果佳 |
| `Hiragino Sans GB` | `/System/Library/Fonts/Hiragino Sans GB.ttc` | 苹果日文/中文备选 |
| `STFangsong` | `/System/Library/AssetsV2/…/STFANGSO.ttf` | 仿宋，备用 |

### 图表脚本的中文支持规则

1. 所有含中文文字的图表脚本，`rcParams` 必须使用上方标准配置
2. 若脚本仅含英文，可继续使用 `"font.family": "Times New Roman"`
3. 新建图表脚本前，可通过以下代码验证字体可用性：

```python
import matplotlib.font_manager as fm
for name in ["Arial Unicode MS", "STHeiti", "Hiragino Sans GB"]:
    try:
        fm.findfont(fm.FontProperties(family=name), fallback_to_default=False)
        print(f"OK: {name}")
    except:
        print(f"MISSING: {name}")
```

---

## 股价数据获取规则

**报告中涉及当前股价、市值、52 周区间等市场数据时，必须通过 `yfinance` 动态获取，禁止手动填写或估算。**

在每个报告脚本的顶部加入以下标准获取逻辑：

```python
import yfinance as yf

def get_market_data(ticker: str) -> dict:
    """从 yfinance 获取实时市场数据"""
    t = yf.Ticker(ticker)
    info = t.fast_info
    return {
        "price":      round(info.last_price, 2),
        "market_cap": info.market_cap,
        "52w_high":   round(info.year_high, 2),
        "52w_low":    round(info.year_low, 2),
    }

mkt = get_market_data("RKLB")  # 替换为对应标的代码
```

- 使用 `fast_info`（轻量接口，速度快）优先于 `info`
- 若网络不可用或获取失败，脚本应打印警告并以 `"N/A"` 占位，不得静默使用估算值
- 市值单位统一换算为「亿美元」后写入报告（`market_cap / 1e9`）
