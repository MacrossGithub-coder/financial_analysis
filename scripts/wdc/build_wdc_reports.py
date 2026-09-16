#!/usr/bin/env python3
"""Build English and complete Chinese WDC Q4 FY2026 earnings updates."""

from __future__ import annotations

import sys
from pathlib import Path

import yfinance as yf
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "WDC"
OUT.mkdir(parents=True, exist_ok=True)

DOC_SCRIPTS = Path(
    "/Users/macrossz/.codex/plugins/cache/openai-primary-runtime/documents/"
    "26.905.11957/skills/documents/scripts"
)
sys.path.insert(0, str(DOC_SCRIPTS))
from table_geometry import apply_table_geometry, column_widths_from_weights  # noqa: E402


PAGE_WIDTH_DXA = 10166
BLACK = "000000"
NAVY = "17365D"
BLUE = "2F5597"
GREEN = "416900"
RED = "A61B1B"
ORANGE = "C65D17"
GREY = "667085"
LIGHT = "F2F6FA"
WHITE = "FFFFFF"
BORDER = "D9D9D9"

URLS = {
    "release": "https://investor.wdc.com/static-files/a6470e24-6dec-4fae-9b35-69777f704e3a",
    "deck": "https://investor.wdc.com/static-files/e1f02f77-4432-42c3-8bd4-024371f40e54",
    "event": "https://investor.wdc.com/events/event-details/western-digital-fourth-quarter-fiscal-2026-earnings-call",
    "transcript": "https://earningscalls.dev/transcripts/western-digital-corporation_wdc_earnings_call_transcript_2026-08-05",
    "tenk": "https://www.sec.gov/Archives/edgar/data/106040/000162828026057139/wdc-20260703.htm",
    "tenk_index": "https://www.sec.gov/Archives/edgar/data/106040/000162828026057139/0001628280-26-057139-index.html",
    "eightk": "https://www.sec.gov/Archives/edgar/data/106040/000162828026053305/0001628280-26-053305-index.htm",
    "q3_release": "https://www.westerndigital.com/company/newsroom/press-releases/2026/2026-04-30-wd-reports-fiscal-third-quarter-2026-financial-results",
    "q3_deck": "https://investor.wdc.com/static-files/5b2d41c1-7d45-4575-b9ea-c51424dbffeb",
    "q2_release": "https://investor.wdc.com/news-releases/news-release-details/western-digital-reports-fiscal-second-quarter-2026-financial",
    "q1_release": "https://investor.wdc.com/news-releases/news-release-details/western-digital-reports-fiscal-first-quarter-2026-financial",
    "q4_25_release": "https://investor.wdc.com/node/27046",
    "consensus": "https://www.marketbeat.com/earnings/reports/2026-8-5-western-digital-co-stock/",
    "quote": "https://longbridge.com/en/quote/WDC.US",
    "yahoo": "https://finance.yahoo.com/quote/WDC/",
    "ir": "https://investor.wdc.com/",
}

CHARTS = {
    i: OUT / name
    for i, name in enumerate(
        [
            "wdc_chart1_quarterly_revenue.png",
            "wdc_chart2_eps.png",
            "wdc_chart3_margins.png",
            "wdc_chart4_exabytes.png",
            "wdc_chart5_end_market_mix.png",
            "wdc_chart6_beat_miss.png",
            "wdc_chart7_fcf.png",
            "wdc_chart8_growth_drivers.png",
            "wdc_chart9_guidance.png",
            "wdc_chart10_estimate_revisions.png",
            "wdc_chart11_price_history.png",
        ],
        1,
    )
}


def market_data():
    result = {"price": "N/A", "market_cap": "N/A", "high": "N/A", "low": "N/A"}
    try:
        info = yf.Ticker("WDC").fast_info
        result = {
            "price": round(float(info.last_price), 2),
            "market_cap": float(info.market_cap),
            "high": round(float(info.year_high), 2),
            "low": round(float(info.year_low), 2),
        }
    except Exception as exc:
        print(f"WARNING: yfinance market-data retrieval failed for WDC: {exc}")
    return result


def set_font(run, lang, size=9.2, bold=False, italic=False, color=BLACK, heading=False):
    if lang == "cn":
        family = "Heiti SC" if heading else "Songti SC"
        fallback = "Arial Unicode MS"
    else:
        family = "Times New Roman"
        fallback = family
    run.font.name = family
    rpr = run._element.get_or_add_rPr()
    # LibreOffice's DOCX renderer resolves CJK glyphs through the Latin font
    # slots. Keep Word's East Asian family contract while providing a
    # Unicode-capable render fallback for visual QA.
    rpr.rFonts.set(qn("w:ascii"), fallback)
    rpr.rFonts.set(qn("w:hAnsi"), fallback)
    rpr.rFonts.set(qn("w:eastAsia"), family)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def paragraph(doc, text, lang, size=9.2, bold=False, italic=False, color=BLACK,
              before=0, after=4.2, align=None, keep=False, heading=False, style=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    set_font(r, lang, size, bold, italic, color, heading)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.keep_with_next = keep
    if align is not None:
        p.alignment = align
    return p


def heading(doc, text, lang, level=1):
    p = doc.add_paragraph(text, style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(1.5)
    p.paragraph_format.space_after = Pt(4.5)
    p.paragraph_format.keep_with_next = True
    for r in p.runs:
        set_font(r, lang, 13.2 if level == 1 else 10.8, True, False, BLACK, True)
    return p


def bullet(doc, title, body, lang, compact=False):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.03)
    p.paragraph_format.first_line_indent = Inches(-0.03)
    p.paragraph_format.space_after = Pt(3.3 if compact else 4.2)
    p.paragraph_format.line_spacing = 1.03
    r = p.add_run("■ ")
    set_font(r, lang, 8.8 if compact else 9.2, True, color=NAVY, heading=True)
    r = p.add_run(title + " ")
    set_font(r, lang, 8.8 if compact else 9.2, True, color=BLACK, heading=True)
    r = p.add_run(body)
    set_font(r, lang, 8.8 if compact else 9.2)
    return p


def add_hyperlink(p, label, url, lang, size=7.4):
    rel_id = p.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    family = "Songti SC" if lang == "cn" else "Times New Roman"
    fallback = "Arial Unicode MS" if lang == "cn" else family
    fonts.set(qn("w:ascii"), fallback)
    fonts.set(qn("w:hAnsi"), fallback)
    fonts.set(qn("w:eastAsia"), family)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(round(size * 2)))
    for item in (fonts, color, underline, sz):
        rpr.append(item)
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    p._p.append(link)


def source(doc, items, lang, lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0.5)
    p.paragraph_format.space_after = Pt(3.2)
    p.paragraph_format.line_spacing = 1.0
    prefix = lead or ("Sources: " if lang == "en" else "资料来源：")
    r = p.add_run(prefix)
    set_font(r, lang, 7.4, italic=True, color=GREY)
    for i, (label, url) in enumerate(items):
        if i:
            r = p.add_run("; ")
            set_font(r, lang, 7.4, italic=True, color=GREY)
        if url:
            add_hyperlink(p, label, url, lang, 7.4)
        else:
            r = p.add_run(label)
            set_font(r, lang, 7.4, italic=True, color=GREY)
    return p


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = OxmlElement(f"w:{edge}")
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), "4")
        tag.set(qn("w:color"), BORDER)
        borders.append(tag)


def cell_margins(cell, top=75, start=95, bottom=75, end=95):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.find(qn("w:tcMar"))
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)


def make_table(doc, headers, rows, lang, weights, font_size=7.8, variance_col=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    for c, label in enumerate(headers):
        cell = table.rows[0].cells[c]
        cell.text = label
        shade(cell, NAVY)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell_borders(cell)
        cell_margins(cell)
        for r in cell.paragraphs[0].runs:
            set_font(r, lang, font_size, True, color=WHITE, heading=True)
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)
    for i, values in enumerate(rows):
        cells = table.add_row().cells
        for c, value in enumerate(values):
            cells[c].text = str(value)
            if i % 2:
                shade(cells[c], LIGHT)
            cells[c].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cells[c].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if c == 0 else WD_ALIGN_PARAGRAPH.CENTER
            cell_borders(cells[c])
            cell_margins(cells[c])
            for r in cells[c].paragraphs[0].runs:
                color = BLACK
                if variance_col == c:
                    color = GREEN if str(value).startswith("+") else RED if str(value).startswith("-") else BLACK
                set_font(r, lang, font_size, c == 0, color=color)
    widths = column_widths_from_weights(weights, PAGE_WIDTH_DXA)
    apply_table_geometry(table, widths, table_width_dxa=PAGE_WIDTH_DXA, indent_dxa=100)
    return table


def figure(doc, idx, caption, lang, sources, width=6.35, display_idx=None):
    number = display_idx if display_idx is not None else idx
    label = f"Figure {number}. {caption}" if lang == "en" else f"图 {number}：{caption}"
    paragraph(doc, label, lang, 8.3, True, color=BLACK, before=3.0, after=1.2, keep=True, heading=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0.4)
    shape = p.add_run().add_picture(str(CHARTS[idx]), width=Inches(width))
    shape._inline.docPr.set("descr", caption)
    source(doc, sources, lang)


def page_field(p):
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])
    set_font(run, "en", 7.6, color=GREY)


def setup_doc(lang):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(0.64)
    sec.bottom_margin = Inches(0.58)
    sec.left_margin = Inches(0.72)
    sec.right_margin = Inches(0.72)
    sec.header_distance = Inches(0.27)
    sec.footer_distance = Inches(0.28)
    family = "Songti SC" if lang == "cn" else "Times New Roman"
    fallback = "Arial Unicode MS" if lang == "cn" else family
    normal = doc.styles["Normal"]
    normal.font.name = family
    normal._element.rPr.rFonts.set(qn("w:ascii"), fallback)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), fallback)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), family)
    normal.font.size = Pt(9.2)
    for style_name, size in (("Title", 18), ("Heading 1", 13.2), ("Heading 2", 10.8), ("Heading 3", 9.8)):
        style = doc.styles[style_name]
        hfamily = "Heiti SC" if lang == "cn" else "Times New Roman"
        style.font.name = hfamily
        style._element.rPr.rFonts.set(qn("w:ascii"), fallback)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), fallback)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), hfamily)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
    title_ppr = doc.styles["Title"]._element.get_or_add_pPr()
    border = title_ppr.find(qn("w:pBdr"))
    if border is not None:
        title_ppr.remove(border)
    header = sec.header.paragraphs[0]
    header.text = "Western Digital  |  Q4 FY2026 Earnings Update" if lang == "en" else "西部数据  |  2026 财年第四季度业绩更新"
    set_font(header.runs[0], lang, 7.6, True, color=BLACK, heading=True)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = footer.add_run("WDC Research  |  16 September 2026  |  ")
    set_font(r, lang, 7.5, color=GREY)
    page_field(footer)
    return doc


def new_page(doc):
    doc.add_page_break()


def title_page(doc, lang, mkt):
    price = f"${mkt['price']:.2f}" if mkt["price"] != "N/A" else "N/A"
    mcap = f"${mkt['market_cap']/1e9:.1f}bn" if mkt["market_cap"] != "N/A" else "N/A"
    range52 = f"${mkt['low']:.2f} to ${mkt['high']:.2f}" if mkt["low"] != "N/A" else "N/A"
    if lang == "en":
        paragraph(doc, "Western Digital Q4 FY2026 Earnings Update", lang, 18, True, align=WD_ALIGN_PARAGRAPH.CENTER, after=1, heading=True, style="Title")
        paragraph(doc, "HDD pricing and mix lift earnings while expectations reset", lang, 11.2, True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, after=5.5, heading=True)
        paragraph(doc, "Published 16 September 2026  Results released 5 August 2026  Delayed reaction update", lang, 8.2, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=5)
        paragraph(doc, f"Rating  OUTPERFORM    Price target  $550    Reference price  {price}    Implied upside  {((550/mkt['price']-1)*100):.1f}%" if mkt["price"] != "N/A" else "Rating  OUTPERFORM    Price target  $550    Reference price  N/A", lang, 9.2, True, align=WD_ALIGN_PARAGRAPH.CENTER, after=2.5)
        paragraph(doc, f"Market capitalization  {mcap}    52 week range  {range52}", lang, 8.4, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
        heading(doc, "Earnings summary", lang, 1)
        headers = ["Metric", "Reported", "Consensus", "Variance"]
        rows = [
            ["Revenue", "$3.747bn", "$3.700bn", "+$47m  +1.3%"],
            ["Non-GAAP EPS", "$3.56", "$3.31", "+$0.25  +7.6%"],
            ["Non-GAAP gross margin", "54.4%", "51.0% to 52.0% guide", "+290 bps vs high end"],
            ["Free cash flow", "$1.281bn", "$0.978bn prior quarter", "+31% QoQ"],
        ]
        make_table(doc, headers, rows, lang, [1.45, 1.0, 1.55, 1.45], 7.8, 3)
        source(doc, [("Q4 FY2026 earnings release dated 5 August 2026", URLS["release"]), ("MarketBeat consensus snapshot for 5 August 2026", URLS["consensus"])], lang)
        bullet(doc, "A clean beat driven by price and mix", "Revenue reached the high end of management's prior range, while adjusted EPS cleared consensus by 7.6%. The more important signal was operating leverage: non-GAAP gross margin expanded 1,310 bps year over year and non-GAAP operating margin rose 1,610 bps to 44.2%.", lang, True)
        bullet(doc, "AI storage demand is becoming a measurable HDD tailwind", "Cloud represented 89% of revenue. Exabytes shipped rose 22% year over year, and management expects growth above 25% as 40TB ePMR ramps and 44TB HAMR follows in the first half of calendar 2027.", lang, True)
        bullet(doc, "The share-price reaction reflected expectations rather than weaker fundamentals", "WDC fell about 13% on the next regular session despite the beat. Investors focused on slower near-term exabyte growth, a gross-margin guide below a competitor's level, and the amount of upside already embedded in the pre-print valuation.", lang, True)
        bullet(doc, "Outperform with a $550 price target", "We raise FY2027E revenue and non-GAAP EPS to $18.8bn and $18.90. Our target blends a 24x FY2028E earnings approach with a cash-flow valuation and implies 33.5% upside from the reference price.", lang, True)
    else:
        paragraph(doc, "西部数据 2026 财年第四季度业绩更新", lang, 18, True, align=WD_ALIGN_PARAGRAPH.CENTER, after=1, heading=True, style="Title")
        paragraph(doc, "硬盘定价与产品组合推升盈利 预期重置带来波动", lang, 11.2, True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, after=5.5, heading=True)
        paragraph(doc, "发布日期 2026 年 9 月 16 日  业绩发布日期 2026 年 8 月 5 日  延迟反应更新", lang, 8.2, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=5)
        paragraph(doc, f"评级  跑赢大市    目标价  550 美元    参考股价  {price}    隐含上涨空间  {((550/mkt['price']-1)*100):.1f}%" if mkt["price"] != "N/A" else "评级  跑赢大市    目标价  550 美元    参考股价  N/A", lang, 9.2, True, align=WD_ALIGN_PARAGRAPH.CENTER, after=2.5)
        paragraph(doc, f"市值  {mcap}    52 周区间  {range52}", lang, 8.4, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, after=6)
        heading(doc, "业绩摘要", lang, 1)
        headers = ["指标", "实际值", "一致预期", "差异"]
        rows = [
            ["营业收入", "37.47 亿美元", "37.00 亿美元", "+0.47 亿美元  +1.3%"],
            ["非 GAAP 每股收益", "3.56 美元", "3.31 美元", "+0.25 美元  +7.6%"],
            ["非 GAAP 毛利率", "54.4%", "指引 51.0% 至 52.0%", "较上限高 290 个基点"],
            ["自由现金流", "12.81 亿美元", "上季 9.78 亿美元", "环比 +31%"],
        ]
        make_table(doc, headers, rows, lang, [1.35, 1.1, 1.65, 1.45], 7.8, 3)
        source(doc, [("2026 年 8 月 5 日 Q4 FY2026 业绩公告", URLS["release"]), ("MarketBeat 2026 年 8 月 5 日一致预期快照", URLS["consensus"])], lang)
        bullet(doc, "定价与组合共同推动高质量超预期", "收入达到公司此前指引区间上端，调整后每股收益比一致预期高 7.6%。更关键的是经营杠杆：非 GAAP 毛利率同比扩大 1,310 个基点，非 GAAP 营业利润率同比扩大 1,610 个基点至 44.2%。", lang, True)
        bullet(doc, "AI 存储需求正转化为可量化的 HDD 顺风", "云业务占收入 89%，出货容量同比增长 22%。随着 40TB ePMR 放量以及 44TB HAMR 在 2027 年上半年导入，管理层预计未来容量增速将超过 25%。", lang, True)
        bullet(doc, "股价下跌反映预期差而非经营恶化", "尽管业绩超预期，WDC 下一常规交易日仍下跌约 13%。市场关注短期容量增速放缓、毛利率指引低于主要竞争对手，以及财报前估值已计入较多乐观预期。", lang, True)
        bullet(doc, "给予跑赢大市评级和 550 美元目标价", "我们将 FY2027 收入和非 GAAP 每股收益预测上调至 188 亿美元和 18.90 美元。目标价综合 24 倍 FY2028 每股收益法与现金流估值，相对参考股价有 33.5% 上涨空间。", lang, True)


def page_results(doc, lang):
    new_page(doc)
    heading(doc, "Detailed results and beat analysis" if lang == "en" else "详细业绩与超预期分析", lang)
    if lang == "en":
        paragraph(doc, "Q4 FY2026 revenue rose 44% year over year and 12% sequentially to $3.747bn. The $47m beat versus the published $3.70bn consensus was modest in absolute terms but high quality: growth came from both 22% exabyte expansion and favorable pricing. Management said blended price per terabyte improved from high-single-digit year-over-year growth in Q3 to high-teens growth in Q4. That combination matters because volume alone can consume manufacturing resources without producing the same profit conversion.", lang)
        paragraph(doc, "The comparison base is clean only on a continuing-operations basis. Western Digital completed the Sandisk separation in February 2025; historical flash results are presented as discontinued operations. We therefore use the HDD-only quarterly series in company materials and do not mix pre-separation consolidated revenue into trend charts.", lang)
    else:
        paragraph(doc, "Q4 FY2026 收入同比增长 44%、环比增长 12% 至 37.47 亿美元，较公开一致预期 37.00 亿美元高 0.47 亿美元。绝对超额不大，但质量较高：出货容量同比增长 22%，同时定价环境改善。管理层表示，每 TB 综合价格同比增幅从上一季度的高个位数提高至本季度的高十位数。该组合很重要，因为单纯依赖出货量会占用制造资源，却未必带来同等的利润转化。", lang)
        paragraph(doc, "可比口径必须限定为持续经营业务。西部数据在 2025 年 2 月完成 Sandisk 分拆，历史闪存业务列为终止经营。因此，本报告使用公司披露的纯 HDD 季度序列，不把分拆前合并收入混入趋势图。", lang)
    figure(doc, 1, "Quarterly revenue progression" if lang == "en" else "季度收入走势", lang, [("Q4 FY2026 earnings presentation slides 3 to 6", URLS["deck"]), ("Q1 to Q3 FY2026 company releases", URLS["q3_release"])], 5.8, 1)
    figure(doc, 6, "Q4 FY2026 beat versus published consensus" if lang == "en" else "Q4 FY2026 实际值相对公开一致预期", lang, [("Q4 FY2026 earnings release", URLS["release"]), ("MarketBeat pre-announcement consensus snapshot", URLS["consensus"])], 5.6, 2)


def page_profitability(doc, lang):
    new_page(doc)
    heading(doc, "Profitability and earnings quality" if lang == "en" else "盈利能力与利润质量", lang)
    if lang == "en":
        paragraph(doc, "Non-GAAP gross margin reached 54.4%, up 390 bps sequentially and 1,310 bps year over year. The drivers were a mix shift toward higher-capacity drives, favorable portfolio pricing and manufacturing execution. Operating expenses fell 4% sequentially to $382m even as revenue grew 12%, lifting non-GAAP operating margin to 44.2%. This is the clearest evidence that the current upcycle is producing operating leverage rather than merely replacing lost flash revenue.", lang)
        paragraph(doc, "GAAP EPS of $8.21 should not be capitalized at face value. It included a large mark-to-market gain associated with the retained Sandisk interest and transaction-related effects. Non-GAAP EPS of $3.56 is a better measure of recurring operating earnings and still increased 109% year over year. The $0.25 beat versus consensus was supported by revenue, margin and a controlled expense base, not by a single below-the-line adjustment.", lang)
    else:
        paragraph(doc, "非 GAAP 毛利率达到 54.4%，环比上升 390 个基点、同比上升 1,310 个基点。主要驱动因素是高容量硬盘占比提升、全产品组合定价改善以及制造执行。收入环比增长 12% 的同时，经营费用环比下降 4% 至 3.82 亿美元，使非 GAAP 营业利润率升至 44.2%。这表明当前上行周期带来真实经营杠杆，而非只是填补分拆闪存业务后的收入缺口。", lang)
        paragraph(doc, "GAAP 每股收益 8.21 美元不宜直接用于估值，其中包含与所持 Sandisk 权益公允价值变动及交易相关事项有关的大额收益。非 GAAP 每股收益 3.56 美元更能反映持续经营盈利，且同比仍增长 109%。相对一致预期高出的 0.25 美元来自收入、毛利率和费用控制的共同贡献，而非单一非经营项目。", lang)
    figure(doc, 2, "GAAP and non-GAAP EPS progression" if lang == "en" else "GAAP 与非 GAAP 每股收益走势", lang, [("Q4 FY2026 earnings presentation slides 6 and 10", URLS["deck"]), ("Company quarterly releases", URLS["q1_release"])], 5.8, 3)
    figure(doc, 3, "Non-GAAP margin progression" if lang == "en" else "非 GAAP 利润率走势", lang, [("Q4 FY2026 earnings presentation slides 5 to 7", URLS["deck"]), ("Q4 FY2026 earnings release", URLS["release"])], 5.8, 4)


def page_metrics(doc, lang):
    new_page(doc)
    heading(doc, "Operating metrics and demand mix" if lang == "en" else "经营指标与需求结构", lang)
    if lang == "en":
        paragraph(doc, "Total exabytes shipped reached 231 in Q4, up from 190 a year earlier. Nearline accounted for 209 exabytes and drove almost all absolute growth, while non-nearline was broadly stable at 22 exabytes. Management attributed the slower sequential cadence to customer and recording-technology mix: a quarter with more CMR demand ships fewer bits per unit than one weighted toward UltraSMR. The lumpiness is real, but it does not invalidate the medium-term capacity trajectory.", lang)
        paragraph(doc, "Cloud generated approximately $3.3bn and 89% of revenue; client and consumer represented 6% and 5%. FY2026 cloud revenue rose 38%, supported by 27% exabyte growth and 8% higher ASP per exabyte. Concentration is the trade-off: the top ten customers produced 73% of FY2026 revenue, and three customers individually represented 16%, 15% and 13%.", lang)
    else:
        paragraph(doc, "Q4 总出货容量达到 231EB，高于上年同期的 190EB。其中近线产品为 209EB，贡献绝大部分绝对增量；非近线产品基本稳定在 22EB。管理层将季度环比节奏放缓归因于客户与记录技术组合：CMR 占比较高的季度，每台硬盘可交付的容量低于 UltraSMR 占比较高的季度。季度波动客观存在，但并未否定中期容量增长路径。", lang)
        paragraph(doc, "云业务收入约 33 亿美元，占总收入 89%；客户端和消费端分别占 6% 和 5%。FY2026 云业务收入增长 38%，其中出货容量增长 27%，每 EB 平均售价增长 8%。相应代价是客户集中度高：前十大客户贡献 FY2026 收入的 73%，其中三家客户分别占 16%、15% 和 13%。", lang)
    figure(doc, 4, "Quarterly exabyte shipments" if lang == "en" else "季度出货容量", lang, [("Q4 FY2026 earnings presentation slide 5", URLS["deck"]), ("Q4 FY2026 earnings-call transcript dated 5 August 2026", URLS["transcript"])], 5.75, 5)
    figure(doc, 5, "End-market revenue mix" if lang == "en" else "终端市场收入结构", lang, [("Q4 FY2026 earnings presentation slide 4", URLS["deck"]), ("FY2026 Form 10-K filed 14 August 2026", URLS["tenk"])], 5.75, 6)


def page_guidance(doc, lang):
    new_page(doc)
    heading(doc, "Guidance and near-term outlook" if lang == "en" else "业绩指引与短期展望", lang)
    if lang == "en":
        paragraph(doc, "For Q1 FY2027, management guided revenue to $4.1bn plus or minus $100m, non-GAAP gross margin to 55% to 56%, operating expenses to $390m to $400m and non-GAAP EPS to $4.00 plus or minus $0.15. At the midpoint, revenue would grow about 45% year over year and 9% sequentially. The guide implies continued margin expansion even with product-transition and investment costs.", lang)
        paragraph(doc, "The prior-quarter Q4 guide was $3.65bn plus or minus $100m, 51% to 52% gross margin and $3.25 plus or minus $0.15 EPS. Actual Q4 results landed at the top of the revenue range, 290 bps above the high end of the margin range and $0.16 above the high end of the EPS range. That execution record supports credibility, but investors should avoid extrapolating the same magnitude of beat every quarter.", lang)
        paragraph(doc, "We view management's greater long-term visibility as more important than the exact Q1 midpoint. Customer LTAs now extend across multiple quarters and years, improving demand planning and pricing discipline. They also create execution risk: if WD misses delivery commitments or a customer fails to honor purchases, contractual remedies may not fully replace lost economics.", lang)
    else:
        paragraph(doc, "公司给出的 Q1 FY2027 指引为：收入 41 亿美元，上下浮动 1 亿美元；非 GAAP 毛利率 55% 至 56%；经营费用 3.90 亿至 4.00 亿美元；非 GAAP 每股收益 4.00 美元，上下浮动 0.15 美元。按中值计算，收入同比增长约 45%、环比增长约 9%，并暗示在产品切换与投资支出存在的情况下，利润率仍继续改善。", lang)
        paragraph(doc, "上一季度公司对 Q4 的指引为收入 36.5 亿美元上下浮动 1 亿美元、毛利率 51% 至 52%、每股收益 3.25 美元上下浮动 0.15 美元。实际 Q4 收入位于区间上端，毛利率比区间上限高 290 个基点，每股收益比区间上限高 0.16 美元。执行记录提升了指引可信度，但投资者不应假设每季都能复制同等幅度的超额。", lang)
        paragraph(doc, "我们认为管理层所称的长期可见度提升，比 Q1 中值本身更重要。客户长期协议覆盖多个季度乃至多年，有助于需求规划和定价纪律；但也带来履约风险。如果西部数据未能完成交付承诺，或客户未履行采购义务，合同救济未必能够完全弥补经济损失。", lang)
    figure(doc, 9, "Q1 FY2027 guidance versus Q4 actual" if lang == "en" else "Q1 FY2027 指引与 Q4 实际值对比", lang, [("Q4 FY2026 earnings release page 3", URLS["release"]), ("Q4 FY2026 earnings presentation slide 8", URLS["deck"])], 5.8, 7)
    figure(doc, 7, "Quarterly free cash flow" if lang == "en" else "季度自由现金流", lang, [("Q4 FY2026 earnings presentation slides 6 and 11", URLS["deck"]), ("Q1 and Q2 FY2026 earnings releases", URLS["q2_release"])], 5.8, 8)


def page_thesis(doc, lang):
    new_page(doc)
    heading(doc, "Updated investment thesis" if lang == "en" else "更新后的投资逻辑", lang)
    if lang == "en":
        bullet(doc, "Pillar one strengthened  AI creates persistent storage demand", "Training, inference and agentic workloads generate data that must be retained economically. HDDs remain the dominant repository for large-scale cold and warm data because cost per terabyte is structurally lower than flash. Cloud revenue growth of 38% and 89% revenue mix provide direct evidence, although they also increase dependence on hyperscaler capital spending.", lang)
        bullet(doc, "Pillar two strengthened  Higher-capacity roadmap supports price cost spread", "WD began shipping ePMR drives up to 40TB in Q4. Management expects these products to exceed half of nearline shipments by Q3 FY2027, followed by 44TB HAMR in the first half of calendar 2027 and 50TB products in the second half. Higher areal density can lower cost per terabyte by roughly 10% over time while allowing customers to save rack space and power.", lang)
        bullet(doc, "Pillar three strengthened  Supply discipline converts demand into margins", "FY2026 revenue rose 36% on 25% more exabytes and 8% higher ASP per exabyte. Gross margin increased 970 bps and operating margin increased 1,290 bps. Management says capacity growth does not require a broad unit-capacity build; investment is concentrated in heads, media and automation. That framework limits the risk that today's pricing is competed away through indiscriminate capacity additions.", lang)
        bullet(doc, "Pillar four unchanged  Concentration and cyclicality remain the key debate", "The top ten customers produced 73% of revenue, up from 68% in FY2025. Large cloud buyers purchase unevenly and may shift between CMR and UltraSMR configurations, creating quarterly volatility. LTAs improve visibility but do not eliminate negotiation power, qualification risk or macro exposure.", lang)
        paragraph(doc, "Our thesis therefore depends less on a single earnings beat and more on whether WD sustains a positive spread between price-per-terabyte gains and cost-per-terabyte declines. The FY2026 evidence is encouraging, but the stock now discounts a meaningful portion of the favorable cycle. Position sizing should reflect the possibility of a sharp multiple contraction even if near-term earnings estimates remain intact.", lang)
        heading(doc, "What changed after the call", lang, 2)
        paragraph(doc, "Management's commentary increased our confidence in the duration of the capacity cycle. It described stronger customer visibility, disciplined use of long-term agreements and a roadmap in which capacity growth is delivered primarily through areal-density gains rather than a large increase in drive units. That distinction matters: density-led output can expand shipped exabytes while preserving industry utilization and limiting the incentive for competitors to chase volume with price. We therefore raise our medium-term gross-margin assumptions, but we do not assume that the high-teens price-per-terabyte growth reported for Q4 persists indefinitely.", lang)
        paragraph(doc, "The principal debate is now timing rather than end demand. A temporary pause in quarterly exabyte growth can occur as hyperscalers digest inventory, qualify a new platform or alter the mix between CMR and UltraSMR. We would view a one-quarter volume slowdown as noise if contract pricing, customer commitments and qualification milestones remain intact. By contrast, simultaneous weakening in price per terabyte, nearline units and gross margin would challenge the thesis because it would indicate that supply discipline is no longer offsetting buyer concentration.", lang)
    else:
        bullet(doc, "逻辑一增强  AI 产生需要长期保存的数据", "训练、推理和智能体工作负载会产生必须以经济方式长期保存的数据。由于每 TB 成本结构性低于闪存，HDD 仍是大规模冷数据和温数据的主要载体。云业务收入增长 38%、占比达到 89%，直接支持该逻辑，但也意味着公司更依赖超大规模云厂商资本开支。", lang)
        bullet(doc, "逻辑二增强  高容量路线图有利于扩大价格成本差", "西部数据在 Q4 开始出货最高 40TB 的 ePMR 硬盘。管理层预计该产品到 Q3 FY2027 将占近线出货量一半以上，随后在 2027 年上半年推出 44TB HAMR、下半年推出 50TB 产品。更高面密度有望长期推动每 TB 成本下降约 10%，同时帮助客户节约机架空间与能耗。", lang)
        bullet(doc, "逻辑三增强  供给纪律把需求转化为利润率", "FY2026 收入增长 36%，其中出货容量增长 25%、每 EB 平均售价增长 8%。毛利率扩大 970 个基点，营业利润率扩大 1,290 个基点。管理层表示，容量增长不需要全面增加硬盘台数产能，投资重点是磁头、盘片和自动化，这降低了无序扩产侵蚀定价的风险。", lang)
        bullet(doc, "逻辑四不变  客户集中与周期性仍是核心争议", "前十大客户占收入 73%，高于 FY2025 的 68%。大型云客户采购节奏不均匀，并可能在 CMR 与 UltraSMR 配置之间切换，从而造成季度波动。长期协议提高可见度，但不能消除客户议价权、产品认证风险或宏观风险。", lang)
        paragraph(doc, "因此，我们的投资逻辑不取决于单季超预期，而取决于每 TB 售价提升能否持续高于每 TB 成本下降。FY2026 证据积极，但当前股价已反映相当一部分景气上行。即使短期盈利预测不变，估值倍数仍可能快速收缩，仓位管理应对此留出余地。", lang)
        heading(doc, "电话会后发生的变化", lang, 2)
        paragraph(doc, "管理层评论提高了我们对容量周期持续性的信心。公司表示客户可见度增强，长期协议运用保持纪律，同时容量增长主要来自面密度提升，而不是大幅增加硬盘台数。这一区别十分关键：依靠密度提升可以扩大出货容量，同时维持行业产能利用率，并降低竞争对手通过价格追逐出货量的动机。因此，我们上调中期毛利率假设，但并未假设 Q4 披露的每 TB 价格高十位数增速能够无限期延续。", lang)
        paragraph(doc, "当前主要争议已从终端需求转向实现时点。超大规模云客户消化库存、认证新平台或调整 CMR 与 UltraSMR 组合时，季度容量增速可能暂时停顿。如果合同定价、客户承诺和认证里程碑保持完好，我们会把单季出货放缓视为噪音；反之，若每 TB 价格、近线硬盘台数和毛利率同步走弱，则会挑战投资逻辑，因为这表明供给纪律已无法抵消客户集中带来的议价压力。", lang)
    figure(doc, 8, "FY2026 revenue growth decomposition" if lang == "en" else "FY2026 收入增长分解", lang, [("FY2026 Form 10-K MD&A filed 14 August 2026", URLS["tenk"]), ("Q4 FY2026 earnings-call transcript", URLS["transcript"])], 5.8, 9)


def page_balance_sheet(doc, lang):
    new_page(doc)
    heading(doc, "Cash flow balance sheet and capital returns" if lang == "en" else "现金流 资产负债表与股东回报", lang)
    if lang == "en":
        paragraph(doc, "Q4 operating cash flow was $1.389bn and free cash flow was $1.281bn, equal to a 34% free-cash-flow margin. FY2026 free cash flow reached $3.511bn, up 145%. Cash ended the year at $1.579bn against approximately $1.1bn of debt, leaving about $0.5bn of net cash. The balance sheet is substantially cleaner after debt repayment and monetization of the retained Sandisk stake.", lang)
        paragraph(doc, "Capital returns were aggressive. WD repurchased $1.0bn of shares in Q4 and returned $3.1bn during FY2026 through repurchases and dividends. The board declared a $0.15 quarterly dividend payable on 17 September 2026. Management reiterated that free cash flow will be returned through both programs, but buyback value depends on the price paid. Repurchases at cycle-peak multiples create less value than debt reduction or investment in technology.", lang)
        paragraph(doc, "Liquidity needs remain material. As of 3 July 2026, $710m principal of 2028 convertible notes remained outstanding and management expected cash settlement of principal, with most conversion premium settled in shares. Accounts receivable rose to $2.026bn and inventory to $1.511bn. Supplier long-term commitments totaled $310m, while fiscal 2027 capital expenditure is expected to exceed FY2026 as WD invests in heads, media and automation.", lang)
        paragraph(doc, "Our base case assumes FY2027 free cash flow of about $5.0bn. The principal sensitivities are gross margin, working-capital absorption and the cadence of capacity-related investment. A two-point gross-margin shortfall could reduce annual operating profit by roughly $0.38bn on our FY2027 revenue estimate before secondary effects.", lang)
        heading(doc, "Capital allocation framework", lang, 2)
        paragraph(doc, "We separate cash generation from cash distribution. The first is supported by margin expansion, modest unit-capacity requirements and a low cash-tax burden; the second remains a board decision and can vary with the share price. Our valuation gives credit to repurchases only through the resulting share count, not by adding the authorization amount to enterprise value. This avoids treating a use of cash as incremental value. It also means that a slower repurchase pace would not reduce our operating forecast, although it could modestly dilute per-share earnings relative to the current model.", lang)
        paragraph(doc, "Working capital is the most important near-term cash-flow bridge. Receivables grew as revenue accelerated, while inventory supports product transitions and customer qualification. In a constructive scenario, collections catch up with sales and inventory turns improve as 40TB products scale, allowing cash conversion to exceed reported earnings. In a downside scenario, a customer order pushout leaves inventory elevated and turns a portion of operating profit into working-capital investment. We therefore view quarterly free cash flow as volatile and emphasize the full-year trajectory.", lang)
        paragraph(doc, "Balance-sheet optionality has nevertheless improved. Net cash and strong free cash flow give management room to settle the convertible principal, fund technology transitions and continue dividends without relying on external financing. This lowers financial risk relative to the previous downcycle. It does not remove operating cyclicality, but it should reduce the probability that a weak demand period forces value-destructive financing or a sharp cut in strategic investment.", lang)
    else:
        paragraph(doc, "Q4 经营现金流为 13.89 亿美元，自由现金流为 12.81 亿美元，对应 34% 的自由现金流率。FY2026 自由现金流达到 35.11 亿美元，同比增长 145%。年末现金 15.79 亿美元，债务约 11 亿美元，净现金约 5 亿美元。偿债与处置所持 Sandisk 股权后，资产负债表显著改善。", lang)
        paragraph(doc, "股东回报力度较大。WDC 在 Q4 回购 10 亿美元股票，FY2026 通过回购与分红向股东返还 31 亿美元；董事会宣布每股 0.15 美元季度股息，于 2026 年 9 月 17 日支付。管理层重申将通过两种方式返还自由现金流，但回购价值取决于成交价格；在周期高点估值下回购，创造的价值可能低于偿债或技术投资。", lang)
        paragraph(doc, "流动性需求仍不可忽视。截至 2026 年 7 月 3 日，2028 年可转债剩余本金 7.10 亿美元，管理层预计本金以现金结算，大部分转股溢价以股票结算。应收账款增至 20.26 亿美元，库存增至 15.11 亿美元；供应商长期承诺总额 3.10 亿美元。由于投资磁头、盘片和自动化，FY2027 资本开支预计高于 FY2026。", lang)
        paragraph(doc, "我们的基准情景假设 FY2027 自由现金流约 50 亿美元，主要敏感因素是毛利率、营运资金占用以及产能相关投资节奏。按我们的 FY2027 收入预测，若毛利率比预期低 2 个百分点，在考虑二阶影响前，年度营业利润可能减少约 3.8 亿美元。", lang)
        heading(doc, "资本配置框架", lang, 2)
        paragraph(doc, "我们区分现金创造与现金分配。前者受益于利润率扩大、有限的硬盘台数产能需求和较低现金税负；后者取决于董事会决策，并会随股价变化。估值只通过回购后的股数变化体现其价值，不把回购授权金额直接加到企业价值中，从而避免把现金用途误认为新增价值。这也意味着，回购节奏放慢不会降低经营预测，但可能使每股收益相对当前模型略有摊薄。", lang)
        paragraph(doc, "营运资金是近期现金流最重要的桥接项目。收入加速推动应收账款增加，库存则服务于产品切换和客户认证。在积极情景下，回款追上销售增速，40TB 产品放量使库存周转改善，现金转化可能超过账面盈利；在下行情景下，客户推迟订单会使库存维持高位，并把部分营业利润转化为营运资金占用。因此，我们认为季度自由现金流会波动，更看重全年趋势。", lang)
        paragraph(doc, "资产负债表的选择空间已经改善。净现金与强劲自由现金流使管理层能够偿付可转债本金、支持技术转换并继续派息，而无需依赖外部融资。相较上一轮下行周期，这降低了财务风险。经营周期性仍然存在，但需求疲弱时被迫进行损害价值的融资或大幅削减战略投资的概率已经下降。", lang)


def page_valuation(doc, lang, mkt):
    new_page(doc)
    heading(doc, "Valuation and price target" if lang == "en" else "估值与目标价", lang)
    if lang == "en":
        paragraph(doc, "At the $411.96 reference price, WDC trades at 21.8x our FY2027E non-GAAP EPS of $18.90 and 16.5x FY2028E EPS of $25.00. Longbridge's trailing P/E screen showed 16.0x on 15 September 2026, but that measure is flattered by large GAAP gains from the Sandisk stake. Forward adjusted earnings are the more decision-useful basis.", lang)
        paragraph(doc, "Our $550 price target blends two approaches. First, a 24x multiple on FY2028E non-GAAP EPS produces $600 per share; the premium reflects AI-related storage growth, improved LTAs and the widening price-cost spread. Second, a DCF with FY2027-FY2031 free cash flow of roughly $5.0bn, $7.5bn, $10.0bn, $12.5bn and $15.0bn, a 10.5% WACC and 3.5% terminal growth yields approximately $473 per share. A 60% earnings-multiple and 40% DCF blend gives about $549, rounded to $550.", lang)
        paragraph(doc, "The target implies 33.5% upside, supporting Outperform, but valuation is not undemanding. The bull case of $700 assumes FY2028 EPS of $28 and a 25x multiple; the bear case of $300 assumes EPS of $20 and a 15x multiple. The resulting range emphasizes that the investment is now highly sensitive to the durability of pricing and the timing of HAMR execution.", lang)
        paragraph(doc, "The market's 13% decline in the session after results illustrates this sensitivity. A strong quarter can coincide with a negative share-price reaction when consensus rises faster than reported fundamentals. We would use pullbacks to build exposure rather than treat a single target as a precise estimate of intrinsic value.", lang)
        heading(doc, "Catalysts and valuation checkpoints", lang, 2)
        paragraph(doc, "The next catalysts are Q1 FY2027 results, evidence that 40TB ePMR exceeds half of nearline shipments by Q3 FY2027, and successful 44TB HAMR qualification in the first half of calendar 2027. The earnings path can support the target before HAMR contributes materially if price per terabyte remains firm and cost per terabyte continues to decline. Conversely, the market is likely to penalize even a small roadmap delay because the current multiple already assumes that density gains extend the margin cycle. We would revisit the target if FY2027 gross margin tracks below 54%, if annual exabyte growth falls below the mid-teens, or if customer concentration rises without stronger contractual protection.", lang)
    else:
        paragraph(doc, "按 411.96 美元参考股价计算，WDC 对应我们 FY2027 非 GAAP 每股收益 18.90 美元的市盈率为 21.8 倍，对应 FY2028 每股收益 25.00 美元的市盈率为 16.5 倍。Longbridge 在 2026 年 9 月 15 日显示的滚动市盈率为 16.0 倍，但该指标受到 Sandisk 股权相关大额 GAAP 收益提振，因此前瞻调整后盈利更适合作为决策依据。", lang)
        paragraph(doc, "550 美元目标价综合两种方法。其一，对 FY2028 非 GAAP 每股收益给予 24 倍市盈率，得到 600 美元；溢价反映 AI 存储增长、长期协议改善和价格成本差扩大。其二，DCF 假设 FY2027 至 FY2031 自由现金流约为 50 亿、75 亿、100 亿、125 亿和 150 亿美元，WACC 为 10.5%，永续增长率为 3.5%，得到约 473 美元。按市盈率法 60%、DCF 40% 加权，结果约 549 美元，取整为 550 美元。", lang)
        paragraph(doc, "目标价对应 33.5% 上涨空间，支持跑赢大市评级，但估值并不便宜。乐观情景 700 美元假设 FY2028 每股收益 28 美元、估值 25 倍；悲观情景 300 美元假设每股收益 20 美元、估值 15 倍。区间表明，当前投资回报高度依赖定价持续性与 HAMR 执行时点。", lang)
        paragraph(doc, "财报后下一交易日下跌约 13%，正体现这种敏感性。当一致预期上升速度快于已披露基本面时，强劲季度也可能伴随负面股价反应。我们倾向于利用回调逐步建仓，而不把单一目标价视为精确的内在价值。", lang)
        heading(doc, "催化剂与估值检查点", lang, 2)
        paragraph(doc, "下一阶段催化剂包括 Q1 FY2027 业绩、40TB ePMR 到 Q3 FY2027 占近线出货量一半以上的证据，以及 44TB HAMR 在 2027 年上半年顺利完成认证。即使 HAMR 尚未形成重大收入，只要每 TB 价格保持坚挺、每 TB 成本继续下降，盈利路径仍可支持目标价。反之，由于当前估值已假设面密度提升会延长利润率周期，即便路线图小幅延误也可能受到市场惩罚。如果 FY2027 毛利率低于 54%、年度容量增速降至中十位数以下，或客户集中度上升却缺少更强合同保护，我们将重新评估目标价。", lang)
    figure(doc, 11, "One-year adjusted share-price history" if lang == "en" else "一年期复权股价走势", lang, [("Yahoo Finance market data accessed 16 September 2026", URLS["yahoo"]), ("Longbridge WDC.US daily K-line accessed 16 September 2026", URLS["quote"])], 6.25, 10)


def page_estimates(doc, lang):
    new_page(doc)
    heading(doc, "Updated estimates" if lang == "en" else "更新后的预测", lang)
    if lang == "en":
        paragraph(doc, "We raise estimates to reflect the Q4 beat, the Q1 guide and stronger price-per-terabyte assumptions. FY2027 revenue increases 3.3% to $18.8bn and non-GAAP EPS rises 7.4% to $18.90. The larger EPS revision reflects gross-margin leverage and a relatively contained expense base. FY2028 assumes the 40TB ePMR ramp is mature and 44TB HAMR has moved beyond initial qualification.", lang)
        headers = ["Metric", "FY2027E old", "FY2027E new", "Change", "FY2028E"]
        rows = [
            ["Revenue ($bn)", "$18.2", "$18.8", "+3.3%", "$22.3"],
            ["Revenue growth", "+40.9%", "+45.5%", "+460 bps", "+18.6%"],
            ["Non-GAAP gross margin", "54.8%", "56.0%", "+120 bps", "57.5%"],
            ["Non-GAAP operating income ($bn)", "$8.0", "$8.4", "+5.0%", "$10.7"],
            ["Non-GAAP EPS", "$17.60", "$18.90", "+7.4%", "$25.00"],
            ["Free cash flow ($bn)", "$4.5", "$5.0", "+11.1%", "$7.5"],
            ["P/E at $411.96", "23.4x", "21.8x", "-1.6x", "16.5x"],
        ]
    else:
        paragraph(doc, "我们根据 Q4 超预期、Q1 指引及更强的每 TB 定价假设上调预测。FY2027 收入预测提高 3.3% 至 188 亿美元，非 GAAP 每股收益提高 7.4% 至 18.90 美元。每股收益上调幅度更大，反映毛利率杠杆与相对受控的费用基数。FY2028 假设 40TB ePMR 放量成熟，44TB HAMR 已越过初始认证阶段。", lang)
        headers = ["指标", "FY2027 原预测", "FY2027 新预测", "变化", "FY2028 预测"]
        rows = [
            ["收入 十亿美元", "18.2", "18.8", "+3.3%", "22.3"],
            ["收入增长", "+40.9%", "+45.5%", "+460 个基点", "+18.6%"],
            ["非 GAAP 毛利率", "54.8%", "56.0%", "+120 个基点", "57.5%"],
            ["非 GAAP 营业利润 十亿美元", "8.0", "8.4", "+5.0%", "10.7"],
            ["非 GAAP 每股收益", "17.60", "18.90", "+7.4%", "25.00"],
            ["自由现金流 十亿美元", "4.5", "5.0", "+11.1%", "7.5"],
            ["按 411.96 美元计算市盈率", "23.4 倍", "21.8 倍", "-1.6 倍", "16.5 倍"],
        ]
    make_table(doc, headers, rows, lang, [1.6, 1.15, 1.15, 1.05, 1.2], 7.6, 3)
    source(doc, [("Company results and guidance", URLS["release"]), ("Analyst estimates dated 16 September 2026", None)], lang)
    if lang == "en":
        paragraph(doc, "Estimate risk is asymmetric around gross margin. Every 100 bps change in FY2027 gross margin changes gross profit by roughly $188m, before tax and share-count effects. Exabyte growth has more muted short-term earnings sensitivity if it is accompanied by lower price per terabyte; conversely, disciplined pricing can drive EPS even when quarterly shipment mix is lumpy.", lang)
        paragraph(doc, "Our FY2027 quarterly framework starts with management's $4.1bn Q1 midpoint, then assumes $4.5bn, $5.0bn and $5.2bn across Q2 to Q4. The profile is not intended as point precision. It expresses a view that customer visibility and high-capacity qualification support a rising revenue base while annual comparisons remain strongest in the first half.", lang)
    else:
        paragraph(doc, "预测风险对毛利率呈非对称。FY2027 毛利率每变化 100 个基点，在税收和股数影响前，毛利润约变化 1.88 亿美元。如果容量增长伴随每 TB 价格下降，短期盈利敏感度相对有限；反之，即便季度出货组合波动，定价纪律仍可推动每股收益。", lang)
        paragraph(doc, "我们的 FY2027 季度框架以公司 Q1 收入中值 41 亿美元为起点，Q2 至 Q4 分别假设 45 亿、50 亿和 52 亿美元。该路径不是精确点预测，而是表达以下判断：客户可见度和高容量产品认证支撑收入基数上升，上半年同比增速仍然最强。", lang)
    figure(doc, 10, "Illustrative estimate revisions" if lang == "en" else "示意性预测调整", lang, [("Q4 FY2026 earnings release and guidance", URLS["release"]), ("Analyst estimates dated 16 September 2026", None)], 5.9, 11)


def page_risks_sources(doc, lang):
    new_page(doc)
    heading(doc, "Risks and sources" if lang == "en" else "风险与资料来源", lang)
    if lang == "en":
        heading(doc, "Risk monitor", lang, 2)
        headers = ["Risk", "Current evidence", "What would change the view"]
        rows = [
            ["Customer concentration", "Top 10 customers were 73% of FY2026 revenue", "A hyperscaler order pushout or LTA dispute"],
            ["Pricing cycle", "High-teens price-per-TB growth in Q4", "ASP growth falling below cost-per-TB improvement"],
            ["Technology execution", "40TB ePMR shipping; 44TB HAMR due in H1 2027", "Qualification delay or weak yields"],
            ["Capital allocation", "$3.1bn returned in FY2026", "Buybacks at high multiples without durable FCF"],
            ["Valuation", "21.8x FY2027E adjusted EPS", "Multiple compression despite estimate delivery"],
        ]
    else:
        heading(doc, "风险监测", lang, 2)
        headers = ["风险", "当前证据", "会改变判断的信号"]
        rows = [
            ["客户集中", "前十大客户占 FY2026 收入 73%", "超大规模客户推迟订单或长期协议争议"],
            ["定价周期", "Q4 每 TB 价格同比增长高十位数", "售价增速低于每 TB 成本改善幅度"],
            ["技术执行", "40TB ePMR 已出货 44TB HAMR 计划 2027 上半年推出", "认证延期或良率不佳"],
            ["资本配置", "FY2026 向股东返还 31 亿美元", "在高估值下回购且自由现金流缺乏持续性"],
            ["估值", "FY2027 调整后每股收益市盈率 21.8 倍", "即使盈利兑现仍发生估值压缩"],
        ]
    make_table(doc, headers, rows, lang, [1.25, 2.05, 2.25], 7.5)
    source(doc, [("FY2026 Form 10-K risk factors and MD&A", URLS["tenk"]), ("Q4 FY2026 earnings-call transcript", URLS["transcript"])], lang)
    heading(doc, "Sources and references" if lang == "en" else "资料来源与参考文件", lang, 2)
    refs = [
        (("Q4 FY2026 earnings release  5 August 2026" if lang == "en" else "Q4 FY2026 业绩公告  2026 年 8 月 5 日"), URLS["release"]),
        (("Q4 FY2026 earnings presentation  5 August 2026" if lang == "en" else "Q4 FY2026 业绩演示文稿  2026 年 8 月 5 日"), URLS["deck"]),
        (("Q4 FY2026 earnings call and webcast  5 August 2026" if lang == "en" else "Q4 FY2026 电话会与网络直播  2026 年 8 月 5 日"), URLS["event"]),
        (("Q4 FY2026 earnings-call transcript  5 August 2026" if lang == "en" else "Q4 FY2026 电话会文字记录  2026 年 8 月 5 日"), URLS["transcript"]),
        (("Form 8-K  filed 5 August 2026" if lang == "en" else "Form 8-K  提交于 2026 年 8 月 5 日"), URLS["eightk"]),
        (("FY2026 Form 10-K  filed 14 August 2026" if lang == "en" else "FY2026 Form 10-K  提交于 2026 年 8 月 14 日"), URLS["tenk_index"]),
        (("Q3 FY2026 earnings release and prior guidance  30 April 2026" if lang == "en" else "Q3 FY2026 业绩公告与此前指引  2026 年 4 月 30 日"), URLS["q3_release"]),
        (("MarketBeat consensus snapshot for 5 August 2026" if lang == "en" else "MarketBeat 2026 年 8 月 5 日一致预期快照"), URLS["consensus"]),
        (("Longbridge quote valuation and K-line data  accessed 16 September 2026" if lang == "en" else "Longbridge 行情 估值与 K 线数据  访问于 2026 年 9 月 16 日"), URLS["quote"]),
        (("Yahoo Finance dynamic market data  accessed 16 September 2026" if lang == "en" else "Yahoo Finance 动态市场数据  访问于 2026 年 9 月 16 日"), URLS["yahoo"]),
    ]
    for label, url in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.2)
        r = p.add_run("• ")
        set_font(r, lang, 8.1)
        add_hyperlink(p, label, url, lang, 8.1)
    if lang == "en":
        paragraph(doc, "Methodology and disclosure  Reported figures are company data. Consensus figures are third-party pre-announcement snapshots. Forward estimates, rating and price target are analytical judgments prepared for this report and are not company guidance. This document is informational research, not personalized investment advice.", lang, 8.0, italic=True, color=GREY, before=2)
        heading(doc, "Analytical conventions", lang, 2)
        paragraph(doc, "All quarterly operating comparisons use continuing HDD operations following the February 2025 Sandisk separation. Non-GAAP figures follow the company's definitions and are used for trend and valuation analysis because GAAP earnings include separation-related items and mark-to-market effects. Free cash flow is operating cash flow less capital expenditure. Market price, market capitalization and 52-week range were retrieved dynamically on 16 September 2026; the reference price is the 15 September 2026 close. Figures may not sum because of rounding.", lang, 8.0, color=GREY)
        paragraph(doc, "Our DCF is a scenario tool rather than a point forecast. The stated cash flows, discount rate and terminal growth rate are explicit analyst assumptions. The earnings-multiple method is cross-checked against the company's growth, margin, customer concentration and technology-execution profile rather than mechanically adopting the market's highest target. The rating would be reconsidered if the risk-adjusted upside fell below the level required for an Outperform or if evidence emerged that the favorable price-cost spread is reversing.", lang, 8.0, color=GREY)
    else:
        paragraph(doc, "方法与披露  已公布数据来自公司资料，一致预期来自第三方财报前快照。前瞻预测、评级与目标价是本报告的分析判断，并非公司指引。本文件仅供研究与信息参考，不构成个性化投资建议。", lang, 8.0, italic=True, color=GREY, before=2)
        heading(doc, "分析口径", lang, 2)
        paragraph(doc, "2025 年 2 月 Sandisk 分拆完成后，本报告所有季度经营比较均采用持续经营的 HDD 业务口径。非 GAAP 数据遵循公司定义，用于趋势和估值分析，因为 GAAP 盈利包含分拆相关项目与公允价值变动。自由现金流定义为经营现金流减资本开支。股价、市值和 52 周区间在 2026 年 9 月 16 日动态获取，参考股价为 2026 年 9 月 15 日收盘价。由于四舍五入，个别合计数可能存在差异。", lang, 8.0, color=GREY)
        paragraph(doc, "DCF 是情景分析工具，而不是精确点预测。所列现金流、折现率和永续增长率均为分析师明确假设。市盈率法会结合公司的增长、利润率、客户集中与技术执行风险进行交叉检验，而不是机械采用市场最高目标价。如果风险调整后的上涨空间低于跑赢大市评级所需水平，或出现价格成本差正在逆转的证据，我们会重新评估评级。", lang, 8.0, color=GREY)


def build(lang, path):
    mkt = market_data()
    doc = setup_doc(lang)
    title_page(doc, lang, mkt)
    page_results(doc, lang)
    page_profitability(doc, lang)
    page_metrics(doc, lang)
    page_guidance(doc, lang)
    page_thesis(doc, lang)
    page_balance_sheet(doc, lang)
    page_valuation(doc, lang, mkt)
    page_estimates(doc, lang)
    page_risks_sources(doc, lang)
    doc.core_properties.title = "Western Digital Q4 FY2026 Earnings Update" if lang == "en" else "西部数据 2026 财年第四季度业绩更新"
    doc.core_properties.subject = "Earnings analysis"
    doc.core_properties.author = "Research"
    doc.save(path)
    print(path)


if __name__ == "__main__":
    build("en", OUT / "WDC_Q4_FY2026_Earnings_Update.docx")
    build("cn", OUT / "WDC_Q4_FY2026_业绩更新报告_中文版.docx")
