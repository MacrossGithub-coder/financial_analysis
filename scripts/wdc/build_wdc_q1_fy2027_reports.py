#!/usr/bin/env python3
"""Build bilingual Western Digital Q1 FY2027 earnings update reports."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

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
    "26.1017.22135/skills/documents/scripts"
)
sys.path.insert(0, str(DOC_SCRIPTS))
from table_geometry import apply_table_geometry, column_widths_from_weights  # noqa: E402

PAGE_WIDTH_DXA = 10050
NAVY = "123B5D"
TEAL = "008A8A"
GREEN = "2F7D32"
ORANGE = "B85C00"
RED = "B42318"
GREY = "667085"
LIGHT = "F2F4F7"
WHITE = "FFFFFF"
BLACK = "000000"

URLS = {
    "release": "https://investor.wdc.com/node/28921",
    "release_pdf": "https://investor.wdc.com/static-files/d4c70967-9cf8-45ff-ba15-93de77163154",
    "presentation": "https://investor.wdc.com/static-files/0c54ec27-0e0d-46b1-a56b-2c72b02815ff",
    "webcast": "https://investor.wdc.com/events/event-details/western-digital-first-quarter-fiscal-2027-earnings-call",
    "transcript": "https://www.marketbeat.com/earnings/reports/2026-10-22-western-digital-co-stock/",
    "eightk": "https://www.sec.gov/Archives/edgar/data/106040/000162828026066882/0001628280-26-066882-index.htm",
    "tenq": "https://www.sec.gov/Archives/edgar/data/106040/000162828026066933/0001628280-26-066933-index.htm",
    "tenq_doc": "https://www.sec.gov/Archives/edgar/data/106040/000162828026066933/wdc-20261002.htm",
    "q4_release": "https://investor.wdc.com/node/28586",
    "consensus": "https://www.marketbeat.com/stocks/NASDAQ/WDC/earnings/",
    "longbridge": "https://longbridge.com/en/quote/WDC.US",
    "ir": "https://investor.wdc.com/financial-information/earnings-documents",
}

CHARTS = {
    1: OUT / "wdc_q1fy27_chart1_revenue.png",
    2: OUT / "wdc_q1fy27_chart2_eps.png",
    3: OUT / "wdc_q1fy27_chart3_margins.png",
    4: OUT / "wdc_q1fy27_chart4_fcf.png",
    5: OUT / "wdc_q1fy27_chart5_mix.png",
    6: OUT / "wdc_q1fy27_chart6_exabytes.png",
    7: OUT / "wdc_q1fy27_chart7_beat.png",
    8: OUT / "wdc_q1fy27_chart8_guidance.png",
    9: OUT / "wdc_q1fy27_chart9_revisions.png",
    10: OUT / "wdc_q1fy27_chart10_valuation.png",
    11: OUT / "wdc_q1fy27_chart11_price.png",
}


def market_data() -> dict:
    result = {"price": 343.67, "previous_close": 354.19, "date": "23 October 2026"}
    try:
        proc = subprocess.run(
            ["longbridge", "quote", "WDC.US", "--format", "json"],
            check=True,
            capture_output=True,
            text=True,
            timeout=30,
        )
        quote = json.loads(proc.stdout)[0]
        result.update(
            price=float(quote["last"]),
            previous_close=float(quote["prev_close"]),
            date="23 October 2026",
        )
    except Exception as exc:
        print(f"WARNING: Longbridge quote unavailable; using dated 23 October snapshot: {exc}")
    result["market_cap"] = result["price"] * 374_000_000 / 1_000_000_000
    result["one_day_change"] = result["price"] / result["previous_close"] - 1
    return result


def set_run_font(run, lang, size=9.4, bold=False, italic=False, color=BLACK, heading=False):
    face = ("Heiti SC" if heading else "Songti SC") if lang == "cn" else "Times New Roman"
    east_asia_face = face
    run.font.name = face
    run._element.get_or_add_rPr()
    for slot in ("w:ascii", "w:hAnsi"):
        run._element.rPr.rFonts.set(qn(slot), face)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east_asia_face)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    if lang == "cn":
        node = run._element.rPr.find(qn("w:lang"))
        if node is None:
            node = OxmlElement("w:lang")
            run._element.rPr.append(node)
        node.set(qn("w:val"), "zh-CN")
        node.set(qn("w:eastAsia"), "zh-CN")


def paragraph(doc, text, lang, size=9.4, bold=False, italic=False, color=BLACK, before=0, after=4.1, align=None, keep=False, heading=False, style=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    set_run_font(r, lang, size=size, bold=bold, italic=italic, color=color, heading=heading)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.03
    p.paragraph_format.keep_with_next = keep
    if align is not None:
        p.alignment = align
    return p


def heading(doc, text, lang, level=1):
    p = doc.add_paragraph(text, style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(1.0)
    p.paragraph_format.space_after = Pt(4.0)
    p.paragraph_format.keep_with_next = True
    return p


def bullet(doc, title, body, lang, compact=False):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.03)
    p.paragraph_format.first_line_indent = Inches(-0.03)
    p.paragraph_format.space_after = Pt(3.0 if compact else 4.0)
    p.paragraph_format.line_spacing = 1.03
    r = p.add_run("■ ")
    set_run_font(r, lang, size=8.9 if compact else 9.3, bold=True, color=NAVY, heading=True)
    r = p.add_run(title + " ")
    set_run_font(r, lang, size=8.9 if compact else 9.3, bold=True, heading=True)
    r = p.add_run(body)
    set_run_font(r, lang, size=8.9 if compact else 9.3)
    return p


def add_hyperlink(p, label, url, lang, size=7.5):
    rel_id = p.part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rel_id)
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    face = "Songti SC" if lang == "cn" else "Times New Roman"
    east_asia_face = face
    for slot in ("w:ascii", "w:hAnsi"):
        fonts.set(qn(slot), face)
    fonts.set(qn("w:eastAsia"), east_asia_face)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(round(size * 2)))
    for node in (fonts, color, underline, sz):
        props.append(node)
    run.append(props)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    p._p.append(link)


def source(doc, items, lang, lead=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0.2)
    p.paragraph_format.space_after = Pt(2.7)
    p.paragraph_format.line_spacing = 1.0
    prefix = lead or ("Sources: " if lang == "en" else "资料来源：")
    r = p.add_run(prefix)
    set_run_font(r, lang, size=7.4, italic=True, color=GREY)
    for idx, (label, url) in enumerate(items):
        if idx:
            r = p.add_run("; ")
            set_run_font(r, lang, size=7.4, italic=True, color=GREY)
        add_hyperlink(p, label, url, lang, size=7.4)
    r = p.add_run(".")
    set_run_font(r, lang, size=7.4, italic=True, color=GREY)
    return p


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=90, bottom=80, end=90):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def make_table(doc, headers, rows, lang, weights, variance_col=None, font_size=7.7):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    for col, label in enumerate(headers):
        cell = table.rows[0].cells[col]
        cell.text = label
        shade(cell, NAVY)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in cell.paragraphs[0].runs:
            set_run_font(run, lang, size=font_size, bold=True, color=WHITE, heading=True)
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    marker = OxmlElement("w:tblHeader")
    marker.set(qn("w:val"), "true")
    tr_pr.append(marker)
    for row_idx, values in enumerate(rows):
        cells = table.add_row().cells
        for col, value in enumerate(values):
            cells[col].text = str(value)
            set_cell_margins(cells[col])
            if row_idx % 2:
                shade(cells[col], LIGHT)
            cells[col].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cells[col].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT if col == 0 else WD_ALIGN_PARAGRAPH.CENTER
            for run in cells[col].paragraphs[0].runs:
                color = BLACK
                if variance_col == col:
                    color = GREEN if str(value).startswith("+") else RED if str(value).startswith("-") else BLACK
                set_run_font(run, lang, size=font_size, bold=(col == 0), color=color)
    widths = column_widths_from_weights(weights, PAGE_WIDTH_DXA)
    apply_table_geometry(table, widths, table_width_dxa=PAGE_WIDTH_DXA, indent_dxa=100)
    return table


def figure(doc, idx, caption, lang, sources, width=5.55):
    label = f"Figure {idx}. {caption}" if lang == "en" else f"图 {idx}：{caption}"
    paragraph(doc, label, lang, size=8.15, bold=True, after=0.8, keep=True, heading=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0.2)
    shape = p.add_run().add_picture(str(CHARTS[idx]), width=Inches(width))
    shape._inline.docPr.set("descr", caption)
    shape._inline.docPr.set("title", caption)
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
    set_run_font(run, "en", size=7.5, color=GREY)


def setup_doc(lang):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.62)
    section.bottom_margin = Inches(0.58)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.header_distance = Inches(0.28)
    section.footer_distance = Inches(0.28)
    normal = doc.styles["Normal"]
    face = "Songti SC" if lang == "cn" else "Times New Roman"
    east_asia_face = face
    normal.font.name = face
    for slot in ("w:ascii", "w:hAnsi"):
        normal._element.rPr.rFonts.set(qn(slot), face)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), east_asia_face)
    normal.font.size = Pt(9.4)
    normal.paragraph_format.space_after = Pt(4.0)
    normal.paragraph_format.line_spacing = 1.03
    title = doc.styles["Title"]
    title_face = "Heiti SC" if lang == "cn" else "Times New Roman"
    title_east_asia_face = title_face
    title.font.name = title_face
    for slot in ("w:ascii", "w:hAnsi"):
        title._element.rPr.rFonts.set(qn(slot), title_face)
    title._element.rPr.rFonts.set(qn("w:eastAsia"), title_east_asia_face)
    title.font.size = Pt(17.5)
    title.font.bold = True
    title.font.color.rgb = RGBColor.from_string(BLACK)
    title_ppr = title._element.get_or_add_pPr()
    title_border = title_ppr.find(qn("w:pBdr"))
    if title_border is not None:
        title_ppr.remove(title_border)
    for style_name, size in (("Heading 1", 13.3), ("Heading 2", 10.7), ("Heading 3", 9.7)):
        style = doc.styles[style_name]
        style_face = "Heiti SC" if lang == "cn" else "Times New Roman"
        style_east_asia_face = style_face
        style.font.name = style_face
        for slot in ("w:ascii", "w:hAnsi"):
            style._element.rPr.rFonts.set(qn(slot), style_face)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), style_east_asia_face)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(BLACK)
    header = section.header.paragraphs[0]
    header.text = (
        "Western Digital (Nasdaq: WDC)  |  Q1 FY2027 Earnings Update"
        if lang == "en"
        else "西部数据（纳斯达克：WDC）｜2027 财年第一季度业绩更新"
    )
    set_run_font(header.runs[0], lang, size=7.6, bold=True, heading=True)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = footer.add_run("Independent Research  |  24 October 2026  |  ")
    set_run_font(r, lang, size=7.4, color=GREY)
    page_field(footer)
    return doc


def new_page(doc):
    doc.add_page_break()


def title_page(doc, lang, mkt):
    if lang == "en":
        paragraph(doc, "WESTERN DIGITAL CORPORATION", lang, size=9.0, bold=True, color=NAVY, after=2.0, heading=True)
        paragraph(doc, "Q1 FY2027 Earnings Update", lang, size=17.5, bold=True, after=2.0, style="Title", heading=True)
        paragraph(doc, "Beat and raise validates the AI storage cycle while investment intensity tempers the rerating", lang, size=11.0, bold=True, after=5.0, heading=True)
        paragraph(doc, "24 October 2026  |  Results released 22 October 2026", lang, size=8.1, color=GREY, after=5.0)
        paragraph(doc, f"Rating: OUTPERFORM   |   Price target: $620 (prior $600)   |   Price: ${mkt['price']:.2f} at 23 Oct. close", lang, size=10.4, bold=True, color=NAVY, after=6.0, heading=True)
        heading(doc, "Earnings summary", lang, 2)
        make_table(
            doc,
            ["Metric", "Reported", "Consensus", "Variance", "YoY"],
            [
                ["Revenue", "$4.381B", "$4.145B", "+$236M / +5.7%", "+55%"],
                ["Non-GAAP EPS", "$4.64", "$4.10", "+$0.54 / +13.2%", "+161%"],
                ["Non-GAAP gross margin", "58.1%", "Prior guide 55%-56%", "+210-310 bps", "+1,420 bps"],
            ],
            lang,
            [1.55, 1.15, 1.45, 1.55, 0.95],
            variance_col=3,
        )
        source(doc, [("WDC Q1 FY2027 release, 22 Oct. 2026", URLS["release"]), ("MarketBeat consensus snapshot, 22 Oct. 2026", URLS["consensus"])], lang)
        bullet(doc, "The quarter beat on both revenue and earnings.", "Revenue exceeded pre-release consensus by $236m, while non-GAAP EPS beat by $0.54. The gross-margin result was more important than the headline EPS: 58.1% cleared the high end of prior guidance by 210 bps and demonstrated that pricing, capacity mix and disciplined supply remain mutually reinforcing.", lang, compact=True)
        bullet(doc, "Q2 guidance resets the earnings run-rate higher.", "Management guided to $5.1bn of revenue, 60%-61% non-GAAP gross margin and $5.80 non-GAAP EPS at the midpoints. Current Street figures have largely moved to that range, but the guide still establishes a materially stronger base than the framework embedded in our prior FY2027 estimates.", lang, compact=True)
        bullet(doc, "The investment debate moves from demand to execution.", "Cloud remained 88% of revenue and demand visibility extends into 2030-2031. The debate is now whether the 40TB ePMR and HAMR ramps can convert that visibility into exabyte supply without excessive capital intensity, yield loss or cost inflation.", lang, compact=True)
        bullet(doc, "We raise the target to $620 and retain Outperform.", "Our FY2027 revenue and non-GAAP EPS estimates rise to $21.3bn and $25.00. A 24.8x multiple on the revised EPS view supports $620, 80% above the 23 October close. We use a lower multiple than the estimate growth might otherwise justify because capex will run modestly above the normal 4%-6% of revenue range for two years.", lang, compact=True)
    else:
        paragraph(doc, "西部数据公司", lang, size=9.0, bold=True, color=NAVY, after=2.0, heading=True)
        paragraph(doc, "2027 财年第一季度业绩更新", lang, size=17.5, bold=True, after=2.0, style="Title", heading=True)
        paragraph(doc, "业绩与指引双超预期验证 AI 存储周期 投资强度则限制估值扩张", lang, size=11.0, bold=True, after=5.0, heading=True)
        paragraph(doc, "2026 年 10 月 24 日｜业绩于 2026 年 10 月 22 日发布", lang, size=8.1, color=GREY, after=5.0)
        paragraph(doc, f"评级：跑赢大盘｜目标价：620 美元（此前 600 美元）｜股价：{mkt['price']:.2f} 美元（10 月 23 日收盘）", lang, size=10.4, bold=True, color=NAVY, after=6.0, heading=True)
        heading(doc, "业绩摘要", lang, 2)
        make_table(
            doc,
            ["指标", "实际", "一致预期", "差异", "同比"],
            [
                ["收入", "43.81 亿美元", "41.45 亿美元", "+2.36 亿 / +5.7%", "+55%"],
                ["Non-GAAP EPS", "4.64 美元", "4.10 美元", "+0.54 / +13.2%", "+161%"],
                ["Non-GAAP 毛利率", "58.1%", "此前指引 55%-56%", "+210 至 310 基点", "+1,420 基点"],
            ],
            lang,
            [1.55, 1.15, 1.45, 1.55, 0.95],
            variance_col=3,
        )
        source(doc, [("WDC 2027 财年 Q1 业绩公告，2026-10-22", URLS["release"]), ("MarketBeat 一致预期快照，2026-10-22", URLS["consensus"])], lang)
        bullet(doc, "收入与盈利均明显超预期。", "收入较财报前一致预期高 2.36 亿美元，Non-GAAP EPS 高 0.54 美元。58.1% 的毛利率比 EPS 更关键：它高出此前指引上限 210 个基点，显示提价、高容量产品组合与供给纪律正形成正反馈。", lang, compact=True)
        bullet(doc, "Q2 指引把盈利运行水平再推高一档。", "管理层给出 51 亿美元收入、60%-61% Non-GAAP 毛利率与 5.80 美元 Non-GAAP EPS 中值。目前市场预期已大体上修至这一范围，但新指引仍显著高于我们上一期 FY2027 预测所隐含的基础。", lang, compact=True)
        bullet(doc, "投资争议从需求转向执行。", "云业务仍占收入 88%，主要客户需求可见度已延伸至 2030-2031 年。核心问题转为 40TB ePMR 与 HAMR 爬坡能否把需求可见度转化为 EB 供给，同时控制资本强度、良率损失与成本通胀。", lang, compact=True)
        bullet(doc, "目标价上调至 620 美元，维持跑赢大盘。", "我们把 FY2027 收入与 Non-GAAP EPS 预测上调至 213 亿美元与 25.00 美元。以 24.8 倍市盈率计，目标价为 620 美元，较 10 月 23 日收盘价高约 80%。未来两年资本开支占收入比将略高于 4%-6% 常态区间，因此我们没有给予更高倍数。", lang, compact=True)


def page_revenue(doc, lang):
    new_page(doc)
    heading(doc, "Revenue analysis" if lang == "en" else "收入分析", lang)
    if lang == "en":
        paragraph(doc, "Revenue of $4.381bn increased 55% YoY and 17% sequentially, beating pre-release consensus of $4.145bn by $236m, or 5.7%. It also exceeded the top of management's prior $4.0bn-$4.2bn range by $181m. The variance was broad rather than dependent on one accounting item: exabytes shipped rose 22% YoY to 248EB, blended price per terabyte increased 28%, and the latest ePMR platform began contributing to the nearline mix. This combination matters because it shows that supply growth and pricing power can coexist when capacity gains lower customers' total cost of ownership.", lang)
        paragraph(doc, "Cloud revenue reached $3.868bn, up 54% YoY and equal to 88% of total revenue. The 10-Q attributes the gain to more than 20% exabyte growth and an approximately 25% increase in average selling price per exabyte. Client revenue rose 119% to $320m on roughly 15% exabyte growth and approximately 90% higher price per exabyte, while Consumer revenue increased 19% to $193m despite about 30% lower exabytes, as mix and price rose sharply. Consumer volume softness is the clearest sign of elasticity, but at 5% of revenue it does not alter the cloud-led thesis.", lang)
        paragraph(doc, "The quarter also strengthens the comparability of post-separation WDC. These figures represent the HDD-focused continuing business after the February 2025 Sandisk separation; flash results are not included. Investors should therefore compare the current five-quarter series with WDC's recast continuing-operations presentation rather than combine it with historical consolidated flash revenue. On that consistent basis, quarterly revenue has expanded every quarter from $2.818bn in Q1 FY2026 to $4.381bn in Q1 FY2027.", lang)
    else:
        paragraph(doc, "收入为 43.81 亿美元，同比增长 55%、环比增长 17%，较财报前 41.45 亿美元一致预期高 2.36 亿美元，即 5.7%；同时比此前 40 亿至 42 亿美元指引上限高 1.81 亿美元。超预期并非依赖单一会计项目：EB 出货同比增长 22% 至 248EB，每 TB 综合售价同比上升 28%，最新一代 ePMR 平台开始贡献近线产品结构。供给增长与定价提升可以同时出现，原因是更高容量降低了客户的总体拥有成本。", lang)
        paragraph(doc, "云业务收入达到 38.68 亿美元，同比增长 54%，占总收入 88%。10-Q 将增长归因于 EB 出货超过 20% 的增幅以及每 EB 平均售价约 25% 的提升。客户端收入增长 119% 至 3.20 亿美元，EB 约增 15%、每 EB 售价约增 90%；消费端收入增长 19% 至 1.93 亿美元，尽管 EB 下降约 30%，更高容量组合与售价仍推动收入增长。消费端销量疲弱反映价格弹性，但其仅占收入 5%，不会改变云业务主导的投资逻辑。", lang)
        paragraph(doc, "本季度也提高了分拆后 WDC 数据的可比性。这里的数据代表 2025 年 2 月 Sandisk 分拆后的 HDD 持续经营业务，不含闪存业务。分析时应采用 WDC 重述后的持续经营口径，而不能把当前 HDD 数据与历史合并闪存收入直接相加。在一致口径下，季度收入已由 FY2026 Q1 的 28.18 亿美元连续增长至 FY2027 Q1 的 43.81 亿美元。", lang)
    figure(doc, 1, "Quarterly revenue progression" if lang == "en" else "季度收入走势", lang, [("Q1 FY2027 presentation, slides 4-6", URLS["presentation"]), ("Q1 FY2027 Form 10-Q, filed 23 Oct. 2026", URLS["tenq_doc"])], width=5.6)


def page_profitability(doc, lang):
    new_page(doc)
    heading(doc, "Profitability and earnings quality" if lang == "en" else "盈利能力与盈利质量", lang)
    if lang == "en":
        paragraph(doc, "Non-GAAP gross margin reached 58.1%, up 370 bps sequentially and 1,420 bps YoY. Non-GAAP operating margin expanded to 49.1%, up 490 bps sequentially and 1,870 bps YoY. The margin bridge is unusually powerful: higher-capacity products lift revenue per drive, favorable contract and spot pricing increases price per terabyte, and areal-density gains add exabytes without equivalent unit-capacity spending. Management said incremental YoY gross margin exceeded 80%, a result that makes pricing sustainability and transition execution more important than simple unit growth.", lang)
        paragraph(doc, "Non-GAAP operating expenses were $391m, only 3% above the prior year despite 55% revenue growth. That restraint produced $2.153bn of non-GAAP operating income, up 152% YoY. GAAP operating income was $2.046bn after $70m of stock compensation and $27m of business realignment charges. The narrow $0.01 difference between GAAP and non-GAAP diluted EPS this quarter contrasts with earlier quarters distorted by gains on the retained Sandisk interest and debt-equity transactions; as a result, Q1's $4.63 GAAP EPS and $4.64 non-GAAP EPS provide a cleaner view of current operating economics.", lang)
        paragraph(doc, "The principal near-term margin risk is not demand but the cost of supplying it. Management expects DRAM and NAND input inflation to limit some cost-per-terabyte improvement, while HAMR heads require more production cycles. Nevertheless, Q2 guidance of 60%-61% gross margin implies another 190-290 bps sequential expansion. We view that guide as credible but not automatically repeatable: continued LTA repricing, spot-market discipline and the 40TB ePMR ramp must offset transition costs and richer component inputs.", lang)
    else:
        paragraph(doc, "Non-GAAP 毛利率达到 58.1%，环比提升 370 个基点、同比提升 1,420 个基点；Non-GAAP 营业利润率达到 49.1%，环比提升 490 个基点、同比提升 1,870 个基点。利润率提升由三方面共同驱动：高容量产品提高单盘收入，合同与现货定价改善每 TB 售价，面密度提升则无需等比例增加整盘产能即可增加 EB。管理层称同比增量毛利率超过 80%，因此定价持续性与技术切换执行比单纯销量增长更关键。", lang)
        paragraph(doc, "Non-GAAP 营业费用为 3.91 亿美元，在收入增长 55% 的情况下仅同比增加 3%，推动 Non-GAAP 营业利润增长 152% 至 21.53 亿美元。扣除 0.70 亿美元股权薪酬与 0.27 亿美元业务调整费用后，GAAP 营业利润为 20.46 亿美元。本季度 GAAP 与 Non-GAAP 摊薄 EPS 仅相差 0.01 美元，远小于此前受 Sandisk 保留权益重估收益及债务股权交易影响的季度，因此 4.63 美元 GAAP EPS 与 4.64 美元 Non-GAAP EPS 更能反映当前经营质量。", lang)
        paragraph(doc, "近期利润率风险主要不是需求，而是满足需求的成本。管理层预计 DRAM 与 NAND 投入品通胀将限制部分每 TB 成本改善，同时 HAMR 磁头需要更多生产循环。不过，Q2 60%-61% 毛利率指引仍意味着环比再提升 190 至 290 个基点。我们认为该指引具备可实现性，但不会自动延续：LTA 重定价、现货市场纪律与 40TB ePMR 爬坡必须抵消技术切换成本和零部件涨价。", lang)
    figure(doc, 2, "GAAP and non-GAAP diluted EPS" if lang == "en" else "GAAP 与 Non-GAAP 摊薄 EPS", lang, [("Q1 FY2027 presentation, slides 6 and 10", URLS["presentation"]), ("Q1 FY2027 earnings release", URLS["release_pdf"])], width=4.95)
    figure(doc, 3, "Non-GAAP margin progression" if lang == "en" else "Non-GAAP 利润率走势", lang, [("Q1 FY2027 presentation, slides 5-6", URLS["presentation"]), ("Q1 FY2027 Form 10-Q", URLS["tenq_doc"])], width=4.95)


def page_cash_mix(doc, lang):
    new_page(doc)
    heading(doc, "Cash generation and end-market mix" if lang == "en" else "现金流与终端市场结构", lang)
    if lang == "en":
        paragraph(doc, "Operating cash flow rose 139% YoY to $1.604bn and free cash flow increased 130% to $1.377bn, equal to 31.4% of revenue. Capital expenditures were $227m, up from $73m a year earlier as WDC invested in production technology and automation. Working capital absorbed $331m, principally through a $269m inventory build and $206m increase in receivables, partly offset by payables and tax timing. The cash conversion is therefore strong despite deliberate inventory preparation for higher demand rather than because working capital temporarily flattered the result.", lang)
        paragraph(doc, "Capital allocation was substantial: WDC repurchased $1.001bn of shares, paid $56m of common dividends and used $570m to settle convertible-note transactions. Cash ended at $1.314bn, while the revolver balance was $650m and only $141m of convertible principal remained. The resulting roughly $0.5bn net cash position supports the raised $0.175 quarterly dividend and targeted investment, but conversion-related share issuance and equity compensation complicate per-share value creation. We therefore focus on diluted shares and free cash flow per share, not only headline repurchases.", lang)
        paragraph(doc, "Revenue remains concentrated in Cloud at 88%, with Client at 7% and Consumer at 5%. Cloud concentration creates customer and negotiating risk, but the mix is economically attractive while hyperscaler and AI workloads sustain high-capacity demand. The 248EB shipment total comprised 223EB of nearline and 25EB of non-nearline products. Nearline grew from 183EB a year ago, while non-nearline moved from 21EB, confirming that the volume engine remains enterprise and cloud storage rather than PC replacement.", lang)
    else:
        paragraph(doc, "经营现金流同比增长 139% 至 16.04 亿美元，自由现金流增长 130% 至 13.77 亿美元，相当于收入的 31.4%。资本开支为 2.27 亿美元，高于上年同期的 0.73 亿美元，反映生产技术与自动化投入。营运资本消耗 3.31 亿美元，主要来自 2.69 亿美元库存增加与 2.06 亿美元应收增加，部分被应付与税务时点抵消。因此，现金转化在主动备货满足更高需求的同时仍然强劲，并非依靠营运资本暂时性释放。", lang)
        paragraph(doc, "资本配置规模较大：公司回购 10.01 亿美元股票，支付 0.56 亿美元普通股股息，并以 5.70 亿美元结算可转债交易。期末现金为 13.14 亿美元，循环信贷余额 6.50 亿美元，可转债本金仅剩 1.41 亿美元。约 5 亿美元净现金支持季度股息上调至每股 0.175 美元与定向投资，但可转债相关增发及股权薪酬会影响每股价值创造，因此我们更关注摊薄股数与每股自由现金流，而不仅是名义回购金额。", lang)
        paragraph(doc, "云业务仍占收入 88%，客户端占 7%，消费端占 5%。云业务集中度带来客户与议价风险，但在超大规模云厂商和 AI 工作负载持续拉动高容量需求时，这一结构具备较高经济价值。248EB 总出货由 223EB 近线与 25EB 非近线构成，分别高于上年同期的 183EB 与 21EB，说明核心销量引擎仍是企业与云存储，而非 PC 换机。", lang)
    figure(doc, 4, "Quarterly free cash flow" if lang == "en" else "季度自由现金流", lang, [("Q1 FY2027 presentation, slides 6 and 10", URLS["presentation"]), ("Q1 FY2027 Form 10-Q, cash-flow statement", URLS["tenq_doc"])], width=4.75)
    figure(doc, 5, "Revenue mix by end market" if lang == "en" else "按终端市场划分的收入结构", lang, [("Q1 FY2027 presentation, slide 4", URLS["presentation"]), ("Q1 FY2027 Form 10-Q, revenue note", URLS["tenq_doc"])], width=4.75)


def page_guidance(doc, lang):
    new_page(doc)
    heading(doc, "Guidance and near-term outlook" if lang == "en" else "指引与近期展望", lang)
    if lang == "en":
        paragraph(doc, "Management guided Q2 FY2027 revenue to $5.1bn plus or minus $100m, representing 66%-72% YoY growth and about 16% sequential growth at the midpoint. Non-GAAP gross margin is expected at 60%-61%, operating expenses at $400m-$410m, other expense near $15m, tax near 17% and diluted shares near 380m. The resulting $5.80 plus or minus $0.20 non-GAAP EPS midpoint implies another 25% sequential increase. Compared with current 23 October consensus of $5.127bn revenue and $5.87 EPS, the guide is broadly aligned after post-release revisions rather than an additional beat still unrecognized by the Street.", lang)
        paragraph(doc, "The comparison with prior guidance is more revealing. Q1 actual revenue of $4.381bn exceeded the old midpoint by 6.9%, actual gross margin exceeded its midpoint by 260 bps and non-GAAP EPS exceeded its midpoint by 16%. Management has consistently used technology transitions, long-term agreements and constrained unit supply to improve price per terabyte; the Q2 margin guide assumes that mechanism persists. The next validation points are the share of nearline shipments on the newest ePMR platform, price resets as LTAs renew and whether cost inflation remains contained.", lang)
        paragraph(doc, "We do not extrapolate Q2's 69% YoY revenue growth indefinitely. The base case assumes Q3 and Q4 growth moderates as comparisons become harder, while higher-capacity mix supports gross margin above 60%. Management's comments that customer visibility reaches 2030-2031 are strategically positive but do not constitute firm volume or price guidance. The key analytical distinction is between demand visibility and contracted economics: some LTA volumes have pricing provisions, some exabytes remain uncontracted, and coverage declines farther into the future.", lang)
    else:
        paragraph(doc, "管理层预计 FY2027 Q2 收入为 51 亿美元，上下浮动 1 亿美元，即同比增长 66%-72%、中值环比增长约 16%。Non-GAAP 毛利率预计为 60%-61%，营业费用 4.00 亿至 4.10 亿美元，其他净费用约 0.15 亿美元，税率约 17%，摊薄股数约 3.80 亿股；对应 Non-GAAP EPS 为 5.80 美元，上下浮动 0.20 美元，中值环比再增长 25%。与 10 月 23 日当前市场预期的 51.27 亿美元收入和 5.87 美元 EPS 相比，新指引在财报后上修后已大体被市场吸收。", lang)
        paragraph(doc, "与此前指引的比较更有信息量。Q1 实际收入较旧指引中值高 6.9%，实际毛利率高 260 个基点，Non-GAAP EPS 高 16%。公司持续利用技术切换、长期协议与整盘供给约束来提高每 TB 售价，Q2 利润率指引假设这一机制延续。下一步验证点包括最新 ePMR 平台占近线出货比例、LTA 续签时的定价调整，以及成本通胀能否受控。", lang)
        paragraph(doc, "我们不会把 Q2 收入同比增长 69% 机械外推。基准情景假设 Q3 与 Q4 因基数抬高而增速放缓，但高容量组合支撑毛利率保持在 60% 以上。管理层称客户需求可见度已延伸至 2030-2031 年，这对战略有利，却不等同于确定的量价指引。分析上必须区分需求可见度与合同经济性：部分 LTA 数量带有定价条款，部分 EB 尚未签约，而且期限越远，合同覆盖越低。", lang)
    figure(doc, 7, "Q1 actual results vs. prior guidance and consensus" if lang == "en" else "Q1 实际业绩与此前指引及一致预期", lang, [("Q4 FY2026 release and Q1 guidance, 5 Aug. 2026", URLS["q4_release"]), ("Q1 FY2027 release, 22 Oct. 2026", URLS["release"]), ("MarketBeat pre-release consensus", URLS["consensus"])], width=4.8)
    figure(doc, 8, "Q2 FY2027 guidance vs. current Street" if lang == "en" else "FY2027 Q2 指引与当前市场预期", lang, [("Q1 FY2027 presentation, slide 7", URLS["presentation"]), ("Longbridge consensus snapshot, 23 Oct. 2026", URLS["longbridge"])], width=4.8)


def page_thesis(doc, lang):
    new_page(doc)
    heading(doc, "Updated investment thesis" if lang == "en" else "更新后的投资逻辑", lang)
    if lang == "en":
        bullet(doc, "AI storage demand and customer visibility  Strengthened.", "Cloud revenue grew 54% and remained 88% of sales, while management described demand signals from hyperscalers, AI frontier labs, NeoCloud operators, enterprise users and hybrid-cloud OEMs. Visibility for major customers now extends into 2030-2031. The breadth reduces dependence on any single AI training cycle, although customer concentration remains high and future visibility is not equivalent to firm take-or-pay volume.", lang)
        bullet(doc, "Technology-led capacity growth  Strengthened but execution-sensitive.", "The newest ePMR platform, delivering up to 40TB, is transitioning faster than previous platforms. Management expects it to represent about 30% of nearline shipments in Q2 and 50% in Q3 FY2027. HAMR volume shipments remain planned for the first half of calendar 2027. If both ramps execute, WDC can accelerate exabyte growth through FY2027-FY2028 while holding unit capacity relatively disciplined.", lang)
        bullet(doc, "Pricing and margin durability  Strengthened.", "Blended price per terabyte increased 28% YoY, and Q2 gross-margin guidance moves to 60%-61%. Pricing is supported by contractual adjustments on LTAs, higher capacity's customer TCO benefits and tight spot supply for uncontracted exabytes. The strongest challenge is that pricing visibility is shorter than demand visibility; contracts cover only part of future output, and aggressive industry supply additions could eventually erode the present scarcity premium.", lang)
        bullet(doc, "Capital allocation and per-share compounding  Mixed.", "Free cash flow supports dividends, repurchases and technology investment, but the company will spend modestly above its normal 4%-6% capex-to-revenue range for two years. Convertible-note settlements also issued shares against conversion premiums, while stock compensation continues. The correct scorecard is growth in free cash flow per diluted share after necessary transition investment, not gross buyback dollars.", lang)
        paragraph(doc, "Overall, the earnings release strengthens the operating thesis more than the stock reaction implies. The 2.97% decline on 23 October appears to reflect fear that elevated capex and input costs mark a cyclical peak. We instead read the call as a deliberate shift toward heads, media and automation needed to lift areal density, not a broad unit-capacity build. That distinction is essential: spending that supports more platters and higher-capacity drives within disciplined unit volumes can preserve pricing, while traditional factory expansion would create a more serious oversupply risk.", lang)
    else:
        bullet(doc, "AI 存储需求与客户可见度  逻辑增强。", "云业务收入增长 54%，占比维持 88%；管理层提到需求来自超大规模云厂商、AI 前沿实验室、NeoCloud、企业客户与混合云 OEM。主要客户需求可见度延伸至 2030-2031 年，覆盖面扩大降低了对单一 AI 训练周期的依赖，但客户集中度仍高，长期可见度也不等于刚性采购量。", lang)
        bullet(doc, "技术驱动的容量增长  逻辑增强但依赖执行。", "最高 40TB 的最新 ePMR 平台切换速度快于以往，管理层预计其占近线出货比例在 FY2027 Q2 约为 30%、Q3 达 50%。HAMR 量产出货仍计划于 2027 年上半年开始。若两项切换顺利，WDC 可在维持整盘供给纪律的同时，于 FY2027-FY2028 加速 EB 增长。", lang)
        bullet(doc, "定价与利润率持续性  逻辑增强。", "每 TB 综合售价同比增长 28%，Q2 毛利率指引升至 60%-61%。定价由 LTA 合同调价、高容量产品带来的客户总体拥有成本优势，以及未签约 EB 的紧张现货供给共同支撑。主要挑战是定价可见度短于需求可见度：合同仅覆盖部分未来产量，行业若激进扩产，当前稀缺溢价最终可能回落。", lang)
        bullet(doc, "资本配置与每股复利  表现中性。", "自由现金流可支持股息、回购与技术投入，但未来两年资本开支占收入比将略高于 4%-6% 常态范围。可转债结算也因转换溢价发行股份，股权薪酬持续存在。正确评价指标应是完成必要技术投入后的摊薄每股自由现金流增长，而非名义回购总额。", lang)
        paragraph(doc, "整体看，本次财报对经营逻辑的强化程度高于股价反应。10 月 23 日下跌 2.97%，更像是市场担忧资本开支与投入成本上升代表周期见顶。我们认为电话会传递的是面向磁头、介质与自动化的密度升级，而非大规模增加整盘产能。二者差别至关重要：在整盘量受控时，通过更多盘片和更高容量提高 EB 的投资可维持价格纪律；传统工厂扩张则会带来更严重的供给过剩风险。", lang)
    figure(doc, 6, "Nearline and non-nearline exabytes shipped" if lang == "en" else "近线与非近线 EB 出货", lang, [("Q1 FY2027 presentation, slide 5", URLS["presentation"]), ("Q1 FY2027 earnings-call transcript, 22 Oct. 2026", URLS["transcript"])], width=5.2)


def page_estimates(doc, lang):
    new_page(doc)
    heading(doc, "Estimate revisions" if lang == "en" else "预测调整", lang)
    if lang == "en":
        paragraph(doc, "We raise FY2027 revenue to $21.3bn from $19.6bn and non-GAAP EPS to $25.00 from $20.50. The revenue change incorporates Q1's $236m consensus beat, the $5.1bn Q2 guide midpoint and a higher second-half baseline supported by the 40TB ePMR ramp. The EPS increase is larger because gross-margin leverage and restrained operating expenses convert incremental revenue at high contribution margins. Our FY2028 view of $25.0bn revenue and $31.00 EPS assumes exabyte supply acceleration, HAMR qualification and continued pricing discipline, but not indefinite repetition of Q2's 69% growth rate.", lang)
        make_table(
            doc,
            ["Metric", "FY2026A", "FY2027E prior", "FY2027E new", "Change", "FY2028E"],
            [
                ["Revenue ($bn)", "12.9", "19.6", "21.3", "+8.7%", "25.0"],
                ["Non-GAAP gross margin", "49.1%", "57.5%", "60.5%", "+300 bps", "61.5%"],
                ["Non-GAAP op. margin", "37.3%", "47.0%", "51.0%", "+400 bps", "52.0%"],
                ["Non-GAAP EPS", "$10.22", "$20.50", "$25.00", "+22.0%", "$31.00"],
                ["Free cash flow ($bn)", "3.5", "5.3", "5.9", "+11.3%", "7.0"],
            ],
            lang,
            [1.55, 1.0, 1.18, 1.18, 0.92, 1.0],
            variance_col=4,
            font_size=7.45,
        )
        source(doc, [("WDC FY2026 Form 10-K", "https://www.sec.gov/Archives/edgar/data/106040/000162828026057139/wdc-20260703.htm"), ("Q1 FY2027 release and Q2 guidance", URLS["release"]), ("Independent estimates, 24 Oct. 2026", URLS["ir"])], lang)
        paragraph(doc, "The prior FY2027 baseline is the estimate published in our Q4 FY2026 update rather than a retrospective Street number. The new FY2027 path is intentionally close to current quarterly consensus on Q2-Q4, but it uses a slightly more conservative share-count and cost framework to reflect dilution and transition spending. Our forecast is not company guidance beyond Q2. Quarterly timing can move with customer deployment schedules, and long-term agreements may shift volume between periods without changing full-year demand.", lang)
        paragraph(doc, "Sensitivity is highest to gross margin. A 100 bp change in FY2027 gross margin, holding revenue and other assumptions constant, changes annual pre-tax profit by roughly $213m and EPS by approximately $0.46 after tax. Revenue upside without pricing discipline would therefore be less valuable than the same revenue delivered through higher-capacity products and controlled unit supply. Conversely, sustained gross margin above 61% would make our $25.00 EPS estimate conservative even if exabyte growth merely tracks the current roadmap.", lang)
    else:
        paragraph(doc, "我们把 FY2027 收入预测由 196 亿美元上调至 213 亿美元，把 Non-GAAP EPS 由 20.50 美元上调至 25.00 美元。收入调整纳入 Q1 较一致预期高 2.36 亿美元、Q2 51 亿美元指引中值，以及 40TB ePMR 爬坡带来的更高下半年基础。EPS 上调幅度更大，因为毛利率杠杆与受控营业费用使增量收入以较高贡献率转化为利润。FY2028 预计收入 250 亿美元、EPS 31.00 美元，假设 EB 供给加速、HAMR 通过认证及定价纪律延续，但不假设 Q2 69% 的同比增速长期持续。", lang)
        make_table(
            doc,
            ["指标", "FY2026A", "FY2027E 此前", "FY2027E 更新", "变化", "FY2028E"],
            [
                ["收入（十亿美元）", "12.9", "19.6", "21.3", "+8.7%", "25.0"],
                ["Non-GAAP 毛利率", "49.1%", "57.5%", "60.5%", "+300 基点", "61.5%"],
                ["Non-GAAP 营业利润率", "37.3%", "47.0%", "51.0%", "+400 基点", "52.0%"],
                ["Non-GAAP EPS", "10.22 美元", "20.50 美元", "25.00 美元", "+22.0%", "31.00 美元"],
                ["自由现金流（十亿美元）", "3.5", "5.3", "5.9", "+11.3%", "7.0"],
            ],
            lang,
            [1.65, 0.95, 1.18, 1.18, 0.92, 1.0],
            variance_col=4,
            font_size=7.35,
        )
        source(doc, [("WDC FY2026 Form 10-K", "https://www.sec.gov/Archives/edgar/data/106040/000162828026057139/wdc-20260703.htm"), ("FY2027 Q1 业绩公告与 Q2 指引", URLS["release"]), ("独立预测，2026-10-24", URLS["ir"])], lang)
        paragraph(doc, "此前 FY2027 基线来自我们上一期 FY2026 Q4 报告，而非事后重构的市场预期。新的 FY2027 路径与当前 Q2-Q4 季度一致预期大体接近，但对摊薄股数与技术切换成本采用略谨慎假设。本预测不代表公司 Q2 之后的指引。季度时点可能随客户部署计划变化，LTA 也可能在季度之间移动数量而不改变全年需求。", lang)
        paragraph(doc, "预测对毛利率最敏感。在收入及其他假设不变时，FY2027 毛利率每变化 100 个基点，全年税前利润约变化 2.13 亿美元，税后 EPS 约变化 0.46 美元。因此，没有定价纪律的收入上行，其价值低于通过高容量产品与受控整盘供给实现的同等收入。反之，若毛利率持续高于 61%，即使 EB 增速仅按现有路线图推进，我们 25.00 美元 EPS 预测也可能偏保守。", lang)
    figure(doc, 9, "FY2027 estimate revisions" if lang == "en" else "FY2027 预测调整", lang, [("Prior Q4 FY2026 independent estimates", URLS["q4_release"]), ("Q1 FY2027 results and current consensus", URLS["release"]), ("Longbridge consensus, 23 Oct. 2026", URLS["longbridge"])], width=5.25)


def page_valuation(doc, lang, mkt):
    new_page(doc)
    heading(doc, "Valuation and price target" if lang == "en" else "估值与目标价", lang)
    upside = 620 / mkt["price"] - 1
    if lang == "en":
        paragraph(doc, f"We raise our price target to $620 from $600, implying {upside:.0%} upside from the 23 October close of ${mkt['price']:.2f}. The target applies 24.8x to our FY2027 non-GAAP EPS estimate of $25.00. That multiple is above WDC's current roughly 13x trailing earnings multiple but below the valuation implied by the strongest AI infrastructure peers and below Seagate's current premium multiple. The premium to WDC's own current multiple is justified by more than 50% revenue growth, expanding margins and a net-cash balance sheet; the discount to premium peers reflects product-transition, customer-concentration and cycle risk.", lang)
        make_table(
            doc,
            ["Scenario", "FY2027 EPS", "P/E", "Value", "Key assumption"],
            [
                ["Bear", "$21.00", "20.0x", "$420", "Pricing normalizes; HAMR ramp slips"],
                ["Base", "$25.00", "24.8x", "$620", "ePMR ramp executes; GM near 60%"],
                ["Bull", "$28.60", "28.0x", "$800", "Supply remains tight; HAMR scales cleanly"],
            ],
            lang,
            [0.9, 1.0, 0.8, 0.8, 2.7],
            font_size=7.5,
        )
        source(doc, [("Independent scenario analysis, 24 Oct. 2026", URLS["ir"]), ("Longbridge price and valuation snapshot, 23 Oct. 2026", URLS["longbridge"])], lang)
        paragraph(doc, "The bear case still sits above the current price because the stock already discounts a sharp margin and multiple reversal. Its $420 value assumes EPS settles near $21 and the market pays only 20x. The bull case requires both sustained scarcity pricing and a clean HAMR ramp, supporting $28.60 of EPS and 28x. We do not probability-weight these endpoints; they are a sensitivity framework showing that technology execution and pricing duration, not terminal revenue growth alone, drive valuation dispersion.", lang)
        paragraph(doc, "Cross-checking against external sentiment, Longbridge's 23 October institutional snapshot showed 17 strong-buy, four buy and five hold ratings, with a $610 average target, $405 low and $900 high. Our $620 target is therefore close to the updated mean rather than an outlier. The main reason to retain Outperform after the post-earnings decline is asymmetric risk-reward: the base case offers substantial upside, while the bear case requires a material deterioration from the current operating trajectory before falling toward the present market price.", lang)
    else:
        paragraph(doc, f"我们把目标价由 600 美元上调至 620 美元，较 10 月 23 日 {mkt['price']:.2f} 美元收盘价具备约 {upside:.0%} 上行空间。目标价以 FY2027 Non-GAAP EPS 25.00 美元乘以 24.8 倍市盈率得到。该倍数高于 WDC 当前约 13 倍追踪市盈率，但低于最强 AI 基础设施同业所隐含的估值，也低于希捷当前溢价倍数。超过 WDC 当前倍数的理由是收入增长超过 50%、利润率扩张及净现金资产负债表；相对高估值同业折价则反映技术切换、客户集中与周期风险。", lang)
        make_table(
            doc,
            ["情景", "FY2027 EPS", "市盈率", "估值", "核心假设"],
            [
                ["悲观", "21.00 美元", "20.0 倍", "420 美元", "定价正常化，HAMR 爬坡推迟"],
                ["基准", "25.00 美元", "24.8 倍", "620 美元", "ePMR 执行顺利，毛利率约 60%"],
                ["乐观", "28.60 美元", "28.0 倍", "800 美元", "供给持续偏紧，HAMR 顺利放量"],
            ],
            lang,
            [0.9, 1.05, 0.8, 0.85, 2.6],
            font_size=7.4,
        )
        source(doc, [("独立情景分析，2026-10-24", URLS["ir"]), ("Longbridge 股价与估值快照，2026-10-23", URLS["longbridge"])], lang)
        paragraph(doc, "悲观情景仍高于当前股价，因为市场价格已反映利润率与估值倍数大幅回落。420 美元估值假设 EPS 回落至约 21 美元，市场仅给予 20 倍。乐观情景需要稀缺定价持续且 HAMR 顺利爬坡，对应 28.60 美元 EPS 与 28 倍估值。我们不对这些端点做概率加权；它们用于显示技术执行与定价持续时间，而非单纯终值收入增速，是估值离散度的核心来源。", lang)
        paragraph(doc, "外部情绪交叉验证方面，Longbridge 10 月 23 日机构快照显示 17 个强力买入、4 个买入与 5 个持有评级，平均目标价 610 美元，区间 405 至 900 美元。我们的 620 美元接近最新均值，并非异常乐观。财报后下跌后仍维持跑赢大盘，关键在于风险收益不对称：基准情景上行空间充足，而悲观情景需要当前经营趋势明显恶化，估值才会接近现价。", lang)
    figure(doc, 10, "Bear base and bull valuation scenarios" if lang == "en" else "悲观 基准与乐观估值情景", lang, [("Independent estimates and scenario analysis, 24 Oct. 2026", URLS["ir"]), ("Longbridge 23 Oct. close", URLS["longbridge"])], width=5.35)


def page_risks(doc, lang, mkt):
    new_page(doc)
    heading(doc, "Market reaction risks and monitoring" if lang == "en" else "市场反应 风险与跟踪要点", lang)
    if lang == "en":
        paragraph(doc, f"WDC fell {abs(mkt['one_day_change']):.2%} on 23 October to ${mkt['price']:.2f} despite the double beat and strong Q2 guide. The reaction suggests investors are looking past the near-term numbers toward the cost of sustaining supply growth. Management expects capex to run modestly above the long-run 4%-6% of revenue range for two years, primarily for heads, media technology and factory automation. The distinction from adding unit capacity is favorable, but it still creates execution risk and reduces near-term free-cash-flow conversion relative to the exceptional Q1 level.", lang)
        bullet(doc, "Technology transition risk.", "ePMR must reach about 30% of nearline shipments in Q2 and 50% in Q3, while HAMR is scheduled for first-half 2027 volume shipments. Qualification delays, yield issues or throughput loss would limit exabyte growth and pressure margins.", lang, compact=True)
        bullet(doc, "Pricing and supply risk.", "Current margins depend on LTA price steps, product value and tight uncontracted supply. A competitor-led capacity response, weaker cloud deployments or LTA renegotiation could compress price per terabyte faster than density lowers cost.", lang, compact=True)
        bullet(doc, "Customer and product concentration.", "Cloud is 88% of revenue and large customer deployment schedules can shift regional and quarterly results. Client and Consumer are too small to offset a broad hyperscaler pause.", lang, compact=True)
        bullet(doc, "Capital structure and dilution.", "Convertible settlements simplified debt but used shares for conversion premiums. Repurchases should be evaluated net of dilution, equity compensation and cash needed for technology investment.", lang, compact=True)
        bullet(doc, "Input-cost and macro risk.", "Management highlighted DRAM and NAND inflation. Trade restrictions, supplier concentration, foreign exchange and customer capex cycles remain material sensitivities in the 10-Q.", lang, compact=True)
        paragraph(doc, "Our monitoring dashboard for the next report has five items: 40TB ePMR shipment mix, total and nearline exabyte growth, blended price per terabyte, gross margin relative to 60%-61% guidance and capex/free-cash-flow conversion. A clean beat on these measures would support further estimate upside. A revenue beat driven only by pricing, without progress on exabyte supply and technology mix, would be lower quality because it would increase the risk of customer pushback and future share loss.", lang)
    else:
        paragraph(doc, f"尽管业绩与 Q2 指引双双强劲，WDC 仍在 10 月 23 日下跌 {abs(mkt['one_day_change']):.2%} 至 {mkt['price']:.2f} 美元。市场反应显示投资者已把注意力从近期数字转向维持供给增长的成本。管理层预计未来两年资本开支占收入比略高于 4%-6% 长期常态，主要用于磁头、介质技术与工厂自动化。与新增整盘产能相比，这种投入更有利于价格纪律，但仍带来执行风险，并使近期自由现金流转化率低于 Q1 的异常高水平。", lang)
        bullet(doc, "技术切换风险。", "ePMR 需要在 Q2 达到约 30% 的近线出货占比、Q3 达到 50%，HAMR 则计划于 2027 年上半年量产出货。认证延期、良率问题或吞吐损失都会限制 EB 增长并压低利润率。", lang, compact=True)
        bullet(doc, "定价与供给风险。", "当前利润率依赖 LTA 调价、产品价值和未签约供给紧张。竞争对手扩产、云部署放缓或 LTA 重新谈判，均可能使每 TB 售价下降速度快于密度降低成本的速度。", lang, compact=True)
        bullet(doc, "客户与产品集中风险。", "云业务占收入 88%，大客户部署计划可改变区域与季度结果；客户端和消费端规模不足以抵消超大规模客户普遍暂停。", lang, compact=True)
        bullet(doc, "资本结构与摊薄风险。", "可转债结算简化了债务，但转换溢价通过发行股份支付。评估回购时应扣除摊薄、股权薪酬和技术投入所需现金。", lang, compact=True)
        bullet(doc, "投入成本与宏观风险。", "管理层明确提及 DRAM 与 NAND 通胀；贸易限制、供应商集中、汇率及客户资本开支周期仍是 10-Q 中的重要敏感项。", lang, compact=True)
        paragraph(doc, "下一期报告的跟踪看板包括五项：40TB ePMR 出货占比、总 EB 与近线 EB 增速、每 TB 综合售价、相对 60%-61% 指引的毛利率，以及资本开支与自由现金流转化。若这些指标同步超预期，预测仍有上调空间；若收入超预期仅来自提价，而 EB 供给和技术结构没有进展，则质量较低，因为客户抵触与未来份额流失风险会升高。", lang)
    figure(doc, 11, "Adjusted one-year share price and target" if lang == "en" else "前复权一年股价与目标价", lang, [("Longbridge adjusted daily K-line through 23 Oct. 2026", URLS["longbridge"]), ("WDC Q1 FY2027 release date", URLS["release"])], width=5.35)


def page_sources(doc, lang):
    new_page(doc)
    heading(doc, "Sources methodology and disclosures" if lang == "en" else "资料来源 方法与披露", lang)
    if lang == "en":
        paragraph(doc, "Today is 24 October 2026. Western Digital Corporation trades on Nasdaq under WDC and files with SEC CIK 106040. The company released Q1 FY2027 results on 22 October 2026 for the period ended 2 October 2026, furnished Form 8-K the same day and filed Form 10-Q on 23 October. The earnings call was held on 22 October at 4:30 p.m. Eastern, matching the release date. These are the latest reported results and fall within the required three-month recency window.", lang)
        paragraph(doc, "Reported financial statements, end-market revenue, cash flow, debt, share count and risk disclosures are anchored to the earnings release, presentation and Form 10-Q. Management commentary and Q&A use the dated 22 October transcript and WDC webcast page. Pre-release consensus for the beat/miss calculation uses the MarketBeat 22 October snapshot; current forward consensus, rating and target-price context use Longbridge data dated 23 October. The stock-price chart uses Longbridge forward-adjusted daily closes through 23 October.", lang)
        paragraph(doc, "Western Digital completed the separation of Sandisk on 21 February 2025. Historical quarterly series in this report use WDC's recast continuing-operations HDD presentation and do not combine post-separation WDC with Sandisk flash revenue. Non-GAAP measures follow WDC's definitions and are not substitutes for GAAP. Prior FY2027 estimates come from the earlier Q4 FY2026 update; updated FY2027-FY2028 estimates, scenarios, rating and target are independent analytical judgments rather than company guidance.", lang)
        paragraph(doc, "This report is independent research for informational purposes only. It is not investment advice, an offer, a solicitation or a fiduciary recommendation. Estimates and valuation scenarios depend on assumptions that may prove wrong. Readers should review the primary filings and consider their objectives, risk tolerance and professional advice before making investment decisions.", lang)
        refs = [
            ("Q1 FY2027 earnings release  22 October 2026", URLS["release"]),
            ("Q1 FY2027 earnings presentation  22 October 2026", URLS["presentation"]),
            ("Form 8-K accession 0001628280-26-066882  filed 22 October 2026", URLS["eightk"]),
            ("Form 10-Q accession 0001628280-26-066933  filed 23 October 2026", URLS["tenq"]),
            ("Q1 FY2027 earnings call and webcast  22 October 2026", URLS["webcast"]),
            ("Q1 FY2027 call transcript  22 October 2026", URLS["transcript"]),
            ("Q4 FY2026 release and prior Q1 guidance  5 August 2026", URLS["q4_release"]),
            ("MarketBeat pre-release consensus snapshot  22 October 2026", URLS["consensus"]),
            ("Longbridge market consensus rating valuation and price data  23 October 2026", URLS["longbridge"]),
            ("WDC earnings documents archive", URLS["ir"]),
        ]
    else:
        paragraph(doc, "当前日期为 2026 年 10 月 24 日。Western Digital Corporation 在纳斯达克以 WDC 交易，SEC CIK 为 106040。公司于 2026 年 10 月 22 日发布截至 10 月 2 日的 FY2027 Q1 业绩，同日提交 Form 8-K，并于 10 月 23 日提交 Form 10-Q。电话会于 10 月 22 日美东时间下午 4:30 举行，与公告日期一致。这是最新一期已披露业绩，并满足三个月时效要求。", lang)
        paragraph(doc, "财务报表、终端市场收入、现金流、债务、股数与风险披露以业绩公告、演示材料和 10-Q 为准。管理层评论与问答使用 10 月 22 日电话会文字稿及 WDC 网络回放页。超预期分析采用 MarketBeat 10 月 22 日财报前一致预期快照；前瞻一致预期、评级与目标价背景采用 10 月 23 日 Longbridge 数据。股价图使用截至 10 月 23 日的 Longbridge 前复权日线收盘价。", lang)
        paragraph(doc, "Western Digital 于 2025 年 2 月 21 日完成 Sandisk 分拆。本报告历史季度序列使用 WDC 重述后的 HDD 持续经营口径，不把分拆后的 WDC 与 Sandisk 闪存收入合并。Non-GAAP 指标遵循 WDC 定义，不能替代 GAAP。此前 FY2027 预测来自上一期 FY2026 Q4 更新；新的 FY2027-FY2028 预测、情景、评级和目标价属于独立分析判断，并非公司指引。", lang)
        paragraph(doc, "本报告为独立研究，仅供信息参考，不构成投资建议、要约、招揽或受托建议。预测与估值情景依赖可能不成立的假设。投资者在决策前应阅读原始文件，并结合自身目标、风险承受能力与专业意见。", lang)
        refs = [
            ("FY2027 Q1 业绩公告  2026-10-22", URLS["release"]),
            ("FY2027 Q1 业绩演示  2026-10-22", URLS["presentation"]),
            ("Form 8-K  0001628280-26-066882  2026-10-22", URLS["eightk"]),
            ("Form 10-Q  0001628280-26-066933  2026-10-23", URLS["tenq"]),
            ("FY2027 Q1 电话会与回放  2026-10-22", URLS["webcast"]),
            ("FY2027 Q1 电话会文字稿  2026-10-22", URLS["transcript"]),
            ("FY2026 Q4 公告与此前 Q1 指引  2026-08-05", URLS["q4_release"]),
            ("MarketBeat 财报前一致预期快照  2026-10-22", URLS["consensus"]),
            ("Longbridge 市场预期 评级 估值与股价数据  2026-10-23", URLS["longbridge"]),
            ("WDC 财报文件档案", URLS["ir"]),
        ]
    heading(doc, "Sources and references" if lang == "en" else "资料来源与参考文件", lang, 2)
    for label, url in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2.0)
        r = p.add_run("• ")
        set_run_font(r, lang, size=8.0)
        add_hyperlink(p, label, url, lang, size=8.0)


def build(lang, mkt):
    doc = setup_doc(lang)
    title_page(doc, lang, mkt)
    page_revenue(doc, lang)
    page_profitability(doc, lang)
    page_cash_mix(doc, lang)
    page_guidance(doc, lang)
    page_thesis(doc, lang)
    page_estimates(doc, lang)
    page_valuation(doc, lang, mkt)
    page_risks(doc, lang, mkt)
    page_sources(doc, lang)
    filename = (
        "WDC_Q1_FY2027_Earnings_Update.docx"
        if lang == "en"
        else "WDC_Q1_FY2027_业绩更新报告_中文版.docx"
    )
    path = OUT / filename
    doc.save(path)
    return path


if __name__ == "__main__":
    missing = [str(path) for path in CHARTS.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("Missing chart files: " + ", ".join(missing))
    mkt = market_data()
    print("Market data:", json.dumps(mkt, ensure_ascii=False))
    print("Created", build("en", mkt))
    print("Created", build("cn", mkt))
