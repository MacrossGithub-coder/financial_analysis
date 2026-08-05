#!/usr/bin/env python3
"""Build bilingual institutional earnings-update DOCX files for SPCX Q2 2026."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "SPCX"
OUT.mkdir(parents=True, exist_ok=True)

EARNINGS_RELEASE = "https://www.sec.gov/Archives/edgar/data/1181412/000162828026052515/earningsreleaseq22608042.htm"
FORM_10Q = "https://www.sec.gov/Archives/edgar/data/1181412/000162828026052535/spcx-20260630.htm"
FORM_8K = "https://www.sec.gov/Archives/edgar/data/1181412/000162828026052515/0001628280-26-052515-index.html"
IR_WEBCAST = "https://ir.spacex.com/updates/releases-details/2026/SpaceX-to-Post-Second-Quarter-2026-Results-and-Host-Webcast-on-August-4-2026-2026-g8layJlbFm/default.aspx"
S1 = "https://www.sec.gov/Archives/edgar/data/1181412/000162828026036936/spaceexplorationtechnologi.htm"
AXIOS = "https://www.axios.com/2026/08/04/spacex-earnings-elon-musk"
KIPLINGER = "https://www.kiplinger.com/investing/stocks/17494/next-week-earnings-calendar-stocks"
AP_CALL = "https://apnews.com/article/3b7b66a3e522e51d75caebc40af7e09e"
NASDAQ = "https://www.nasdaq.com/market-activity/stocks/spcx"
YAHOO = "https://finance.yahoo.com/quote/SPCX/"
LONG_BRIDGE = "https://longbridge.com/quote/SPCX.US"

CHARTS = {
    1: "spcx_q2_chart1_beat_miss.png",
    2: "spcx_q2_chart2_segment_revenue.png",
    3: "spcx_q2_chart3_segment_operating_income.png",
    4: "spcx_q2_chart4_starlink_metrics.png",
    5: "spcx_q2_chart5_ebitda_margin.png",
    6: "spcx_q2_chart6_ai_revenue_capex.png",
    7: "spcx_q2_chart7_capex_mix.png",
    8: "spcx_q2_chart8_cash_flow.png",
    9: "spcx_q2_chart9_estimate_revisions.png",
    10: "spcx_q2_chart10_price_since_ipo.png",
    11: "spcx_q2_chart11_valuation_scenarios.png",
}

NAVY = "0B2545"
BLUE = "005288"
LIGHT_BLUE = "DCE6F1"
PALE_BLUE = "EEF5FA"
GRAY = "667085"
LIGHT_GRAY = "F2F4F7"
GREEN = "2E7D32"
RED = "9B1C1C"
GOLD = "7A5A00"
WHITE = "FFFFFF"

# documents-skill preset: standard_business_brief.
# Named override: institutional_equity_research -- Times New Roman typography,
# 9.5 pt body, denser paragraph rhythm, and restrained SpaceX-blue accents.
BODY_FONT = "Times New Roman"
CN_BODY = "Songti SC"
CN_HEAD = "Heiti SC"


def fetch_market_data() -> dict:
    """Fetch dynamic market data with yfinance; return N/A rather than estimates."""
    code = r'''
import json, yfinance as yf
t = yf.Ticker("SPCX")
fi = t.fast_info
info = t.info
name = info.get("longName") or info.get("shortName") or ""
if "Space Exploration Technologies" not in name:
    raise RuntimeError(f"Ticker identity mismatch: {name}")
print(json.dumps({
    "price": float(fi.last_price),
    "market_cap": float(fi.market_cap),
    "year_high": float(fi.year_high),
    "year_low": float(fi.year_low),
    "name": name,
}))
'''
    try:
        result = subprocess.run(
            ["/Users/macrossz/anaconda3/bin/python", "-c", code],
            check=True,
            text=True,
            capture_output=True,
            timeout=45,
            env={"MPLCONFIGDIR": "/private/tmp/spcx-mpl"},
        )
        return json.loads(result.stdout.strip().splitlines()[-1])
    except Exception as exc:
        print(f"Warning: yfinance market-data fetch failed: {exc}")
        return {"price": None, "market_cap": None, "year_high": None, "year_low": None, "name": "N/A"}


def set_run_font(run, lang: str, size: float, bold=False, color=None, italic=False):
    run.font.name = BODY_FONT
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), BODY_FONT)
    rfonts.set(qn("w:hAnsi"), BODY_FONT)
    if lang == "cn":
        rfonts.set(qn("w:eastAsia"), CN_HEAD if bold else CN_BODY)


def add_hyperlink(paragraph, url: str, text: str, lang: str, size=7.5):
    part = paragraph.part
    rid = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run_el = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), BODY_FONT)
    rfonts.set(qn("w:hAnsi"), BODY_FONT)
    if lang == "cn":
        rfonts.set(qn("w:eastAsia"), CN_BODY)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.extend([rfonts, color, underline, sz])
    text_el = OxmlElement("w:t")
    text_el.text = text
    run_el.extend([rpr, text_el])
    hyperlink.append(run_el)
    paragraph._p.append(hyperlink)


def shade_cell(cell, fill: str):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc = cell._tc
    tcpr = tc.get_or_add_tcPr()
    tcmar = tcpr.first_child_found_in("w:tcMar")
    if tcmar is None:
        tcmar = OxmlElement("w:tcMar")
        tcpr.append(tcmar)
    for key, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcmar.find(qn(f"w:{key}"))
        if node is None:
            node = OxmlElement(f"w:{key}")
            tcmar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def mark_header_row(row):
    trpr = row._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    trpr.append(hdr)


def set_table_geometry(table, widths_dxa: list[int], indent_dxa=120):
    if sum(widths_dxa) != 9360:
        raise ValueError(f"Table widths must sum to 9360 DXA, got {sum(widths_dxa)}")
    table.autofit = False
    tblpr = table._tbl.tblPr
    layout = tblpr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblpr.append(layout)
    layout.set(qn("w:type"), "fixed")
    tblw = tblpr.find(qn("w:tblW"))
    if tblw is None:
        tblw = OxmlElement("w:tblW")
        tblpr.append(tblw)
    tblw.set(qn("w:w"), "9360")
    tblw.set(qn("w:type"), "dxa")
    tblind = tblpr.find(qn("w:tblInd"))
    if tblind is None:
        tblind = OxmlElement("w:tblInd")
        tblpr.append(tblind)
    tblind.set(qn("w:w"), str(indent_dxa))
    tblind.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            tcpr = cell._tc.get_or_add_tcPr()
            tcw = tcpr.find(qn("w:tcW"))
            if tcw is None:
                tcw = OxmlElement("w:tcW")
                tcpr.append(tcw)
            tcw.set(qn("w:w"), str(widths_dxa[idx]))
            tcw.set(qn("w:type"), "dxa")
            cell.width = Inches(widths_dxa[idx] / 1440)
            set_cell_margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_square_numbering(doc: Document) -> int:
    numbering = doc.part.numbering_part.element
    abstract_ids = [int(x.get(qn("w:abstractNumId"))) for x in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(x.get(qn("w:numId"))) for x in numbering.findall(qn("w:num"))]
    abstract_id = max(abstract_ids or [0]) + 1
    num_id = max(num_ids or [0]) + 1
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "singleLevel")
    abstract.append(multi)
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    numfmt = OxmlElement("w:numFmt")
    numfmt.set(qn("w:val"), "bullet")
    lvltext = OxmlElement("w:lvlText")
    lvltext.set(qn("w:val"), "■")
    suff = OxmlElement("w:suff")
    suff.set(qn("w:val"), "space")
    ppr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), "540")
    tabs.append(tab)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "720")
    ind.set(qn("w:hanging"), "180")
    ppr.extend([tabs, ind])
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), BODY_FONT)
    fonts.set(qn("w:hAnsi"), BODY_FONT)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    rpr.extend([fonts, color])
    lvl.extend([start, numfmt, lvltext, suff, ppr, rpr])
    abstract.append(lvl)
    numbering.append(abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    absid = OxmlElement("w:abstractNumId")
    absid.set(qn("w:val"), str(abstract_id))
    num.append(absid)
    numbering.append(num)
    return num_id


def apply_numbering(paragraph, num_id: int):
    ppr = paragraph._p.get_or_add_pPr()
    numpr = ppr.find(qn("w:numPr"))
    if numpr is None:
        numpr = OxmlElement("w:numPr")
        ppr.append(numpr)
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numid = OxmlElement("w:numId")
    numid.set(qn("w:val"), str(num_id))
    numpr.extend([ilvl, numid])


def set_styles(doc: Document, lang: str):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(9.0)
    normal._element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
    if lang == "cn":
        normal._element.rPr.rFonts.set(qn("w:eastAsia"), CN_BODY)
    pf = normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(3)
    pf.line_spacing = 1.04
    for name, size, before, after, color in (
        ("Heading 1", 14, 0, 6, BLUE),
        ("Heading 2", 11.5, 7, 4, NAVY),
        ("Heading 3", 10.5, 5, 3, NAVY),
    ):
        style = styles[name]
        style.font.name = BODY_FONT
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style._element.rPr.rFonts.set(qn("w:ascii"), BODY_FONT)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), BODY_FONT)
        if lang == "cn":
            style._element.rPr.rFonts.set(qn("w:eastAsia"), CN_HEAD)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_begin, instr, fld_sep, text, fld_end])


def configure_doc(doc: Document, lang: str):
    set_styles(doc, lang)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.header_distance = Inches(0.36)
    section.footer_distance = Inches(0.36)
    hp = section.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    text = "SpaceX (SPCX) | Q2 2026 Earnings Update" if lang == "en" else "SpaceX (SPCX) | 2026年第二季度业绩更新"
    run = hp.add_run(text)
    set_run_font(run, lang, 7.5, bold=True, color=GRAY)
    fp = section.footer.paragraphs[0]
    label = fp.add_run("Equity Research  |  " if lang == "en" else "股票研究  |  ")
    set_run_font(label, lang, 7.5, color=GRAY)
    add_page_number(fp)


def add_paragraph(doc, text: str, lang: str, size=9.0, bold=False, color=None, after=3, before=0, align=None, italic=False):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.04
    run = p.add_run(text)
    set_run_font(run, lang, size, bold=bold, color=color, italic=italic)
    return p


def add_rich_paragraph(doc, parts: Iterable[tuple[str, bool]], lang: str, size=9.0, after=3):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.04
    for text, bold in parts:
        run = p.add_run(text)
        set_run_font(run, lang, size, bold=bold)
    return p


def add_heading(doc, text: str, lang: str, level=1):
    p = doc.add_paragraph(text, style=f"Heading {level}")
    for run in p.runs:
        set_run_font(run, lang, 14 if level == 1 else 11.5, bold=True, color=BLUE if level == 1 else NAVY)
    return p


def add_bullet(doc, num_id: int, header: str, body: str, lang: str):
    p = doc.add_paragraph()
    apply_numbering(p, num_id)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.03
    r1 = p.add_run(header + " ")
    set_run_font(r1, lang, 8.9, bold=True, color=NAVY)
    r2 = p.add_run(body)
    set_run_font(r2, lang, 8.9)


def add_source_line(doc, lang: str, label: str, links: list[tuple[str, str]]):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    prefix = "Source: " if lang == "en" else "资料来源："
    run = p.add_run(prefix)
    set_run_font(run, lang, 7.3, italic=True, color=GRAY)
    for idx, (text, url) in enumerate(links):
        add_hyperlink(p, url, text, lang, size=7.3)
        if idx < len(links) - 1:
            sep = p.add_run("; ")
            set_run_font(sep, lang, 7.3, italic=True, color=GRAY)
    if label:
        tail = p.add_run(f". {label}")
        set_run_font(tail, lang, 7.3, italic=True, color=GRAY)
    return p


def add_table(doc, headers: list[str], rows: list[list[str]], widths: list[int], lang: str):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = header
        shade_cell(cell, NAVY)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs:
            set_run_font(run, lang, 7.8, bold=True, color=WHITE)
    mark_header_row(table.rows[0])
    for r_idx, row_data in enumerate(rows):
        cells = table.add_row().cells
        for c_idx, value in enumerate(row_data):
            cells[c_idx].text = value
            if r_idx % 2 == 1:
                shade_cell(cells[c_idx], LIGHT_GRAY)
            p = cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            for run in p.runs:
                set_run_font(run, lang, 7.6, bold=(r_idx == len(rows) - 1 and c_idx == 0))
    set_table_geometry(table, widths)
    return table


def add_figure(doc, number: int, title: str, alt_text: str, lang: str, width=5.35, sources=None, chart_number=None):
    cap_label = "Figure" if lang == "en" else "图"
    p = add_paragraph(
        doc,
        f"{cap_label} {number} - {title}",
        lang,
        size=8.6,
        bold=True,
        color=NAVY,
        after=2,
        before=3,
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )
    p.paragraph_format.keep_with_next = True
    pic_p = doc.add_paragraph()
    pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic_p.paragraph_format.space_after = Pt(0)
    run = pic_p.add_run()
    inline = run.add_picture(str(OUT / CHARTS[chart_number or number]), width=Inches(width))
    inline._inline.docPr.set("descr", alt_text)
    if sources:
        add_source_line(doc, lang, "", sources)


def add_page_break(doc):
    doc.add_page_break()


def build_report(lang: str, market: dict | None = None) -> Path:
    if lang not in {"en", "cn"}:
        raise ValueError(lang)
    market = market or fetch_market_data()
    price = market["price"]
    market_cap = market["market_cap"]
    year_high = market["year_high"]
    year_low = market["year_low"]
    target = 150.0
    upside = (target / price - 1) * 100 if price else None
    price_text = f"${price:.2f}" if price else "N/A"
    mcap_text = f"${market_cap / 1e12:.2f}T" if market_cap else "N/A"
    range_text = f"${year_low:.2f}-${year_high:.2f}" if year_low and year_high else "N/A"
    upside_text = f"{upside:.1f}%" if upside is not None else "N/A"

    doc = Document()
    configure_doc(doc, lang)
    num_id = add_square_numbering(doc)
    doc.core_properties.title = "SPCX Q2 2026 Earnings Update" if lang == "en" else "SPCX 2026年第二季度业绩更新报告"
    doc.core_properties.subject = "SpaceX Q2 2026 results, estimates, valuation, and investment thesis"
    doc.core_properties.author = "Equity Research"

    # PAGE 1 - summary
    if lang == "en":
        add_paragraph(doc, "SPACE EXPLORATION TECHNOLOGIES CORP. (SPCX)", lang, 16, True, BLUE, 1, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, "Q2 2026 EARNINGS UPDATE", lang, 13.5, True, NAVY, 2, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, "Revenue and EBITDA Beat; AI Capex Becomes the Valuation Battleground", lang, 11.5, True, BLUE, 7, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, f"5 August 2026  |  Rating: UPGRADE TO OUTPERFORM  |  Price: {price_text}  |  PT: $150  |  Upside: {upside_text}", lang, 8.8, True, GRAY, 5, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, f"Market cap: {mcap_text}  |  Since-listing range: {range_text}  |  IPO price: $135", lang, 8.3, False, GRAY, 6, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_heading(doc, "Earnings summary", lang, 2)
        headers = ["Metric", "Reported", "Reference", "Variance", "Reference basis"]
        rows = [
            ["Revenue", "$7.814B", "$6.900B", "+$0.914B / +13.2%", "S&P Visible Alpha consensus"],
            ["GAAP EPS", "$(0.09)", "$(0.29)", "+$0.20 better", "Pre-results Street consensus"],
            ["Adjusted EBITDA", "$3.538B", "$1.127B", "+214% QoQ", "Q1 2026 actual"],
            ["Capital expenditures", "$18.369B", "$10.107B", "+81.7% QoQ", "Q1 2026 actual"],
        ]
        add_table(doc, headers, rows, [1650, 1200, 1200, 1800, 3510], lang)
        add_source_line(doc, lang, "Results released and filed 4 August 2026", [("Q2 earnings release", EARNINGS_RELEASE), ("Form 10-Q", FORM_10Q), ("Axios / Visible Alpha consensus", AXIOS), ("Kiplinger pre-results consensus", KIPLINGER)])
        add_heading(doc, "Investment impact", lang, 2)
        add_bullet(doc, num_id, "The P&L beat was real and broad-based.", "Revenue rose 92% YoY and beat consensus by 13%; adjusted EBITDA nearly tripled to $3.5B. Connectivity remained the earnings engine, while AI adjusted EBITDA turned positive as contracted cloud capacity began to monetize.", lang)
        add_bullet(doc, num_id, "The cash-flow debate worsened materially.", "Q2 capex reached $18.4B, 86% of it in AI. H1 operating cash flow of $3.5B was overwhelmed by $28.5B of capex, leaving roughly $(25.0)B of free cash flow before acquisitions.", lang)
        add_bullet(doc, num_id, "Starlink quality improved despite lower ARPU.", "Subscribers doubled YoY to 12.0M and ARPU held at $66 QoQ. Connectivity operating income rose 79% to $1.7B and adjusted EBITDA margin remained above 60%.", lang)
        add_bullet(doc, num_id, "Upgrade to Outperform; maintain $150 PT.", f"The share-price reset to {price_text} creates {upside_text} upside to our unchanged target. We raise revenue and EBITDA estimates sharply, but the target remains capped by higher capex, execution risk, and the near-term post-IPO unlock overhang.", lang)
    else:
        add_paragraph(doc, "太空探索技术公司 (SPCX)", lang, 16, True, BLUE, 1, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, "2026年第二季度业绩更新报告", lang, 13.5, True, NAVY, 2, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, "营收与EBITDA显著超预期；AI资本开支成为估值核心矛盾", lang, 11.5, True, BLUE, 7, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, f"2026年8月5日  |  评级：上调至 跑赢大盘  |  股价：{price_text}  |  目标价：$150  |  上涨空间：{upside_text}", lang, 8.8, True, GRAY, 5, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_paragraph(doc, f"市值：{mcap_text}  |  上市以来区间：{range_text}  |  IPO发行价：$135", lang, 8.3, False, GRAY, 6, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_heading(doc, "业绩摘要", lang, 2)
        headers = ["指标", "报告值", "参考值", "差异", "参考基准"]
        rows = [
            ["收入", "$78.14亿", "$69.00亿", "+$9.14亿 / +13.2%", "S&P Visible Alpha一致预期"],
            ["GAAP每股收益", "$(0.09)", "$(0.29)", "好于预期$0.20", "财报前市场一致预期"],
            ["调整后EBITDA", "$35.38亿", "$11.27亿", "环比+214%", "Q1 2026实际值"],
            ["资本开支", "$183.69亿", "$101.07亿", "环比+81.7%", "Q1 2026实际值"],
        ]
        add_table(doc, headers, rows, [1650, 1200, 1200, 1800, 3510], lang)
        add_source_line(doc, lang, "业绩于2026年8月4日发布并提交", [("Q2业绩新闻稿", EARNINGS_RELEASE), ("Form 10-Q", FORM_10Q), ("Axios / Visible Alpha一致预期", AXIOS), ("Kiplinger财报前一致预期", KIPLINGER)])
        add_heading(doc, "投资影响", lang, 2)
        add_bullet(doc, num_id, "利润表超预期具有广度与质量。", "收入同比增长92%，较一致预期高13%；调整后EBITDA接近三倍增长至35亿美元。连接业务仍是盈利核心，AI则在云算力合同开始变现后实现调整后EBITDA转正。", lang)
        add_bullet(doc, num_id, "现金流争议显著加剧。", "Q2资本开支达到184亿美元，其中86%投向AI。H1经营现金流35亿美元远不足以覆盖285亿美元资本开支，自由现金流约为负250亿美元（未计并购）。", lang)
        add_bullet(doc, num_id, "尽管ARPU下滑，星链经营质量继续提升。", "用户数同比翻倍至1,200万，ARPU环比维持在$66。连接业务营业利润增长79%至17亿美元，调整后EBITDA利润率仍超过60%。", lang)
        add_bullet(doc, num_id, "评级上调至跑赢大盘，维持$150目标价。", f"股价回落至{price_text}，相对目标价的上涨空间为{upside_text}。我们大幅上调收入与EBITDA预测，但更高资本开支、执行风险和IPO后解禁压力限制了目标价上调空间。", lang)

    # PAGE 2 - beat/miss
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "1. Detailed results: a genuine operating beat", lang)
        add_paragraph(doc, "SpaceX reported Q2 revenue of $7.814B, up 92% YoY and 66% QoQ, exceeding S&P Visible Alpha consensus of $6.9B by $914M, or 13.2%. GAAP loss per share of $0.09 was $0.20 better than the pre-results Street expectation of a $0.29 loss. The quarter therefore cleared the two headline hurdles decisively; the negative share reaction was driven by investment intensity and supply mechanics rather than a weak P&L.", lang)
        add_paragraph(doc, "Adjusted EBITDA of $3.538B increased 191% YoY and 214% QoQ, lifting consolidated adjusted EBITDA margin to 45.3% from 24.0% in Q1. Operating loss narrowed to $143M from $1.943B in Q1 and $970M a year ago. Net loss of $541M improved by $467M YoY despite higher interest expense, reflecting operating leverage and a sharply better AI contribution.", lang)
        add_paragraph(doc, "Comparability requires care. xAI was combined with SpaceX in February 2026, and the AI segment is ramping large cloud-services contracts. YoY growth therefore captures both organic expansion and a materially different operating perimeter. We treat the Q2 run-rate as a forward indicator, but not as a clean steady-state margin.", lang)
        add_paragraph(doc, "The quality of the beat is stronger than the headline revenue variance alone suggests. All three reportable segments grew sequentially, consolidated operating loss nearly disappeared, and the AI segment crossed into positive adjusted EBITDA. At the same time, $3.679B of depreciation, amortization, and share-based compensation add-backs explain why adjusted EBITDA exceeded operating income by such a wide margin. The correct read-through is accelerating commercial traction, not near-term free-cash-flow maturity.", lang)
        add_figure(doc, 1, "Q2 2026 beat/miss snapshot", "Two-panel chart comparing reported Q2 revenue and GAAP EPS with pre-earnings consensus.", lang, 5.55, [("Q2 earnings release (4 Aug 2026)", EARNINGS_RELEASE), ("Axios / Visible Alpha consensus (4 Aug 2026)", AXIOS), ("Kiplinger pre-results consensus (31 Jul 2026)", KIPLINGER)])
    else:
        add_heading(doc, "1. 详细业绩：经营层面的真实超预期", lang)
        add_paragraph(doc, "SpaceX Q2收入为78.14亿美元，同比增长92%、环比增长66%，较S&P Visible Alpha 69亿美元一致预期高9.14亿美元，即13.2%。GAAP每股亏损$0.09，比财报前市场预期的$0.29亏损好$0.20。因此，本季度明确跨过了营收与每股收益两项核心门槛；股价负面反应主要来自投资强度和供给机制，而非利润表疲弱。", lang)
        add_paragraph(doc, "调整后EBITDA为35.38亿美元，同比增长191%、环比增长214%，合并调整后EBITDA利润率由Q1的24.0%升至45.3%。营业亏损从Q1的19.43亿美元和去年同期的9.70亿美元收窄至1.43亿美元。尽管利息费用上升，净亏损仍同比改善4.67亿美元至5.41亿美元，体现经营杠杆和AI贡献的显著改善。", lang)
        add_paragraph(doc, "同比口径需要谨慎。xAI于2026年2月并入SpaceX，AI板块正快速执行大型云服务合同。因此，增速既反映自然增长，也反映经营范围的重大变化。我们把Q2运行速度作为前瞻指标，但不将其视为稳定状态利润率。", lang)
        add_paragraph(doc, "超预期的质量不只体现在收入差异。三个报告板块均实现环比增长，合并营业亏损接近消失，AI板块也跨过调整后EBITDA转正门槛。与此同时，折旧、摊销及股权激励合计36.79亿美元，解释了调整后EBITDA与营业利润之间的巨大差距。正确解读应是商业化进展加速，而非近期自由现金流已经成熟。", lang)
        add_figure(doc, 1, "Q2 2026业绩与预期对比", "双面板图，对比Q2实际收入和GAAP每股收益与财报前一致预期。", lang, 5.55, [("Q2业绩新闻稿（2026年8月4日）", EARNINGS_RELEASE), ("Axios / Visible Alpha一致预期（2026年8月4日）", AXIOS), ("Kiplinger财报前一致预期（2026年7月31日）", KIPLINGER)])

    # PAGE 3 - segment charts
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "2. Segment mix: Connectivity funds two option-heavy businesses", lang)
        add_paragraph(doc, "Connectivity contributed 55% of Q2 revenue but more than all consolidated operating profit. AI revenue inflected sharply and adjusted EBITDA turned positive, while Space remained loss-making as Starship R&D accelerated. The portfolio is becoming more balanced on revenue, yet still highly concentrated on Starlink for cash earnings.", lang)
        add_paragraph(doc, "For portfolio construction, the segment mix creates a barbell. Connectivity is a scaled, recurring-revenue asset with visible subscriber economics; Space and AI are long-duration options whose outcomes are driven by technical milestones and utilization. This structure supports premium valuation when execution is strong, but it also increases covariance in a risk-off market because both optional businesses require external capital or Starlink cash generation before they can fund themselves.", lang)
        add_figure(doc, 2, "Segment revenue progression", "Grouped bars showing Space, Connectivity, and AI revenue for Q2 2025, Q1 2026, and Q2 2026.", lang, 5.25, [("Q2 earnings release (4 Aug 2026), financial highlights", EARNINGS_RELEASE), ("Form 10-Q (filed 4 Aug 2026)", FORM_10Q)])
        add_figure(doc, 3, "Segment operating income and loss", "Grouped bars showing operating income or loss by segment across three quarters.", lang, 5.25, [("Q2 earnings release (4 Aug 2026), segment schedules", EARNINGS_RELEASE), ("Form 10-Q (filed 4 Aug 2026)", FORM_10Q)])
    else:
        add_heading(doc, "2. 板块结构：连接业务为两项高期权业务提供资金", lang)
        add_paragraph(doc, "连接业务贡献Q2收入的55%，但其营业利润超过公司合并营业利润总额。AI收入显著加速并实现调整后EBITDA转正；太空业务则因星舰研发加速仍处亏损。收入结构正趋于均衡，但现金盈利仍高度依赖星链。", lang)
        add_paragraph(doc, "从组合构建角度看，板块结构呈现杠铃特征：连接业务是具备可见用户经济性的规模化经常性收入资产；太空与AI则是由技术里程碑和利用率驱动的长期期权。执行顺利时，这种结构支持估值溢价；风险偏好下降时，其相关性也会放大，因为两个期权型业务在实现自我融资前，都依赖外部资本或星链现金创造。", lang)
        add_figure(doc, 2, "板块收入变化", "分组柱状图，展示Q2 2025、Q1 2026和Q2 2026的太空、连接和AI收入。", lang, 5.25, [("Q2业绩新闻稿（2026年8月4日），财务摘要", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日提交）", FORM_10Q)])
        add_figure(doc, 3, "板块营业利润与亏损", "分组柱状图，展示三个季度各板块营业利润或亏损。", lang, 5.25, [("Q2业绩新闻稿（2026年8月4日），板块附表", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日提交）", FORM_10Q)])

    # PAGE 4 - segment table and EBITDA
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "3. Segment scorecard and margin bridge", lang)
        headers = ["Segment", "Revenue", "YoY", "Op. income/(loss)", "Adj. EBITDA", "Capex"]
        rows = [
            ["Space", "$0.962B", "+29%", "$(0.542)B", "$(0.205)B", "$1.174B"],
            ["Connectivity", "$4.291B", "+66%", "$1.656B", "$2.597B", "$1.367B"],
            ["AI", "$2.561B", "+247%", "$(1.257)B", "$1.146B", "$15.828B"],
            ["Total", "$7.814B", "+92%", "$(0.143)B", "$3.538B", "$18.369B"],
        ]
        add_table(doc, headers, rows, [1500, 1300, 900, 2000, 1800, 1860], lang)
        add_source_line(doc, lang, "Three months ended 30 June 2026", [("Q2 earnings release (4 Aug 2026), pages 2-8", EARNINGS_RELEASE), ("Form 10-Q (4 Aug 2026), pages 5 and 34-49", FORM_10Q)])
        add_paragraph(doc, "The consolidated adjusted EBITDA improvement was driven by three effects: higher Connectivity scale, initial monetization of AI cloud-service agreements, and the large non-cash depreciation and share-based compensation add-backs embedded in the non-GAAP measure. Depreciation and amortization alone totaled $2.848B in Q2, while share-based compensation was $831M. Investors should therefore avoid treating adjusted EBITDA as cash generation.", lang)
        add_paragraph(doc, "On a GAAP basis, Connectivity operating margin expanded to 38.6% from 35.7% a year ago. AI operating loss narrowed 17% YoY and 49% QoQ, but remained $1.257B. Space operating loss worsened YoY to $542M as R&D rose to $1.076B, more than the segment's quarterly revenue. This is the central cross-subsidy: Starlink earnings support Starship development and the AI build-out.", lang)
        add_figure(doc, 4, "Adjusted EBITDA and margin progression", "Bars showing adjusted EBITDA with a line showing adjusted EBITDA margin for three quarters.", lang, 5.35, [("Q2 earnings release (4 Aug 2026), adjusted EBITDA reconciliation", EARNINGS_RELEASE), ("Form 10-Q (4 Aug 2026)", FORM_10Q)], chart_number=5)
    else:
        add_heading(doc, "3. 板块评分与利润率桥接", lang)
        headers = ["板块", "收入", "同比", "营业利润/（亏损）", "调整后EBITDA", "资本开支"]
        rows = [
            ["太空", "$9.62亿", "+29%", "$(5.42)亿", "$(2.05)亿", "$11.74亿"],
            ["连接", "$42.91亿", "+66%", "$16.56亿", "$25.97亿", "$13.67亿"],
            ["AI", "$25.61亿", "+247%", "$(12.57)亿", "$11.46亿", "$158.28亿"],
            ["合计", "$78.14亿", "+92%", "$(1.43)亿", "$35.38亿", "$183.69亿"],
        ]
        add_table(doc, headers, rows, [1500, 1300, 900, 2000, 1800, 1860], lang)
        add_source_line(doc, lang, "截至2026年6月30日止三个月", [("Q2业绩新闻稿（2026年8月4日），第2-8页", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日），第5页及第34-49页", FORM_10Q)])
        add_paragraph(doc, "合并调整后EBITDA改善来自三方面：连接业务规模扩大、AI云服务协议开始变现，以及非GAAP指标中对大额非现金折旧与股权激励的加回。仅Q2折旧及摊销即达28.48亿美元，股权激励为8.31亿美元。因此，投资者不应把调整后EBITDA等同于现金创造。", lang)
        add_paragraph(doc, "按GAAP口径，连接业务营业利润率由去年同期35.7%升至38.6%。AI营业亏损同比收窄17%、环比收窄49%，但仍为12.57亿美元。太空业务营业亏损同比扩大至5.42亿美元，原因是研发费用增至10.76亿美元，超过该板块季度收入。由此形成核心交叉补贴：星链盈利支持星舰开发和AI建设。", lang)
        add_figure(doc, 4, "调整后EBITDA与利润率变化", "柱状图展示三个季度调整后EBITDA，折线展示调整后EBITDA利润率。", lang, 5.35, [("Q2业绩新闻稿（2026年8月4日），调整后EBITDA调节表", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日）", FORM_10Q)], chart_number=5)

    # PAGE 5 - Connectivity
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "4. Connectivity: the thesis strengthened", lang)
        add_paragraph(doc, "Starlink subscribers reached 12.0M at quarter-end, doubling from 6.0M a year earlier and increasing 1.7M sequentially. Monthly ARPU held at $66 vs. Q1 but remained 22% below Q2 2025's $85 as international mix and lower-priced plans broadened access. Stable sequential ARPU alongside 16.5% subscriber growth is a favorable near-term combination.", lang)
        add_paragraph(doc, "Consumer revenue increased 44% YoY to $2.485B, while Enterprise & Government revenue more than doubled to $1.806B. The latter reflects airline activations, mobile partnerships, and more than $6B of multi-year U.S. government awards for Starshield. Diversification beyond residential broadband should support ARPU and reduce churn, although customer concentration and government procurement timing add volatility.", lang)
        add_paragraph(doc, "Connectivity generated $1.656B of operating income and $2.597B of adjusted EBITDA, while capex was $1.367B. The segment therefore produced positive operating economics even before considering working capital. We mark this thesis pillar as strengthened: Starlink is scaling quickly enough to fund satellite refreshes and contribute substantial cash toward group-level investment.", lang)
        add_paragraph(doc, "The key sensitivity is the interaction between subscriber additions and pricing. At the current 12M subscriber base, each $1 change in monthly ARPU represents roughly $144M of annualized revenue before churn and mix effects. That arithmetic shows why enterprise, aviation, maritime, and direct-to-device services matter: they can stabilize blended pricing while lower-cost consumer plans expand the addressable market. We expect volume to remain the dominant driver through 2027.", lang)
        add_figure(doc, 5, "Starlink subscribers and ARPU", "Bars show Starlink subscribers; line shows monthly ARPU for Q2 2025, Q1 2026, and Q2 2026.", lang, 5.55, [("Q2 earnings release (4 Aug 2026), Connectivity segment data", EARNINGS_RELEASE), ("Form 10-Q (4 Aug 2026), segment discussion", FORM_10Q)], chart_number=4)
    else:
        add_heading(doc, "4. 连接业务：投资逻辑进一步增强", lang)
        add_paragraph(doc, "季度末星链用户达到1,200万，较去年同期600万翻倍，环比增加170万。月度ARPU与Q1持平于$66，但较Q2 2025的$85低22%，反映国际市场和低价套餐扩大覆盖。ARPU环比稳定、同时用户增长16.5%，是有利的短期组合。", lang)
        add_paragraph(doc, "消费者收入同比增长44%至24.85亿美元；企业与政府收入翻倍以上至18.06亿美元，受航空公司上线、移动通信合作以及Starshield获得逾60亿美元美国政府多年合同推动。住宅宽带之外的多元化应有助于支持ARPU并降低流失，但客户集中和政府采购节奏也会增加波动。", lang)
        add_paragraph(doc, "连接业务实现16.56亿美元营业利润和25.97亿美元调整后EBITDA，资本开支为13.67亿美元。即使不考虑营运资本，该板块也具备正向经营经济性。我们将这一投资逻辑支柱评为“增强”：星链扩张速度足以支持卫星更新，并为集团层面的投资贡献可观现金。", lang)
        add_paragraph(doc, "核心敏感性在于新增用户与定价的相互作用。按当前1,200万用户计算，月度ARPU每变化$1，对应约1.44亿美元年化收入（未计流失和组合效应）。这也说明企业、航空、海事和直连手机业务的重要性：它们可在低价消费者套餐扩张市场时稳定综合定价。我们预计到2027年，用户量仍将是主要增长驱动。", lang)
        add_figure(doc, 5, "星链用户数与ARPU", "柱状图展示星链用户数，折线展示Q2 2025、Q1 2026和Q2 2026月度ARPU。", lang, 5.55, [("Q2业绩新闻稿（2026年8月4日），连接业务数据", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日），板块讨论", FORM_10Q)], chart_number=4)

    # PAGE 6 - AI and capex
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "5. AI: monetization arrives, but capital intensity dominates", lang)
        add_paragraph(doc, "AI revenue rose 247% YoY and 213% QoQ to $2.561B. AI solutions and infrastructure revenue accounted for $2.194B, while advertising contributed $367M. New cloud-service agreements generated $1.6B of incremental quarterly infrastructure revenue and brought contracted sales to $14.1B. Nameplate compute increased to 1.4 GW from 1.0 GW in Q1 and 0.4 GW a year ago.", lang)
        add_paragraph(doc, "The improvement was not cheap. AI capex reached $15.828B, 6.2x segment revenue and more than double Q1. R&D expense was $2.178B. AI adjusted EBITDA turned positive at $1.146B, but GAAP operating loss remained $1.257B after depreciation and share-based compensation. The $60B Cursor acquisition adds strategic enterprise distribution but increases integration and capital-allocation complexity.", lang)
        add_paragraph(doc, "We mark the AI thesis as improved operationally but unproven economically. Contracted revenue and positive adjusted EBITDA validate demand. However, returns depend on utilization, pricing durability, GPU depreciation, power availability, and customer concentration. A small number of cloud customers can terminate after contractual ramp periods, making backlog quality as important as headline size.", lang)
        add_paragraph(doc, "A simple capital-efficiency check remains demanding. Q2 AI capex was about $6.18 for every dollar of segment revenue and nearly ten times the quarter's incremental cloud-infrastructure revenue. Those ratios will fall if contracted capacity ramps on schedule, but the depreciation clock starts before full utilization is guaranteed. We therefore require sequential evidence on booked revenue conversion, power delivery, and cash gross profit before assigning the segment a software-like multiple.", lang)
        add_figure(doc, 6, "AI revenue vs. AI capital expenditures", "Grouped bars compare AI segment revenue and capital expenditures across three quarters.", lang, 5.25, [("Q2 earnings release (4 Aug 2026), AI segment schedule", EARNINGS_RELEASE), ("Form 10-Q (4 Aug 2026), AI risk-factor update", FORM_10Q)])
    else:
        add_heading(doc, "5. AI：商业化到来，但资本密集度主导估值", lang)
        add_paragraph(doc, "AI收入同比增长247%、环比增长213%至25.61亿美元。其中，AI解决方案与基础设施收入21.94亿美元，广告收入3.67亿美元。新签云服务协议带来16亿美元季度新增基础设施收入，合同销售额达到141亿美元。名义算力由Q1的1.0GW和去年同期0.4GW升至1.4GW。", lang)
        add_paragraph(doc, "改善代价高昂。AI资本开支达到158.28亿美元，是板块收入的6.2倍、超过Q1两倍；研发费用21.78亿美元。AI调整后EBITDA转正至11.46亿美元，但计入折旧和股权激励后，GAAP营业亏损仍为12.57亿美元。以600亿美元收购Cursor增强企业分发能力，也提高了整合和资本配置复杂度。", lang)
        add_paragraph(doc, "我们将AI投资逻辑评为“经营改善、经济性未证实”。合同收入和正调整后EBITDA验证了需求，但投资回报取决于利用率、定价持续性、GPU折旧、电力可得性和客户集中度。少数云客户在合同爬坡期后可终止协议，因此积压订单质量与规模同样重要。", lang)
        add_paragraph(doc, "简单的资本效率检验仍然严苛。Q2 AI资本开支相当于每1美元板块收入投入约6.18美元，且接近本季度新增云基础设施收入的十倍。若合同产能按期爬坡，这些比率会下降；但折旧在利用率完全确定前已经开始。我们需要看到合同收入转化、电力交付和现金毛利的连续证据，才会给予该板块类似软件公司的估值倍数。", lang)
        add_figure(doc, 6, "AI收入与AI资本开支", "分组柱状图，对比三个季度AI板块收入和资本开支。", lang, 5.25, [("Q2业绩新闻稿（2026年8月4日），AI板块附表", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日），AI风险因素更新", FORM_10Q)])

    # PAGE 7 - cash flow and balance sheet
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "6. Cash flow and balance sheet: funded, not self-funding", lang)
        add_paragraph(doc, "H1 operating cash flow was $3.466B, up from $351M a year ago. Capital expenditures of $28.476B produced approximate free cash flow of $(25.010)B. This calculation excludes acquisitions and financing flows and is the most important counterweight to the adjusted EBITDA narrative. The company is producing operating cash, but the investment program remains far larger.", lang)
        add_paragraph(doc, "Liquidity is unusually strong after the IPO and bond issuance. Cash and equivalents were $93.522B and marketable securities $6.487B, totaling $100.009B. Debt and finance leases were $39.364B, implying roughly $60.6B of net cash before acquisition consideration. The IPO generated $85.675B of net proceeds, while the June bond issuance refinanced bridge borrowing and extended maturities.", lang)
        add_paragraph(doc, "At the H1 free-cash-flow run-rate, liquidity provides meaningful but finite runway. Our FY2026 capex estimate rises to $65B, broadly assuming the Q2 investment cadence remains elevated. If AI utilization, Starship milestones, or Starlink cash conversion disappoint, management would need to moderate spending or return to capital markets. Balance-sheet capacity lowers near-term solvency risk; it does not resolve return-on-capital risk.", lang)
        add_paragraph(doc, "Liquidity should also be viewed against committed strategic uses. The Cursor consideration, continued Starship development, satellite replenishment, and multi-gigawatt data-center construction compete for the same balance sheet. Even before acquisition cash outflows, our $65B FY2026 capex estimate would consume most of reported cash and marketable securities within two years if operating cash flow did not scale. The financing question is therefore about sequencing and returns, not immediate access to capital.", lang)
        add_figure(doc, 7, "H1 2026 operating cash flow, capex, and free cash flow", "Bar chart showing positive operating cash flow offset by much larger capital expenditures and negative free cash flow.", lang, 5.45, [("Q2 earnings release (4 Aug 2026), selected cash-flow information", EARNINGS_RELEASE), ("Form 10-Q (4 Aug 2026), cash-flow statement", FORM_10Q)], chart_number=8)
    else:
        add_heading(doc, "6. 现金流与资产负债表：资金充足，但尚未自我融资", lang)
        add_paragraph(doc, "H1经营现金流为34.66亿美元，高于去年同期3.51亿美元。资本开支284.76亿美元，对应约负250.10亿美元自由现金流。该计算未包括并购和融资现金流，是调整后EBITDA叙事最重要的制衡因素：公司能够产生经营现金，但投资计划规模仍远大于经营现金流。", lang)
        add_paragraph(doc, "IPO与债券发行后，流动性异常充裕。现金及等价物935.22亿美元、可交易证券64.87亿美元，合计1,000.09亿美元。债务及融资租赁393.64亿美元，对应并购对价前约606亿美元净现金。IPO净募资856.75亿美元；6月债券发行则再融资过桥贷款并延长到期结构。", lang)
        add_paragraph(doc, "按H1自由现金流消耗速度，流动性提供了有意义但并非无限的资金跑道。我们将FY2026资本开支预测上调至650亿美元，基本假设Q2投资节奏仍处高位。若AI利用率、星舰里程碑或星链现金转换不及预期，管理层需放缓支出或再次融资。资产负债表能力降低短期偿付风险，但不能消除资本回报风险。", lang)
        add_paragraph(doc, "流动性还应与已承诺的战略用途对照。Cursor交易对价、持续的星舰开发、卫星更新以及多吉瓦数据中心建设都在争夺同一张资产负债表。即使不计并购现金流，若经营现金流未能扩大，我们预计的FY2026资本开支650亿美元也会在两年内消耗大部分现金及可交易证券。因此，融资问题的核心是项目排序与回报，而不是近期能否获得资金。", lang)
        add_figure(doc, 7, "H1 2026经营现金流、资本开支与自由现金流", "柱状图显示正经营现金流被更大规模资本开支抵消，并形成负自由现金流。", lang, 5.45, [("Q2业绩新闻稿（2026年8月4日），现金流精选信息", EARNINGS_RELEASE), ("Form 10-Q（2026年8月4日），现金流量表", FORM_10Q)], chart_number=8)

    # PAGE 8 - call/guidance and catalysts
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "7. Management outlook and catalyst calendar", lang)
        add_paragraph(doc, "SpaceX did not provide conventional quarterly revenue or EPS guidance. Instead, management framed long-duration operating targets. On the 4 August call, Elon Musk said the company's internal projection for $1T of annual revenue moved forward to 2030 from 2031, with a non-zero probability of 2029. He also described a path toward as much as 10 GW of compute capacity by the end of 2027 and said SpaceX would standardize on Nvidia's Vera Rubin architecture.", lang)
        add_paragraph(doc, "Gwynne Shotwell reiterated a goal of landing people on the Moon in 2028. The earnings release highlighted two successful Starship V3 tests in the prior 90 days and management's belief that full reusability can lower cost to orbit by 99% or more vs. historical averages. These statements are strategically important but sit well beyond the normal forecasting window and carry unusually high execution risk.", lang)
        add_heading(doc, "Near-term catalysts", lang, 2)
        add_bullet(doc, num_id, "6 August 2026 - first major post-IPO unlock tranche.", "The release of eligible insider and employee shares can increase free float sharply and may dominate near-term price discovery even after a fundamental beat.", lang)
        add_bullet(doc, num_id, "Q3 2026 - expected Cursor transaction close.", "The acquisition could accelerate enterprise distribution and recurring software revenue, but investors need clarity on consideration, integration, and incremental operating expense.", lang)
        add_bullet(doc, num_id, "Second half 2026 - cloud capacity ramp and compute build-out.", "Quarterly AI revenue, contracted sales conversion, and utilization will determine whether Q2's positive adjusted EBITDA is durable.", lang)
        add_bullet(doc, num_id, "Starship flight cadence and satellite V3 deployment.", "Successful reusability milestones would improve the economic case for launch, Starlink capacity, and future orbital compute; delays would extend R&D losses.", lang)
        add_paragraph(doc, "Transcript status: the company stated that a transcript would be available the day after the call, but an official transcript was not discoverable in the IR index at this report's cut-off. Call commentary above was cross-checked against the official webcast notice and same-day Axios/AP coverage; no unmatched third-party transcript was used.", lang, size=8.4, italic=True, color=GRAY)
        add_source_line(doc, lang, "Call and release dated 4 August 2026", [("Official webcast/replay notice", IR_WEBCAST), ("Q2 earnings release", EARNINGS_RELEASE), ("Axios call coverage", AXIOS), ("AP call coverage", AP_CALL)])
    else:
        add_heading(doc, "7. 管理层展望与催化剂日历", lang)
        add_paragraph(doc, "SpaceX未给出传统的季度收入或每股收益指引，而是提出长期经营目标。在8月4日电话会上，Elon Musk表示公司达到年收入1万亿美元的内部预测由2031年提前至2030年，且2029年实现并非零概率。他还描述了到2027年底算力可能达到10GW的路径，并称SpaceX将统一采用Nvidia Vera Rubin架构。", lang)
        add_paragraph(doc, "Gwynne Shotwell重申2028年载人登月目标。业绩新闻稿强调过去90天完成两次成功的星舰V3测试，管理层认为完全复用可较历史平均水平降低99%或更多入轨成本。这些表述具有重大战略意义，但远超常规预测周期，执行风险也异常高。", lang)
        add_heading(doc, "近期催化剂", lang, 2)
        add_bullet(doc, num_id, "2026年8月6日——IPO后首批大规模解禁。", "符合条件的内部人士和员工股份释放将大幅增加流通盘，即使基本面超预期，也可能主导短期价格发现。", lang)
        add_bullet(doc, num_id, "2026年第三季度——预计完成Cursor交易。", "并购可加速企业分发和软件经常性收入，但投资者需要了解对价、整合和新增运营费用。", lang)
        add_bullet(doc, num_id, "2026年下半年——云算力产能爬坡。", "季度AI收入、合同销售额转化率和利用率将决定Q2正调整后EBITDA能否持续。", lang)
        add_bullet(doc, num_id, "星舰试飞节奏与V3卫星部署。", "成功实现复用里程碑将改善发射、星链容量和未来轨道算力的经济性；延迟则会延长研发亏损。", lang)
        add_paragraph(doc, "文字稿状态：公司称电话会次日将提供文字稿，但截至本报告截稿时，官方文字稿尚未在IR索引中检索到。以上电话会内容已与官方网络直播通知以及Axios/AP同日报道交叉核验；未使用日期或标的不匹配的第三方文字稿。", lang, size=8.4, italic=True, color=GRAY)
        add_source_line(doc, lang, "电话会与新闻稿日期均为2026年8月4日", [("官方网络直播/回放通知", IR_WEBCAST), ("Q2业绩新闻稿", EARNINGS_RELEASE), ("Axios电话会报道", AXIOS), ("AP电话会报道", AP_CALL)])

    # PAGE 9 - thesis/risk and price
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "8. Updated thesis and risk assessment", lang)
        add_bullet(doc, num_id, "Thesis pillar 1 - Starlink global scale: strengthened.", "Subscriber growth, stable sequential ARPU, and 79% operating-income growth demonstrate strong unit economics. Enterprise, aviation, mobile, and government channels broaden the addressable market beyond residential broadband.", lang)
        add_bullet(doc, num_id, "Thesis pillar 2 - integrated platform advantage: strengthened, but early.", "SpaceX can deploy satellites internally, monetize connectivity, and reuse infrastructure and capital-market access across AI. The Q2 AI revenue inflection is the first financial evidence of this platform logic.", lang)
        add_bullet(doc, num_id, "Thesis pillar 3 - self-funded compounding: weakened.", "H1 free cash flow of approximately $(25)B and $15.8B of quarterly AI capex show that strategic breadth has not yet translated into self-funding growth. Adjusted EBITDA substantially overstates cash economics during the build-out.", lang)
        add_heading(doc, "Principal risks", lang, 2)
        add_paragraph(doc, "Valuation and duration remain the leading risks. At roughly $1.5T market capitalization, investors are underwriting years of high growth and successful commercialization across multiple frontier technologies. AI customer concentration, cancellable cloud agreements after ramp periods, GPU obsolescence, power constraints, Starship technical milestones, regulatory approvals, government-contract timing, dual-class governance, and management key-person risk can all impair the path.", lang)
        add_paragraph(doc, "Near-term technical risk is unusually high because the first large unlock follows immediately after earnings. The stock has already fallen below the $135 IPO price and almost 50% from its $225.64 high. A cheaper entry price improves expected return but does not eliminate the possibility of further multiple compression as free float expands.", lang)
        add_paragraph(doc, "Our upgrade is therefore tactical on entry point but strategic on duration. The post-earnings decline improves asymmetry versus the unchanged $150 target, while the operating beat raises confidence that Starlink can carry group investment through the next milestone cycle. Position sizing should still reflect an unusually wide distribution of outcomes. Investors unable to tolerate a drawdown toward the $80 bear case should wait for evidence that AI capex growth has peaked.", lang)
        add_figure(doc, 8, "SPCX share-price performance since IPO", "Line chart of SPCX closes since the June 2026 IPO, with IPO price and Q2 results marked.", lang, 5.55, [("yfinance market data, accessed 5 Aug 2026", YAHOO), ("Longbridge live quote, accessed 5 Aug 2026", LONG_BRIDGE), ("Nasdaq listing page", NASDAQ)], chart_number=10)
    else:
        add_heading(doc, "8. 更新后的投资逻辑与风险评估", lang)
        add_bullet(doc, num_id, "投资逻辑支柱1——星链全球规模：增强。", "用户增长、ARPU环比稳定以及营业利润增长79%证明强劲单位经济性。企业、航空、移动和政府渠道把可服务市场扩展到住宅宽带之外。", lang)
        add_bullet(doc, num_id, "投资逻辑支柱2——一体化平台优势：增强，但仍处早期。", "SpaceX可内部部署卫星、变现连接服务，并在AI板块复用基础设施与资本市场能力。Q2 AI收入加速是该平台逻辑首次得到财务验证。", lang)
        add_bullet(doc, num_id, "投资逻辑支柱3——自我融资复利：弱化。", "H1约负250亿美元自由现金流和单季158亿美元AI资本开支显示，战略广度尚未转化为自我融资增长。建设期内，调整后EBITDA显著高估现金经济性。", lang)
        add_heading(doc, "主要风险", lang, 2)
        add_paragraph(doc, "估值与久期仍是首要风险。约1.5万亿美元市值意味着投资者正押注多年高速增长，以及多项前沿技术成功商业化。AI客户集中、云协议爬坡期后可取消、GPU淘汰、电力约束、星舰技术里程碑、监管审批、政府合同节奏、双重股权治理以及关键人物风险，均可能破坏实现路径。", lang)
        add_paragraph(doc, "近期技术面风险异常高，因为首批大规模解禁紧随财报而来。股价已跌破$135发行价，较$225.64高点接近腰斩。更低买入价改善预期回报，但在流通盘扩大时，不能排除估值倍数进一步压缩。", lang)
        add_paragraph(doc, "因此，本次上调在入场点上具有战术性，在持有期限上具有战略性。财报后下跌改善了相对不变$150目标价的收益不对称，而经营超预期提升了星链在下一轮里程碑周期中支持集团投资的可信度。仓位规模仍应反映结果分布异常宽广；无法承受股价回撤至$80熊市情景的投资者，应等待AI资本开支增速见顶的证据。", lang)
        add_figure(doc, 8, "SPCX自IPO以来股价表现", "折线图展示2026年6月IPO以来SPCX收盘价，并标记发行价和Q2财报。", lang, 5.55, [("yfinance市场数据，2026年8月5日获取", YAHOO), ("Longbridge实时行情，2026年8月5日获取", LONG_BRIDGE), ("Nasdaq上市页面", NASDAQ)], chart_number=10)

    # PAGE 10 - estimates
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "9. Estimate revisions: both growth and spending reset higher", lang)
        add_paragraph(doc, "We raise FY2026 revenue to $35.0B from $22.0B and adjusted EBITDA to $15.5B from $5.5B, reflecting the Q2 beat, $14.1B of contracted AI sales, and continued Starlink momentum. We also raise capex to $65B from $35B. Our model assumes H2 revenue of $22.5B, a meaningful acceleration that still sits well below management's long-duration aspirations.", lang)
        headers = ["Metric", "FY26 old", "FY26 new", "FY27 old", "FY27 new", "Post-Q2 rationale"]
        rows = [
            ["Revenue", "$22.0B", "$35.0B", "$28.0B", "$58.0B", "AI contracts + Starlink scale"],
            ["Adj. EBITDA", "$5.5B", "$15.5B", "$9.0B", "$26.0B", "AI turns positive; mix leverage"],
            ["Operating income", "$(6.0)B", "$(3.0)B", "$(3.0)B", "$4.0B", "Q2 near breakeven; D&A remains high"],
            ["Net income", "$(10.0)B", "$(6.0)B", "$(5.0)B", "$0.0B", "Operating improvement; higher interest"],
            ["Capital expenditures", "$35.0B", "$65.0B", "$30.0B", "$75.0B", "AI compute and power build-out"],
        ]
        add_table(doc, headers, rows, [1450, 1050, 1050, 1050, 1050, 3710], lang)
        add_source_line(doc, lang, "Old estimates from SPCX Q1 2026 earnings update dated 14 Jun 2026; new estimates are independent analyst estimates dated 5 Aug 2026", [("Q2 earnings release", EARNINGS_RELEASE), ("Form 10-Q", FORM_10Q), ("S-1 baseline", S1)])
        add_paragraph(doc, "FY2027 estimates assume Starlink subscribers and enterprise mix continue expanding, AI contracted sales convert into revenue, and depreciation catches up with the 2026 capex surge. We forecast GAAP operating profitability in 2027 but only breakeven net income because interest and other below-the-line costs remain material. The model is highly sensitive to AI utilization and capex timing; it should be interpreted as a scenario, not company guidance.", lang)
        add_figure(doc, 9, "Old vs. post-Q2 financial estimates", "Horizontal bars compare prior and post-Q2 estimates for revenue, adjusted EBITDA, and capex in 2026 and 2027.", lang, 5.55, [("SPCX Q1 2026 earnings update (14 Jun 2026), prior estimates", S1), ("Q2 earnings release and 10-Q (4 Aug 2026)", EARNINGS_RELEASE)])
    else:
        add_heading(doc, "9. 预测修正：增长与支出同时重置上行", lang)
        add_paragraph(doc, "我们将FY2026收入由220亿美元上调至350亿美元，调整后EBITDA由55亿美元上调至155亿美元，反映Q2超预期、141亿美元AI合同销售额和星链持续增长。同时将资本开支由350亿美元上调至650亿美元。模型假设H2收入225亿美元，明显加速，但仍远低于管理层长期愿景。", lang)
        headers = ["指标", "FY26旧", "FY26新", "FY27旧", "FY27新", "Q2后修正理由"]
        rows = [
            ["收入", "$220亿", "$350亿", "$280亿", "$580亿", "AI合同 + 星链规模"],
            ["调整后EBITDA", "$55亿", "$155亿", "$90亿", "$260亿", "AI转正；结构杠杆"],
            ["营业利润", "$(60)亿", "$(30)亿", "$(30)亿", "$40亿", "Q2接近盈亏平衡；折旧仍高"],
            ["净利润", "$(100)亿", "$(60)亿", "$(50)亿", "$0亿", "经营改善；利息费用更高"],
            ["资本开支", "$350亿", "$650亿", "$300亿", "$750亿", "AI算力与电力建设"],
        ]
        add_table(doc, headers, rows, [1450, 1050, 1050, 1050, 1050, 3710], lang)
        add_source_line(doc, lang, "旧预测来自2026年6月14日SPCX Q1 2026业绩更新；新预测为2026年8月5日独立分析师估计", [("Q2业绩新闻稿", EARNINGS_RELEASE), ("Form 10-Q", FORM_10Q), ("S-1基线", S1)])
        add_paragraph(doc, "FY2027预测假设星链用户和企业业务组合继续扩大、AI合同销售额转化为收入、折旧开始追上2026年资本开支激增。我们预测2027年实现GAAP营业盈利，但由于利息及其他线下费用仍高，净利润仅实现盈亏平衡。模型对AI利用率和资本开支节奏高度敏感，应被视为情景而非公司指引。", lang)
        add_figure(doc, 9, "旧预测与Q2后新预测对比", "横向柱状图，对比2026年和2027年收入、调整后EBITDA及资本开支的旧预测与新预测。", lang, 5.55, [("SPCX Q1 2026业绩更新（2026年6月14日），旧预测", S1), ("Q2业绩新闻稿及10-Q（2026年8月4日）", EARNINGS_RELEASE)])

    # PAGE 11 - valuation
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "10. Valuation and price target", lang)
        add_paragraph(doc, "We maintain a $150 12-month price target and upgrade SPCX to Outperform after the post-results decline. At the reference price, SpaceX trades at approximately 43x our FY2026 revenue estimate and 97x FY2026 adjusted EBITDA. Traditional near-term multiples are extreme and are distorted by heavy depreciation, so we use a scenario-weighted sum-of-the-parts framework.", lang)
        add_paragraph(doc, "Our operating SOTP values Connectivity at roughly $420B, based on 30x FY2027 adjusted EBITDA of approximately $14B; Space at $140B, including an explicit $100B Starship option value; and AI at approximately $725B, or 25x FY2027 revenue of about $29B. Adding roughly $61B of net cash yields about $1.35T, or $103 per share, before platform value. We apply a 45% strategic premium for vertical integration, cross-segment distribution, and capital-market access, producing a base value near $150 per share.", lang)
        add_paragraph(doc, "The bear case of $80 assumes slower AI utilization, capex remains elevated, and the platform premium disappears. The bull case of $225 assumes sustained Starlink margins, fast conversion of AI contracted sales, and successful Starship reuse milestones. The wide range is unavoidable: most value rests in duration-sensitive optionality rather than current free cash flow.", lang)
        add_figure(doc, 10, "Twelve-month valuation scenarios", "Bar chart showing bear, base, and bull per-share valuation scenarios with the reference price marked.", lang, 5.25, [("Independent analyst SOTP dated 5 Aug 2026", FORM_10Q), ("Q2 earnings release (4 Aug 2026)", EARNINGS_RELEASE), ("Dynamic market data", YAHOO)], chart_number=11)
        add_paragraph(doc, f"Rating / target: Outperform / $150 (prior: Market Perform / $150). Reference price: {price_text}. Implied upside: {upside_text}.", lang, 9.2, True, BLUE, 4)
    else:
        add_heading(doc, "10. 估值与目标价", lang)
        add_paragraph(doc, "财报后股价下跌，我们维持12个月目标价$150，并把SPCX评级上调至跑赢大盘。按参考股价，SpaceX约对应我们FY2026收入预测的43倍和调整后EBITDA预测的97倍。传统短期倍数极高且受巨额折旧扭曲，因此采用情景加权分部加总法。", lang)
        add_paragraph(doc, "经营SOTP中，连接业务估值约4,200亿美元，对应FY2027约140亿美元调整后EBITDA的30倍；太空业务估值1,400亿美元，其中包含明确的1,000亿美元星舰期权价值；AI估值约7,250亿美元，对应FY2027约290亿美元收入的25倍。加上约610亿美元净现金，平台溢价前价值约1.35万亿美元，即每股$103。我们为垂直整合、跨板块分发和资本市场能力给予45%战略溢价，得到每股约$150基准价值。", lang)
        add_paragraph(doc, "熊市情景$80假设AI利用率较低、资本开支维持高位且平台溢价消失。牛市情景$225假设星链利润率持续、AI合同销售快速转化并成功实现星舰复用里程碑。宽广区间无法避免：大部分价值来自对久期敏感的期权，而非当前自由现金流。", lang)
        add_figure(doc, 10, "十二个月估值情景", "柱状图展示熊市、基准和牛市每股估值情景，并标记参考股价。", lang, 5.25, [("2026年8月5日独立分析师SOTP", FORM_10Q), ("Q2业绩新闻稿（2026年8月4日）", EARNINGS_RELEASE), ("动态市场数据", YAHOO)], chart_number=11)
        add_paragraph(doc, f"评级 / 目标价：跑赢大盘 / $150（此前：中性 / $150）。参考股价：{price_text}。隐含上涨空间：{upside_text}。", lang, 9.2, True, BLUE, 4)

    # PAGE 12 - sources
    add_page_break(doc)
    if lang == "en":
        add_heading(doc, "Sources and references", lang)
        add_paragraph(doc, "Primary earnings materials", lang, 10.5, True, NAVY, 4)
        sources = [
            ("SpaceX Q2 2026 earnings release (4 August 2026)", EARNINGS_RELEASE),
            ("Form 10-Q for quarter ended 30 June 2026 (filed 4 August 2026)", FORM_10Q),
            ("Form 8-K and Exhibit 99.1 filing index (filed 4 August 2026)", FORM_8K),
            ("Official earnings webcast and replay notice (call 4 August 2026)", IR_WEBCAST),
            ("Form S-1 registration statement (filed 20 May 2026; prior-estimate baseline)", S1),
        ]
        for text, url in sources:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            add_hyperlink(p, url, text, lang, 8.5)
        add_paragraph(doc, "Consensus, call verification, and market data", lang, 10.5, True, NAVY, 4, before=5)
        sources2 = [
            ("Axios / S&P Visible Alpha Q2 results and call coverage (4 August 2026)", AXIOS),
            ("Kiplinger pre-earnings revenue and EPS consensus (31 July 2026)", KIPLINGER),
            ("Associated Press earnings-call coverage (4 August 2026)", AP_CALL),
            ("Yahoo Finance / yfinance dynamic market data (accessed 5 August 2026)", YAHOO),
            ("Longbridge live quote and historical market data (accessed 5 August 2026)", LONG_BRIDGE),
            ("Nasdaq SPCX listing page", NASDAQ),
        ]
        for text, url in sources2:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            add_hyperlink(p, url, text, lang, 8.5)
        add_paragraph(doc, "Methodology notes", lang, 10.5, True, NAVY, 4, before=6)
        add_paragraph(doc, "Beat/miss uses pre-results public consensus. Free cash flow is operating cash flow less capital expenditures. Adjusted EBITDA is company-defined and reconciled in the earnings release; it excludes large depreciation, share-based compensation, interest, and other items. Forecasts and valuation are independent analyst estimates, not company guidance. All dates, quarter labels, and filing identifiers were verified against the company release and SEC documents.", lang, size=8.5)
        add_paragraph(doc, "DISCLAIMER", lang, 8.5, True, GRAY, 2, before=8)
        add_paragraph(doc, "This report is for informational purposes only and does not constitute investment advice, an offer, or a solicitation. Estimates and price targets are based on publicly available information and may change without notice. The analyst is not a licensed investment adviser and has no position in SPCX. Investors should conduct independent due diligence and consider their own objectives and risk tolerance.", lang, size=7.6, color=GRAY)
    else:
        add_heading(doc, "资料来源与参考文献", lang)
        add_paragraph(doc, "主要业绩材料", lang, 10.5, True, NAVY, 4)
        sources = [
            ("SpaceX Q2 2026业绩新闻稿（2026年8月4日）", EARNINGS_RELEASE),
            ("截至2026年6月30日季度Form 10-Q（2026年8月4日提交）", FORM_10Q),
            ("Form 8-K及附件99.1索引（2026年8月4日提交）", FORM_8K),
            ("官方业绩网络直播及回放通知（电话会：2026年8月4日）", IR_WEBCAST),
            ("Form S-1注册声明（2026年5月20日提交；旧预测基线）", S1),
        ]
        for text, url in sources:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            add_hyperlink(p, url, text, lang, 8.5)
        add_paragraph(doc, "一致预期、电话会核验与市场数据", lang, 10.5, True, NAVY, 4, before=5)
        sources2 = [
            ("Axios / S&P Visible Alpha Q2业绩及电话会报道（2026年8月4日）", AXIOS),
            ("Kiplinger财报前收入和EPS一致预期（2026年7月31日）", KIPLINGER),
            ("Associated Press业绩电话会报道（2026年8月4日）", AP_CALL),
            ("Yahoo Finance / yfinance动态市场数据（2026年8月5日获取）", YAHOO),
            ("Longbridge实时行情及历史数据（2026年8月5日获取）", LONG_BRIDGE),
            ("Nasdaq SPCX上市页面", NASDAQ),
        ]
        for text, url in sources2:
            p = doc.add_paragraph(style="List Bullet")
            p.paragraph_format.space_after = Pt(2)
            add_hyperlink(p, url, text, lang, 8.5)
        add_paragraph(doc, "方法说明", lang, 10.5, True, NAVY, 4, before=6)
        add_paragraph(doc, "超预期/不及预期分析采用财报前公开一致预期。自由现金流定义为经营现金流减资本开支。调整后EBITDA为公司定义的非GAAP指标，调节表见业绩新闻稿；该指标排除大额折旧、股权激励、利息及其他项目。预测和估值为独立分析师估计，并非公司指引。所有日期、季度标签和申报编号均已与公司新闻稿及SEC文件核验。", lang, size=8.5)
        add_paragraph(doc, "免责声明", lang, 8.5, True, GRAY, 2, before=8)
        add_paragraph(doc, "本报告仅供信息参考，不构成投资建议、要约或招揽。预测和目标价基于公开信息，可能随时调整而不另行通知。分析师不是持牌投资顾问，且未持有SPCX。投资者应进行独立尽调，并结合自身目标与风险承受能力作出决策。", lang, size=7.6, color=GRAY)

    filename = "SPCX_Q2_2026_Earnings_Update.docx" if lang == "en" else "SPCX_Q2_2026_业绩更新报告_中文版.docx"
    path = OUT / filename
    doc.save(path)
    print(f"Saved {path}")
    return path


if __name__ == "__main__":
    build_report("en")
    build_report("cn")
