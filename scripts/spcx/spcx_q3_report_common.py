#!/usr/bin/env python3
"""Build bilingual institutional earnings-update DOCX files for SPCX Q3 2026."""

from __future__ import annotations

import json
from pathlib import Path

import yfinance as yf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

import spcx_q2_report_common as base


ROOT = Path("/Users/macrossz/DevTools/VscodeProject/ClaudeCode/financial_analysis")
OUT = ROOT / "output" / "SPCX"
DATA = ROOT / "data" / "spcx" / "q3_2026_facts.json"
OUT.mkdir(parents=True, exist_ok=True)

with DATA.open(encoding="utf-8") as handle:
    FACTS = json.load(handle)

EARNINGS_RELEASE = FACTS["sources"]["earnings_release"]
FORM_8K = FACTS["sources"]["form_8k"]
FORM_10Q = FACTS["sources"]["q2_form_10q"]
IR_EVENT = FACTS["sources"]["ir_event"]
TRANSCRIPT = FACTS["sources"]["transcript"]
CONSENSUS = FACTS["sources"]["consensus_preview"]
SEGMENT_CONSENSUS = FACTS["sources"]["segment_consensus"]
LONG_BRIDGE = FACTS["sources"]["market_data"]
Q2_EVENT = "https://ir.spacex.com/events/event-details/2026/Q2-2026-Earnings/default.aspx"
AP = "https://apnews.com/article/spacex-musk-turns-profit-losses-67064446bf0afdf529db55bdec82af69"

CHARTS = {
    1: "spcx_q3_chart1_beat_miss.png",
    2: "spcx_q3_chart2_revenue_ebitda.png",
    3: "spcx_q3_chart3_segment_revenue.png",
    4: "spcx_q3_chart4_segment_operating_income.png",
    5: "spcx_q3_chart5_margin_trends.png",
    6: "spcx_q3_chart6_starlink_metrics.png",
    7: "spcx_q3_chart7_si_revenue_capex_compute.png",
    8: "spcx_q3_chart8_cash_flow.png",
    9: "spcx_q3_chart9_guidance.png",
    10: "spcx_q3_chart10_estimate_revisions.png",
    11: "spcx_q3_chart11_valuation_scenarios.png",
}

# Reuse the tested Q2 report primitives while giving the new report its own
# chart map and complying with the document skill's black-heading requirement.
base.OUT = OUT
base.CHARTS = CHARTS
base.BLUE = "000000"
base.NAVY = "000000"
# LibreOffice's macOS font resolver reliably exposes the Songti collection by
# its PostScript-compatible STSong family name.  Keep headings in Heiti SC.
base.CN_BODY = "Songti SC"
base.CN_HEAD = "Heiti SC"

_base_set_run_font = base.set_run_font
_base_add_hyperlink = base.add_hyperlink


def _set_run_font_with_cjk_fallback(run, lang: str, size: float, bold=False, color=None, italic=False):
    """Use the named Chinese family in every OOXML font slot for LO parity."""
    _base_set_run_font(run, lang, size, bold=bold, color=color, italic=italic)
    if lang == "cn":
        family = base.CN_HEAD if bold else base.CN_BODY
        run.font.name = family
        rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
        for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            rfonts.set(qn(slot), family)


def _add_hyperlink_with_cjk_fallback(paragraph, url: str, text: str, lang: str, size=7.5):
    _base_add_hyperlink(paragraph, url, text, lang, size=size)
    if lang == "cn":
        hyperlink = paragraph._p[-1]
        rfonts = hyperlink.find(".//" + qn("w:rFonts"))
        if rfonts is not None:
            for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                rfonts.set(qn(slot), base.CN_BODY)


base.set_run_font = _set_run_font_with_cjk_fallback
base.add_hyperlink = _add_hyperlink_with_cjk_fallback


def fetch_market_data() -> dict:
    """Fetch current market fields from yfinance; fail visibly to N/A."""
    try:
        ticker = yf.Ticker("SPCX")
        info = ticker.info
        name = info.get("longName") or info.get("shortName") or ""
        if "Space Exploration Technologies" not in name:
            raise RuntimeError(f"Ticker identity mismatch: {name}")
        history = ticker.history(period="5d", auto_adjust=False)
        if history.empty:
            raise RuntimeError("SPCX price history is empty")
        fast = ticker.fast_info
        return {
            "price": float(history["Close"].dropna().iloc[-1]),
            "market_cap": float(fast.market_cap),
            "year_high": float(fast.year_high),
            "year_low": float(fast.year_low),
            "name": name,
            "source": "yfinance",
        }
    except Exception as exc:
        print(f"Warning: yfinance market-data fetch failed: {exc}")
        return {
            "price": None,
            "market_cap": None,
            "year_high": None,
            "year_low": None,
            "name": "N/A",
            "source": "N/A",
        }


def configure_doc(doc: Document, lang: str) -> None:
    base.set_styles(doc, lang)
    if lang == "cn":
        normal = doc.styles["Normal"]
        normal.font.name = base.CN_BODY
        for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            normal._element.rPr.rFonts.set(qn(slot), base.CN_BODY)
        for heading_name in ("Heading 1", "Heading 2", "Heading 3"):
            heading = doc.styles[heading_name]
            heading.font.name = base.CN_HEAD
            for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                heading._element.rPr.rFonts.set(qn(slot), base.CN_HEAD)
    title_style = doc.styles["Title"]
    title_style.font.name = base.CN_HEAD if lang == "cn" else base.BODY_FONT
    title_style.font.size = Pt(16)
    title_style.font.bold = True
    title_style.font.color.rgb = RGBColor(0, 0, 0)
    title_style._element.rPr.rFonts.set(qn("w:ascii"), base.BODY_FONT)
    title_style._element.rPr.rFonts.set(qn("w:hAnsi"), base.BODY_FONT)
    if lang == "cn":
        for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            title_style._element.rPr.rFonts.set(qn(slot), base.CN_HEAD)
    title_style.paragraph_format.space_before = Pt(0)
    title_style.paragraph_format.space_after = Pt(1)

    for heading_name in ("Heading 1", "Heading 2", "Heading 3"):
        doc.styles[heading_name].font.color.rgb = RGBColor(0, 0, 0)

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
    header = "SpaceX (SPCX) | Q3 2026 Earnings Update" if lang == "en" else "SpaceX (SPCX) | 2026年第三季度业绩更新"
    run = hp.add_run(header)
    base.set_run_font(run, lang, 7.5, bold=True, color=base.GRAY)

    fp = section.footer.paragraphs[0]
    label = fp.add_run("Equity Research  |  " if lang == "en" else "股票研究  |  ")
    base.set_run_font(label, lang, 7.5, color=base.GRAY)
    base.add_page_number(fp)


def add_title(doc: Document, text: str, lang: str) -> None:
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    run = p.add_run(text)
    base.set_run_font(run, lang, 16, bold=True, color="000000")


def add_disclaimer(doc: Document, lang: str) -> None:
    if lang == "en":
        base.add_paragraph(doc, "DISCLAIMER", lang, 8.5, True, base.GRAY, 2, before=8)
        base.add_paragraph(
            doc,
            "This report is for informational purposes only and does not constitute investment advice, an offer, or a solicitation. Forecasts and price targets are independent estimates based on public information and may change without notice. Investors should perform their own due diligence and consider their objectives and risk tolerance.",
            lang,
            size=7.6,
            color=base.GRAY,
        )
    else:
        base.add_paragraph(doc, "免责声明", lang, 8.5, True, base.GRAY, 2, before=8)
        base.add_paragraph(
            doc,
            "本报告仅供信息参考，不构成投资建议、要约或招揽。预测与目标价为基于公开信息作出的独立估计，可能随时调整而不另行通知。投资者应自行尽调，并结合自身目标与风险承受能力作出决策。",
            lang,
            size=7.6,
            color=base.GRAY,
        )


def add_source_item(doc: Document, text: str, url: str, lang: str) -> None:
    """Create a stable hanging bullet; LibreOffice can misplace List Bullet tabs."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.22)
    p.paragraph_format.first_line_indent = Inches(-0.18)
    p.paragraph_format.space_after = Pt(2)
    bullet = p.add_run("• ")
    base.set_run_font(bullet, lang, 8.5)
    base.add_hyperlink(p, url, text, lang, 8.5)


def build_report(lang: str, market: dict | None = None) -> Path:
    if lang not in {"en", "cn"}:
        raise ValueError(lang)
    market = market or fetch_market_data()
    price = market["price"]
    market_cap = market["market_cap"]
    year_high = market["year_high"]
    year_low = market["year_low"]
    target = FACTS["valuation"]["new_target"]
    upside = (target / price - 1) * 100 if price else None
    price_text = f"${price:.2f}" if price is not None else "N/A"
    mcap_text = f"${market_cap / 1e12:.2f}T" if market_cap is not None else "N/A"
    range_text = f"${year_low:.2f}-${year_high:.2f}" if year_low is not None and year_high is not None else "N/A"
    upside_text = f"{upside:.1f}%" if upside is not None else "N/A"

    doc = Document()
    configure_doc(doc, lang)
    num_id = base.add_square_numbering(doc)
    doc.core_properties.title = "SPCX Q3 2026 Earnings Update" if lang == "en" else "SPCX 2026年第三季度业绩更新报告"
    doc.core_properties.subject = "SpaceX Q3 2026 results, guidance, estimates, valuation, thesis, and risks"
    doc.core_properties.author = "Equity Research"

    # PAGE 1 - summary
    if lang == "en":
        add_title(doc, "SPACE EXPLORATION TECHNOLOGIES CORP SPCX", lang)
        base.add_paragraph(doc, "Q3 2026 EARNINGS UPDATE", lang, 13.5, True, "000000", 2, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_paragraph(doc, "SI Profit Inflection Resets the Earnings Curve while Capex Accelerates Again", lang, 11.5, True, "000000", 7, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_paragraph(doc, f"21 October 2026  |  Rating: MAINTAIN OUTPERFORM  |  Price: {price_text}  |  PT: $150 to $210  |  Upside: {upside_text}", lang, 8.8, True, base.GRAY, 5, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_paragraph(doc, f"Market cap: {mcap_text}  |  52-week range: {range_text}  |  After-hours: $159.71 (+8.1%)", lang, 8.3, False, base.GRAY, 6, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_heading(doc, "Earnings summary", lang)
        headers = ["Metric", "Reported", "Consensus", "Variance", "Read-through"]
        rows = [
            ["Revenue", "$14.575B", "$12.900B", "+$1.675B / +13.0%", "All three segments beat"],
            ["GAAP diluted EPS", "$0.35", "$0.14", "+$0.21 / +150%", "Operating and other income"],
            ["Adjusted EBITDA", "$8.809B", "$6.590B", "+$2.219B / +33.7%", "Cloud-service mix"],
            ["Capital expenditures", "$19.593B", "$18.369B", "+6.7% QoQ", "SI infrastructure"],
        ]
        base.add_table(doc, headers, rows, [1600, 1250, 1250, 1900, 3360], lang)
        base.add_source_line(doc, lang, "Results filed 20 October 2026; consensus dated 19 October 2026", [("Q3 earnings release and Exhibit 99.1", EARNINGS_RELEASE), ("SEC Form 8-K", FORM_8K), ("TipRanks pre-earnings consensus", CONSENSUS), ("Zacks segment consensus", SEGMENT_CONSENSUS)])
        base.add_heading(doc, "Investment impact", lang, 2)
        base.add_bullet(doc, num_id, "The beat was broad and the earnings curve changed shape.", "Revenue exceeded pre-earnings consensus by 13%, while diluted EPS of $0.35 was $0.21 above consensus. SI contributed $7.56B of revenue and $431M of GAAP operating income, crossing profitability earlier and at greater scale than our post-Q2 model assumed.", lang)
        base.add_bullet(doc, num_id, "Connectivity remains the quality anchor.", "Revenue increased 69% YoY to $5.25B, operating income rose 99% to $2.37B, subscribers reached 13.6M, and monthly ARPU improved 5% QoQ to $69. Starlink remains the most dependable source of recurring earnings and group-level funding.", lang)
        base.add_bullet(doc, num_id, "Management raised the growth and spending ceiling.", "The call introduced $33-37B of Q1 2027 revenue guidance and $35-40B of Q4 2026 capex. The $100B December ARR target and 2.5-3.0 GW year-end compute target imply another rapid step-up, but free cash flow remains deeply negative during deployment.", lang)
        base.add_bullet(doc, num_id, "Maintain Outperform; raise PT to $210 from $150.", "Higher SI utilization, a $63B backlog, and positive GAAP segment profit support a higher SOTP. We retain an Outperform rating, but use a wide $110-$320 scenario range because capital intensity, customer concentration, merger uncertainty, and execution risk remain exceptional.", lang)
    else:
        add_title(doc, "太空探索技术公司 SPCX", lang)
        base.add_paragraph(doc, "2026年第三季度业绩更新报告", lang, 13.5, True, "000000", 2, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_paragraph(doc, "SI盈利拐点重塑盈利曲线 资本开支再次提速", lang, 11.5, True, "000000", 7, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_paragraph(doc, f"2026年10月21日  |  评级：维持 跑赢大盘  |  股价：{price_text}  |  目标价：$150上调至$210  |  上涨空间：{upside_text}", lang, 8.8, True, base.GRAY, 5, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_paragraph(doc, f"市值：{mcap_text}  |  52周区间：{range_text}  |  盘后：$159.71（+8.1%）", lang, 8.3, False, base.GRAY, 6, align=WD_ALIGN_PARAGRAPH.CENTER)
        base.add_heading(doc, "业绩摘要", lang)
        headers = ["指标", "报告值", "一致预期", "差异", "解读"]
        rows = [
            ["收入", "$145.75亿", "$129.00亿", "+$16.75亿 / +13.0%", "三个板块均超预期"],
            ["GAAP摊薄EPS", "$0.35", "$0.14", "+$0.21 / +150%", "经营改善及其他收益"],
            ["调整后EBITDA", "$88.09亿", "$65.90亿", "+$22.19亿 / +33.7%", "云服务结构改善"],
            ["资本开支", "$195.93亿", "$183.69亿", "环比+6.7%", "SI基础设施"],
        ]
        base.add_table(doc, headers, rows, [1600, 1250, 1250, 1900, 3360], lang)
        base.add_source_line(doc, lang, "业绩于2026年10月20日提交；一致预期截至2026年10月19日", [("Q3业绩新闻稿及附件99.1", EARNINGS_RELEASE), ("SEC Form 8-K", FORM_8K), ("TipRanks财报前一致预期", CONSENSUS), ("Zacks板块一致预期", SEGMENT_CONSENSUS)])
        base.add_heading(doc, "投资影响", lang, 2)
        base.add_bullet(doc, num_id, "超预期具有广度，盈利曲线已改变。", "收入较财报前一致预期高13%，摊薄EPS为$0.35，超预期$0.21。SI贡献75.6亿美元收入和4.31亿美元GAAP营业利润，盈利拐点早于且规模显著高于我们Q2后的模型。", lang)
        base.add_bullet(doc, num_id, "连接业务仍是质量锚。", "收入同比增长69%至52.5亿美元，营业利润增长99%至23.7亿美元；用户达到1,360万，月度ARPU环比提升5%至$69。星链仍是最可靠的经常性盈利和集团资金来源。", lang)
        base.add_bullet(doc, num_id, "管理层同时抬高增长与支出上限。", "电话会首次给出2027Q1收入330-370亿美元和2026Q4资本开支350-400亿美元指引。12月1000亿美元ARR目标及年末2.5-3.0GW算力目标意味着又一次快速跃升，但部署期自由现金流仍将大幅为负。", lang)
        base.add_bullet(doc, num_id, "维持跑赢大盘 目标价由$150上调至$210。", "更高SI利用率、630亿美元积压订单和板块GAAP盈利支持更高SOTP。由于资本强度、客户集中、合并不确定性和执行风险仍极高，我们采用$110-$320的宽情景区间。", lang)

    # PAGE 2 - headline beat
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "1 Results were materially ahead of a moving consensus", lang)
        base.add_paragraph(doc, "SpaceX reported Q3 revenue of $14.575B, up 161% YoY and 87% QoQ. Against the $12.9B pre-earnings consensus published on 19 October, the company beat by $1.675B, or 13.0%. GAAP diluted EPS of $0.35 exceeded the $0.14 consensus by $0.21. Zacks' nine-analyst segment expectations were $1.43B for Space, $4.78B for Connectivity, and $6.03B for SI; reported results exceeded each by 24%, 10%, and 25%, respectively.", lang)
        base.add_paragraph(doc, "The beat was not merely a below-the-line event. Consolidated operating income improved to $2.791B from a $273M loss a year ago and a $143M loss in Q2. Adjusted EBITDA reached $8.809B, up 306% YoY and 149% QoQ. Revenue mix mattered: higher-margin SI cloud-service agreements scaled rapidly, while Connectivity maintained strong incremental margins. Other income of $1.025B and a $390M tax benefit also lifted net income, so the $0.35 EPS result should not be treated as a pure recurring run-rate.", lang)
        base.add_paragraph(doc, "Comparability remains imperfect. SI includes the August acquisition of Cursor, which contributed $608M of revenue from closing through quarter-end. The acquisition also generated $1.217B of Q3 share-based compensation, depressing GAAP SI operating income. Excluding that acquisition-related charge, management indicated SI would have produced about $1.6B of operating profit. We therefore view the SI margin inflection as economically meaningful but still subject to purchase-accounting and customer-mix volatility.", lang)
        base.add_figure(doc, 1, "Q3 2026 beat and miss snapshot", "Two-panel chart comparing Q3 revenue and GAAP diluted EPS with pre-earnings consensus.", lang, 5.35, [("Q3 earnings release, 20 Oct 2026, pp. 1 and 9", EARNINGS_RELEASE), ("TipRanks pre-earnings consensus, 19 Oct 2026", CONSENSUS), ("Zacks key-metric consensus, 20 Oct 2026", SEGMENT_CONSENSUS)])
        base.add_figure(doc, 2, "Quarterly revenue and adjusted EBITDA", "Grouped bars showing quarterly revenue and adjusted EBITDA from Q3 2025 through Q3 2026.", lang, 5.35, [("Q3 earnings release, 20 Oct 2026, pp. 3 and 9", EARNINGS_RELEASE), ("Q2 Form 10-Q, filed 4 Aug 2026, pp. 5 and 40", FORM_10Q)])
    else:
        base.add_heading(doc, "1 业绩显著高于快速上修的一致预期", lang)
        base.add_paragraph(doc, "SpaceX Q3收入145.75亿美元，同比增长161%、环比增长87%。相较10月19日发布的129亿美元财报前一致预期，公司超出16.75亿美元，即13.0%。GAAP摊薄EPS为$0.35，较$0.14一致预期高$0.21。Zacks九位分析师对太空、连接和SI收入的预期分别为14.3亿、47.8亿和60.3亿美元；实际值分别高24%、10%和25%。", lang)
        base.add_paragraph(doc, "超预期并非只来自线下项目。合并营业利润由去年同期亏损2.73亿美元、Q2亏损1.43亿美元改善至27.91亿美元。调整后EBITDA达88.09亿美元，同比增长306%、环比增长149%。收入结构是关键：高利润率SI云服务协议快速扩张，连接业务保持强劲增量利润率。10.25亿美元其他收益和3.90亿美元所得税收益也推高净利润，因此不应把$0.35 EPS完全视为可持续运行速度。", lang)
        base.add_paragraph(doc, "可比性仍不完美。SI包含8月完成的Cursor收购，交割日至季末贡献6.08亿美元收入；该交易也产生12.17亿美元Q3并购相关股权激励，压低GAAP SI营业利润。管理层表示，剔除该费用后SI营业利润约16亿美元。因此，SI利润率拐点具备经济意义，但仍受购买会计和客户结构波动影响。", lang)
        base.add_figure(doc, 1, "Q3 2026实际值与预期对比", "双面板图，对比Q3实际收入和GAAP摊薄EPS与财报前一致预期。", lang, 5.35, [("Q3业绩新闻稿，2026年10月20日，第1及第9页", EARNINGS_RELEASE), ("TipRanks财报前一致预期，2026年10月19日", CONSENSUS), ("Zacks关键指标一致预期，2026年10月20日", SEGMENT_CONSENSUS)])
        base.add_figure(doc, 2, "季度收入与调整后EBITDA", "分组柱状图，展示Q3 2025至Q3 2026的季度收入与调整后EBITDA。", lang, 5.35, [("Q3业绩新闻稿，2026年10月20日，第3及第9页", EARNINGS_RELEASE), ("Q2 Form 10-Q，2026年8月4日提交，第5及第40页", FORM_10Q)])

    # PAGE 3 - segments
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "2 Segment mix now has two profit engines", lang)
        base.add_paragraph(doc, "SI became the largest revenue segment in Q3, contributing 52% of consolidated sales vs. 36% from Connectivity and 12% from Space. SI revenue rose 195% QoQ and 830% YoY to $7.560B, driven by new cloud-service activations, Cursor, and higher Grok adoption. More important, SI moved from a $1.257B operating loss in Q2 to $431M of operating income, while adjusted EBITDA expanded from $1.146B to $4.953B.", lang)
        base.add_paragraph(doc, "Connectivity remains the cleaner earnings stream. Revenue increased 22% QoQ and 69% YoY to $5.248B. Operating income of $2.367B equated to a 45.1% margin, up from 38.6% in Q2, while adjusted EBITDA reached $3.499B. Consumer revenue grew to $3.121B and Enterprise and Government revenue to $2.127B, reducing reliance on any single channel.", lang)
        base.add_paragraph(doc, "Space revenue improved 84% QoQ to $1.767B, helped by customer mix and 12 customer launches. The segment was almost breakeven at a $7M operating loss, but adjusted EBITDA was $357M. R&D remained $1.225B as Starship work accelerated. The first revenue-generating Starship mission deployed 26 Starlink V3 satellites, providing strategic proof that launch economics can improve connectivity capacity, even though the near-term Space P&L remains burdened by development spending.", lang)
        base.add_paragraph(doc, "The portfolio is therefore no longer a single-profit-engine story. Connectivity still supplies the highest-quality recurring cash earnings, but SI has crossed into positive GAAP contribution and can become the dominant value driver if utilization remains high. The key analytical shift is from asking whether SI can monetize to asking whether its returns can remain attractive after depreciation, stock compensation, and a much larger capital base are fully reflected.", lang)
        base.add_figure(doc, 3, "Segment revenue progression", "Grouped bars showing Space, Connectivity, and SI revenue from Q3 2025 through Q3 2026.", lang, 5.2, [("Q3 earnings release, 20 Oct 2026, pp. 3-7", EARNINGS_RELEASE), ("Q2 Form 10-Q, filed 4 Aug 2026, pp. 37-44", FORM_10Q)])
        base.add_figure(doc, 4, "Segment operating income and loss", "Grouped bars showing operating income or loss by segment over four reported quarters.", lang, 5.2, [("Q3 earnings release, 20 Oct 2026, pp. 3-7", EARNINGS_RELEASE), ("Q2 Form 10-Q, filed 4 Aug 2026, segment note", FORM_10Q)])
    else:
        base.add_heading(doc, "2 板块结构已有两个盈利引擎", lang)
        base.add_paragraph(doc, "SI在Q3成为最大收入板块，占合并收入52%；连接和太空分别占36%和12%。SI收入环比增长195%、同比增长830%至75.60亿美元，受新云服务上线、Cursor并表和Grok采用率提高推动。更重要的是，SI由Q2营业亏损12.57亿美元转为盈利4.31亿美元，调整后EBITDA由11.46亿美元增至49.53亿美元。", lang)
        base.add_paragraph(doc, "连接业务仍是更纯粹的盈利来源。收入环比增长22%、同比增长69%至52.48亿美元。营业利润23.67亿美元，对应45.1%利润率，高于Q2的38.6%；调整后EBITDA达34.99亿美元。消费者收入增至31.21亿美元，企业与政府收入增至21.27亿美元，降低对单一渠道的依赖。", lang)
        base.add_paragraph(doc, "太空业务收入环比增长84%至17.67亿美元，受客户结构和12次客户发射推动。营业亏损仅700万美元，接近盈亏平衡；调整后EBITDA为3.57亿美元。星舰研发加速使研发费用仍达12.25亿美元。首次产生收入的星舰任务部署26颗星链V3卫星，证明发射经济性可改善连接容量，但近期太空利润表仍受开发支出压制。", lang)
        base.add_paragraph(doc, "因此，公司已不再只有一个盈利引擎。连接业务仍提供质量最高的经常性现金盈利，但SI已实现GAAP正贡献；若利用率保持高位，SI可成为最大价值驱动。分析焦点由“SI能否变现”转向“在充分计入折旧、股权激励及更大资本基数后，其回报能否维持吸引力”。", lang)
        # Slightly smaller than the English figures so the explicit next-page
        # break stays on this page with Songti metrics in LibreOffice and Word.
        base.add_figure(doc, 3, "板块收入变化", "分组柱状图，展示Q3 2025至Q3 2026的太空、连接和SI收入。", lang, 4.75, [("Q3业绩新闻稿，2026年10月20日，第3-7页", EARNINGS_RELEASE), ("Q2 Form 10-Q，2026年8月4日提交，第37-44页", FORM_10Q)])
        base.add_figure(doc, 4, "板块营业利润与亏损", "分组柱状图，展示四个报告季度各板块营业利润或亏损。", lang, 4.75, [("Q3业绩新闻稿，2026年10月20日，第3-7页", EARNINGS_RELEASE), ("Q2 Form 10-Q，2026年8月4日提交，板块附注", FORM_10Q)])

    # PAGE 4 - scorecard and margins
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "3 Profitability crossed a threshold but quality needs adjustment", lang)
        headers = ["Segment", "Revenue", "YoY", "Op income", "Adj EBITDA", "Capex"]
        rows = [
            ["Space", "$1.767B", "+6%", "$(0.007)B", "$0.357B", "$1.606B"],
            ["Connectivity", "$5.248B", "+69%", "$2.367B", "$3.499B", "$1.543B"],
            ["Super Intelligence", "$7.560B", "+830%", "$0.431B", "$4.953B", "$16.444B"],
            ["Total", "$14.575B", "+161%", "$2.791B", "$8.809B", "$19.593B"],
        ]
        base.add_table(doc, headers, rows, [1500, 1300, 900, 1800, 1800, 2060], lang)
        base.add_source_line(doc, lang, "Three months ended 30 September 2026", [("Q3 earnings release, pp. 3-7", EARNINGS_RELEASE), ("SEC Form 8-K, filed 20 Oct 2026", FORM_8K)])
        base.add_paragraph(doc, "Gross margin expanded to 62.4% from 51.5% a year ago and 55.3% in Q2 as higher-margin cloud agreements increased SI mix. Operating margin reached 19.1%, a 24-point YoY improvement, and net margin reached 30.2%. Management cautioned that Q3 benefited from one-time items and customer mix, and expects Q4 gross margin to be closer to Q2. We therefore model partial normalization rather than extrapolating the Q3 peak.", lang)
        base.add_paragraph(doc, "Adjusted EBITDA requires a disciplined bridge. The $8.809B figure adds back $3.644B of depreciation and amortization, $1.146B of recurring share-based compensation, $1.217B of acquisition-related share-based compensation, and other items. Those add-backs are economically relevant for a company investing $19.6B in a single quarter. Adjusted EBITDA demonstrates utilization and pricing power, but it substantially overstates cash available to equity holders during the build-out.", lang)
        base.add_paragraph(doc, "Below the operating line, $873M of interest income and $1.025B of other income more than offset $675M of interest expense. The resulting $4.404B net income exceeded operating income by $1.613B. Our forecast normalizes other income and the tax benefit, while retaining positive operating leverage from SI and Connectivity. This approach produces lower near-term EPS than a simple annualization of Q3 but a more durable base for valuation.", lang)
        base.add_figure(doc, 5, "Consolidated margin progression", "Lines showing gross, operating, and net margins from Q3 2025 through Q3 2026.", lang, 5.45, [("Q3 earnings release, 20 Oct 2026, pp. 9-10", EARNINGS_RELEASE), ("Q2 Form 10-Q, filed 4 Aug 2026, consolidated statements", FORM_10Q)])
    else:
        base.add_heading(doc, "3 盈利跨过门槛 但质量需要调整", lang)
        headers = ["板块", "收入", "同比", "营业利润", "调整后EBITDA", "资本开支"]
        rows = [
            ["太空", "$17.67亿", "+6%", "$(0.07)亿", "$3.57亿", "$16.06亿"],
            ["连接", "$52.48亿", "+69%", "$23.67亿", "$34.99亿", "$15.43亿"],
            ["超级智能", "$75.60亿", "+830%", "$4.31亿", "$49.53亿", "$164.44亿"],
            ["合计", "$145.75亿", "+161%", "$27.91亿", "$88.09亿", "$195.93亿"],
        ]
        base.add_table(doc, headers, rows, [1500, 1300, 900, 1800, 1800, 2060], lang)
        base.add_source_line(doc, lang, "截至2026年9月30日止三个月", [("Q3业绩新闻稿，第3-7页", EARNINGS_RELEASE), ("SEC Form 8-K，2026年10月20日提交", FORM_8K)])
        base.add_paragraph(doc, "毛利率由去年同期51.5%和Q2的55.3%升至62.4%，原因是高利润率云协议提升SI占比。营业利润率达19.1%，同比改善24个百分点；净利率达30.2%。管理层提醒，Q3受一次性项目和客户结构提振，预计Q4毛利率更接近Q2。因此，我们预测部分正常化，而非外推Q3峰值。", lang)
        base.add_paragraph(doc, "调整后EBITDA需要严格桥接。88.09亿美元包括加回36.44亿美元折旧及摊销、11.46亿美元经常性股权激励、12.17亿美元并购相关股权激励及其他项目。对于单季资本开支196亿美元的公司，这些加回具有真实经济成本。调整后EBITDA证明利用率与定价能力，但在建设期显著高估可供股东使用的现金。", lang)
        base.add_paragraph(doc, "线下项目方面，8.73亿美元利息收入和10.25亿美元其他收益超过6.75亿美元利息费用，使44.04亿美元净利润比营业利润高16.13亿美元。我们的预测对其他收益和所得税收益作正常化处理，同时保留SI与连接业务的正向经营杠杆；由此得到的近期EPS低于简单年化Q3，但估值基础更可持续。", lang)
        base.add_figure(doc, 5, "合并利润率变化", "折线图展示Q3 2025至Q3 2026的毛利率、营业利润率和净利率。", lang, 5.45, [("Q3业绩新闻稿，2026年10月20日，第9-10页", EARNINGS_RELEASE), ("Q2 Form 10-Q，2026年8月4日提交，合并财务报表", FORM_10Q)])

    # PAGE 5 - Connectivity
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "4 Connectivity strengthened the core thesis", lang)
        base.add_paragraph(doc, "Starlink consumer subscribers reached 13.6M at quarter-end, up 89% YoY and 13% QoQ. Monthly consumer ARPU improved 5% sequentially to $69, reversing the prior decline even as the company continued expanding internationally. Consumer revenue increased 57% YoY to $3.121B. Enterprise and Government revenue rose 91% to $2.127B, supported by aviation, maritime, government, and other managed deployments.", lang)
        base.add_paragraph(doc, "Operating economics improved faster than revenue. Connectivity operating income rose 99% YoY to $2.367B and adjusted EBITDA increased 82% to $3.499B. Capex of $1.543B was only modestly above Q2 despite rapid growth, implying that the deployed constellation and customer base are generating operating leverage. The segment's 45.1% operating margin and 66.7% adjusted EBITDA margin are the strongest evidence that Starlink can finance broader group ambitions.", lang)
        base.add_paragraph(doc, "The next growth vector is Starlink Mobile. SpaceX agreed on 8 October to acquire up to 14 MHz of paired nationwide 800 MHz spectrum, with closing expected in H1 2027. Management described a hybrid architecture using S-band capacity, low-band spectrum for building penetration, and radios installed at existing Starlink locations. This expands the addressable market but introduces regulatory, integration, and capital-allocation risk.", lang)
        base.add_paragraph(doc, "V3 satellites should improve capacity economics. Flight 14 deployed 26 V3 satellites with roughly ten times the downlink capacity added by a Falcon 9 Starlink launch. Management expects users to feel a meaningful improvement after a few hundred V3 satellites are operating, potentially by mid-2027. Execution on Starship cadence therefore directly influences Starlink service quality and long-term margin.", lang)
        base.add_figure(doc, 6, "Starlink consumer subscribers and ARPU", "Bars show Starlink consumer subscribers while the line shows monthly consumer ARPU.", lang, 5.55, [("Q3 earnings release, 20 Oct 2026, pp. 5-6", EARNINGS_RELEASE), ("Q3 earnings call transcript, 20 Oct 2026", TRANSCRIPT)])
    else:
        base.add_heading(doc, "4 连接业务强化核心投资逻辑", lang)
        base.add_paragraph(doc, "季度末星链消费者用户达到1,360万，同比增长89%、环比增长13%。月度消费者ARPU环比提升5%至$69，在公司继续国际扩张的同时扭转此前下滑。消费者收入同比增长57%至31.21亿美元；企业与政府收入增长91%至21.27亿美元，受航空、海事、政府及其他托管部署推动。", lang)
        base.add_paragraph(doc, "经营经济性改善快于收入。连接业务营业利润同比增长99%至23.67亿美元，调整后EBITDA增长82%至34.99亿美元。尽管业务高速增长，15.43亿美元资本开支仅略高于Q2，说明已部署星座和客户基础正在产生经营杠杆。45.1%营业利润率和66.7%调整后EBITDA利润率，是星链能够为集团更广泛目标提供资金的最强证据。", lang)
        base.add_paragraph(doc, "下一增长向量是Starlink Mobile。SpaceX于10月8日签署协议，拟收购全国范围最多14MHz成对800MHz频谱，预计2027年上半年完成。管理层描述的混合架构包括S频段容量、用于建筑物穿透的低频谱，以及在现有星链地点部署无线电设备。可服务市场扩大，但也带来监管、整合和资本配置风险。", lang)
        base.add_paragraph(doc, "V3卫星应改善容量经济性。Flight 14部署26颗V3卫星，其下行容量约等于一次猎鹰9号星链发射新增容量的十倍。管理层预计，在轨V3卫星达到数百颗后，用户将明显感受到改善，时间可能在2027年中。星舰发射频率因此直接影响星链服务质量和长期利润率。", lang)
        base.add_figure(doc, 6, "星链消费者用户与ARPU", "柱状图展示星链消费者用户数，折线展示月度消费者ARPU。", lang, 5.55, [("Q3业绩新闻稿，2026年10月20日，第5-6页", EARNINGS_RELEASE), ("Q3业绩电话会文字稿，2026年10月20日", TRANSCRIPT)])

    # PAGE 6 - SI
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "5 SI monetization validated demand but not yet cash returns", lang)
        base.add_paragraph(doc, "SI revenue increased to $7.560B, of which $7.022B came from solutions and infrastructure and $538M from advertising. Nameplate compute reached 1.7 GW, up from 1.4 GW in Q2 and 0.6 GW a year ago. The revenue step-up shows that new capacity is being monetized quickly: management described robust third-party demand and flexibility to allocate compute between internal model training and external cloud customers.", lang)
        base.add_paragraph(doc, "The segment's adjusted EBITDA of $4.953B exceeded revenue less cash operating costs by a wide margin, but capex remained $16.444B, or 2.2 times quarterly revenue. Management plans to deploy roughly 200,000-220,000 GB300 systems per month through year-end, reach 2.5-3.0 GW of nameplate compute in Q4, and add at least 10 GW during 2027. The upside case contemplates 20 GW, but management explicitly framed that as difficult and uncertain.", lang)
        base.add_paragraph(doc, "Cursor broadens the stack from infrastructure and models into applications. The acquisition carried a $60B implied equity value and $42.5B GAAP purchase consideration. Approximately $23.4B of equity awards are tied to post-acquisition service and will create material stock compensation over roughly three years. Cursor contributed $608M of Q3 revenue, but the retention package dilutes the translation from adjusted EBITDA to shareholder earnings.", lang)
        base.add_paragraph(doc, "Our thesis status is strengthened on demand and weakened on capital intensity. A backlog of $63B, including about $42B expected within twelve months, supports near-term visibility. Yet Q4 capex guidance of $35-40B is almost twice Q3 spending, and GPU, power, cooling, networking, customer concentration, contract repricing, and technology obsolescence can all compress returns. The debate is no longer whether SI can grow; it is whether return on invested capital remains attractive through a much larger depreciation cycle.", lang)
        base.add_figure(doc, 7, "SI revenue capex and nameplate compute", "Bars compare SI revenue and capex while the line shows nameplate compute capacity.", lang, 5.45, [("Q3 earnings release, 20 Oct 2026, pp. 6-7", EARNINGS_RELEASE), ("Q3 earnings call transcript, 20 Oct 2026, CFO remarks", TRANSCRIPT), ("Q2 Form 10-Q, filed 4 Aug 2026", FORM_10Q)])
    else:
        base.add_heading(doc, "5 SI变现验证需求 但尚未验证现金回报", lang)
        base.add_paragraph(doc, "SI收入增至75.60亿美元，其中解决方案与基础设施收入70.22亿美元、广告收入5.38亿美元。名义算力达到1.7GW，高于Q2的1.4GW和去年同期0.6GW。收入跃升说明新增容量正快速变现；管理层称第三方需求旺盛，并可在内部模型训练与外部云客户之间灵活分配算力。", lang)
        base.add_paragraph(doc, "板块调整后EBITDA为49.53亿美元，但资本开支仍达164.44亿美元，相当于季度收入的2.2倍。管理层计划年底前每月部署约20万至22万套GB300系统，Q4名义算力达到2.5-3.0GW，并在2027年新增至少10GW。上行情景为20GW，但管理层明确表示该目标困难且不确定。", lang)
        base.add_paragraph(doc, "Cursor把业务栈从基础设施和模型扩展至应用。收购隐含股权价值600亿美元，GAAP购买对价425亿美元；约234亿美元股权奖励与收购后服务挂钩，将在约三年产生大量股权激励。Cursor在Q3贡献6.08亿美元收入，但留任方案削弱调整后EBITDA向股东盈利的转化。", lang)
        base.add_paragraph(doc, "我们的投资逻辑在需求上增强、在资本强度上弱化。630亿美元积压订单中约420亿美元预计在未来十二个月确认，提供近期可见度。然而，Q4资本开支指引350-400亿美元，接近Q3两倍；GPU、电力、冷却、网络、客户集中、合同重新定价和技术淘汰均可能压缩回报。争论已不再是SI能否增长，而是在更大折旧周期中投入资本回报率能否保持吸引力。", lang)
        base.add_figure(doc, 7, "SI收入资本开支与名义算力", "柱状图比较SI收入和资本开支，折线展示名义算力。", lang, 5.45, [("Q3业绩新闻稿，2026年10月20日，第6-7页", EARNINGS_RELEASE), ("Q3业绩电话会文字稿，2026年10月20日，CFO发言", TRANSCRIPT), ("Q2 Form 10-Q，2026年8月4日提交", FORM_10Q)])

    # PAGE 7 - cash flow and balance sheet
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "6 Cash flow remains the limiting variable", lang)
        base.add_paragraph(doc, "Nine-month operating cash flow improved to $13.097B from $4.120B a year ago, supported by higher earnings, depreciation, stock compensation, and a $4.922B increase in deferred revenue. Capex nevertheless reached $48.069B, producing approximately $(34.972)B of free cash flow before acquisitions and marketable-securities activity. The business is operationally profitable but not self-funding at the current deployment pace.", lang)
        base.add_paragraph(doc, "Liquidity is substantial. SpaceX ended Q3 with $51.294B of cash and $38.561B of marketable securities against $38.844B of debt and finance leases, implying roughly $51.0B of net cash. The IPO contributed $85.675B of net financing proceeds during the first nine months. Those resources can absorb another investment surge, but the midpoint of Q4 capex guidance alone equals about 74% of quarter-end cash.", lang)
        base.add_paragraph(doc, "Working capital also deserves attention. Accounts receivable increased by $2.882B and inventory by $1.381B during the first nine months, while deferred revenue rose by $4.922B. Large upfront customer commitments improve cash visibility, but growth in receivables and inventory can consume cash rapidly as compute and satellite deployments accelerate. Backlog conversion is therefore as important as signed contract value.", lang)
        base.add_paragraph(doc, "Management cited operating cash flow, investment-grade debt markets, and partnerships across suppliers and customers as the funding stack for 2027. We consider that credible given the balance sheet and demand, but not costless. Higher debt, vendor financing, customer prepayments, and additional equity-linked compensation each transfer risk differently. Our valuation applies a discount for the possibility that capital needs grow faster than contract cash receipts.", lang)
        base.add_figure(doc, 8, "Nine-month operating cash flow capex and free cash flow", "Bars show operating cash flow, capital expenditures, and free cash flow for the nine months ended September 2026.", lang, 5.4, [("Q3 earnings release, 20 Oct 2026, pp. 11-12", EARNINGS_RELEASE), ("SEC Form 8-K and Exhibit 99.1", FORM_8K)])
    else:
        base.add_heading(doc, "6 现金流仍是限制变量", lang)
        base.add_paragraph(doc, "九个月经营现金流由去年同期41.20亿美元改善至130.97亿美元，受盈利增加、折旧、股权激励和递延收入增加49.22亿美元支持。但资本开支达到480.69亿美元，自由现金流约为负349.72亿美元（未计并购及有价证券活动）。业务在经营层面已盈利，但按当前部署速度仍无法自我融资。", lang)
        base.add_paragraph(doc, "流动性充足。Q3末现金512.94亿美元、有价证券385.61亿美元，债务及融资租赁388.44亿美元，净现金约510亿美元。前九个月IPO净融资所得856.75亿美元。这些资源可承受又一轮投资激增，但Q4资本开支指引中点已相当于季末现金的74%。", lang)
        base.add_paragraph(doc, "营运资本也值得关注。前九个月应收账款增加28.82亿美元、库存增加13.81亿美元，递延收入增加49.22亿美元。大额客户预付款改善现金可见度，但随着算力和卫星部署加速，应收和库存也会迅速消耗现金。因此，积压订单转化速度与已签合同价值同等重要。", lang)
        base.add_paragraph(doc, "管理层把经营现金流、投资级债务市场以及供应商和客户合作伙伴列为2027年融资组合。鉴于资产负债表和需求，我们认为路径可信，但并非无成本。更高债务、供应商融资、客户预付款和更多股权挂钩激励，以不同方式转移风险。估值对资本需求增长快于合同现金回款的可能性给予折价。", lang)
        base.add_figure(doc, 8, "九个月经营现金流资本开支与自由现金流", "柱状图展示截至2026年9月的九个月经营现金流、资本开支与自由现金流。", lang, 5.4, [("Q3业绩新闻稿，2026年10月20日，第11-12页", EARNINGS_RELEASE), ("SEC Form 8-K及附件99.1", FORM_8K)])

    # PAGE 8 - guidance
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "7 Guidance moves the debate from quarter to capacity conversion", lang)
        base.add_paragraph(doc, "Management introduced Q1 2027 revenue guidance of $33-37B, more than double Q3 actual revenue. It also reiterated a target to exit 2026 at a $100B annualized revenue run-rate based on December revenue. The difference between a run-rate target and recognized quarterly revenue is important: the ARR target implies roughly $8.3B of December revenue, while the Q1 guidance midpoint implies $11.7B per month on average. Both require a rapid contract-conversion curve.", lang)
        base.add_paragraph(doc, "Q4 capex guidance of $35-40B is the clearest near-term risk. At the midpoint, full-year 2026 capex would approach $85.6B, above our prior $65B estimate. Management expects gross margin to normalize toward Q2 because Q3 included one-time benefits and favorable customer mix. We therefore forecast a sharp Q4 revenue increase but lower margin than Q3, followed by renewed operating leverage as Q1 2027 capacity ramps.", lang)
        base.add_paragraph(doc, "Operational milestones support the revenue plan but are aggressive. Management expects 2.5-3.0 GW of compute by year-end, at least 10 GW of incremental capacity during 2027, weekly Starship launches in H1 2027, and twice-weekly cadence later in the year. A few hundred V3 satellites should begin producing noticeable Starlink service improvements around mid-2027. Each milestone requires coordinated execution across power, cooling, chips, launch licensing, manufacturing, and network operations.", lang)
        base.add_paragraph(doc, "We rate guidance credibility as medium-high on demand and medium on timing. The $63B backlog and strong Q3 conversion substantiate customer demand, while the balance sheet supports procurement. However, management is simultaneously scaling cloud infrastructure, Starship, V3 satellites, Starlink Mobile, and internal chip fabrication. The number of critical paths makes schedule slippage likely even if long-term demand remains intact.", lang)
        base.add_figure(doc, 9, "Revenue and capital spending outlook", "Two-panel chart comparing Q3 revenue with Q1 2027 guidance and Q3 capex with Q4 guidance.", lang, 5.55, [("Q3 earnings call transcript, 20 Oct 2026, CFO guidance", TRANSCRIPT), ("Q3 earnings release, 20 Oct 2026", EARNINGS_RELEASE), ("Q2 earnings webcast, 4 Aug 2026, prior ARR target", Q2_EVENT)])
    else:
        base.add_heading(doc, "7 指引把争论从季度业绩转向产能转化", lang)
        base.add_paragraph(doc, "管理层首次给出2027Q1收入330-370亿美元指引，超过Q3实际收入两倍；同时重申按12月收入计算，2026年底达到1000亿美元年化收入运行速度。运行速度目标与季度确认收入的差异很重要：ARR目标对应12月单月约83亿美元收入，而Q1指引中点对应月均约117亿美元。两者都要求合同快速转化。", lang)
        base.add_paragraph(doc, "Q4资本开支350-400亿美元是最明确的近期风险。按中点计算，2026全年资本开支将接近856亿美元，高于我们此前650亿美元预测。管理层预计Q4毛利率向Q2回归，因为Q3包含一次性收益和有利客户结构。因此，我们预测Q4收入大幅增长但利润率低于Q3，随后随着2027Q1产能爬坡重新体现经营杠杆。", lang)
        base.add_paragraph(doc, "经营里程碑支持收入计划，但目标激进。管理层预计年末算力2.5-3.0GW、2027年至少新增10GW、2027年上半年星舰周度发射、下半年提升至每周两次。数百颗V3卫星有望在2027年中开始明显改善星链服务。每项里程碑都要求电力、冷却、芯片、发射许可、制造和网络运营协同执行。", lang)
        base.add_paragraph(doc, "我们对指引的需求可信度评为中高，对时间可信度评为中等。630亿美元积压订单和Q3强劲转化验证客户需求，资产负债表也支持采购。然而，公司同时扩张云基础设施、星舰、V3卫星、Starlink Mobile及内部芯片制造。关键路径数量使进度延误的概率较高，即便长期需求不变。", lang)
        base.add_figure(doc, 9, "收入与资本开支展望", "双面板图，对比Q3收入与2027Q1指引，以及Q3资本开支与Q4指引。", lang, 5.55, [("Q3业绩电话会文字稿，2026年10月20日，CFO指引", TRANSCRIPT), ("Q3业绩新闻稿，2026年10月20日", EARNINGS_RELEASE), ("Q2业绩网络直播，2026年8月4日，此前ARR目标", Q2_EVENT)])

    # PAGE 9 - estimates
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "8 Estimate revisions incorporate Q1 visibility and Q4 spending", lang)
        base.add_paragraph(doc, "We raise FY2026 revenue to $50.5B from $35.0B and adjusted EBITDA to $29.0B from $15.5B. Nine-month reported revenue of $27.1B leaves $23.4B for Q4 in our model, consistent with the year-end ARR target but below a simple extrapolation of Q1 2027 guidance. We lift FY2026 capex to $85.5B from $65.0B, reflecting management's $35-40B Q4 range.", lang)
        headers = ["Metric", "FY26 old", "FY26 new", "Change", "FY27 old", "FY27 new"]
        rows = [
            ["Revenue", "$35.0B", "$50.5B", "+44%", "$58.0B", "$142.0B"],
            ["Adjusted EBITDA", "$15.5B", "$29.0B", "+87%", "$26.0B", "$72.0B"],
            ["Operating income", "$(3.0)B", "$8.5B", "NM", "$4.0B", "$31.0B"],
            ["Net income", "$(6.0)B", "$9.1B", "NM", "$0.0B", "$27.0B"],
            ["GAAP diluted EPS", "$(0.50)", "$0.69", "NM", "$0.00", "$1.95"],
            ["Capital expenditures", "$65.0B", "$85.5B", "+32%", "$75.0B", "$150.0B"],
        ]
        base.add_table(doc, headers, rows, [1850, 1300, 1300, 1050, 1300, 2560], lang)
        base.add_source_line(doc, lang, "Old estimates from independent post-Q2 update dated 5 Aug 2026; new estimates dated 21 Oct 2026", [("Q3 earnings release", EARNINGS_RELEASE), ("Q3 earnings call guidance", TRANSCRIPT), ("Q2 Form 10-Q baseline", FORM_10Q)])
        base.add_paragraph(doc, "For FY2027, we raise revenue to $142B and adjusted EBITDA to $72B. The revenue forecast assumes Q1 guidance is achieved near the midpoint, followed by continued SI and Connectivity growth but some sequential normalization after the initial capacity ramp. We forecast $31B of operating income and $27B of net income, below management's most ambitious ARR scenarios and before any Tesla combination.", lang)
        base.add_paragraph(doc, "Our FY2027 capex estimate rises to $150B because 10 GW of incremental compute, higher Starship cadence, V3 deployment, and mobile-spectrum integration require sustained spending. The model therefore shows strong accounting earnings alongside negative free cash flow. The largest sensitivities are SI utilization, realized cloud pricing, power availability, depreciation life, stock compensation, and the timing of customer prepayments.", lang)
        base.add_figure(doc, 10, "Old and post-Q3 estimates", "Horizontal bars compare post-Q2 and post-Q3 estimates for revenue, adjusted EBITDA, and capex in FY2026 and FY2027.", lang, 5.5, [("Independent post-Q2 estimates, 5 Aug 2026", FORM_10Q), ("Q3 results and guidance, 20 Oct 2026", EARNINGS_RELEASE), ("Q3 earnings call transcript", TRANSCRIPT)])
    else:
        base.add_heading(doc, "8 预测修正纳入Q1可见度和Q4支出", lang)
        base.add_paragraph(doc, "我们将FY2026收入由350亿美元上调至505亿美元，调整后EBITDA由155亿美元上调至290亿美元。九个月已报告收入270.83亿美元，模型对Q4收入预测为234亿美元，与年末ARR目标一致，但低于简单外推2027Q1指引。考虑管理层Q4 350-400亿美元区间，FY2026资本开支由650亿美元上调至855亿美元。", lang)
        headers = ["指标", "FY26旧", "FY26新", "变化", "FY27旧", "FY27新"]
        rows = [
            ["收入", "$350亿", "$505亿", "+44%", "$580亿", "$1,420亿"],
            ["调整后EBITDA", "$155亿", "$290亿", "+87%", "$260亿", "$720亿"],
            ["营业利润", "$(30)亿", "$85亿", "不适用", "$40亿", "$310亿"],
            ["净利润", "$(60)亿", "$91亿", "不适用", "$0亿", "$270亿"],
            ["GAAP摊薄EPS", "$(0.50)", "$0.69", "不适用", "$0.00", "$1.95"],
            ["资本开支", "$650亿", "$855亿", "+32%", "$750亿", "$1,500亿"],
        ]
        base.add_table(doc, headers, rows, [1850, 1300, 1300, 1050, 1300, 2560], lang)
        base.add_source_line(doc, lang, "旧预测来自2026年8月5日独立Q2后更新；新预测日期为2026年10月21日", [("Q3业绩新闻稿", EARNINGS_RELEASE), ("Q3业绩电话会指引", TRANSCRIPT), ("Q2 Form 10-Q基线", FORM_10Q)])
        base.add_paragraph(doc, "FY2027方面，我们将收入上调至1,420亿美元、调整后EBITDA上调至720亿美元。收入预测假设Q1接近指引中点，随后SI和连接业务继续增长，但初始产能爬坡后环比增速正常化。我们预测营业利润310亿美元、净利润270亿美元，低于管理层最激进的ARR情景，且不计任何特斯拉合并。", lang)
        base.add_paragraph(doc, "FY2027资本开支预测上调至1,500亿美元，因为新增10GW算力、更高星舰发射频率、V3部署和移动频谱整合需要持续支出。因此，模型同时呈现强劲会计盈利和负自由现金流。最大敏感性包括SI利用率、云服务实际定价、电力可用性、折旧年限、股权激励及客户预付款时点。", lang)
        base.add_figure(doc, 10, "旧预测与Q3后预测", "横向柱状图，对比FY2026和FY2027收入、调整后EBITDA及资本开支的Q2后与Q3后预测。", lang, 5.5, [("2026年8月5日独立Q2后预测", FORM_10Q), ("2026年10月20日Q3业绩及指引", EARNINGS_RELEASE), ("Q3业绩电话会文字稿", TRANSCRIPT)])

    # PAGE 10 - valuation
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "9 Valuation supports a higher target but not a narrow range", lang)
        base.add_paragraph(doc, "We raise our 12-month price target to $210 from $150 and maintain Outperform. The target implies an equity value of approximately $2.62T on the current diluted share base. At the 20 October regular close of $147.79, it represents 42% upside; at the $159.71 after-hours close, upside is 31%. The market is already capitalizing substantial long-duration optionality, so our rating depends on continued capacity conversion rather than multiple expansion alone.", lang)
        base.add_paragraph(doc, "Our SOTP assigns approximately $650B to Connectivity, reflecting its recurring revenue, improving ARPU, 45% operating margin, and mobile optionality. We assign $250B to Space, including launch leadership and explicit Starship option value. SI receives $1.67T, supported by the $63B backlog, Q1 2027 revenue guidance, positive Q3 operating income, and the possibility of rapid compute scaling. Adding roughly $51B of net cash and modest corporate adjustments produces the $2.62T equity value.", lang)
        base.add_paragraph(doc, "The base target equates to about 18.5x FY2027 revenue and 36x FY2027 adjusted EBITDA. Those multiples are extreme for mature infrastructure but lower than the current value placed on FY2026 earnings because our model assumes a large step-up. We do not use near-term P/E as the primary tool: below-the-line volatility, acquisition stock compensation, and a rapidly growing share base make it unstable.", lang)
        base.add_paragraph(doc, "The $110 bear case assumes slower SI utilization, lower cloud pricing, delayed power and Starship milestones, and no strategic premium for a Tesla combination. The $320 bull case assumes Q1 guidance is exceeded, 2027 compute approaches the high end of management's capacity ambition, Connectivity sustains current margins, and Starship achieves rapid reuse. No merger value is included in the $210 base case because the special committees have not agreed terms and shareholder approval is uncertain.", lang)
        base.add_figure(doc, 11, "Twelve-month valuation scenarios", "Bar chart showing bear, base, and bull per-share valuation scenarios with the 20 October close marked.", lang, 5.55, [("Independent analyst SOTP dated 21 Oct 2026", EARNINGS_RELEASE), ("Q3 earnings call transcript, 20 Oct 2026", TRANSCRIPT), ("Longbridge close and market cap, accessed 21 Oct 2026", LONG_BRIDGE)])
    else:
        base.add_heading(doc, "9 估值支持更高目标价 但无法支持窄区间", lang)
        base.add_paragraph(doc, "我们将12个月目标价由$150上调至$210，并维持跑赢大盘。按当前摊薄股本，目标价对应约2.62万亿美元股权价值。相对10月20日常规收盘价$147.79，上涨空间42%；相对盘后$159.71，上涨空间31%。市场已计入大量长期期权，因此评级依赖产能持续转化，而非单纯估值扩张。", lang)
        base.add_paragraph(doc, "SOTP中，连接业务估值约6,500亿美元，反映其经常性收入、ARPU改善、45%营业利润率及移动业务期权。太空业务估值2,500亿美元，包括发射领先和明确的星舰期权价值。SI估值1.67万亿美元，依据630亿美元积压订单、2027Q1收入指引、Q3营业利润转正及算力快速扩张的可能性。加上约510亿美元净现金及少量公司层面调整，得到2.62万亿美元股权价值。", lang)
        base.add_paragraph(doc, "基准目标价相当于约18.5倍FY2027收入和36倍FY2027调整后EBITDA。对成熟基础设施而言倍数极高，但因模型假设盈利大幅跃升，低于当前基于FY2026盈利的估值。我们不以近期市盈率为主要工具：线下项目波动、并购股权激励和快速增长的股本使其不稳定。", lang)
        base.add_paragraph(doc, "$110熊市情景假设SI利用率较低、云服务定价下降、电力和星舰里程碑延误，且不计特斯拉合并战略溢价。$320牛市情景假设Q1指引被超越、2027算力接近管理层产能雄心高端、连接业务维持当前利润率、星舰实现快速复用。$210基准情景不计合并价值，因为特别委员会尚未就条款达成一致，股东批准也不确定。", lang)
        base.add_figure(doc, 11, "十二个月估值情景", "柱状图展示每股熊市、基准和牛市估值情景，并标记10月20日收盘价。", lang, 5.55, [("2026年10月21日独立分析师SOTP", EARNINGS_RELEASE), ("Q3业绩电话会文字稿，2026年10月20日", TRANSCRIPT), ("Longbridge收盘价与市值，2026年10月21日获取", LONG_BRIDGE)])

    # PAGE 11 - thesis, catalysts, risks, call
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "10 Updated thesis catalysts and risks", lang)
        base.add_bullet(doc, num_id, "Thesis pillar 1 Connectivity scale strengthened.", "Subscriber growth, sequential ARPU improvement, and 99% operating-income growth confirm that Starlink is a high-quality recurring earnings engine. Enterprise, aviation, government, mobile, and V3 capacity broaden the opportunity beyond consumer broadband.", lang)
        base.add_bullet(doc, num_id, "Thesis pillar 2 Vertical integration strengthened.", "Flight 14 connected Starship economics directly to V3 deployment, while SI combines compute, models, and applications after Cursor. The same engineering and capital base is increasingly shared across launch, connectivity, and compute.", lang)
        base.add_bullet(doc, num_id, "Thesis pillar 3 Self-funded compounding remains weakened.", "Despite positive GAAP operating income, nine-month free cash flow was about $(35)B and Q4 capex guidance is $35-40B. The company still depends on IPO cash, debt markets, supplier and customer financing, and stock compensation.", lang)
        base.add_heading(doc, "Catalysts", lang, 2)
        base.add_paragraph(doc, "The next twelve months include Starship Flight 15 and a planned dual-stage catch attempt; progress toward weekly Starship launches; visible V3 service improvements; the 800 MHz spectrum transaction; Q4 year-end ARR evidence; Q1 2027 revenue delivery; Cursor enterprise adoption; and updates on Terafab test chips. Each can materially change either the revenue curve or the cost of capital.", lang)
        base.add_heading(doc, "Principal risks", lang, 2)
        base.add_paragraph(doc, "The leading risks are SI customer concentration and contract duration, cloud pricing, GPU and power availability, depreciation and obsolescence, launch failures or licensing delays, spectrum approval, Starlink competition, government-contract timing, acquisition integration, stock-based compensation, dual-class governance, key-person dependence, and further capital needs. A potential Tesla transaction adds conflict, valuation, approval, dilution, and execution risks before any synergy is realized.", lang)
        base.add_heading(doc, "Call takeaways", lang, 2)
        base.add_paragraph(doc, "Management's tone was highly confident on demand and technically ambitious on capacity. It framed 10 GW of incremental 2027 compute as highly likely and 20 GW as possible but difficult. The company disclosed a rough 50/50 target for internal vs. external compute allocation, with flexibility to shift. On a Tesla combination, Elon Musk confirmed independent special committees but emphasized that no terms have been agreed and a transaction is not certain. Our estimates exclude merger effects.", lang)
        base.add_source_line(doc, lang, "Call held 20 October 2026; transcript date matches the earnings release", [("Official earnings event and replay", IR_EVENT), ("Q3 earnings call transcript", TRANSCRIPT), ("Associated Press call coverage", AP)])
    else:
        base.add_heading(doc, "10 更新后的投资逻辑催化剂与风险", lang)
        base.add_bullet(doc, num_id, "投资逻辑支柱1 连接业务规模化增强。", "用户增长、ARPU环比改善及营业利润增长99%确认星链是高质量经常性盈利引擎。企业、航空、政府、移动通信及V3容量把机会扩展至消费者宽带之外。", lang)
        base.add_bullet(doc, num_id, "投资逻辑支柱2 垂直整合增强。", "Flight 14把星舰经济性直接连接至V3部署；Cursor收购后，SI同时覆盖算力、模型和应用。发射、连接与算力正更多共享同一工程与资本基础。", lang)
        base.add_bullet(doc, num_id, "投资逻辑支柱3 自我融资复利仍弱化。", "尽管GAAP营业利润转正，九个月自由现金流仍约负350亿美元，Q4资本开支指引为350-400亿美元。公司仍依赖IPO现金、债务市场、供应商及客户融资以及股权激励。", lang)
        base.add_heading(doc, "催化剂", lang, 2)
        base.add_paragraph(doc, "未来十二个月催化剂包括星舰Flight 15和两级回收尝试、向周度星舰发射推进、V3服务改善、800MHz频谱交易、Q4年末ARR验证、2027Q1收入兑现、Cursor企业采用，以及Terafab测试芯片更新。每项都可能显著改变收入曲线或资本成本。", lang)
        base.add_heading(doc, "主要风险", lang, 2)
        base.add_paragraph(doc, "主要风险包括SI客户集中和合同期限、云服务定价、GPU与电力可用性、折旧与淘汰、发射失败或许可延误、频谱审批、星链竞争、政府合同节奏、并购整合、股权激励、双重股权治理、关键人物依赖及进一步融资需求。潜在特斯拉交易在协同兑现前还会增加利益冲突、估值、批准、摊薄和执行风险。", lang)
        base.add_heading(doc, "电话会要点", lang, 2)
        base.add_paragraph(doc, "管理层对需求语气高度自信，对产能目标极具技术雄心。公司把2027新增10GW算力描述为高度可能，把20GW描述为可能但困难；内部与外部算力分配大致目标为50/50，但可动态调整。关于与特斯拉合并，马斯克确认双方成立独立特别委员会，但强调尚未谈妥条款，交易并不确定。我们的预测不计合并影响。", lang)
        base.add_source_line(doc, lang, "电话会于2026年10月20日举行；文字稿日期与业绩新闻稿一致", [("官方业绩活动及回放", IR_EVENT), ("Q3业绩电话会文字稿", TRANSCRIPT), ("美联社电话会报道", AP)])

    # PAGE 12 - sources
    base.add_page_break(doc)
    if lang == "en":
        base.add_heading(doc, "Sources and references", lang)
        base.add_paragraph(doc, "Primary earnings materials", lang, 10.5, True, "000000", 4)
        sources = [
            ("SpaceX Q3 2026 earnings release and unaudited financial statements, 20 October 2026", EARNINGS_RELEASE),
            ("SEC Form 8-K and filing index, filed 20 October 2026", FORM_8K),
            ("Official Q3 2026 earnings event, webcast, Q&A, and earnings update, 20 October 2026", IR_EVENT),
            ("Q3 2026 earnings call transcript, 20 October 2026", TRANSCRIPT),
            ("Form 10-Q for quarter ended 30 June 2026, filed 4 August 2026", FORM_10Q),
            ("Official Q2 2026 earnings event and prior outlook, 4 August 2026", Q2_EVENT),
        ]
        for text, url in sources:
            add_source_item(doc, text, url, lang)
        base.add_paragraph(doc, "Consensus and market context", lang, 10.5, True, "000000", 4, before=6)
        sources2 = [
            ("TipRanks Q3 preview and pre-earnings consensus, 19 October 2026", CONSENSUS),
            ("Zacks key-metric consensus and results comparison, 20 October 2026", SEGMENT_CONSENSUS),
            ("Associated Press results and call coverage, 20 October 2026", AP),
            ("Longbridge live quote and calculated market data, accessed 21 October 2026", LONG_BRIDGE),
            ("Yahoo Finance yfinance check attempted 21 October 2026; unavailable, so dynamic fields are N/A", "https://finance.yahoo.com/quote/SPCX/"),
        ]
        for text, url in sources2:
            add_source_item(doc, text, url, lang)
        base.add_paragraph(doc, "Methodology notes", lang, 10.5, True, "000000", 4, before=6)
        base.add_paragraph(doc, "Beat and miss calculations use public pre-earnings consensus dated 19 October 2026. Free cash flow equals operating cash flow less purchases of property, plant, and equipment. Adjusted EBITDA is company-defined and excludes depreciation, share-based compensation, interest, taxes, and other items. Forecasts, SOTP, rating, and price target are independent analyst estimates, not company guidance. Dynamic yfinance fields are shown as N/A because the 21 October query failed; valuation references instead use the independently retrieved 20 October Longbridge close. The Q3 Form 10-Q had not been filed as of publication; Q3 figures use the SEC-filed 8-K and Exhibit 99.1, while the Q2 Form 10-Q supplies detailed accounting and historical context.", lang, size=8.5)
        add_disclaimer(doc, lang)
    else:
        base.add_heading(doc, "资料来源与参考文献", lang)
        base.add_paragraph(doc, "主要业绩材料", lang, 10.5, True, "000000", 4)
        sources = [
            ("SpaceX Q3 2026业绩新闻稿及未经审计财务报表，2026年10月20日", EARNINGS_RELEASE),
            ("SEC Form 8-K及申报索引，2026年10月20日提交", FORM_8K),
            ("官方Q3 2026业绩活动、网络直播、问答及业绩更新，2026年10月20日", IR_EVENT),
            ("Q3 2026业绩电话会文字稿，2026年10月20日", TRANSCRIPT),
            ("截至2026年6月30日季度Form 10-Q，2026年8月4日提交", FORM_10Q),
            ("官方Q2 2026业绩活动及此前展望，2026年8月4日", Q2_EVENT),
        ]
        for text, url in sources:
            add_source_item(doc, text, url, lang)
        base.add_paragraph(doc, "一致预期与市场背景", lang, 10.5, True, "000000", 4, before=6)
        sources2 = [
            ("TipRanks Q3前瞻及财报前一致预期，2026年10月19日", CONSENSUS),
            ("Zacks关键指标一致预期及业绩对比，2026年10月20日", SEGMENT_CONSENSUS),
            ("美联社业绩及电话会报道，2026年10月20日", AP),
            ("Longbridge实时行情及计算市场数据，2026年10月21日获取", LONG_BRIDGE),
            ("Yahoo Finance yfinance于2026年10月21日尝试查询但不可用，动态字段因此显示N/A", "https://finance.yahoo.com/quote/SPCX/"),
        ]
        for text, url in sources2:
            add_source_item(doc, text, url, lang)
        base.add_paragraph(doc, "方法说明", lang, 10.5, True, "000000", 4, before=6)
        base.add_paragraph(doc, "超预期和不及预期计算采用截至2026年10月19日的公开财报前一致预期。自由现金流等于经营现金流减不动产、厂房和设备采购。调整后EBITDA为公司定义指标，排除折旧、股权激励、利息、税项及其他项目。预测、SOTP、评级和目标价均为独立分析师估计，并非公司指引。2026年10月21日yfinance动态查询失败，因此相关字段显示N/A；估值改用独立取得的10月20日Longbridge收盘价。截至发布日期，Q3 Form 10-Q尚未提交；Q3数据来自SEC备案的8-K及附件99.1，Q2 Form 10-Q用于详细会计和历史背景。", lang, size=8.5)
        add_disclaimer(doc, lang)

    filename = "SPCX_Q3_2026_Earnings_Update.docx" if lang == "en" else "SPCX_Q3_2026_业绩更新报告_中文版.docx"
    path = OUT / filename
    doc.save(path)
    print(f"Saved {path}")
    return path


if __name__ == "__main__":
    market_data = fetch_market_data()
    build_report("en", market_data)
    build_report("cn", market_data)
