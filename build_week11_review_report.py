from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


OUT = Path(r"D:\PycharmProjects\PythonProject6\outputs\week11_simulated_dataset")
DATA = json.loads((OUT / "review_analysis.json").read_text(encoding="utf-8"))
DOCX = OUT / "模拟数据复盘分析报告.docx"

NAVY = "1F4E78"
BLUE = "4472C4"
LIGHT_BLUE = "D9EAF7"
LIGHT_GRAY = "F2F4F7"
MID_GRAY = "667085"
INK = "1F2937"
GREEN = "70AD47"
AMBER = "ED7D31"
RED = "C00000"
WHITE = "FFFFFF"


def shade(cell, color):
    tc_pr = cell._tc.get_or_add_tcPr()
    node = tc_pr.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd")
        tc_pr.append(node)
    node.set(qn("w:fill"), color)


def width(cell, dxa):
    tc_pr = cell._tc.get_or_add_tcPr()
    node = tc_pr.find(qn("w:tcW"))
    if node is None:
        node = OxmlElement("w:tcW")
        tc_pr.append(node)
    node.set(qn("w:w"), str(dxa))
    node.set(qn("w:type"), "dxa")


def margins(cell, top=80, start=100, bottom=80, end=100):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def fixed_table(table, widths):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    pr = table._tbl.tblPr
    layout = pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    tw = pr.find(qn("w:tblW"))
    if tw is None:
        tw = OxmlElement("w:tblW")
        pr.append(tw)
    tw.set(qn("w:w"), str(sum(widths)))
    tw.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for value in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(value))
        grid.append(col)
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        if tr_pr.find(qn("w:cantSplit")) is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        for i, cell in enumerate(row.cells):
            width(cell, widths[i])
            margins(cell)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def font(run, size=10.5, bold=False, color=INK, italic=False):
    run.font.name = "Microsoft YaHei"
    rfonts = run._element.get_or_add_rPr().rFonts
    for key in ("eastAsia", "ascii", "hAnsi"):
        rfonts.set(qn(f"w:{key}"), "Microsoft YaHei")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def para_style(p, before=0, after=6, line=1.15, align=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if align is not None:
        p.alignment = align


def body(doc, text, bold_prefix=None, after=6):
    p = doc.add_paragraph()
    para_style(p, after=after)
    if bold_prefix and text.startswith(bold_prefix):
        font(p.add_run(bold_prefix), bold=True)
        font(p.add_run(text[len(bold_prefix):]))
    else:
        font(p.add_run(text))
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    para_style(p, after=4)
    font(p.add_run(text))
    return p


def heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    font(p.add_run(text), size={1: 16, 2: 13, 3: 11.5}[level], bold=True,
         color=NAVY if level < 3 else "2F5597")
    return p


def table(doc, headers, rows, widths, font_size=8.8):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    fixed_table(t, widths)
    tr_pr = t.rows[0]._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        tr_pr.append(OxmlElement("w:tblHeader"))
    for i, value in enumerate(headers):
        c = t.rows[0].cells[i]
        shade(c, BLUE)
        p = c.paragraphs[0]
        para_style(p, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        font(p.add_run(str(value)), size=font_size, bold=True, color=WHITE)
    for r_index, values in enumerate(rows):
        cells = t.add_row().cells
        for i, value in enumerate(values):
            p = cells[i].paragraphs[0]
            para_style(p, after=0, line=1.0,
                       align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT)
            font(p.add_run(str(value)), size=font_size)
        if r_index % 2:
            for c in cells:
                shade(c, "F8FAFC")
    fixed_table(t, widths)
    return t


def callout(doc, title, text, color=LIGHT_BLUE):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    fixed_table(t, [9360])
    c = t.cell(0, 0)
    shade(c, color)
    p = c.paragraphs[0]
    para_style(p, after=2)
    font(p.add_run(title + "："), bold=True, color=NAVY)
    font(p.add_run(text))
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def page_number(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    font(p.add_run("第 "), size=9, color=MID_GRAY)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    p._p.append(fld)
    font(p.add_run(" 页"), size=9, color=MID_GRAY)


def pct(v):
    return f"{v:.1f}%"


O = DATA["overall"]
G = DATA["groupMetrics"]
F = DATA["functionMetrics"]
Q = DATA["frequencyMetrics"]
A = DATA["anomalyMetrics"]

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.8)
sec.bottom_margin = Inches(0.75)
sec.left_margin = Inches(0.85)
sec.right_margin = Inches(0.85)

styles = doc.styles
styles["Normal"].font.name = "Microsoft YaHei"
styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
styles["Normal"].font.size = Pt(10.5)
for level in (1, 2, 3):
    st = styles[f"Heading {level}"]
    st.font.name = "Microsoft YaHei"
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    st.paragraph_format.space_before = Pt({1: 14, 2: 10, 3: 8}[level])
    st.paragraph_format.space_after = Pt({1: 7, 2: 5, 3: 4}[level])

header = sec.header.paragraphs[0]
font(header.add_run("百度办公 AI 助手｜第11周交付材料"), size=9, color=MID_GRAY)
page_number(sec.footer.paragraphs[0])

# Cover
p = doc.add_paragraph()
para_style(p, before=50, after=8)
font(p.add_run("WEEK 11 · DATA REVIEW"), size=10, bold=True, color=BLUE)
p = doc.add_paragraph()
para_style(p, after=12)
font(p.add_run("模拟数据复盘分析报告"), size=27, bold=True, color=NAVY)
p = doc.add_paragraph()
para_style(p, after=28)
font(p.add_run("基于300名模拟用户、1500次会话与3000条使用事件"), size=13, color=MID_GRAY)
table(doc, ["项目", "内容"], [
    ["项目名称", "百度办公 AI 助手 MVP"],
    ["复盘周期", "2026年8月1日—2026年8月30日"],
    ["分析范围", "用户群体、功能表现、使用频率、反馈与异常风险"],
    ["数据性质", "规则生成的模拟数据，不代表百度真实业务表现"],
    ["版本", "V1.0｜2026年8月19日"],
], [1850, 7510], font_size=9.5)
doc.add_paragraph()
callout(doc, "重要说明", "本报告用于验证指标体系、复盘方法和优化决策流程。所有用户、行为、反馈及异常记录均为模拟数据，不包含真实个人信息，不应被解释为真实运营结论。", "FFF2CC")
doc.add_page_break()

heading(doc, "摘要", 1)
body(doc, f"本报告基于300名模拟用户、1500次模拟会话、3000条功能使用事件、500条反馈记录和100条异常记录，对百度办公 AI 助手 MVP 进行阶段性数据复盘。复盘从总体表现、用户群体、核心功能、使用频率、用户反馈与异常风险六个角度展开，目的不是证明产品已经达到正式上线标准，而是检验既有指标体系是否能够发现问题并支持后续优化。")
body(doc, f"模拟结果显示，系统总体生成成功率为{pct(O['successRate'])}，平均响应时间为{O['avgResponseSeconds']:.1f}秒，任务完成率为{pct(O['completionRate'])}，有效反馈平均评分为{O['avgScore']:.2f}分。会议纪要与自然语言咨询表现相对稳定；任务自动化、智能推荐和任务拆解是主要改进对象。管理人员群体的生成成功率和再次使用意愿低于其他群体，说明高复杂度、强决策型场景仍需加强。")
callout(doc, "综合判断", "当前模拟结果可支持产品继续迭代，但不适合直接作为真实发布决策依据。建议优先修复接口稳定性、任务字段完整性和任务自动化状态更新，再优化推荐解释、文档来源定位与管理人员场景体验。")

heading(doc, "一、复盘背景与目标", 1)
heading(doc, "1.1 复盘背景", 2)
body(doc, "第10周已完成数据看板原型、数据复盘机制、MVP功能测试、AI能力测试和竞品对比。本周在此基础上构建模拟用户数据集，并将指标体系应用到一轮完整复盘中，以验证数据字段能否关联、指标口径能否计算、问题能否被识别以及优化建议能否落到具体模块。")
heading(doc, "1.2 复盘目标", 2)
for x in [
    "验证模拟数据集能否覆盖用户、会话、事件、反馈和异常五类分析对象；",
    "比较不同用户群体、功能和使用频率下的产品表现；",
    "识别影响成功率、完成率、满意度和再次使用意愿的关键问题；",
    "形成产品功能、AI融合、用户体验和产品增长四个方向的迭代依据。",
]: bullet(doc, x)

heading(doc, "二、数据范围与分析口径", 1)
heading(doc, "2.1 数据范围", 2)
table(doc, ["数据对象", "记录数", "主要用途"], [
    ["模拟用户", "300", "用户群体、使用频率和偏好分析"],
    ["模拟会话", "1500", "访问规模与会话层级分析"],
    ["使用事件", "3000", "成功率、耗时、完成率、保存率和导出率计算"],
    ["用户反馈", "500", "评分、帮助率、再次使用意愿和标签分析"],
    ["异常记录", "100", "异常类型、严重程度和处理状态分析"],
], [1700, 1200, 6460])
heading(doc, "2.2 核心指标口径", 2)
table(doc, ["指标", "计算口径", "用途"], [
    ["生成成功率", "成功生成事件数÷生成事件总数", "判断功能和AI调用稳定性"],
    ["平均响应时间", "全部事件响应时间的算术平均值", "评估等待成本"],
    ["任务完成率", "成功事件中完成任务数÷成功事件数", "判断结果是否支持用户完成工作"],
    ["保存率", "成功事件中保存结果数÷成功事件数", "观察结果沉淀意愿"],
    ["导出率", "具备导出场景的成功事件中导出数÷成功数", "观察成果交付需求"],
    ["再次使用意愿", "有效反馈中选择愿意再次使用的比例", "衡量持续使用意向"],
], [1700, 4050, 3610])
body(doc, "说明：反馈指标仅基于500条反馈样本，异常指标仅基于抽取出的100条异常记录。两者不能代替对全部事件的完整统计。", after=8)

heading(doc, "三、总体运行表现", 1)
heading(doc, "3.1 核心结果", 2)
table(doc, ["指标", "模拟结果", "判断"], [
    ["生成成功率", pct(O['successRate']), "总体可用，但仍有296次失败"],
    ["平均响应时间", f"{O['avgResponseSeconds']:.1f}秒", "达到30秒内目标，仍需关注慢场景"],
    ["任务完成率", pct(O['completionRate']), "多数结果可支持完成任务"],
    ["保存率", pct(O['saveRate']), "超过半数成功结果被保存"],
    ["导出率", pct(O['exportRate']), "受功能是否支持导出影响，不宜孤立判断"],
    ["平均评分", f"{O['avgScore']:.2f}/5", "整体满意度较好"],
    ["再次使用意愿", pct(O['reuseIntentRate']), "具备持续使用基础"],
], [2100, 1800, 5460])
heading(doc, "3.2 总体问题", 2)
body(doc, "3000条事件中共有296条失败，失败率为9.9%。失败原因主要集中在接口调用失败、生成超时和字段缺失。接口调用失败105次，占全部失败事件的35.5%，说明外部模型调用和网络链路仍是首要稳定性风险。")
table(doc, ["失败原因", "次数", "占全部失败事件"], [
    [name, str(count), pct(count / O['errorCount'] * 100)]
    for name, count in list(O['topErrors'].items())[:6]
], [3700, 1800, 3860])

heading(doc, "四、用户群体复盘", 1)
heading(doc, "4.1 分群结果", 2)
group_rows = []
for name in ["职场新人", "普通职员", "项目负责人", "管理人员"]:
    m = G[name]
    group_rows.append([name, m['users'], m['events'], pct(m['successRate']), f"{m['avgResponseSeconds']:.1f}秒",
                       pct(m['completionRate']), f"{m['avgScore']:.2f}", pct(m['reuseIntentRate'])])
table(doc, ["用户群体", "人数", "事件数", "成功率", "响应时间", "完成率", "评分", "再用意愿"],
      group_rows, [1550, 800, 1000, 1150, 1200, 1150, 1000, 1510], font_size=8.2)
heading(doc, "4.2 主要发现", 2)
for x in [
    "普通职员的任务完成率为86.5%、平均评分为4.30分，整体表现最均衡，适合作为首批核心运营人群。",
    "职场新人成功率为90.9%，但完成率仅82.6%，说明‘生成成功’并不等于‘用户完成工作’，应增加结果引导和模板说明。",
    "项目负责人偏好文档问答、任务拆解和智能推荐，对信息准确性、负责人和截止时间字段更敏感。",
    "管理人员活跃天数最高，但成功率仅86.9%，再次使用意愿仅68.4%，明显低于其他群体，是优先改善对象。",
]: bullet(doc, x)
callout(doc, "人群结论", "先稳住普通职员的高频办公场景，再针对管理人员补充管理摘要、决策依据和跨任务状态跟踪；不建议用同一套模板覆盖所有用户群体。")

heading(doc, "五、核心功能复盘", 1)
heading(doc, "5.1 功能表现对比", 2)
function_order = ["会议纪要", "文档问答", "任务拆解", "周报生成", "自然语言咨询", "智能推荐", "任务自动化"]
rows = []
for name in function_order:
    m = F[name]
    rows.append([name, m['events'], pct(m['share']), pct(m['successRate']), f"{m['avgResponseSeconds']:.1f}秒",
                 pct(m['completionRate']), f"{m['avgScore']:.2f}"])
table(doc, ["功能", "使用次数", "使用占比", "成功率", "响应时间", "完成率", "评分"], rows,
      [1600, 1150, 1250, 1250, 1400, 1250, 1460], font_size=8.4)
heading(doc, "5.2 优势功能", 2)
body(doc, "会议纪要生成成功率为94.9%，任务完成率为89.4%，平均评分为4.33分，是当前最稳定的结构化生成能力。自然语言咨询成功率达到93.1%，响应时间为6.2秒，适合承担低门槛入口和即时答疑。文档问答使用占比最高（19.2%），再次使用意愿达到86.0%，说明资料型问答具有较强持续价值。")
heading(doc, "5.3 待优化功能", 2)
table(doc, ["功能", "主要问题", "数据表现", "优先动作"], [
    ["任务自动化", "Function Calling和状态更新失败", "成功率81.2%，响应11.4秒", "增加重试、降级、状态幂等和失败提示"],
    ["智能推荐", "推荐未命中、推荐理由不足", "成功率86.5%，完成率79.2%", "补充可解释理由与人工改选入口"],
    ["任务拆解", "负责人、截止时间等字段缺失", "成功率87.0%，响应9.7秒", "结构化校验，缺失字段标记待确认"],
    ["文档问答", "解析失败和来源定位不足", "55次失败，21次接口失败", "增强解析兼容与引用定位"],
    ["周报生成", "历史资料不足、导出格式异常", "51次失败", "增加资料不足提示并统一导出样式"],
], [1500, 2700, 2200, 2960], font_size=8.2)

heading(doc, "六、使用频率与场景复盘", 1)
heading(doc, "6.1 使用频率", 2)
freq_rows = []
for name in ["高频", "中频", "低频"]:
    m = Q[name]
    freq_rows.append([name, m['users'], f"{m['activeDaysAvg']:.1f}天", pct(m['successRate']),
                      pct(m['completionRate']), f"{m['avgScore']:.2f}", pct(m['reuseIntentRate'])])
table(doc, ["频率层级", "用户数", "平均活跃天数", "成功率", "完成率", "评分", "再用意愿"],
      freq_rows, [1300, 1000, 1700, 1200, 1200, 1100, 1860])
body(doc, "高频用户成功率和帮助率略高，说明使用熟练度可能降低操作成本；但高频用户完成率并未显著高于中频用户，后续需要验证其是否更多使用复杂功能。低频用户再次使用意愿为79.3%，主要障碍并非单一成功率，而可能与场景触发频率和价值感知有关。")
heading(doc, "6.2 典型人群—功能组合", 2)
table(doc, ["组合", "模拟表现", "解释"], [
    ["普通职员—会议纪要", "240次，成功率95.4%，完成率91.3%", "高频、稳定，可作为核心拉新体验"],
    ["普通职员—周报生成", "269次，成功率92.6%，完成率90.0%", "适合形成固定周度使用习惯"],
    ["项目负责人—文档问答", "180次，完成率90.2%", "资料型决策支持价值较高"],
    ["项目负责人—任务拆解", "165次，成功率87.3%", "需求高，但字段完整性仍影响交付"],
    ["管理人员—智能推荐", "68次，成功率80.9%", "复杂决策场景下推荐可靠性不足"],
    ["管理人员—任务自动化", "73次，成功率82.2%", "使用需求较强，稳定性尚未匹配"],
], [2600, 3100, 3660])

heading(doc, "七、反馈与满意度复盘", 1)
heading(doc, "7.1 反馈概览", 2)
body(doc, f"本轮共生成500条反馈记录，其中有效反馈476条。有效反馈平均评分为{O['avgScore']:.2f}分，认为结果有帮助的比例为{pct(O['helpfulRate'])}，愿意再次使用的比例为{pct(O['reuseIntentRate'])}。正向标签主要为‘节省时间’、‘操作方便’、‘结构清晰’和‘结果准确’。")
table(doc, ["反馈标签", "次数", "解释"], [
    ["节省时间", 122, "产品效率价值最明确"],
    ["操作方便", 116, "基础流程理解成本较低"],
    ["结构清晰", 103, "模板化输出获得认可"],
    ["结果准确", 96, "整体质量较好，但仍需来源支撑"],
    ["响应偏慢", 17, "集中关注复杂任务和接口波动"],
    ["内容不完整", 9, "与字段缺失、资料不足相关"],
    ["格式需优化", 9, "主要影响导出和二次编辑"],
], [2200, 1100, 6060])
heading(doc, "7.2 反馈解释边界", 2)
body(doc, "反馈样本主要从成功事件中抽取，因此评分和帮助率存在正向选择偏差，不能直接代表全部3000条事件。正式运营时应提高失败场景的反馈采集率，并分别展示成功用户、失败用户和中断用户的体验。")

heading(doc, "八、异常与风险复盘", 1)
heading(doc, "8.1 异常分布", 2)
table(doc, ["维度", "主要结果", "判断"], [
    ["异常功能", "文档问答20条、任务拆解18条、周报生成15条、任务自动化15条", "资料处理与复杂任务更易暴露问题"],
    ["异常原因", "接口调用失败35条、生成超时11条、字段缺失11条", "稳定性和结构化输出是首要风险"],
    ["严重程度", "中52条、低48条", "模拟集未设置高等级事件，正式上线前需补充"],
    ["处理状态", "已关闭38条、处理中31条、待处理31条", "仍有62%的异常未关闭"],
], [1600, 4800, 2960])
heading(doc, "8.2 风险判断", 2)
for x in [
    "接口调用失败在全部失败事件和异常样本中均排名第一，应设置熔断、重试、本地降级和错误埋点；",
    "任务自动化失败可能造成用户误以为任务已经执行，必须区分‘任务已生成’与‘外部动作已执行’；",
    "任务拆解字段缺失会直接影响负责人、截止时间和验收标准，应在输出前进行结构化完整性校验；",
    "模拟异常中没有高等级安全事件，不能据此判断真实系统不存在高风险，仍需单独进行安全测试。",
]: bullet(doc, x)

heading(doc, "九、问题优先级", 1)
table(doc, ["优先级", "问题", "影响", "判定依据"], [
    ["P0", "任务自动化调用与状态更新不稳定", "任务可能未执行或状态失真", "成功率最低81.2%，响应最慢11.4秒"],
    ["P0", "接口调用失败缺少稳定降级闭环", "影响全部核心功能", "105次失败，占失败事件35.5%"],
    ["P1", "任务拆解关键字段缺失", "影响任务可执行性", "字段缺失28次，其中任务拆解24次"],
    ["P1", "文档问答来源定位不足", "影响答案可信度", "解析失败18次、来源定位不足16次"],
    ["P1", "管理人员场景适配不足", "影响高价值用户留存", "成功率86.9%，再用意愿68.4%"],
    ["P2", "周报导出格式与历史资料提示不足", "影响结果交付体验", "导出格式异常15次、历史资料不足22次"],
    ["P2", "推荐理由和改选机制不足", "影响推荐采纳", "推荐未命中15次、理由不足11次"],
], [900, 3000, 2600, 2860], font_size=8.2)

heading(doc, "十、四大模块优化建议", 1)
heading(doc, "10.1 产品功能", 2)
for x in [
    "任务自动化增加执行前确认、任务生成/执行状态分离、失败重试和人工接管入口；",
    "任务拆解固定输出任务名称、负责人、优先级、截止时间、依赖、验收标准和风险，缺失项统一标记‘待确认’；",
    "周报生成增加历史资料完整性检查，资料不足时先提示补充，不直接生成确定性结论；",
    "文档问答增加引用片段、页码或段落定位，并显示解析失败文件及处理建议。",
]: bullet(doc, x)
heading(doc, "10.2 AI融合", 2)
for x in [
    "建立‘真实模型成功—本地降级成功—整体结果返回—最终失败’四层事件口径；",
    "对接口超时和短暂失败采用有限次数重试，对连续失败触发本地模板降级；",
    "智能推荐输出推荐理由、适用条件和替代选项，允许用户采纳、修改或跳过；",
    "针对管理人员和项目负责人建立差异化Prompt，强化决策摘要、风险、责任人和时间节点。",
]: bullet(doc, x)
heading(doc, "10.3 用户体验", 2)
for x in [
    "统一加载、成功、降级和失败状态提示，让用户知道系统当前处于哪一步；",
    "在生成结果页提供编辑、确认、保存、导出和再次生成的连续操作链路；",
    "为首次使用者提供示例输入与字段说明，降低生成成功但无法完成任务的情况；",
    "对高频复杂功能显示处理进度和预计等待时间，减少无反馈等待。",
]: bullet(doc, x)
heading(doc, "10.4 产品增长", 2)
for x in [
    "以会议纪要和周报生成为核心拉新场景，以文档问答提升持续使用；",
    "针对普通职员设计周度提醒和模板复用，促进稳定使用习惯；",
    "针对管理人员推出管理摘要、项目风险总览和待办跟踪组合场景；",
    "将推荐采纳率、连续使用率、核心功能复用率和失败后恢复率纳入后续增长看板。",
]: bullet(doc, x)

heading(doc, "十一、迭代计划与验收建议", 1)
table(doc, ["阶段", "建议周期", "重点任务", "验收建议"], [
    ["第一阶段", "1周", "修复接口失败、任务自动化状态和关键字段缺失", "任务自动化成功率≥90%，关键字段完整率≥95%"],
    ["第二阶段", "1周", "优化文档解析、来源定位和周报导出", "文档问答成功率≥93%，导出格式异常显著下降"],
    ["第三阶段", "1周", "优化推荐解释和管理人员场景", "管理人员再用意愿模拟值提升至75%以上"],
    ["第四阶段", "持续", "补充真实小流量数据并开展A/B验证", "模拟结论与真实行为趋势基本一致"],
], [1300, 1100, 3850, 3110])
body(doc, "上述目标是下一轮模拟验证建议，不是对真实运营结果的承诺。进入真实用户测试前，应重新确认数据合规、埋点准确性和样本代表性。")

heading(doc, "十二、局限性与后续工作", 1)
for x in [
    "数据由规则生成，用户行为之间的相关性可能比真实环境更规整；",
    "反馈主要来自成功事件，满意度指标可能偏高；",
    "模拟数据未覆盖真实网络波动、模型版本变化、权限差异和复杂安全事件；",
    "部分指标受功能适用范围影响，例如文档问答和咨询并不以导出作为主要目标；",
    "本轮结果用于确定复盘方法和优化优先级，后续仍需用真实小流量数据校准。",
]: bullet(doc, x)

heading(doc, "十三、结论", 1)
body(doc, "本轮模拟数据复盘证明，现有数据结构能够从用户、会话、事件、反馈和异常五个层级识别产品问题。总体表现具备继续迭代的基础，会议纪要、自然语言咨询和文档问答体现出较明确的用户价值；任务自动化、智能推荐、任务拆解及管理人员场景仍需重点优化。下一步应优先完成稳定性和结构化质量整改，再通过真实小流量验证模拟结论，形成‘数据发现—方案调整—原型修改—回归验证’的闭环。")
callout(doc, "最终结论", "本报告完成了第11周第2项交付物的目标：使用模拟数据运行指标体系，比较不同用户群体与不同功能场景的表现，并形成可执行的优化优先级。")

heading(doc, "附录A：主要数据来源", 1)
table(doc, ["文件", "用途"], [
    ["模拟用户数据集.xlsx", "提供模拟用户、会话、事件、反馈和异常明细"],
    ["模拟用户数据集构建说明.docx", "说明分群规则、字段定义、生成逻辑与适用边界"],
    ["第10周项目材料", "提供PRD、指标体系、功能测试、AI测试和风险口径"],
], [3200, 6160])
heading(doc, "附录B：复盘结论使用原则", 1)
for x in [
    "只把模拟结果用于验证分析方法和提出假设，不用于对外宣传；",
    "所有指标必须保留分母、适用范围和数据来源；",
    "在模型、Prompt、接口、数据结构或安全策略变化后，使用相同口径重新复盘；",
    "真实测试出现与模拟结论相反的结果时，以真实、合规、可追溯的数据为准。",
]: bullet(doc, x)

props = doc.core_properties
props.title = "模拟数据复盘分析报告"
props.subject = "百度办公 AI 助手第11周交付材料"
props.author = "AI产品经理实习项目"
props.keywords = "模拟数据, 数据复盘, 用户分群, 功能分析, AI产品"

doc.save(DOCX)
print(DOCX)
