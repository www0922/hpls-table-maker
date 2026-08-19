"""Build the approved HPLS leadership solution PDF."""

from __future__ import annotations

import argparse
import math
from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


PAGE_TITLES = ["封面", "成果摘要", "业务痛点", "方案设计", "人机分工", "量化成效", "复用方案"]
GUARDRAILS = ["输入契约", "失败即中止", "输出规则校验", "源表与原模板不改动"]
PAGE_WIDTH, PAGE_HEIGHT = landscape(A4)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
EVIDENCE_ASSET_DIR = PROJECT_ROOT / "assets" / "report" / "evidence"
EVIDENCE_ASSETS = {
    "precheck": EVIDENCE_ASSET_DIR / "pe100-input-check.png",
    "success": EVIDENCE_ASSET_DIR / "xp-generation-success.png",
    "review": EVIDENCE_ASSET_DIR / "manual-review-warning.png",
}

FONT_REGULAR = "HPLS-Regular"
FONT_BOLD = "HPLS-Bold"
FONT_REGULAR_PATH = Path(r"C:\Windows\Fonts\msyh.ttc")
FONT_BOLD_PATH = Path(r"C:\Windows\Fonts\msyhbd.ttc")

CHARCOAL = "#17202A"
INK = "#252B33"
TEAL = "#16A085"
DEEP_TEAL = "#0E7668"
ORANGE = "#E05D3F"
AMBER = "#F2A65A"
LIGHT_GRAY = "#F4F6F7"
MID_GRAY = "#D9DEE3"
TEXT_GRAY = "#5E6873"
LIGHT_GREEN = "#EDF7F4"
LIGHT_ORANGE = "#FFF4EB"
WHITE = "#FFFFFF"

SOLUTION_BODY_FONT_SIZE = 10.5
CONTENT_BODY_FONT_SIZE = 10.0
CONTENT_COMPACT_FONT_SIZE = 9.5
HUMAN_MACHINE_PRINCIPLE = "机器执行确定性规则；业务判断与质量责任保留给人。"
ORGANIZATION_VALUE = "专业人员转向异常判断、质量复核与规则优化。"
PLATFORM_ADAPTATION = "按平台适配源表字段、匹配规则、分组逻辑、计算公式、模板与业务阈值。"
SOLUTION_STAGES = [
    ("01", "源数据准备", "准备 3 或 5 类 Excel 源表", LIGHT_GRAY, CHARCOAL),
    ("02", "AI 统一入口", "识别表型，路由自动化管线", INK, WHITE),
    ("03", "前置检查", "缺失 / 重复 / 模板 / 依赖 / 路径", LIGHT_GRAY, CHARCOAL),
    ("04", "规则流水线", "迁移 / 匹配 / 分组 / 公式 / 稀释 / qPCR", LIGHT_GREEN, CHARCOAL),
    ("05", "校验与交付", "规则通过且输出存在后交付", LIGHT_GREEN, CHARCOAL),
]


def validate_evidence_assets(assets=None):
    """Ensure all report evidence assets are present before PDF generation."""
    assets = EVIDENCE_ASSETS if assets is None else assets
    normalized_assets = {key: Path(path) for key, path in assets.items()}
    missing = [path for path in normalized_assets.values() if not path.is_file()]
    if missing:
        raise FileNotFoundError(
            "Missing evidence assets: " + ", ".join(str(path) for path in missing)
        )
    return assets


def compute_metrics():
    """Derive leadership metrics from the confirmed measured ranges."""
    manual_mid = (60 + 120) / 2
    auto_mid = (1 + 2) / 2
    return {
        "manual_mid_minutes": manual_mid,
        "auto_mid_minutes": auto_mid,
        "minutes_saved": manual_mid - auto_mid,
        "time_reduction_percent": round((manual_mid - auto_mid) / manual_mid * 100, 2),
    }


def register_fonts():
    """Register embedded Chinese fonts once per interpreter."""
    if FONT_REGULAR in pdfmetrics.getRegisteredFontNames():
        return
    if not FONT_REGULAR_PATH.exists() or not FONT_BOLD_PATH.exists():
        raise FileNotFoundError("Microsoft YaHei fonts are required to generate the PDF")
    pdfmetrics.registerFont(TTFont(FONT_REGULAR, str(FONT_REGULAR_PATH), subfontIndex=0))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, str(FONT_BOLD_PATH), subfontIndex=0))


def _color(value):
    return value if hasattr(value, "red") else HexColor(value)


def draw_rounded_box(c, x, y, w, h, fill, stroke=None, radius=4):
    c.saveState()
    c.setFillColor(_color(fill))
    if stroke:
        c.setStrokeColor(_color(stroke))
        c.setLineWidth(0.8)
    else:
        c.setStrokeColor(_color(fill))
        c.setLineWidth(0)
    c.roundRect(x, y, w, h, radius, fill=1, stroke=1 if stroke else 0)
    c.restoreState()


def draw_contained_image(c, image_path, x, y, width, height, background=INK):
    image = ImageReader(str(image_path))
    image_width, image_height = image.getSize()
    scale = min(width / image_width, height / image_height)
    draw_width = image_width * scale
    draw_height = image_height * scale
    draw_rounded_box(c, x, y, width, height, background, radius=3)
    c.drawImage(
        image,
        x + (width - draw_width) / 2,
        y + (height - draw_height) / 2,
        draw_width,
        draw_height,
        preserveAspectRatio=True,
        anchor="c",
        mask="auto",
    )


def _wrap_text(text, font, size, width):
    lines = []
    for paragraph in str(text).split("\n"):
        if paragraph == "":
            lines.append("")
            continue
        current = ""
        for char in paragraph:
            candidate = current + char
            if current and pdfmetrics.stringWidth(candidate, font, size) > width:
                lines.append(current.rstrip())
                current = char.lstrip()
            else:
                current = candidate
        if current or not lines:
            lines.append(current.rstrip())
    return lines


def draw_text(c, text, x, y, width, font=FONT_REGULAR, size=10, color=TEXT_GRAY, leading=None, max_lines=None):
    """Draw wrapped text from a top coordinate and return the next y."""
    leading = leading or size * 1.42
    lines = _wrap_text(text, font, size, width)
    if max_lines is not None and len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]
        while last and pdfmetrics.stringWidth(last + "...", font, size) > width:
            last = last[:-1]
        lines[-1] = last.rstrip() + "..."
    c.saveState()
    c.setFillColor(_color(color))
    c.setFont(font, size)
    cursor = y
    for line in lines:
        c.drawString(x, cursor - size, line)
        cursor -= leading
    c.restoreState()
    return cursor


def draw_centered_text(c, text, center_x, y, font, size, color):
    c.saveState()
    c.setFillColor(_color(color))
    c.setFont(font, size)
    c.drawCentredString(center_x, y, text)
    c.restoreState()


def draw_arrow(c, x1, y1, x2, y2, color=ORANGE):
    angle = math.atan2(y2 - y1, x2 - x1)
    head = 7
    c.saveState()
    c.setStrokeColor(_color(color))
    c.setFillColor(_color(color))
    c.setLineWidth(2)
    c.line(x1, y1, x2, y2)
    p1 = (x2, y2)
    p2 = (x2 - head * math.cos(angle - 0.52), y2 - head * math.sin(angle - 0.52))
    p3 = (x2 - head * math.cos(angle + 0.52), y2 - head * math.sin(angle + 0.52))
    path = c.beginPath()
    path.moveTo(*p1)
    path.lineTo(*p2)
    path.lineTo(*p3)
    path.close()
    c.drawPath(path, fill=1, stroke=0)
    c.restoreState()


def draw_header(c, section, title, page_number):
    c.saveState()
    c.setFillColor(_color(CHARCOAL))
    c.setFont(FONT_BOLD, 8.5)
    c.drawString(38, PAGE_HEIGHT - 34, section.upper())
    c.setFillColor(_color(ORANGE))
    c.rect(38, PAGE_HEIGHT - 43, 42, 3, fill=1, stroke=0)
    c.setFillColor(_color(CHARCOAL))
    c.setFont(FONT_BOLD, 24)
    c.drawString(38, PAGE_HEIGHT - 70, title)
    c.setFillColor(_color(TEXT_GRAY))
    c.setFont(FONT_BOLD, 11)
    c.drawRightString(PAGE_WIDTH - 38, PAGE_HEIGHT - 54, f"{page_number:02d}")
    c.setStrokeColor(_color(MID_GRAY))
    c.setLineWidth(0.7)
    c.line(38, PAGE_HEIGHT - 83, PAGE_WIDTH - 38, PAGE_HEIGHT - 83)
    c.restoreState()


def draw_footer(c, page_number):
    y = 19
    c.saveState()
    c.setStrokeColor(_color(MID_GRAY))
    c.setLineWidth(0.6)
    c.line(38, 33, PAGE_WIDTH - 38, 33)
    c.setFillColor(_color(TEXT_GRAY))
    c.setFont(FONT_BOLD, 7.5)
    c.drawString(38, y, "HPLS AUTOMATION")
    c.setFont(FONT_REGULAR, 7.5)
    c.drawCentredString(PAGE_WIDTH / 2, y, "成果汇报 · 2026.08")
    c.drawRightString(PAGE_WIDTH - 38, y, f"{page_number} / 07")
    c.restoreState()


def draw_metric_card(c, x, y, w, h, value, label, accent=TEAL, note=None):
    draw_rounded_box(c, x, y, w, h, LIGHT_GRAY)
    c.saveState()
    c.setFillColor(_color(accent))
    c.rect(x, y + h - 4, w, 4, fill=1, stroke=0)
    c.setFillColor(_color(CHARCOAL))
    c.setFont(FONT_BOLD, 25)
    c.drawString(x + 14, y + h - 39, value)
    c.restoreState()
    draw_text(c, label, x + 14, y + h - 48, w - 28, size=9.5, color=TEXT_GRAY, max_lines=2)
    if note:
        draw_text(c, note, x + 14, y + 19, w - 28, size=9.5, color=TEXT_GRAY, max_lines=1)


def draw_section_label(c, text, x, y, color=TEAL):
    c.saveState()
    c.setFillColor(_color(color))
    c.setFont(FONT_BOLD, 8)
    c.drawString(x, y, text.upper())
    c.restoreState()


def draw_cover(c):
    c.saveState()
    c.setFillColor(_color(CHARCOAL))
    c.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, fill=1, stroke=0)
    c.setFillColor(_color(TEAL))
    c.rect(0, 0, 12, PAGE_HEIGHT, fill=1, stroke=0)
    c.setFillColor(_color(AMBER))
    c.rect(PAGE_WIDTH - 165, 0, 165, 10, fill=1, stroke=0)
    c.restoreState()

    draw_section_label(c, "HPLS AUTOMATION / LEADERSHIP BRIEF", 48, PAGE_HEIGHT - 56, AMBER)
    draw_text(c, "HPLS 文库表\n自动化生成方案", 48, PAGE_HEIGHT - 98, 510, FONT_BOLD, 34, WHITE, leading=45)
    draw_text(c, "把实验室制表经验固化为可校验、可复用的自动化能力", 50, PAGE_HEIGHT - 205, 520, size=13, color="#D7DDE2")

    c.saveState()
    c.setFillColor(_color(WHITE))
    c.setFont(FONT_BOLD, 22)
    c.drawString(50, 305, "从手工做表")
    c.setFillColor(_color(ORANGE))
    c.setFont(FONT_BOLD, 39)
    c.drawString(184, 294, "2 小时")
    c.setFillColor(_color("#8D98A3"))
    c.setFont(FONT_BOLD, 22)
    c.drawString(308, 305, "到")
    c.setFillColor(_color("#45C4AE"))
    c.setFont(FONT_BOLD, 39)
    c.drawString(345, 294, "1 分钟")
    c.restoreState()

    stages = [("放入源表", "人"), ("一句话触发", "AI"), ("规则流水线", "程序"), ("校验后交付", "人机")]
    start_x = 50
    y = 196
    box_w = 153
    for idx, (title, owner) in enumerate(stages):
        x = start_x + idx * 187
        draw_rounded_box(c, x, y, box_w, 64, "#252F38", "#46515C")
        draw_section_label(c, owner, x + 12, y + 43, AMBER if owner in {"人", "人机"} else "#45C4AE")
        draw_text(c, title, x + 12, y + 34, box_w - 24, FONT_BOLD, 12, WHITE, max_lines=1)
        if idx < len(stages) - 1:
            draw_arrow(c, x + box_w + 8, y + 32, x + box_w + 27, y + 32, ORANGE)

    c.saveState()
    c.setFillColor(_color("#9DA7B1"))
    c.setFont(FONT_REGULAR, CONTENT_COMPACT_FONT_SIZE)
    c.drawString(50, 92, "覆盖表型")
    c.restoreState()
    for idx, name in enumerate(["PE100", "PE150", "XP", "真迈"]):
        x = 50 + idx * 102
        draw_rounded_box(c, x, 49, 88, 28, "#303B45", "#4A5661", radius=3)
        draw_centered_text(c, name, x + 44, 58, FONT_BOLD, 9, WHITE)
    c.saveState()
    c.setFillColor(_color("#9DA7B1"))
    c.setFont(FONT_REGULAR, 8)
    c.drawRightString(PAGE_WIDTH - 38, 56, "成果汇报 · 2026.08")
    c.restoreState()


def draw_executive_summary(c):
    draw_header(c, "EXECUTIVE SUMMARY", "成果摘要", 2)
    metrics = compute_metrics()
    margin = 38
    gap = 12
    card_w = (PAGE_WIDTH - margin * 2 - gap * 3) / 4
    y = 357
    cards = [
        (f"约 {round(metrics['time_reduction_percent']):.0f}%", "处理时长下降", ORANGE, "按区间中位数计算"),
        (f"{metrics['minutes_saved']:.1f} 分钟", "单批释放人工时间", TEAL, "从重复操作转向复核"),
        ("4 类", "测序表型覆盖", AMBER, "PE100 / PE150 / XP / 真迈"),
        ("1 个入口", "自然语言统一调度", CHARCOAL, "非技术同事可直接使用"),
    ]
    for idx, args in enumerate(cards):
        draw_metric_card(c, margin + idx * (card_w + gap), y, card_w, 112, *args)

    draw_section_label(c, "VALUE SHIFT", margin, 326)
    draw_text(c, "价值不只在提速，更在于把经验变成组织能力", margin, 310, 520, FONT_BOLD, 18, CHARCOAL, max_lines=1)

    left_x = margin
    left_w = 238
    center_x = left_x + left_w + 22
    right_x = center_x + 44
    right_w = PAGE_WIDTH - margin - right_x
    draw_rounded_box(c, left_x, 137, left_w, 137, INK)
    draw_section_label(c, "BEFORE", left_x + 16, 250, AMBER)
    draw_text(c, "复制粘贴\n公式维护\n规则依赖个人经验", left_x + 16, 232, left_w - 32, FONT_BOLD, 13, WHITE, leading=25)
    draw_arrow(c, center_x, 204, center_x + 33, 204, ORANGE)
    draw_rounded_box(c, right_x, 137, right_w, 137, LIGHT_GREEN)
    draw_section_label(c, "AFTER", right_x + 16, 250, DEEP_TEAL)
    draw_text(c, "异常判断 / 质量复核 / 规则优化", right_x + 16, 231, right_w - 32, FONT_BOLD, 15, CHARCOAL, max_lines=1)
    draw_text(c, "输入预检、规则执行和输出校验由机器稳定完成；专业人员把精力放在真正需要判断的环节。", right_x + 16, 199, right_w - 32, size=10, color=TEXT_GRAY, max_lines=3)

    draw_footer(c, 2)


def draw_pain_points(c):
    draw_header(c, "BUSINESS PAIN POINTS", "业务痛点", 3)
    margin = 38

    draw_rounded_box(c, margin, 328, 244, 150, INK)
    draw_section_label(c, "COMPLEX INPUT", margin + 16, 452, AMBER)
    draw_text(c, "一批任务，多类源表", margin + 16, 432, 205, FONT_BOLD, 18, WHITE, max_lines=1)
    c.saveState()
    c.setFillColor(_color(ORANGE))
    c.setFont(FONT_BOLD, 35)
    c.drawString(margin + 16, 370, "3")
    c.setFillColor(_color("#C5CDD4"))
    c.setFont(FONT_REGULAR, CONTENT_COMPACT_FONT_SIZE)
    c.drawString(margin + 44, 379, "类源表")
    c.setFillColor(_color("#6F7A84"))
    c.rect(margin + 91, 367, 1, 48, fill=1, stroke=0)
    c.setFillColor(_color("#45C4AE"))
    c.setFont(FONT_BOLD, 35)
    c.drawString(margin + 111, 370, "5")
    c.setFillColor(_color("#C5CDD4"))
    c.setFont(FONT_REGULAR, 9)
    c.drawString(margin + 139, 379, "类源表")
    c.restoreState()
    draw_text(c, "PE100 / PE150", margin + 16, 353, 84, size=CONTENT_COMPACT_FONT_SIZE, color="#AEB7BF", max_lines=1)
    draw_text(c, "XP / 真迈", margin + 111, 353, 90, size=CONTENT_COMPACT_FONT_SIZE, color="#AEB7BF", max_lines=1)

    draw_rounded_box(c, margin + 258, 328, PAGE_WIDTH - margin * 2 - 258, 150, LIGHT_ORANGE)
    draw_section_label(c, "TIME COST", margin + 274, 452, ORANGE)
    draw_text(c, "单批手工处理", margin + 274, 432, 220, FONT_BOLD, 18, CHARCOAL, max_lines=1)
    c.saveState()
    c.setFillColor(_color(ORANGE))
    c.setFont(FONT_BOLD, 45)
    c.drawString(margin + 274, 365, "1-2 小时")
    c.restoreState()
    draw_text(c, "时间被消耗在数据搬运、查找匹配与公式维护", margin + 274, 350, 410, size=10, color=TEXT_GRAY, max_lines=2)

    draw_section_label(c, "RULE CHAIN", margin, 293)
    draw_text(c, "十几道操作相互依赖，任一处错漏都可能传导到最终结果", margin, 277, 620, FONT_BOLD, 16, CHARCOAL, max_lines=1)
    chain = ["数据迁移", "浓度匹配", "分组排序", "公式计算", "命名标记", "稀释 / qPCR"]
    chain_y = 201
    chain_w = 112
    chain_gap = 15
    for idx, name in enumerate(chain):
        x = margin + idx * (chain_w + chain_gap)
        fill = LIGHT_GRAY if idx % 2 == 0 else LIGHT_GREEN
        draw_rounded_box(c, x, chain_y, chain_w, 48, fill)
        draw_centered_text(c, name, x + chain_w / 2, chain_y + 19, FONT_BOLD, 10, CHARCOAL)
        if idx < len(chain) - 1:
            draw_arrow(c, x + chain_w + 4, chain_y + 24, x + chain_w + 12, chain_y + 24, ORANGE)

    risks = [
        ("差错风险", "复制、公式和匹配规则容易错漏"),
        ("经验风险", "规则集中在熟练人员个人经验中"),
        ("组织风险", "新人上手慢，跨平台维护成本高"),
    ]
    risk_y = 80
    risk_w = (PAGE_WIDTH - margin * 2 - 20) / 3
    for idx, (title, body) in enumerate(risks):
        x = margin + idx * (risk_w + 10)
        c.saveState()
        c.setFillColor(_color(ORANGE))
        c.rect(x, risk_y, 4, 67, fill=1, stroke=0)
        c.restoreState()
        draw_text(c, title, x + 13, risk_y + 55, risk_w - 22, FONT_BOLD, 11, CHARCOAL, max_lines=1)
        draw_text(c, body, x + 13, risk_y + 35, risk_w - 22, size=CONTENT_BODY_FONT_SIZE, color=TEXT_GRAY, max_lines=2)
    draw_footer(c, 3)


def draw_solution_design(c):
    draw_header(c, "SOLUTION DESIGN", "方案设计", 4)
    margin = 38
    draw_text(c, "一句话触发，五层闭环完成交付", margin, 490, 520, FONT_BOLD, 17, CHARCOAL, max_lines=1)
    draw_rounded_box(c, PAGE_WIDTH - 214, 465, 176, 30, LIGHT_ORANGE)
    draw_centered_text(c, "本地处理 · 模板不被修改", PAGE_WIDTH - 126, 475, FONT_BOLD, CONTENT_COMPACT_FONT_SIZE, CHARCOAL)

    stages = SOLUTION_STAGES
    stage_gap = 24
    stage_w = (PAGE_WIDTH - margin * 2 - stage_gap * 4) / 5
    stage_y = 272
    stage_h = 158
    for idx, (num, title, body, fill, text_color) in enumerate(stages):
        x = margin + idx * (stage_w + stage_gap)
        draw_rounded_box(c, x, stage_y, stage_w, stage_h, fill, MID_GRAY if fill != INK else "#46515C")
        c.saveState()
        c.setFillColor(_color(AMBER if fill == INK else TEAL))
        c.circle(x + 22, stage_y + stage_h - 24, 12, fill=1, stroke=0)
        c.setFillColor(_color(INK if fill == INK else WHITE))
        c.setFont(FONT_BOLD, 8.5)
        c.drawCentredString(x + 22, stage_y + stage_h - 27, num)
        c.restoreState()
        draw_text(c, title, x + 12, stage_y + stage_h - 54, stage_w - 24, FONT_BOLD, 12, text_color, max_lines=2)
        draw_text(
            c,
            body,
            x + 12,
            stage_y + 75,
            stage_w - 24,
            size=SOLUTION_BODY_FONT_SIZE,
            color="#D5DCE2" if fill == INK else TEXT_GRAY,
            max_lines=4,
        )
        if idx < 4:
            draw_arrow(c, x + stage_w + 5, stage_y + stage_h / 2, x + stage_w + 18, stage_y + stage_h / 2, ORANGE)

    lanes = [("PE100", "6 个核心步骤"), ("PE150", "7 个核心步骤"), ("XP", "9 步 + qPCR"), ("真迈", "9 步 + qPCR")]
    lane_y = 192
    lane_gap = 9
    lane_w = (PAGE_WIDTH - margin * 2 - lane_gap * 3) / 4
    for idx, (name, detail) in enumerate(lanes):
        x = margin + idx * (lane_w + lane_gap)
        draw_rounded_box(c, x, lane_y, lane_w, 49, CHARCOAL)
        draw_text(c, name, x + 12, lane_y + 37, 65, FONT_BOLD, 10, WHITE, max_lines=1)
        c.saveState()
        c.setStrokeColor(_color(TEAL))
        c.line(x + 75, lane_y + 12, x + 75, lane_y + 37)
        c.restoreState()
        draw_text(c, detail, x + 88, lane_y + 35, lane_w - 98, size=CONTENT_COMPACT_FONT_SIZE, color="#D7DDE2", max_lines=2)

    draw_rounded_box(c, margin, 82, PAGE_WIDTH - margin * 2, 77, LIGHT_ORANGE)
    draw_section_label(c, "PROTECTION MECHANISMS", margin + 16, 138, ORANGE)
    guards = GUARDRAILS
    guard_w = (PAGE_WIDTH - margin * 2 - 32 - 30) / 4
    for idx, guard in enumerate(guards):
        x = margin + 16 + idx * (guard_w + 10)
        draw_rounded_box(c, x, 96, guard_w, 29, WHITE, "#F0CAB7", radius=3)
        draw_centered_text(c, guard, x + guard_w / 2, 106, FONT_BOLD, CONTENT_BODY_FONT_SIZE, CHARCOAL)
    draw_footer(c, 4)


def draw_human_machine(c):
    draw_header(c, "HUMAN + MACHINE", "人机分工", 5)
    margin = 38
    draw_text(c, "人定边界，机器执行，异常回到人", margin, 490, 470, FONT_BOLD, 17, CHARCOAL, max_lines=1)
    draw_rounded_box(c, PAGE_WIDTH - 342, 458, 304, 43, INK)
    draw_text(c, HUMAN_MACHINE_PRINCIPLE, PAGE_WIDTH - 328, 489, 278, size=CONTENT_BODY_FONT_SIZE, color=WHITE, max_lines=2)

    table_x = margin
    table_y = 173
    table_w = PAGE_WIDTH - margin * 2
    role_w = 116
    col_w = (table_w - role_w) / 3
    header_h = 37
    row_h = 116

    headers = ["角色", "执行前", "执行中", "执行后"]
    widths = [role_w, col_w, col_w, col_w]
    x = table_x
    for idx, (label, w) in enumerate(zip(headers, widths)):
        draw_rounded_box(c, x, table_y + row_h * 2, w - 4, header_h, CHARCOAL, radius=2)
        draw_centered_text(c, label, x + (w - 4) / 2, table_y + row_h * 2 + 13, FONT_BOLD, CONTENT_BODY_FONT_SIZE, WHITE)
        x += w

    rows = [
        ("业务人员", ORANGE, [
            ("确认任务", "选择表型与批次，准备并核对源表，保证来源可信。"),
            ("处理例外", "对缺失、重复或来源警告作出判断，不让系统自行猜测。"),
            ("质量放行", "复核关键结果、确认交付对象，并承担最终业务责任。"),
        ]),
        ("AI + 程序", TEAL, [
            ("识别与预检", "路由四类表型，检查文件、模板、依赖、路径及覆盖风险。"),
            ("稳定执行", "完成迁移、匹配、分组、计算、命名、稀释及 qPCR。"),
            ("校验与交付", "执行关键规则校验；仅成功且输出存在时返回结果。"),
        ]),
    ]
    for row_idx, (role, accent, cells) in enumerate(rows):
        y = table_y + row_h * (1 - row_idx)
        draw_rounded_box(c, table_x, y, role_w - 4, row_h - 6, WHITE, MID_GRAY, radius=2)
        c.saveState()
        c.setFillColor(_color(accent))
        c.rect(table_x, y, 5, row_h - 6, fill=1, stroke=0)
        c.restoreState()
        draw_text(c, role, table_x + 16, y + 70, role_w - 28, FONT_BOLD, 12, CHARCOAL, max_lines=2)
        for col_idx, (title, body) in enumerate(cells):
            x = table_x + role_w + col_idx * col_w
            fill = LIGHT_ORANGE if row_idx == 0 else LIGHT_GREEN
            draw_rounded_box(c, x, y, col_w - 4, row_h - 6, fill, radius=2)
            draw_text(c, title, x + 12, y + row_h - 25, col_w - 28, FONT_BOLD, 10, accent, max_lines=1)
            draw_text(c, body, x + 12, y + row_h - 49, col_w - 28, size=CONTENT_BODY_FONT_SIZE, color=TEXT_GRAY, max_lines=4)

    draw_rounded_box(c, margin, 78, 472, 68, LIGHT_ORANGE)
    draw_section_label(c, "EXCEPTION LOOP", margin + 14, 126, ORANGE)
    draw_text(c, "发现输入歧义、规则冲突或来源数据不足时，系统停止并指出问题，由人修正后重跑。", margin + 14, 114, 442, size=CONTENT_BODY_FONT_SIZE, color=CHARCOAL, max_lines=3)
    draw_rounded_box(c, margin + 484, 78, PAGE_WIDTH - margin * 2 - 484, 68, LIGHT_GREEN)
    draw_section_label(c, "ORGANIZATION VALUE", margin + 498, 126, DEEP_TEAL)
    draw_text(c, ORGANIZATION_VALUE, margin + 498, 114, PAGE_WIDTH - margin * 2 - 512, size=CONTENT_BODY_FONT_SIZE, color=CHARCOAL, max_lines=3)
    draw_footer(c, 5)


def draw_impact(c):
    draw_header(c, "MEASURED IMPACT", "量化成效", 6)
    margin = 38
    metrics = compute_metrics()
    draw_text(c, "单批处理从小时级进入分钟级", margin, 490, 500, FONT_BOLD, 17, CHARCOAL, max_lines=1)
    draw_rounded_box(c, PAGE_WIDTH - 178, 465, 140, 29, LIGHT_GREEN, "#A8D5CD", radius=3)
    draw_centered_text(c, "数据状态：实测确认", PAGE_WIDTH - 108, 475, FONT_BOLD, CONTENT_COMPACT_FONT_SIZE, DEEP_TEAL)

    chart_x = margin
    chart_y = 182
    chart_w = 493
    chart_h = 247
    draw_rounded_box(c, chart_x, chart_y, chart_w, chart_h, LIGHT_GRAY)
    draw_text(c, "单批处理时长对比", chart_x + 17, chart_y + chart_h - 17, 250, FONT_BOLD, 13, CHARCOAL, max_lines=1)

    labels = [("手工制表", "60-120 分钟", 1.0, ORANGE, "中位 90"), ("自动生成", "1-2 分钟", 0.0167, TEAL, "中位 1.5")]
    for idx, (label, range_text, ratio, color, midpoint) in enumerate(labels):
        y = chart_y + 145 - idx * 70
        row_text_y = y + 8
        track_x = chart_x + 100
        track_w = 290
        c.saveState()
        c.setFillColor(_color(CHARCOAL))
        c.setFont(FONT_BOLD, CONTENT_COMPACT_FONT_SIZE)
        c.drawString(chart_x + 17, row_text_y, label)
        c.setFillColor(_color("#E2E6E9"))
        c.rect(track_x, y, track_w, 26, fill=1, stroke=0)
        bar_w = max(8, track_w * ratio)
        c.setFillColor(_color(color))
        c.rect(track_x, y, bar_w, 26, fill=1, stroke=0)
        if ratio > 0.3:
            c.setFillColor(white)
            c.setFont(FONT_BOLD, CONTENT_COMPACT_FONT_SIZE)
            c.drawString(track_x + 8, row_text_y, range_text)
        else:
            c.setFillColor(_color(CHARCOAL))
            c.setFont(FONT_BOLD, CONTENT_COMPACT_FONT_SIZE)
            c.drawString(track_x + bar_w + 8, row_text_y, range_text)
        c.setFillColor(_color(CHARCOAL))
        c.setFont(FONT_BOLD, CONTENT_BODY_FONT_SIZE)
        c.drawRightString(chart_x + chart_w - 17, row_text_y, midpoint)
        c.restoreState()

    delta_y = chart_y + 15
    delta_w = (chart_w - 44) / 2
    draw_metric_card(c, chart_x + 17, delta_y, delta_w, 67, f"约 {round(metrics['time_reduction_percent']):.0f}%", "处理时长下降", TEAL)
    draw_metric_card(c, chart_x + 27 + delta_w, delta_y, delta_w, 67, f"{metrics['minutes_saved']:.1f} 分钟", "单批释放人工时间", TEAL)

    side_x = chart_x + chart_w + 14
    side_w = PAGE_WIDTH - margin - side_x
    draw_section_label(c, "SYSTEM GAINS", side_x, chart_y + chart_h - 5)
    draw_rounded_box(c, side_x, 300, side_w, 106, INK)
    draw_text(c, "60%+", side_x + 16, 384, side_w - 32, FONT_BOLD, 29, WHITE, max_lines=1)
    draw_text(c, "减少重复 Excel I/O 后的\n流水线性能提升", side_x + 16, 344, side_w - 32, size=CONTENT_BODY_FONT_SIZE, color="#D7DDE2", max_lines=3)
    draw_rounded_box(c, side_x, 182, side_w, 106, LIGHT_GREEN)
    draw_text(c, "4 类表型", side_x + 16, 266, side_w - 32, FONT_BOLD, 25, CHARCOAL, max_lines=1)
    draw_text(c, "由一个 AI 入口统一调度", side_x + 16, 225, side_w - 32, size=CONTENT_BODY_FONT_SIZE, color=TEXT_GRAY, max_lines=2)

    draw_rounded_box(c, margin, 77, PAGE_WIDTH - margin * 2, 76, LIGHT_ORANGE)
    draw_section_label(c, "MEASUREMENT NOTES", margin + 14, 132, ORANGE)
    note = "实测：手工 1-2 小时、自动化 1-2 分钟、性能提升 60%+。约 98% 与 88.5 分钟按两组区间中位数计算；不外推年度收益，不宣称绝对零错误。"
    draw_text(c, note, margin + 14, 120, PAGE_WIDTH - margin * 2 - 28, size=CONTENT_BODY_FONT_SIZE, color=CHARCOAL, max_lines=3)
    draw_footer(c, 6)


def draw_reuse(c):
    draw_header(c, "REUSE FRAMEWORK", "复用方案", 7)
    margin = 38
    draw_text(c, "五类标准资产，复制到更多制表场景", margin, 490, 520, FONT_BOLD, 17, CHARCOAL, max_lines=1)
    draw_rounded_box(c, PAGE_WIDTH - 257, 465, 219, 29, INK)
    draw_centered_text(c, "新场景 = 通用骨架 + 平台规则适配", PAGE_WIDTH - 147.5, 475, FONT_BOLD, CONTENT_COMPACT_FONT_SIZE, WHITE)

    assets = [
        ("01", "输入契约", "定义关键词、数量、表结构与唯一性。"),
        ("02", "输出模板", "约定文件名、Sheet 与交付目录。"),
        ("03", "规则模块", "迁移、匹配、分组和公式独立化。"),
        ("04", "校验清单", "关键规则自动检查，失败即停止。"),
        ("05", "AI 入口", "识别场景、执行预检、调度与交付。"),
    ]
    asset_gap = 10
    asset_w = (PAGE_WIDTH - margin * 2 - asset_gap * 4) / 5
    asset_y = 304
    for idx, (num, title, body) in enumerate(assets):
        x = margin + idx * (asset_w + asset_gap)
        draw_rounded_box(c, x, asset_y, asset_w, 127, LIGHT_GRAY)
        c.saveState()
        c.setFillColor(_color(TEAL))
        c.rect(x, asset_y + 123, asset_w, 4, fill=1, stroke=0)
        c.restoreState()
        draw_section_label(c, num, x + 12, asset_y + 104, ORANGE)
        draw_text(c, title, x + 12, asset_y + 90, asset_w - 24, FONT_BOLD, 12, CHARCOAL, max_lines=1)
        draw_text(c, body, x + 12, asset_y + 61, asset_w - 24, size=CONTENT_BODY_FONT_SIZE, color=TEXT_GRAY, max_lines=4)

    layer_gap = 12
    layer_w = (PAGE_WIDTH - margin * 2 - layer_gap) / 2
    draw_rounded_box(c, margin, 179, layer_w, 91, LIGHT_GREEN)
    draw_section_label(c, "DIRECT REUSE", margin + 15, 249, DEEP_TEAL)
    draw_text(c, "直接复用：通用骨架", margin + 15, 234, layer_w - 30, FONT_BOLD, 11, CHARCOAL, max_lines=1)
    draw_text(c, "文件发现、覆盖保护、步骤编排、错误翻译、日志输出、结果存在性检查与交付机制。", margin + 15, 211, layer_w - 30, size=CONTENT_BODY_FONT_SIZE, color=TEXT_GRAY, max_lines=3)
    adapt_x = margin + layer_w + layer_gap
    draw_rounded_box(c, adapt_x, 179, layer_w, 91, LIGHT_ORANGE)
    draw_section_label(c, "ADAPT BY PLATFORM", adapt_x + 15, 249, ORANGE)
    draw_text(c, "按需适配：平台差异", adapt_x + 15, 234, layer_w - 30, FONT_BOLD, 11, CHARCOAL, max_lines=1)
    draw_text(c, PLATFORM_ADAPTATION, adapt_x + 15, 211, layer_w - 30, size=CONTENT_BODY_FONT_SIZE, color=TEXT_GRAY, max_lines=3)

    draw_rounded_box(c, margin, 79, PAGE_WIDTH - margin * 2, 67, INK)
    draw_section_label(c, "REPLICATION PATH", margin + 15, 126, AMBER)
    path_steps = ["盘点人工步骤", "固化输入输出", "模块化规则", "建立校验", "接入 AI 入口"]
    start_x = margin + 15
    available = PAGE_WIDTH - margin * 2 - 30
    step_w = 118
    path_gap = (available - step_w * 5) / 4
    for idx, label in enumerate(path_steps):
        x = start_x + idx * (step_w + path_gap)
        draw_rounded_box(c, x, 91, step_w, 25, "#303B45", "#4A5661", radius=3)
        draw_centered_text(c, label, x + step_w / 2, 99, FONT_BOLD, CONTENT_COMPACT_FONT_SIZE, WHITE)
        if idx < 4:
            draw_arrow(c, x + step_w + 4, 103, x + step_w + path_gap - 4, 103, TEAL)
    draw_footer(c, 7)


def draw_evidence_precheck(c):
    draw_header(c, "MEASURED EVIDENCE", "输入预检与失败保护", 8)
    margin = 38
    draw_text(
        c,
        "缺文件即停止，并把问题翻译成可执行的补救步骤",
        margin,
        490,
        590,
        FONT_BOLD,
        17,
        CHARCOAL,
        max_lines=1,
    )
    image_x, image_y, image_w, image_h = margin, 86, 565, 359
    draw_contained_image(c, EVIDENCE_ASSETS["precheck"], image_x, image_y, image_w, image_h)
    card_x = image_x + image_w + 14
    card_w = PAGE_WIDTH - margin - card_x
    cards = [
        (338, "缺文件即停止", "系统不猜测、不带错运行。", LIGHT_ORANGE, ORANGE),
        (220, "问题可定位", "明确指出缺少的 3 类文件。", LIGHT_GREEN, DEEP_TEAL),
        (102, "修复可执行", "直接给出目录与下一步操作。", LIGHT_GRAY, TEAL),
    ]
    for y, title, body, fill, accent in cards:
        draw_rounded_box(c, card_x, y, card_w, 100, fill)
        c.saveState()
        c.setFillColor(_color(accent))
        c.rect(card_x, y, 4, 100, fill=1, stroke=0)
        c.restoreState()
        draw_text(c, title, card_x + 14, y + 74, card_w - 28, FONT_BOLD, 12, CHARCOAL, max_lines=1)
        draw_text(c, body, card_x + 14, y + 48, card_w - 28, size=10, color=TEXT_GRAY, max_lines=3)
    draw_footer(c, 8)


def draw_evidence_success(c):
    draw_header(c, 'MEASURED EVIDENCE', '成功交付与人工复核闭环', 9)
    margin = 38
    draw_text(c, '机器完成确定性任务，异常判断与质量责任回到人', margin, 490, 600, FONT_BOLD, 17, CHARCOAL, max_lines=1)
    draw_contained_image(c, EVIDENCE_ASSETS['success'], margin, 86, 523, 359)
    side_x = margin + 537
    side_w = PAGE_WIDTH - margin - side_x
    draw_contained_image(c, EVIDENCE_ASSETS['review'], side_x, 327, side_w, 118)
    draw_rounded_box(c, side_x, 86, side_w, 225, INK)
    draw_section_label(c, 'HUMAN REVIEW', side_x+16, 286, AMBER)
    draw_text(c, '17 项', side_x+16, 264, side_w-32, FONT_BOLD, 28, '#45C4AE', max_lines=1)
    draw_text(c, '异常记录回到人工复核', side_x+16, 218, side_w-32, FONT_BOLD, 11, WHITE, max_lines=2)
    metric_w = (side_w - 42) / 2
    first_x = side_x + 16
    second_x = first_x + metric_w + 10
    draw_rounded_box(c, first_x, 118, metric_w, 54, '#303B45', '#4A5661', radius=3)
    draw_rounded_box(c, second_x, 118, metric_w, 54, '#303B45', '#4A5661', radius=3)
    draw_centered_text(c, 'A 面 8 项', first_x + metric_w/2, 138, FONT_BOLD, 10, WHITE)
    draw_centered_text(c, 'B 面 9 项', second_x + metric_w/2, 138, FONT_BOLD, 10, WHITE)
    draw_footer(c, 9)


PAGE_RENDERERS = [
    draw_cover,
    draw_executive_summary,
    draw_pain_points,
    draw_solution_design,
    draw_human_machine,
    draw_impact,
    draw_reuse,
]


def build_pdf(output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    register_fonts()
    c = canvas.Canvas(str(output_path), pagesize=landscape(A4), pageCompression=1)
    c.setTitle("HPLS 文库表自动化生成方案")
    c.setAuthor("HPLS Project")
    for renderer in PAGE_RENDERERS:
        renderer(c)
        c.showPage()
    c.save()
    return output_path


def parse_args():
    parser = argparse.ArgumentParser(description="生成 HPLS 领导汇报方案 PDF")
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(r"C:\Users\admin\Desktop\HPLS文库表自动化生成方案_领导汇报版.pdf"),
    )
    return parser.parse_args()


def main():
    output = build_pdf(parse_args().output)
    print(output.resolve())


if __name__ == "__main__":
    main()
