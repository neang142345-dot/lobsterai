from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from docx import Document
from docx.shared import Pt as DocxPt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import requests
import os
import re
import json
import time

app = Flask(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "my-secret")
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"

AI_API_KEY = os.getenv("AI_API_KEY")
AI_API_URL = os.getenv("AI_API_URL")


# ========= 基础工具 =========

def send_message(chat_id: int, text: str):
    requests.post(
        f"{TELEGRAM_API}/sendMessage",
        json={"chat_id": chat_id, "text": text},
        timeout=30
    )


def send_document(chat_id: int, file_path: str, caption: str = "文件已生成"):
    with open(file_path, "rb") as f:
        requests.post(
            f"{TELEGRAM_API}/sendDocument",
            data={"chat_id": chat_id, "caption": caption},
            files={"document": f},
            timeout=120
        )


def normalize_text(text: str) -> str:
    if not text:
        return ""
    return text.replace("\r\n", "\n").strip()


def sanitize_filename(name: str) -> str:
    if not name:
        return "文件"
    name = name.strip()
    name = re.sub(r'[\/\\:\*\?"<>\|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return (name or "文件")[:50]


def extract_custom_filename(text: str):
    lines = normalize_text(text).splitlines()
    custom_name = None
    remaining_lines = []

    for line in lines:
        stripped = line.strip()
        if custom_name is None and (stripped.startswith("文件名：") or stripped.startswith("文件名:")):
            custom_name = stripped.split("：", 1)[1].strip() if "：" in stripped else stripped.split(":", 1)[1].strip()
        else:
            remaining_lines.append(line)

    return custom_name, "\n".join(remaining_lines).strip()


def extract_style(text: str):
    lines = normalize_text(text).splitlines()
    style = None
    remaining_lines = []

    for line in lines:
        stripped = line.strip()
        if style is None and (stripped.startswith("风格：") or stripped.startswith("风格:")):
            style = stripped.split("：", 1)[1].strip() if "：" in stripped else stripped.split(":", 1)[1].strip()
        else:
            remaining_lines.append(line)

    return style, "\n".join(remaining_lines).strip()


def extract_theme(text: str):
    lines = normalize_text(text).splitlines()
    theme = None
    remaining_lines = []

    for line in lines:
        stripped = line.strip()
        if theme is None and (stripped.startswith("主题：") or stripped.startswith("主题:")):
            theme = stripped.split("：", 1)[1].strip() if "：" in stripped else stripped.split(":", 1)[1].strip()
        else:
            remaining_lines.append(line)

    return theme, "\n".join(remaining_lines).strip()


def strip_control_phrases(text: str) -> str:
    cleaned = normalize_text(text)

    prefixes = [
        "生成 pdf：", "生成pdf：", "生成 pdf:", "生成pdf:",
        "生成 word：", "生成word：", "生成 word:", "生成word:",
        "生成 excel：", "生成excel：", "生成 excel:", "生成excel:",
        "生成 ppt：", "生成ppt：", "生成 ppt:", "生成ppt:",
        "生成 docx：", "生成docx：", "生成 docx:", "生成docx:",
        "生成 xlsx：", "生成xlsx：", "生成 xlsx:", "生成xlsx:",
        "生成 pptx：", "生成pptx：", "生成 pptx:", "生成pptx:",
        "制作 pdf：", "制作pdf：", "制作 pdf:", "制作pdf:",
        "制作 word：", "制作word：", "制作 word:", "制作word:",
        "制作 excel：", "制作excel：", "制作 excel:", "制作excel:",
        "制作 ppt：", "制作ppt：", "制作 ppt:", "制作ppt:",
        "帮我生成", "帮我制作", "生成文件", "生成文档", "生成", "制作文件", "制作文档", "制作",
        "发我pdf", "回传pdf", "回传我", "发我", "整理成pdf", "整理成word", "整理成ppt", "整理成excel"
    ]

    for prefix in prefixes:
        if cleaned.lower().startswith(prefix.lower()):
            cleaned = cleaned[len(prefix):].strip()
            break

    return cleaned


def build_title(text: str) -> str:
    base = normalize_text(text)
    if not base:
        return "自动生成文档"
    first_line = base.splitlines()[0].strip()
    return first_line if first_line else "自动生成文档"


def build_output_filename(content: str, ext: str, custom_name: str = None, ai_title: str = None) -> str:
    if custom_name:
        base_name = sanitize_filename(custom_name)
    elif ai_title:
        base_name = sanitize_filename(ai_title)
    else:
        base_name = sanitize_filename(build_title(content))

    timestamp = int(time.time())
    return f"/tmp/{base_name}_{timestamp}.{ext}"


def extract_effective_request(text: str, reply_to_message: dict):
    current_text = normalize_text(text)
    reply_text = ""

    if reply_to_message:
        reply_text = normalize_text(reply_to_message.get("text", ""))

    command_like = any(kw in current_text.lower() for kw in [
        "pdf", "docx", "word", "xlsx", "excel", "ppt", "pptx",
        "生成", "回传", "发我", "整理"
    ])

    if reply_text and command_like and len(current_text) <= 30:
        return reply_text + "\n" + current_text

    return current_text


# ========= V5 主题系统 =========

def get_theme_palette(theme: str):
    theme = (theme or "商务蓝").strip()

    if theme == "深色商务":
        return {
            "name": "深色商务",
            "primary": RGBColor(25, 42, 86),
            "secondary": RGBColor(46, 64, 113),
            "accent": RGBColor(93, 173, 226),
            "text": RGBColor(255, 255, 255),
            "subtext": RGBColor(220, 220, 220),
            "light_fill": "D6EAF8",
            "excel_header": "2E4053"
        }

    if theme == "极简白":
        return {
            "name": "极简白",
            "primary": RGBColor(40, 40, 40),
            "secondary": RGBColor(150, 150, 150),
            "accent": RGBColor(91, 155, 213),
            "text": RGBColor(40, 40, 40),
            "subtext": RGBColor(120, 120, 120),
            "light_fill": "F2F4F4",
            "excel_header": "DDEBF7"
        }

    # 默认：商务蓝
    return {
        "name": "商务蓝",
        "primary": RGBColor(44, 95, 153),
        "secondary": RGBColor(91, 155, 213),
        "accent": RGBColor(23, 162, 184),
        "text": RGBColor(30, 30, 30),
        "subtext": RGBColor(100, 100, 100),
        "light_fill": "D9EAF7",
        "excel_header": "D9EAF7"
    }


def choose_subtitle_by_style(style: str):
    if style == "汇报":
        return "管理层汇报材料"
    if style == "商务":
        return "商务演示文稿"
    if style == "正式":
        return "正式演示文稿"
    return "业务分析报告"


# ========= 本地兜底 =========

def detect_intent_and_file_type_fallback(text: str):
    raw = text or ""
    t = raw.lower()

    if "pdf" in t:
        return {"file_type": "pdf", "intent": "generic"}
    if "docx" in t or "word" in t:
        return {"file_type": "docx", "intent": "document"}
    if "xlsx" in t or "excel" in t:
        return {"file_type": "xlsx", "intent": "table"}
    if "pptx" in t or re.search(r"\bppt\b", t):
        return {"file_type": "pptx", "intent": "presentation"}

    if "表格" in raw or "工资表" in raw or "销售表" in raw or "统计表" in raw or "清单" in raw:
        if "工资" in raw:
            return {"file_type": "xlsx", "intent": "salary_table"}
        if "销售" in raw:
            return {"file_type": "xlsx", "intent": "sales_table"}
        return {"file_type": "xlsx", "intent": "table"}

    if "汇报" in raw or "演示" in raw or "路演" in raw or "老板汇报" in raw or "公司经营分析" in raw:
        return {"file_type": "pptx", "intent": "presentation"}

    if "请假申请" in raw or "申请书" in raw:
        return {"file_type": "docx", "intent": "leave"}

    if "合同" in raw or "协议" in raw:
        return {"file_type": "docx", "intent": "contract"}

    if "会议纪要" in raw:
        return {"file_type": "docx", "intent": "meeting_minutes"}

    if "周报" in raw or "月报" in raw or "日报" in raw or "总结" in raw:
        return {"file_type": "docx", "intent": "weekly_report"}

    return {"file_type": "pdf", "intent": "generic"}


def enhance_content(intent: str, content: str, style: str = None) -> str:
    clean = content.strip() if content and content.strip() else "（空内容）"
    lines = [line.strip() for line in clean.splitlines() if line.strip()]

    if intent == "leave":
        if len(lines) <= 3:
            clean = (
                "请假申请\n\n"
                "尊敬的领导：\n"
                f"本人因{clean}，特申请请假一天，请予批准。\n\n"
                "此致\n"
                "敬礼"
            )

    elif intent == "meeting_minutes":
        clean = (
            "会议纪要\n\n"
            f"{clean}\n\n"
            "一、会议结论\n"
            "待补充\n\n"
            "二、行动项\n"
            "1. 待补充\n"
            "2. 待补充"
        )

    elif intent == "weekly_report":
        clean = (
            "工作汇报\n\n"
            "一、本期完成\n"
            f"{clean}\n\n"
            "二、存在问题\n"
            "待补充\n\n"
            "三、下一步计划\n"
            "待补充"
        )

    elif intent == "contract":
        if len(lines) <= 5:
            clean = (
                "合作协议\n\n"
                "甲方：\n"
                "乙方：\n"
                "合作背景：\n"
                f"{clean}\n\n"
                "合作内容：\n"
                "付款方式：\n"
                "双方权责：\n"
                "违约责任：\n"
            )

    return clean


# ========= AI =========

def call_ai(prompt: str):
    if not AI_API_KEY or not AI_API_URL:
        return None

    try:
        resp = requests.post(
            AI_API_URL,
            headers={
                "Authorization": f"Bearer {AI_API_KEY}",
                "Content-Type": "application/json"
            },
            json={
                "model": "gpt-4o-mini",
                "messages": [
                    {
                        "role": "system",
                        "content": (
                            "你是一个专业办公文件助手。"
                            "你必须把用户需求转换成最终可生成文件的结构化结果。"
                            "不要把纯文本回复当成最终结果。"
                            "你的目标是帮助系统生成真实文件：pdf、docx、xlsx、pptx。"
                        )
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "temperature": 0.3
            },
            timeout=45
        )
        if resp.status_code != 200:
            return None

        data = resp.json()
        return data["choices"][0]["message"]["content"]
    except Exception:
        return None


def ai_analyze(text: str, style: str = None):
    style_text = style if style else "默认"

    prompt = f"""
请分析用户需求，并返回严格 JSON，不要输出任何额外解释，不要使用 markdown 代码块。

用户输入：
{text}

文件风格：
{style_text}

输出格式：
{{
  "file_type": "pdf/docx/xlsx/pptx",
  "intent": "generic/document/leave/contract/meeting_minutes/weekly_report/table/salary_table/sales_table/presentation",
  "title": "文件标题",
  "style": "正式/商务/汇报/默认",
  "summary": "一句话概括文件用途",
  "content_text": "适合 PDF/Word 的完整正文",
  "table_headers": ["列1", "列2"],
  "table_rows": [["值1", "值2"]],
  "slides": [
    {{
      "title": "页标题",
      "bullets": ["要点1", "要点2", "要点3"]
    }}
  ]
}}

规则：
1. 目标是生成真实文件，而不是返回说明文字。
2. 表格、工资表、销售表、清单、统计 → xlsx
3. 汇报、演示、路演、PPT、老板汇报、公司经营分析 → pptx
4. 请假申请、合同、会议纪要、周报、日报、月报、正式文书 → docx
5. 其他 → pdf
6. title 必须简洁明确
7. style 使用：正式 / 商务 / 汇报 / 默认
8. content_text 必须是完整可直接写入文件的内容
9. 如果 file_type 不是 xlsx，table_headers 和 table_rows 返回空数组
10. 如果 file_type 不是 pptx，slides 返回空数组
11. 如果 file_type 是 xlsx，尽量生成完整表头，并至少生成 3 行合理示例数据
12. 如果 file_type 是 pptx：
   - 尽量拆成 4 到 6 页
   - 每页只保留最核心信息
   - 每页最多 4 个 bullet
   - bullet 必须是短句，不要写成长段
   - 要适合口头汇报，像老板汇报材料
13. 如果用户说“整理成文件”“发我pdf”“回传我”“生成pdf”之类，仍然必须返回文件所需结构，不要输出说明。
14. 如果用户信息不足，也要尽量合理补全。
"""
    result = call_ai(prompt)
    if not result:
        return None

    try:
        return json.loads(result)
    except Exception:
        return None


# ========= PDF =========

def wrap_text_lines(text: str, c, font_name: str, font_size: int, max_width: float):
    if not text or not text.strip():
        return ["（空内容）"]

    wrapped = []
    raw_lines = text.splitlines()

    for raw in raw_lines:
        line = raw.rstrip()
        if not line.strip():
            wrapped.append("")
            continue

        current = ""
        for ch in line:
            test = current + ch
            if c.stringWidth(test, font_name, font_size) <= max_width:
                current = test
            else:
                if current:
                    wrapped.append(current)
                current = ch

        if current:
            wrapped.append(current)

    return wrapped if wrapped else ["（空内容）"]


def create_pdf(text: str, output_path: str, style: str = None, theme: str = None):
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    palette = get_theme_palette(theme)

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    title = build_title(text)

    # 顶部色带
    c.setFillColorRGB(
        palette["primary"][0] / 255.0,
        palette["primary"][1] / 255.0,
        palette["primary"][2] / 255.0
    )
    c.rect(0, height - 90, width, 90, fill=1, stroke=0)

    # 标题
    c.setFillColorRGB(1, 1, 1)
    c.setFont("STSong-Light", 20)
    c.drawCentredString(width / 2, height - 45, title)

    # 风格 / 主题
    c.setFont("STSong-Light", 10)
    info = f"风格：{style or '默认'}    主题：{palette['name']}"
    c.drawString(40, height - 78, info)

    # 正文
    y = height - 120
    c.setFillColorRGB(0, 0, 0)
    c.setFont("STSong-Light", 12)
    max_body_width = width - 80
    lines = wrap_text_lines(text, c, "STSong-Light", 12, max_body_width)

    for line in lines:
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont("STSong-Light", 12)
        c.drawString(40, y, line)
        y -= 24

    c.save()


# ========= Word =========

def add_doc_divider(doc):
    p = doc.add_paragraph()
    run = p.add_run("─" * 42)
    run.font.size = DocxPt(10)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def create_docx(text: str, output_path: str, style: str = None, theme: str = None):
    palette = get_theme_palette(theme)
    doc = Document()
    title = build_title(text)

    # 标题
    h = doc.add_heading("", level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = h.add_run(title)
    run.font.size = DocxPt(20)
    run.font.bold = True

    # 副标题
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta.add_run(f"风格：{style or '默认'}   主题：{palette['name']}")
    meta_run.font.size = DocxPt(10)

    add_doc_divider(doc)
    doc.add_paragraph("")

    paragraphs = text.splitlines() if text else ["（空内容）"]
    for line in paragraphs:
        p = doc.add_paragraph(line if line.strip() else "")
        p.paragraph_format.space_after = DocxPt(8)
        p.paragraph_format.line_spacing = 1.5
        for r in p.runs:
            r.font.size = DocxPt(12)

    doc.save(output_path)


# ========= Excel =========

def parse_table_text(text: str):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return [["（空内容）"]]

    rows = []
    for line in lines:
        if "\t" in line:
            rows.append([cell.strip() for cell in line.split("\t")])
        elif "," in line:
            rows.append([cell.strip() for cell in line.split(",")])
        elif "，" in line:
            rows.append([cell.strip() for cell in line.split("，")])
        else:
            rows.append([line])
    return rows


def autofit_worksheet(ws):
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value is not None else 0
                if cell_len > max_length:
                    max_length = cell_len
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_length + 4, 30)


def style_excel_sheet(ws, style: str = None, theme: str = None):
    palette = get_theme_palette(theme)
    thin = Side(border_style="thin", color="CCCCCC")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center", horizontal="center")

    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF" if theme == "深色商务" else "000000")
            cell.alignment = Alignment(vertical="center", horizontal="center")
            cell.fill = PatternFill(fill_type="solid", fgColor=palette["excel_header"])


def create_xlsx(text: str, output_path: str, style: str = None, theme: str = None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    rows = parse_table_text(text)
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    style_excel_sheet(ws, style, theme)
    autofit_worksheet(ws)
    wb.save(output_path)


def create_xlsx_structured(output_path: str, headers, rows, style: str = None, theme: str = None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    if not headers:
        headers = ["内容"]
    if not rows:
        rows = [["（空内容）"]]

    for col_idx, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=header)

    for row_idx, row in enumerate(rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    style_excel_sheet(ws, style, theme)
    autofit_worksheet(ws)
    wb.save(output_path)


# ========= PPT =========

def create_slide_header(slide, title_text, palette):
    title_box = slide.shapes.add_textbox(Pt(32), Pt(24), Pt(640), Pt(40))
    tf_title = title_box.text_frame
    tf_title.clear()
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.name = "Microsoft YaHei"
    p_title.font.color.rgb = palette["primary"]

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(32), Pt(70), Pt(640), Pt(3))
    line.fill.solid()
    line.fill.fore_color.rgb = palette["secondary"]
    line.line.fill.background()


def create_pptx_structured(output_path: str, title: str, slides_data, style: str = None, theme: str = None):
    prs = Presentation()
    palette = get_theme_palette(theme)

    # 封面页
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 背景标题块
    block = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(48), Pt(110), Pt(620), Pt(120))
    block.fill.solid()
    block.fill.fore_color.rgb = palette["primary"]
    block.line.fill.background()

    # 标题
    title_box = slide.shapes.add_textbox(Pt(70), Pt(135), Pt(580), Pt(50))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.name = "Microsoft YaHei"
    p.font.color.rgb = RGBColor(255, 255, 255)

    # 副标题
    subtitle = choose_subtitle_by_style(style)
    sub_box = slide.shapes.add_textbox(Pt(70), Pt(185), Pt(580), Pt(30))
    tf2 = sub_box.text_frame
    tf2.clear()
    p2 = tf2.paragraphs[0]
    p2.text = f"{subtitle}｜主题：{palette['name']}"
    p2.font.size = Pt(14)
    p2.font.name = "Microsoft YaHei"
    p2.font.color.rgb = RGBColor(255, 255, 255)

    # 内容页
    if not slides_data:
        slides_data = [{"title": "内容", "bullets": ["（空内容）"]}]

    for item in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        create_slide_header(slide, item.get("title", "内容"), palette)

        content_box = slide.shapes.add_textbox(Pt(52), Pt(110), Pt(610), Pt(330))
        tf = content_box.text_frame
        tf.word_wrap = True
        tf.clear()

        bullets = item.get("bullets", []) or ["（空内容）"]
        bullets = bullets[:4]

        for i, bullet in enumerate(bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.size = Pt(20 if style == "汇报" else 18)
            p.font.name = "Microsoft YaHei"
            p.font.color.rgb = palette["text"]
            p.space_after = Pt(10)

        # 页脚
        footer = slide.shapes.add_textbox(Pt(52), Pt(470), Pt(300), Pt(20))
        tf_footer = footer.text_frame
        p_footer = tf_footer.paragraphs[0]
        p_footer.text = f"{subtitle}"
        p_footer.font.size = Pt(9)
        p_footer.font.name = "Microsoft YaHei"
        p_footer.font.color.rgb = palette["subtext"]

    prs.save(output_path)


def create_pptx(text: str, output_path: str, style: str = None, theme: str = None):
    title = build_title(text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    bullets = lines[1:] if len(lines) > 1 else ["（空内容）"]
    slides_data = [{"title": title, "bullets": bullets[:4]}]
    create_pptx_structured(output_path, title, slides_data, style, theme)


# ========= 路由 =========

@app.route("/")
def home():
    return "TG office assistant bot is running!"


@app.route("/health")
def health():
    return {"status": "ok"}


@app.route("/webhook", methods=["POST"])
def webhook():
    secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if secret != WEBHOOK_SECRET:
        return jsonify({"ok": False, "error": "unauthorized"}), 403

    update = request.get_json(silent=True) or {}
    message = update.get("message", {})
    chat = message.get("chat", {})
    chat_id = chat.get("id")
    text = message.get("text", "")
    reply_to_message = message.get("reply_to_message", {})

    if not chat_id:
        return jsonify({"ok": True})

    if text == "/start":
        send_message(
            chat_id,
            "你好，直接发送内容即可。我会自动识别并生成 PDF / Word / Excel / PPT 真文件。支持“文件名：xxx”“风格：商务/正式/汇报”“主题：商务蓝/深色商务/极简白”。"
        )
        return jsonify({"ok": True})

    if text.startswith("/") and text != "/start":
        send_message(chat_id, "请直接发送正文内容，不需要输入命令。")
        return jsonify({"ok": True})

    merged_text = extract_effective_request(text, reply_to_message)

    custom_name, text_without_filename = extract_custom_filename(merged_text)
    style, text_without_style = extract_style(text_without_filename)
    theme, text_without_theme = extract_theme(text_without_style)
    cleaned = strip_control_phrases(text_without_theme)

    if not cleaned:
        send_message(chat_id, "请发送你要生成成文件的文字内容。")
        return jsonify({"ok": True})

    lowered = cleaned.lower()
    forced_type = None
    if "pdf" in lowered:
        forced_type = "pdf"
    elif "docx" in lowered or "word" in lowered:
        forced_type = "docx"
    elif "xlsx" in lowered or "excel" in lowered:
        forced_type = "xlsx"
    elif "pptx" in lowered or re.search(r"\bppt\b", lowered):
        forced_type = "pptx"

    ai_result = ai_analyze(cleaned, style)

    if ai_result:
        file_type = forced_type or ai_result.get("file_type", "pdf")
        ai_title = ai_result.get("title", build_title(cleaned))
        ai_style = ai_result.get("style", style or "默认")
        content_text = ai_result.get("content_text", cleaned)
        table_headers = ai_result.get("table_headers", [])
        table_rows = ai_result.get("table_rows", [])
        slides = ai_result.get("slides", [])
    else:
        route = detect_intent_and_file_type_fallback(cleaned)
        file_type = forced_type or route["file_type"]
        intent = route["intent"]
        ai_title = build_title(cleaned)
        ai_style = style or "默认"
        content_text = enhance_content(intent, cleaned, style)
        table_headers = []
        table_rows = []
        slides = []

    if not content_text or not normalize_text(content_text):
        content_text = cleaned

    send_message(chat_id, f"已收到，正在为你生成 {file_type} 文件...")

    try:
        if file_type == "pdf":
            file_path = build_output_filename(content_text, "pdf", custom_name, ai_title)
            create_pdf(content_text, file_path, ai_style, theme)
            send_document(chat_id, file_path, "你的 PDF 已生成")

        elif file_type == "docx":
            file_path = build_output_filename(content_text, "docx", custom_name, ai_title)
            create_docx(content_text, file_path, ai_style, theme)
            send_document(chat_id, file_path, "你的 Word 文档已生成")

        elif file_type == "xlsx":
            file_path = build_output_filename(content_text, "xlsx", custom_name, ai_title)
            if table_headers or table_rows:
                create_xlsx_structured(file_path, table_headers, table_rows, ai_style, theme)
            else:
                create_xlsx(content_text, file_path, ai_style, theme)
            send_document(chat_id, file_path, "你的 Excel 文件已生成")

        elif file_type == "pptx":
            file_path = build_output_filename(content_text, "pptx", custom_name, ai_title)
            if slides:
                create_pptx_structured(file_path, ai_title, slides, ai_style, theme)
            else:
                create_pptx(content_text, file_path, ai_style, theme)
            send_document(chat_id, file_path, "你的 PPT 文件已生成")

        else:
            file_path = build_output_filename(content_text, "pdf", custom_name, ai_title)
            create_pdf(content_text, file_path, ai_style, theme)
            send_document(chat_id, file_path, "你的 PDF 已生成")

    except Exception as e:
        send_message(chat_id, f"生成失败：{str(e)}")
        return jsonify({"ok": False, "error": str(e)}), 500

    return jsonify({"ok": True})


@app.route("/set_webhook")
def set_webhook():
    url = os.getenv("RENDER_EXTERNAL_URL", "").rstrip("/") + "/webhook"
    if not BOT_TOKEN or not url:
        return {
            "ok": False,
            "error": "missing BOT_TOKEN or RENDER_EXTERNAL_URL"
        }, 400

    resp = requests.post(
        f"{TELEGRAM_API}/setWebhook",
        json={
            "url": url,
            "secret_token": WEBHOOK_SECRET
        },
        timeout=30
    )
    return resp.json(), resp.status_code


if __name__ == "__main__":
    port = int(os.getenv("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
