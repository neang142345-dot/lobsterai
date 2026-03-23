from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from pptx import Presentation
from pptx.util import Pt
import requests
import os
import re
import json

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
            timeout=60
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


def strip_leading_type_command(text: str) -> str:
    cleaned = normalize_text(text)
    prefixes = [
        "生成 pdf：", "生成pdf：", "生成 pdf:", "生成pdf:",
        "生成 word：", "生成word：", "生成 word:", "生成word:",
        "生成 excel：", "生成excel：", "生成 excel:", "生成excel:",
        "生成 ppt：", "生成ppt：", "生成 ppt:", "生成ppt:",
        "制作 pdf：", "制作pdf：", "制作 pdf:", "制作pdf:",
        "制作 word：", "制作word：", "制作 word:", "制作word:",
        "制作 excel：", "制作excel：", "制作 excel:", "制作excel:",
        "制作 ppt：", "制作ppt：", "制作 ppt:", "制作ppt:",
        "帮我生成", "帮我制作", "生成文件", "生成文档", "生成", "制作文件", "制作文档", "制作"
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
    return f"/tmp/{base_name}.{ext}"


# ========= 本地兜底 =========

def detect_intent_and_file_type_fallback(text: str):
    raw = text or ""
    t = raw.lower()

    if "xlsx" in t or "excel" in t or "表格" in raw or "工资表" in raw or "销售表" in raw or "统计表" in raw:
        if "工资" in raw:
            return {"file_type": "xlsx", "intent": "salary_table"}
        if "销售" in raw:
            return {"file_type": "xlsx", "intent": "sales_table"}
        return {"file_type": "xlsx", "intent": "table"}

    if "pptx" in t or re.search(r"\bppt\b", t) or "幻灯片" in raw or "演示文稿" in raw or "汇报" in raw or "路演" in raw:
        return {"file_type": "pptx", "intent": "presentation"}

    if "docx" in t or "word" in t or "请假申请" in raw or "合同" in raw or "申请书" in raw:
        if "请假" in raw:
            return {"file_type": "docx", "intent": "leave"}
        if "合同" in raw:
            return {"file_type": "docx", "intent": "contract"}
        return {"file_type": "docx", "intent": "document"}

    if "会议纪要" in raw:
        return {"file_type": "docx", "intent": "meeting_minutes"}

    if "周报" in raw or "月报" in raw or "日报" in raw:
        return {"file_type": "docx", "intent": "weekly_report"}

    if "pdf" in t:
        return {"file_type": "pdf", "intent": "generic"}

    return {"file_type": "pdf", "intent": "generic"}


def enhance_content(intent: str, content: str, style: str = None) -> str:
    clean = content.strip() if content and content.strip() else "（空内容）"
    lines = [line.strip() for line in clean.splitlines() if line.strip()]

    if intent == "leave":
        if len(lines) <= 3:
            clean = (
                "请假申请\n"
                "尊敬的领导：\n"
                f"本人因{clean}，特申请请假一天，请予批准。\n"
                "此致\n"
                "敬礼"
            )

    elif intent == "meeting_minutes":
        clean = (
            "会议纪要\n"
            f"{clean}\n\n"
            "一、会议结论\n"
            "待补充\n\n"
            "二、行动项\n"
            "1. 待补充\n"
            "2. 待补充"
        )

    elif intent == "weekly_report":
        clean = (
            "工作汇报\n"
            "一、本周完成\n"
            f"{clean}\n\n"
            "二、存在问题\n"
            "待补充\n\n"
            "三、下周计划\n"
            "待补充"
        )

    elif intent == "contract":
        if len(lines) <= 5:
            clean = (
                "合作协议\n"
                "甲方：\n"
                "乙方：\n"
                "合作背景：\n"
                f"{clean}\n"
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
                            "你擅长判断文件类型、补全文档内容、生成更完整的表格结构、"
                            "以及把汇报拆成更适合PPT展示的页面结构。"
                        )
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                "temperature": 0.3
            },
            timeout=30
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
1. 表格、工资表、销售表、清单、统计 → xlsx
2. 汇报、演示、路演、PPT → pptx
3. 请假申请、合同、会议纪要、周报、日报、月报、正式文书 → docx
4. 其他 → pdf
5. title 必须简洁明确
6. style 使用：正式 / 商务 / 汇报 / 默认
7. content_text 必须是完整可直接写入文件的内容
8. 如果 file_type 不是 xlsx，table_headers 和 table_rows 返回空数组
9. 如果 file_type 不是 pptx，slides 返回空数组
10. 如果 file_type 是 xlsx，尽量生成完整表头，并至少生成 3 行合理示例数据
11. 如果 file_type 是 pptx：
   - 尽量拆成 4 到 6 页
   - 每页只保留最核心信息
   - 每页最多 4 个 bullet
   - bullet 必须是短句，不要写成长段
   - 要适合口头汇报，像老板汇报材料
12. 如果用户信息不足，也要尽量合理补全
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
        line = raw.strip()
        if not line:
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


def create_pdf(text: str, output_path: str, style: str = None):
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    title = build_title(text)
    y = height - 50
    max_title_width = width - 100

    c.setFont("STSong-Light", 18)

    original_title = title
    while c.stringWidth(title, "STSong-Light", 18) > max_title_width and len(title) > 1:
        title = title[:-1]
    if title != original_title and len(title) > 1:
        title = title[:-1] + "…"

    title_width = c.stringWidth(title, "STSong-Light", 18)
    c.drawString((width - title_width) / 2, y, title)
    y -= 40

    if style:
        c.setFont("STSong-Light", 10)
        c.drawString(50, y, f"风格：{style}")
        y -= 20

    c.setFont("STSong-Light", 12)
    max_body_width = width - 100
    lines = wrap_text_lines(text, c, "STSong-Light", 12, max_body_width)

    for line in lines:
        if y < 50:
            c.showPage()
            c.setFont("STSong-Light", 12)
            y = height - 50
        c.drawString(50, y, line)
        y -= 22

    c.save()


# ========= Word =========

def create_docx(text: str, output_path: str, style: str = None):
    doc = Document()
    title = build_title(text)

    doc.add_heading(title, level=1)

    if style:
        doc.add_paragraph(f"风格：{style}")

    paragraphs = text.splitlines() if text else ["（空内容）"]
    for p in paragraphs:
        doc.add_paragraph(p if p.strip() else "")

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


def style_excel_sheet(ws, style: str = None):
    thin = Side(border_style="thin", color="CCCCCC")

    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center")

    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(vertical="center", horizontal="center")
            if style in ("正式", "商务", "汇报"):
                cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")


def create_xlsx(text: str, output_path: str, style: str = None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    rows = parse_table_text(text)
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    style_excel_sheet(ws, style)
    autofit_worksheet(ws)
    wb.save(output_path)


def create_xlsx_structured(output_path: str, headers, rows, style: str = None):
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

    style_excel_sheet(ws, style)
    autofit_worksheet(ws)
    wb.save(output_path)


# ========= PPT =========

def parse_ppt_content(text: str):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return "自动生成演示文稿", ["（空内容）"]

    title = lines[0]
    bullets = lines[1:] if len(lines) > 1 else ["（无补充内容）"]
    return title, bullets


def create_pptx_structured(output_path: str, title: str, slides_data, style: str = None):
    prs = Presentation()

    # ========= 封面页 =========
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title

    subtitle = "业务分析报告"
    if style == "汇报":
        subtitle = "管理层汇报材料"
    elif style == "商务":
        subtitle = "商务演示文稿"
    elif style == "正式":
        subtitle = "正式演示文稿"

    slide.placeholders[1].text = subtitle

    title_para = slide.shapes.title.text_frame.paragraphs[0]
    title_para.font.size = Pt(30)
    title_para.font.bold = True

    sub_para = slide.placeholders[1].text_frame.paragraphs[0]
    sub_para.font.size = Pt(16)

    # ========= 内容页 =========
    if not slides_data:
        slides_data = [{"title": "内容", "bullets": ["（空内容）"]}]

    for item in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[5])

        title_box = slide.shapes.add_textbox(Pt(36), Pt(24), Pt(620), Pt(40))
        tf_title = title_box.text_frame
        tf_title.clear()
        p_title = tf_title.paragraphs[0]
        p_title.text = item.get("title", "内容")
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.name = "微软雅黑"

        sub_box = slide.shapes.add_textbox(Pt(36), Pt(60), Pt(620), Pt(24))
        tf_sub = sub_box.text_frame
        tf_sub.clear()
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = "核心要点"
        p_sub.font.size = Pt(10)

        content_box = slide.shapes.add_textbox(Pt(50), Pt(110), Pt(620), Pt(320))
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
            p.font.name = "微软雅黑"

    prs.save(output_path)


def create_pptx(text: str, output_path: str, style: str = None):
    title, bullets = parse_ppt_content(text)
    slides_data = [{"title": title, "bullets": bullets[:4]}]
    create_pptx_structured(output_path, title, slides_data, style)


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

    if not chat_id:
        return jsonify({"ok": True})

    if text == "/start":
        send_message(
            chat_id,
            "你好，直接发送内容即可。我会自动识别并生成 PDF / Word / Excel / PPT 真文件。支持“文件名：xxx”和“风格：商务/正式/汇报”。"
        )
        return jsonify({"ok": True})

    if text.startswith("/"):
        send_message(chat_id, "请直接发送正文内容，不需要输入命令。")
        return jsonify({"ok": True})

    custom_name, text_without_filename = extract_custom_filename(text)
    style, text_without_style = extract_style(text_without_filename)
    cleaned = strip_leading_type_command(text_without_style)

    if not cleaned:
        send_message(chat_id, "请发送你要生成成文件的文字内容。")
        return jsonify({"ok": True})

    ai_result = ai_analyze(cleaned, style)

    if ai_result:
        file_type = ai_result.get("file_type", "pdf")
        ai_title = ai_result.get("title", build_title(cleaned))
        ai_style = ai_result.get("style", style or "默认")
        content_text = ai_result.get("content_text", cleaned)
        table_headers = ai_result.get("table_headers", [])
        table_rows = ai_result.get("table_rows", [])
        slides = ai_result.get("slides", [])
    else:
        route = detect_intent_and_file_type_fallback(cleaned)
        file_type = route["file_type"]
        intent = route["intent"]
        ai_title = build_title(cleaned)
        ai_style = style or "默认"
        content_text = enhance_content(intent, cleaned, style)
        table_headers = []
        table_rows = []
        slides = []

    send_message(chat_id, f"已收到，正在为你生成 {file_type} 文件...")

    try:
        if file_type == "pdf":
            file_path = build_output_filename(content_text, "pdf", custom_name, ai_title)
            create_pdf(content_text, file_path, ai_style)
            send_document(chat_id, file_path, "你的 PDF 已生成")

        elif file_type == "docx":
            file_path = build_output_filename(content_text, "docx", custom_name, ai_title)
            create_docx(content_text, file_path, ai_style)
            send_document(chat_id, file_path, "你的 Word 文档已生成")

        elif file_type == "xlsx":
            file_path = build_output_filename(content_text, "xlsx", custom_name, ai_title)
            if table_headers or table_rows:
                create_xlsx_structured(file_path, table_headers, table_rows, ai_style)
            else:
                create_xlsx(content_text, file_path, ai_style)
            send_document(chat_id, file_path, "你的 Excel 文件已生成")

        elif file_type == "pptx":
            file_path = build_output_filename(content_text, "pptx", custom_name, ai_title)
            if slides:
                create_pptx_structured(file_path, ai_title, slides, ai_style)
            else:
                create_pptx(content_text, file_path, ai_style)
            send_document(chat_id, file_path, "你的 PPT 文件已生成")

        else:
            file_path = build_output_filename(content_text, "pdf", custom_name, ai_title)
            create_pdf(content_text, file_path, ai_style)
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
