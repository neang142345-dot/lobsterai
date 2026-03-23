from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches, Pt
import requests
import os
import re
import time

app = Flask(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "my-secret")
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"


# ========= 通用工具 =========

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


def detect_file_type(text: str) -> str:
    t = (text or "").lower()

    if "xmind" in t or "思维导图" in text:
        return "xmind"
    if "pptx" in t or re.search(r"\bppt\b", t) or "幻灯片" in text or "演示文稿" in text:
        return "pptx"
    if "xlsx" in t or "excel" in t or "表格" in text:
        return "xlsx"
    if "docx" in t or "word" in t or "文档" in text:
        return "docx"
    if "pdf" in t:
        return "pdf"

    # 默认仍然给 PDF
    return "pdf"


def strip_leading_type_command(text: str) -> str:
    cleaned = normalize_text(text)

    prefixes = [
        "生成 pdf：", "生成pdf：", "生成 pdf:", "生成pdf:",
        "生成 word：", "生成word：", "生成 word:", "生成word:",
        "生成 excel：", "生成excel：", "生成 excel:", "生成excel:",
        "生成 ppt：", "生成ppt：", "生成 ppt:", "生成ppt:",
        "生成 xmind：", "生成xmind：", "生成 xmind:", "生成xmind:",
        "生成 思维导图：", "生成 思维导图:", "生成思维导图：", "生成思维导图:",
        "制作 pdf：", "制作pdf：", "制作 pdf:", "制作pdf:",
        "制作 word：", "制作word：", "制作 word:", "制作word:",
        "制作 excel：", "制作excel：", "制作 excel:", "制作excel:",
        "制作 ppt：", "制作ppt：", "制作 ppt:", "制作ppt:",
        "制作 xmind：", "制作xmind：", "制作 xmind:", "制作xmind:",
        "制作 思维导图：", "制作 思维导图:", "制作思维导图：", "制作思维导图:",
        "帮我生成 pdf：", "帮我生成word：", "帮我生成 excel：", "帮我生成 ppt：",
        "帮我生成 xmind：", "帮我制作 pdf：", "帮我制作word：", "帮我制作 excel：",
        "帮我制作 ppt：", "帮我制作 xmind：",
        "生成文件", "生成文档", "生成", "制作文件", "制作文档", "制作", "帮我生成", "帮我制作"
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


def safe_filename(ext: str) -> str:
    return f"/tmp/generated_{int(time.time())}.{ext}"


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


def parse_table_text(text: str):
    """
    Excel 简单解析规则：
    1. 多行，逗号/中文逗号/制表符分列
    2. 没有明显分隔符时，一行一列
    """
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


def parse_ppt_content(text: str):
    """
    简单规则：
    - 第一行做标题
    - 后续每行做 bullet
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return "自动生成演示文稿", ["（空内容）"]

    title = lines[0]
    bullets = lines[1:] if len(lines) > 1 else ["（无补充内容）"]
    return title, bullets


def parse_mindmap_outline(text: str):
    """
    简单导图规则：
    - 第一行做主题
    - 后续每行做一级节点
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return "思维导图主题", ["（空内容）"]

    root = lines[0]
    children = lines[1:] if len(lines) > 1 else ["（空内容）"]
    return root, children


# ========= 文件生成 =========

def create_pdf(text: str, output_path: str):
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


def create_docx(text: str, output_path: str):
    doc = Document()
    title = build_title(text)

    doc.add_heading(title, level=1)

    paragraphs = text.splitlines() if text else ["（空内容）"]
    for p in paragraphs:
        doc.add_paragraph(p if p.strip() else "")

    doc.save(output_path)


def create_xlsx(text: str, output_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    rows = parse_table_text(text)
    for row_idx, row in enumerate(rows, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    wb.save(output_path)


def create_pptx(text: str, output_path: str):
    prs = Presentation()
    title, bullets = parse_ppt_content(text)

    # 标题页
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = "由 TG 文件机器人自动生成"

    # 内容页
    chunk_size = 6
    for i in range(0, len(bullets), chunk_size):
        chunk = bullets[i:i + chunk_size]
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"{title}（续）" if i > 0 else "内容"

        tf = slide.placeholders[1].text_frame
        tf.clear()

        for j, bullet in enumerate(chunk):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.size = Pt(20)

    prs.save(output_path)


def create_xmind_markdown(text: str, output_path: str):
    """
    先输出 Markdown 脑图大纲文件。
    后续如果要原生 .xmind，只替换这个函数。
    """
    root, children = parse_mindmap_outline(text)

    lines = [f"# {root}", ""]
    for child in children:
        lines.append(f"- {child}")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# ========= 路由 =========

@app.route("/")
def home():
    return "TG multi-file bot is running!"


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
            "你好，直接发送内容即可。我会自动识别并生成 PDF / Word / Excel / PPT / 脑图大纲文件。"
        )
        return jsonify({"ok": True})

    if text.startswith("/"):
        send_message(chat_id, "请直接发送正文内容，不需要输入命令。")
        return jsonify({"ok": True})

    file_type = detect_file_type(text)
    content = strip_leading_type_command(text)

    if not content:
        send_message(chat_id, "请发送你要生成成文件的文字内容。")
        return jsonify({"ok": True})

    send_message(chat_id, f"已收到，正在为你生成 {file_type} 文件...")

    try:
        if file_type == "pdf":
            file_path = safe_filename("pdf")
            create_pdf(content, file_path)
            send_document(chat_id, file_path, "你的 PDF 已生成")

        elif file_type == "docx":
            file_path = safe_filename("docx")
            create_docx(content, file_path)
            send_document(chat_id, file_path, "你的 Word 文档已生成")

        elif file_type == "xlsx":
            file_path = safe_filename("xlsx")
            create_xlsx(content, file_path)
            send_document(chat_id, file_path, "你的 Excel 文件已生成")

        elif file_type == "pptx":
            file_path = safe_filename("pptx")
            create_pptx(content, file_path)
            send_document(chat_id, file_path, "你的 PPT 文件已生成")

        elif file_type == "xmind":
            file_path = safe_filename("md")
            create_xmind_markdown(content, file_path)
            send_document(chat_id, file_path, "你的脑图大纲文件已生成（Markdown 版）")

        else:
            file_path = safe_filename("pdf")
            create_pdf(content, file_path)
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
