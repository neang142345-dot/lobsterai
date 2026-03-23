from flask import Flask, request, jsonify
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import requests
import os
import time

app = Flask(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "my-secret")
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"

def create_pdf(text: str, output_path: str):
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    y = height - 50
    title = text.strip() if text and text.strip() else "自动生成文档"
    max_width = width - 100

    c.setFont("STSong-Light", 18)

    while c.stringWidth(title, "STSong-Light", 18) > max_width and len(title) > 1:
          title = title[:-1]

    if title != text.strip():
    title = title[:-1] + "…"

    text_width = c.stringWidth(title, "STSong-Light", 18)
    c.drawString((width - text_width) / 2, y, title)
    y -= 40

    c.setFont("STSong-Light", 12)

    max_chars_per_line = 28
    lines = []
    raw_lines = text.splitlines() if text else ["(empty)"]

    for raw in raw_lines:
        if not raw:
            lines.append("")
            continue
        for i in range(0, len(raw), max_chars_per_line):
            lines.append(raw[i:i + max_chars_per_line])

    for line in lines:
        if y < 50:
            c.showPage()
            c.setFont("STSong-Light", 12)
            y = height - 50
        c.drawString(50, y, line)
        y -= 22

    c.save()

def send_message(chat_id: int, text: str):
    requests.post(
        f"{TELEGRAM_API}/sendMessage",
        json={
            "chat_id": chat_id,
            "text": text
        },
        timeout=30
    )

def send_document(chat_id: int, file_path: str, caption: str = "文件已生成"):
    with open(file_path, "rb") as f:
        requests.post(
            f"{TELEGRAM_API}/sendDocument",
            data={
                "chat_id": chat_id,
                "caption": caption
            },
            files={
                "document": f
            },
            timeout=60
        )

@app.route("/")
def home():
    return "TG PDF bot is running!"

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
    clean_text = text.replace("生成文件", "").strip()

    if not chat_id:
        return jsonify({"ok": True})

    if text == "/start":
        send_message(
            chat_id,
            "你好，直接发送内容，我会自动生成 PDF 给你。"
        )
        return jsonify({"ok": True})

    if text.startswith("/"):
        send_message(chat_id, "请直接发送正文内容。")
        return jsonify({"ok": True})

    if text.strip():
        send_message(chat_id, "已收到，正在生成 PDF...")

        timestamp = int(time.time())
        file_path = f"/tmp/generated_{timestamp}.pdf"

        create_pdf(clean_text, file_path)
        send_document(chat_id, file_path, "你的 PDF 已生成")

        return jsonify({"ok": True})

    send_message(chat_id, "请发送内容。")
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
