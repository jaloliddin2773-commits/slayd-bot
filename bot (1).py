"""Telegram slayd boti. Ishga tushirish: python bot.py"""
import logging
import os
import re
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

from telegram import Update
from telegram.constants import ChatAction
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from slides import ask_gemini, build_pptx

logging.basicConfig(level=logging.INFO)
TOKEN = os.environ["TELEGRAM_TOKEN"]

WELCOME = (
    "Salom! 👋 Men taqdimot (slayd) tayyorlab beraman.\n\n"
    "Mavzuni yozing, men PowerPoint fayl yuboraman.\n\n"
    "Masalan:\n"
    "• Global isish haqida taqdimot, 8 slayd\n"
    "• O'zbekiston tarixi\n"
    "• Sog'lom ovqatlanish, 6 slayd"
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(WELCOME)


async def make_slides(update: Update, context: ContextTypes.DEFAULT_TYPE):
    topic = (update.message.text or "").strip()
    if len(topic) < 3:
        await update.message.reply_text("Mavzuni to'liqroq yozing, iltimos.")
        return

    wait = await update.message.reply_text("⏳ Taqdimot tayyorlanmoqda, 20-30 soniya kuting...")
    await update.message.chat.send_action(ChatAction.UPLOAD_DOCUMENT)

    try:
        data = ask_gemini(topic)
        name = re.sub(r"[^\w\-]+", "_", data.get("title", "taqdimot"))[:40] or "taqdimot"
        with tempfile.TemporaryDirectory() as tmp:
            path = build_pptx(data, os.path.join(tmp, f"{name}.pptx"))
            with open(path, "rb") as f:
                await update.message.reply_document(
                    f,
                    filename=f"{name}.pptx",
                    caption=f"✅ Tayyor! {len(data['slides'])} ta slayd.",
                )
        await wait.delete()
    except Exception:
        logging.exception("Xatolik")
        await wait.edit_text("😔 Xatolik bo'ldi. Birozdan keyin qayta urinib ko'ring.")


class _Ping(BaseHTTPRequestHandler):
    """Render kabi xizmatlar port so'raydi, shu uchun kichik server."""

    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")

    def log_message(self, *args):
        pass


def _serve_ping():
    port = int(os.environ.get("PORT", "10000"))
    HTTPServer(("0.0.0.0", port), _Ping).serve_forever()


def main():
    threading.Thread(target=_serve_ping, daemon=True).start()
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, make_slides))
    app.run_polling()


if __name__ == "__main__":
    main()
