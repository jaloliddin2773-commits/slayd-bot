import asyncio
import logging
import os
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, HTTPServer

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from slides import ask_gemini, build_pptx


# =========================================================
# CONFIG
# =========================================================

TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()

PORT = int(os.getenv("PORT", "10000"))

MAX_REQUEST_LENGTH = 2000
COOLDOWN_SECONDS = 10


# =========================================================
# LOGGING
# =========================================================

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    level=logging.INFO,
)

logger = logging.getLogger("slayd-bot")


# =========================================================
# RENDER HEALTH SERVER
# =========================================================

class HealthHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"SLAYD BOT OK")

    def log_message(self, format, *args):
        return


def start_health_server():
    server = HTTPServer(
        ("0.0.0.0", PORT),
        HealthHandler,
    )

    thread = threading.Thread(
        target=server.serve_forever,
        daemon=True,
    )

    thread.start()

    logger.info("Health server started on port %s", PORT)


# =========================================================
# START
# =========================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "👋 Assalomu alaykum!\n\n"
        "Men AI yordamida PowerPoint prezentatsiya yarataman.\n\n"
        "Masalan:\n\n"
        "📚 8 slaydlik \"Sun'iy intellekt\" "
        "mavzusida prezentatsiya yarat.\n\n"
        "Yoki:\n\n"
        "🎓 Ingliz tilida 10 slaydlik "
        "IELTS haqida professional prezentatsiya qil.\n\n"
        "Mavzuni yozing — qolganini men qilaman."
    )

    await update.message.reply_text(text)


# =========================================================
# HELP
# =========================================================

async def help_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    text = (
        "📖 Qanday ishlaydi?\n\n"
        "1️⃣ Mavzuni yozasiz.\n"
        "2️⃣ Slayd sonini ko'rsatasiz.\n"
        "3️⃣ AI kontentni tayyorlaydi.\n"
        "4️⃣ Men PowerPoint fayl yarataman.\n"
        "5️⃣ Tayyor .pptx faylni yuboraman.\n\n"
        "Misol:\n"
        "«Global warming haqida 7 slaydlik "
        "inglizcha prezentatsiya yarat»"
    )

    await update.message.reply_text(text)


# =========================================================
# GENERATE
# =========================================================

async def generate_presentation(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    if not update.message or not update.message.text:
        return

    user_id = update.effective_user.id

    request_text = update.message.text.strip()

    if not request_text:
        return

    if len(request_text) > MAX_REQUEST_LENGTH:
        await update.message.reply_text(
            f"❌ So'rov juda uzun.\n"
            f"Maksimal: {MAX_REQUEST_LENGTH} ta belgi."
        )
        return

    # -----------------------------------------------------
    # Cooldown
    # -----------------------------------------------------

    last_requests = context.application.bot_data.setdefault(
        "last_requests",
        {},
    )

    now = time.time()
    last_time = last_requests.get(user_id, 0)

    if now - last_time < COOLDOWN_SECONDS:
        wait = int(COOLDOWN_SECONDS - (now - last_time)) + 1

        await update.message.reply_text(
            f"⏳ Biroz kuting: {wait} soniya."
        )

        return

    last_requests[user_id] = now

    # -----------------------------------------------------
    # Concurrency limit
    # -----------------------------------------------------

    semaphore = context.application.bot_data["generation_semaphore"]

    async with semaphore:

        status = await update.message.reply_text(
            "🧠 Mavzu tahlil qilinmoqda...\n"
            "Reja va slaydlar tayyorlanmoqda."
        )

        try:
            # ---------------------------------------------
            # AI
            # ---------------------------------------------

            presentation_data = await asyncio.to_thread(
                ask_gemini,
                request_text,
            )

            model = presentation_data.get(
                "_model",
                "Gemini",
            )

            await status.edit_text(
                "🎨 Dizayn tayyorlanmoqda...\n"
                f"AI: {model}"
            )

            # ---------------------------------------------
            # PPTX
            # ---------------------------------------------

            with tempfile.TemporaryDirectory() as temp_dir:

                file_path = os.path.join(
                    temp_dir,
                    "presentation.pptx",
                )

                await asyncio.to_thread(
                    build_pptx,
                    presentation_data,
                    file_path,
                )

                title = presentation_data.get(
                    "title",
                    "Presentation",
                )

                await status.edit_text(
                    "✅ Tayyor!\n"
                    "PowerPoint fayl yuborilmoqda..."
                )

                with open(file_path, "rb") as document:

                    await update.message.reply_document(
                        document=document,
                        filename="presentation.pptx",
                        caption=(
                            f"🎓 {title}\n\n"
                            "Generated by AI Presentation Bot"
                        ),
                    )

                await status.delete()

        except Exception as exc:

            logger.exception(
                "Presentation generation failed: %s",
                exc,
            )

            await status.edit_text(
                "❌ Prezentatsiyani yaratishda xatolik yuz berdi.\n\n"
                "Server yoki AI vaqtincha band bo‘lishi mumkin.\n"
                "Birozdan keyin yana urinib ko‘ring."
            )


# =========================================================
# ERROR HANDLER
# =========================================================

async def error_handler(
    update: object,
    context: ContextTypes.DEFAULT_TYPE,
):
    logger.exception(
        "Telegram error:",
        exc_info=context.error,
    )


# =========================================================
# MAIN
# =========================================================

def main():

    if not TOKEN:
        raise RuntimeError(
            "TELEGRAM_TOKEN environment variable is missing."
        )

    if not os.getenv("GEMINI_API_KEY", "").strip():
        raise RuntimeError(
            "GEMINI_API_KEY environment variable is missing."
        )

    start_health_server()

    application = (
        Application.builder()
        .token(TOKEN)
        .build()
    )

    application.bot_data["generation_semaphore"] = asyncio.Semaphore(2)
    application.bot_data["last_requests"] = {}

    application.add_handler(
        CommandHandler("start", start)
    )

    application.add_handler(
        CommandHandler("help", help_command)
    )

    application.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            generate_presentation,
        )
    )

    application.add_error_handler(
        error_handler
    )

    logger.info("SLAYD BOT starting...")

    application.run_polling(
        drop_pending_updates=True
    )


if __name__ == "__main__":
    main()
