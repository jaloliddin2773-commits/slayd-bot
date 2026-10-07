import asyncio
import logging
import os
import re
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, HTTPServer

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from slides import (
    ask_gemini,
    build_pptx,
    build_pdf,
)


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
        self.send_header(
            "Content-Type",
            "text/plain; charset=utf-8",
        )
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

    logger.info(
        "Health server started on port %s",
        PORT,
    )


# =========================================================
# FORMAT MENU
# =========================================================

def format_keyboard():
    keyboard = [
        [
            InlineKeyboardButton(
                "📊 PPTX",
                callback_data="format:pptx",
            ),
            InlineKeyboardButton(
                "📄 PDF",
                callback_data="format:pdf",
            ),
        ],
        [
            InlineKeyboardButton(
                "📊📄 PPTX + PDF",
                callback_data="format:both",
            ),
        ],
    ]

    return InlineKeyboardMarkup(keyboard)


async def show_format_menu(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    text = (
        "🎨 Prezentatsiya formatini tanlang:\n\n"
        "📊 PPTX — tahrirlash mumkin\n"
        "📄 PDF — tayyor hujjat\n"
        "📊📄 PPTX + PDF — ikkalasi ham"
    )

    if update.message:
        await update.message.reply_text(
            text,
            reply_markup=format_keyboard(),
        )

    elif update.callback_query:
        await update.callback_query.edit_message_text(
            text,
            reply_markup=format_keyboard(),
        )


# =========================================================
# START
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    context.user_data["format"] = None

    text = (
        "👋 Assalomu alaykum!\n\n"
        "🤖 Men AI yordamida professional "
        "prezentatsiyalar yarataman.\n\n"
        "Avval formatni tanlang:"
    )

    await update.message.reply_text(
        text,
        reply_markup=format_keyboard(),
    )


# =========================================================
# FORMAT CALLBACK
# =========================================================

async def format_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    query = update.callback_query

    await query.answer()

    value = query.data.split(":", 1)[1]

    context.user_data["format"] = value

    names = {
        "pptx": "📊 PPTX",
        "pdf": "📄 PDF",
        "both": "📊📄 PPTX + PDF",
    }

    selected = names.get(
        value,
        "📊 PPTX",
    )

    await query.edit_message_text(
        f"✅ Tanlandi: {selected}\n\n"
        "Endi menga mavzuni yozing.\n\n"
        "Masalan:\n"
        "«Sun'iy intellekt haqida "
        "8 slaydlik o'zbekcha prezentatsiya»"
    )


# =========================================================
# COMMANDS
# =========================================================

async def pptx_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    context.user_data["format"] = "pptx"

    await update.message.reply_text(
        "📊 PPTX rejimi tanlandi.\n\n"
        "Mavzuni yozing."
    )


async def pdf_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    context.user_data["format"] = "pdf"

    await update.message.reply_text(
        "📄 PDF rejimi tanlandi.\n\n"
        "Mavzuni yozing."
    )


async def both_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    context.user_data["format"] = "both"

    await update.message.reply_text(
        "📊📄 PPTX + PDF rejimi tanlandi.\n\n"
        "Mavzuni yozing."
    )


async def help_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    text = (
        "📖 BUYRUQLAR\n\n"
        "/start — boshlash\n"
        "/pptx — PPTX yaratish\n"
        "/pdf — PDF yaratish\n"
        "/both — PPTX + PDF\n"
        "/help — yordam\n\n"
        "Yoki /start bosib formatni "
        "tugma orqali tanlang."
    )

    await update.message.reply_text(text)


# =========================================================
# FILE NAME
# =========================================================

def safe_filename(title: str) -> str:
    title = re.sub(
        r'[\\/:*?"<>|]+',
        "",
        title,
    )

    title = re.sub(
        r"\s+",
        "_",
        title.strip(),
    )

    if not title:
        title = "presentation"

    return title[:80]


# =========================================================
# GENERATE PRESENTATION
# =========================================================

async def generate_presentation(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
):
    if not update.message:
        return

    request_text = update.message.text.strip()

    if not request_text:
        return

    if len(request_text) > MAX_REQUEST_LENGTH:
        await update.message.reply_text(
            f"❌ So'rov juda uzun.\n"
            f"Maksimum {MAX_REQUEST_LENGTH} belgi."
        )
        return

    # -----------------------------------------------------
    # FORMAT
    # -----------------------------------------------------

    output_format = context.user_data.get(
        "format"
    )

    if output_format not in (
        "pptx",
        "pdf",
        "both",
    ):
        await update.message.reply_text(
            "Avval formatni tanlang:",
            reply_markup=format_keyboard(),
        )
        return

    # -----------------------------------------------------
    # COOLDOWN
    # -----------------------------------------------------

    user_id = update.effective_user.id

    last_requests = context.application.bot_data.setdefault(
        "last_requests",
        {},
    )

    now = time.time()
    last_time = last_requests.get(
        user_id,
        0,
    )

    if now - last_time < COOLDOWN_SECONDS:

        wait = int(
            COOLDOWN_SECONDS
            - (now - last_time)
        ) + 1

        await update.message.reply_text(
            f"⏳ {wait} soniya kuting."
        )

        return

    last_requests[user_id] = now

    # -----------------------------------------------------
    # CONCURRENCY
    # -----------------------------------------------------

    semaphore = (
        context.application
        .bot_data["generation_semaphore"]
    )

    async with semaphore:

        status = await update.message.reply_text(
            "🧠 AI mavzuni tahlil qilmoqda..."
        )

        try:

            # =============================================
            # AI
            # =============================================

            presentation_data = await asyncio.to_thread(
                ask_gemini,
                request_text,
            )

            model = presentation_data.get(
                "_model",
                "Gemini",
            )

            await status.edit_text(
                "🎨 Professional dizayn "
                "tayyorlanmoqda...\n\n"
                f"AI: {model}"
            )

            # =============================================
            # FILES
            # =============================================

            with tempfile.TemporaryDirectory() as temp_dir:

                title = presentation_data.get(
                    "title",
                    "Presentation",
                )

                filename = safe_filename(title)

                pptx_path = os.path.join(
                    temp_dir,
                    f"{filename}.pptx",
                )

                pdf_path = os.path.join(
                    temp_dir,
                    f"{filename}.pdf",
                )

                # -----------------------------------------
                # PPTX
                # -----------------------------------------

                if output_format in (
                    "pptx",
                    "both",
                ):

                    await asyncio.to_thread(
                        build_pptx,
                        presentation_data,
                        pptx_path,
                    )

                # -----------------------------------------
                # PDF
                # -----------------------------------------

                if output_format in (
                    "pdf",
                    "both",
                ):

                    await asyncio.to_thread(
                        build_pdf,
                        presentation_data,
                        pdf_path,
                    )

                # =========================================
                # SEND
                # =========================================

                await status.edit_text(
                    "✅ Tayyor!\n"
                    "Fayl yuborilmoqda..."
                )

                # -----------------------------------------
                # PPTX
                # -----------------------------------------

                if output_format in (
                    "pptx",
                    "both",
                ):

                    with open(
                        pptx_path,
                        "rb",
                    ) as document:

                        await update.message.reply_document(
                            document=document,
                            filename=f"{filename}.pptx",
                            caption=(
                                f"📊 {title}\n\n"
                                "PowerPoint presentation"
                            ),
                        )

                # -----------------------------------------
                # PDF
                # -----------------------------------------

                if output_format in (
                    "pdf",
                    "both",
                ):

                    with open(
                        pdf_path,
                        "rb",
                    ) as document:

                        await update.message.reply_document(
                            document=document,
                            filename=f"{filename}.pdf",
                            caption=(
                                f"📄 {title}\n\n"
                                "PDF presentation"
                            ),
                        )

                await status.delete()

                # -----------------------------------------
                # READY FOR NEXT REQUEST
                # -----------------------------------------

                context.user_data["format"] = None

        except Exception as exc:

            logger.exception(
                "Presentation generation failed: %s",
                exc,
            )

            await status.edit_text(
                "❌ Prezentatsiya yaratishda "
                "xatolik yuz berdi.\n\n"
                "Birozdan keyin yana urinib ko'ring."
            )


# =========================================================
# ERROR
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

    if not os.getenv(
        "GEMINI_API_KEY",
        "",
    ).strip():

        raise RuntimeError(
            "GEMINI_API_KEY environment variable is missing."
        )

    start_health_server()

    application = (
        Application.builder()
        .token(TOKEN)
        .build()
    )

    application.bot_data[
        "generation_semaphore"
    ] = asyncio.Semaphore(2)

    application.bot_data[
        "last_requests"
    ] = {}

    # Commands
    application.add_handler(
        CommandHandler(
            "start",
            start,
        )
    )

    application.add_handler(
        CommandHandler(
            "pptx",
            pptx_command,
        )
    )

    application.add_handler(
        CommandHandler(
            "pdf",
            pdf_command,
        )
    )

    application.add_handler(
        CommandHandler(
            "both",
            both_command,
        )
    )

    application.add_handler(
        CommandHandler(
            "help",
            help_command,
        )
    )

    # Format buttons
    application.add_handler(
        CallbackQueryHandler(
            format_callback,
            pattern=r"^format:",
        )
    )

    # User text
    application.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            generate_presentation,
        )
    )

    application.add_error_handler(
        error_handler
    )

    logger.info(
        "SLAYD BOT starting..."
    )

    application.run_polling(
        drop_pending_updates=True
    )


if __name__ == "__main__":
    main()
