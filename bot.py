import os
import tempfile
import logging

from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

from slides import ask_gemini, build_pptx


# ============================================================
# SETTINGS
# ============================================================

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()


# ============================================================
# LOGGING
# ============================================================

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)

logger = logging.getLogger(__name__)


# ============================================================
# /START
# ============================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):

    await update.message.reply_text(
        "👋 Assalomu alaykum!\n\n"
        "Men AI yordamida PowerPoint slayd yarataman.\n\n"
        "Masalan:\n\n"
        "📚 Sun'iy intellekt haqida 5 ta slayd\n\n"
        "🌍 Global isish haqida 7 ta slayd\n\n"
        "🧬 Biologiya: hujayra haqida 10 ta slayd\n\n"
        "Mavzuni yozing 👇"
    )


# ============================================================
# HELP
# ============================================================

async def help_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    await update.message.reply_text(
        "📖 Qanday foydalanish kerak?\n\n"
        "Shunchaki mavzu va slayd sonini yozing.\n\n"
        "Misol:\n"
        "👉 O'zbekiston tarixi haqida 8 ta slayd\n\n"
        "Agar slayd sonini yozmasangiz, 7 ta slayd yaratiladi."
    )


# ============================================================
# CREATE SLIDE
# ============================================================

async def create_presentation(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if not update.message:
        return

    request_text = update.message.text.strip()

    if not request_text:
        await update.message.reply_text(
            "❗ Mavzuni yozing."
        )
        return

    # Juda uzun so'rovni cheklaymiz
    if len(request_text) > 2000:

        await update.message.reply_text(
            "❗ So'rov juda uzun.\n"
            "Iltimos, qisqaroq yozing."
        )

        return

    # Kutish xabari
    status_message = await update.message.reply_text(
        "⏳ Taqdimotingiz tayyorlan
