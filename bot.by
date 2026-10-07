import asyncio
import traceback
import logging
import os
import time
from aiohttp import web, ClientSession
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters
)

from config import TELEGRAM_TOKEN, PORT, logger
from gemini_service import generate_presentation_structure
from pptx_generator import create_pptx
from pdf_generator import create_pdf

USER_MODES = {}
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

# Statistika va foydalanuvchilar ro'yxatini xotirada saqlash
USERS_SET = set()
STATS = {"total_generated": 0}
USER_LAST_REQUEST = {}

# Render Anti-Sleep uchun URL
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL")


async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    USERS_SET.add(user_id)

    welcome_text = (
        "👋 **Xush kelibsiz! Professional Slayd Botiga xush kelibsiz.**\n\n"
        "Menga istalgan mavzuni yuboring, men sizga mos ravishda taqdimot tayyorlab beraman.\n\n"
        "**Formatni tanlash komandalari:**\n"
        "/pptx - Faqat PPTX formatida\n"
        "/pdf - Faqat PDF formatida\n"
        "/both - PPTX va PDF formatlarida (Standart)\n"
        "/help - Yordam"
    )
    await update.message.reply_text(welcome_text, parse_mode="Markdown")


async def set_mode(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cmd = update.message.text.replace("/", "").lower().strip()
    user_id = update.effective_user.id
    if cmd in ["pptx", "pdf", "both"]:
        USER_MODES[user_id] = cmd
        await update.message.reply_text(
            f"✅ Format belgilandi: **{cmd.upper()}**", parse_mode="Markdown"
        )


async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        return

    stats_text = (
        "📊 **BOT STATISTIKASI**\n\n"
        f"👥 **Jami foydalanuvchilar:** `{len(USERS_SET)}` ta\n"
        f"📊 **Yaratilgan taqdimotlar:** `{STATS['total_generated']}` ta"
    )
    await update.message.reply_text(stats_text, parse_mode="Markdown")


async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        return

    msg_text = update.message.text.replace("/broadcast", "").strip()
    if not msg_text:
        await update.message.reply_text(
            "⚠️ Xabar matnini kiriting! Masalan: `/broadcast Salom barchaga!`",
            parse_mode="Markdown",
        )
        return

    count = 0
    for uid in USERS_SET:
        try:
            await context.bot.send_message(
                chat_id=uid, text=msg_text, parse_mode="Markdown"
            )
            count += 1
            await asyncio.sleep(0.05)
        except Exception:
            pass

    await update.message.reply_text(
        f"✅ Xabar `{count}` ta foydalanuvchiga yuborildi.", parse_mode="Markdown"
    )


async def handle_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    USERS_SET.add(user_id)
    prompt = update.message.text

    # Rate Limiting (RAM xavfsizligi uchun 10 soniyalik cheklov)
    now = time.time()
    last_req = USER_LAST_REQUEST.get(user_id, 0)
    if now - last_req < 10:
        await update.message.reply_text(
            "⏳ Iltimos, keyingi so'rovdan oldin 10 soniya kuting."
        )
        return
    USER_LAST_REQUEST[user_id] = now

    fmt = USER_MODES.get(user_id, "both")
    status_msg = await update.message.reply_text(
        "🧠 AI mavzuni tahlil qilmoqda va slaydlarni tuzmoqda..."
    )

    try:
        slide_data = await generate_presentation_structure(prompt)
        await status_msg.edit_text("🎨 Prezentatsiya fayllari shakllantirilmoqda...")

        clean_title = slide_data.get("title", "Prezentatsiya").replace(" ", "_")[
            :30
        ]

        if fmt in ["pptx", "both"]:
            pptx_stream = create_pptx(slide_data)
            await update.message.reply_document(
                document=pptx_stream,
                filename=f"{clean_title}.pptx",
                caption="📊 PPTX Formati",
            )

        if fmt in ["pdf", "both"]:
            pdf_stream = create_pdf(slide_data)
            await update.message.reply_document(
                document=pdf_stream,
                filename=f"{clean_title}.pdf",
                caption="📄 PDF Formati",
            )

        STATS["total_generated"] += 1
        await status_msg.delete()

    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Slayd yaratishda xatolik:\n{error_details}")

        await status_msg.edit_text(
            "❌ Slayd tayyorlashda xatolik yuz berdi. Qaytadan urinib ko'ring."
        )

        # Adminga diagnostika xabarini yuborish
        if ADMIN_ID != 0:
            try:
                admin_msg = (
                    f"⚠️ **BOTDA XATOLIK YUZ BERDI!**\n\n"
                    f"👤 User ID: `{user_id}`\n"
                    f"📝 Prompt: `{prompt}`\n\n"
                    f"🔍 **Aniq xatolik sababi:**\n```text\n{error_details[:3500]}\n```"
                )
                await context.bot.send_message(
                    chat_id=ADMIN_ID, text=admin_msg, parse_mode="Markdown"
                )
            except Exception as admin_err:
                logger.error(f"Adminga xatoni yuborishda muammo: {admin_err}")


# Health-check HTTP Server (Render port binding uchun)
async def health_check(request):
    return web.Response(text="Bot is active and running 24/7!", status=200)


# Keep-Alive Self Ping Task (Render Anti-Sleep)
async def keep_alive_task():
    await asyncio.sleep(10)
    url = RENDER_EXTERNAL_URL or f"http://127.0.0.1:{PORT}"
    logger.info(f"Keep-Alive tayyorlandi URL: {url}")

    async with ClientSession() as session:
        while True:
            try:
                async with session.get(url) as resp:
                    logger.info(f"Self-ping yuborildi: status {resp.status}")
            except Exception as e:
                logger.warning(f"Self-ping xatosi: {e}")
            await asyncio.sleep(600)  # Har 10 minutda ping yuboradi


async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", start_command))
    app.add_handler(CommandHandler("stats", stats_command))
    app.add_handler(CommandHandler("broadcast", broadcast_command))
    app.add_handler(CommandHandler(["pptx", "pdf", "both"], set_mode))
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, handle_prompt)
    )

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    # Webserver setup
    server = web.Application()
    server.router.add_get("/", health_check)
    runner = web.AppRunner(server)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()

    # Background Anti-Sleep task
    asyncio.create_task(keep_alive_task())

    logger.info(
        f"Bot va Health Server {PORT}-portda muvaffaqiyatli ishga tushdi."
    )
    await asyncio.Event().wait()


if __name__ == "__main__":
    asyncio.run(main())
