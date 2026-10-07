import asyncio
import traceback
import logging
from aiohttp import web
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

# Admin Telegram ID (O'zingizning Telegram ID'ingizni kiritishingiz mumkin)
# Masalan: ADMIN_ID = 123456789
import os
ADMIN_ID = int(os.getenv("ADMIN_ID", "0"))

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
        await update.message.reply_text(f"✅ Format belgilandi: **{cmd.upper()}**", parse_mode="Markdown")

async def handle_prompt(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    prompt = update.message.text
    fmt = USER_MODES.get(user_id, "both")

    status_msg = await update.message.reply_text("🧠 AI mavzuni tahlil qilmoqda va slaydlarni tuzmoqda...")

    try:
        slide_data = await generate_presentation_structure(prompt)
        await status_msg.edit_text("🎨 Prezentatsiya fayllari shakllantirilmoqda...")

        clean_title = slide_data.get("title", "Prezentatsiya").replace(" ", "_")[:30]

        if fmt in ["pptx", "both"]:
            pptx_stream = create_pptx(slide_data)
            await update.message.reply_document(
                document=pptx_stream,
                filename=f"{clean_title}.pptx",
                caption="📊 PPTX Formati"
            )

        if fmt in ["pdf", "both"]:
            pdf_stream = create_pdf(slide_data)
            await update.message.reply_document(
                document=pdf_stream,
                filename=f"{clean_title}.pdf",
                caption="📄 PDF Formati"
            )

        await status_msg.delete()

    except Exception as e:
        error_details = traceback.format_exc()
        logger.error(f"Slayd yaratishda xatolik:\n{error_details}")

        # Foydalanuvchiga oddiy xabar
        await status_msg.edit_text("❌ Slayd tayyorlashda xatolik yuz berdi. Qaytadan urinib ko'ring.")

        # Adminga aniq xatolik sababini yuborish
        if ADMIN_ID != 0:
            try:
                admin_msg = (
                    f"⚠️ **BOTDA XATOLIK YUZ BERDI!**\n\n"
                    f"👤 User ID: `{user_id}`\n"
                    f"📝 Prompt: `{prompt}`\n\n"
                    f"🔍 **Aniq xatolik sababi:**\n```text\n{error_details[:3500]}\n```"
                )
                await context.bot.send_message(chat_id=ADMIN_ID, text=admin_msg, parse_mode="Markdown")
            except Exception as admin_err:
                logger.error(f"Adminga xatolik xabarini yuborishda muammo: {admin_err}")

async def health_check(request):
    return web.Response(text="Bot runs OK", status=200)

async def main():
    logger.info("Bot ishga tushirish tayyorgarligi boshlandi...")
    
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", start_command))
    app.add_handler(CommandHandler(["pptx", "pdf", "both"], set_mode))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_prompt))

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    # Webserver setup (Render Port Binding va Health check uchun)
    server = web.Application()
    server.router.add_get("/", health_check)
    runner = web.AppRunner(server)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()

    logger.info(f"✅ Bot va Webserver {PORT}-portda muvaffaqiyatli ishga tushdi.")
    await asyncio.Event().wait()

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except Exception as startup_err:
        logger.critical(f"❌ BOT ISHGA TUSHISHIDA CRITICAL XATOLIK YUZ BERDI:\n{traceback.format_exc()}")
        raise startup_err
