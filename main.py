import asyncio
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

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    welcome_text = (
        "👋 **Xush kelibsiz! Professional Slayd Botiga hush kelibsiz.**\n\n"
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
        logger.error(f"Slayd yaratishda xatolik: {e}", exc_info=True)
        await status_msg.edit_text("❌ Slayd tayyorlashda xatolik yuz berdi. Qaytadan urinib ko'ring.")

# Render uchun Health-check endpoint
async def health_check(request):
    return web.Response(text="Bot is running smoothly!", status=200)

async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", start_command))
    app.add_handler(CommandHandler(["pptx", "pdf", "both"], set_mode))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_prompt))

    await app.initialize()
    await app.start()
    await app.updater.start_polling()

    # Webserver setup (Render Port Binding uchun)
    server = web.Application()
    server.router.add_get("/", health_check)
    runner = web.AppRunner(server)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()

    logger.info(f"Bot va Health Server {PORT}-portda muvaffaqiyatli ishga tushdi.")
    await asyncio.Event().wait()

if __name__ == "__main__":
    asyncio.run(main())
