import logging
import os
import uuid
from pathlib import Path

from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from docx_handler import fill_template
from gemini_client import generate_kp_content

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TEMPLATE_PATH = os.getenv(
    "TEMPLATE_PATH", "КП_№SP26-10_от_18_02_26г (1) (1) (1).docx"
)
OUTPUT_DIR = Path("temp_output")
OUTPUT_DIR.mkdir(exist_ok=True)

# Per-user state keys stored in context.user_data:
#   "history"  — list of Gemini dialog turns (for context-aware edits)
#   "content"  — last generated KP content dict
#   "waiting_for_request" — True once the user sends /start


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a greeting and usage instructions."""
    context.user_data.clear()
    context.user_data["history"] = []
    await update.message.reply_text(
        "👋 Привет! Я помогу составить коммерческое предложение.\n\n"
        "Напиши мне запрос в свободной форме, например:\n"
        "📝 *Составь КП для компании «Рога и Копыта», услуга: разработка сайта, "
        "сумма: 500 000 тенге*\n\n"
        "После отправки готового файла ты сможешь попросить правки — "
        "просто напиши что именно изменить.",
        parse_mode="Markdown",
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle all user text messages — both initial requests and edit instructions."""
    user_text = update.message.text.strip()
    user_id = update.effective_user.id

    history: list = context.user_data.get("history", [])
    previous_content: dict | None = context.user_data.get("content")

    await update.message.reply_text("⏳ Генерирую КП, подождите...")

    # Build the effective prompt.
    # If there is a previous content, instruct Gemini to apply the edits.
    if previous_content:
        prompt = (
            f"Предыдущий вариант КП был таким:\n{previous_content}\n\n"
            f"Пользователь просит внести следующие правки:\n{user_text}\n\n"
            "Верни ПОЛНЫЙ обновлённый JSON с учётом правок."
        )
    else:
        prompt = user_text

    try:
        content = generate_kp_content(prompt, history)
    except Exception as exc:
        logger.exception("Error generating KP for user %s", user_id)
        await update.message.reply_text(
            f"❌ Ошибка при генерации КП: {exc}\n\nПопробуй ещё раз."
        )
        return

    # Update dialog history so Gemini has full context on next edit
    history.append({"role": "user", "parts": [user_text]})
    history.append({"role": "model", "parts": [str(content)]})
    context.user_data["history"] = history
    context.user_data["content"] = content

    # Fill the template
    output_filename = OUTPUT_DIR / f"КП_{user_id}_{uuid.uuid4().hex[:8]}.docx"
    try:
        fill_template(TEMPLATE_PATH, content, str(output_filename))
    except Exception as exc:
        logger.exception("Error filling template for user %s", user_id)
        await update.message.reply_text(
            f"❌ Ошибка при заполнении шаблона: {exc}\n\nПроверь шаблон и попробуй ещё раз."
        )
        return

    # Send the file
    try:
        caption = (
            f"✅ *КП готово!*\n"
            f"Компания: {content.get('company_name', '—')}\n"
            f"Услуга: {content.get('service_title', '—')}\n"
            f"Сумма: {content.get('total_amount', '—')}\n\n"
            "Если нужны правки — просто напиши что изменить 😊"
        )
        with open(output_filename, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"КП_{content.get('company_name', 'draft')}.docx",
                caption=caption,
                parse_mode="Markdown",
            )
    finally:
        # Clean up temp file
        try:
            output_filename.unlink()
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not set in environment / .env file")

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("Bot started. Polling...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
