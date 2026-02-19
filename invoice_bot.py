import asyncio
import logging
import os
import tempfile
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.error import NetworkError, RetryAfter, TimedOut
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

from generate_invoice_from_docx import generate_invoice_pdf


ASK_TYPE, ASK_BUYER, ASK_BASIS, ASK_ITEM_NAME, ASK_ITEM_PRICE = range(5)
TYPE_OOO = "🏢 ООО"
TYPE_IP = "👤 ИП"
BACK_TEXT = "⬅️ Назад"
logger = logging.getLogger(__name__)


async def safe_reply_text(message, text: str, retries: int = 3, **kwargs) -> bool:
    for attempt in range(1, retries + 1):
        try:
            await message.reply_text(text, **kwargs)
            return True
        except RetryAfter as exc:
            await asyncio.sleep(float(exc.retry_after))
        except (TimedOut, NetworkError):
            if attempt == retries:
                return False
            await asyncio.sleep(1.5 * attempt)
    return False


async def safe_reply_document(message, document_path: Path, retries: int = 3) -> bool:
    for attempt in range(1, retries + 1):
        try:
            with document_path.open("rb") as file:
                await message.reply_document(document=file)
            return True
        except RetryAfter as exc:
            await asyncio.sleep(float(exc.retry_after))
        except (TimedOut, NetworkError):
            if attempt == retries:
                return False
            await asyncio.sleep(1.5 * attempt)
    return False


def back_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton(BACK_TEXT, callback_data="back")]]
    )


async def send_with_back(message, text: str, **kwargs) -> bool:
    return await safe_reply_text(message, text, reply_markup=back_keyboard(), **kwargs)


async def send_type_picker(message) -> bool:
    keyboard = [
        [
            InlineKeyboardButton(TYPE_OOO, callback_data="type_ooo"),
            InlineKeyboardButton(TYPE_IP, callback_data="type_ip"),
        ]
    ]
    sent = await safe_reply_text(
        message,
        "📋 Выберите тип счета:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return sent


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not update.message:
        return ConversationHandler.END
    sent = await send_type_picker(update.message)
    if not sent:
        return ConversationHandler.END
    return ASK_TYPE


async def start_from_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()
    sent = await send_type_picker(query.message)
    if not sent:
        return ConversationHandler.END
    return ASK_TYPE


async def ask_buyer(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ASK_TYPE
    await query.answer()
    selected_data = query.data or ""
    if selected_data not in ("type_ooo", "type_ip"):
        return ASK_TYPE
    context.user_data["invoice_type_label"] = TYPE_IP if selected_data == "type_ip" else TYPE_OOO
    context.user_data["invoice_type"] = "ip" if selected_data == "type_ip" else "ooo"

    sent = await send_with_back(
        query.message,
        'Введите данные покупателя.\n\n'
        'Пример:\n'
        '`ООО "Название", ИНН 1234567890, КПП 123456789, '
        '123456, г. Москва, ул. Примерная, д. 1, стр. 2, пом. 3`',
        parse_mode="Markdown",
    )
    if not sent:
        return ASK_TYPE
    return ASK_BUYER


# --- Back handlers for each step ---

async def back_to_type(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query:
        await query.answer()
        await send_type_picker(query.message)
    return ASK_TYPE


async def back_to_buyer(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query:
        await query.answer()
        await send_with_back(
            query.message,
            'Введите данные покупателя.\n\n'
            'Пример:\n'
            '`ООО "Название", ИНН 1234567890, КПП 123456789, '
            '123456, г. Москва, ул. Примерная, д. 1, стр. 2, пом. 3`',
            parse_mode="Markdown",
        )
    return ASK_BUYER


async def back_to_basis(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query:
        await query.answer()
        await send_with_back(query.message, "Введите основание договора:")
    return ASK_BASIS


async def back_to_item_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query:
        await query.answer()
        await send_with_back(query.message, "Введите название товара/услуги:")
    return ASK_ITEM_NAME


# --- Step handlers ---

async def ask_basis(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["buyer"] = update.message.text.strip()
    sent = await send_with_back(update.message, "Введите основание договора:")
    if not sent:
        return ASK_BUYER
    return ASK_BASIS


async def ask_item_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["contract_basis"] = update.message.text.strip()
    sent = await send_with_back(update.message, "Введите название товара/услуги:")
    if not sent:
        return ASK_BASIS
    return ASK_ITEM_NAME


async def ask_item_price(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["item_name"] = update.message.text.strip()
    sent = await send_with_back(update.message, "Введите цену (например 50 000,00):")
    if not sent:
        return ASK_ITEM_NAME
    return ASK_ITEM_PRICE


async def generate(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data["item_price"] = update.message.text.strip()

    buyer = context.user_data.get("buyer", "")
    contract_basis = context.user_data.get("contract_basis", "")
    item_name = context.user_data.get("item_name", "")
    item_price = context.user_data.get("item_price", "")
    invoice_type = context.user_data.get("invoice_type", "ooo")

    template_path = Path("invoice_ip.docx" if invoice_type == "ip" else "invoice.docx")
    if not template_path.exists():
        await safe_reply_text(update.message, f"Не найден шаблон {template_path.name}")
        return ConversationHandler.END

    await safe_reply_text(update.message, "⏳ Счет создается, подождите пару секунд...")

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_docx = Path(tmp_dir) / "invoice_filled.docx"
        out_pdf = Path(tmp_dir) / "invoice.pdf"
        try:
            generate_invoice_pdf(
                template_path=template_path,
                output_docx=out_docx,
                output_pdf=out_pdf,
                buyer=buyer,
                contract_basis=contract_basis,
                item_name=item_name,
                item_price=item_price,
                invoice_type=invoice_type,
            )
        except Exception as exc:
            await safe_reply_text(update.message, f"Ошибка генерации: {exc}")
            return ConversationHandler.END

        sent = await safe_reply_document(update.message, out_pdf)
        if not sent:
            await safe_reply_text(
                update.message,
                "Не удалось отправить PDF из-за проблем с сетью. Попробуй еще раз.",
            )
            return ASK_ITEM_PRICE

    await safe_reply_text(
        update.message,
        "✅ Готово! Хотите создать еще один счет?",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("🔄 Создать еще", callback_data="create_more")]]
        ),
    )

    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if update.message:
        await safe_reply_text(update.message, "❌ Отменено.")
    elif update.callback_query:
        await update.callback_query.answer()
        await safe_reply_text(update.callback_query.message, "❌ Отменено.")
    return ConversationHandler.END


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.exception("Unhandled bot error", exc_info=context.error)


def main() -> None:
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("Укажи TELEGRAM_BOT_TOKEN в переменных окружения.")

    app = (
        ApplicationBuilder()
        .token(token)
        .connect_timeout(30.0)
        .read_timeout(30.0)
        .write_timeout(30.0)
        .pool_timeout(30.0)
        .build()
    )

    import warnings
    warnings.filterwarnings("ignore", message=".*per_message.*", category=UserWarning)

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CallbackQueryHandler(start_from_callback, pattern=r"^create_more$"),
        ],
        states={
            ASK_TYPE: [
                CallbackQueryHandler(ask_buyer, pattern=r"^type_(ooo|ip)$"),
            ],
            ASK_BUYER: [
                CallbackQueryHandler(back_to_type, pattern=r"^back$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ask_basis),
            ],
            ASK_BASIS: [
                CallbackQueryHandler(back_to_buyer, pattern=r"^back$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ask_item_name),
            ],
            ASK_ITEM_NAME: [
                CallbackQueryHandler(back_to_basis, pattern=r"^back$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, ask_item_price),
            ],
            ASK_ITEM_PRICE: [
                CallbackQueryHandler(back_to_item_name, pattern=r"^back$"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, generate),
            ],
        },
        fallbacks=[
            CommandHandler("start", start),
            CommandHandler("cancel", cancel),
        ],
        per_message=False,
    )

    app.add_handler(conv)

    # Catch clicks on old/stale inline buttons that are outside the conversation
    async def stale_button(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        query = update.callback_query
        if query:
            await query.answer("Эта кнопка уже неактивна. Нажмите /start", show_alert=True)

    app.add_handler(CallbackQueryHandler(stale_button))

    app.add_error_handler(error_handler)
    app.run_polling()


if __name__ == "__main__":
    main()
