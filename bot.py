import os
import asyncio
import logging
import json
import re
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
import pytz
import gspread
from google.oauth2.service_account import Credentials
from telegram import Bot, Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes
from telegram.error import TelegramError

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
SPREADSHEET_ID  = os.environ["SPREADSHEET_ID"].strip()
SHEET_NAME      = os.environ.get("SHEET_NAME", "Лист1")
TIMEZONE        = os.environ.get("TIMEZONE", "Europe/Kiev")
CHECK_INTERVAL  = int(os.environ.get("CHECK_INTERVAL", "60"))

# Автоответ на личные сообщения
AUTO_REPLY_TEXT = (
    "⚠️ Важно!\n\n"
    "Данный менеджер не осуществляет поддержку пользователей и не занимается решением технических или организационных вопросов.\n\n"
    "Для получения помощи, пожалуйста, обращайтесь в Help Desk SlideEdu."
)
AUTO_REPLY_PHOTO = "https://drive.google.com/uc?export=download&id=1ME_JpmIRw95EldEXNVNNWvtn17vEb92X"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

COL_DATE      = 0
COL_TIME      = 1
COL_CHAT_RU   = 2
COL_CHAT_UKR  = 3
COL_TEXT      = 4
COL_MEDIA_RU  = 5
COL_MEDIA_UKR = 6
COL_STATUS    = 7


def get_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON не задан!")
    creds_info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def translate_to_ukrainian(text: str) -> str:
    try:
        encoded = urllib.parse.quote(text)
        url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=ru&tl=uk&dt=t&q={encoded}"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=10) as response:
            result = json.loads(response.read().decode('utf-8'))
        translated = ''
        for part in result[0]:
            if part[0]:
                translated += part[0]
        logger.info(f"Перевод успешен, длина: {len(translated)}")
        return translated
    except Exception as e:
        logger.error(f"Ошибка перевода: {e}")
        return text


def excel_date_to_datetime(excel_date_float: float, tz) -> datetime | None:
    try:
        base = datetime(1899, 12, 30)
        days = int(excel_date_float)
        time_fraction = excel_date_float - days
        total_seconds = round(time_fraction * 86400)
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        dt = base + timedelta(days=days, hours=hours, minutes=minutes, seconds=seconds)
        return tz.localize(dt)
    except Exception:
        return None


def parse_datetime(date_val: str, time_val: str, tz) -> datetime | None:
    date_val = date_val.strip()
    time_val = time_val.strip()
    try:
        excel_num = float(date_val)
        if time_val:
            try:
                time_num = float(time_val)
                combined = int(excel_num) + time_num
                return excel_date_to_datetime(combined, tz)
            except ValueError:
                dt = excel_date_to_datetime(excel_num, tz)
                if dt:
                    t = datetime.strptime(time_val, "%H:%M")
                    return dt.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0)
        else:
            return excel_date_to_datetime(excel_num, tz)
    except ValueError:
        pass
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            d = datetime.strptime(date_val, fmt)
            if time_val:
                t = datetime.strptime(time_val, "%H:%M")
                d = d.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0)
            return tz.localize(d)
        except ValueError:
            continue
    return None


def convert_drive_url(url: str) -> str:
    url = url.strip()
    match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url


async def send_post(bot: Bot, chat_id: str, text: str, media_url: str):
    chat_id = chat_id.strip()
    if media_url and media_url.strip():
        url = convert_drive_url(media_url)
        await bot.send_photo(chat_id=chat_id, photo=url, caption=text)
    else:
        await bot.send_message(chat_id=chat_id, text=text)


async def handle_private_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Автоответ на личные сообщения."""
    if update.message and update.message.chat.type == "private":
        user = update.message.from_user
        logger.info(f"Личное сообщение от {user.first_name} (@{user.username}): {update.message.text}")
        try:
            await update.message.reply_photo(
                photo=AUTO_REPLY_PHOTO,
                caption=AUTO_REPLY_TEXT
            )
            logger.info(f"Автоответ отправлен пользователю {user.first_name}")
        except Exception as e:
            logger.error(f"Ошибка отправки автоответа: {e}")
            try:
                await update.message.reply_text(AUTO_REPLY_TEXT)
            except Exception as e2:
                logger.error(f"Ошибка отправки текстового автоответа: {e2}")


async def check_and_send(bot: Bot):
    tz = pytz.timezone(TIMEZONE)
    now = datetime.now(tz)
    window_start = now - timedelta(minutes=2)

    try:
        sheet = get_sheet()
        rows = sheet.get_all_values()
    except Exception as e:
        logger.error(f"Ошибка чтения таблицы: {e}")
        return

    if not rows:
        return

    start_row = 1 if rows[0][COL_DATE].lower() in ("дата", "date") else 0
    logger.info(f"Всего строк: {len(rows)}, проверяю с строки {start_row + 1}")

    for i, row in enumerate(rows[start_row:], start=start_row + 2):
        if len(row) < 5:
            continue

        status = row[COL_STATUS].strip() if len(row) > COL_STATUS else ""
        if status.startswith("✅"):
            continue

        date_val = row[COL_DATE]
        time_val = row[COL_TIME] if len(row) > COL_TIME else ""

        if not date_val.strip():
            continue

        post_dt = parse_datetime(date_val, time_val, tz)
        if post_dt is None:
            continue

        logger.info(f"Строка {i}: дата поста {post_dt.strftime('%d.%m.%Y %H:%M')}, сейчас {now.strftime('%d.%m.%Y %H:%M')}")

        if window_start <= post_dt <= now:
            chat_ru   = row[COL_CHAT_RU].strip() if len(row) > COL_CHAT_RU else ""
            chat_ukr  = row[COL_CHAT_UKR].strip() if len(row) > COL_CHAT_UKR else ""
            text      = row[COL_TEXT]
            media_ru  = row[COL_MEDIA_RU].strip() if len(row) > COL_MEDIA_RU else ""
            media_ukr = row[COL_MEDIA_UKR].strip() if len(row) > COL_MEDIA_UKR else ""

            statuses = []

            if chat_ru:
                try:
                    await send_post(bot, chat_ru, text, media_ru)
                    statuses.append("✅ РУ")
                    logger.info(f"Строка {i}: РУ отправлено!")
                except TelegramError as e:
                    statuses.append(f"❌ РУ: {e}")
                    logger.error(f"Строка {i}: РУ ошибка — {e}")

            if chat_ukr:
                text_ukr = translate_to_ukrainian(text)
                media_for_ukr = media_ukr if media_ukr else media_ru
                try:
                    await send_post(bot, chat_ukr, text_ukr, media_for_ukr)
                    statuses.append("✅ УКР")
                    logger.info(f"Строка {i}: УКР отправлено!")
                except TelegramError as e:
                    statuses.append(f"❌ УКР: {e}")
                    logger.error(f"Строка {i}: УКР ошибка — {e}")

            new_status = f"{' | '.join(statuses)} {now.strftime('%d.%m %H:%M')}"
            try:
                sheet.update_cell(i, COL_STATUS + 1, new_status)
            except Exception as e:
                logger.error(f"Не удалось записать статус: {e}")


async def scheduler(bot: Bot):
    """Фоновая задача проверки расписания."""
    while True:
        try:
            await check_and_send(bot)
        except Exception as e:
            logger.error(f"Ошибка в планировщике: {e}")
        await asyncio.sleep(CHECK_INTERVAL)


async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Обработчик личных сообщений
    app.add_handler(MessageHandler(filters.TEXT & filters.ChatType.PRIVATE, handle_private_message))

    bot = app.bot
    me = await bot.get_me()
    logger.info(f"✅ Бот запущен: @{me.username}")
    logger.info(f"Проверка каждые {CHECK_INTERVAL} секунд")
    logger.info(f"Таблица ID: {SPREADSHEET_ID}")

    # Запускаем планировщик постов параллельно
    async with app:
        await app.start()
        await app.updater.start_polling()
        await scheduler(bot)


if __name__ == "__main__":
    asyncio.run(main())
