import os
import asyncio
import logging
import json
from datetime import datetime, timedelta
import pytz
import gspread
from google.oauth2.service_account import Credentials
from telegram import Bot
from telegram.error import TelegramError

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

TELEGRAM_TOKEN   = os.environ["TELEGRAM_TOKEN"]
SPREADSHEET_ID   = os.environ["SPREADSHEET_ID"].strip()
SHEET_NAME       = os.environ.get("SHEET_NAME", "Лист1")
TIMEZONE         = os.environ.get("TIMEZONE", "Europe/Kiev")
CHECK_INTERVAL   = int(os.environ.get("CHECK_INTERVAL", "60"))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

COL_DATE   = 0
COL_TIME   = 1
COL_CHAT   = 2
COL_TEXT   = 3
COL_MEDIA  = 4
COL_STATUS = 5


def get_sheet():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    logger.info(f"SPREADSHEET_ID: '{SPREADSHEET_ID}'")
    logger.info(f"SHEET_NAME: '{SHEET_NAME}'")
    if creds_json:
        creds_json = creds_json.strip()
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def parse_datetime(date_str: str, time_str: str, tz) -> datetime | None:
    date_str = date_str.strip()
    time_str = time_str.strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            d = datetime.strptime(date_str, fmt)
            t = datetime.strptime(time_str, "%H:%M")
            return tz.localize(d.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0))
        except ValueError:
            continue
    return None


async def send_post(bot: Bot, chat_id: str, text: str, media_url: str | None):
    chat_id = chat_id.strip()
    if media_url and media_url.strip():
        await bot.send_photo(chat_id=chat_id, photo=media_url.strip(), caption=text)
    else:
        await bot.send_message(chat_id=chat_id, text=text, parse_mode="HTML")


async def check_and_send(bot: Bot):
    tz = pytz.timezone(TIMEZONE)
    now = datetime.now(tz)
    window_start = now - timedelta(minutes=2)

    try:
        sheet = get_sheet()
        rows = sheet.get_all_values()
        logger.info(f"Таблица прочитана, строк: {len(rows)}")
    except Exception as e:
        logger.error(f"Ошибка чтения таблицы: {e}")
        return

    if not rows:
        return

    start_row = 1 if rows and rows[0][COL_DATE].lower() in ("дата", "date", "a") else 0

    for i, row in enumerate(rows[start_row:], start=start_row + 2):
        if len(row) < 4:
            continue
        status = row[COL_STATUS].strip() if len(row) > COL_STATUS else ""
        if status.startswith("✅"):
            continue

        post_dt = parse_datetime(row[COL_DATE], row[COL_TIME], tz)
        if post_dt is None:
            continue

        if window_start <= post_dt <= now:
            chat_id  = row[COL_CHAT]
            text     = row[COL_TEXT]
            media    = row[COL_MEDIA] if len(row) > COL_MEDIA else ""

            logger.info(f"Строка {i}: отправляю в {chat_id}")
            try:
                await send_post(bot, chat_id, text, media)
                new_status = f"✅ Отправлено {now.strftime('%d.%m %H:%M')}"
                logger.info(f"Строка {i}: успешно отправлено")
            except TelegramError as e:
                new_status = f"❌ Ошибка: {e}"
                logger.error(f"Строка {i}: ошибка Telegram — {e}")
            except Exception as e:
                new_status = f"❌ Ошибка: {e}"
                logger.error(f"Строка {i}: неизвестная ошибка — {e}")

            try:
                sheet.update_cell(i, COL_STATUS + 1, new_status)
            except Exception as e:
                logger.error(f"Не удалось записать статус: {e}")


async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    me = await bot.get_me()
    logger.info(f"Бот запущен: @{me.username}")
    logger.info(f"Проверка каждые {CHECK_INTERVAL} секунд.")

    while True:
        try:
            await check_and_send(bot)
        except Exception as e:
            logger.error(f"Ошибка в главном цикле: {e}")
        await asyncio.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    asyncio.run(main())
