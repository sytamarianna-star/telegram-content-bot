import os
import asyncio
import logging
import json
import re
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

# ── Настройки из переменных окружения Railway ─────────────────────────────────
TELEGRAM_TOKEN  = os.environ["TELEGRAM_TOKEN"]
SPREADSHEET_ID  = os.environ["SPREADSHEET_ID"].strip()
SHEET_NAME      = os.environ.get("SHEET_NAME", "Лист1")
TIMEZONE        = os.environ.get("TIMEZONE", "Europe/Kiev")
CHECK_INTERVAL  = int(os.environ.get("CHECK_INTERVAL", "60"))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Колонки таблицы (индексы с 0)
COL_DATE   = 0  # A: Дата
COL_TIME   = 1  # B: Время
COL_CHAT   = 2  # C: Chat ID
COL_TEXT   = 3  # D: Текст поста
COL_MEDIA  = 4  # E: Ссылка на фото (необязательно)
COL_STATUS = 5  # F: Статус (бот пишет сам)


def get_sheet():
    """Подключение к Google Sheets через Service Account."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON не задан!")
    creds_info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def excel_date_to_datetime(excel_date_float: float, tz) -> datetime | None:
    """Конвертирует числовую дату Excel в datetime."""
    try:
        # Excel считает дни с 30.12.1899
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
    """Парсит дату и время из ячеек — поддерживает числа Excel и текст."""
    date_val = date_val.strip()
    time_val = time_val.strip()

    # Числовой формат Excel (дата и время вместе в одной ячейке)
    try:
        excel_num = float(date_val)
        # Если время задано отдельно — берём дату из числа даты, время отдельно
        if time_val:
            try:
                time_num = float(time_val)
                # Складываем дату + время
                combined = int(excel_num) + time_num
                return excel_date_to_datetime(combined, tz)
            except ValueError:
                # Время как текст HH:MM
                dt = excel_date_to_datetime(excel_num, tz)
                if dt:
                    t = datetime.strptime(time_val, "%H:%M")
                    return dt.replace(hour=t.hour, minute=t.minute, second=0, microsecond=0)
        else:
            return excel_date_to_datetime(excel_num, tz)
    except ValueError:
        pass

    # Текстовый формат даты
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
    """
    Конвертирует ссылку Google Drive в прямую ссылку для скачивания.
    https://drive.google.com/file/d/FILE_ID/view?... → https://drive.google.com/uc?export=download&id=FILE_ID
    """
    url = url.strip()
    match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url


async def send_post(bot: Bot, chat_id: str, text: str, media_url: str | None):
    """Отправляет пост — текст или текст с фото."""
    chat_id = chat_id.strip()
    if media_url and media_url.strip():
        url = convert_drive_url(media_url)
        logger.info(f"Медиа URL: {url}")
        await bot.send_photo(chat_id=chat_id, photo=url, caption=text)
    else:
        await bot.send_message(chat_id=chat_id, text=text)


async def check_and_send(bot: Bot):
    """Читает таблицу и отправляет посты по расписанию."""
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
        logger.warning("Таблица пустая")
        return

    # Пропускаем заголовок
    start_row = 1 if rows[0][COL_DATE].lower() in ("дата", "date") else 0
    logger.info(f"Всего строк: {len(rows)}, проверяю с строки {start_row + 1}")

    for i, row in enumerate(rows[start_row:], start=start_row + 2):
        if len(row) < 4:
            continue

        # Пропускаем уже отправленные
        status = row[COL_STATUS].strip() if len(row) > COL_STATUS else ""
        if status.startswith("✅"):
            continue

        date_val = row[COL_DATE]
        time_val = row[COL_TIME] if len(row) > COL_TIME else ""

        # Пропускаем строки без даты
        if not date_val.strip():
            continue

        post_dt = parse_datetime(date_val, time_val, tz)
        if post_dt is None:
            logger.warning(f"Строка {i}: не удалось распарсить дату '{date_val}' время '{time_val}'")
            continue

        logger.info(f"Строка {i}: дата поста {post_dt.strftime('%d.%m.%Y %H:%M')}, сейчас {now.strftime('%d.%m.%Y %H:%M')}")

        if window_start <= post_dt <= now:
            chat_id = row[COL_CHAT]
            text    = row[COL_TEXT]
            media   = row[COL_MEDIA].strip() if len(row) > COL_MEDIA else ""

            logger.info(f"Строка {i}: отправляю в {chat_id}")
            try:
                await send_post(bot, chat_id, text, media)
                new_status = f"✅ Отправлено {now.strftime('%d.%m %H:%M')}"
                logger.info(f"Строка {i}: успешно!")
            except TelegramError as e:
                new_status = f"❌ Telegram ошибка: {e}"
                logger.error(f"Строка {i}: {e}")
            except Exception as e:
                new_status = f"❌ Ошибка: {e}"
                logger.error(f"Строка {i}: {e}")

            try:
                sheet.update_cell(i, COL_STATUS + 1, new_status)
            except Exception as e:
                logger.error(f"Не удалось записать статус: {e}")


async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    me = await bot.get_me()
    logger.info(f"✅ Бот запущен: @{me.username}")
    logger.info(f"Проверка каждые {CHECK_INTERVAL} секунд")
    logger.info(f"Таблица ID: {SPREADSHEET_ID}")
    logger.info(f"Лист: {SHEET_NAME}")
    logger.info(f"Часовой пояс: {TIMEZONE}")

    while True:
        try:
            await check_and_send(bot)
        except Exception as e:
            logger.error(f"Ошибка в главном цикле: {e}")
        await asyncio.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    asyncio.run(main())
