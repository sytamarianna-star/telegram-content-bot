import os
import asyncio
import logging
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

# ── Настройки (берутся из переменных окружения Railway) ──────────────────────
TELEGRAM_TOKEN   = os.environ["TELEGRAM_TOKEN"]
SPREADSHEET_ID   = os.environ["SPREADSHEET_ID"]   # ID из URL таблицы
SHEET_NAME       = os.environ.get("SHEET_NAME", "Лист1")
TIMEZONE         = os.environ.get("TIMEZONE", "Europe/Kiev")   # твой часовой пояс
CHECK_INTERVAL   = int(os.environ.get("CHECK_INTERVAL", "60")) # секунды между проверками

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Колонки таблицы (индексы, начиная с 0)
COL_DATE   = 0   # A: Дата (формат: DD.MM.YYYY или YYYY-MM-DD)
COL_TIME   = 1   # B: Время (формат: HH:MM)
COL_CHAT   = 2   # C: Chat ID (число со знаком минус, например -1001234567890)
COL_TEXT   = 3   # D: Текст поста
COL_MEDIA  = 4   # E: Ссылка на фото/видео (необязательно)
COL_STATUS = 5   # F: Статус (бот сам пишет "✅ Отправлено" или "❌ Ошибка")


def get_sheet():
    """Подключение к Google Sheets через Service Account."""
    import json
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


def parse_datetime(date_str: str, time_str: str, tz) -> datetime | None:
    """Парсит дату и время из ячеек таблицы."""
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
    """Отправляет пост в чат — текст или текст с фото."""
    chat_id = chat_id.strip()
    if media_url and media_url.strip():
        await bot.send_photo(chat_id=chat_id, photo=media_url.strip(), caption=text)
    else:
        await bot.send_message(chat_id=chat_id, text=text, parse_mode="HTML")


async def check_and_send(bot: Bot):
    """Главная логика: читает таблицу, ищет посты к публикации."""
    tz = pytz.timezone(TIMEZONE)
    now = datetime.now(tz)
    window_start = now - timedelta(minutes=1)

    try:
        sheet = get_sheet()
        rows = sheet.get_all_values()
    except Exception as e:
        logger.error(f"Ошибка чтения таблицы: {e}")
        return

    if not rows:
        return

    # Пропускаем первую строку если это заголовок
    start_row = 1 if rows and rows[0][COL_DATE].lower() in ("дата", "date", "a") else 0

    for i, row in enumerate(rows[start_row:], start=start_row + 2):  # +2 = номер строки в Sheets
        # Пропускаем короткие строки и уже отправленные
        if len(row) < 4:
            continue
        status = row[COL_STATUS].strip() if len(row) > COL_STATUS else ""
        if status.startswith("✅"):
            continue

        post_dt = parse_datetime(row[COL_DATE], row[COL_TIME], tz)
        if post_dt is None:
            continue

        # Публикуем если время пришло (с окном 1 минута назад — на случай задержки)
        if window_start <= post_dt <= now:
            chat_id  = row[COL_CHAT]
            text     = row[COL_TEXT]
            media    = row[COL_MEDIA] if len(row) > COL_MEDIA else ""

            logger.info(f"Строка {i}: отправляю в {chat_id} — «{text[:40]}...»")
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

            # Записываем статус обратно в таблицу
            try:
                sheet.update_cell(i, COL_STATUS + 1, new_status)
            except Exception as e:
                logger.error(f"Не удалось записать статус в строку {i}: {e}")


async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    me = await bot.get_me()
    logger.info(f"Бот запущен: @{me.username}")

    # Выводим chat_id всех чатов где есть бот (удобно при первом запуске)
    logger.info("Бот работает. Проверка каждые %s секунд.", CHECK_INTERVAL)

    while True:
        try:
            await check_and_send(bot)
        except Exception as e:
            logger.error(f"Неожиданная ошибка в главном цикле: {e}")
        await asyncio.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    asyncio.run(main())
