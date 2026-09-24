import os
import asyncio
import io
import logging
import json
import re
import shutil
import tempfile
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

# Локальный кэш отправленных постов (защита от дублей)
sent_posts = set()


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


VIDEO_EXTENSIONS = (".mp4", ".mov", ".m4v", ".webm", ".mkv", ".avi")
MAX_VIDEO_BYTES = 49 * 1024 * 1024


def convert_drive_url(url: str) -> str:
    url = url.strip()
    match = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=download&id={file_id}"
    return url


def _drive_file_id(url: str) -> str | None:
    match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    match = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)
    return None


def _is_youtube(url: str) -> bool:
    host = urllib.parse.urlparse(url).netloc.lower()
    if host.startswith("www."):
        host = host[4:]
    return host in {
        "youtu.be",
        "youtube.com",
        "m.youtube.com",
        "music.youtube.com",
        "youtube-nocookie.com",
    }


def _path_is_video(url: str) -> bool:
    path = urllib.parse.urlparse(url).path.lower()
    return path.endswith(VIDEO_EXTENSIONS)


def _video_filename(name: str, mime: str = "") -> str:
    clean = (name or "").strip()
    if clean.lower().endswith(VIDEO_EXTENSIONS):
        return clean
    mime_ext = {
        "video/mp4": ".mp4",
        "video/quicktime": ".mov",
        "video/webm": ".webm",
        "video/x-matroska": ".mkv",
        "video/x-msvideo": ".avi",
    }.get(mime, ".mp4")
    return f"video{mime_ext}"


def _google_access_token() -> str | None:
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON", "").strip()
    if not creds_json:
        return None
    try:
        from google.auth.transport.requests import Request as GoogleRequest
        creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
        creds.refresh(GoogleRequest())
        return creds.token
    except Exception as exc:
        logger.info(f"Токен Google Drive недоступен: {exc}")
        return None


def _drive_meta(file_id: str) -> dict | None:
    token = _google_access_token()
    if not token:
        return None
    url = (
        "https://www.googleapis.com/drive/v3/files/"
        f"{file_id}?fields=mimeType,name,size&supportsAllDrives=true"
    )
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        logger.info(f"Метаданные Drive недоступны ({file_id}): {exc}")
        return None


def _download_bytes(url: str, headers: dict | None = None) -> tuple[bytes, str]:
    req = urllib.request.Request(url, headers=headers or {"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=180) as response:
        data = response.read()
        content_type = response.headers.get_content_type()
    return data, content_type


def _looks_like_html(data: bytes, content_type: str) -> bool:
    if content_type.startswith("text/html"):
        return True
    head = data.lstrip()[:20].lower()
    return head.startswith(b"<!doctype") or head.startswith(b"<html")


def _download_drive_media(file_id: str) -> bytes:
    token = _google_access_token()
    if token:
        url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&supportsAllDrives=true"
        try:
            data, content_type = _download_bytes(url, {"Authorization": f"Bearer {token}"})
            if data and not _looks_like_html(data, content_type):
                return data
        except Exception as exc:
            logger.info(f"Скачивание через Drive API не удалось: {exc}")
    public = f"https://drive.google.com/uc?export=download&confirm=t&id={file_id}"
    data, content_type = _download_bytes(public)
    if _looks_like_html(data, content_type):
        raise TelegramError(
            "Не удалось скачать видео с Google Drive. Откройте файл по ссылке «всем, у кого есть ссылка»."
        )
    return data


# У YouTube часто нет готового mp4 со звуком. Скачиваем картинку и звук
# отдельно и склеиваем через ffmpeg. Если файл больше лимита Telegram — ниже качество.
YOUTUBE_FORMATS = [
    "bv*[ext=mp4][height<=720]+ba[ext=m4a]/bv*[height<=720]+ba/b[height<=720]",
    "bv*[ext=mp4][height<=480]+ba[ext=m4a]/bv*[height<=480]+ba/b[height<=480]",
    "bv*[ext=mp4][height<=360]+ba[ext=m4a]/bv*[height<=360]+ba/b[height<=360]",
]
# С IP сервера обычная страница YouTube просит войти. Эти клиенты страницу не открывают.
YOUTUBE_CLIENTS = ("visionos", "android", "ios")


def _download_youtube(url: str) -> str:
    try:
        import yt_dlp
    except ImportError as exc:
        raise TelegramError("Для видео с YouTube на сервере нужен пакет yt-dlp") from exc

    last_error = "неизвестная ошибка"
    for client in YOUTUBE_CLIENTS:
        for fmt in YOUTUBE_FORMATS:
            tmp = tempfile.mkdtemp(prefix="tgvideo_")
            opts = {
                "format": fmt,
                "merge_output_format": "mp4",
                "outtmpl": os.path.join(tmp, "%(id)s.%(ext)s"),
                "noplaylist": True,
                "quiet": True,
                "no_warnings": True,
                "retries": 3,
                "fragment_retries": 3,
                "socket_timeout": 30,
                "extractor_args": {
                    "youtube": {
                        "player_client": [client],
                        "player_skip": ["webpage", "configs"],
                    }
                },
            }
            try:
                with yt_dlp.YoutubeDL(opts) as ydl:
                    ydl.download([url])
                files = [
                    os.path.join(tmp, name)
                    for name in os.listdir(tmp)
                    if os.path.isfile(os.path.join(tmp, name))
                ]
                if not files:
                    last_error = "YouTube не отдал видеофайл"
                    shutil.rmtree(tmp, ignore_errors=True)
                    continue
                path = max(files, key=os.path.getsize)
                if os.path.getsize(path) > MAX_VIDEO_BYTES:
                    last_error = "Видео больше 49 МБ — Telegram не примет его от бота"
                    shutil.rmtree(tmp, ignore_errors=True)
                    continue
                return path
            except Exception as exc:
                last_error = str(exc)
                shutil.rmtree(tmp, ignore_errors=True)
                lowered = last_error.lower()
                if ("not a bot" in lowered || "sign in to confirm" in lowered) {
                    break
                }

    if ("ffmpeg" in last_error.lower()) {
        raise TelegramError("На сервере нет ffmpeg, видео с YouTube не собралось")
    }
    raise TelegramError(f"Не удалось скачать видео с YouTube: {last_error}")


def _ensure_video_size(size: int) -> None:
    if size > MAX_VIDEO_BYTES:
        raise TelegramError("Видео больше 49 МБ — Telegram не примет его от бота")


async def _send_video_bytes(bot: Bot, chat_id: str, text: str, data: bytes, filename: str) -> None:
    _ensure_video_size(len(data))
    await bot.send_video(
        chat_id=chat_id,
        video=io.BytesIO(data),
        filename=filename,
        caption=text,
        supports_streaming=True,
    )


async def send_post(bot: Bot, chat_id: str, text: str, media_url: str):
    """Картинка уходит фото, видео — файлом, который можно смотреть в чате."""
    chat_id = chat_id.strip()
    raw = (media_url or "").strip()
    try:
        if not raw:
            await bot.send_message(chat_id=chat_id, text=text)
            return

        if _is_youtube(raw):
            path = await asyncio.to_thread(_download_youtube, raw)
            try:
                with open(path, "rb") as handle:
                    data = handle.read()
                name = _video_filename(os.path.basename(path))
                logger.info(f"YouTube скачан: {name}, {len(data)} байт")
                await _send_video_bytes(bot, chat_id, text, data, name)
            finally:
                shutil.rmtree(os.path.dirname(path), ignore_errors=True)
            return

        if _path_is_video(raw):
            logger.info(f"Прямая ссылка на видео: {raw}")
            data, _ = await asyncio.to_thread(_download_bytes, raw)
            filename = _video_filename(os.path.basename(urllib.parse.urlparse(raw).path))
            await _send_video_bytes(bot, chat_id, text, data, filename)
            return

        file_id = _drive_file_id(raw)
        if file_id:
            meta = await asyncio.to_thread(_drive_meta, file_id) or {}
            mime = meta.get("mimeType", "")
            name = meta.get("name", "")
            if mime.startswith("video/") or name.lower().endswith(VIDEO_EXTENSIONS):
                logger.info(f"Видео с Google Drive: {name or file_id} ({mime})")
                data = await asyncio.to_thread(_download_drive_media, file_id)
                await _send_video_bytes(bot, chat_id, text, data, _video_filename(name, mime))
                return

        url = convert_drive_url(raw)
        logger.info(f"Картинка: {url}")
        await bot.send_photo(chat_id=chat_id, photo=url, caption=text)
    except TelegramError:
        raise
    except Exception as exc:
        raise TelegramError(f"Не удалось отправить медиа: {exc}") from exc


async def handle_private_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message and update.message.chat.type == "private":
        user = update.message.from_user
        logger.info(f"Личное сообщение от {user.first_name}")
        try:
            await update.message.reply_photo(photo=AUTO_REPLY_PHOTO, caption=AUTO_REPLY_TEXT)
        except Exception as e:
            logger.error(f"Ошибка автоответа: {e}")
            try:
                await update.message.reply_text(AUTO_REPLY_TEXT)
            except Exception as e2:
                logger.error(f"Ошибка текстового автоответа: {e2}")


async def check_and_send(bot: Bot):
    global sent_posts
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

    # Находим строки где есть дата — это настоящие посты
    posts = []
    for i, row in enumerate(rows[start_row:], start=start_row):
        if len(row) < 5:
            continue
        date_val = row[COL_DATE].strip()
        if not date_val:
            continue
        # Номер строки в Sheets (1-based)
        sheet_row = i + 1
        posts.append((sheet_row, row))

    logger.info(f"Найдено постов с датой: {len(posts)}")

    for sheet_row, row in posts:
        # Пропускаем уже отправленные (по статусу в таблице)
        status = row[COL_STATUS].strip() if len(row) > COL_STATUS else ""
        if status.startswith("✅"):
            continue

        # Пропускаем уже отправленные (локальный кэш — защита от дублей)
        post_key = f"{row[COL_DATE]}_{row[COL_TIME]}_{row[COL_CHAT_RU]}"
        if post_key in sent_posts:
            continue

        date_val = row[COL_DATE]
        time_val = row[COL_TIME] if len(row) > COL_TIME else ""

        post_dt = parse_datetime(date_val, time_val, tz)
        if post_dt is None:
            continue

        logger.info(f"Строка {sheet_row}: дата поста {post_dt.strftime('%d.%m.%Y %H:%M')}, сейчас {now.strftime('%d.%m.%Y %H:%M')}")

        if window_start <= post_dt <= now:
            chat_ru   = row[COL_CHAT_RU].strip() if len(row) > COL_CHAT_RU else ""
            chat_ukr  = row[COL_CHAT_UKR].strip() if len(row) > COL_CHAT_UKR else ""
            text      = row[COL_TEXT]
            media_ru  = row[COL_MEDIA_RU].strip() if len(row) > COL_MEDIA_RU else ""
            media_ukr = row[COL_MEDIA_UKR].strip() if len(row) > COL_MEDIA_UKR else ""

            # Сразу добавляем в кэш чтобы не отправить дважды
            sent_posts.add(post_key)

            statuses = []

            if chat_ru:
                try:
                    await send_post(bot, chat_ru, text, media_ru)
                    statuses.append("✅ РУ")
                    logger.info(f"Строка {sheet_row}: РУ отправлено!")
                except TelegramError as e:
                    statuses.append(f"❌ РУ: {e}")
                    logger.error(f"Строка {sheet_row}: РУ ошибка — {e}")

            if chat_ukr:
                text_ukr = translate_to_ukrainian(text)
                media_for_ukr = media_ukr if media_ukr else media_ru
                try:
                    await send_post(bot, chat_ukr, text_ukr, media_for_ukr)
                    statuses.append("✅ УКР")
                    logger.info(f"Строка {sheet_row}: УКР отправлено!")
                except TelegramError as e:
                    statuses.append(f"❌ УКР: {e}")
                    logger.error(f"Строка {sheet_row}: УКР ошибка — {e}")

            new_status = f"{' | '.join(statuses)} {now.strftime('%d.%m %H:%M')}"
            try:
                sheet.update_cell(sheet_row, COL_STATUS + 1, new_status)
                logger.info(f"Статус записан в строку {sheet_row}")
            except Exception as e:
                logger.error(f"Не удалось записать статус в строку {sheet_row}: {e}")


async def scheduler(bot: Bot):
    while True:
        try:
            await check_and_send(bot)
        except Exception as e:
            logger.error(f"Ошибка в планировщике: {e}")
        await asyncio.sleep(CHECK_INTERVAL)


async def main():
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & filters.ChatType.PRIVATE, handle_private_message))

    bot = app.bot
    me = await bot.get_me()
    logger.info(f"✅ Бот запущен: @{me.username}")
    logger.info(f"Проверка каждые {CHECK_INTERVAL} секунд")
    logger.info(f"Таблица ID: {SPREADSHEET_ID}")

    async with app:
        await app.start()
        await app.updater.start_polling()
        await scheduler(bot)


if __name__ == "__main__":
    asyncio.run(main())
