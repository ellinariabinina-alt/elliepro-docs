# -*- coding: utf-8 -*-
"""
Бот сбора документов самозанятых для Ellie.pro.

Что делает:
  1) Сотрудник при /start выбирает свою именную папку на Гугл Диске (один раз).
  2) Присылает счёт -> бот грузит файл в папку и пишет Эллине с кнопкой "Оплачено".
  3) Эллина жмёт "Оплачено" -> бот просит у сотрудника чек и акт.
  4) Сотрудник присылает чек/акт -> бот грузит в ту же папку.
  5) Когда чек и акт собраны -> оплата закрывается, Эллине уходит уведомление.
  6) Чего не хватает -> бот напоминает сотруднику, через N часов эскалирует Эллине.

Всё состояние хранится в Google-таблице (не в файлах) — переживает рестарт Railway.
"""

import os
import json
import time
import asyncio
import logging
import tempfile
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, filters,
)

# ============================ НАСТРОЙКИ ============================

BOT_TOKEN = os.environ["BOT_TOKEN"]                       # из переменных Railway
GOOGLE_CREDENTIALS_JSON = os.environ["GOOGLE_CREDENTIALS_JSON"]  # из переменных Railway

PARENT_FOLDER_ID = "17dMbmbx6WLgNDRWYQb3L0wFNtdY9-jQG"     # родительская папка на Диске
SPREADSHEET_ID = "1obb5f5CS1VmDnVaXJzpcEAsjb0uqxOYvQBHg1MOpENM"  # таблица-трекер
ELLINA_ID = 87998099                                      # личный Telegram ID для эскалаций

MSK = ZoneInfo("Europe/Moscow")

# Тайминги напоминаний (в часах)
FIRST_REMINDER_AFTER_H = 4     # первое напоминание сотруднику после "Оплачено"
REPEAT_EVERY_H = 24            # повтор напоминания
ESCALATE_AFTER_H = 48         # после этого — эскалация Эллине
REMINDER_JOB_INTERVAL_SEC = 1800  # как часто проверять (раз в 30 минут)

# ============================ ЛОГИ ============================

logging.basicConfig(
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO,
)
log = logging.getLogger("ellie-docs-bot")

# ============================ GOOGLE ============================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

_creds_info = json.loads(GOOGLE_CREDENTIALS_JSON)
_creds = Credentials.from_service_account_info(_creds_info, scopes=SCOPES)
_gc = gspread.authorize(_creds)
_drive = build("drive", "v3", credentials=_creds, cache_discovery=False)
_sh = _gc.open_by_key(SPREADSHEET_ID)

EMP_HEADERS = ["user_id", "username", "ФИО", "folder_id", "folder_name", "registered_at"]
PAY_HEADERS = [
    "payment_id", "user_id", "ФИО", "folder_id", "дата", "сумма", "статус",
    "счёт", "акт", "чек", "создан_at", "оплачен_at",
    "последнее_напоминание", "эскалация", "закрыт_at",
]


def _ensure_ws(title, headers):
    try:
        ws = _sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = _sh.add_worksheet(title=title, rows=2000, cols=len(headers))
    values = ws.get_all_values()
    if not values:
        # лист пустой — пишем заголовок
        ws.append_row(headers)
    elif values[0][:len(headers)] != headers:
        # заголовка нет (первая строка — данные) — вставляем заголовок сверху
        ws.insert_row(headers, index=1)
    return ws


EMP_WS = _ensure_ws("Сотрудники", EMP_HEADERS)
PAY_WS = _ensure_ws("Оплаты", PAY_HEADERS)

# ============================ ХЕЛПЕРЫ (синхронные) ============================
# Все обращения к gspread/Drive — синхронные. В хендлерах вызываем их через
# asyncio.to_thread(), чтобы не блокировать бота.


def now_utc_iso():
    return datetime.now(timezone.utc).isoformat()


def parse_iso(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def msk_date():
    return datetime.now(MSK).strftime("%Y-%m-%d")


def _safe_q(text):
    return text.replace("\\", "\\\\").replace("'", "\\'")


# ---- Сотрудники ----

def find_employee(user_id):
    values = EMP_WS.get_all_values()
    for r in values:
        if not r or not r[0] or r[0] == EMP_HEADERS[0]:
            continue  # пустые строки и строка-заголовок
        if str(r[0]) == str(user_id):
            return {EMP_HEADERS[j]: (r[j] if j < len(r) else "") for j in range(len(EMP_HEADERS))}
    return None


def save_employee(user_id, username, fio, folder_id, folder_name):
    values = EMP_WS.get_all_values()
    for i, r in enumerate(values, start=1):
        if not r or not r[0] or r[0] == EMP_HEADERS[0]:
            continue
        if str(r[0]) == str(user_id):
            # перезапись папки при /repick
            EMP_WS.update_cell(i, EMP_HEADERS.index("username") + 1, username or "")
            EMP_WS.update_cell(i, EMP_HEADERS.index("folder_id") + 1, folder_id)
            EMP_WS.update_cell(i, EMP_HEADERS.index("folder_name") + 1, folder_name)
            return
    EMP_WS.append_row([
        str(user_id), username or "", fio, folder_id, folder_name, now_utc_iso(),
    ])


# ---- Папки на Диске ----

def search_folders(text):
    q = (
        f"'{PARENT_FOLDER_ID}' in parents and "
        f"mimeType='application/vnd.google-apps.folder' and trashed=false and "
        f"name contains '{_safe_q(text)}'"
    )
    res = _drive.files().list(
        q=q, fields="files(id,name)", pageSize=25, orderBy="name",
    ).execute()
    return res.get("files", [])


def upload_to_drive(folder_id, local_path, filename):
    meta = {"name": filename, "parents": [folder_id]}
    media = MediaFileUpload(local_path, resumable=False)
    f = _drive.files().create(
        body=meta, media_body=media, fields="id,webViewLink",
    ).execute()
    return f.get("webViewLink", "")


# ---- Оплаты ----

def _pay_rows():
    """Возвращает список (номер_строки, dict) по листу Оплаты."""
    values = PAY_WS.get_all_values()
    out = []
    for i, r in enumerate(values, start=1):
        if not r or not r[0] or r[0] == PAY_HEADERS[0]:
            continue  # пустые строки и строка-заголовок
        d = {PAY_HEADERS[j]: (r[j] if j < len(r) else "") for j in range(len(PAY_HEADERS))}
        out.append((i, d))
    return out


def _pay_update(rownum, field, value):
    PAY_WS.update_cell(rownum, PAY_HEADERS.index(field) + 1, value)


def create_payment(user_id, fio, folder_id, schet_link):
    pid = f"P{int(time.time())}"
    PAY_WS.append_row([
        pid, str(user_id), fio, folder_id, msk_date(), "", "новый",
        schet_link, "", "", now_utc_iso(), "", "", "", "",
    ])
    return pid


def get_payment(pid):
    for rownum, d in _pay_rows():
        if d["payment_id"] == pid:
            return rownum, d
    return None, None


def set_amount(user_id, amount):
    """Записывает сумму в последний счёт сотрудника без суммы."""
    target = None
    for rownum, d in _pay_rows():
        if str(d["user_id"]) == str(user_id) and not d["сумма"] and d["счёт"]:
            target = (rownum, d)
    if target:
        _pay_update(target[0], "сумма", amount)
        return True
    return False


def mark_paid(pid):
    rownum, d = get_payment(pid)
    if not rownum:
        return None
    _pay_update(rownum, "статус", "оплачен")
    _pay_update(rownum, "оплачен_at", now_utc_iso())
    return d


def attach_doc(user_id, doc_type, link):
    """Прикрепляет акт/чек к последней оплаченной незакрытой оплате сотрудника."""
    field = "акт" if doc_type == "акт" else "чек"
    target = None
    for rownum, d in _pay_rows():
        if str(d["user_id"]) == str(user_id) and d["статус"] == "оплачен":
            target = (rownum, d)  # берём самую позднюю подходящую
    if not target:
        return None, False
    rownum, d = target
    _pay_update(rownum, field, link)
    d[field] = link
    # проверяем комплект
    closed = bool(d.get("акт")) and bool(d.get("чек"))
    if closed:
        _pay_update(rownum, "статус", "закрыт")
        _pay_update(rownum, "закрыт_at", now_utc_iso())
    return d, closed


# ============================ ТЕКСТЫ ============================

WELCOME = (
    "Привет! Я бот для сбора документов по оплатам самозанятым.\n\n"
    "Через меня проходят 3 документа: счёт, чек из «Мой налог» и акт.\n\n"
    "Сначала выберите вашу папку. Напишите вашу фамилию (как называется ваша папка на Диске):"
)

HELP = (
    "Как пользоваться:\n"
    "1. Пришлите счёт — я загружу его и сообщу руководителю.\n"
    "2. После оплаты я попрошу чек и акт.\n"
    "3. Пришлите чек и акт — я разложу их по папке.\n\n"
    "Команды: /start — выбрать папку, /repick — сменить папку."
)

# ============================ ХЕНДЛЕРЫ ============================


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    emp = await asyncio.to_thread(find_employee, user.id)
    if emp:
        await update.message.reply_text(
            f"Вы уже привязаны к папке: {emp['folder_name']}.\n"
            f"Можно сразу присылать счёт.\n\n{HELP}"
        )
        return
    context.user_data["reg"] = True
    await update.message.reply_text(WELCOME)


async def cmd_repick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["reg"] = True
    await update.message.reply_text(
        "Напишите вашу фамилию (как называется ваша папка на Диске):"
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(HELP)


async def cmd_open(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Список незакрытых оплат — только для Эллины."""
    if update.effective_user.id != ELLINA_ID:
        return
    rows = await asyncio.to_thread(_pay_rows)
    lines = []
    for _, d in rows:
        if d["статус"] == "оплачен":
            missing = []
            if not d["акт"]:
                missing.append("акт")
            if not d["чек"]:
                missing.append("чек")
            lines.append(f"• {d['ФИО']} (от {d['дата']}): не хватает {', '.join(missing) or '—'}")
    text = "Открытые оплаты:\n" + ("\n".join(lines) if lines else "всё собрано ✅")
    await update.message.reply_text(text)


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    text = (update.message.text or "").strip()

    # 1) Идёт регистрация — ищем папку
    if context.user_data.get("reg"):
        folders = await asyncio.to_thread(search_folders, text)
        if not folders:
            await update.message.reply_text(
                "Не нашёл такую папку. Напишите точнее или обратитесь к руководителю."
            )
            return
        context.user_data["folder_candidates"] = folders
        kb = [
            [InlineKeyboardButton(f["name"], callback_data=f"folder:{i}")]
            for i, f in enumerate(folders[:20])
        ]
        await update.message.reply_text(
            "Выберите вашу папку:", reply_markup=InlineKeyboardMarkup(kb)
        )
        return

    # 2) Ждём сумму по счёту
    if context.user_data.get("awaiting_amount"):
        ok = await asyncio.to_thread(set_amount, user.id, text)
        context.user_data["awaiting_amount"] = False
        if ok:
            await update.message.reply_text("Сумма записана, спасибо.")
        return

    # 3) Прочий текст
    emp = await asyncio.to_thread(find_employee, user.id)
    if not emp:
        await update.message.reply_text("Сначала /start и выберите вашу папку.")
    else:
        await update.message.reply_text("Пришлите документ файлом или фото. " + HELP)


async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    emp = await asyncio.to_thread(find_employee, user.id)
    if not emp:
        await update.message.reply_text("Сначала /start и выберите вашу папку.")
        return

    msg = update.message
    if msg.document:
        file_id = msg.document.file_id
        orig_name = msg.document.file_name or "file"
        ext = os.path.splitext(orig_name)[1] or ".pdf"
    elif msg.photo:
        file_id = msg.photo[-1].file_id
        ext = ".jpg"
    else:
        return

    context.user_data["pending_file"] = {"file_id": file_id, "ext": ext}
    kb = [[
        InlineKeyboardButton("Счёт", callback_data="type:счёт"),
        InlineKeyboardButton("Акт", callback_data="type:акт"),
        InlineKeyboardButton("Чек", callback_data="type:чек"),
    ]]
    await update.message.reply_text("Что это за документ?", reply_markup=InlineKeyboardMarkup(kb))


async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    data = q.data
    user = q.from_user

    # ----- выбор папки при регистрации -----
    if data.startswith("folder:"):
        idx = int(data.split(":", 1)[1])
        cands = context.user_data.get("folder_candidates", [])
        if idx >= len(cands):
            await q.edit_message_text("Попробуйте /start заново.")
            return
        folder = cands[idx]
        # ФИО берём из имени папки
        await asyncio.to_thread(
            save_employee, user.id, user.username, folder["name"],
            folder["id"], folder["name"],
        )
        context.user_data["reg"] = False
        context.user_data.pop("folder_candidates", None)
        await q.edit_message_text(
            f"Готово. Вы привязаны к папке: {folder['name']}.\n\n{HELP}"
        )
        return

    # ----- тип присланного документа -----
    if data.startswith("type:"):
        doc_type = data.split(":", 1)[1]
        pending = context.user_data.get("pending_file")
        if not pending:
            await q.edit_message_text("Файл не найден, пришлите заново.")
            return
        emp = await asyncio.to_thread(find_employee, user.id)
        if not emp:
            await q.edit_message_text("Сначала /start и выберите папку.")
            return

        # скачиваем файл из Telegram
        tg_file = await context.bot.get_file(pending["file_id"])
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=pending["ext"])
        tmp.close()
        await tg_file.download_to_drive(tmp.name)

        type_label = {"счёт": "Счёт", "акт": "Акт", "чек": "Чек"}[doc_type]
        fio_safe = emp["ФИО"].replace(" ", "_")
        filename = f"{msk_date()}_{type_label}_{fio_safe}{pending['ext']}"

        try:
            link = await asyncio.to_thread(
                upload_to_drive, emp["folder_id"], tmp.name, filename
            )
        finally:
            try:
                os.unlink(tmp.name)
            except OSError:
                pass

        context.user_data.pop("pending_file", None)

        if doc_type == "счёт":
            pid = await asyncio.to_thread(
                create_payment, user.id, emp["ФИО"], emp["folder_id"], link
            )
            await q.edit_message_text("Счёт получен и сохранён ✅")
            # уведомляем Эллину
            kb = [[InlineKeyboardButton("Оплачено ✅", callback_data=f"paid:{pid}")]]
            await context.bot.send_message(
                chat_id=ELLINA_ID,
                text=(
                    f"💳 Новый счёт\n"
                    f"Сотрудник: {emp['ФИО']}\n"
                    f"Файл: {link}\n\n"
                    f"После оплаты нажмите кнопку — бот запросит чек и акт."
                ),
                reply_markup=InlineKeyboardMarkup(kb),
            )
            # опционально просим сумму (не блокирует)
            context.user_data["awaiting_amount"] = True
            await context.bot.send_message(
                chat_id=user.id,
                text="Если не сложно, укажите сумму по счёту (для учёта). Или пропустите."
            )
        else:
            res, closed = await asyncio.to_thread(attach_doc, user.id, doc_type, link)
            if not res:
                await q.edit_message_text(
                    f"{type_label} сохранён, но открытой оплаты не нашёл. "
                    f"Если это новый счёт — пришлите его как «Счёт»."
                )
                return
            await q.edit_message_text(f"{type_label} получен и сохранён ✅")
            if closed:
                await context.bot.send_message(
                    chat_id=user.id,
                    text="Спасибо! Комплект документов собран полностью ✅"
                )
                await context.bot.send_message(
                    chat_id=ELLINA_ID,
                    text=(
                        f"✅ Комплект собран\n"
                        f"Сотрудник: {res['ФИО']} (счёт от {res['дата']})\n"
                        f"Счёт, акт и чек в папке."
                    ),
                )
            else:
                missing = "чек" if not res.get("чек") else "акт"
                await context.bot.send_message(
                    chat_id=user.id,
                    text=f"Принято. Осталось прислать: {missing}."
                )
        return

    # ----- кнопка "Оплачено" (только Эллина) -----
    if data.startswith("paid:"):
        if user.id != ELLINA_ID:
            await q.answer("Кнопка только для руководителя.", show_alert=True)
            return
        pid = data.split(":", 1)[1]
        d = await asyncio.to_thread(mark_paid, pid)
        if not d:
            await q.edit_message_text("Оплата не найдена.")
            return
        await q.edit_message_text(
            f"💳 Счёт {d['ФИО']} (от {d['дата']}) отмечен как оплаченный ✅\n"
            f"Бот запросил у сотрудника чек и акт."
        )
        await context.bot.send_message(
            chat_id=int(d["user_id"]),
            text=(
                f"Оплата по счёту от {d['дата']} получена ✅\n\n"
                f"Пришлите, пожалуйста:\n"
                f"• чек из приложения «Мой налог»\n"
                f"• акт\n\n"
                f"Просто отправьте файлы сюда — я разложу их по папке."
            ),
        )
        return


# ============================ НАПОМИНАНИЯ ============================


async def reminder_job(context: ContextTypes.DEFAULT_TYPE):
    rows = await asyncio.to_thread(_pay_rows)
    now = datetime.now(timezone.utc)
    for rownum, d in rows:
        if d["статус"] != "оплачен":
            continue
        missing = []
        if not d["акт"]:
            missing.append("акт")
        if not d["чек"]:
            missing.append("чек")
        if not missing:
            continue

        paid_at = parse_iso(d["оплачен_at"]) or now
        hours_since_paid = (now - paid_at).total_seconds() / 3600

        # напоминание сотруднику
        last = parse_iso(d["последнее_напоминание"])
        should_remind = False
        if hours_since_paid >= FIRST_REMINDER_AFTER_H:
            if last is None:
                should_remind = True
            elif (now - last).total_seconds() / 3600 >= REPEAT_EVERY_H:
                should_remind = True
        if should_remind:
            try:
                await context.bot.send_message(
                    chat_id=int(d["user_id"]),
                    text=(
                        f"Напоминание по оплате от {d['дата']}.\n"
                        f"Не хватает: {', '.join(missing)}.\n"
                        f"Пришлите, пожалуйста, сюда файлом."
                    ),
                )
                await asyncio.to_thread(_pay_update, rownum, "последнее_напоминание", now_utc_iso())
            except Exception as e:
                log.warning("Не смог напомнить %s: %s", d["user_id"], e)

        # эскалация Эллине
        if hours_since_paid >= ESCALATE_AFTER_H and d["эскалация"] != "да":
            try:
                await context.bot.send_message(
                    chat_id=ELLINA_ID,
                    text=(
                        f"⚠️ {d['ФИО']} не прислал {', '.join(missing)} "
                        f"по оплате от {d['дата']} уже больше {ESCALATE_AFTER_H} ч."
                    ),
                )
                await asyncio.to_thread(_pay_update, rownum, "эскалация", "да")
            except Exception as e:
                log.warning("Не смог эскалировать %s: %s", d["user_id"], e)


# ============================ ЗАПУСК ============================


def main():
    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("repick", cmd_repick))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("open", cmd_open))
    app.add_handler(MessageHandler(filters.Document.ALL | filters.PHOTO, on_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.add_handler(CallbackQueryHandler(on_callback))

    app.job_queue.run_repeating(
        reminder_job, interval=REMINDER_JOB_INTERVAL_SEC, first=60
    )

    log.info("Бот запущен.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
