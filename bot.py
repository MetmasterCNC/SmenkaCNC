import asyncio
import json
import os
import re
from datetime import datetime, timedelta, timezone
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import (
    Message,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
    FSInputFile,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ================= НАСТРОЙКИ =================
BOT_TOKEN = os.getenv("BOT_TOKEN")
ACCESS_PASSWORD = "123456789"  # Пароль для первого входа операторов
ADMIN_IDS = [1011177377]   # Ваш Telegram ID (числом без кавычек)

# Часовой пояс: 3 — Москва / Казань (UTC+3)
LOCAL_TZ = timezone(timedelta(hours=3))

EXCEL_FILE = "cnc_work_log.xlsx"
USERS_FILE = "operators.json"

# Список ваших станков (кнопки)
MACHINES = [
    ["VMH", "Micron","CUT20P"],
    ["Deg","Das"],
    ["Hitech","DMG"],
    ["GA2000","LG","LG-mini"],
    ["JSL","Tornoss"]
]

# Список операций (кнопки)
OPERATIONS = [
    ["Наладка", "Черновая"],
    ["Чистовая", "Сверление / Резьба"],
    ["Контроль / Доработка"]
]
# =============================================

router = Router()


def get_now() -> datetime:
    return datetime.now(LOCAL_TZ)


class WorkForm(StatesGroup):
    waiting_password = State()
    waiting_fio = State()
    choosing_machine = State()
    entering_part = State()
    choosing_operation = State()
    entering_qty = State()
    entering_note = State()


# --- Работа с базой операторов (JSON) ---
def load_users() -> dict:
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_user(user_id: int, fio: str):
    users = load_users()
    users[str(user_id)] = fio
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


# --- Работа с Excel (Общий + Месяцы + Операторы + Фильтры + Удаление) ---
def apply_sheet_enhancements(ws):
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:G1"


def format_sheet(ws):
    headers = ["Дата и время", "Оператор", "Станок", "Деталь / Чертеж", "Операция", "Кол-во (шт)", "Примечание"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    widths = [20, 25, 20, 25, 22, 14, 35]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w

    apply_sheet_enhancements(ws)


def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Общий журнал"
        format_sheet(ws)
        wb.save(EXCEL_FILE)


def get_or_create_sheet(wb, sheet_title: str):
    clean_title = re.sub(r'[\\/*?:\[\]]', '', sheet_title).strip()[:31] or "Лист"
    if clean_title not in wb.sheetnames:
        ws = wb.create_sheet(title=clean_title)
        format_sheet(ws)
    else:
        ws = wb[clean_title]
        apply_sheet_enhancements(ws)
    return ws


def append_to_excel(row_data: list):
    init_excel()
    wb = load_workbook(EXCEL_FILE)

    # 1. Запись в "Общий журнал"
    ws_main = wb.worksheets[0]
    apply_sheet_enhancements(ws_main)
    ws_main.append(row_data)

    # 2. Запись в лист текущего месяца (например, "11.2026")
    month_title = get_now().strftime("%m.%Y")
    ws_month = get_or_create_sheet(wb, month_title)
    ws_month.append(row_data)

    # 3. Запись в именной лист оператора (row_data[1] — ФИО)
    operator_fio = str(row_data[1])
    ws_operator = get_or_create_sheet(wb, operator_fio)
    ws_operator.append(row_data)

    wb.save(EXCEL_FILE)


def delete_last_entry(target_fio: str | None = None) -> list | None:
    if not os.path.exists(EXCEL_FILE):
        return None

    wb = load_workbook(EXCEL_FILE)
    ws_main = wb.worksheets[0]

    target_row_idx = None
    deleted_data = None

    for row_idx in range(ws_main.max_row, 1, -1):
        row_values = [ws_main.cell(row=row_idx, column=col).value for col in range(1, 8)]
        if target_fio is None or str(row_values[1]) == target_fio:
            target_row_idx = row_idx
            deleted_data = row_values
            break

    if not target_row_idx or not deleted_data:
        return None

    # 1. Удаляем из "Общего журнала"
    ws_main.delete_rows(target_row_idx, 1)

    # 2. Синхронно удаляем из вкладок месяца и оператора
    for ws in wb.worksheets[1:]:
        for r_idx in range(ws.max_row, 1, -1):
            r_vals = [ws.cell(row=r_idx, column=col).value for col in range(1, 8)]
            if r_vals == deleted_data:
                ws.delete_rows(r_idx, 1)
                break

    wb.save(EXCEL_FILE)
    return deleted_data


# --- Клавиатуры ---
def to_kb_rows(matrix: list) -> list:
    return [[KeyboardButton(text=str(item)) for item in row] for row in matrix]


def main_menu_kb(user_id: int):
    buttons = [
        [KeyboardButton(text="➕ Внести работу")],
        [KeyboardButton(text="🗑 Удалить последнюю запись")]
    ]
    if user_id in ADMIN_IDS:
        buttons.append([KeyboardButton(text="📊 Скачать Excel-отчёт"), KeyboardButton(text="👤 Сменить ФИО")])
    else:
        buttons.append([KeyboardButton(text="👤 Сменить ФИО")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def cancel_kb():
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True
    )


# --- Обработчики ---
@router.message(F.text == "❌ Отмена")
async def cancel_action(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Действие отменено.", reply_markup=main_menu_kb(message.from_user.id))


@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    users = load_users()
    user_id = str(message.from_user.id)

    if user_id in users:
        await message.answer(
            f"Добро пожаловать, **{users[user_id]}**!\nНажмите «➕ Внести работу», чтобы записать смену.",
            reply_markup=main_menu_kb(message.from_user.id),
            parse_mode="Markdown"
        )
    else:
        await message.answer(
            "Доступ ограничен. Введите цеховой пароль для регистрации:",
            reply_markup=ReplyKeyboardRemove()
        )
        await state.set_state(WorkForm.waiting_password)


@router.message(WorkForm.waiting_password)
async def process_password(message: Message, state: FSMContext):
    if message.text and message.text.strip() == ACCESS_PASSWORD:
        await message.answer("Пароль верный. Введите ваши **Фамилию и Имя** (например: *Иванов А.С.*):", parse_mode="Markdown")
        await state.set_state(WorkForm.waiting_fio)
    else:
        await message.answer("Неверный пароль. Попробуйте ещё раз:")


@router.message(WorkForm.waiting_fio)
async def process_fio(message: Message, state: FSMContext):
    fio = message.text.strip() if message.text else ""
    if len(fio) < 3:
        await message.answer("Слишком короткое имя. Введите Фамилию и Имя:")
        return
    save_user(message.from_user.id, fio)
    await state.clear()
    await message.answer(
        f"Готово! Вы зарегистрированы как **{fio}**.",
        reply_markup=main_menu_kb(message.from_user.id),
        parse_mode="Markdown"
    )


@router.message(F.text == "👤 Сменить ФИО")
async def change_fio(message: Message, state: FSMContext):
    users = load_users()
    if str(message.from_user.id) not in users:
        await cmd_start(message, state)
        return
    await message.answer("Введите новые Фамилию и Имя:", reply_markup=cancel_kb())
    await state.set_state(WorkForm.waiting_fio)


@router.message(F.text == "📊 Скачать Excel-отчёт")
async def send_excel(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("⛔ У вас нет прав для скачивания общего журнала.")
        return

    init_excel()
    doc = FSInputFile(EXCEL_FILE)
    await message.answer_document(doc, caption=f"Актуальный журнал работ на {get_now().strftime('%d.%m.%Y %H:%M')}")


@router.message(Command("undo"))
@router.message(F.text == "🗑 Удалить последнюю запись")
async def undo_last_entry(message: Message, bot: Bot):
    users = load_users()
    user_id = str(message.from_user.id)
    if user_id not in users:
        await message.answer("Сначала пройдите регистрацию через /start")
        return

    is_admin = message.from_user.id in ADMIN_IDS
    target_fio = None if is_admin else users[user_id]

    try:
        deleted = delete_last_entry(target_fio=target_fio)
    except PermissionError:
        await message.answer("⚠️ Ошибка! Файл Excel сейчас открыт. Закройте его и попробуйте снова.")
        return

    if not deleted:
        await message.answer("ℹ️ Записей для удаления не найдено.")
        return

    details = (
        f"📅 Дата: `{deleted[0]}`\n"
        f"👤 Оператор: {deleted[1]}\n"
        f"⚙️ Станок: {deleted[2]}\n"
        f"📐 Деталь: {deleted[3]}\n"
        f"🛠 Операция: {deleted[4]}\n"
        f"🔢 Кол-во: {deleted[5]} шт.\n"
        f"📝 Примечание: {deleted[6] or '—'}"
    )
    await message.answer(
        f"🗑 **Запись удалена из всех вкладок:**\n\n{details}",
        reply_markup=main_menu_kb(message.from_user.id),
        parse_mode="Markdown"
    )

    for admin_id in ADMIN_IDS:
        if admin_id != message.from_user.id:
            try:
                await bot.send_message(
                    admin_id,
                    f"⚠️ **Оператор {users[user_id]} удалил свою последнюю запись:**\n\n{details}",
                    parse_mode="Markdown"
                )
            except Exception:
                pass


# --- Пошаговое заполнение работы ---
@router.message(F.text == "➕ Внести работу")
async def start_work_entry(message: Message, state: FSMContext):
    users = load_users()
    user_id = str(message.from_user.id)
    if user_id not in users:
        await cmd_start(message, state)
        return

    await state.update_data(fio=users[user_id])
    kb = ReplyKeyboardMarkup(
        keyboard=to_kb_rows(MACHINES) + [[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True
    )
    await message.answer("1/5. Выберите **станок** (или напишите вручную):", reply_markup=kb, parse_mode="Markdown")
    await state.set_state(WorkForm.choosing_machine)


@router.message(WorkForm.choosing_machine)
async def process_machine(message: Message, state: FSMContext):
    await state.update_data(machine=message.text.strip())
    await message.answer(
        "2/5. Введите **обозначение детали / номер чертежа**:",
        reply_markup=cancel_kb(),
        parse_mode="Markdown"
    )
    await state.set_state(WorkForm.entering_part)


@router.message(WorkForm.entering_part)
async def process_part(message: Message, state: FSMContext):
    await state.update_data(part=message.text.strip())
    kb = ReplyKeyboardMarkup(
        keyboard=to_kb_rows(OPERATIONS) + [[KeyboardButton(text="❌ Отмена")]],
        resize_keyboard=True
    )
    await message.answer("3/5. Выберите **операцию** (или напишите свою):", reply_markup=kb, parse_mode="Markdown")
    await state.set_state(WorkForm.choosing_operation)


@router.message(WorkForm.choosing_operation)
async def process_operation(message: Message, state: FSMContext):
    await state.update_data(operation=message.text.strip())
    await message.answer(
        "4/5. Введите **количество изготовленных деталей** (числом):",
        reply_markup=cancel_kb(),
        parse_mode="Markdown"
    )
    await state.set_state(WorkForm.entering_qty)


@router.message(WorkForm.entering_qty)
async def process_qty(message: Message, state: FSMContext):
    text = message.text.strip() if message.text else ""
    if not text.isdigit():
        await message.answer("Пожалуйста, введите целое число (например: 15):")
        return

    await state.update_data(qty=int(text))
    kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="Без примечаний")],
            [KeyboardButton(text="❌ Отмена")]
        ],
        resize_keyboard=True
    )
    await message.answer("5/5. Добавьте **примечание** (брак, смена пластин и т.д.) или нажмите «Без примечаний»:", reply_markup=kb, parse_mode="Markdown")
    await state.set_state(WorkForm.entering_note)


@router.message(WorkForm.entering_note)
async def process_note(message: Message, state: FSMContext, bot: Bot):
    note = "" if message.text.strip() == "Без примечаний" else message.text.strip()
    data = await state.get_data()
    now_str = get_now().strftime("%d.%m.%Y %H:%M")

    row = [
        now_str,
        data["fio"],
        data["machine"],
        data["part"],
        data["operation"],
        data["qty"],
        note
    ]

    try:
        append_to_excel(row)
        details = (
            f"📅 Дата: `{now_str}`\n"
            f"👤 Оператор: {data['fio']}\n"
            f"⚙️ Станок: {data['machine']}\n"
            f"📐 Деталь: {data['part']}\n"
            f"🛠 Операция: {data['operation']}\n"
            f"🔢 Кол-во: {data['qty']} шт.\n"
            f"📝 Примечание: {note or '—'}"
        )
        await message.answer(
            f"✅ **Запись сохранена!**\n\n{details}",
            reply_markup=main_menu_kb(message.from_user.id),
            parse_mode="Markdown"
        )

        for admin_id in ADMIN_IDS:
            if admin_id != message.from_user.id:
                try:
                    await bot.send_message(
                        admin_id,
                        f"🔔 **Новая запись в журнале ЧПУ**\n\n{details}",
                        parse_mode="Markdown"
                    )
                except Exception:
                    pass

    except PermissionError:
        await message.answer(
            "⚠️ **Ошибка записи!** Файл Excel сейчас открыт. Закройте файл и попробуйте снова.",
            reply_markup=main_menu_kb(message.from_user.id),
            parse_mode="Markdown"
        )

    await state.clear()


async def main():
    init_excel()
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher()
    dp.include_router(router)
    print("Бот запущен (полная версия)...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
    
