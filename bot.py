import asyncio
import json
import os
from datetime import datetime
from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove, FSInputFile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ================= НАСТРОЙКИ =================
BOT_TOKEN = "9185160463:AAGFZwG1ylAulfZdgj8LcwKfvVeeb5SH-jw"
ACCESS_PASSWORD = "1234"  # Пароль для входа операторов
ADMIN_IDS = [1011177377]   # Впишите сюда ваш Telegram ID (можно несколько через запятую)

EXCEL_FILE = "cnc_work_log.xlsx"
USERS_FILE = "operators.json"

# Список ваших станков (кнопки)
MACHINES = [
    ["Deg", "DAS"],
    ["MICRON", "DMG"],
    ["Эрозионный"]
]

# Список операций (кнопки)
OPERATIONS = [
    ["Наладка", "Черновая"],
    ["Чистовая", "Сверление / Резьба"],
    ["Контроль / Доработка"]
]
# =============================================

router = Router()


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


# --- Работа с Excel ---
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Журнал работ"
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

        wb.save(EXCEL_FILE)


def append_to_excel(row_data: list):
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append(row_data)
    wb.save(EXCEL_FILE)


# --- Клавиатуры ---
def main_menu_kb(user_id: int):
    buttons = [[KeyboardButton(text="➕ Внести работу")]]
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
@router.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    users = load_users()
    user_id = str(message.from_user.id)

    if user_id in users:
        await message.answer(
            f"Добро пожаловать, **{users[user_id]}**!\nНажмите «Внести работу», чтобы записать смену.",
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
    if message.text.strip() == ACCESS_PASSWORD:
        await message.answer("Пароль верный. Введите ваши **Фамилию и Имя** (например: *Иванов А.С.*):", parse_mode="Markdown")
        await state.set_state(WorkForm.waiting_fio)
    else:
        await message.answer("Неверный пароль. Попробуйте ещё раз:")


@router.message(WorkForm.waiting_fio)
async def process_fio(message: Message, state: FSMContext):
    fio = message.text.strip()
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
    await message.answer("Введите новые Фамилию и Имя:", reply_markup=ReplyKeyboardRemove())
    await state.set_state(WorkForm.waiting_fio)


@router.message(F.text == "❌ Отмена")
async def cancel_action(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Действие отменено.", reply_markup=main_menu_kb(message.from_user.id))


@router.message(F.text == "📊 Скачать Excel-отчёт")
async def send_excel(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("⛔ У вас нет прав для скачивания общего журнала.")
        return

    init_excel()
    doc = FSInputFile(EXCEL_FILE)
    await message.answer_document(doc, caption=f"Актуальный журнал работ на {datetime.now().strftime('%d.%m.%Y %H:%M')}")


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
        keyboard=MACHINES + [[KeyboardButton(text="❌ Отмена")]],
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
        keyboard=OPERATIONS + [[KeyboardButton(text="❌ Отмена")]],
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
    text = message.text.strip()
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
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

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

        # Рассылка уведомления администраторам
        for admin_id in ADMIN_IDS:
            if admin_id != message.from_user.id:
                try:
                    await bot.send_message(
                        admin_id,
                        f"🔔 **Новая запись в журнале ЧПУ**\n\n{details}",
                        parse_mode="Markdown"
                    )
                except Exception:
                    pass  # Если админ ещё не нажал /start в боте

    except PermissionError:
        await message.answer(
            "⚠️ **Ошибка записи!** Файл Excel сейчас открыт на компьютере. Закройте файл и попробуйте снова.",
            reply_markup=main_menu_kb(message.from_user.id),
            parse_mode="Markdown"
        )

    await state.clear()


async def main():
    init_excel()
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher()
    dp.include_router(router)
    print("Бот запущен...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
      
