import io
import re
import zipfile

from aiogram import Router, F
from aiogram.filters import CommandStart, Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import CallbackQuery, Message
from aiogram.types.input_file import BufferedInputFile
from aiogram.utils.keyboard import InlineKeyboardBuilder

from access import AccessManager
from render import (
    Fonts,
    build_xlsx_products_template,
    build_xlsx_tea_template,
    load_rows_products,
    load_rows_tea,
    make_pdf_products_two_sides,
    make_pdf_tea_bank,
    make_pdf_tea_box,
    make_pdf_tips_two_sides,
    safe_filename,
    unique_names,
)

router = Router()

WELCOME_TEXT = (
    "Привет! Я Чайный Ассистент ☕️\n\n"
    "Я умею делать аккуратные PDF-карточки:\n"
    "• 🧾 Ценники для чая (Excel → ZIP с PDF)\n"
    "• 📦Ценники для товаров (Excel → ZIP с двухсторонними PDF)\n"
    "• 💸 Карточки для чаевых (пошаговый ввод → PDF с QR)\n\n"
    "Выбирай действие кнопками ниже 👇"
)

DONE_TEXT = (
    "Готово. Что делаем дальше ?\n"
    "• 🧾 Ценники для чая (Excel → ZIP с PDF)\n"
    "• 📦Ценники для товаров (Excel → ZIP с двухсторонними PDF)\n"
    "• 💸 Карточки для чаевых (пошаговый ввод → PDF с QR)\n\n"
    "Выбирай действие кнопками ниже 👇"
)


class TipsFSM(StatesGroup):
    name = State()
    goal = State()
    link = State()


class WaitFilesFSM(StatesGroup):
    wait_tea_xlsx = State()
    wait_products_xlsx = State()


def main_menu_kb():
    kb = InlineKeyboardBuilder()
    kb.button(text="🧾 Ценники: Чай", callback_data="menu:tea")
    kb.button(text="📦 Ценники: Товары", callback_data="menu:products")
    kb.button(text="💸 Карточка: Чаевые", callback_data="menu:tips")
    kb.adjust(1)
    return kb.as_markup()


def back_cancel_kb(back_cb: str):
    kb = InlineKeyboardBuilder()
    kb.button(text="⬅️ Назад", callback_data=back_cb)
    kb.button(text="⛔️ Отмена", callback_data="menu:cancel")
    kb.adjust(2)
    return kb.as_markup()


async def deny(message: Message):
    await message.answer("⛔️ У тебя нет доступа к этому боту. Напиши администратору.")


def is_xlsx(message: Message) -> bool:
    if not message.document:
        return False
    return (message.document.file_name or "").lower().endswith(".xlsx")


@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext, access: AccessManager):
    if not access.is_allowed(message.from_user.id):
        return await deny(message)
    await state.clear()
    await message.answer(WELCOME_TEXT, reply_markup=main_menu_kb())


@router.callback_query(F.data == "menu:cancel")
async def cb_cancel(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return
    await state.clear()
    await query.message.answer(WELCOME_TEXT, reply_markup=main_menu_kb())
    await query.answer()


# -------------------------
# Админ-команды
# -------------------------
@router.message(Command("add_user"))
async def admin_add_user(message: Message, access: AccessManager):
    if not access.is_admin(message.from_user.id):
        return await deny(message)

    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("Использование: /add_user 123456789")
        return
    uid = int(parts[1])
    access.add_user(uid)
    await message.answer(f"✅ Добавил доступ пользователю {uid}.")


@router.message(Command("del_user"))
async def admin_del_user(message: Message, access: AccessManager):
    if not access.is_admin(message.from_user.id):
        return await deny(message)

    parts = (message.text or "").split()
    if len(parts) != 2 or not parts[1].isdigit():
        await message.answer("Использование: /del_user 123456789")
        return
    uid = int(parts[1])
    access.del_user(uid)
    await message.answer(f"🗑️ Убрал доступ пользователю {uid}.")


@router.message(Command("list_users"))
async def admin_list_users(message: Message, access: AccessManager):
    if not access.is_admin(message.from_user.id):
        return await deny(message)
    users = access.list_users()
    await message.answer("Разрешённые пользователи:\n" + ("\n".join(map(str, users)) if users else "Список пуст."))


# -------------------------
# Чай: Excel → один ZIP (две папки)
# -------------------------
@router.callback_query(F.data == "menu:tea")
async def cb_tea(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return

    xlsx = build_xlsx_tea_template()
    await state.set_state(WaitFilesFSM.wait_tea_xlsx)

    await query.message.answer(
        "🧾 Ценники: Чай\n\n"
        "Заполни Excel и отправь обратно.\n"
        "В ответ пришлю ZIP, внутри две папки:\n"
        "• Ценники для банок\n"
        "• Ценники для коробок"
    )
    await query.message.answer_document(BufferedInputFile(xlsx, filename="tea_template.xlsx"))
    await query.answer()


@router.message(WaitFilesFSM.wait_tea_xlsx)
async def tea_receive_xlsx(message: Message, state: FSMContext, access: AccessManager, fonts: Fonts):
    if not access.is_allowed(message.from_user.id):
        return await deny(message)

    if not is_xlsx(message):
        await message.answer("Пришли файл .xlsx (Excel).")
        return

    file = await message.bot.get_file(message.document.file_id)
    fb = await message.bot.download_file(file.file_path)
    xlsx_bytes = fb.read()

    try:
        rows = load_rows_tea(xlsx_bytes)
    except Exception as e:
        await message.answer(f"Ошибка в Excel: {e}")
        return

    await message.answer(f"⏳ Генерирую PDF… строк: {len(rows)}")
    base_names = unique_names([safe_filename(r[1]) for r in rows])

    out_zip = io.BytesIO()
    with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for (tea_type, name, price), fname in zip(rows, base_names):
            z.writestr(f"Ценники для банок/{fname}.pdf", make_pdf_tea_bank(fonts, tea_type, name, price))
            z.writestr(f"Ценники для коробок/{fname}.pdf", make_pdf_tea_box(fonts, tea_type, name, price))

    out_zip.seek(0)
    await message.answer_document(BufferedInputFile(out_zip.read(), filename="Ценники Чай.zip"))

    await state.clear()
    await message.answer(DONE_TEXT, reply_markup=main_menu_kb())


# -------------------------
# Товары: Excel → ZIP (двухсторонние PDF)
# -------------------------
@router.callback_query(F.data == "menu:products")
async def cb_products(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return

    xlsx = build_xlsx_products_template()
    await state.set_state(WaitFilesFSM.wait_products_xlsx)

    await query.message.answer(
        "📦 Ценники: Товары\n\n"
        "Заполни Excel и отправь обратно.\n"
        "В ответ пришлю ZIP, каждый PDF будет 2 страницы (перед/зад)."
    )
    await query.message.answer_document(BufferedInputFile(xlsx, filename="products_template.xlsx"))
    await query.answer()


@router.message(WaitFilesFSM.wait_products_xlsx)
async def products_receive_xlsx(message: Message, state: FSMContext, access: AccessManager, fonts: Fonts):
    if not access.is_allowed(message.from_user.id):
        return await deny(message)

    if not is_xlsx(message):
        await message.answer("Пришли файл .xlsx (Excel).")
        return

    file = await message.bot.get_file(message.document.file_id)
    fb = await message.bot.download_file(file.file_path)
    xlsx_bytes = fb.read()

    try:
        rows = load_rows_products(xlsx_bytes)
    except Exception as e:
        await message.answer(f"Ошибка в Excel: {e}")
        return

    await message.answer(f"⏳ Генерирую двухсторонние PDF… строк: {len(rows)}")
    base_names = unique_names([safe_filename(r[0]) for r in rows])

    out_zip = io.BytesIO()
    with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
        for (name, price, hours), fname in zip(rows, base_names):
            z.writestr(f"{fname}.pdf", make_pdf_products_two_sides(fonts, name, price, hours))

    out_zip.seek(0)
    await message.answer_document(BufferedInputFile(out_zip.read(), filename="Ценники Товары.zip"))

    await state.clear()
    await message.answer(DONE_TEXT, reply_markup=main_menu_kb())


# -------------------------
# Чаевые: пошагово → PDF (2 страницы)
# -------------------------
@router.callback_query(F.data == "menu:tips")
async def cb_tips(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return

    await state.clear()
    await state.set_state(TipsFSM.name)
    await query.message.answer("💸 Введи имя (как на карточке):", reply_markup=back_cancel_kb("tips:back_name"))
    await query.answer()


@router.callback_query(F.data == "tips:back_name")
async def tips_back_name(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return
    await state.set_state(TipsFSM.name)
    await query.message.answer("💸 Введи имя (как на карточке):", reply_markup=back_cancel_kb("tips:back_name"))
    await query.answer()


@router.callback_query(F.data == "tips:back_goal")
async def tips_back_goal(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return
    await state.set_state(TipsFSM.goal)
    await query.message.answer("🎯 Введи цель (на что копишь):", reply_markup=back_cancel_kb("tips:back_name"))
    await query.answer()


@router.callback_query(F.data == "tips:back_link")
async def tips_back_link(query: CallbackQuery, state: FSMContext, access: AccessManager):
    if not access.is_allowed(query.from_user.id):
        await query.answer("Нет доступа", show_alert=True)
        return
    await state.set_state(TipsFSM.link)
    await query.message.answer("🔗 Вставь ссылку Netmonet (для QR):", reply_markup=back_cancel_kb("tips:back_goal"))
    await query.answer()


@router.message(TipsFSM.name)
async def tips_name(message: Message, state: FSMContext, access: AccessManager):
    if not access.is_allowed(message.from_user.id):
        return await deny(message)

    text = (message.text or "").strip()
    if not text or len(text) > 40:
        await message.answer("Имя должно быть 1–40 символов. Попробуй ещё раз:", reply_markup=back_cancel_kb("tips:back_name"))
        return

    await state.update_data(tips_name=text)
    await state.set_state(TipsFSM.goal)
    await message.answer("🎯 Введи цель (на что копишь):", reply_markup=back_cancel_kb("tips:back_name"))


@router.message(TipsFSM.goal)
async def tips_goal(message: Message, state: FSMContext, access: AccessManager):
    if not access.is_allowed(message.from_user.id):
        return await deny(message)

    text = (message.text or "").strip()
    if not text or len(text) > 80:
        await message.answer("Цель должна быть 1–80 символов. Попробуй ещё раз:", reply_markup=back_cancel_kb("tips:back_goal"))
        return

    await state.update_data(tips_goal=text)
    await state.set_state(TipsFSM.link)
    await message.answer("🔗 Вставь ссылку Netmonet (для QR):", reply_markup=back_cancel_kb("tips:back_goal"))


@router.message(TipsFSM.link)
async def tips_link(message: Message, state: FSMContext, access: AccessManager, fonts: Fonts):
    if not access.is_allowed(message.from_user.id):
        return await deny(message)

    link = (message.text or "").strip()
    if not link or len(link) > 300:
        await message.answer("Ссылка выглядит странно. Вставь корректную ссылку Netmonet:", reply_markup=back_cancel_kb("tips:back_link"))
        return
    if not re.match(r"^https?://", link, flags=re.I):
        await message.answer("Ссылка должна начинаться с http:// или https://", reply_markup=back_cancel_kb("tips:back_link"))
        return

    data = await state.get_data()
    person_name = data.get("tips_name", "Имя")
    goal = data.get("tips_goal", "Цель")

    await message.answer("⏳ Генерирую карточку чаевых…")
    pdf_bytes = make_pdf_tips_two_sides(fonts, person_name, goal, link)

    filename = safe_filename(f"Чаевые_{person_name}") + ".pdf"
    await message.answer_document(BufferedInputFile(pdf_bytes, filename=filename))

    await state.clear()
    await message.answer(DONE_TEXT, reply_markup=main_menu_kb())
