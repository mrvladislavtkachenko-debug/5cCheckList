# -*- coding: utf-8 -*-
"""
Telegram-бот для аудита 5С на овощном участке.

Запуск:
    export BOT_TOKEN="<токен от @BotFather>"      # или в .env
    python bot.py

Команды:
    /start — приветствие и кнопка «Начать аудит»
    /help  — краткая справка
    /cancel — отменить текущий аудит
"""
import os
import logging

from aiogram import Bot, Dispatcher, Router, F
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup, default_state
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    Message,
)

from audit_engine import Audit, TOTAL_CRITERIA
from checklist import STEPS, steps_summary

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# ----------------------------------------------------------------------------
# Состояния диалога (FSM)
# ----------------------------------------------------------------------------
class AuditStates(StatesGroup):
    naming = State()      # ожидаем название объекта/участка
    answering = State()   # идём по чек-листу


# ----------------------------------------------------------------------------
# Клавиатуры
# ----------------------------------------------------------------------------
def main_menu_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🧾 Начать аудит 5С",
                                  callback_data="start_audit")],
        ]
    )
    return kb


def answer_kb() -> InlineKeyboardMarkup:
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Выполнено (1)",
                                     callback_data="ans:1"),
                InlineKeyboardButton(text="❌ Не выполнено (0)",
                                     callback_data="ans:0"),
            ],
            [
                InlineKeyboardButton(text="◀️ Назад",
                                     callback_data="back"),
                InlineKeyboardButton(text="🚫 Отменить",
                                     callback_data="cancel"),
            ],
        ]
    )
    return kb


def cancel_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🚫 Отменить",
                                  callback_data="cancel")]
        ]
    )


# ----------------------------------------------------------------------------
# Тексты
# ----------------------------------------------------------------------------
HELP_TEXT = (
    "<b>Как пользоваться ботом аудита 5С</b>\n\n"
    "Бот проводит аудит участка по системе <b>5С</b>:\n"
    "1. 🗂️ Сортировка — убери лишнее\n"
    "2. 📦 Соблюдение порядка — всему своё место\n"
    "3. 🧹 Содержание в чистоте — убирай регулярно\n"
    "4. 📋 Стандартизация — закрепи правила\n"
    "5. 🔄 Совершенствование — улучшай постоянно\n\n"
    "Проверьте каждый пункт чек-листа и нажимайте "
    "«✅ Выполнено» или «❌ Не выполнено».\n"
    "В конце бот покажет результат в % по каждому шагу "
    "и перечень невыполненных пунктов.\n\n"
    "Команды: /start — главное меню, /cancel — отменить аудит."
)

ABOUT_TEXT = (
    "🧾 <b>Аудит 5С на овощном участке</b>\n\n"
    "Чек-лист адаптирован по методическим указаниям "
    "Федерального центра компетенций\n"
    "<i>МУ-52-2024 «5С в производстве» (редакция № 3, Приложение № 8)</i>\n"
    "под реалии овощного участка, огорода или теплицы.\n\n"
    "Каждый пункт оценивается бинарно:\n"
    "<code>1</code> — требование выполнено,\n"
    "<code>0</code> — не выполнено.\n\n"
    "Всего критериев в чек-листе: <b>{total}</b>"
).format(total=TOTAL_CRITERIA)


def steps_list_text() -> str:
    lines = ["<b>Шаги аудита:</b>", ""]
    for i, s in enumerate(STEPS, 1):
        lines.append(f"{s['emoji']} <b>{i}. {s['name']}</b> — {s['subtitle']}")
    return "\n".join(lines)


def welcome_text() -> str:
    total = TOTAL_CRITERIA
    lines = [
        "🥕 <b>Добро пожаловать! Это бот аудита 5С на овощном участке.</b>",
        "",
        "Бот проведёт вас по чек-листу из {t} вопросов по системе 5С и "
        "составит отчёт о состоянии участка.".format(t=total),
        "",
        steps_list_text(),
        "",
        "Нажмите «Начать аудит», назовите проверяемый объект "
        "(например «Теплица №1» или «Грядки у склада») и отвечайте "
        "на вопросы кнопками.",
    ]
    return "\n".join(lines)


# ----------------------------------------------------------------------------
# Приложение
# ----------------------------------------------------------------------------
class Application:
    """Обёртка, связывающая aiogram-сущности. Позволяет тестировать."""

    def __init__(self, token: str):
        self.bot = Bot(
            token=token,
            default=DefaultBotProperties(parse_mode=ParseMode.HTML),
        )
        self.storage = MemoryStorage()
        self.dp = Dispatcher(storage=self.storage)
        self.dp.message(CommandStart())(self.cmd_start)
        self.dp.message(Command("help"))(self.cmd_help)
        self.dp.message(Command("cancel"))(self.cmd_cancel)
        self.dp.callback_query(F.data == "start_audit")(self.on_start_audit)
        self.dp.callback_query(F.data == "cancel")(self.on_cancel)
        self.dp.callback_query(F.data == "back")(self.on_back)
        self.dp.callback_query(F.data.startswith("ans:"))(self.on_answer)
        # Текстовое сообщение в состоянии naming
        self.dp.message(AuditStates.naming, F.text)(self.on_object_name)
        # Всё, что не кнопка, во время answering — напоминание
        self.dp.message(AuditStates.answering)(self.on_stray_message)
        self.dp.message(default_state)(self.on_unhandled_message)

    # --- публичный API ----------------------------------------------------
    async def _delete(self, message: Message) -> None:
        try:
            await message.delete()
        except Exception:  # нет прав или сообщение уже удалено — не страшно
            pass

    async def _send_question(self, message: Message, audit: Audit,
                             state: FSMContext) -> None:
        txt = audit.question_text(audit.pos)
        await state.update_data(msg_id=message.message_id)
        await message.answer(txt, reply_markup=answer_kb())

    # --- обработчики команд ------------------------------------------------
    async def cmd_start(self, message: Message, state: FSMContext) -> None:
        await state.clear()
        await message.answer(welcome_text(), reply_markup=main_menu_kb())

    async def cmd_help(self, message: Message) -> None:
        await message.answer(HELP_TEXT)

    async def cmd_cancel(self, message: Message, state: FSMContext) -> None:
        await self._cancel_state(message, state)

    async def _cancel_state(self, message: Message, state: FSMContext) -> None:
        if state is not None:
            await state.clear()
        await message.answer(
            "🚫 Аудит отменён.\n\nВы можете начать новый.",
            reply_markup=main_menu_kb(),
        )

    # --- обработчики кнопок -------------------------------------------------
    async def on_start_audit(self, callback: CallbackQuery,
                             state: FSMContext) -> None:
        await callback.answer()
        await state.clear()
        await state.set_state(AuditStates.naming)
        await callback.message.answer(
            "🏡 <b>Укажите название проверяемого объекта</b> "
            "(участок, грядка, теплица, парник и т. п.):\n\n"
            "Например: «Теплица №1», «Овощной участок за складом».",
            reply_markup=cancel_kb(),
        )

    async def on_object_name(self, message: Message,
                             state: FSMContext) -> None:
        name = message.text.strip()
        if not name:
            return
        if message.message_id:
            await self._delete(message)
        audit = Audit(object_name=name)
        await state.set_state(AuditStates.answering)
        await state.update_data(audit=audit)
        await message.answer(
            audit.question_text(audit.pos),
            reply_markup=answer_kb(),
        )

    async def on_answer(self, callback: CallbackQuery,
                        state: FSMContext) -> None:
        await callback.answer()
        data = await state.get_data()
        audit: Audit = data.get("audit")
        if audit is None:
            await state.clear()
            await callback.message.answer("Аудит уже завершён.",
                                          reply_markup=main_menu_kb())
            return
        value = callback.data.split(":", 1)[1]  # "1" или "0"
        done = value == "1"
        audit.answer(done)
        pos = audit.pos - 1  # тот вопрос, на который только что ответили
        # Убираем кнопки с уже отвеченного вопроса и помечаем выбор,
        # чтобы не нажимать старые кнопки повторно.
        try:
            await callback.message.edit_text(
                audit.question_text(pos) + "\n\n" +
                ("→ ✅ Отмечено: выполнено (1)" if done
                 else "→ ❌ Отмечено: не выполнено (0)"),
                reply_markup=None,
            )
        except Exception:
            pass

        if audit.finished:
            report = audit.report()
            await state.clear()
            await callback.message.answer(
                report + "\n\nХотите провести ещё один аудит?",
                reply_markup=main_menu_kb(),
            )
        else:
            await callback.message.answer(
                audit.question_text(audit.pos),
                reply_markup=answer_kb(),
            )

    async def on_back(self, callback: CallbackQuery,
                      state: FSMContext) -> None:
        await callback.answer()
        data = await state.get_data()
        audit: Audit = data.get("audit")
        if audit is None or audit.pos == 0:
            return
        audit.back()
        await self._delete(callback.message)
        await callback.message.answer(
            audit.question_text(audit.pos),
            reply_markup=answer_kb(),
        )

    async def on_cancel(self, callback: CallbackQuery,
                        state: FSMContext) -> None:
        await callback.answer()
        await self._cancel_state(callback.message, state)

    async def on_stray_message(self, message: Message,
                               state: FSMContext) -> None:
        """Во время прохождения чек-листа текст игнорируем — нужны кнопки."""
        data = await state.get_data()
        audit: Audit = data.get("audit")
        if audit is None:
            return
        await message.answer(
            "Используйте кнопки под вопросом: «✅ Выполнено» "
            "или «❌ Не выполнено». Вопрос ниже:",
        )
        await message.answer(
            audit.question_text(audit.pos),
            reply_markup=answer_kb(),
        )

    async def on_unhandled_message(self, message: Message) -> None:
        await message.answer("Отправьте /start, чтобы начать.",
                             reply_markup=main_menu_kb())

    # --- запуск поллинга ---------------------------------------------------
    async def run_polling(self) -> None:
        try:
            await self.dp.start_polling(self.bot)
        finally:
            await self.bot.session.close()


# ----------------------------------------------------------------------------
# Точка входа
# ----------------------------------------------------------------------------
def load_token() -> str:
    token = os.environ.get("BOT_TOKEN", "").strip()
    if not token:
        try:
            from dotenv import load_dotenv
            load_dotenv()
            token = os.environ.get("BOT_TOKEN", "").strip()
        except Exception:
            pass
    return token


def main() -> None:
    token = load_token()
    if not token:
        print(
            "Не задан токен бота.\n"
            "Получите токен у @BotFather и задайте переменную окружения "
            "BOT_TOKEN, например:\n"
            "    export BOT_TOKEN='123456:ABC...'\n"
            "или создайте файл .env со строкой BOT_TOKEN='123456:ABC...'"
        )
        raise SystemExit(1)

    app = Application(token)
    logger.info("Бот запущен (polling). Нажмите Ctrl+C для остановки.")
    import asyncio
    try:
        asyncio.run(app.run_polling())
    except KeyboardInterrupt:
        logger.info("Остановлено.")


if __name__ == "__main__":
    main()
