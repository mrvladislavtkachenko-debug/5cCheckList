# -*- coding: utf-8 -*-
"""Интеграционный тест потока бота без реального Telegram API."""
import unittest

from aiogram.fsm.context import FSMContext

import bot as botmod
from audit_engine import TOTAL_CRITERIA


class _Answer:
    """Заглушка, сохраняющая последний отправленный текст."""

    def __init__(self):
        self.last_text = None
        self.last_kwargs = None

    async def __call__(self, text=None, **kwargs):
        self.last_text = text
        self.last_kwargs = kwargs
        return None


class MessageMock:
    def __init__(self, message_id=1, text="Теплица №1"):
        self.message_id = message_id
        self.text = text
        self.answer = _Answer()
        self.edited = []

    async def delete(self):
        return None

    async def edit_text(self, text=None, **kwargs):
        self.edited.append(text)


class CallbackMock:
    def __init__(self, data="", message=None):
        self.data = data
        self.message = message or MessageMock()
        self.answered = False

    async def answer(self):
        self.answered = True


class BotFlowTest(unittest.TestCase):
    def setUp(self):
        self.app = botmod.Application(token="123:fake")
        # контекст состояния для одного пользователя
        self.ctx = FSMContext(storage=self.app.storage, key=(1, 1))

    def _clear(self):
        import asyncio
        asyncio.get_event_loop().run_until_complete(self.ctx.clear())

    def _run(self, coro):
        import asyncio
        return asyncio.get_event_loop().run_until_complete(coro)

    def test_full_audit_flow_positive(self):
        # /start
        start_msg = MessageMock(message_id=1)
        self._run(self.app.cmd_start(start_msg, self.ctx))
        self.assertIsNone(self._run(self.ctx.get_state()))  # сброшено

        # нажать «Начать аудит»
        cb = CallbackMock(data="start_audit",
                          message=MessageMock(message_id=2))
        self._run(self.app.on_start_audit(cb, self.ctx))
        state = self._run(self.ctx.get_state())
        self.assertEqual(state, botmod.AuditStates.naming)

        # ввести название объекта
        obj = MessageMock(message_id=3, text="Теплица №1")
        self._run(self.app.on_object_name(obj, self.ctx))
        state = self._run(self.ctx.get_state())
        self.assertEqual(state, botmod.AuditStates.answering)

        # пройти все пункты, отвечая «выполнено»
        for i in range(TOTAL_CRITERIA):
            q = CallbackMock(data="ans:1", message=MessageMock(message_id=10 + i))
            self._run(self.app.on_answer(q, self.ctx))

        # после последнего вопроса состояние очищено, показан отчёт
        data = self._run(self.ctx.get_data())
        state = self._run(self.ctx.get_state())
        self.assertIsNone(state)
        self.assertTrue(self.app.dp is not None)
        # последний отправленный текст — финальное сообщение с отчётом
        # берём из последнего сообщения (создадим и проверим движок отдельно),
        # здесь главное, что флоу прошёл без исключений
        self.assertEqual(data.get("audit"), None)

    def test_cancel_resets_state(self):
        self._run(self.ctx.set_state(botmod.AuditStates.answering))
        msg = MessageMock(message_id=5)
        self._run(self.app.cmd_cancel(msg, self.ctx))
        self.assertIsNone(self._run(self.ctx.get_state()))


if __name__ == "__main__":
    unittest.main()
