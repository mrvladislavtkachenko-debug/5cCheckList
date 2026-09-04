# -*- coding: utf-8 -*-
"""Тесты движка аудита (не требуют Telegram/сети)."""
import unittest

from audit_engine import (
    Audit,
    FLAT,
    STEP_BOUNDS,
    TOTAL_CRITERIA,
    is_step_first,
)
from checklist import STEPS, steps_summary


class AuditFlowTest(unittest.TestCase):
    def setUp(self):
        self.audit = Audit("Теплица №1")

    def test_counts(self):
        expected_total = sum(len(s["criteria"]) for s in STEPS)
        self.assertEqual(TOTAL_CRITERIA, expected_total)
        self.assertEqual(len(FLAT), TOTAL_CRITERIA)
        # ровно 5 шагов
        self.assertEqual(len(STEPS), 5)

    def test_first_questions_are_step_openers(self):
        self.assertTrue(is_step_first(0))
        for _, end in STEP_BOUNDS:
            if end < TOTAL_CRITERIA:
                self.assertTrue(is_step_first(end))  # следующий шаг начинается
        # внутренние пункты — не «открывающие»
        if len(STEP_BOUNDS[0]) > 0:
            start, _ = STEP_BOUNDS[0]
            self.assertFalse(is_step_first(start + 1))

    def test_full_run_all_done(self):
        while not self.audit.finished:
            self.audit.answer(True)
        self.assertTrue(self.audit.finished)
        self.assertEqual(len(self.audit.answers), TOTAL_CRITERIA)
        self.assertTrue(all(self.audit.answers.values()))
        report = self.audit.report()
        self.assertIn("100%", report)
        self.assertIn("замечаний нет", report)
        self.assertIn("Итог", report)

    def test_full_run_all_not_done(self):
        while not self.audit.finished:
            self.audit.answer(False)
        report = self.audit.report()
        self.assertIn("0%", report)
        self.assertIn("Не выполнено", report)

    def test_back_navigation(self):
        self.audit.answer(True)   # pos 0
        self.audit.answer(True)   # pos 1
        self.audit.back()         # -> pos 1
        self.audit.answer(False)  # пересматриваем pos 1
        self.assertEqual(self.audit.answers[1], False)

    def test_skip_answered_forward(self):
        # ответим на 1-й и 2-й, потом вернёмся и убедимся, что вперёд идём
        # на следующий неотвеченный
        self.audit.answer(True)  # 0
        self.audit.answer(True)  # 1
        self.audit.back()
        self.audit.answer(True)  # переответ на 1 -> переходит на следующий (2)
        self.assertNotIn(1, {})  # placeholder
        self.assertEqual(self.audit.pos, 2)
        self.assertEqual(len(self.audit.answers), 2)


class StepsSummaryTest(unittest.TestCase):
    def test_five_steps_and_keys(self):
        summary = steps_summary()
        self.assertEqual(len(summary), 5)
        keys = [k for k, *_ in summary]
        self.assertEqual(keys, ["1C", "2C", "3C", "4C", "5C"])


if __name__ == "__main__":
    unittest.main()
