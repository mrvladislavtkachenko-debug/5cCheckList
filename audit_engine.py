# -*- coding: utf-8 -*-
"""
Движок прохождения аудита: линейная навигация по пунктам чек-листа,
хранение ответов (0/1) и формирование итогового отчёта.

Не зависит от aiogram — логику удобно покрывать тестами.
Тексты отдаются в HTML-разметке (для aiogram parse_mode="HTML").
"""
from datetime import datetime

from checklist import STEPS, TOTAL_CRITERIA, interpret_total


def escape(text: str) -> str:
    """Экранирование HTML-спецсимволов для подстановки в сообщения."""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


# Плоский список пунктов: каждая позиция = (номер шага, номер пункта в шаге).
FLAT = []
for _si, _step in enumerate(STEPS):
    for _ci in range(len(_step["criteria"])):
        FLAT.append((_si, _ci))

# Границы шагов в плоском списке.
STEP_BOUNDS = []
_acc = 0
for _step in STEPS:
    STEP_BOUNDS.append((_acc, _acc + len(_step["criteria"])))
    _acc += len(_step["criteria"])


def is_step_first(pos: int) -> bool:
    """Позиция является первым вопросом своего шага."""
    for start, _end in STEP_BOUNDS:
        if pos == start:
            return True
    return False


class Audit:
    """Один сеанс аудита одного объекта (участка)."""

    def __init__(self, object_name: str):
        self.object_name = (object_name or "").strip()
        self.answers: dict[int, bool] = {}  # pos -> True(выполнено)/False
        self.pos = 0                        # индекс текущего отображаемого пункта
        self.started_at = datetime.now()

    @property
    def finished(self) -> bool:
        return self.pos >= TOTAL_CRITERIA

    # --- навигация -------------------------------------------------------
    def _next_unanswered_after(self, start: int) -> int:
        p = start + 1
        while p < TOTAL_CRITERIA and p in self.answers:
            p += 1
        return p

    def answer(self, done: bool) -> None:
        """Зафиксировать ответ на текущем пункте и перейти вперёд."""
        if self.finished:
            return
        self.answers[self.pos] = done
        self.pos = self._next_unanswered_after(self.pos)

    def back(self) -> None:
        """Вернуться на один пункт назад (можно пересмотреть ответ)."""
        if self.pos > 0:
            self.pos -= 1

    def has_answer(self, pos: int) -> bool:
        return pos in self.answers

    # --- отображение пункта ----------------------------------------------
    def question_text(self, pos: int) -> str:
        si, ci = FLAT[pos]
        step = STEPS[si]
        start, end = STEP_BOUNDS[si]
        num_in_step = pos - start + 1
        total_in_step = end - start

        parts = [
            f"🗓 <b>Шаг {si + 1} из {len(STEPS)} · {step['emoji']} {escape(step['name'])}</b>",
            f"<i>{escape(step['subtitle'])}</i>",
            "",
        ]
        if is_step_first(pos):
            parts += [f"<i>{escape(step['description'])}</i>", ""]
        parts += [
            f"📌 <b>Пункт {num_in_step}/{total_in_step}:</b>",
            escape(step["criteria"][ci]),
            "",
            "<b>Оценка:</b> 1 — требование выполнено, 0 — не выполнено.",
        ]
        return "\n".join(parts)

    def question_markup_label(self, pos: int) -> str:
        si, ci = FLAT[pos]
        step = STEPS[si]
        start, end = STEP_BOUNDS[si]
        return f"{step['key']} · {pos - start + 1}/{end - start}"

    # --- отчёт ------------------------------------------------------------
    @staticmethod
    def _bar(percent: int, width: int = 10) -> str:
        filled = round(percent / 100 * width)
        return "█" * filled + "░" * (width - filled)

    def report(self) -> str:
        if not self.answers:
            return "Нет данных для отчёта."

        lines = []
        lines.append("🧾 <b>Отчёт аудита 5С — овощной участок</b>")
        lines.append("")
        lines.append(f"📍 <b>Объект:</b> {escape(self.object_name or 'не указан')}")
        lines.append(f"🗓 <b>Дата:</b> {self.started_at:%d.%m.%Y %H:%M}")
        lines.append("")

        lines.append("<b>Степень реализации шагов:</b>")
        total_done = 0
        for si, step in enumerate(STEPS):
            start, end = STEP_BOUNDS[si]
            done = sum(1 for p in range(start, end) if self.answers.get(p))
            total_in_step = end - start
            total_done += done
            pct = round(done / total_in_step * 100) if total_in_step else 0
            lines.append(
                f"{step['emoji']} {escape(step['name'])}: "
                f"<code>{done}/{total_in_step}</code> · <code>{pct}%</code> "
                f"{self._bar(pct)}"
            )

        overall = round(total_done / TOTAL_CRITERIA * 100)
        lines.append("")
        lines.append(
            f"⭐ <b>Итог:</b> <code>{total_done}/{TOTAL_CRITERIA}</code> "
            f"пунктов — <b>{overall}%</b>"
        )
        lines.append(self._bar(overall, 20))
        lines.append("")
        lines.append(escape(interpret_total(overall)))

        # Невыполненные пункты
        unmet = []
        for si, step in enumerate(STEPS):
            start, end = STEP_BOUNDS[si]
            for p in range(start, end):
                if p in self.answers and not self.answers[p]:
                    unmet.append((si, p - start))
        if unmet:
            lines.append("")
            lines.append("<b>Не выполнено (над чем работать):</b>")
            for si, ci in unmet:
                lines.append(
                    f"• {STEPS[si]['emoji']} {STEPS[si]['key']}: "
                    f"{escape(STEPS[si]['criteria'][ci])}"
                )
        else:
            lines.append("")
            lines.append("🎉 Все пункты выполнены — замечаний нет!")

        return "\n".join(lines)
