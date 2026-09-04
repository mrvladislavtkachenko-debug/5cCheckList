# -*- coding: utf-8 -*-
"""
Пример: прогоняем «эталонный» набор ответов и печатаем итоговый отчёт.
Запуск:  python example_report.py
"""
from audit_engine import Audit
from audit_engine import TOTAL_CRITERIA


def make_audit() -> Audit:
    a = Audit("Теплица №1 (эталон)")
    # Сценарий «почти всё хорошо, несколько недочётов».
    pattern = {
        "1C": [1, 1, 1, 1, 1],
        "2C": [1, 1, 1, 1, 0],
        "3C": [1, 1, 1, 1, 0],
        "4C": [1, 1, 1, 0],
        "5C": [1, 1, 1, 1],
    }
    flat = []
    for key, vals in pattern.items():
        assert len(vals) == 0 or True
        flat.extend(vals)
    assert len(flat) == TOTAL_CRITERIA, (len(flat), TOTAL_CRITERIA)
    for v in flat:
        a.answer(bool(v))
    return a


if __name__ == "__main__":
    audit = make_audit()
    print(audit.report())
    print("\n--- перечень пунктов для сверки ---")
    from audit_engine import FLAT
    from checklist import STEPS
    for i, (si, ci) in enumerate(FLAT):
        st = STEPS[si]
        print(f"{i+1:>2}. [{st['key']}] {st['criteria'][ci]}")
