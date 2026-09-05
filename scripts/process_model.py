"""Independent reference model. All generated observations are explicitly synthetic."""
from __future__ import annotations
import json
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / 'data/process.json').read_text())
PRODUCTS = DATA['products']
ELEMENTS = DATA['elements']
BLOCK = 21
RAW_LAST = 7 + (len(PRODUCTS) - 1) * BLOCK + 18
CAT_FIRST, CAT_LAST, CAT_TOTAL = 8, 25, 27


def block_row(product_index: int) -> int:
    return 7 + product_index * BLOCK


def auto_seconds(program: int) -> float:
    p = DATA['programs'][program - 1]
    return p['wash_s'] + p['rinse_s'] + p['spin_s']


def samples(product_index: int, target: bool = False) -> list[list[float]]:
    """Element durations, NOT cumulative stopwatch readings (seconds)."""
    p = PRODUCTS[product_index]
    result = []
    # Minima occur in different cycles; hence a nonzero regulating time.
    variations = [0, 0.04, 0.08, 0.02, 0.06, 0.10, 0.03, 0.07, 0.01, 0.05]
    for i, e in enumerate(ELEMENTS):
        if e['kind'] == 'Авто/ожидание':
            result.append([auto_seconds(p['program'])] * 10)
            continue
        base = p['prep_s'] if i == 1 else e['base_s']
        if target:
            base *= e['target_factor']
        result.append([round(base * (1 + variations[(j + 3 * i + product_index) % 10]), 2)
                       for j in range(10)])
    return result


def regulate(values: list[list[float]], valid: list[bool] | None = None) -> dict:
    valid = valid or [True] * len(values[0])
    included = [j for j, ok in enumerate(valid) if ok]
    if not included:
        raise ValueError('No complete valid cycles')
    minimum = [min(row[j] for j in included) for row in values]
    maximum = [max(row[j] for j in included) for row in values]
    cycles = [sum(row[j] for row in values) for j in included]
    t_min = min(cycles)
    delta = max(0, t_min - sum(minimum))
    eligible = [i for i, e in enumerate(ELEMENTS) if e['kind'] == 'Ручная']
    span = [hi - lo for lo, hi in zip(minimum, maximum)]
    denominator = sum(span[i] for i in eligible)
    adjustments = [0.0] * len(values)
    if denominator:
        for i in eligible:
            adjustments[i] = delta * span[i] / denominator
    elif delta:
        adjustments[max(eligible, key=lambda i: minimum[i])] = delta
    adjusted = [lo + reg for lo, reg in zip(minimum, adjustments)]
    assert math.isclose(sum(adjusted), t_min, abs_tol=1e-8)
    return dict(minimum=minimum, maximum=maximum, adjustments=adjustments,
                adjusted=adjusted, cycles=cycles, t_min=t_min,
                t_max=max(cycles), mean=sum(cycles) / len(cycles),
                residual=[m - a for m, a in zip(maximum, adjusted)])


def cleaning_events(cycles: int, initial_counter: int = 0, closing: bool = True) -> tuple[int, int]:
    if not 0 <= initial_counter <= 3 or cycles < 0:
        raise ValueError('Invalid counter or cycle count')
    regular = (cycles + initial_counter) // 4 if cycles else 0
    final = int(cycles > 0 and closing and (cycles + initial_counter) % 4 != 0)
    return regular, final


def reference_result(target=False, shift_minutes=480, breaks=(30, 20, 10, 10),
                     quantities=None, yields_wash=None, batches=None) -> dict:
    q = quantities or [p['demo_batches'] * p['max_kg'] for p in PRODUCTS]
    y = yields_wash or [1.0] * len(PRODUCTS)
    b = batches or [p['max_kg'] for p in PRODUCTS]
    n = [math.ceil(round(qi / yi / bi, 10)) for qi, yi, bi in zip(q, y, b)]
    count = sum(n)
    available = (shift_minutes - sum(breaks)) * 60
    regs = [regulate(samples(i, target)) for i in range(len(PRODUCTS))]
    before_or_target = 'target' if target else 'before'
    cleaning_time = sum(min(x[before_or_target]) for x in DATA['periodic'][:5])
    cleaning_count = sum(cleaning_events(count))
    supply_time = min(DATA['periodic'][5][before_or_target])
    supply_count = math.ceil(count / DATA['periodic'][5]['frequency']) if count else 0
    periodic = cleaning_count * cleaning_time + supply_count * supply_time
    cyclic = sum(ni * r['t_min'] for ni, r in zip(n, regs))
    mean_cyclic = sum(ni * r['mean'] for ni, r in zip(n, regs))
    active = sum(ni * sum(t for t, e in zip(r['adjusted'], ELEMENTS)
                          if e['kind'] != 'Авто/ожидание') for ni, r in zip(n, regs)) + periodic
    machine = sum(ni * (r['adjusted'][4] + r['adjusted'][6] + r['adjusted'][5])
                  for ni, r in zip(n, regs)) + cleaning_count * cleaning_time
    return dict(cycles=count, kilograms=sum(q), available=available,
                takt=available / count if count else None,
                cleaning_count=cleaning_count, periodic=periodic,
                cyclic=cyclic, total=cyclic + periodic,
                mean_total=mean_cyclic + periodic, active=active, machine=machine,
                weighted_cycle=cyclic / count if count else None,
                load=(cyclic + periodic) / available if available > 0 else None)


if __name__ == '__main__':
    for target in [False, True]:
        print('ЦЕЛЬ' if target else 'ДО', json.dumps(reference_result(target), ensure_ascii=False, indent=2))
