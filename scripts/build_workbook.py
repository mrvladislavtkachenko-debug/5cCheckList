"""Build the self-contained, macro-free workbook; original templates stay untouched.

The forms of Appendix 3 are adapted for 18 products and a single sequential
SOLIA workcell. Appendix 7's unrelated hidden audit sheets are not transplanted.
"""
from __future__ import annotations
import argparse
import datetime as dt
import json
from pathlib import Path
import sys
import xlsxwriter
import openpyxl
from xlsxwriter.utility import xl_col_to_name as col

from process_model import (ROOT, DATA, PRODUCTS, ELEMENTS, BLOCK, RAW_LAST,
                           block_row, samples, reference_result)

OUT = ROOT / 'deliverables' / 'Стандартизированная_работа_И-ГЦ-02.xlsx'
DEMO = 'ДЕМО / ПРОЕКТ • Жёлтые ячейки — ваш ввод. Расчёты не являются утверждёнными нормативами.'
SHEETS = ['Начало', 'Ввод', 'Режимы И-ГЦ-02', 'Ассортимент', 'Замеры ДО', 'Замеры ЦЕЛЬ',
          'Периодическая работа', 'Расчет ДО', 'Расчет ЦЕЛЬ', 'Время такта',
          'Лист наблюдений', 'Ручная работа', 'Карта СР', 'Производственная способность',
          'ОТ ДО', 'ОТ ЦЕЛЬ', 'Диаграмма', 'Общий список улучшений', 'Отчет об улучшении',
          'СОП', 'СОП санобработка', 'План внедрения', 'Критерии качества СР',
          'Чек-лист СОП', 'Показатели', 'Контроль', 'Допущения и источники', 'Лист ознакомления']
INPUTS = []
F = {}
BOOK = None
WS = {}


def ref(sheet, cell):
    return f"'{sheet}'!{cell}"


def write(ws, cell, value, style='body'):
    ws.write(cell, value, F[style])


def formula(ws, cell, value, style='calc'):
    assert value.startswith('=')
    ws.write_formula(cell, value, F[style])


def merge(ws, address, text, style='body'):
    ws.merge_range(address, text, F[style])


def inp(ws, cell, value, note, style='input', validation=None):
    ws.write(cell, value, F[style])
    ws.write_comment(cell, 'ЗАМЕНИТЕ НА РЕАЛЬНЫЕ ДАННЫЕ.\n' + note,
                     {'author': 'Расчётная книга', 'width': 330, 'height': 140})
    if validation:
        ws.data_validation(cell, {**validation, 'error_type': 'stop', 'show_error': True,
                                 'error_title': 'Проверьте ввод',
                                 'error_message': 'Введите значение в допустимых пределах. См. примечание к ячейке.'})
    INPUTS.append((ws.name, cell, note))


def decimal(minimum=0, maximum=86400):
    return {'validate': 'decimal', 'criteria': 'between', 'minimum': minimum, 'maximum': maximum}


def integer(minimum=0, maximum=100000):
    return {'validate': 'integer', 'criteria': 'between', 'minimum': minimum, 'maximum': maximum}


def choices(values):
    return {'validate': 'list', 'source': values}


def headers(ws, row, labels):
    for i, text in enumerate(labels):
        write(ws, f'{col(i)}{row}', text, 'header')
    ws.set_row(row - 1, 44)


def setup(name, title, end_col='H', landscape=True, a3=False):
    ws = WS[name]
    ws.hide_gridlines(2)
    ws.set_tab_color('#147D92')
    ws.set_default_row(23)
    ws.set_column('A:A', 6)
    ws.set_column('B:B', 43)
    ws.set_column('C:H', 17)
    merge(ws, f'A1:{end_col}1', title, 'title')
    merge(ws, f'A2:{end_col}2', DEMO, 'warning')
    ws.set_row(0, 35)
    ws.set_row(1, 31)
    ws.write_url('A3', "internal:'Начало'!A1", F['link'], '← Начало')
    if name != 'Начало':
        merge(ws, f'B3:{end_col}3', 'Синий фон — формула; серый — инструкция; жёлтый — ввод. Защита снимается без пароля.', 'muted')
    ws.freeze_panes(6, 2)
    ws.set_zoom(85)
    ws.set_paper(8 if a3 else 9)
    if landscape:
        ws.set_landscape()
    else:
        ws.set_portrait()
    ws.fit_to_pages(1, 0)
    ws.set_margins(0.25, 0.25, 0.4, 0.4)
    ws.set_header('&LИ-ГЦ-02 · расчётная версия бланков МУ-41-2024&RДЕМО / НЕ УТВЕРЖДЕНО', {'margin': 0.15})
    ws.set_footer('&L'+name+'&RСтр. &P / &N', {'margin': 0.15})
    ws.repeat_rows(0, 2)
    ws.protect('', {'select_locked_cells': True, 'select_unlocked_cells': True,
                    'autofilter': True, 'sort': True, 'objects': False})
    return ws


def selected(column, offset, scenario=None):
    """Bounded INDEX, no INDIRECT, full-column arrays, macros or external links."""
    idx = "'Ввод'!$B$29"
    row = f'7+({idx}-1)*{BLOCK}+{offset}'
    def one(state):
        return f"INDEX('Замеры {state}'!${column}$1:${column}${RAW_LAST},{row})"
    expr = one(scenario) if scenario else f'IF(\'Ввод\'!$B$13="ДО",{one("ДО")},{one("ЦЕЛЬ")})'
    return f'IF({idx}>0,{expr},"")'


def selected_cat(column):
    return f'IF(\'Ввод\'!$B$29>0,INDEX(\'Ассортимент\'!${column}$8:${column}$25,\'Ввод\'!$B$29),"")'


def selected_calc(column, scenario='ДО'):
    return f'IF(\'Ввод\'!$B$29>0,INDEX(\'Расчет {scenario}\'!${column}$8:${column}$25,\'Ввод\'!$B$29),"")'


def metadata(ws, end_col='H'):
    merge(ws, f'A4:{end_col}4', '', 'subhead')
    formula(ws, 'A4', '="Продукт: "&\'Ввод\'!B12&"  |  РМ: "&\'Ввод\'!B8&"  |  Тип II; одна последовательная ячейка"', 'subhead')
    merge(ws, f'A5:{end_col}5', '', 'muted')
    formula(ws, 'A5', '="От: осмотр партии. До: возврат на исходную позицию. Подготовил: "&\'Ввод\'!B9', 'muted')


def build_inputs():
    ws = setup('Ввод', '01 / ИСХОДНЫЕ ДАННЫЕ И ВЫБОР ПРОДУКТА', 'D')
    ws.set_tab_color('#E9B949')
    ws.set_column('A:A', 44)
    ws.set_column('B:B', 58)
    ws.set_column('C:C', 15)
    ws.set_column('D:D', 76)
    headers(ws, 5, ['Параметр', 'Ваше значение', 'Единица', 'Что заменить / источник'])
    values = {
        6: ('Предприятие', 'АО «ПП «Русский хлеб» — уточнить юрлицо', '', 'В титуле И-ГЦ-02 также указано ООО «Курортное кафе». Перед утверждением сверить реквизиты.'),
        7: ('Цех', 'Кулинарный цех', '', 'Уточнить наименование цеха.'),
        8: ('Рабочее место', 'Овощной участок / SOLIA', '', 'Фактическое рабочее место и инвентарный номер оборудования.'),
        9: ('Подготовил', 'Впишите ФИО составителя', '', 'Без вымышленных подписей и подтверждения проведения наблюдений.'),
        10: ('Дата начала проекта', dt.datetime(2026, 9, 5), 'дата', 'Демонстрационная дата; от неё рассчитываются плановые сроки.'),
        12: ('Продукт для подробных бланков', PRODUCTS[0]['name'], '', 'Выбор перестраивает наблюдения, ручную работу, ОТ, карту и СОП. Общие расчёты всегда учитывают все 18 продуктов.'),
        13: ('Сценарий подробных наблюдений', 'ДО', '', 'ДО — модель исходного состояния; ЦЕЛЬ — прогноз. Ни один не является фактом до замены замеров.'),
        14: ('Основание данных ДО', 'ДЕМО: синтетические замеры; заменить фактическими', '', 'Укажите дату, смену, исполнителя и ссылку на лист/видеозапись реальных замеров.'),
        15: ('Основание данных ЦЕЛЬ', 'ДЕМО: прогноз; улучшения не внедрены', '', 'Не переименовывать прогноз в факт без контрольных наблюдений.'),
        17: ('Продолжительность одной смены', 480, 'мин/смену', 'Модель рассчитывает ОДНУ смену. Для разных смен используйте отдельные копии книги.'),
        18: ('Обед', 30, 'мин/смену', 'Только регламентированное время. Не включайте сюда периодическую санобработку.'),
        19: ('Прочие регламентные перерывы', 20, 'мин/смену', 'Сумма перерывов; не учитывать повторно в периодической работе.'),
        20: ('Фиксированная подготовка смены', 10, 'мин/смену', 'Не включает подачу тары и санобработку бака, учтённые отдельно.'),
        21: ('Фиксированное завершение смены', 10, 'мин/смену', 'Не включает санобработку бака. Она целиком находится в периодической работе.'),
        22: ('Доступно моечных машин', 1, 'ед.', 'При значении не 1 ресурсные оценки агрегированные. ОТ и карта описывают только одну ячейку; нужен фактический график распределения.'),
        23: ('Операторов в смене', 1, 'чел.', 'В базовой схеме оператор не выполняет другую работу во время автоматической программы. Дополнительные люди не ускоряют машину автоматически.'),
        25: ('Помывок с последней санобработки к началу смены', 0, '0…3', 'Фактический счётчик. 0 означает подготовленный чистый бак. Для неполного предыдущего блока введите 1–3.'),
        26: ('Резерв завершающей санобработки неполного блока', 1, '1 / 0', 'ДЕМО: 1 — резерв полного комплекса после неполной четвёрки в конце смены. Это плановое допущение; уточнить по санпрограмме. Регулярную обработку после 4 помывок отключить нельзя.'),
        33: ('Ответственный за внедрение', 'Начальник кулинарного цеха — впишите ФИО', '', 'Проект распределения ролей, не назначение приказом.'),
        34: ('Ответственный за технологию', 'Инженер-технолог — впишите ФИО', '', 'Проверяет режимы, нормы загрузки, санпрограмму и качество.'),
        35: ('Исполнитель', 'Оператор SOLIA — впишите ФИО', '', 'Указать реально наблюдаемого и обученного исполнителя.'),
        36: ('Ответственный за ОТ', 'Ответственный по ОТ — впишите ФИО', '', 'Согласовать безопасные действия, СИЗ, инструкции и допуск.'),
        37: ('Утверждает', 'Уполномоченный руководитель — впишите ФИО', '', 'Утверждение выполняется вне расчётной модели в установленном порядке.'),
        38: ('Статус утверждения', 'Не утверждено', '', 'Редактировать только после реального согласования и утверждения; подпись не генерируется.'),
        40: ('Номер улучшения для отчёта', 1, '1…6', 'Переключает подробный отчёт об улучшении.')
    }
    for r, (label, value, unit, note) in values.items():
        write(ws, f'A{r}', label)
        validation = None
        style = 'input_text'
        if isinstance(value, (int, float)):
            validation = integer(0, 1440) if 17 <= r <= 21 else integer(1, 100)
            if r == 25: validation = integer(0, 3)
            if r == 26: validation = integer(0, 1)
            if r == 40: validation = integer(1, 6)
            style = 'input'
        if r == 12: validation = {'validate': 'list', 'source': "='Ассортимент'!$B$8:$B$25"}
        if r == 13: validation = choices(['ДО', 'ЦЕЛЬ'])
        if r == 38: validation = choices(['Не утверждено', 'На согласовании', 'Утверждено'])
        if r == 10: style = 'input_date'
        inp(ws, f'B{r}', value, note, style, validation)
        write(ws, f'C{r}', unit, 'muted')
        write(ws, f'D{r}', note)
        ws.set_row(r - 1, 43 if r not in (14, 15, 26) else 60)
    for r, label, expression, unit in [
        (11, 'Принятый тип работы', None, 'II'),
        (24, 'Обязательная периодичность санобработки', None, '4 помывки'),
        (27, 'Фонд рабочего времени одной ячейки', '=IF(AND(COUNT(B17:B21)=5,MIN(B17:B21)>=0,B17>SUM(B18:B21)),(B17-SUM(B18:B21))*60,"")', 'с/смену'),
        (28, 'Всего партий в плане', "=IF(COUNT('Ассортимент'!O8:O25)=18,'Ассортимент'!O27,\"\")", 'партий/смену'),
        (29, 'Индекс выбранного продукта', '=IFERROR(MATCH(B12,\'Ассортимент\'!B8:B25,0),0)', ''),
        (30, 'Такт выпуска по всему ассортименту', '=IF(AND(ISNUMBER(B27),B27>0,ISNUMBER(B28),B28>0),B27/B28,"")', 'с/партию')]:
        write(ws, f'A{r}', label)
        if expression: formula(ws, f'B{r}', expression)
        else: write(ws, f'B{r}', 4 if r == 24 else 'II — многономенклатурная работа', 'source')
        write(ws, f'C{r}', unit, 'muted')
    write(ws, 'D11', 'Принят для модели; подтвердить наблюдением. Внутри каждого продукта используется поэлементный подход I типа.')
    write(ws, 'D24', 'И-ГЦ-02, стр. 3. Частота — требование источника, не объект оптимизации.')
    merge(ws, 'A42:D43', 'Проверяйте лист «Контроль» после каждого изменения. Количество людей не умножает время такта. Для нескольких ячеек нужна отдельная балансировка и распределение целых партий.', 'warning')
    ws.print_area('A1:D43')


def build_regimes():
    ws = setup('Режимы И-ГЦ-02', '02 / ТЕХНОЛОГИЧЕСКИЕ ДАННЫЕ ИЗ ИНСТРУКЦИИ', 'J')
    ws.set_column('A:A', 12); ws.set_column('B:H', 14); ws.set_column('I:J', 24)
    headers(ws, 7, ['Программа', 'Мытьё, об/мин', 'Мытьё, с', 'Ополаск., об/мин', 'Ополаск., с', 'Вращение, об/мин', 'Вращение, с', 'Всего авто, с', 'Тара после обработки', 'Источник'])
    for r, p in enumerate(DATA['programs'], 8):
        fields = [p['id'], p['wash_rpm'], p['wash_s'], p['rinse_rpm'], p['rinse_s'], p['spin_rpm'], p['spin_s']]
        for c, value in enumerate(fields): write(ws, f'{col(c)}{r}', value, 'source')
        formula(ws, f'H{r}', f'=SUM(C{r},E{r},G{r})')
        write(ws, f'I{r}', '«ЗЕЛЕНЬ» для зелени; «СО» для овощей', 'source')
        write(ws, f'J{r}', 'И-ГЦ-02, PDF стр. 3, таблица №1', 'source')
        ws.set_row(r-1, 52)
    merge(ws, 'A13:J14', 'ВАЖНО: 01 = 450 с; 02 = 180 с; 03 = 300 с. Значения «1:00 мин:сек» переведены в 60 секунд, «6:00» — в 360 секунд. Это не 1 и не 6 секунд.', 'warning')
    merge(ws, 'A16:J18', 'После каждой 4-й помывки: слить бак → после слива выключить и обесточить оборудование → промыть бак моющими средствами и продезинфицировать. В книге это отдельный периодический комплекс, а не вычет из фонда времени.', 'source')
    merge(ws, 'A20:J22', 'И-ГЦ-02 не задаёт средство, концентрацию, температуру и экспозицию дезинфекции, перечень СИЗ и длительности ручных действий. Они НЕ придуманы как технологические нормы. Демонстрационное время в периодических замерах служит только иллюстрацией расчёта.', 'warning')
    merge(ws, 'A24:J25', 'Названия и загрузки всех 18 позиций перенесены в «Ассортимент». Группировка для расчёта — проектная; отдельные названия («сельдерей ствол») сохранены как в источнике. Предел трактуется консервативно как верхняя масса одной загрузки.', 'muted')
    ws.print_area('A1:J25')


def build_catalog():
    ws = setup('Ассортимент', '03 / ПЛАН СМЕНЫ · МАССОВЫЙ БАЛАНС · ЧИСЛО ПАРТИЙ', 'T', a3=True)
    ws.set_tab_color('#E9B949')
    ws.set_column('A:A', 5); ws.set_column('B:B', 29); ws.set_column('C:C', 18)
    ws.set_column('D:R', 12); ws.set_column('S:T', 30)
    merge(ws, 'A5:T5', 'Меняйте F:J. План задаётся в кг годного после мойки. Выход подготовки и выход мойки — разные показатели. Партии считаются по массе, поступающей непосредственно в SOLIA.', 'subhead')
    headers(ws, 7, ['№', 'Продукт', 'Группа', 'Программа', 'Предел загрузки, кг', 'План годного, кг', 'Выход подготовки', 'Выход мойки', 'Партия план, кг', 'Партия замера, кг', 'В мойку, кг', 'Исходное сырьё, кг', 'Отсев до мойки, кг', 'Потери при мойке, кг', 'Партий, шт.', 'Последняя партия, кг', 'Доля партий', 'Авто, с/партию', 'Контроль вводов', 'Сопоставимость замеров'])
    for i, p in enumerate(PRODUCTS):
        r = 8 + i
        for c, value in enumerate([i+1, p['name'], p['group'], p['program'], p['max_kg']]):
            write(ws, f'{col(c)}{r}', value, 'source')
        inp(ws, f'F{r}', p['demo_batches'] * p['max_kg'], 'Потребность в годном продукте после мойки за одну смену. 0 исключает позицию из плана.', validation=decimal(0, 1000000))
        inp(ws, f'G{r}', .95 if p['group'] == 'Зелень' else .98, 'ДЕМО. Масса подготовленного сырья, поступившего в мойку / масса полученного сырья. Не норматив отходов.', 'input_pct', decimal(.000001, 1))
        inp(ws, f'H{r}', 1, 'ДЕМО 100%. Годная масса после мойки / масса загруженного сырья. Не путать с выходом подготовки.', 'input_pct', decimal(.000001, 1))
        inp(ws, f'I{r}', p['max_kg'], 'Плановая масса загрузки. Не превышать серый предел. Неполную последнюю партию модель обрабатывает полным циклом, без линейного сокращения времени.', validation=decimal(.001, p['max_kg']))
        inp(ws, f'J{r}', p['max_kg'], 'Масса, при которой выполнены замеры ДО и ЦЕЛЬ. При смене партии нужны новые замеры: времена не масштабируются пропорционально кг.', validation=decimal(.001, p['max_kg']))
        formula(ws, f'S{r}', f'=IF(AND(COUNT(F{r}:J{r})=5,F{r}>=0,G{r}>0,G{r}<=1,H{r}>0,H{r}<=1,I{r}>0,I{r}<=E{r},J{r}>0,J{r}<=E{r}),"ОК","ОШИБКА: проверьте F:J")', 'calc_text')
        for c, expr in {
            'K': f'F{r}/H{r}', 'L': f'K{r}/G{r}', 'M': f'L{r}-K{r}', 'N': f'K{r}-F{r}',
            'O': f'ROUNDUP(ROUND(K{r}/I{r},10),0)',
            'P': f'IF(O{r}>0,MAX(0,K{r}-(O{r}-1)*I{r}),0)',
        }.items(): formula(ws, f'{c}{r}', f'=IF(S{r}="ОК",{expr},"")', 'calc_int' if c=='O' else 'calc')
        formula(ws, f'Q{r}', f'=IF(AND(ISNUMBER(O{r}),ISNUMBER($O$27),$O$27>0),O{r}/$O$27,0)', 'pct')
        formula(ws, f'R{r}', f'=VLOOKUP(D{r},\'Режимы И-ГЦ-02\'!$A$8:$H$10,8,FALSE)')
        formula(ws, f'T{r}', f'=IF(S{r}<>"ОК","Исправьте ввод",IF(F{r}=0,"Нет в плане",IF(ABS(I{r}-J{r})<0.000001,"Партия совпадает","НУЖЕН НОВЫЙ ХРОНОМЕТРАЖ")))', 'calc_text')
        ws.set_row(r-1, 45)
    write(ws, 'B27', 'ИТОГО', 'subhead')
    for c in ['F','K','L','M','N','O']:
        formula(ws, f'{c}27', f'=IF(COUNT({c}8:{c}25)=18,SUM({c}8:{c}25),"")', 'total')
    formula(ws, 'Q27', '=SUM(Q8:Q25)', 'pct')
    merge(ws, 'A29:T30', 'Число партий = ROUNDUP(масса в мойку / плановая партия). Усреднение времени ведётся по количеству партий, НЕ по кг. Для неполной партии сохранены полные ручные и автоматические времена — консервативное допущение. Значение 0 в плане допустимо; пустые/ошибочные исходные данные не означают нулевой спрос.', 'warning')
    ws.autofilter('A7:T25'); ws.freeze_panes(7, 2)
    ws.conditional_format('S8:T25', {'type':'text','criteria':'containing','value':'ОШИБКА','format':F['bad']})
    ws.print_area('A1:T30')


def build_raw(state):
    ws = setup(f'Замеры {state}', f'04 / ПОЭЛЕМЕНТНЫЕ ЗАМЕРЫ · {state} · ВСЕ 18 ПРОДУКТОВ', 'X', a3=True)
    ws.set_tab_color('#E9B949')
    ws.set_column('A:A', 4); ws.set_column('B:B', 36); ws.set_column('C:C', 16); ws.set_column('D:D', 27)
    ws.set_column('E:N', 8); ws.set_column('O:U', 10); ws.set_column('V:X', 18)
    ws.set_column('Y:AQ', None, None, {'hidden': True})
    merge(ws, 'A5:X5', 'Вводите длительность каждого элемента в секундах, НЕ накопительные отсечки. Флаг 0 исключает ВЕСЬ цикл. Если хотя бы один элемент пустой/ошибочный — весь цикл автоматически исключается.', 'subhead')
    for i, p in enumerate(PRODUCTS):
        r = block_row(i); first, last, total, stats = r+5, r+15, r+16, r+17
        merge(ws, f'A{r}:X{r}', f'{i+1:02}. {p["name"]} · {state} · 10 демонстрационных циклов, заменить реальными', 'subhead')
        merge(ws, f'A{r+1}:X{r+1}', '', 'muted')
        formula(ws, f'A{r+1}', f'="Масса замера, кг: "&\'Ассортимент\'!J{8+i}&". Все замеры одной партии/исполнителя/условий. Основание: "&\'Ввод\'!B{14 if state=="ДО" else 15}', 'muted')
        headers(ws, r+2, ['№','Рабочий элемент','Тип','Точка отсчёта']+[str(j) for j in range(1,11)]+['t min','t max','Размах','Рег. с','t отрег.','Колебания','Среднее','Контроль','Служебно','Статус'])
        write(ws, f'B{r+3}', 'Учитывать цикл: 1 / 0', 'body')
        write(ws, f'B{r+4}', 'Причина исключения / условия', 'body')
        for j in range(10):
            c = col(4+j)
            inp(ws, f'{c}{r+3}', 1, f'{p["name"]}, цикл {j+1}. 0 исключает все элементы данного цикла, а не один выброс.', validation=integer(0,1))
            inp(ws, f'{c}{r+4}', 'ДЕМО', 'Замените на описание условий. При исключении обязательна конкретная причина.', 'input_text')
        ws.set_row(r+3, 33)
        raw = samples(i, state=='ЦЕЛЬ')
        for e_idx, e in enumerate(ELEMENTS):
            rr = first + e_idx
            write(ws, f'A{rr}', e_idx+1 if e['kind']=='Ручная' else '—')
            write(ws, f'B{rr}', e['name'])
            write(ws, f'C{rr}', e['kind'], 'source')
            write(ws, f'D{rr}', e['endpoint'], 'muted')
            ws.set_row(rr-1, 40)
            for j in range(10):
                c = col(4+j)
                if e['kind']=='Авто/ожидание':
                    formula(ws, f'{c}{rr}', f"='Ассортимент'!R{8+i}", 'source_num')
                else:
                    inp(ws, f'{c}{rr}', raw[e_idx][j], f'ДЕМО, {state}, {p["name"]}, цикл {j+1}. Длительность элемента «{e["name"]}», секунд. Нулевая длительность означает реальное отсутствие действия/потерь; пустая — нет замера.', validation=decimal())
                formula(ws, f'{col(32+j)}{rr}', f'=IF(ISNUMBER({c}${stats}),{c}{rr},"")')
            helper = f'AG{rr}:AP{rr}'
            formula(ws, f'O{rr}', f'=IF(COUNT({helper})>0,MIN({helper}),"")')
            formula(ws, f'P{rr}', f'=IF(COUNT({helper})>0,MAX({helper}),"")')
            formula(ws, f'Q{rr}', f'=IF(ISNUMBER(O{rr}),P{rr}-O{rr},"")')
            formula(ws, f'AQ{rr}', f'=IF(C{rr}="Ручная",O{rr},0)')
            formula(ws, f'R{rr}', f'=IF(ISNUMBER(O{rr}),IF(C{rr}="Ручная",IF($S${stats}>0,$R${stats}*Q{rr}/$S${stats},IF($W${stats}={e_idx+1},$R${stats},0)),0),"")')
            formula(ws, f'S{rr}', f'=IF(ISNUMBER(O{rr}),O{rr}+R{rr},"")')
            formula(ws, f'T{rr}', f'=IF(ISNUMBER(S{rr}),P{rr}-S{rr},"")')
            formula(ws, f'U{rr}', f'=IF(COUNT({helper})>0,AVERAGE({helper}),"")')
        write(ws, f'B{total}', 'ИТОГИ: все элементы цикла', 'subhead')
        write(ws, f'B{stats}', 'ТОЛЬКО полные учтённые циклы', 'subhead')
        for j in range(10):
            c = col(4+j)
            formula(ws, f'{c}{total}', f'=IF(COUNT({c}{first}:{c}{last})=11,SUM({c}{first}:{c}{last}),"")', 'total')
            formula(ws, f'{c}{stats}', f'=IF(AND({c}{r+3}=1,COUNT({c}{first}:{c}{last})=11,MIN({c}{first}:{c}{last})>=0),{c}{total},"")', 'total')
            ws.conditional_format(f'{c}{first}:{c}{last}', {'type':'formula','criteria':f'={c}${r+3}=0','format':F['excluded']})
        for c in ['O','P','Q','R','S','T','U']:
            formula(ws, f'{c}{total}', f'=IF(COUNT({c}{first}:{c}{last})=11,SUM({c}{first}:{c}{last}),"")', 'total')
        formulas = {
            'O': f'IF(COUNT(E{stats}:N{stats})>0,MIN(E{stats}:N{stats}),"")',
            'P': f'IF(COUNT(E{stats}:N{stats})>0,MAX(E{stats}:N{stats}),"")',
            'Q': f'IF(ISNUMBER(O{stats}),P{stats}-O{stats},"")',
            'R': f'IF(ISNUMBER(O{stats}),MAX(0,O{stats}-O{total}),"")',
            'S': f'SUMIF(C{first}:C{last},"Ручная",Q{first}:Q{last})',
            'T': f'COUNT(E{stats}:N{stats})',
            'U': f'IF(COUNT(E{stats}:N{stats})>0,AVERAGE(E{stats}:N{stats}),"")',
            'V': f'IF(ISNUMBER(O{stats}),S{total}-O{stats},"")',
            'W': f'MATCH(MAX(AQ{first}:AQ{last}),AQ{first}:AQ{last},0)',
            'X': f'IF(T{stats}=0,"НЕТ ЗАМЕРОВ",IF(T{stats}<10,"МЕНЕЕ 10 ЦИКЛОВ",IF(ABS(V{stats})<0.000001,"ОК","ОШИБКА РЕГУЛИРОВКИ")))'
        }
        for c, expr in formulas.items(): formula(ws, f'{c}{stats}', '='+expr, 'calc_text' if c=='X' else 'total')
        merge(ws, f'A{r+18}:X{r+18}', 'В строке учтённых циклов: O = T min; P = T max; Q = размах цикла; R = T min − сумма t min; S = сумма размахов ручных элементов; T = число циклов; U = среднее; V = сумма t отрег − T min (должно быть 0).', 'muted')
        ws.conditional_format(f'E{first}:N{last}', {'type':'formula','criteria':f'=AND(E{first}=$O{first},ISNUMBER(E${stats}))','format':F['min']})
        ws.conditional_format(f'E{first}:N{last}', {'type':'formula','criteria':f'=AND(E{first}=$P{first},ISNUMBER(E${stats}),$P{first}>$O{first})','format':F['max']})
    ws.set_h_pagebreaks([block_row(i)-1 for i in range(1,len(PRODUCTS))])
    ws.print_area(f'A1:X{RAW_LAST}'); ws.freeze_panes(5,4)


def build_periodic():
    ws = setup('Периодическая работа', '05 / ЛИСТ НАБЛЮДЕНИЯ ПЕРИОДИЧЕСКОЙ РАБОТЫ', 'R', a3=True)
    ws.set_tab_color('#E9B949'); ws.set_column('B:B', 53); ws.set_column('C:C', 18); ws.set_column('D:Q', 12); ws.set_column('R:R', 48)
    merge(ws, 'A5:R5', 'Минимум из 3 замеров / периодичность = нормативное распределение на цикл. Для проверки плана применяются ЦЕЛЫЕ события и отдельный резерв завершения неполного блока.', 'subhead')
    headers(ws, 7, ['№','Рабочий элемент','Событие','Каждые N циклов','ДО 1, с','ДО 2, с','ДО 3, с','min ДО, с','ДО, с/цикл','ЦЕЛЬ 1, с','ЦЕЛЬ 2, с','ЦЕЛЬ 3, с','min ЦЕЛЬ, с','ЦЕЛЬ, с/цикл','Событий/смену','ДО за смену, с','ЦЕЛЬ за смену, с','Источник / ограничения'])
    for i, p in enumerate(DATA['periodic']):
        r = 8+i
        for c, value in enumerate([i+1,p['name'],p['event']]): write(ws, f'{col(c)}{r}', value)
        if i<5: formula(ws, f'D{r}', "='Ввод'!B24", 'source_num')
        else: inp(ws, f'D{r}', p['frequency'], 'ДЕМО. Комплект до 1-го цикла, далее каждые N циклов. Уточните реальную частоту подачи чистой тары.', validation=integer(1,10000))
        for j, value in enumerate(p['before']): inp(ws, f'{col(4+j)}{r}', value, p['source']+'; длительность всего элемента, с. Для дезинфекции это НЕ заданная экспозиция. Не сокращать ниже требований утверждённых средств.', validation=decimal(.001,86400))
        for j, value in enumerate(p['target']):
            if i<5: formula(ws, f'{col(9+j)}{r}', f'={col(4+j)}{r}', 'source_num')
            else: inp(ws, f'{col(9+j)}{r}', value, 'Прогноз улучшения логистики. Санитарные операции не сокращаются.', validation=decimal(.001,86400))
        formula(ws, f'H{r}', f'=IF(AND(COUNT(E{r}:G{r})=3,MIN(E{r}:G{r})>0),MIN(E{r}:G{r}),"")')
        formula(ws, f'M{r}', f'=IF(AND(COUNT(J{r}:L{r})=3,MIN(J{r}:L{r})>0),MIN(J{r}:L{r}),"")')
        for output, minimum in [('I','H'),('N','M')]: formula(ws, f'{output}{r}', f'=IF(AND(ISNUMBER({minimum}{r}),ISNUMBER(D{r}),D{r}>0),{minimum}{r}/D{r},"")')
        expr = '$B$20' if i<5 else f'IF(AND(ISNUMBER(\'Ввод\'!B28),\'Ввод\'!B28>=0,ISNUMBER(D{r}),D{r}>0),ROUNDUP(\'Ввод\'!B28/D{r},0),"")'
        formula(ws, f'O{r}', '='+expr, 'calc_int')
        for output, minimum in [('P','H'),('Q','M')]: formula(ws, f'{output}{r}', f'=IF(AND(ISNUMBER({minimum}{r}),ISNUMBER(O{r})),{minimum}{r}*O{r},"")')
        write(ws, f'R{r}', p['source'], 'muted'); ws.set_row(r-1, 74)
    write(ws, 'B15', 'ИТОГО', 'subhead')
    for c in ['I','N','P','Q']: formula(ws, f'{c}15', f'=IF(COUNT({c}8:{c}13)=6,SUM({c}8:{c}13),"")', 'total')
    for r, label, expr in [
        (18,'Регулярных полных санобработок', '=IF(AND(ISNUMBER(\'Ввод\'!B28),\'Ввод\'!B28>0,ISNUMBER(\'Ввод\'!B25),\'Ввод\'!B25>=0,\'Ввод\'!B25<=3,MOD(\'Ввод\'!B25,1)=0),INT((\'Ввод\'!B28+\'Ввод\'!B25)/4),IF(\'Ввод\'!B28=0,0,""))'),
        (19,'Дополнительных после неполного блока', '=IF(COUNT(\'Ввод\'!B28,\'Ввод\'!B25)=2,IF(AND(\'Ввод\'!B28>0,\'Ввод\'!B26=1,MOD(\'Ввод\'!B28+\'Ввод\'!B25,4)>0),1,0),"")'),
        (20,'Всего полных комплексов за смену', '=IF(ISNUMBER(B18),B18+B19,"")')]:
        merge(ws, f'C{r}:H{r}', label)
        formula(ws, f'B{r}', expr, 'total')
    merge(ws, 'A22:R24', 'Частота 4 — фиксированное требование И-ГЦ-02. Комплекс состоит из 5 строк, но число санобработок НЕ суммируется по этим строкам. Завершение неполного блока — отдельный резерв, не утверждённое требование источника. Не вычитать эти же секунды ещё раз из фонда рабочего времени.', 'warning')
    merge(ws, 'A26:R28', 'ДО и ЦЕЛЬ используют одинаковые санитарные времена. Улучшение не достигается за счёт экспозиции, концентрации, температуры или пропуска очистки. Обесточивание следует ПОСЛЕ слива, как указано в И-ГЦ-02. Восстановление готовности — проектное дополнение по паспорту машины, который требуется проверить.', 'source')
    ws.print_area('A1:R28')


def build_calc(state):
    ws = setup(f'Расчет {state}', f'06 / II ТИП · СРЕДНЕВЗВЕШЕННЫЕ ВРЕМЕНА · {state}', 'U', a3=True)
    ws.set_column('B:B', 30); ws.set_column('C:T', 13); ws.set_column('U:U', 27)
    headers(ws, 7, ['№','Продукт','Партий','Вес','Ручная, с','Переходы, с','Потери, с','Авто/ожид., с','Цикл min = отрег., с','Цикл сред., с','Цикл max, с','Σ колебаний элементов, с','Периодич. норматив, с/цикл','Цикл + периодич., с','Циклическое min, с/смену','Циклическое сред., с/смену','Активное занятие, с/цикл','Машина, с/цикл','Учтено циклов','Контроль регулировки','Полнота данных'])
    raw = f'Замеры {state}'
    pc = 'I' if state=='ДО' else 'N'
    ps = 'P' if state=='ДО' else 'Q'
    for i, p in enumerate(PRODUCTS):
        r=8+i; b=block_row(i); first,last,stats=b+5,b+15,b+17
        write(ws,f'A{r}',i+1); write(ws,f'B{r}',p['name'])
        expressions = {
            'C': f"'Ассортимент'!O{r}", 'D': f"'Ассортимент'!Q{r}",
            'E': f'SUMIF(\'{raw}\'!C{first}:C{last},"Ручная",\'{raw}\'!S{first}:S{last})',
            'F': f'SUMIF(\'{raw}\'!C{first}:C{last},"Переход",\'{raw}\'!S{first}:S{last})',
            'G': f"'{raw}'!S{last}", 'H': f"'Ассортимент'!R{r}",
            'I': f"'{raw}'!O{stats}", 'J': f"'{raw}'!U{stats}", 'K': f"'{raw}'!P{stats}",
            'L': f"'{raw}'!T{b+16}", 'M': f"'Периодическая работа'!{pc}15",
            'N': f'IF(AND(ISNUMBER(I{r}),ISNUMBER(M{r})),I{r}+M{r},"")',
            'O': f'IF(AND(ISNUMBER(C{r}),ISNUMBER(I{r})),C{r}*I{r},IF(C{r}=0,0,""))',
            'P': f'IF(AND(ISNUMBER(C{r}),ISNUMBER(J{r})),C{r}*J{r},IF(C{r}=0,0,""))',
            'Q': f'IF(ISNUMBER(I{r}),SUM(E{r}:G{r}),"")',
            'R': f'IF(ISNUMBER(I{r}),\'{raw}\'!S{first+4}+\'{raw}\'!S{first+6}+H{r},"")',
            'S': f"'{raw}'!T{stats}", 'T': f"'{raw}'!V{stats}",
            'U': f'IF(NOT(ISNUMBER(C{r})),"ОШИБКА ПЛАНА",IF(C{r}=0,"Нет в плане",IF(AND(ISNUMBER(I{r}),S{r}>0),"ОК","НЕТ ЗАМЕРОВ")))'
        }
        # Empty statistics remain empty: a direct Excel reference to an empty formula is a string, not 0.
        for c, expr in expressions.items(): formula(ws,f'{c}{r}','='+expr,'pct' if c=='D' else ('calc_text' if c=='U' else 'calc'))
        ws.set_row(r-1,42)
    write(ws,'B27','Ср. вз. / итоги','subhead')
    formula(ws,'C27','=IF(COUNT(C8:C25)=18,SUM(C8:C25),"")','total')
    formula(ws,'D27','=SUM(D8:D25)','pct')
    gate='AND(ISNUMBER($C$27),$C$27>0,COUNTIF($U$8:$U$25,"НЕТ ЗАМЕРОВ")=0,COUNTIF($U$8:$U$25,"ОШИБКА ПЛАНА")=0)'
    for c in list('EFGHIJKLMNQR'):
        formula(ws,f'{c}27',f'=IF(AND({gate},SUMPRODUCT($C$8:$C$25,ISNUMBER({c}8:{c}25)*1)=$C$27),SUMPRODUCT($C$8:$C$25,{c}8:{c}25)/$C$27,"")','total')
    for c in ['O','P']:
        formula(ws,f'{c}27',f'=IF(AND(COUNT({c}8:{c}25)=18,COUNT(C8:C25)=18),SUM({c}8:{c}25),"")','total')
    summary = {
        30: ('Циклическое время плана по отрегулированному циклу', '=O27', 'с/смену'),
        31: ('Циклическое время плана по среднему наблюдаемому циклу', '=P27', 'с/смену'),
        32: ('Периодическая работа: точное число целых событий', f"='Периодическая работа'!{ps}15", 'с/смену'),
        33: ('Полное время плана, отрегулированный цикл + целые события', '=IF(COUNT(B30,B32)=2,B30+B32,"")', 'с/смену'),
        34: ('Полное время плана, средний цикл + целые события', '=IF(COUNT(B31,B32)=2,B31+B32,"")', 'с/смену'),
        35: ('Доступный фонд ОДНОЙ ячейки', "='Ввод'!B27", 'с/смену'),
        36: ('Загрузка одной последовательной ячейки — отрегулированная', '=IF(AND(COUNT(B33,B35)=2,B35>0),B33/B35,"")', '%'),
        37: ('Загрузка одной последовательной ячейки — средняя', '=IF(AND(COUNT(B34,B35)=2,B35>0),B34/B35,"")', '%'),
        38: ('Активное занятие оператора, включая переходы и потери, без автоожидания', f'=IF(AND({gate},ISNUMBER(B32)),SUMPRODUCT(C8:C25,Q8:Q25)+B32,IF(C27=0,0,""))', 'чел·с/смену'),
        39: ('Теоретическая потребность по активному занятию, НЕ штатный норматив', '=IF(AND(COUNT(B38,B35)=2,B35>0),B38/B35,"")', 'экв. чел.'),
        40: ('Нижняя оценка операторов при ПОЛНОМ ожидании программы', '=IF(AND(COUNT(B33,B35)=2,B35>0),ROUNDUP(B33/B35,0),"")', 'чел.'),
        41: ('Занятость машины: загрузка + авто + выгрузка + санобработка', f'=IF(AND({gate},COUNT(\'Периодическая работа\'!{ps}8:{ps}12)=5),SUMPRODUCT(C8:C25,R8:R25)+SUM(\'Периодическая работа\'!{ps}8:{ps}12),IF(C27=0,0,""))', 'маш·с/смену'),
        42: ('Теоретическая потребность в машинах по их занятости', '=IF(AND(COUNT(B41,B35)=2,B35>0),B41/B35,"")', 'экв. машин'),
        43: ('Нижняя целая оценка числа машин', '=IF(ISNUMBER(B42),ROUNDUP(B42,0),"")', 'ед.'),
        44: ('Периодическая работа с завершением неполного блока', '=IF(AND(ISNUMBER(C27),C27>0,ISNUMBER(B32)),B32/C27,"")', 'с/партию'),
        45: ('Ср. вз. средний цикл + точная доля периодической работы', '=IF(AND(ISNUMBER(C27),C27>0,ISNUMBER(B34)),B34/C27,"")', 'с/партию'),
        46: ('Ср. вз. отрег. цикл + точная доля периодической работы', '=IF(AND(ISNUMBER(C27),C27>0,ISNUMBER(B33)),B33/C27,"")', 'с/партию'),
        47: ('Ориентир способности по нормативному усреднению (не точный сменный выпуск)', '=IF(AND(COUNT(B35,I27,M27)=3,B35>0,SUM(I27,M27)>0),INT(B35/(I27+M27)),"")', 'экв. партий'),
        49: ('Проверка плана для 1 машины + 1 оператора, отрег. цикл', '=IF(AND(COUNT(B33,B35)=2,B35>0,C27>0),IF(B33<=B35,"ВПИСЫВАЕТСЯ","НЕ ВПИСЫВАЕТСЯ"),"НЕТ РАСЧЁТА")', ''),
        50: ('Проверка плана для 1 машины + 1 оператора, средний цикл', '=IF(AND(COUNT(B34,B35)=2,B35>0,C27>0),IF(B34<=B35,"ВПИСЫВАЕТСЯ","НЕ ВПИСЫВАЕТСЯ"),"НЕТ РАСЧЁТА")', ''),
    }
    # A compact, deliberately separate summary uses B as value and C:I as label.
    for r,(label,expr,unit) in summary.items():
        formula(ws,f'B{r}',expr,'pct' if unit=='%' else ('calc_text' if r>=49 else 'total'))
        merge(ws,f'C{r}:I{r}',label)
        write(ws,f'J{r}',unit,'muted'); ws.set_row(r-1,32)
    merge(ws,'A52:U54','Машинное время и одновременное ожидание оператора не складываются дважды. В этой модели нет перекрытия подготовки следующей партии с мойкой. Ресурсные оценки для нескольких операторов/машин — нижние агрегированные границы, не готовый график. Ориентир выпуска по усреднению может отличаться от точного плана из-за целых партий и санобработок; основной критерий — строки 33–37 и 49–50.','warning')
    for r, label, expression in [
        (56, 'Агрегированная загрузка доступных машин (нижняя оценка)', '=IF(AND(COUNT(B41,B35,\'Ввод\'!B22)=3,B35>0,\'Ввод\'!B22>0),B41/(B35*\'Ввод\'!B22),"")'),
        (57, 'Агрегированная занятость доступных операторов с ожиданием (нижняя оценка)', '=IF(AND(COUNT(B33,B35,\'Ввод\'!B23)=3,B35>0,\'Ввод\'!B23>0),B33/(B35*\'Ввод\'!B23),"")'),
        (58, 'Агрегированное активное занятие доступных операторов (нижняя оценка)', '=IF(AND(COUNT(B38,B35,\'Ввод\'!B23)=3,B35>0,\'Ввод\'!B23>0),B38/(B35*\'Ввод\'!B23),"")')]:
        formula(ws, f'B{r}', expression, 'pct'); merge(ws, f'C{r}:K{r}', label); ws.set_row(r-1, 35)
    merge(ws, 'A60:U62', 'Строки 56–58 используют фактическое число доступных ресурсов. Это только нижние агрегированные оценки: не учитывают неравномерное распределение целых партий, отдельные счётчики санобработки машин, перемещения между ними и ограничения многомашинного обслуживания. При числе ресурсов не 1 требуется отдельный график; проверка на листе «Контроль» напоминает об этом.', 'warning')
    ws.print_area('A1:U62'); ws.set_h_pagebreaks([29])


def make_formats(book):
    base = {'font_name':'Arial','font_size':10,'valign':'vcenter','text_wrap':True,'border':1,'border_color':'#DDE5EA'}
    def fmt(name, **kw): F[name]=book.add_format({**base,**kw})
    fmt('body',font_color='#213547')
    fmt('title',font_size=20,bold=True,font_color='white',bg_color='#163444',border=0)
    fmt('header',bold=True,font_color='white',bg_color='#246577',align='center')
    fmt('subhead',bold=True,font_color='#163444',bg_color='#DCEEF0')
    fmt('warning',font_color='#7B4600',bg_color='#FFF1D2',bold=True,border=0)
    fmt('muted',font_color='#556778',bg_color='#F5F7F9',font_size=9)
    fmt('source',font_color='#445468',bg_color='#E9EDF1')
    fmt('source_num',font_color='#445468',bg_color='#E9EDF1',num_format='0.00')
    fmt('calc',font_color='#165B7A',bg_color='#EDF6FC',num_format='0.00')
    fmt('calc_int',font_color='#165B7A',bg_color='#EDF6FC',num_format='0')
    fmt('calc_text',font_color='#165B7A',bg_color='#EDF6FC')
    fmt('input',font_color='#163444',bg_color='#FFF2AC',num_format='0.00',locked=False)
    fmt('input_text',font_color='#163444',bg_color='#FFF2AC',locked=False)
    fmt('input_pct',font_color='#163444',bg_color='#FFF2AC',num_format='0.0%',locked=False)
    fmt('input_date',font_color='#163444',bg_color='#FFF2AC',num_format='dd.mm.yyyy',locked=False)
    fmt('date',font_color='#165B7A',bg_color='#EDF6FC',num_format='dd.mm.yyyy')
    fmt('pct',font_color='#165B7A',bg_color='#EDF6FC',num_format='0.0%')
    fmt('total',bold=True,font_color='#163444',bg_color='#DCEEF0',num_format='0.00')
    fmt('link',font_color='#006C8E',underline=True,border=0)
    fmt('bad',font_color='#8F263A',bg_color='#FCE4E8',bold=True)
    fmt('good',font_color='#176647',bg_color='#DCF3E8')
    fmt('excluded',font_color='#999999',font_strikeout=True,bg_color='#EFEFEF')
    fmt('min',font_color='#176647',underline=True)
    fmt('max',font_color='#9B233B',bold=True)
    for name,color in [('manual','#2789A7'),('auto','#7A6AB2'),('walk','#EBA84A'),('wait','#A1ADB7')]:
        fmt(name,bg_color=color,font_color=color,align='center',border=0)


def build(path=OUT):
    global BOOK,WS
    path.parent.mkdir(parents=True, exist_ok=True)
    BOOK=xlsxwriter.Workbook(path, {'strings_to_formulas':False, 'strings_to_urls':False})
    BOOK.set_properties({'title':'Стандартизированная работа: И-ГЦ-02',
                         'subject':'Расчётная версия бланков МУ-41-2024; демонстрационные данные',
                         'author':'Проектный расчёт — не утверждено',
                         'comments':'Исходные требования И-ГЦ-02 сохранены. Без макросов и внешних формульных связей.'})
    BOOK.set_calc_mode('auto')
    WS={name:BOOK.add_worksheet(name) for name in SHEETS}
    make_formats(BOOK)
    build_inputs(); build_regimes(); build_catalog()
    build_raw('ДО'); build_raw('ЦЕЛЬ'); build_periodic()
    build_calc('ДО'); build_calc('ЦЕЛЬ')
    # Reporting builders share the same address map and formatting primitives.
    from workbook_reports import build_reports
    build_reports(sys.modules[__name__])
    WS['Начало'].activate(); WS['Начало'].set_first_sheet()
    BOOK.close()
    (ROOT/'data/input_cells.json').write_text(json.dumps(INPUTS,ensure_ascii=False,indent=2))
    print(f'Built {path.name}: {len(SHEETS)} sheets; {len(INPUTS)} editable cells')
    return path


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--output',type=Path,default=OUT)
    args=parser.parse_args()
    build(args.output)
