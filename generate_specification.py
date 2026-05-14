# -*- coding: utf-8 -*-
"""
Генератор спецификации АПС в формате Excel (specification.xlsx).

Содержит 4 листа:
    1. Спецификация    — полный состав оборудования (ППКП, извещатели, оповещение)
    2. Расчёт ИП       — расчёт количества извещателей по помещениям обоих этажей
    3. Шлейфы          — распределение устройств по шлейфам сигнализации Гранит-12
    4. Кабельный журнал — кабели и металлорукав по типам и трассам
"""
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

import coordinates as C


# ============================================================================
# СТИЛИ
# ============================================================================

TITLE_FONT = Font(name="Times New Roman", size=14, bold=True)
HEADER_FONT = Font(name="Times New Roman", size=11, bold=True)
CELL_FONT = Font(name="Times New Roman", size=11)
TOTAL_FONT = Font(name="Times New Roman", size=11, bold=True)

THIN = Side(border_style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def style_header(ws, row, ncols):
    for col in range(1, ncols + 1):
        c = ws.cell(row=row, column=col)
        c.font = HEADER_FONT
        c.alignment = CENTER
        c.border = BORDER
        c.fill = HEADER_FILL


def style_row(ws, row, ncols, bold=False, fill=False):
    font = TOTAL_FONT if bold else CELL_FONT
    for col in range(1, ncols + 1):
        c = ws.cell(row=row, column=col)
        c.font = font
        c.border = BORDER
        c.alignment = CENTER if col != 2 else LEFT
        if fill:
            c.fill = TOTAL_FILL


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================================
# ЛИСТ 1 — СПЕЦИФИКАЦИЯ ОБОРУДОВАНИЯ
# ============================================================================

SPEC_DATA = [
    # (поз, наименование, марка/тип, ед., кол-во, примечание)
    ("1",  "Прибор приёмно-контрольный пожарный, 12 шлейфов",
            "Гранит-12",                "шт.", 1,
            "НПО «Сибирский Арсенал», САПО.425519.028"),
    ("2",  "Источник питания резервированный 12 В / 3 А, АКБ 17 А·ч",
            "РИП-12-3/17М",             "шт.", 1,
            "Для оповещения; выход ОПВ Гранит-12 = 1 А недостаточен"),
    ("3",  "Извещатель пожарный дымовой оптико-электронный точечный",
            "ИП 212-141",               "шт.", 48,
            "1 эт. — 24 шт.; 2 эт. — 24 шт."),
    ("4",  "Извещатель пожарный тепловой максимальный (класс А3)",
            "ИП 101-1А-А3",             "шт.", 4,
            "Бойлерная (2) + кухня-столовая (2)"),
    ("5",  "Извещатель пожарный ручной",
            "ИПР 513-10",               "шт.", 6,
            "Высота установки 1,5 м, у эвакуационных выходов"),
    ("6",  "Оповещатель пожарный речевой настенный",
            "Соната-М",                 "шт.", 8,
            "12 В, 0,25 А, уровень 96 дБ"),
    ("7",  "Оповещатель пожарный световой",
            "Призма 201",               "шт.", 5,
            "В архивах, серверной, бойлерной, коридорах"),
    ("8",  "Табло световое эвакуационное «ВЫХОД»",
            "Т 12-ОП «Выход»",          "шт.", 4,
            "Над эвакуационными выходами, 12 В"),
    ("9",  "Аккумуляторная батарея",
            "АКБ 12В / 7А·ч",           "шт.", 1,
            "Встроена в Гранит-12 (резерв ≥ 24 ч)"),
    ("10", "Аккумуляторная батарея",
            "АКБ 12В / 17А·ч",          "шт.", 1,
            "Для РИП-12 (резерв оповещения ≥ 24 ч)"),
    ("11", "Кабель пожарный КПСВВнг(А)-LS 1×2×0,5",
            "КПСВВнг(А)-LS 1×2×0,5",    "м",   654,
            "Для шлейфов сигнализации ИП и ИПР, с запасом 20%"),
    ("12", "Кабель пожарный КПСВВнг(А)-LS 1×2×0,75",
            "КПСВВнг(А)-LS 1×2×0,75",   "м",   192,
            "Для линий оповещения, с запасом 20%"),
    ("13", "Кабель силовой ВВГнг(А)-LS 3×1,5",
            "ВВГнг(А)-LS 3×1,5",        "м",   36,
            "Питание ППКП от щита, автомат 6 А"),
    ("14", "Металлорукав в ПВХ оболочке РЗ-Ц-Х ⌀ 15",
            "МР-15",                    "м",   1000,
            "Защита кабельных трасс по СП 6.13130 п.4.2"),
    ("15", "Коробка распределительная пожарная огнестойкая",
            "КРП IP54",                 "шт.", 12,
            "Места соединения шлейфов и ответвлений"),
    ("16", "Извещатель пожарный — выносной устройство световой сигнализации",
            "УС-02",                    "шт.", 5,
            "Для ИП в архивах и серверной (СП 484 п.6.6.13)"),
]


def make_specification_sheet(ws):
    ws.title = "Спецификация"
    ws.merge_cells("A1:F1")
    ws["A1"] = "СПЕЦИФИКАЦИЯ ОБОРУДОВАНИЯ И МАТЕРИАЛОВ"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER

    ws.merge_cells("A2:F2")
    ws["A2"] = "Автоматическая пожарная сигнализация. Административное здание, S общ. = 491,1 м²"
    ws["A2"].font = HEADER_FONT
    ws["A2"].alignment = CENTER

    headers = ["№\nп/п", "Наименование", "Тип, марка", "Ед.\nизм.", "Кол-во", "Примечание"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 6)

    row = 5
    for pos, name, mark, unit, qty, note in SPEC_DATA:
        ws.cell(row=row, column=1, value=pos)
        ws.cell(row=row, column=2, value=name)
        ws.cell(row=row, column=3, value=mark)
        ws.cell(row=row, column=4, value=unit)
        ws.cell(row=row, column=5, value=qty)
        ws.cell(row=row, column=6, value=note)
        style_row(ws, row, 6)
        row += 1

    set_widths(ws, [6, 50, 22, 8, 10, 45])
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[4].height = 38
    for r in range(5, row):
        ws.row_dimensions[r].height = 32


# ============================================================================
# ЛИСТ 2 — РАСЧЁТ ИП ПО ПОМЕЩЕНИЯМ
# ============================================================================

CALC_ROOMS = [
    # (этаж, №, помещение, S м², ИП 212, ИП 101, обоснование)
    (1, 1,  "Тамбур",                3.8,  1, 0, "СП 484 п.6.6.5 — мин. 1 ИП при резервировании смежным шлейфом"),
    (1, 2,  "Охрана",                6.3,  2, 0, "СП 484 п.6.6.1 — минимум 2 ИП в помещении"),
    (1, 3,  "Комната отдыха",        7.3,  2, 0, "Минимум 2 ИП"),
    (1, 4,  "С/У",                   4.3,  0, 0, "СП 484 п.6.2 — санузлы не защищаются"),
    (1, 5,  "Коридор",              51.2,  4, 0, "Длинный, шаг ≤9 м в два ряда, СП 484 табл.А.1"),
    (1, 6,  "Кабинет",              28.3,  2, 0, "S < 85 м², минимум 2 ИП"),
    (1, 7,  "Шоурум",               44.8,  3, 0, "S < 85 м², в один ряд с шагом ≤9 м"),
    (1, 8,  "Архив",                20.0,  3, 0, "Логика «И» по СП 484 п. 6.4.4: 2 ИП в ШС1, 1 ИП в ШС2"),
    (1, 9,  "Архив",                20.0,  3, 0, "Логика «И» по СП 484 п. 6.4.4: 2 ИП в ШС1, 1 ИП в ШС2"),
    (1, 10, "Кабинет",              26.2,  2, 0, "S < 85 м², минимум 2 ИП"),
    (1, 11, "С/У",                  10.5,  0, 0, "СП 484 п.6.2 — санузлы не защищаются"),
    (1, 12, "Бойлерная",             9.0,  0, 2, "Тепловые ИП 101-1А-А3 (запылённость, пар)"),
    (1, 13, "Подсобное помещение",   5.2,  2, 0, "Минимум 2 ИП"),
    (2, 1,  "Коридор + лестн. кл.", 67.0,  5, 0, "Длинный, шаг ≤9 м + ИП на лестничную клетку"),
    (2, 2,  "Кабинет",              18.7,  2, 0, "Минимум 2 ИП"),
    (2, 3,  "Кабинет",              17.8,  2, 0, "Минимум 2 ИП"),
    (2, 4,  "Кабинет",              12.2,  2, 0, "Минимум 2 ИП"),
    (2, 5,  "Кабинет",              12.2,  2, 0, "Минимум 2 ИП"),
    (2, 6,  "Переговорная",         35.5,  2, 0, "S < 85 м², минимум 2 ИП"),
    (2, 7,  "Приёмная",             17.8,  2, 0, "Минимум 2 ИП"),
    (2, 8,  "Комната отдыха",       11.0,  2, 0, "Минимум 2 ИП"),
    (2, 9,  "Кабинет директора",    11.0,  2, 0, "Минимум 2 ИП"),
    (2, 10, "Кухня-столовая",       26.0,  0, 2, "Тепловые ИП 101-1А-А3 (наличие пара)"),
    (2, 11, "С/У",                  10.5,  0, 0, "СП 484 п.6.2 — санузлы не защищаются"),
    (2, 12, "Серверная",            14.5,  3, 0, "Логика «И» по СП 484 п. 6.4.4: 2 ИП в ШС7, 1 ИП в ШС8"),
]


def make_calc_sheet(ws):
    ws.title = "Расчёт ИП"
    ws.merge_cells("A1:G1")
    ws["A1"] = "РАСЧЁТ КОЛИЧЕСТВА ПОЖАРНЫХ ИЗВЕЩАТЕЛЕЙ ПО ПОМЕЩЕНИЯМ"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER

    ws.merge_cells("A2:G2")
    ws["A2"] = "ИП 212-141 (дымовые), ИП 101-1А-А3 (тепловые) — СП 484.1311500.2020"
    ws["A2"].font = HEADER_FONT
    ws["A2"].alignment = CENTER

    headers = ["Этаж", "№", "Помещение", "S, м²", "ИП 212\n(дымовых)", "ИП 101\n(тепловых)", "Обоснование"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 7)

    row = 5
    f1_212 = f1_101 = f1_s = 0
    f2_212 = f2_101 = f2_s = 0

    for fl, no, room, area, q212, q101, base in CALC_ROOMS:
        ws.cell(row=row, column=1, value=fl)
        ws.cell(row=row, column=2, value=no)
        ws.cell(row=row, column=3, value=room)
        ws.cell(row=row, column=4, value=area)
        ws.cell(row=row, column=5, value=q212)
        ws.cell(row=row, column=6, value=q101)
        ws.cell(row=row, column=7, value=base)
        style_row(ws, row, 7)
        if fl == 1:
            f1_212 += q212; f1_101 += q101; f1_s += area
        else:
            f2_212 += q212; f2_101 += q101; f2_s += area
        row += 1

    # Итоги по этажам
    for label, s, q212, q101 in [
        ("Итого по 1-му этажу", f1_s, f1_212, f1_101),
        ("Итого по 2-му этажу", f2_s, f2_212, f2_101),
        ("ВСЕГО по проекту",   f1_s + f2_s, f1_212 + f2_212, f1_101 + f2_101),
    ]:
        ws.cell(row=row, column=3, value=label)
        ws.cell(row=row, column=4, value=round(s, 1))
        ws.cell(row=row, column=5, value=q212)
        ws.cell(row=row, column=6, value=q101)
        style_row(ws, row, 7, bold=True, fill=True)
        row += 1

    set_widths(ws, [8, 8, 30, 10, 14, 14, 50])
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[4].height = 36


# ============================================================================
# ЛИСТ 3 — РАСПРЕДЕЛЕНИЕ ПО ШЛЕЙФАМ
# ============================================================================

LOOPS_INFO = [
    # (ШС, назначение, состав, кол-во, ток мА)
    (1,  "1 эт. сев. (осн.)", "Архив 8 (2 ИП-А), Архив 9 (2 ИП-А), Каб.10 (2), Бойлерная (2 ИП 101), Подсобное (2)", 10, 0.45),
    (2,  "1 эт. сев. (дубль логики «И»)", "Архив 8 (1 ИП-Б), Архив 9 (1 ИП-Б) — только дубль", 2, 0.09),
    (3,  "1 эт. юж. (осн.)", "Тамбур (1), Охрана (2), Ком.отд (2), Каб.6 (2), Коридор (4), Шоурум (3)", 14, 0.63),
    (4,  "1 эт. — ИПР", "4 × ИПР 513-10 у эвакуационных выходов", 4, 0.20),
    (5,  "1 эт. — оповещение (через РИП)", "5 Соната-М + 3 Призма 201 + 3 Табло «Выход»", 11, "—"),
    (6,  "2 эт. юж.", "Каб 2-5 (8), Переговорная (2), Приёмная (2)", 12, 0.54),
    (7,  "2 эт. центр+сев. (осн.)", "Ком.отд (2), Каб.дир (2), Кухня (2 ИП 101), Серверная (2 ИП-А), Коридор (5)", 13, 0.59),
    (8,  "2 эт. (дубль логики «И»)", "Серверная (1 ИП-Б) — только дубль", 1, 0.05),
    (9,  "2 эт. — ИПР", "2 × ИПР 513-10", 2, 0.10),
    (10, "2 эт. — оповещение (через РИП)", "3 Соната-М + 2 Призма 201 + 1 Табло «Выход»", 6, "—"),
    (11, "РЕЗЕРВ", "Свободный шлейф (для расширения)", 0, "—"),
    (12, "РЕЗЕРВ", "Свободный шлейф (для расширения)", 0, "—"),
]


def make_loops_sheet(ws):
    ws.title = "Шлейфы"
    ws.merge_cells("A1:E1")
    ws["A1"] = "РАСПРЕДЕЛЕНИЕ УСТРОЙСТВ ПО ШЛЕЙФАМ ППКП ГРАНИТ-12"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER

    ws.merge_cells("A2:E2")
    ws["A2"] = "Лимит шлейфа Гранит-12: 1,5 мА (СП 484 п.6.2.4 для логики «И»: ИП в разных шлейфах)"
    ws["A2"].font = HEADER_FONT
    ws["A2"].alignment = CENTER

    headers = ["№ ШС", "Назначение", "Состав", "Устройств", "Ток дежурки, мА"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 5)

    row = 5
    for loop, purpose, content, qty, current in LOOPS_INFO:
        ws.cell(row=row, column=1, value=f"ШС {loop}")
        ws.cell(row=row, column=2, value=purpose)
        ws.cell(row=row, column=3, value=content)
        ws.cell(row=row, column=4, value=qty)
        ws.cell(row=row, column=5, value=current)
        style_row(ws, row, 5)
        row += 1

    set_widths(ws, [10, 32, 60, 14, 16])
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[4].height = 32


# ============================================================================
# ЛИСТ 4 — КАБЕЛЬНЫЙ ЖУРНАЛ
# ============================================================================

CABLE_LOG = [
    # (№, участок, тип кабеля, длина, способ прокладки, примечание)
    ("1", "ППКП — ШС 1 (1 эт. сев.)",
        "КПСВВнг(А)-LS 1×2×0,5", 95, "В МР-15", "10 устройств"),
    ("2", "ППКП — ШС 2 (1 эт. сев. дубль логики «И»)",
        "КПСВВнг(А)-LS 1×2×0,5", 40, "В МР-15", "2 ИП-Б архивов (только дубль)"),
    ("3", "ППКП — ШС 3 (1 эт. юж.)",
        "КПСВВнг(А)-LS 1×2×0,5", 95, "В МР-15", "14 устройств"),
    ("4", "ППКП — ШС 4 (ИПР 1 эт.)",
        "КПСВВнг(А)-LS 1×2×0,5", 55, "В МР-15", "4 ИПР"),
    ("5", "ППКП — ШС 5 (оповещение 1 эт.) через РИП-12",
        "КПСВВнг(А)-LS 1×2×0,75", 95, "В МР-15", "11 устройств"),
    ("6", "ППКП — стояк межэтажный — ШС 6 (2 эт. юж.)",
        "КПСВВнг(А)-LS 1×2×0,5", 95, "В МР-15, стояк", "12 устройств"),
    ("7", "ППКП — стояк — ШС 7 (2 эт. центр+сев.)",
        "КПСВВнг(А)-LS 1×2×0,5", 115, "В МР-15, стояк", "13 устройств"),
    ("8", "ППКП — стояк — ШС 8 (2 эт. дубль логики «И»)",
        "КПСВВнг(А)-LS 1×2×0,5", 20, "В МР-15, стояк", "1 устройство (ИП-Б серверной)"),
    ("9", "ППКП — стояк — ШС 9 (ИПР 2 эт.)",
        "КПСВВнг(А)-LS 1×2×0,5", 30, "В МР-15", "2 ИПР"),
    ("10", "РИП-12 — стояк — ШС 10 (оповещение 2 эт.)",
        "КПСВВнг(А)-LS 1×2×0,75", 65, "В МР-15", "6 устройств"),
    ("11", "Щит — ППКП (силовое питание)",
        "ВВГнг(А)-LS 3×1,5", 25, "В МР-15", "Автомат защиты 6 А"),
    ("12", "Щит — РИП-12 (силовое питание)",
        "ВВГнг(А)-LS 3×1,5", 5, "В МР-15", "Автомат защиты 6 А"),
]


def make_cable_sheet(ws):
    ws.title = "Кабельный журнал"
    ws.merge_cells("A1:F1")
    ws["A1"] = "КАБЕЛЬНЫЙ ЖУРНАЛ"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER

    ws.merge_cells("A2:F2")
    ws["A2"] = "Все кабели — нг(А)-LS по СП 6.13130 п.4.2; защита трасс в МР-15"
    ws["A2"].font = HEADER_FONT
    ws["A2"].alignment = CENTER

    headers = ["№", "Участок (от — до)", "Марка кабеля", "Длина,\nм",
               "Способ\nпрокладки", "Примечание"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 6)

    row = 5
    total = {}
    for no, span, cab, length, way, note in CABLE_LOG:
        ws.cell(row=row, column=1, value=no)
        ws.cell(row=row, column=2, value=span)
        ws.cell(row=row, column=3, value=cab)
        ws.cell(row=row, column=4, value=length)
        ws.cell(row=row, column=5, value=way)
        ws.cell(row=row, column=6, value=note)
        style_row(ws, row, 6)
        total[cab] = total.get(cab, 0) + length
        row += 1

    # Сводка по типам кабеля
    row += 1
    ws.cell(row=row, column=2, value="СВОДКА ПО ТИПАМ КАБЕЛЯ (с запасом 20%)")
    style_row(ws, row, 6, bold=True, fill=True)
    row += 1
    for cab, length in total.items():
        ws.cell(row=row, column=2, value=cab)
        ws.cell(row=row, column=4, value=round(length * 1.20))
        ws.cell(row=row, column=5, value="м")
        style_row(ws, row, 6, bold=True)
        row += 1

    # Металлорукав отдельно
    mr_total = sum(length for _, _, _, length, way, _ in CABLE_LOG if "МР" in way)
    ws.cell(row=row, column=2, value="Металлорукав МР-15 (∑ × 1,2)")
    ws.cell(row=row, column=4, value=round(mr_total * 1.20))
    ws.cell(row=row, column=5, value="м")
    style_row(ws, row, 6, bold=True)

    set_widths(ws, [6, 45, 25, 10, 18, 25])
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[4].height = 36


# ============================================================================
# MAIN
# ============================================================================

def main():
    wb = Workbook()
    make_specification_sheet(wb.active)
    make_calc_sheet(wb.create_sheet())
    make_loops_sheet(wb.create_sheet())
    make_cable_sheet(wb.create_sheet())
    wb.save("specification.xlsx")
    print("✓ Сохранено: specification.xlsx")
    print(f"  Листов: {len(wb.sheetnames)}")
    for name in wb.sheetnames:
        print(f"    - {name}")


if __name__ == "__main__":
    main()
