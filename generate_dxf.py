# -*- coding: utf-8 -*-
"""
DXF-генератор плана автоматической пожарной сигнализации (АПС).

Создаёт два файла:
    fire_alarm_floor1.dxf — план оборудования 1-го этажа
    fire_alarm_floor2.dxf — план оборудования 2-го этажа

Координаты оборудования берутся из coordinates.py.
В DXF единицей измерения принят миллиметр (стандарт строительных чертежей).
Реальные координаты (метры) автоматически умножаются на 1000.
Размеры символов даны для масштаба 1:100 (например, Ø5 мм на печати = 500 мм в чертеже).
"""
import math
import os
import ezdxf
from ezdxf.enums import TextEntityAlignment

import coordinates as C


# ============================================================================
# КОНСТАНТЫ
# ============================================================================

# Единицы: метры -> миллиметры (для масштаба 1:100)
SCALE = 1000

# Размеры символов оборудования (мм, в реальном чертеже при М1:100)
D_IP212 = 500       # диаметр ИП 212-141, мм -> 5 мм на печати
D_IP101 = 500       # диаметр ИП 101 (тепловой)
SIDE_IPR = 600      # сторона квадрата ИПР, мм -> 6 мм
TRI_SONATA = 700    # высота треугольника Соната-М
D_PRIZMA = 400      # диаметр Призмы 201
EXIT_W, EXIT_H = 1000, 500   # прямоугольник табло «Выход» (10×5 мм)
PPKP_W, PPKP_H = 1500, 1000  # ППКП Гранит-12 (15×10 мм)
RIP_W, RIP_H = 1200, 800     # РИП-12

TEXT_H = 350        # высота шрифта подписей
TAG_OFFSET = 700    # смещение подписи от символа

# Слои и их цвета (AutoCAD Color Index)
LAYERS = {
    "АПС_Архитектура": {"color": 8,  "linetype": "CONTINUOUS"},  # серый — растровая подложка
    "АПС_План":       {"color": 7,  "linetype": "CONTINUOUS"},  # белый/чёрный
    "АПС_ИП":         {"color": 1,  "linetype": "CONTINUOUS"},  # красный
    "АПС_ИПР":        {"color": 5,  "linetype": "CONTINUOUS"},  # синий
    "АПС_Оповещение": {"color": 2,  "linetype": "CONTINUOUS"},  # жёлтый
    "АПС_Призма":     {"color": 6,  "linetype": "CONTINUOUS"},  # фиолетовый
    "АПС_Выход":      {"color": 3,  "linetype": "CONTINUOUS"},  # зелёный
    "АПС_ППКП":       {"color": 4,  "linetype": "CONTINUOUS"},  # голубой
    "АПС_Шлейфы":     {"color": 8,  "linetype": "DASHDOT"},     # серый штрих-пункт
    "АПС_Подписи":    {"color": 7,  "linetype": "CONTINUOUS"},  # текст
}

# Цвета шлейфов (для разноцветной трассировки на слое АПС_Шлейфы)
LOOP_COLORS = {
    1: 1,    # красный
    2: 30,   # оранжевый
    3: 2,    # жёлтый
    4: 5,    # синий
    5: 6,    # фиолетовый
    6: 3,    # зелёный
    7: 4,    # голубой
    8: 200,  # розовый
    9: 21,   # тёмно-красный
    10: 140, # светло-синий
}

# Габариты этажа (м -> мм)
FLOOR_W = 30.0 * SCALE
FLOOR_H = 21.0 * SCALE

# Архитектурная подложка (растр архитектурного плана).
# Файлы JPG расположены в корне репозитория; путь относительный, чтобы DXF
# открывался и в директории проекта.
FLOOR_BG = {
    1: "1 этаж .jpg",
    2: "2 этаж .jpg",
}
# Размеры подложки совмещаются с FLOOR_W × FLOOR_H. Если архитектурный план
# содержит экспликацию справа, при необходимости откорректировать BG_WIDTH/HEIGHT
# и BG_OFFSET для точной привязки осей.
BG_WIDTH = FLOOR_W
BG_HEIGHT = FLOOR_H
BG_OFFSET = (0, 0)


# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================

def m(value):
    """Перевод метров в миллиметры (масштаб чертежа)."""
    return value * SCALE


def setup_document():
    """Создание нового DXF-документа со слоями и типами линий."""
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM

    # Создание слоёв
    for name, props in LAYERS.items():
        if name not in doc.layers:
            doc.layers.add(name, color=props["color"], linetype=props["linetype"])

    return doc


def attach_floor_background(doc, msp, floor_no):
    """Вставляет архитектурный план как растровую подложку (DXF IMAGE).

    Изображение `1 этаж .jpg` / `2 этаж .jpg` берётся из текущей директории
    проекта. Размеры подложки заданы FLOOR_W × FLOOR_H (мм). При необходимости
    подложка калибруется в AutoCAD после открытия чертежа.
    """
    bg = FLOOR_BG.get(floor_no)
    if not bg or not os.path.exists(bg):
        print(f"  [предупреждение] архитектурная подложка не найдена: {bg}")
        return

    try:
        from PIL import Image
        with Image.open(bg) as im:
            px_w, px_h = im.size
    except Exception:
        px_w, px_h = 1280, 839  # ориентировочные размеры исходных JPG

    image_def = doc.add_image_def(filename=bg, size_in_pixel=(px_w, px_h))
    msp.add_image(
        insert=BG_OFFSET,
        size_in_units=(BG_WIDTH, BG_HEIGHT),
        image_def=image_def,
        rotation=0,
        dxfattribs={"layer": "АПС_Архитектура"},
    )


def draw_room_outline(msp, x, y, w, h, name=""):
    """Рисует контур помещения (прямоугольник) и подпись."""
    msp.add_lwpolyline(
        [(x, y), (x + w, y), (x + w, y + h), (x, y + h), (x, y)],
        dxfattribs={"layer": "АПС_План"},
    )
    if name:
        msp.add_text(
            name,
            dxfattribs={"layer": "АПС_План", "height": TEXT_H * 1.2},
        ).set_placement((x + w / 2, y + h / 2), align=TextEntityAlignment.MIDDLE_CENTER)


# ============================================================================
# ОТРИСОВКА СИМВОЛОВ ОБОРУДОВАНИЯ
# ============================================================================

def draw_ip212(msp, x, y, label):
    """Дымовой ИП 212-141: окружность + крест внутри."""
    r = D_IP212 / 2
    msp.add_circle((x, y), r, dxfattribs={"layer": "АПС_ИП"})
    # крест внутри (две диагонали)
    s = r * 0.7
    msp.add_line((x - s, y - s), (x + s, y + s), dxfattribs={"layer": "АПС_ИП"})
    msp.add_line((x - s, y + s), (x + s, y - s), dxfattribs={"layer": "АПС_ИП"})
    msp.add_text(label, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}).set_placement(
        (x + r + 100, y + r + 100), align=TextEntityAlignment.BOTTOM_LEFT
    )


def draw_ip101(msp, x, y, label):
    """Тепловой ИП 101-1А-А3: окружность + буква T внутри."""
    r = D_IP101 / 2
    msp.add_circle((x, y), r, dxfattribs={"layer": "АПС_ИП"})
    msp.add_text(
        "T", dxfattribs={"layer": "АПС_ИП", "height": r * 1.2, "style": "Standard"}
    ).set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)
    msp.add_text(label, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}).set_placement(
        (x + r + 100, y + r + 100), align=TextEntityAlignment.BOTTOM_LEFT
    )


def draw_ipr(msp, x, y, label):
    """ИПР 513-10: квадрат + буква Р внутри."""
    s = SIDE_IPR / 2
    msp.add_lwpolyline(
        [(x - s, y - s), (x + s, y - s), (x + s, y + s), (x - s, y + s), (x - s, y - s)],
        dxfattribs={"layer": "АПС_ИПР"},
    )
    msp.add_text(
        "Р", dxfattribs={"layer": "АПС_ИПР", "height": s * 1.2}
    ).set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)
    msp.add_text(label, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}).set_placement(
        (x + s + 100, y + s + 100), align=TextEntityAlignment.BOTTOM_LEFT
    )


def draw_sonata(msp, x, y, label):
    """Соната-М: равносторонний треугольник."""
    h = TRI_SONATA
    a = h / math.sin(math.radians(60))  # сторона
    pts = [
        (x, y + 2 * h / 3),
        (x - a / 2, y - h / 3),
        (x + a / 2, y - h / 3),
        (x, y + 2 * h / 3),
    ]
    msp.add_lwpolyline(pts, dxfattribs={"layer": "АПС_Оповещение"})
    msp.add_text(label, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}).set_placement(
        (x + a / 2 + 100, y + h / 3), align=TextEntityAlignment.BOTTOM_LEFT
    )


def draw_prizma(msp, x, y, label):
    """Призма 201: окружность с штриховкой (4 короткие диагональные линии)."""
    r = D_PRIZMA / 2
    msp.add_circle((x, y), r, dxfattribs={"layer": "АПС_Призма"})
    # Штриховка — 3 параллельные линии под 45°
    for k in (-0.6, 0, 0.6):
        dx = r * 0.9
        msp.add_line(
            (x - dx, y + k * r - dx),
            (x + dx, y + k * r + dx),
            dxfattribs={"layer": "АПС_Призма"},
        )
    msp.add_text(label, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}).set_placement(
        (x + r + 100, y + r + 100), align=TextEntityAlignment.BOTTOM_LEFT
    )


def draw_exit_sign(msp, x, y, label):
    """Табло «Выход»: прямоугольник с надписью."""
    w, h = EXIT_W, EXIT_H
    msp.add_lwpolyline(
        [(x - w / 2, y - h / 2), (x + w / 2, y - h / 2),
         (x + w / 2, y + h / 2), (x - w / 2, y + h / 2),
         (x - w / 2, y - h / 2)],
        dxfattribs={"layer": "АПС_Выход"},
    )
    msp.add_text(
        "ВЫХОД", dxfattribs={"layer": "АПС_Выход", "height": h * 0.5}
    ).set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)
    if label:
        msp.add_text(
            label, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}
        ).set_placement((x, y - h), align=TextEntityAlignment.TOP_CENTER)


def draw_ppkp(msp, x, y):
    """ППКП Гранит-12: прямоугольник с надписью."""
    w, h = PPKP_W, PPKP_H
    msp.add_lwpolyline(
        [(x - w / 2, y - h / 2), (x + w / 2, y - h / 2),
         (x + w / 2, y + h / 2), (x - w / 2, y + h / 2),
         (x - w / 2, y - h / 2)],
        dxfattribs={"layer": "АПС_ППКП"},
    )
    msp.add_text(
        "ППКП\nГранит-12", dxfattribs={"layer": "АПС_ППКП", "height": h * 0.3}
    ).set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)


def draw_rip(msp, x, y):
    """РИП-12: прямоугольник с надписью."""
    w, h = RIP_W, RIP_H
    msp.add_lwpolyline(
        [(x - w / 2, y - h / 2), (x + w / 2, y - h / 2),
         (x + w / 2, y + h / 2), (x - w / 2, y + h / 2),
         (x - w / 2, y - h / 2)],
        dxfattribs={"layer": "АПС_ППКП"},
    )
    msp.add_text(
        "РИП-12", dxfattribs={"layer": "АПС_ППКП", "height": h * 0.4}
    ).set_placement((x, y), align=TextEntityAlignment.MIDDLE_CENTER)


# ============================================================================
# ТРАССИРОВКА ШЛЕЙФОВ
# ============================================================================

def draw_loop(msp, points, loop_no):
    """Рисует полилинию шлейфа по списку точек."""
    if len(points) < 2:
        return
    color = LOOP_COLORS.get(loop_no, 8)
    msp.add_lwpolyline(
        points,
        dxfattribs={
            "layer": "АПС_Шлейфы",
            "color": color,
            "linetype": "DASHDOT",
        },
    )


def build_loops(ip212_list, ip101_list, ipr_list, ppkp_pos):
    """Группирует устройства по шлейфам, возвращает словарь {loop: [(x,y), ...]}."""
    loops = {}
    for x, y, room, loop, tag in ip212_list + ip101_list:
        loops.setdefault(loop, []).append((m(x), m(y)))
    for x, y, place, loop in ipr_list:
        loops.setdefault(loop, []).append((m(x), m(y)))

    # Добавляем ППКП в начало каждого шлейфа (для наглядности трассы)
    px, py = m(ppkp_pos[0]), m(ppkp_pos[1])
    for loop in loops:
        # Сортируем точки по расстоянию от ППКП — простая эвристика трассировки
        loops[loop].sort(key=lambda p: (p[0] - px) ** 2 + (p[1] - py) ** 2)
        loops[loop].insert(0, (px, py))
    return loops


# ============================================================================
# СТЕМП И РАМКА
# ============================================================================

def draw_titleblock(msp, floor_no):
    """Рисует рамку чертежа и основную надпись по ГОСТ 21.1101 (форма 3)."""
    # Внешняя рамка
    margin = 500
    msp.add_lwpolyline(
        [(-margin, -margin),
         (FLOOR_W + margin, -margin),
         (FLOOR_W + margin, FLOOR_H + margin),
         (-margin, FLOOR_H + margin),
         (-margin, -margin)],
        dxfattribs={"layer": "АПС_План"},
    )
    # Заголовок
    title = f"План АПС {floor_no}-го этажа. М 1:100"
    msp.add_text(
        title, dxfattribs={"layer": "АПС_Подписи", "height": 600}
    ).set_placement(
        (FLOOR_W / 2, FLOOR_H + margin + 800), align=TextEntityAlignment.BOTTOM_CENTER
    )

    # ---- Основная надпись (штамп) по ГОСТ 21.1101 форма 3 ----
    # Размещаем штамп в правом нижнем углу под рамкой
    sx = FLOOR_W + margin - 18500   # 185 мм при М1:100 = 18500 мм в модели
    sy = -margin - 5500             # 55 мм высота штампа
    sw, sh = 18500, 5500
    # Внешний контур
    msp.add_lwpolyline(
        [(sx, sy), (sx + sw, sy), (sx + sw, sy + sh), (sx, sy + sh), (sx, sy)],
        dxfattribs={"layer": "АПС_План"},
    )
    # Горизонтальные линии (форма 3 имеет 4 ряда)
    for k in (1, 2, 3, 4):
        y = sy + sh * k / 5
        msp.add_line((sx, y), (sx + sw, y), dxfattribs={"layer": "АПС_План"})
    # Вертикальные разделители — упрощённо: одна линия 65 мм от левого края
    msp.add_line((sx + 6500, sy), (sx + 6500, sy + sh),
                 dxfattribs={"layer": "АПС_План"})

    # Текстовое заполнение
    fields = [
        (0.5, 4.5, "Шифр: АПС-АДМ-01"),
        (0.5, 3.5, "Стадия: РД"),
        (0.5, 2.5, f"Лист: {floor_no} плана"),
        (0.5, 1.5, "Листов: 2"),
        (0.5, 0.5, "Заказчик: __________"),
        (7.0, 4.5, "Автоматическая пожарная сигнализация"),
        (7.0, 3.5, "Административное здание, S = 491,1 м²"),
        (7.0, 2.5, f"План АПС {floor_no}-го этажа. М 1:100"),
        (7.0, 1.5, "Разработал: ___________"),
        (7.0, 0.5, "Проверил:   ___________"),
    ]
    for fx, fy, txt in fields:
        msp.add_text(
            txt, dxfattribs={"layer": "АПС_Подписи", "height": 320}
        ).set_placement(
            (sx + fx * 1000, sy + fy * sh / 5),
            align=TextEntityAlignment.MIDDLE_LEFT,
        )


def draw_legend(msp, x0, y0):
    """Условные обозначения — справа от плана."""
    items = [
        ("ИП 212-141 (дымовой)",   "ip212"),
        ("ИП 101-1А-А3 (тепловой)", "ip101"),
        ("ИПР 513-10",              "ipr"),
        ("Соната-М",                "sonata"),
        ("Призма 201",              "prizma"),
        ("Табло «Выход»",           "exit"),
        ("ППКП Гранит-12",          "ppkp"),
        ("Шлейф сигнализации",      "loop"),
    ]
    msp.add_text(
        "Условные обозначения:",
        dxfattribs={"layer": "АПС_Подписи", "height": 500},
    ).set_placement((x0, y0 + 1500), align=TextEntityAlignment.BOTTOM_LEFT)

    step = 1200
    for i, (text, kind) in enumerate(items):
        cy = y0 - i * step
        cx = x0 + 500
        if kind == "ip212":
            draw_ip212(msp, cx, cy, "")
        elif kind == "ip101":
            draw_ip101(msp, cx, cy, "")
        elif kind == "ipr":
            draw_ipr(msp, cx, cy, "")
        elif kind == "sonata":
            draw_sonata(msp, cx, cy, "")
        elif kind == "prizma":
            draw_prizma(msp, cx, cy, "")
        elif kind == "exit":
            draw_exit_sign(msp, cx, cy, "")
        elif kind == "ppkp":
            draw_ppkp(msp, cx, cy)
        elif kind == "loop":
            msp.add_lwpolyline(
                [(cx - 500, cy), (cx + 500, cy)],
                dxfattribs={"layer": "АПС_Шлейфы", "linetype": "DASHDOT"},
            )
        msp.add_text(
            text, dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}
        ).set_placement((cx + 1500, cy), align=TextEntityAlignment.MIDDLE_LEFT)


# ============================================================================
# ГЕНЕРАЦИЯ ЭТАЖА
# ============================================================================

def draw_riser(msp, x, y):
    """Отметка межэтажного стояка (для 2 этажа — точка ввода шлейфов от ППКП)."""
    r = 400
    msp.add_circle((x, y), r, dxfattribs={"layer": "АПС_ППКП"})
    msp.add_circle((x, y), r * 0.4, dxfattribs={"layer": "АПС_ППКП"})
    msp.add_text(
        "Стояк к ППКП", dxfattribs={"layer": "АПС_Подписи", "height": TEXT_H}
    ).set_placement((x + r + 200, y), align=TextEntityAlignment.MIDDLE_LEFT)


def generate_floor(floor_no, ip212, ip101, ipr, sonata, prizma, exit_signs,
                   ppkp_pos=None, rip_pos=None, riser_pos=None,
                   output="fire_alarm.dxf"):
    """Генерирует DXF-файл для одного этажа."""
    doc = setup_document()
    msp = doc.modelspace()

    # 0. Архитектурная подложка (растр плана этажа)
    attach_floor_background(doc, msp, floor_no)

    # 1. Внешний контур этажа (упрощённый прямоугольник — рамка над подложкой)
    msp.add_lwpolyline(
        [(0, 0), (FLOOR_W, 0), (FLOOR_W, FLOOR_H), (0, FLOOR_H), (0, 0)],
        dxfattribs={"layer": "АПС_План"},
    )

    # 2. Подписи помещений (центры) — упрощённо, без отрисовки внутренних стен
    msp.add_text(
        f"Этаж {floor_no} — план оборудования АПС",
        dxfattribs={"layer": "АПС_Подписи", "height": 500},
    ).set_placement((m(0.5), FLOOR_H - 700), align=TextEntityAlignment.BOTTOM_LEFT)

    # 3. ППКП / РИП (только 1 эт.) или стояк (только 2 эт.)
    if ppkp_pos is not None:
        draw_ppkp(msp, m(ppkp_pos[0]), m(ppkp_pos[1]))
    if rip_pos is not None:
        draw_rip(msp, m(rip_pos[0]), m(rip_pos[1]))
    if riser_pos is not None:
        draw_riser(msp, m(riser_pos[0]), m(riser_pos[1]))

    # 4. Оборудование
    for x, y, room, loop, tag in ip212:
        label = f"П{room}/ШС{loop}" + (f"-{tag}" if tag else "")
        draw_ip212(msp, m(x), m(y), label)

    for x, y, room, loop, tag in ip101:
        label = f"П{room}/ШС{loop}"
        draw_ip101(msp, m(x), m(y), label)

    for x, y, place, loop in ipr:
        draw_ipr(msp, m(x), m(y), f"ШС{loop}")

    for x, y, place, loop in sonata:
        draw_sonata(msp, m(x), m(y), f"ШС{loop}")

    for x, y, place, loop in prizma:
        draw_prizma(msp, m(x), m(y), f"ШС{loop}")

    for x, y, place, loop in exit_signs:
        draw_exit_sign(msp, m(x), m(y), f"ШС{loop}")

    # 5. Трассировка шлейфов извещателей (ИП + ИПР)
    # Точка ввода = ППКП (1 эт.) или стояк (2 эт.)
    entry = ppkp_pos if ppkp_pos is not None else riser_pos
    if entry is not None:
        loops = build_loops(ip212, ip101, ipr, entry)
        for loop_no, pts in loops.items():
            draw_loop(msp, pts, loop_no)

    # 6. Рамка и заголовок
    draw_titleblock(msp, floor_no)

    # 7. Условные обозначения справа
    draw_legend(msp, FLOOR_W + 1500, FLOOR_H - 1500)

    # Сохранение
    doc.saveas(output)
    print(f"✓ Сохранено: {output}")


# ============================================================================
# MAIN
# ============================================================================

def main():
    # 1 этаж — с ППКП и РИП
    generate_floor(
        floor_no=1,
        ip212=C.FLOOR1_IP212,
        ip101=C.FLOOR1_IP101,
        ipr=C.FLOOR1_IPR,
        sonata=C.FLOOR1_SONATA,
        prizma=C.FLOOR1_PRIZMA,
        exit_signs=C.FLOOR1_EXIT,
        ppkp_pos=(C.PPKP_GRANIT12[0], C.PPKP_GRANIT12[1]),
        rip_pos=(C.RIP_12[0], C.RIP_12[1]),
        output="fire_alarm_floor1.dxf",
    )

    # 2 этаж — ППКП не дублируется, вместо него отметка стояка над ППКП 1 этажа
    riser_point_2 = (C.PPKP_GRANIT12[0], C.PPKP_GRANIT12[1])  # та же точка по плану
    generate_floor(
        floor_no=2,
        ip212=C.FLOOR2_IP212,
        ip101=C.FLOOR2_IP101,
        ipr=C.FLOOR2_IPR,
        sonata=C.FLOOR2_SONATA,
        prizma=C.FLOOR2_PRIZMA,
        exit_signs=C.FLOOR2_EXIT,
        ppkp_pos=None,
        rip_pos=None,
        riser_pos=riser_point_2,
        output="fire_alarm_floor2.dxf",
    )


if __name__ == "__main__":
    main()
