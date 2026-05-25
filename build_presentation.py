#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка ВКР-презентации АО «Унистрой» в стиле шаблона
  «Шаблон презентации ВКР ИиИ_2026 (2).pptx»
  (КФУ, ИУЭФ, кафедра инноваций и инвестиций).

Принцип: открываем сам шаблон, у которого уже заданы:
  - размер слайда 10″ × 5,625″ (16:9);
  - тёмно-синий титул и финал с фирменными изображениями;
  - вертикальная полоса слева и угловой логотип на каждом контентном слайде;
  - правильные шрифты PT Sans / Times New Roman.

Затем:
  1) на титульном слайде — заполняем тему ВКР, ФИО / группу / руководителя;
  2) на слайдах 2–9 — оставляем фирменные декоративные элементы и заголовок,
     добавляем инфографику Унистрой (плитки, KPI, схемы);
  3) финальный слайд (10) — оставляем «Спасибо за внимание!».

Результат сохраняется в VKR_Unistroy_Presentation.pptx.

Запуск:
    python3 build_presentation.py
"""

from __future__ import annotations

import os
from copy import deepcopy

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "Шаблон презентации ВКР ИиИ_2026 (2).pptx")
OUTPUT = os.path.join(HERE, "VKR_Unistroy_Presentation.pptx")

# Палитра (см. presentation_style.md)
DARK_BLUE = RGBColor(0x17, 0x36, 0x5D)   # #17365D — основной фирменный
BLUE2     = RGBColor(0x1F, 0x49, 0x7D)   # #1F497D — вспомогательный
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
BLACK     = RGBColor(0x00, 0x00, 0x00)
GRAY      = RGBColor(0x88, 0x88, 0x88)
GRAY_BG   = RGBColor(0xF2, 0xF2, 0xF2)
GRAY_BD   = RGBColor(0xBF, 0xBF, 0xBF)
ACCENT_GREEN = RGBColor(0x70, 0xAD, 0x47)  # для ▲
ACCENT_RED   = RGBColor(0xC0, 0x50, 0x4D)  # для ▼

FONT_MAIN = "PT Sans"

# Геометрия слайда (10″ × 5,625″ = 25.4 × 14.29 см)
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)

# Зона контента (под заголовком, справа от вертикальной полосы шаблона)
CONTENT_LEFT = Inches(1.10)
CONTENT_TOP  = Inches(0.95)
CONTENT_W    = Inches(8.78)
CONTENT_BOT  = Inches(5.10)
CONTENT_H    = CONTENT_BOT - CONTENT_TOP


# ---------------------------------------------------------------------------
# Низкоуровневые помощники
# ---------------------------------------------------------------------------

def remove_shape(shp):
    """Удалить shape со слайда (через XML — у python-pptx нет прямого API)."""
    sp = shp._element
    sp.getparent().remove(sp)


def set_solid_fill(shape, rgb: RGBColor):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb


def set_no_line(shape):
    line = shape.line
    line.fill.background()


def set_line(shape, rgb: RGBColor, width_pt: float = 1.0):
    shape.line.color.rgb = rgb
    shape.line.width = Pt(width_pt)


def style_run(run, *, font=FONT_MAIN, size=None, bold=None,
              color: RGBColor | None = None, italic=None):
    run.font.name = font
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color


def set_text(tf, text: str, *, size=14, bold=False, color=BLACK,
             align=PP_ALIGN.LEFT, font=FONT_MAIN, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.15):
    """Полностью переписать текст текст-фрейма в едином стиле."""
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    style_run(r, font=font, size=size, bold=bold, color=color)
    return p, r


def add_paragraph(tf, text: str, *, size=14, bold=False, color=BLACK,
                  align=PP_ALIGN.LEFT, font=FONT_MAIN, line_spacing=1.15,
                  space_before=None, space_after=None, level=0):
    p = tf.add_paragraph()
    p.alignment = align
    p.line_spacing = line_spacing
    p.level = level
    if space_before is not None:
        p.space_before = Pt(space_before)
    if space_after is not None:
        p.space_after = Pt(space_after)
    r = p.add_run()
    r.text = text
    style_run(r, font=font, size=size, bold=bold, color=color)
    return p, r


def add_text_box(slide, left, top, width, height, *, fill=None,
                 line=None, line_w=1.0):
    """Прямоугольник со встроенным text frame."""
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    if fill is None:
        box.fill.background()
    else:
        set_solid_fill(box, fill)
    if line is None:
        set_no_line(box)
    else:
        set_line(box, line, line_w)
    box.shadow.inherit = False
    return box


def add_rounded_box(slide, left, top, width, height, *, fill=None,
                    line=None, line_w=1.0, corner=0.08):
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    # установить радиус скругления (0..0.5 от меньшей стороны)
    try:
        sp = box.element
        avLst = sp.find(".//" + qn("a:avLst"))
        if avLst is not None:
            for child in list(avLst):
                avLst.remove(child)
            gd = etree.SubElement(avLst, qn("a:gd"))
            gd.set("name", "adj")
            gd.set("fmla", f"val {int(corner*50000)}")
    except Exception:
        pass
    if fill is None:
        box.fill.background()
    else:
        set_solid_fill(box, fill)
    if line is None:
        set_no_line(box)
    else:
        set_line(box, line, line_w)
    box.shadow.inherit = False
    return box


def add_arrow(slide, left, top, width, height, *,
              fill=DARK_BLUE, line=None):
    """Стрелка вправо (используется в схемах)."""
    arr = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, left, top, width, height)
    set_solid_fill(arr, fill)
    if line is None:
        set_no_line(arr)
    else:
        set_line(arr, line)
    arr.shadow.inherit = False
    return arr


def add_chevron(slide, left, top, width, height, *, fill=DARK_BLUE):
    chev = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, left, top, width, height)
    set_solid_fill(chev, fill)
    set_no_line(chev)
    chev.shadow.inherit = False
    return chev


def add_circle(slide, left, top, size, *, fill=DARK_BLUE, line=None):
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    set_solid_fill(c, fill)
    if line is None:
        set_no_line(c)
    else:
        set_line(c, line)
    c.shadow.inherit = False
    return c


# ---------------------------------------------------------------------------
# Очистка контентных слайдов (2–9) — оставляем только декор шаблона
# ---------------------------------------------------------------------------

# В шаблоне на каждом контентном слайде есть:
#   - картинка-полоса слева  (Google Shape;… ~0.98″ ширина, на всю высоту);
#   - картинка-логотип        (~0.75 × 0.63 в верхнем углу полосы);
#   - заголовок-плейсхолдер   (с шаблонным текстом «АКТУАЛЬНОСТЬ» и т.п.);
#   - текст-бокс с номером слайда (правый нижний угол);
#   - картинка-«рисунок 21/31» (≈0.85×0.85 у нижней части полосы).
# Сохраняем картинки и номер, чистим/перенастраиваем заголовок,
# а все прочие подложки/фигуры удаляем.

def clean_content_slide(slide, *, new_title: str):
    """Очистить контентный слайд от служебного текста-плейсхолдера,
    выставить новый заголовок. Возвращает оставшийся плейсхолдер заголовка
    (для возможных дальнейших настроек) или None."""
    title_shp = None
    to_remove = []
    for shp in list(slide.shapes):
        if shp.has_text_frame:
            txt = shp.text_frame.text.strip().upper()
            # это плейсхолдер заголовка — оставляем, но переписываем текст
            if any(key in txt for key in (
                "АКТУАЛЬНОСТЬ", "ЦЕЛИ И ЗАДАЧИ", "РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ",
                "ВЫВОДЫ",
            )):
                title_shp = shp
                # Переписать заголовок целиком
                set_text(
                    shp.text_frame, new_title,
                    size=24, bold=True, color=DARK_BLUE,
                    align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                    line_spacing=1.0,
                )
                # Скорректировать рамку заголовка под единый размер
                shp.left = Inches(1.10)
                shp.top = Inches(0.18)
                shp.width = Inches(8.78)
                shp.height = Inches(0.65)
                continue
            # это номер слайда — оставляем
            if txt.isdigit():
                continue
        # картинки — оставляем все
        if shp.shape_type and str(shp.shape_type).startswith("PICTURE"):
            continue
        # прочие текстовые подписи (если есть) — удаляем
        if shp.has_text_frame:
            to_remove.append(shp)

    for shp in to_remove:
        remove_shape(shp)

    return title_shp


# ---------------------------------------------------------------------------
# Слайд 1 — титульный
# ---------------------------------------------------------------------------

THEME_TITLE = (
    "МОДЕЛИРОВАНИЕ И ОПТИМИЗАЦИЯ БИЗНЕС-ПРОЦЕССОВ\n"
    "В ДЕВЕЛОПЕРСКОЙ КОМПАНИИ\n"
    "(на примере АО «УНИСТРОЙ»)"
)
AUTHOR_LINES = [
    "ФИО: ___________________________",
    "Группа: ___________",
    "Научный руководитель: ___________________________",
]


def fill_title_slide(slide):
    """Заполнить плейсхолдеры титульного слайда нашими данными."""
    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        txt = shp.text_frame.text.strip()

        if txt == "НАЗВАНИЕ ТЕМЫ":
            tf = shp.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            lines = THEME_TITLE.split("\n")
            for i, line in enumerate(lines):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.CENTER
                p.line_spacing = 1.05
                r = p.add_run()
                r.text = line
                style_run(r, font=FONT_MAIN, size=22, bold=True, color=WHITE)

        elif txt.startswith("ФИО"):
            tf = shp.text_frame
            tf.clear()
            tf.word_wrap = True
            for i, line in enumerate(AUTHOR_LINES):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.line_spacing = 1.15
                r = p.add_run()
                r.text = line
                style_run(r, font=FONT_MAIN, size=12, bold=True, color=WHITE)

        elif "Казань" in txt:
            # оставляем «Казань – 2026» как есть, но обновим стиль и текст
            tf = shp.text_frame
            tf.clear()
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = "Казань – 2026"
            style_run(r, font=FONT_MAIN, size=12, bold=False, color=WHITE)


# ---------------------------------------------------------------------------
# Слайд 2 — Актуальность (5 KPI-плиток + итог)
# ---------------------------------------------------------------------------

ACTUAL_TILES = [
    ("+2,5%",   "рост строительной отрасли РФ\n(2025 г.; в 2023 — +7,4%)"),
    ("14,5%",   "ключевая ставка ЦБ\n→ ипотека до 24,9%"),
    ("−29%",    "выдачи ипотеки\nв Республике Татарстан, 2025 г."),
    ("149,6%",  "незавершённое строительство\nк годовому вводу жилья"),
    ("30–60",   "минут в день — потери сотрудника\nот ручной работы между системами"),
]


def build_slide_actuality(slide):
    """Слайд 2 — АКТУАЛЬНОСТЬ."""
    clean_content_slide(slide, new_title="АКТУАЛЬНОСТЬ")

    # 5 плиток в одну строку
    n = 5
    gap = Inches(0.10)
    total_w = CONTENT_W - gap * (n - 1)
    tile_w = Emu(int(total_w / n))
    tile_h = Inches(2.75)
    top = Inches(1.00)

    for i, (big, sub) in enumerate(ACTUAL_TILES):
        left = Emu(int(CONTENT_LEFT + (tile_w + gap) * i))
        # внешняя плитка с тонкой рамкой
        card = add_rounded_box(
            slide, left, top, tile_w, tile_h,
            fill=WHITE, line=DARK_BLUE, line_w=1.25, corner=0.08,
        )
        # верхняя «шапка» плитки — тёмно-синяя
        head_h = Inches(0.30)
        head = add_rounded_box(
            slide, left, top, tile_w, head_h,
            fill=DARK_BLUE, line=None, corner=0.10,
        )
        # текст «#N»
        tf = head.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        set_text(
            tf, f"#{i+1}", size=11, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )

        # большая цифра
        big_top = Emu(int(top + head_h + Inches(0.10)))
        big_box = add_text_box(
            slide, left, big_top, tile_w, Inches(1.00),
        )
        tfb = big_box.text_frame
        tfb.margin_left = tfb.margin_right = Inches(0.05)
        tfb.margin_top = tfb.margin_bottom = Inches(0.02)
        set_text(
            tfb, big, size=32, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.0,
        )

        # подпись
        sub_top = Emu(int(big_top + Inches(1.00)))
        sub_box = add_text_box(
            slide, left, sub_top, tile_w, Inches(1.20),
        )
        tfs = sub_box.text_frame
        tfs.margin_left = tfs.margin_right = Inches(0.08)
        tfs.margin_top = tfs.margin_bottom = Inches(0.02)
        tfs.word_wrap = True
        set_text(
            tfs, sub, size=10, bold=False, color=BLACK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.15,
        )

    # Итоговая полоса внизу
    bar_top = Emu(int(top + tile_h + Inches(0.15)))
    bar = add_rounded_box(
        slide, CONTENT_LEFT, bar_top, CONTENT_W, Inches(0.45),
        fill=DARK_BLUE, line=None, corner=0.20,
    )
    tf = bar.text_frame
    tf.margin_left = tf.margin_right = Inches(0.10)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    set_text(
        tf,
        "Девелоперам нужны прозрачные и оптимизированные бизнес-процессы",
        size=14, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        line_spacing=1.0,
    )


# ---------------------------------------------------------------------------
# Слайд 3 — Цель и задачи
# ---------------------------------------------------------------------------

GOAL_TEXT = (
    "Изучение теоретических основ моделирования и оптимизации "
    "бизнес-процессов в девелопменте и их практическое применение "
    "в деятельности АО «Унистрой»"
)
TASKS = [
    "Систематизировать теоретические подходы\nк моделированию бизнес-процессов",
    "Проанализировать строительную отрасль РФ\nи Республики Татарстан",
    "Дать организационно-экономическую\nхарактеристику АО «Унистрой»",
    "Провести диагностику бизнес-процессов\nи выявить «узкие места»",
    "Разработать решения по оптимизации\nи оценить их эффективность",
]


def build_slide_goals(slide):
    """Слайд 3 — ЦЕЛИ И ЗАДАЧИ."""
    clean_content_slide(slide, new_title="ЦЕЛЬ И ЗАДАЧИ")

    # «Цель» — широкая плашка тёмно-синего цвета
    goal_top = Inches(0.95)
    goal_h = Inches(0.95)
    label = add_rounded_box(
        slide, CONTENT_LEFT, goal_top, Inches(1.10), goal_h,
        fill=DARK_BLUE, line=None, corner=0.15,
    )
    tf = label.text_frame
    tf.margin_left = tf.margin_right = Inches(0.05)
    set_text(
        tf, "ЦЕЛЬ", size=14, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    body_left = Emu(int(CONTENT_LEFT + Inches(1.18)))
    body_w = Emu(int(CONTENT_W - Inches(1.18)))
    goal_body = add_rounded_box(
        slide, body_left, goal_top, body_w, goal_h,
        fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.08,
    )
    tf2 = goal_body.text_frame
    tf2.margin_left = Inches(0.15)
    tf2.margin_right = Inches(0.15)
    tf2.margin_top = Inches(0.05)
    tf2.margin_bottom = Inches(0.05)
    tf2.word_wrap = True
    set_text(
        tf2, GOAL_TEXT, size=13, bold=False, color=BLACK,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
        line_spacing=1.2,
    )

    # «Задачи» — 5 плиток в линию
    tasks_top = Inches(2.10)
    n = 5
    gap = Inches(0.10)
    total_w = CONTENT_W - gap * (n - 1)
    tile_w = Emu(int(total_w / n))
    tile_h = Inches(2.65)

    for i, t in enumerate(TASKS):
        left = Emu(int(CONTENT_LEFT + (tile_w + gap) * i))
        card = add_rounded_box(
            slide, left, tasks_top, tile_w, tile_h,
            fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.08,
        )
        # круг с номером (вверху по центру)
        circ_d = Inches(0.62)
        circ_left = Emu(int(left + (tile_w - circ_d) / 2))
        circ_top = Emu(int(tasks_top - Inches(0.30)))
        circ = add_circle(
            slide, circ_left, circ_top, circ_d,
            fill=DARK_BLUE, line=None,
        )
        tfc = circ.text_frame
        tfc.margin_left = tfc.margin_right = Inches(0.0)
        tfc.margin_top = tfc.margin_bottom = Inches(0.0)
        set_text(
            tfc, str(i+1), size=18, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.0,
        )
        # текст задачи
        txt_top = Emu(int(tasks_top + Inches(0.45)))
        txt_box = add_text_box(
            slide, left, txt_top, tile_w, Inches(2.10),
        )
        tft = txt_box.text_frame
        tft.margin_left = tft.margin_right = Inches(0.10)
        tft.margin_top = tft.margin_bottom = Inches(0.05)
        tft.word_wrap = True
        set_text(
            tft, t, size=11, bold=False, color=BLACK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.20,
        )


# ---------------------------------------------------------------------------
# Слайд 4 — Задача 1: теория моделирования
# ---------------------------------------------------------------------------

def build_slide_task1(slide):
    clean_content_slide(slide, new_title="РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 1 · ТЕОРЕТИЧЕСКИЕ ОСНОВЫ")

    # ── Блок 1 (слева сверху): ВХОДЫ → ПРОЦЕСС → ВЫХОДЫ
    block_top = Inches(1.00)
    block_h = Inches(1.20)
    cap_h = Inches(0.30)

    # подпись блока
    cap = add_text_box(slide, CONTENT_LEFT, block_top, Inches(4.20), cap_h)
    set_text(
        cap.text_frame, "БИЗНЕС-ПРОЦЕСС: ВХОДЫ → ПРОЦЕСС → ВЫХОДЫ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )

    # 3 шеврона: ВХОДЫ — ПРОЦЕСС — ВЫХОДЫ
    sch_top = Emu(int(block_top + cap_h + Inches(0.05)))
    sch_h = Inches(0.55)
    labels = [("ВХОДЫ", BLUE2), ("ПРОЦЕСС", DARK_BLUE), ("ВЫХОДЫ", BLUE2)]
    cw = Inches(1.30)
    cgap = Inches(0.05)
    for i, (lab, col) in enumerate(labels):
        left = Emu(int(CONTENT_LEFT + i * (cw + cgap)))
        chev = add_chevron(slide, left, sch_top, cw, sch_h, fill=col)
        tf = chev.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        set_text(
            tf, lab, size=11, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.0,
        )
    # пояснения сверху/снизу
    top_lbl = add_text_box(
        slide, CONTENT_LEFT, Emu(int(sch_top - Inches(0.05))),
        Inches(4.10), Inches(0.30),
    )
    set_text(
        top_lbl.text_frame, "↑ управление", size=9, bold=False, color=GRAY,
        align=PP_ALIGN.CENTER,
    )
    bot_lbl = add_text_box(
        slide, CONTENT_LEFT, Emu(int(sch_top + sch_h + Inches(0.02))),
        Inches(4.10), Inches(0.30),
    )
    set_text(
        bot_lbl.text_frame, "↓ исполнители и ресурсы",
        size=9, bold=False, color=GRAY, align=PP_ALIGN.CENTER,
    )

    # ── Блок 2 (справа сверху): три категории процессов
    right_left = Inches(5.50)
    right_w = Inches(4.38)
    cap2 = add_text_box(slide, right_left, block_top, right_w, cap_h)
    set_text(
        cap2.text_frame, "КЛАССИФИКАЦИЯ ПРОЦЕССОВ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    cats = [
        ("ОСНОВНЫЕ",        "ценность для клиента\n(проектирование, стройка, продажи)"),
        ("ВСПОМОГАТЕЛЬНЫЕ", "поддержка\n(HR, юристы, IT)"),
        ("УПРАВЛЕНЧЕСКИЕ",  "координация\n(бюджет, стратегия)"),
    ]
    cat_top = Emu(int(block_top + cap_h + Inches(0.05)))
    cat_w = Emu(int((right_w - Inches(0.20)) / 3))
    cat_h = Inches(0.95)
    for i, (head, desc) in enumerate(cats):
        cl = Emu(int(right_left + i * (cat_w + Inches(0.10))))
        card = add_rounded_box(
            slide, cl, cat_top, cat_w, cat_h,
            fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.08,
        )
        head_box = add_text_box(slide, cl, cat_top, cat_w, Inches(0.30))
        set_solid_fill(head_box, DARK_BLUE)
        set_no_line(head_box)
        head_box.shadow.inherit = False
        tf = head_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        set_text(
            tf, head, size=10, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.0,
        )
        body = add_text_box(
            slide, cl, Emu(int(cat_top + Inches(0.32))),
            cat_w, Inches(0.62),
        )
        tfb = body.text_frame
        tfb.margin_left = tfb.margin_right = Inches(0.05)
        tfb.margin_top = tfb.margin_bottom = Inches(0.02)
        tfb.word_wrap = True
        set_text(
            tfb, desc, size=9, bold=False, color=BLACK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.10,
        )

    # ── Блок 3 (снизу слева): сравнение нотаций — мини-таблица
    tab_top = Inches(2.95)
    tab_left = CONTENT_LEFT
    tab_w = Inches(5.30)

    cap3 = add_text_box(slide, tab_left, tab_top, tab_w, cap_h)
    set_text(
        cap3.text_frame, "СРАВНЕНИЕ НОТАЦИЙ МОДЕЛИРОВАНИЯ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    # таблица: 4 строки × 3 колонки
    rows = 4
    cols = 3
    t_top = Emu(int(tab_top + cap_h + Inches(0.02)))
    row_h = Inches(0.40)
    col_widths = [Inches(1.30), Inches(1.30), Inches(2.70)]
    headers = ["Нотация", "Уровень", "Сильная сторона"]
    data = [
        ("IDEF0",     "верхний",   "архитектура функций"),
        ("BPMN 2.0",  "детальный", "логика, ветвления, ответственность"),
        ("EPC",       "средний",   "цепочка «событие → функция»"),
    ]
    # шапка
    cur_left = tab_left
    for ci, (htxt, cw_) in enumerate(zip(headers, col_widths)):
        cell = add_text_box(slide, cur_left, t_top, cw_, row_h,
                            fill=DARK_BLUE)
        tf = cell.text_frame
        tf.margin_left = tf.margin_right = Inches(0.06)
        set_text(
            tf, htxt, size=10, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        cur_left = Emu(int(cur_left + cw_))
    # строки
    for ri, row in enumerate(data):
        cur_left = tab_left
        cur_top = Emu(int(t_top + row_h * (ri + 1)))
        bg = WHITE if ri % 2 == 0 else GRAY_BG
        for ci, (val, cw_) in enumerate(zip(row, col_widths)):
            cell = add_text_box(
                slide, cur_left, cur_top, cw_, row_h,
                fill=bg, line=GRAY_BD, line_w=0.5,
            )
            tf = cell.text_frame
            tf.margin_left = tf.margin_right = Inches(0.06)
            tf.word_wrap = True
            set_text(
                tf, val, size=10,
                bold=(ci == 0),
                color=BLACK,
                align=PP_ALIGN.LEFT if ci == 2 else PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE,
            )
            cur_left = Emu(int(cur_left + cw_))

    # ── Блок 4 (снизу справа): инструменты — 4 «логотипа»
    tools_left = Inches(6.55)
    tools_top = tab_top
    tools_w = Inches(3.30)
    cap4 = add_text_box(slide, tools_left, tools_top, tools_w, cap_h)
    set_text(
        cap4.text_frame, "ИНСТРУМЕНТЫ МОДЕЛИРОВАНИЯ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    tools = ["ARIS", "Bizagi", "MS Visio", "Stormbpmn ★"]
    grid_top = Emu(int(tools_top + cap_h + Inches(0.05)))
    cw2 = Emu(int((tools_w - Inches(0.10)) / 2))
    ch2 = Inches(0.55)
    rgap = Inches(0.10)
    for i, tname in enumerate(tools):
        r, c = divmod(i, 2)
        l = Emu(int(tools_left + c * (cw2 + Inches(0.10))))
        t = Emu(int(grid_top + r * (ch2 + rgap)))
        is_star = tname.endswith("★")
        card = add_rounded_box(
            slide, l, t, cw2, ch2,
            fill=DARK_BLUE if is_star else WHITE,
            line=DARK_BLUE, line_w=1.0,
            corner=0.20,
        )
        tf = card.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        set_text(
            tf, tname, size=11, bold=True,
            color=WHITE if is_star else DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
    # подпись под звездой
    note = add_text_box(
        slide, tools_left, Emu(int(grid_top + 2 * (ch2 + rgap))),
        tools_w, Inches(0.30),
    )
    set_text(
        note.text_frame,
        "★ — отечественное решение, реестр Минцифры",
        size=9, bold=False, color=GRAY,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


# ---------------------------------------------------------------------------
# Слайд 5 — Задача 2: отрасль РФ и РТ
# ---------------------------------------------------------------------------

def build_slide_task2(slide):
    clean_content_slide(slide, new_title="РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 2 · АНАЛИЗ ОТРАСЛИ")

    # ── Блок «РФ — 2025»
    section_top = Inches(0.95)

    # Подзаголовок РФ
    rf_label = add_rounded_box(
        slide, CONTENT_LEFT, section_top, Inches(1.50), Inches(0.32),
        fill=DARK_BLUE, line=None, corner=0.20,
    )
    set_text(
        rf_label.text_frame, "РФ · 2025",
        size=12, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )

    rf_kpis = [
        ("18,82 трлн ₽", "объём работ\n(+2,5% г/г)"),
        ("108,1 млн м²", "ввод жилья"),
        ("31–32%",       "распроданность\nстроящихся домов"),
        ("−29%",         "выдачи ипотеки\nк 2024 г."),
    ]
    rf_top = Emu(int(section_top + Inches(0.40)))
    n = 4
    gap = Inches(0.10)
    tile_w = Emu(int((CONTENT_W - gap * (n - 1)) / n))
    tile_h = Inches(1.05)
    for i, (big, sub) in enumerate(rf_kpis):
        left = Emu(int(CONTENT_LEFT + i * (tile_w + gap)))
        card = add_rounded_box(
            slide, left, rf_top, tile_w, tile_h,
            fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.08,
        )
        big_box = add_text_box(slide, left, rf_top, tile_w, Inches(0.50))
        tf = big_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        set_text(
            tf, big, size=16, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM,
        )
        sub_box = add_text_box(
            slide, left, Emu(int(rf_top + Inches(0.50))),
            tile_w, Inches(0.55),
        )
        tfs = sub_box.text_frame
        tfs.margin_left = tfs.margin_right = Inches(0.05)
        tfs.margin_top = tfs.margin_bottom = Inches(0.02)
        tfs.word_wrap = True
        set_text(
            tfs, sub, size=9, bold=False, color=BLACK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.10,
        )

    # ── Блок «Татарстан»
    rt_section_top = Inches(2.20)
    rt_label = add_rounded_box(
        slide, CONTENT_LEFT, rt_section_top, Inches(1.50), Inches(0.32),
        fill=BLUE2, line=None, corner=0.20,
    )
    set_text(
        rt_label.text_frame, "ТАТАРСТАН",
        size=12, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )
    rt_kpis = [
        ("3,515 млн м²", "ввод жилья (+8% — рекорд)"),
        ("265 тыс ₽/м²", "новостройки Казани\n(×1,5 за 3 года)"),
        ("18% → 32%",    "доля федеральных\nигроков в Казани"),
        ("49 место",     "доступность жилья\n(РТ в РФ)"),
    ]
    rt_top = Emu(int(rt_section_top + Inches(0.40)))
    for i, (big, sub) in enumerate(rt_kpis):
        left = Emu(int(CONTENT_LEFT + i * (tile_w + gap)))
        card = add_rounded_box(
            slide, left, rt_top, tile_w, tile_h,
            fill=GRAY_BG, line=BLUE2, line_w=1.0, corner=0.08,
        )
        big_box = add_text_box(slide, left, rt_top, tile_w, Inches(0.50))
        tf = big_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        set_text(
            tf, big, size=15, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM,
        )
        sub_box = add_text_box(
            slide, left, Emu(int(rt_top + Inches(0.50))),
            tile_w, Inches(0.55),
        )
        tfs = sub_box.text_frame
        tfs.margin_left = tfs.margin_right = Inches(0.05)
        tfs.margin_top = tfs.margin_bottom = Inches(0.02)
        tfs.word_wrap = True
        set_text(
            tfs, sub, size=9, bold=False, color=BLACK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.10,
        )

    # ── Блок «Жизненный цикл проекта» — 5 шевронов
    cycle_top = Inches(3.55)
    cap = add_text_box(
        slide, CONTENT_LEFT, cycle_top, CONTENT_W, Inches(0.30),
    )
    set_text(
        cap.text_frame,
        "ЖИЗНЕННЫЙ ЦИКЛ ДЕВЕЛОПЕРСКОГО ПРОЕКТА",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    stages = ["Инициация", "Проектирование", "Стройка", "Продажи", "Ввод"]
    sgap = Inches(0.04)
    sn = len(stages)
    sw = Emu(int((CONTENT_W - sgap * (sn - 1)) / sn))
    sh = Inches(0.50)
    s_top = Emu(int(cycle_top + Inches(0.32)))
    for i, st in enumerate(stages):
        l = Emu(int(CONTENT_LEFT + i * (sw + sgap)))
        col = DARK_BLUE if i % 2 == 0 else BLUE2
        chev = add_chevron(slide, l, s_top, sw, sh, fill=col)
        tf = chev.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        set_text(
            tf, st, size=11, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )

    # ── Нижняя единая мысль
    bot_top = Inches(4.55)
    bar = add_rounded_box(
        slide, CONTENT_LEFT, bot_top, CONTENT_W, Inches(0.42),
        fill=DARK_BLUE, line=None, corner=0.20,
    )
    set_text(
        bar.text_frame,
        "Драйверы: BIM · ИИ · цифровые двойники · BPMS  →  процессная зрелость = конкурентное преимущество",
        size=11, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


# ---------------------------------------------------------------------------
# Слайд 6 — Задача 3: визитка АО «Унистрой»
# ---------------------------------------------------------------------------

def build_slide_task3(slide):
    clean_content_slide(
        slide, new_title="РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 3 · ОРГАНИЗАЦИОННО-ЭКОНОМИЧЕСКАЯ ХАРАКТЕРИСТИКА"
    )

    # ── Шапка-визитка: 4 плитки
    cap = add_text_box(
        slide, CONTENT_LEFT, Inches(0.95), CONTENT_W, Inches(0.30),
    )
    set_text(
        cap.text_frame, "ВИЗИТКА КОМПАНИИ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    visit = [
        ("1996",       "год основания\nбренд «Унистрой» в G-Group"),
        ("8",          "городов присутствия\n(Казань, СПб, Екатеринбург и др.)"),
        ("2,4 млн м²", "построено · 42 тыс. семей"),
        (">400",       "сотрудников\n(271 — управляющая компания)"),
    ]
    v_top = Inches(1.27)
    n = 4
    gap = Inches(0.10)
    tile_w = Emu(int((CONTENT_W - gap * (n - 1)) / n))
    tile_h = Inches(1.05)
    for i, (big, sub) in enumerate(visit):
        l = Emu(int(CONTENT_LEFT + i * (tile_w + gap)))
        card = add_rounded_box(
            slide, l, v_top, tile_w, tile_h,
            fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.08,
        )
        big_box = add_text_box(slide, l, v_top, tile_w, Inches(0.50))
        tf = big_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        set_text(
            tf, big, size=18, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM,
        )
        sub_box = add_text_box(
            slide, l, Emu(int(v_top + Inches(0.50))),
            tile_w, Inches(0.55),
        )
        tfs = sub_box.text_frame
        tfs.margin_left = tfs.margin_right = Inches(0.05)
        tfs.margin_top = tfs.margin_bottom = Inches(0.02)
        tfs.word_wrap = True
        set_text(
            tfs, sub, size=9, bold=False, color=BLACK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
            line_spacing=1.10,
        )

    # ── Финансы 2025: 4 KPI с динамикой
    cap2 = add_text_box(
        slide, CONTENT_LEFT, Inches(2.45), CONTENT_W, Inches(0.30),
    )
    set_text(
        cap2.text_frame, "ФИНАНСОВЫЕ ПОКАЗАТЕЛИ · 2025",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    fin = [
        ("43,9 млрд ₽", "Выручка",         "▲ +26%",        ACCENT_GREEN),
        ("11,1 млрд ₽", "EBITDA",          "рент. 25,2% ▼", ACCENT_RED),
        ("2,4 млрд ₽",  "Чистая прибыль",  "▼ −47%",        ACCENT_RED),
        ("2,79",        "Долг / EBITDA",   "▲ выше нормы",  ACCENT_RED),
    ]
    f_top = Inches(2.77)
    f_h = Inches(1.30)
    for i, (big, lbl, dyn, col) in enumerate(fin):
        l = Emu(int(CONTENT_LEFT + i * (tile_w + gap)))
        card = add_rounded_box(
            slide, l, f_top, tile_w, f_h,
            fill=GRAY_BG, line=DARK_BLUE, line_w=1.0, corner=0.08,
        )
        head = add_text_box(slide, l, f_top, tile_w, Inches(0.30),
                            fill=DARK_BLUE)
        set_text(
            head.text_frame, lbl, size=11, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        big_box = add_text_box(
            slide, l, Emu(int(f_top + Inches(0.32))),
            tile_w, Inches(0.55),
        )
        tf = big_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = Inches(0.02)
        set_text(
            tf, big, size=18, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        dyn_box = add_text_box(
            slide, l, Emu(int(f_top + Inches(0.90))),
            tile_w, Inches(0.40),
        )
        tfd = dyn_box.text_frame
        tfd.margin_left = tfd.margin_right = Inches(0.05)
        tfd.margin_top = tfd.margin_bottom = Inches(0.02)
        tfd.word_wrap = True
        set_text(
            tfd, dyn, size=10, bold=True, color=col,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )

    # ── Позиции на рынке
    cap3 = add_text_box(
        slide, CONTENT_LEFT, Inches(4.20), CONTENT_W, Inches(0.30),
    )
    set_text(
        cap3.text_frame, "ПОЗИЦИИ НА РЫНКЕ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    pos = [
        ("24",     "место в РФ\nпо объёму строительства"),
        ("3",      "место в Казани\n(после «Суварстроит» и ПИК)"),
        ("5,0 / 5","рейтинг ЕРЗ\nза соблюдение сроков"),
    ]
    p_top = Inches(4.52)
    n2 = 3
    p_w = Emu(int((CONTENT_W - gap * (n2 - 1)) / n2))
    p_h = Inches(0.58)
    for i, (big, sub) in enumerate(pos):
        l = Emu(int(CONTENT_LEFT + i * (p_w + gap)))
        card = add_rounded_box(
            slide, l, p_top, p_w, p_h,
            fill=DARK_BLUE, line=None, corner=0.20,
        )
        # большая цифра — слева
        num_w = Inches(0.95)
        num_box = add_text_box(slide, l, p_top, num_w, p_h)
        tf = num_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.05)
        set_text(
            tf, big, size=18, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        # подпись справа
        txt_w = Emu(int(p_w - num_w))
        txt_box = add_text_box(
            slide, Emu(int(l + num_w)), p_top, txt_w, p_h,
        )
        tf2 = txt_box.text_frame
        tf2.margin_left = tf2.margin_right = Inches(0.05)
        tf2.margin_top = tf2.margin_bottom = Inches(0.02)
        tf2.word_wrap = True
        set_text(
            tf2, sub, size=9, bold=False, color=WHITE,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.10,
        )


# ---------------------------------------------------------------------------
# Слайд 7 — Задача 4: 4 «узких места»
# ---------------------------------------------------------------------------

def build_slide_task4(slide):
    clean_content_slide(
        slide, new_title="РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 4 · 4 «УЗКИХ МЕСТА» БИЗНЕС-ПРОЦЕССОВ"
    )

    issues = [
        ("№1 СОГЛАСОВАНИЕ СМЕТ И ПИСЕМ УДОРОЖАНИЯ ВРУЧНУЮ В EXCEL",
         "до 5 дней",
         "одна смета · письма удорожания — до 7 дней"),
        ("№2 ДЕЦЕНТРАЛИЗОВАННАЯ БАЗА ЦЕН (Excel + сетевые папки)",
         "30 мин",
         "поиск одной позиции · экономия аналитиков 3,6%"),
        ("№3 ОТСУТСТВИЕ ГАРАНТИИ ЗАКУПКИ ПРИ ЗАПРОСЕ КП",
         "+10–15%",
         "завышение цен · 25% отказов · 94% нарушений"),
        ("№4 РАЗОРВАННЫЙ ДОКУМЕНТООБОРОТ (СБИС + 1С + бумага)",
         "30–60 мин/день",
         "потери на сотрудника · тысячи человеко-часов в год"),
    ]
    # 2 × 2 сетка карточек
    grid_top = Inches(0.95)
    grid_h = Inches(3.55)
    rows = 2
    cols = 2
    gap = Inches(0.15)
    cw = Emu(int((CONTENT_W - gap) / cols))
    rh = Emu(int((grid_h - gap) / rows))

    for i, (head, big, sub) in enumerate(issues):
        r, c = divmod(i, cols)
        l = Emu(int(CONTENT_LEFT + c * (cw + gap)))
        t = Emu(int(grid_top + r * (rh + gap)))
        card = add_rounded_box(
            slide, l, t, cw, rh,
            fill=WHITE, line=DARK_BLUE, line_w=1.25, corner=0.06,
        )
        # шапка
        head_h = Inches(0.55)
        head_box = add_text_box(slide, l, t, cw, head_h, fill=DARK_BLUE)
        tf = head_box.text_frame
        tf.margin_left = tf.margin_right = Inches(0.10)
        tf.margin_top = tf.margin_bottom = Inches(0.04)
        tf.word_wrap = True
        set_text(
            tf, head, size=11, bold=True, color=WHITE,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.10,
        )
        # большая цифра — слева внизу
        big_top = Emu(int(t + head_h + Inches(0.05)))
        big_w = Emu(int(cw * 0.42))
        big_box = add_text_box(slide, l, big_top, big_w,
                               Emu(int(rh - head_h - Inches(0.10))))
        tfb = big_box.text_frame
        tfb.margin_left = Inches(0.10)
        tfb.margin_right = Inches(0.05)
        tfb.margin_top = tfb.margin_bottom = Inches(0.02)
        set_text(
            tfb, big, size=22, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.05,
        )
        # подпись — справа внизу
        sub_box = add_text_box(
            slide, Emu(int(l + big_w)), big_top,
            Emu(int(cw - big_w)),
            Emu(int(rh - head_h - Inches(0.10))),
        )
        tfs = sub_box.text_frame
        tfs.margin_left = Inches(0.05)
        tfs.margin_right = Inches(0.10)
        tfs.margin_top = tfs.margin_bottom = Inches(0.02)
        tfs.word_wrap = True
        set_text(
            tfs, sub, size=11, bold=False, color=BLACK,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
            line_spacing=1.20,
        )

    # ── «Воронка последствий» внизу
    funnel_top = Inches(4.65)
    bar = add_rounded_box(
        slide, CONTENT_LEFT, funnel_top, CONTENT_W, Inches(0.40),
        fill=DARK_BLUE, line=None, corner=0.20,
    )
    set_text(
        bar.text_frame,
        "Ручной труд  →  задержки  →  срывы поставок  →  штрафы и удорожание",
        size=11, bold=True, color=WHITE,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )


# ---------------------------------------------------------------------------
# Слайд 8 — Задача 5: решения и эффективность
# ---------------------------------------------------------------------------

def build_slide_task5(slide):
    clean_content_slide(
        slide, new_title="РЕЗУЛЬТАТЫ ПО ЗАДАЧЕ 5 · РЕШЕНИЯ И ЭКОНОМИЧЕСКАЯ ЭФФЕКТИВНОСТЬ"
    )

    # Левая часть — 4 решения
    left_x = CONTENT_LEFT
    left_w = Inches(4.85)
    cap = add_text_box(slide, left_x, Inches(0.95), left_w, Inches(0.30))
    set_text(
        cap.text_frame, "ПОРТФЕЛЬ РЕШЕНИЙ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    sols = [
        ("1", "Сквозной модуль согласования смет",
         "1С: ДО + Stormbpmn"),
        ("2", "Централизованная база цен",
         "1С: Долстрой + API поставщиков"),
        ("3", "Регламент гарантированного запроса КП",
         "Маркет Унистрой · 2-стадийный"),
        ("4", "Унификация документооборота",
         "1С: ДО + СБИС + BPM-движок"),
    ]
    s_top = Inches(1.27)
    s_h = Inches(0.78)
    s_gap = Inches(0.08)
    for i, (n, head, sub) in enumerate(sols):
        t = Emu(int(s_top + i * (s_h + s_gap)))
        # круг с номером
        circ_d = Inches(0.65)
        circ = add_circle(
            slide, left_x, Emu(int(t + (s_h - circ_d) / 2)), circ_d,
            fill=DARK_BLUE,
        )
        tfc = circ.text_frame
        set_text(
            tfc, n, size=18, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        # карточка
        card_left = Emu(int(left_x + circ_d + Inches(0.10)))
        card_w = Emu(int(left_w - circ_d - Inches(0.10)))
        card = add_rounded_box(
            slide, card_left, t, card_w, s_h,
            fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.08,
        )
        head_box = add_text_box(
            slide, card_left, t, card_w, Inches(0.40),
        )
        tf = head_box.text_frame
        tf.margin_left = Inches(0.10)
        tf.margin_right = Inches(0.10)
        tf.margin_top = Inches(0.03)
        set_text(
            tf, head, size=12, bold=True, color=DARK_BLUE,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        )
        sub_box = add_text_box(
            slide, card_left, Emu(int(t + Inches(0.40))),
            card_w, Inches(0.40),
        )
        tfs = sub_box.text_frame
        tfs.margin_left = Inches(0.10)
        tfs.margin_right = Inches(0.10)
        tfs.margin_top = Inches(0.0)
        tfs.word_wrap = True
        set_text(
            tfs, sub, size=10, bold=False, color=BLACK,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        )

    # Правая часть — KPI эффективности
    right_x = Inches(6.10)
    right_w = Inches(3.78)
    cap2 = add_text_box(slide, right_x, Inches(0.95), right_w, Inches(0.30))
    set_text(
        cap2.text_frame, "ЭКОНОМИЧЕСКАЯ ЭФФЕКТИВНОСТЬ",
        size=11, bold=True, color=DARK_BLUE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
    )
    kpis = [
        ("6,17 млн ₽",   "совокупные инвестиции"),
        ("2,89 млн ₽/мес", "экономия ФОТ"),
        ("12,9 млн ₽/мес", "совокупная экономия"),
        (">112 млн ₽",   "NPV (5 лет)"),
        ("19,3",         "PI · индекс рентабельности"),
        ("≈ 2,1 мес",    "срок окупаемости портфеля"),
        ("4 835 ч/мес",  "сэкономленное время"),
    ]
    k_top = Inches(1.27)
    k_h = Inches(0.46)
    k_gap = Inches(0.05)
    for i, (big, lbl) in enumerate(kpis):
        t = Emu(int(k_top + i * (k_h + k_gap)))
        # большая цифра — синий блок слева
        num_w = Inches(1.55)
        num = add_rounded_box(
            slide, right_x, t, num_w, k_h,
            fill=DARK_BLUE, line=None, corner=0.10,
        )
        tfn = num.text_frame
        tfn.margin_left = tfn.margin_right = Inches(0.05)
        set_text(
            tfn, big, size=12, bold=True, color=WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        # подпись — серая полоса справа
        lbl_left = Emu(int(right_x + num_w + Inches(0.05)))
        lbl_w = Emu(int(right_w - num_w - Inches(0.05)))
        lbl_box = add_rounded_box(
            slide, lbl_left, t, lbl_w, k_h,
            fill=GRAY_BG, line=GRAY_BD, line_w=0.5, corner=0.10,
        )
        tfl = lbl_box.text_frame
        tfl.margin_left = Inches(0.10)
        tfl.margin_right = Inches(0.10)
        tfl.word_wrap = True
        set_text(
            tfl, lbl, size=10, bold=False, color=BLACK,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
        )


# ---------------------------------------------------------------------------
# Слайд 9 — Выводы
# ---------------------------------------------------------------------------

def build_slide_conclusions(slide):
    clean_content_slide(slide, new_title="ВЫВОДЫ")

    cards = [
        ("ГЛАВА 1 · ТЕОРИЯ",
         "Оптимальное сочетание для девелопмента — BPMN 2.0 + Stormbpmn (импортозамещение, реестр Минцифры)"),
        ("ГЛАВА 2 · ОТРАСЛЬ",
         "РФ в стагнации; Татарстан +8% по вводу, доля федералов в Казани 18% → 32% — конкуренция требует процессной зрелости"),
        ("ГЛАВА 3 · ДИАГНОСТИКА",
         "Выявлено 4 «узких места»: 30–60 мин/день потерь на сотрудника, экономия по сметам всего 3,6%"),
        ("ГЛАВА 3 · РЕШЕНИЯ",
         "Портфель из 4 проектов: инвестиции 6,17 млн ₽, экономия до 12,9 млн ₽/мес, NPV > 112 млн ₽, окупаемость ≈ 2 мес"),
        ("ПРАКТИЧЕСКАЯ ЗНАЧИМОСТЬ",
         "Решения тиражируемы на других застройщиков отрасли"),
    ]

    grid_top = Inches(1.00)
    grid_h = Inches(4.00)
    # 5 карточек: 2 в верхнем ряду, 2 в среднем, 1 широкая внизу
    rows = [
        (cards[0:2], Inches(1.00), Inches(1.20)),
        (cards[2:4], Inches(2.30), Inches(1.20)),
        ([cards[4]], Inches(3.60), Inches(1.20)),
    ]
    gap = Inches(0.15)
    for row_cards, row_top, row_h in rows:
        n = len(row_cards)
        cw = Emu(int((CONTENT_W - gap * (n - 1)) / n))
        for i, (head, body) in enumerate(row_cards):
            l = Emu(int(CONTENT_LEFT + i * (cw + gap)))
            card = add_rounded_box(
                slide, l, row_top, cw, row_h,
                fill=WHITE, line=DARK_BLUE, line_w=1.0, corner=0.06,
            )
            head_box = add_text_box(slide, l, row_top, cw, Inches(0.38),
                                    fill=DARK_BLUE)
            tf = head_box.text_frame
            tf.margin_left = tf.margin_right = Inches(0.10)
            set_text(
                tf, head, size=11, bold=True, color=WHITE,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE,
            )
            body_box = add_text_box(
                slide, l, Emu(int(row_top + Inches(0.40))),
                cw, Emu(int(row_h - Inches(0.40))),
            )
            tfb = body_box.text_frame
            tfb.margin_left = Inches(0.12)
            tfb.margin_right = Inches(0.12)
            tfb.margin_top = Inches(0.05)
            tfb.margin_bottom = Inches(0.05)
            tfb.word_wrap = True
            set_text(
                tfb, body, size=12, bold=False, color=BLACK,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                line_spacing=1.20,
            )


# ---------------------------------------------------------------------------
# Главная сборка
# ---------------------------------------------------------------------------

BUILDERS = {
    2: build_slide_actuality,
    3: build_slide_goals,
    4: build_slide_task1,
    5: build_slide_task2,
    6: build_slide_task3,
    7: build_slide_task4,
    8: build_slide_task5,
    9: build_slide_conclusions,
    # 10 — оставляем как в шаблоне (Спасибо за внимание!)
}


def build():
    if not os.path.exists(TEMPLATE):
        raise SystemExit(f"Шаблон не найден: {TEMPLATE}")

    prs = Presentation(TEMPLATE)
    # Размер уже корректный (10x5.625), но подстрахуемся
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # 1 — титульный
    fill_title_slide(prs.slides[0])

    # 2..9 — контентные
    for n, fn in BUILDERS.items():
        fn(prs.slides[n - 1])

    # 10 — финал — оставляем без изменений (можно подкорректировать шрифт надписи)
    final = prs.slides[9]
    for shp in final.shapes:
        if shp.has_text_frame and shp.text_frame.text.strip().startswith("СПАСИБО"):
            for para in shp.text_frame.paragraphs:
                para.alignment = PP_ALIGN.CENTER
                for run in para.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(40)
                    run.font.bold = True
                    run.font.color.rgb = WHITE

    prs.save(OUTPUT)
    print(f"OK: сохранено в {OUTPUT}")
    print(f"   слайдов: {len(prs.slides)}, размер: "
          f"{Emu(prs.slide_width).inches:.3f} × "
          f"{Emu(prs.slide_height).inches:.3f} in")


if __name__ == "__main__":
    build()
