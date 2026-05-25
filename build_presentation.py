# -*- coding: utf-8 -*-
"""
Генерация PPTX-презентации для защиты ВКР
"АО Унистрой: Моделирование и оптимизация бизнес-процессов"
Шаблон ИУЭФ КФУ 2026, 16:9, тёмно-синий + белый, инфографический стиль.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ----------------------------- Constants ----------------------------------- #
NAVY = RGBColor(0x17, 0x36, 0x5D)
NAVY_LIGHT = RGBColor(0x2E, 0x5A, 0x8C)
ACCENT = RGBColor(0xC0, 0x50, 0x4D)        # тёплый акцент для ▲▼
ACCENT_GREEN = RGBColor(0x3F, 0x88, 0x4A)
GREY = RGBColor(0x59, 0x59, 0x59)
GREY_LIGHT = RGBColor(0xE7, 0xEB, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x10, 0x10, 0x10)

FONT = "Times New Roman"   # как в шаблоне ИУЭФ КФУ

# 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H

BLANK = prs.slide_layouts[6]

# ---------------------------- Helper utils --------------------------------- #
def add_rect(slide, x, y, w, h, fill=NAVY, line=None, line_w=0.0, shadow=False):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w if line_w else 0.75)
    if not shadow:
        # remove default shadow
        sppr = sh.shadow
        sppr.inherit = False
    sh.text_frame.margin_left = Inches(0.1)
    sh.text_frame.margin_right = Inches(0.1)
    sh.text_frame.margin_top = Inches(0.05)
    sh.text_frame.margin_bottom = Inches(0.05)
    sh.text_frame.word_wrap = True
    return sh


def add_round_rect(slide, x, y, w, h, fill=NAVY, line=None, line_w=0.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.adjustments[0] = 0.10
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w if line_w else 0.75)
    sppr = sh.shadow
    sppr.inherit = False
    sh.text_frame.margin_left = Inches(0.12)
    sh.text_frame.margin_right = Inches(0.12)
    sh.text_frame.margin_top = Inches(0.08)
    sh.text_frame.margin_bottom = Inches(0.08)
    sh.text_frame.word_wrap = True
    return sh


def set_text(shape, lines, default_size=14, default_color=BLACK, default_bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=FONT):
    """
    lines: list of dicts or str. Dict keys: text, size, color, bold, italic, align
    """
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if isinstance(lines, str):
        lines = [{"text": lines}]
    for i, ln in enumerate(lines):
        if isinstance(ln, str):
            ln = {"text": ln}
        if i == 0:
            p = tf.paragraphs[0]
            for r in list(p.runs):
                r.text = ""
        else:
            p = tf.add_paragraph()
        p.alignment = ln.get("align", align)
        run = p.add_run()
        run.text = ln.get("text", "")
        run.font.name = font
        run.font.size = Pt(ln.get("size", default_size))
        run.font.bold = ln.get("bold", default_bold)
        run.font.italic = ln.get("italic", False)
        run.font.color.rgb = ln.get("color", default_color)
        if ln.get("space_after"):
            p.space_after = Pt(ln["space_after"])


def add_text_box(slide, x, y, w, h, lines, **kwargs):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.text_frame.word_wrap = True
    set_text(tb, lines, **kwargs)
    return tb


def add_header(slide, title, subtitle=None):
    # Top navy bar
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.9), fill=NAVY)
    # Title
    add_text_box(slide, Inches(0.4), Inches(0.10), SLIDE_W - Inches(0.8), Inches(0.5),
                 [{"text": title, "size": 24, "bold": True, "color": WHITE}],
                 anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text_box(slide, Inches(0.4), Inches(0.50), SLIDE_W - Inches(0.8), Inches(0.4),
                     [{"text": subtitle, "size": 13, "color": WHITE, "italic": True}],
                     anchor=MSO_ANCHOR.MIDDLE)
    # Bottom accent line
    add_rect(slide, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), fill=NAVY)


def add_slide_number(slide, n, total=10):
    add_text_box(slide, SLIDE_W - Inches(1.0), SLIDE_H - Inches(0.16),
                 Inches(0.9), Inches(0.16),
                 [{"text": f"{n} / {total}", "size": 10, "color": WHITE, "bold": True,
                   "align": PP_ALIGN.RIGHT}],
                 anchor=MSO_ANCHOR.MIDDLE)


def kpi_tile(slide, x, y, w, h, icon, big, label, sub=None,
             fill=WHITE, big_color=NAVY, line=NAVY):
    box = add_round_rect(slide, x, y, w, h, fill=fill, line=line, line_w=1.25)
    # icon
    add_text_box(slide, x, y + Inches(0.05), w, Inches(0.45),
                 [{"text": icon, "size": 26, "align": PP_ALIGN.CENTER, "color": NAVY}],
                 anchor=MSO_ANCHOR.MIDDLE)
    # big number
    add_text_box(slide, x, y + Inches(0.50), w, Inches(0.65),
                 [{"text": big, "size": 30, "bold": True,
                   "align": PP_ALIGN.CENTER, "color": big_color}],
                 anchor=MSO_ANCHOR.MIDDLE)
    # label
    add_text_box(slide, x + Inches(0.08), y + Inches(1.15), w - Inches(0.16), h - Inches(1.20),
                 [{"text": label, "size": 11, "align": PP_ALIGN.CENTER, "color": BLACK}],
                 anchor=MSO_ANCHOR.TOP)
    if sub:
        # additional subline at bottom
        pass
    return box


def small_caption(slide, x, y, w, h, text, color=NAVY, size=14, bold=True,
                  align=PP_ALIGN.LEFT):
    return add_text_box(slide, x, y, w, h,
                        [{"text": text, "size": size, "bold": bold,
                          "color": color, "align": align}],
                        anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================== #
#                            СЛАЙД 1 — ТИТУЛЬНЫЙ                              #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)

# Full navy background
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)
# White inset frame
add_rect(s, Inches(0.4), Inches(0.4), SLIDE_W - Inches(0.8), SLIDE_H - Inches(0.8),
         fill=WHITE, line=NAVY, line_w=2.0)

# Top header (institute)
add_text_box(s, Inches(0.6), Inches(0.6), SLIDE_W - Inches(1.2), Inches(1.4),
             [
                 {"text": "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
                  "size": 11, "bold": True, "align": PP_ALIGN.CENTER, "color": NAVY},
                 {"text": "ФГАОУ ВО «КАЗАНСКИЙ (ПРИВОЛЖСКИЙ) ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ»",
                  "size": 11, "bold": True, "align": PP_ALIGN.CENTER, "color": NAVY},
                 {"text": "ИНСТИТУТ УПРАВЛЕНИЯ, ЭКОНОМИКИ И ФИНАНСОВ",
                  "size": 11, "bold": True, "align": PP_ALIGN.CENTER, "color": NAVY},
                 {"text": "Кафедра инноваций и инвестиций",
                  "size": 11, "italic": True, "align": PP_ALIGN.CENTER, "color": NAVY},
                 {"text": "Направление 38.03.02 «Менеджмент» · Профиль «Бизнес-аналитика в управленческой деятельности»",
                  "size": 10, "italic": True, "align": PP_ALIGN.CENTER, "color": GREY},
             ])

# Divider
add_rect(s, Inches(2.0), Inches(2.20), SLIDE_W - Inches(4.0), Inches(0.04), fill=NAVY)

# "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА"
add_text_box(s, Inches(0.6), Inches(2.30), SLIDE_W - Inches(1.2), Inches(0.5),
             [{"text": "ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА",
               "size": 16, "bold": True, "align": PP_ALIGN.CENTER, "color": NAVY}],
             anchor=MSO_ANCHOR.MIDDLE)

# Title (theme)
add_text_box(s, Inches(0.8), Inches(2.95), SLIDE_W - Inches(1.6), Inches(1.4),
             [{"text": "МОДЕЛИРОВАНИЕ И ОПТИМИЗАЦИЯ\nБИЗНЕС-ПРОЦЕССОВ В ДЕВЕЛОПЕРСКОЙ КОМПАНИИ",
               "size": 22, "bold": True, "align": PP_ALIGN.CENTER, "color": NAVY},
              {"text": "(на примере АО «Унистрой»)",
               "size": 18, "italic": True, "align": PP_ALIGN.CENTER, "color": NAVY_LIGHT}],
             anchor=MSO_ANCHOR.MIDDLE)

# Author / supervisor block
add_rect(s, Inches(1.3), Inches(4.7), SLIDE_W - Inches(2.6), Inches(0.04), fill=NAVY)
add_text_box(s, Inches(1.5), Inches(4.85), SLIDE_W - Inches(3.0), Inches(1.5),
             [
                 {"text": "Выполнил(а): _______________________________________",
                  "size": 13, "align": PP_ALIGN.LEFT, "color": BLACK},
                 {"text": "Группа: _______________",
                  "size": 13, "align": PP_ALIGN.LEFT, "color": BLACK, "space_after": 6},
                 {"text": "Научный руководитель: ___________________________",
                  "size": 13, "align": PP_ALIGN.LEFT, "color": BLACK},
             ])

# Footer — city/year
add_text_box(s, Inches(0.6), SLIDE_H - Inches(1.0), SLIDE_W - Inches(1.2), Inches(0.4),
             [{"text": "Казань — 2026", "size": 14, "bold": True,
               "align": PP_ALIGN.CENTER, "color": NAVY}],
             anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================== #
#                          СЛАЙД 2 — АКТУАЛЬНОСТЬ                             #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "АКТУАЛЬНОСТЬ ИССЛЕДОВАНИЯ",
           "Девелопменту нужны прозрачные и оптимизированные бизнес-процессы")
add_slide_number(s, 2)

# 5 KPI tiles — equal width grid
gx = Inches(0.4); gy = Inches(1.15)
gap = Inches(0.15)
total_w = SLIDE_W - Inches(0.8)
tile_w = (total_w - gap * 4) / 5
tile_h = Inches(2.7)

tiles = [
    ("\u25B2",  "+2,5%",
     "рост строительной отрасли РФ\nв 2025 г. — минимум за 5 лет\n(в 2023 г. было +7,4%)"),
    ("%",       "14,5%",
     "ключевая ставка ЦБ (май 2026)\n→ ипотека до 24,9% годовых"),
    ("\u25BC",  "−29%",
     "объём выданной ипотеки\nв Республике Татарстан\nза 2025 год"),
    ("\u29D6",  "149,6%",
     "соотношение незавершённого\nстроительства к годовому вводу\n(было 106,8%)"),
    ("\u29D7",  "30–60",
     "минут в день теряет сотрудник\n«Унистроя» на ручной перенос\nданных между 3 системами"),
]

for i, (icon, big, label) in enumerate(tiles):
    x = gx + i * (tile_w + gap)
    kpi_tile(s, x, gy, tile_w, tile_h, icon, big, label,
             fill=WHITE, big_color=NAVY, line=NAVY)

# Bottom highlight bar
add_rect(s, Inches(0.4), Inches(4.10), SLIDE_W - Inches(0.8), Inches(0.04), fill=NAVY)

# Macro insight blocks below
mb_y = Inches(4.35)
mb_h = Inches(2.7)
mb_w = (SLIDE_W - Inches(0.8) - Inches(0.3)) / 2
# left — Federal context
left = add_round_rect(s, Inches(0.4), mb_y, mb_w, mb_h, fill=GREY_LIGHT, line=NAVY, line_w=1.0)
set_text(left, [
    {"text": "ФЕДЕРАЛЬНЫЙ КОНТЕКСТ", "size": 14, "bold": True, "color": NAVY,
     "align": PP_ALIGN.LEFT, "space_after": 6},
    {"text": "• Замедление прироста отрасли · падение ипотеки · накопление нераспроданных квартир",
     "size": 12, "color": BLACK, "space_after": 4},
    {"text": "• Распроданность строящегося жилья:  31–32%  (было 50–55%)",
     "size": 12, "color": BLACK, "space_after": 4},
    {"text": "• Цена м² Казань (новостройка): 265 тыс. ₽  (×1,5 за 3 года)",
     "size": 12, "color": BLACK},
])

# right — Operational pain
right = add_round_rect(s, Inches(0.4) + mb_w + Inches(0.3), mb_y, mb_w, mb_h,
                       fill=NAVY, line=NAVY, line_w=1.0)
set_text(right, [
    {"text": "ВНУТРЕННИЕ ПРОЦЕССЫ", "size": 14, "bold": True, "color": WHITE,
     "align": PP_ALIGN.LEFT, "space_after": 6},
    {"text": "• 3 параллельных системы документооборота: СБИС + 1С: ДО + бумага",
     "size": 12, "color": WHITE, "space_after": 4},
    {"text": "• Согласование сметы вручную в Excel — до 5 дней",
     "size": 12, "color": WHITE, "space_after": 4},
    {"text": "• Письма удорожания зависают до 7 дней — срывы поставок и штрафы",
     "size": 12, "color": WHITE},
])


# =========================================================================== #
#                       СЛАЙД 3 — ЦЕЛЬ И ЗАДАЧИ                               #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ЦЕЛЬ И ЗАДАЧИ ИССЛЕДОВАНИЯ")
add_slide_number(s, 3)

# Goal — large navy plate
goal = add_round_rect(s, Inches(0.4), Inches(1.15), SLIDE_W - Inches(0.8), Inches(1.55),
                      fill=NAVY, line=NAVY)
set_text(goal, [
    {"text": "ЦЕЛЬ", "size": 16, "bold": True, "color": WHITE,
     "align": PP_ALIGN.LEFT, "space_after": 4},
    {"text": "Изучение теоретических основ моделирования и оптимизации\n"
             "бизнес-процессов в девелопменте и их практическое применение\n"
             "в деятельности АО «Унистрой».",
     "size": 18, "bold": True, "color": WHITE, "align": PP_ALIGN.LEFT},
])

# Tasks header
small_caption(s, Inches(0.4), Inches(2.85), SLIDE_W - Inches(0.8), Inches(0.4),
              "ЗАДАЧИ", color=NAVY, size=16)

# 5 tasks tiles in a row
tasks = [
    ("1", "Систематизировать\nтеоретические подходы\nк моделированию\nбизнес-процессов"),
    ("2", "Проанализировать\nстроительную отрасль\nРФ и Татарстана"),
    ("3", "Дать организационно-\nэкономическую\nхарактеристику\nАО «Унистрой»"),
    ("4", "Провести диагностику\nбизнес-процессов\nкомпании, выявить\n«узкие места»"),
    ("5", "Разработать решения\nи оценить их\nэкономическую\nэффективность"),
]
ty = Inches(3.35)
th = Inches(3.55)
for i, (num, txt) in enumerate(tasks):
    x = gx + i * (tile_w + gap)
    box = add_round_rect(s, x, ty, tile_w, th, fill=WHITE, line=NAVY, line_w=1.5)
    # number circle
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL,
                              x + tile_w / 2 - Inches(0.4),
                              ty + Inches(0.25),
                              Inches(0.8), Inches(0.8))
    circ.fill.solid(); circ.fill.fore_color.rgb = NAVY
    circ.line.fill.background()
    set_text(circ, [{"text": num, "size": 30, "bold": True, "color": WHITE,
                     "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, x + Inches(0.1), ty + Inches(1.2), tile_w - Inches(0.2),
                 th - Inches(1.3),
                 [{"text": txt, "size": 13, "color": BLACK,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.TOP)


# =========================================================================== #
#               СЛАЙД 4 — ГЛАВА 1 (ТЕОРИЯ)                                    #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ГЛАВА 1. ТЕОРЕТИЧЕСКИЕ ОСНОВЫ МОДЕЛИРОВАНИЯ БИЗНЕС-ПРОЦЕССОВ")
add_slide_number(s, 4)

# Block A — process schema
small_caption(s, Inches(0.4), Inches(1.05), Inches(6.2), Inches(0.4),
              "Бизнес-процесс = преобразование ресурсов в ценность", color=NAVY)

# 3 chevron-like boxes (input → process → output)
sx = Inches(0.4); sy = Inches(1.5); sw = Inches(1.8); sh_ = Inches(1.0)
gap_s = Inches(0.10)
labels = [("ВХОДЫ", "ресурсы"), ("ПРОЦЕСС", "преобразование"), ("ВЫХОДЫ", "ценность для клиента")]
for i, (t1, t2) in enumerate(labels):
    box = add_round_rect(s, sx + i * (sw + gap_s), sy, sw, sh_,
                         fill=NAVY if i == 1 else WHITE, line=NAVY, line_w=1.5)
    set_text(box, [
        {"text": t1, "size": 13, "bold": True,
         "color": WHITE if i == 1 else NAVY, "align": PP_ALIGN.CENTER},
        {"text": t2, "size": 11,
         "color": WHITE if i == 1 else GREY, "align": PP_ALIGN.CENTER, "italic": True},
    ], anchor=MSO_ANCHOR.MIDDLE)

# arrows between
for i in range(2):
    ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                            sx + (i + 1) * sw + i * gap_s + Inches(0.005),
                            sy + Inches(0.40),
                            Inches(0.10), Inches(0.20))
    ar.fill.solid(); ar.fill.fore_color.rgb = NAVY
    ar.line.fill.background()

# Block B — classification (3 colored tiles)
small_caption(s, Inches(0.4), Inches(2.75), Inches(6.2), Inches(0.4),
              "Классификация процессов", color=NAVY)
cls = [
    ("ОСНОВНЫЕ", "ценность для клиента\nпроектирование · стройка · продажи", NAVY),
    ("ВСПОМОГАТЕЛЬНЫЕ", "поддержка\nHR · юристы · IT", NAVY_LIGHT),
    ("УПРАВЛЕНЧЕСКИЕ", "координация\nбюджет · стратегия", GREY),
]
cy = Inches(3.20); ch = Inches(1.30); cw = Inches(1.95)
for i, (h_, txt, color) in enumerate(cls):
    box = add_round_rect(s, Inches(0.4) + i * (cw + Inches(0.10)), cy, cw, ch,
                         fill=color, line=color)
    set_text(box, [
        {"text": h_, "size": 13, "bold": True, "color": WHITE,
         "align": PP_ALIGN.CENTER, "space_after": 4},
        {"text": txt, "size": 11, "color": WHITE, "align": PP_ALIGN.CENTER},
    ], anchor=MSO_ANCHOR.MIDDLE)

# Block C — notations comparison (right side)
small_caption(s, Inches(7.0), Inches(1.05), Inches(5.9), Inches(0.4),
              "Нотации моделирования", color=NAVY)
not_data = [
    ("IDEF0", "верхний уровень", "архитектура функций"),
    ("BPMN 2.0", "детальный", "логика, ветвления, ответственность"),
    ("EPC", "средний", "событие → функция"),
]
nx = Inches(7.0); ny = Inches(1.50); nh = Inches(0.55); nw = Inches(5.9)
# header row
hdr = add_rect(s, nx, ny, nw, nh, fill=NAVY)
set_text(hdr, [{"text": "Нотация       Уровень       Сильная сторона",
                "size": 12, "bold": True, "color": WHITE, "align": PP_ALIGN.LEFT}],
         anchor=MSO_ANCHOR.MIDDLE)
# rows
for i, row in enumerate(not_data):
    y = ny + nh + i * Inches(0.50)
    bg = WHITE if i % 2 == 0 else GREY_LIGHT
    add_rect(s, nx, y, nw, Inches(0.50), fill=bg, line=NAVY, line_w=0.5)
    cells = [(row[0], 1.5, True, NAVY),
             (row[1], 1.7, False, BLACK),
             (row[2], nw.inches - 1.5 - 1.7 - 0.2, False, BLACK)]
    cx = nx + Inches(0.1)
    for txt, w_in, bold, col in cells:
        tb = s.shapes.add_textbox(cx, y, Inches(w_in), Inches(0.50))
        set_text(tb, [{"text": txt, "size": 12, "bold": bold,
                       "color": col, "align": PP_ALIGN.LEFT}],
                 anchor=MSO_ANCHOR.MIDDLE)
        cx += Inches(w_in)

# Block D — platforms
small_caption(s, Inches(7.0), Inches(4.10), Inches(5.9), Inches(0.4),
              "Платформы", color=NAVY)
plats = [
    ("ARIS", GREY),
    ("Bizagi", GREY),
    ("MS Visio", GREY),
    ("Stormbpmn\n★ отеч., Минцифры", NAVY),
]
px = Inches(7.0); py = Inches(4.55); pw = Inches(1.40); ph = Inches(1.05)
for i, (name, color) in enumerate(plats):
    x = px + i * (pw + Inches(0.10))
    box = add_round_rect(s, x, py, pw, ph, fill=color, line=color)
    set_text(box, [{"text": name, "size": 12, "bold": True, "color": WHITE,
                    "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)

# Bottom takeaway bar
add_rect(s, Inches(0.4), Inches(6.55), SLIDE_W - Inches(0.8), Inches(0.55), fill=NAVY)
add_text_box(s, Inches(0.5), Inches(6.55), SLIDE_W - Inches(1.0), Inches(0.55),
             [{"text": "ВЫВОД ПО ГЛАВЕ 1: для девелопмента оптимально BPMN 2.0 + Stormbpmn (импортозамещение)",
               "size": 13, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================== #
#               СЛАЙД 5 — ГЛАВА 2 (ОТРАСЛЬ)                                   #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ГЛАВА 2. АНАЛИЗ ОТРАСЛИ И ТИПОВЫХ БИЗНЕС-ПРОЦЕССОВ")
add_slide_number(s, 5)

# Left — Russia 2025
small_caption(s, Inches(0.4), Inches(1.05), Inches(6.2), Inches(0.4),
              "Россия — 2025", color=NAVY)
rf = [
    ("\u20BD", "18,82",  "трлн ₽\nобъём работ (+2,5%)"),
    ("\u2302", "108,1",  "млн м²\nввод жилья"),
    ("\u29D7", "31–32%", "распроданность\nстроящихся домов"),
    ("\u25BC", "−29%",   "ипотека к 2024 г."),
]
rx = Inches(0.4); ry = Inches(1.50); rh = Inches(1.85)
rw = (Inches(6.2) - Inches(0.30)) / 4
for i, (ic, big, lab) in enumerate(rf):
    kpi_tile(s, rx + i * (rw + Inches(0.10)), ry, rw, rh, ic, big, lab,
             fill=WHITE, line=NAVY)

# Right — Tatarstan
small_caption(s, Inches(7.0), Inches(1.05), Inches(5.9), Inches(0.4),
              "Татарстан / Казань", color=NAVY)
tt = [
    ("\u2191", "3,515",  "млн м²\nввод жилья РТ (+8% — рекорд)"),
    ("\u20BD", "265",    "тыс. ₽ за м²\nКазань — ×1,5 за 3 года"),
    ("\u25B2", "32%",    "доля федералов\nв Казани (было 18%)"),
    ("\u2691", "49",     "место РТ в РФ\nпо доступности жилья"),
]
tx = Inches(7.0); ty2 = Inches(1.50); tw_ = (Inches(5.9) - Inches(0.30)) / 4
for i, (ic, big, lab) in enumerate(tt):
    kpi_tile(s, tx + i * (tw_ + Inches(0.10)), ty2, tw_, rh, ic, big, lab,
             fill=NAVY, line=NAVY, big_color=WHITE)
    # repaint label color to white & icon to white for navy tiles — quick override
    # since kpi_tile assumes white tiles, we add overlay text for contrast
    # (workaround: redraw labels in white)
    # icon
    add_rect(s, tx + i * (tw_ + Inches(0.10)) + Inches(0.02),
             ty2 + Inches(0.02), tw_ - Inches(0.04), rh - Inches(0.04),
             fill=NAVY, line=NAVY)
    add_text_box(s, tx + i * (tw_ + Inches(0.10)),
                 ty2 + Inches(0.05), tw_, Inches(0.45),
                 [{"text": ic, "size": 26, "color": WHITE,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, tx + i * (tw_ + Inches(0.10)),
                 ty2 + Inches(0.50), tw_, Inches(0.65),
                 [{"text": big, "size": 30, "bold": True, "color": WHITE,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, tx + i * (tw_ + Inches(0.10)) + Inches(0.08),
                 ty2 + Inches(1.15), tw_ - Inches(0.16), rh - Inches(1.20),
                 [{"text": lab, "size": 11, "color": WHITE,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.TOP)

# Lifecycle row (5 chevrons)
small_caption(s, Inches(0.4), Inches(3.55), SLIDE_W - Inches(0.8), Inches(0.4),
              "Жизненный цикл девелоперского проекта", color=NAVY)
phases = ["Инициация", "Проектирование", "Строительство", "Продажи", "Ввод в эксплуатацию"]
ph_y = Inches(4.0); ph_h = Inches(0.7)
ph_total = SLIDE_W - Inches(0.8)
ph_w = (ph_total - Inches(0.10) * (len(phases) - 1)) / len(phases)
for i, ph in enumerate(phases):
    sh = s.shapes.add_shape(MSO_SHAPE.PENTAGON,
                            Inches(0.4) + i * (ph_w + Inches(0.10)),
                            ph_y, ph_w, ph_h)
    sh.fill.solid(); sh.fill.fore_color.rgb = NAVY if i % 2 == 0 else NAVY_LIGHT
    sh.line.fill.background()
    set_text(sh, [{"text": ph, "size": 13, "bold": True,
                   "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)

# Methods row (5 tiles)
small_caption(s, Inches(0.4), Inches(4.95), SLIDE_W - Inches(0.8), Inches(0.4),
              "Методы оптимизации", color=NAVY)
methods = [
    ("Стандарти-\nзация", NAVY),
    ("Параллельные\nпроцессы", NAVY),
    ("Электронный\nдокументооборот", NAVY),
    ("Цифровой\nконтроль\nподрядчиков (BIM)", NAVY),
    ("Управление\nрассрочками", NAVY),
]
m_y = Inches(5.40); m_h = Inches(1.10)
m_w = (ph_total - Inches(0.10) * (len(methods) - 1)) / len(methods)
for i, (txt, col) in enumerate(methods):
    box = add_round_rect(s, Inches(0.4) + i * (m_w + Inches(0.10)),
                         m_y, m_w, m_h, fill=WHITE, line=NAVY, line_w=1.5)
    set_text(box, [{"text": txt, "size": 12, "bold": True, "color": NAVY,
                    "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)

# Bottom takeaway
add_rect(s, Inches(0.4), Inches(6.65), SLIDE_W - Inches(0.8), Inches(0.45), fill=NAVY)
add_text_box(s, Inches(0.5), Inches(6.65), SLIDE_W - Inches(1.0), Inches(0.45),
             [{"text": "ВЫВОД ПО ГЛАВЕ 2: РФ в стагнации, РТ +8% — но 32% Казани заняли федералы → процессная зрелость = условие выживания",
               "size": 12, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================== #
#                СЛАЙД 6 — АО «УНИСТРОЙ» (ВИЗИТКА)                            #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ГЛАВА 3. АО «УНИСТРОЙ» — ОРГАНИЗАЦИОННО-ЭКОНОМИЧЕСКАЯ ХАРАКТЕРИСТИКА")
add_slide_number(s, 6)

# Top row — visit card 4 tiles
visit = [
    ("\u2691", "1996", "год основания\nбренд в составе G-Group"),
    ("\u2316", "8",    "городов присутствия\nКазань · СПб · Екатеринбург ·\nПермь · Уфа · Тольятти ·\nМахачкала · Н.Новгород"),
    ("\u29D7", "2,4",  "млн м²\nпостроено / 42 тыс. семей"),
    ("\u2630", ">400", "сотрудников\n(271 — управляющая\nкомпания)"),
]
vy = Inches(1.10); vh = Inches(2.30)
vw = (SLIDE_W - Inches(0.8) - Inches(0.30)) / 4
for i, (ic, big, lab) in enumerate(visit):
    kpi_tile(s, Inches(0.4) + i * (vw + Inches(0.10)), vy, vw, vh, ic, big, lab,
             fill=NAVY, line=NAVY, big_color=WHITE)
    # override colors (white)
    box_x = Inches(0.4) + i * (vw + Inches(0.10))
    add_rect(s, box_x + Inches(0.02), vy + Inches(0.02),
             vw - Inches(0.04), vh - Inches(0.04), fill=NAVY, line=NAVY)
    add_text_box(s, box_x, vy + Inches(0.05), vw, Inches(0.45),
                 [{"text": ic, "size": 26, "color": WHITE, "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, box_x, vy + Inches(0.50), vw, Inches(0.65),
                 [{"text": big, "size": 36, "bold": True, "color": WHITE,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, box_x + Inches(0.08), vy + Inches(1.20), vw - Inches(0.16),
                 vh - Inches(1.25),
                 [{"text": lab, "size": 11, "color": WHITE, "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.TOP)

# Mid section — Финансы 2025
small_caption(s, Inches(0.4), Inches(3.55), Inches(6.2), Inches(0.4),
              "Финансовые показатели 2025", color=NAVY)
fin = [
    ("\u20BD", "43,9",  "млрд ₽ выручка", "▲ +26%", ACCENT_GREEN),
    ("E",      "11,1",  "млрд ₽ EBITDA\n(рентаб. 25,2%)", "▼ с 31,2%", ACCENT),
    ("\u20BD", "2,4",   "млрд ₽ чистая прибыль", "▼ −47%", ACCENT),
    ("D",      "2,79",  "Долг / EBITDA", "▲ рост", ACCENT),
]
fy = Inches(4.00); fh = Inches(2.20)
fw = (Inches(6.2) - Inches(0.30)) / 4
for i, (ic, big, lab, delta, dcolor) in enumerate(fin):
    x = Inches(0.4) + i * (fw + Inches(0.10))
    kpi_tile(s, x, fy, fw, fh, ic, big, lab, fill=WHITE, line=NAVY)
    # delta tag
    tag = add_round_rect(s, x + fw / 2 - Inches(0.55), fy + fh - Inches(0.45),
                         Inches(1.10), Inches(0.35), fill=dcolor, line=dcolor)
    set_text(tag, [{"text": delta, "size": 11, "bold": True, "color": WHITE,
                    "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)

# Right — positions
small_caption(s, Inches(7.0), Inches(3.55), Inches(5.9), Inches(0.4),
              "Позиции компании", color=NAVY)
pos = [
    ("24", "место в РФ\nпо объёму строительства"),
    ("3",  "место в Казани\n(после «Суварстроит» и ПИК)"),
    ("5",  "баллов рейтинга ЕРЗ\n(соблюдение сроков)"),
]
py3 = Inches(4.00); ph3 = Inches(2.20)
pw3 = (Inches(5.9) - Inches(0.20)) / 3
for i, (big, lab) in enumerate(pos):
    x = Inches(7.0) + i * (pw3 + Inches(0.10))
    box = add_round_rect(s, x, py3, pw3, ph3, fill=GREY_LIGHT, line=NAVY, line_w=1.5)
    add_text_box(s, x, py3 + Inches(0.30), pw3, Inches(1.0),
                 [{"text": big, "size": 56, "bold": True, "color": NAVY,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, x + Inches(0.1), py3 + Inches(1.30), pw3 - Inches(0.2),
                 ph3 - Inches(1.40),
                 [{"text": lab, "size": 12, "color": BLACK,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.TOP)

# Bottom strip
add_rect(s, Inches(0.4), Inches(6.55), SLIDE_W - Inches(0.8), Inches(0.55), fill=NAVY)
add_text_box(s, Inches(0.5), Inches(6.55), SLIDE_W - Inches(1.0), Inches(0.55),
             [{"text": "Федеральный девелопер · 30 лет на рынке · ТОП-3 в Казани, ТОП-25 в России",
               "size": 13, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================== #
#                  СЛАЙД 7 — ПРОБЛЕМЫ В БИЗНЕС-ПРОЦЕССАХ                      #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ГЛАВА 3. 4 «УЗКИХ МЕСТА» БИЗНЕС-ПРОЦЕССОВ АО «УНИСТРОЙ»")
add_slide_number(s, 7)

problems = [
    ("01", "Ручное согласование смет\nи писем удорожания",
     "Excel + почта, без BPM-движка. Передача через нескольких аналитиков.",
     "до 5 дней", "одна смета", "до 7 дней — письма удорожания"),
    ("02", "Децентрализованная база\nцен на материалы",
     "Excel + сетевые папки, 1С: Долстрой не используется.",
     "30 мин", "поиск 1 позиции", "экономия аналитиков всего 3,6%"),
    ("03", "Отсутствие гарантии\nзакупки при запросе КП",
     "Подрядчики, не уверенные в договоре, завышают цены или отказываются.",
     "+10–15%", "завышение цен", "94% неустранённых нарушений"),
    ("04", "Разорванный\nдокументооборот",
     "3 параллельных канала: СБИС + 1С: ДО + бумага.",
     "30–60 мин/день", "потери на сотрудника", "тысячи человеко-часов в год"),
]

py4 = Inches(1.10); ph4 = Inches(2.55)
pw4 = (SLIDE_W - Inches(0.8) - Inches(0.30)) / 4
for i, (num, title, desc, big, big_lab, sub) in enumerate(problems):
    x = Inches(0.4) + i * (pw4 + Inches(0.10))
    # card
    card = add_round_rect(s, x, py4, pw4, ph4, fill=WHITE, line=NAVY, line_w=1.5)
    # number badge
    badge = add_rect(s, x, py4, Inches(0.85), Inches(0.55), fill=NAVY, line=NAVY)
    set_text(badge, [{"text": num, "size": 18, "bold": True, "color": WHITE,
                      "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
    # title
    add_text_box(s, x + Inches(0.95), py4 + Inches(0.05), pw4 - Inches(1.0), Inches(0.55),
                 [{"text": title, "size": 12, "bold": True, "color": NAVY,
                   "align": PP_ALIGN.LEFT}],
                 anchor=MSO_ANCHOR.MIDDLE)
    # description
    add_text_box(s, x + Inches(0.10), py4 + Inches(0.65), pw4 - Inches(0.2), Inches(0.65),
                 [{"text": desc, "size": 10, "italic": True, "color": GREY,
                   "align": PP_ALIGN.LEFT}],
                 anchor=MSO_ANCHOR.TOP)
    # big number
    add_text_box(s, x, py4 + Inches(1.30), pw4, Inches(0.55),
                 [{"text": big, "size": 24, "bold": True, "color": ACCENT,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, x + Inches(0.10), py4 + Inches(1.85), pw4 - Inches(0.2), Inches(0.30),
                 [{"text": big_lab, "size": 11, "color": BLACK,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, x + Inches(0.10), py4 + Inches(2.15), pw4 - Inches(0.2), Inches(0.35),
                 [{"text": sub, "size": 10, "italic": True, "color": NAVY,
                   "align": PP_ALIGN.CENTER}],
                 anchor=MSO_ANCHOR.MIDDLE)

# Funnel arrow on bottom
small_caption(s, Inches(0.4), Inches(3.85), SLIDE_W - Inches(0.8), Inches(0.4),
              "Кумулятивный эффект потерь", color=NAVY)
funnel = [
    ("Ручной труд", NAVY),
    ("Задержки", NAVY_LIGHT),
    ("Срывы поставок", ACCENT),
    ("Штрафы и удорожание", ACCENT),
]
fy2 = Inches(4.30); fh2 = Inches(0.85)
fw2_total = SLIDE_W - Inches(0.8)
fw2 = (fw2_total - Inches(0.10) * (len(funnel) - 1)) / len(funnel)
for i, (txt, color) in enumerate(funnel):
    sh = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                            Inches(0.4) + i * (fw2 + Inches(0.10)),
                            fy2, fw2, fh2)
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    set_text(sh, [{"text": txt, "size": 14, "bold": True, "color": WHITE,
                   "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)

# Pull-quote section
add_rect(s, Inches(0.4), Inches(5.55), SLIDE_W - Inches(0.8), Inches(1.55),
         fill=GREY_LIGHT, line=NAVY, line_w=1.5)
add_text_box(s, Inches(0.6), Inches(5.65), SLIDE_W - Inches(1.2), Inches(1.4),
             [
                 {"text": "ОБЩИЙ ДИАГНОЗ", "size": 14, "bold": True,
                  "color": NAVY, "align": PP_ALIGN.LEFT, "space_after": 6},
                 {"text": "Бизнес-процессы фрагментированы, ключевые операции выполняются вручную, "
                          "ИТ-системы не интегрированы — это снижает экономическую эффективность "
                          "и создаёт риск срыва сроков сдачи объектов.",
                  "size": 13, "color": BLACK, "align": PP_ALIGN.LEFT}
             ])


# =========================================================================== #
#         СЛАЙД 8 — РЕШЕНИЯ И ЭКОНОМИЧЕСКАЯ ЭФФЕКТИВНОСТЬ                     #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ГЛАВА 3. РЕШЕНИЯ ПО ОПТИМИЗАЦИИ И ЭКОНОМИЧЕСКАЯ ЭФФЕКТИВНОСТЬ")
add_slide_number(s, 8)

# Left half — 4 solutions
small_caption(s, Inches(0.4), Inches(1.05), Inches(6.7), Inches(0.4),
              "Портфель из 4 взаимоувязанных решений", color=NAVY)
sols = [
    ("01", "\u21BB", "Сквозной модуль согласования смет",
     "1С: ДО + Stormbpmn — авто-распределение позиций между аналитиками, сметным отделом и закупками"),
    ("02", "\u29C9", "Централизованная база цен в 1С: Долстрой",
     "API поставщиков (Лемана, Металлоторг, ЭТМ) + победителей «Маркет Унистрой»"),
    ("03", "\u2691", "Регламент гарантированного запроса КП",
     "Двухстадийный: подтверждение закупки до выставления счёта; рейтинг подрядчиков"),
    ("04", "\u29C8", "Унификация документооборота",
     "Всё в 1С: ДО + BPM-движок; синхронизация с СБИС через API; бумага — только госорганам"),
]
sx2 = Inches(0.4); sy2 = Inches(1.50); sh2 = Inches(1.18); sw2 = Inches(6.7)
for i, (num, ic, title, desc) in enumerate(sols):
    y = sy2 + i * (sh2 + Inches(0.10))
    card = add_round_rect(s, sx2, y, sw2, sh2, fill=WHITE, line=NAVY, line_w=1.5)
    # number circle
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL, sx2 + Inches(0.10), y + Inches(0.20),
                              Inches(0.78), Inches(0.78))
    circ.fill.solid(); circ.fill.fore_color.rgb = NAVY
    circ.line.fill.background()
    set_text(circ, [{"text": num, "size": 18, "bold": True, "color": WHITE,
                     "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
    # title + desc
    add_text_box(s, sx2 + Inches(1.00), y + Inches(0.10), sw2 - Inches(1.10), Inches(0.45),
                 [{"text": f"{ic}  {title}", "size": 14, "bold": True, "color": NAVY,
                   "align": PP_ALIGN.LEFT}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, sx2 + Inches(1.00), y + Inches(0.55), sw2 - Inches(1.10), Inches(0.55),
                 [{"text": desc, "size": 11, "color": BLACK, "align": PP_ALIGN.LEFT}],
                 anchor=MSO_ANCHOR.TOP)

# Right half — economic effect
small_caption(s, Inches(7.30), Inches(1.05), Inches(5.6), Inches(0.4),
              "Экономическая эффективность портфеля", color=NAVY)

eff = [
    ("Инвестиции",            "6,17",   "млн ₽",     NAVY),
    ("Экономия ФОТ/мес",     "2,89",   "млн ₽",      NAVY),
    ("Полная экономия/мес", "до 12,9", "млн ₽",      NAVY),
    ("NPV (5 лет)",           ">112",   "млн ₽",      NAVY),
    ("PI",                    "19,3",   "× 1",        NAVY),
    ("Срок окупаемости",     "~2,1",   "мес.",       NAVY),
]
ex = Inches(7.30); ey = Inches(1.50); eh = Inches(0.95)
ew = (Inches(5.6) - Inches(0.10)) / 2
for i, (lab, big, unit, col) in enumerate(eff):
    r = i // 2; c = i % 2
    x = ex + c * (ew + Inches(0.10))
    y = ey + r * (eh + Inches(0.08))
    box = add_round_rect(s, x, y, ew, eh, fill=WHITE, line=col, line_w=1.5)
    # left part — label, right part — number
    add_text_box(s, x + Inches(0.10), y, ew * 0.55, eh,
                 [{"text": lab, "size": 11, "color": BLACK,
                   "align": PP_ALIGN.LEFT, "bold": True}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, x + ew * 0.55, y, ew * 0.45 - Inches(0.05), eh,
                 [{"text": big, "size": 22, "bold": True, "color": col,
                   "align": PP_ALIGN.RIGHT}],
                 anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(s, x + ew - Inches(0.55), y + eh - Inches(0.32),
                 Inches(0.50), Inches(0.30),
                 [{"text": unit, "size": 9, "italic": True, "color": GREY,
                   "align": PP_ALIGN.RIGHT}],
                 anchor=MSO_ANCHOR.BOTTOM)

# Time savings highlight
add_rect(s, Inches(7.30), Inches(4.62), Inches(5.6), Inches(0.85), fill=NAVY)
add_text_box(s, Inches(7.40), Inches(4.62), Inches(5.4), Inches(0.85),
             [
                 {"text": "4 835 часов / мес.", "size": 22, "bold": True,
                  "color": WHITE, "align": PP_ALIGN.CENTER},
                 {"text": "сэкономленного рабочего времени по портфелю",
                  "size": 11, "italic": True, "color": WHITE, "align": PP_ALIGN.CENTER},
             ],
             anchor=MSO_ANCHOR.MIDDLE)

# Mini-table at bottom
small_caption(s, Inches(0.4), Inches(6.25), SLIDE_W - Inches(0.8), Inches(0.3),
              "Детализация по проектам (для проектора)", color=NAVY, size=12)
table_rows = [
    ["Проект",            "Инв., млн ₽", "Эконом./мес, млн ₽", "PP, мес.", "NPV, млн ₽", "PI"],
    ["Согласование смет", "2,92",        "0,23",                "12,7",     "6,55",        "3,25"],
    ["База цен",          "1,90",        "0,425",               "4,5",      "15,6",        "9,2"],
    ["Регламент КП",      "0,54",        "операционный",        "—",        "—",           "—"],
    ["Документооборот",   "0,81",        "2,235",               "0,4",      "91,2",        "113,6"],
    ["ИТОГО",             "6,17",        "2,89",                "≈2,1",     "112,9",       "19,3"],
]
ty3 = Inches(6.55); th3 = Inches(0.42)
total_w_t = SLIDE_W - Inches(0.8)
col_w = [0.32, 0.13, 0.18, 0.11, 0.13, 0.13]
for r, row in enumerate(table_rows):
    cx = Inches(0.4); y = ty3 + r * (Inches(0.15))
    is_header = (r == 0)
    is_total = (r == len(table_rows) - 1)
    bg = NAVY if is_header else (GREY_LIGHT if is_total else (WHITE if r % 2 == 1 else GREY_LIGHT))
    fg = WHITE if is_header else (NAVY if is_total else BLACK)
    bold = is_header or is_total
    for c, val in enumerate(row):
        w = total_w_t * col_w[c]
        cell = add_rect(s, cx, y, w, Inches(0.32), fill=bg, line=NAVY, line_w=0.4)
        set_text(cell, [{"text": val, "size": 9, "bold": bold, "color": fg,
                         "align": PP_ALIGN.CENTER if c > 0 else PP_ALIGN.LEFT}],
                 anchor=MSO_ANCHOR.MIDDLE)
        cx += w


# =========================================================================== #
#                            СЛАЙД 9 — ВЫВОДЫ                                  #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_header(s, "ВЫВОДЫ ПО РАБОТЕ")
add_slide_number(s, 9)

conclusions = [
    ("ГЛАВА 1",
     "Теория",
     "Для девелопмента оптимально сочетание BPMN 2.0 + Stormbpmn (импортозамещение, реестр Минцифры).",
     NAVY),
    ("ГЛАВА 2",
     "Отрасль",
     "РФ в стагнации; РТ +8% — рекорд. Доля федералов в Казани выросла с 18% до 32% — конкуренция требует процессной зрелости.",
     NAVY_LIGHT),
    ("ГЛАВА 3 ▸ ПРОБЛЕМЫ",
     "Диагностика",
     "В АО «Унистрой» выявлено 4 «узких места»: ручное согласование, децентрализованная база цен, отсутствие гарантии закупки, разорванный документооборот.",
     ACCENT),
    ("ГЛАВА 3 ▸ РЕШЕНИЯ",
     "Эффективность",
     "Портфель из 4 проектов: инвестиции 6,17 млн ₽ → экономия до 12,9 млн ₽/мес, NPV >112 млн ₽, PI = 19,3.",
     ACCENT_GREEN),
    ("ГЛАВА 3 ▸ ЗНАЧИМОСТЬ",
     "Тиражируемость",
     "Решения тиражируемы на других застройщиков отрасли — практическая значимость работы подтверждена.",
     NAVY),
]

cy2 = Inches(1.10); ch2 = Inches(1.05); cw_total = SLIDE_W - Inches(0.8)
for i, (chap, sub, txt, col) in enumerate(conclusions):
    y = cy2 + i * (ch2 + Inches(0.10))
    # left chap badge
    badge = add_round_rect(s, Inches(0.4), y, Inches(2.6), ch2, fill=col, line=col)
    set_text(badge, [
        {"text": chap, "size": 13, "bold": True, "color": WHITE,
         "align": PP_ALIGN.CENTER},
        {"text": sub, "size": 11, "italic": True, "color": WHITE,
         "align": PP_ALIGN.CENTER},
    ], anchor=MSO_ANCHOR.MIDDLE)
    # right text
    body = add_round_rect(s, Inches(3.10), y, cw_total - Inches(2.7), ch2,
                          fill=WHITE, line=col, line_w=1.5)
    set_text(body, [{"text": txt, "size": 13, "color": BLACK,
                     "align": PP_ALIGN.LEFT}],
             anchor=MSO_ANCHOR.MIDDLE)


# =========================================================================== #
#                  СЛАЙД 10 — СПАСИБО ЗА ВНИМАНИЕ                              #
# =========================================================================== #
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, fill=NAVY)

# Decorative diagonal stripes (3 thin lines)
for i, off in enumerate([Inches(0.5), Inches(0.7), Inches(0.9)]):
    add_rect(s, Inches(0.4) + off, Inches(0.9), Inches(0.05), SLIDE_H - Inches(1.8),
             fill=WHITE)
for i, off in enumerate([Inches(0.5), Inches(0.7), Inches(0.9)]):
    add_rect(s, SLIDE_W - Inches(0.45) - off, Inches(0.9), Inches(0.05),
             SLIDE_H - Inches(1.8), fill=WHITE)

add_text_box(s, Inches(1.0), Inches(2.5), SLIDE_W - Inches(2.0), Inches(1.5),
             [{"text": "СПАСИБО ЗА ВНИМАНИЕ!", "size": 54, "bold": True,
               "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
add_text_box(s, Inches(1.0), Inches(4.2), SLIDE_W - Inches(2.0), Inches(0.7),
             [{"text": "Готов(а) ответить на вопросы", "size": 22, "italic": True,
               "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)
add_text_box(s, Inches(1.0), Inches(5.4), SLIDE_W - Inches(2.0), Inches(0.5),
             [{"text": "АО «Унистрой» · Казань · 2026", "size": 14,
               "color": WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE)


# ----------------------------- Save ---------------------------------------- #
output = "/projects/sandbox/nick-onion/VKR_Unistroy_Presentation.pptx"
prs.save(output)
print(f"OK: {output}")
