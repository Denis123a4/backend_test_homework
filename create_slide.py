# -*- coding: utf-8 -*-
"""Generate a modern dark-themed PPTX slide (slide.pptx)."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree


def rgb(r, g, b):
    return RGBColor(r, g, b)


# ── Colour palette ────────────────────────────────────────────────────────
C_BG    = rgb(0x0D, 0x1B, 0x2A)   # background – dark navy
C_HDR   = rgb(0x0F, 0x27, 0x44)   # header card
C_CARD  = rgb(0x13, 0x24, 0x3A)   # body card
C_CARD2 = rgb(0x17, 0x2D, 0x4A)   # alternate card
C_TEAL  = rgb(0x4E, 0xCD, 0xC4)   # teal accent
C_GOLD  = rgb(0xF5, 0xC5, 0x42)   # gold accent
C_WHITE = rgb(0xFF, 0xFF, 0xFF)   # white
C_LIGHT = rgb(0xB8, 0xCC, 0xE0)   # light blue-grey
C_RTOP  = rgb(0x0A, 0x4D, 0x55)   # right-panel header bg
C_RBOT  = rgb(0x0D, 0x32, 0x38)   # right-panel body bg

IMG_COLOURS = [
    rgb(0x2D, 0x6A, 0x4F),   # forest green – ecology
    rgb(0x1D, 0x4E, 0x89),   # deep blue   – social
    rgb(0x6D, 0x35, 0x0E),   # earthy brown – event
    rgb(0x1A, 0x53, 0x5C),   # dark teal
    rgb(0x50, 0x2B, 0x5C),   # purple
    rgb(0x2C, 0x3E, 0x50),   # navy-grey
]

W = 13.333   # slide width  (inches, widescreen 16:9)
H = 7.5      # slide height (inches)

# ── Create presentation ───────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(W)
prs.slide_height = Inches(H)
slide = prs.slides.add_slide(prs.slide_layouts[6])   # blank

# Background
bg = slide.background
bg.fill.solid()
bg.fill.fore_color.rgb = C_BG


# ── Helpers ───────────────────────────────────────────────────────────────
def _no_line(shape):
    """Remove visible border from a shape via XML."""
    sp_pr = shape._element.spPr
    ln = sp_pr.find(qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(sp_pr, qn('a:ln'))
    for child in list(ln):
        ln.remove(child)
    etree.SubElement(ln, qn('a:noFill'))


def rect(x, y, w, h, color, border=False, border_color=C_LIGHT):
    """Add a solid-filled rectangle."""
    shp = slide.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if border:
        shp.line.color.rgb = border_color
        shp.line.width = Pt(0.5)
    else:
        _no_line(shp)
    return shp


def tb(x, y, w, h):
    """Add an empty text-box."""
    return slide.shapes.add_textbox(
        Inches(x), Inches(y), Inches(w), Inches(h))


def rfmt(run, size, bold=False, italic=False, color=C_WHITE):
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.italic = italic
    run.font.color.rgb = color


def para(tf, text, size=9.5, bold=False, italic=False,
         color=C_WHITE, align=PP_ALIGN.LEFT, sb=0):
    """Append a paragraph to a text frame."""
    p = tf.add_paragraph()
    p.alignment = align
    if sb:
        p.space_before = Pt(sb)
    r = p.add_run()
    r.text = text
    rfmt(r, size, bold, italic, color)
    return p


# ═══════════════════════════════════════════════════════════════════════════
# HEADER  (y 0 → 1.15)
# ═══════════════════════════════════════════════════════════════════════════
rect(0,    0,    W,    1.15, C_HDR)           # header bar
rect(0,    0,    0.12, 1.15, C_TEAL)          # left teal accent
rect(0,    1.15, W,    0.04, C_TEAL)          # bottom separator

# Title
tx = tb(0.2, 0.04, W - 0.3, 0.72)
tx.text_frame.word_wrap = True
p = tx.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = ('«ОТКРЫТОЕ СЕРДЦЕ». ВОЛОНТЁРСТВО КАК ПРОФИЛАКТИКА '
          'ДЕВИАНТНОГО ПОВЕДЕНИЯ ОБУЧАЮЩИХСЯ.')
rfmt(r, 19, bold=True)

# Teacher
tx2 = tb(0.25, 0.78, W - 0.5, 0.35)
p2 = tx2.text_frame.paragraphs[0]
r2 = p2.add_run()
r2.text = 'Педагог-психолог Вербицкая Н.А. Т. 8 905 556 72 80'
rfmt(r2, 11, color=C_TEAL)


# ═══════════════════════════════════════════════════════════════════════════
# QUOTE CARD  (y 1.22 → 2.17)
# ═══════════════════════════════════════════════════════════════════════════
rect(0.15, 1.22, W - 0.3, 0.93, C_CARD)
rect(0.15, 1.22, 0.07,   0.93, C_TEAL)       # left accent

tx3 = tb(0.28, 1.24, W - 0.5, 0.89)
tf3 = tx3.text_frame
tf3.word_wrap = True
p3 = tf3.paragraphs[0]
p3.alignment = PP_ALIGN.CENTER
r3 = p3.add_run()
r3.text = ('Волонтёрство в школьной среде рассматривается как одна из самых мощных '
           'методик позитивной профилактики. В отличие от традиционных лекций, оно не '
           'запрещает негативное поведение, а вытесняет его, предлагая подростку '
           'конструктивную альтернативу и чувство собственной значимости.')
rfmt(r3, 10, italic=True, color=C_LIGHT)


# ═══════════════════════════════════════════════════════════════════════════
# GOAL CARD  (y 2.19 → 2.96)
# ═══════════════════════════════════════════════════════════════════════════
rect(0.15, 2.19, W - 0.3, 0.76, C_CARD2)
rect(0.15, 2.19, 0.07,   0.76, C_GOLD)       # gold accent

tx4 = tb(0.28, 2.21, W - 0.5, 0.72)
tf4 = tx4.text_frame
tf4.word_wrap = True
p4 = tf4.paragraphs[0]
r4a = p4.add_run()
r4a.text = 'Основная цель : '
rfmt(r4a, 11, bold=True, color=C_GOLD)
r4b = p4.add_run()
r4b.text = ('Социализация и самореализация подростка через общественно полезную деятельность, '
            'которая формирует внутренний иммунитет к деструктивному поведению и зависимостям')
rfmt(r4b, 10.5, color=C_WHITE)


# ═══════════════════════════════════════════════════════════════════════════
# LEFT PANEL  (x 0.15, y 3.01 → 7.41)
# ═══════════════════════════════════════════════════════════════════════════
LP_X, LP_Y = 0.15, 3.01
LP_W, LP_H = 8.45, 4.40

rect(LP_X, LP_Y, LP_W, LP_H, C_CARD)
rect(LP_X, LP_Y, 0.07, LP_H, C_TEAL)

tx5 = tb(LP_X + 0.13, LP_Y + 0.1, LP_W - 0.2, LP_H - 0.15)
tf5 = tx5.text_frame
tf5.word_wrap = True

# Header line (first paragraph already exists)
p0 = tf5.paragraphs[0]
p0.alignment = PP_ALIGN.LEFT
r0 = p0.add_run()
r0.text = 'Задачи волонтерства как профилактики:'
rfmt(r0, 10.5, bold=True, color=C_TEAL)

BULLETS = [
    ('1.   Социально-педагогические задачи',                                                                                                 10,   True,  C_GOLD,  5),
    ('❖  Формирование ответственности: Когда подростку доверяют заботу о ком-то (детях, животных, экологии), он переходит из позиции «потребителя» в позицию «созидателя».',         9,    False, C_WHITE, 2),
    ('❖  Развитие эмпатии: Помощь другим помогает лучше понимать чувства людей, что является естественной профилактикой буллинга и агрессии.',                                        9,    False, C_WHITE, 2),
    ('❖  Приобретение новых навыков : Умение работать в команде, планировать время и вести переговоры повышает уверенность в себе.',                                                   9,    False, C_WHITE, 2),
    ('2. Психологические задачи.',                                                                                                           10,   True,  C_GOLD,  5),
    ('❖  Смена социального статуса: Для детей из «группы риска» волонтерство — это шанс избавиться от ярлыка «трудного ребенка» и получить признание в позитивном ключе.',            9,    False, C_WHITE, 2),
    ('❖  Снижение уровня тревожности и агрессии: Физическая и эмоциональная активность в волонтерских проектах помогает экологично «сбрасывать» накопленное напряжение.',              9,    False, C_WHITE, 2),
    ('❖  Поиск смысла и принадлежности: Волонтерская группа дает подростку чувство общности («я — часть важного дела»), что снижает риск попадания в опасные субкультуры.',           9,    False, C_WHITE, 2),
    ('3. Просветительские задачи',                                                                                                          10,   True,  C_GOLD,  5),
    ('❖  Обучение по принципу «Равный — равному»: Волонтеры-сверстники транслируют ценности здорового образа жизни эффективнее, чем взрослые.',                                       9,    False, C_WHITE, 2),
    ('❖  Правовое воспитание: Участие в социально значимых проектах формирует уважение к закону и нормам общества.',                                                                  9,    False, C_WHITE, 2),
]

for text, size, bold, color, sb in BULLETS:
    para(tf5, text, size=size, bold=bold, color=color, sb=sb)


# ═══════════════════════════════════════════════════════════════════════════
# RIGHT PANEL  (x 8.73, y 3.01 → 7.41)
# ═══════════════════════════════════════════════════════════════════════════
RP_X = 8.73
RP_Y = 3.01
RP_W = W - RP_X - 0.1
RP_H = 4.40

# Header card
rect(RP_X, RP_Y, RP_W, 1.0, C_RTOP)
tx6 = tb(RP_X + 0.1, RP_Y + 0.1, RP_W - 0.2, 0.82)
tf6 = tx6.text_frame
tf6.word_wrap = True
p6 = tf6.paragraphs[0]
p6.alignment = PP_ALIGN.CENTER
r6 = p6.add_run()
r6.text = 'Волонтерские направления: экологическое, социальное, событийное.'
rfmt(r6, 11, bold=True, color=C_WHITE)

# Body
RP_BY = RP_Y + 1.05
RP_BH = RP_H - 1.1
rect(RP_X, RP_BY, RP_W, RP_BH, C_RBOT)

# 6 photo placeholders – 2 rows × 3 columns
IMG_W = (RP_W - 0.28) / 3
IMG_H = (RP_BH - 0.21) / 2
GAP_X = 0.07
GAP_Y = 0.07

for row in range(2):
    for col in range(3):
        ix = RP_X + 0.07 + col * (IMG_W + GAP_X)
        iy = RP_BY + 0.07 + row * (IMG_H + GAP_Y)
        rect(ix, iy, IMG_W, IMG_H, IMG_COLOURS[row * 3 + col],
             border=True, border_color=C_LIGHT)


# ── Save ──────────────────────────────────────────────────────────────────
prs.save('slide.pptx')
print('slide.pptx saved successfully.')
