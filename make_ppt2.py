# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

DARK    = RGBColor(0x0F, 0x2E, 0x2A)
PRIMARY = RGBColor(0x1F, 0x6E, 0x5C)
GREEN2  = RGBColor(0x2E, 0x8B, 0x74)
ACCENT  = RGBColor(0xFF, 0x7A, 0x45)
LIGHT   = RGBColor(0xF4, 0xF7, 0xF6)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
TEXT    = RGBColor(0x1A, 0x1A, 0x1A)
GRAY    = RGBColor(0x6B, 0x77, 0x74)
RED     = RGBColor(0xC0, 0x3A, 0x2B)
GOLD    = RGBColor(0xF0, 0xA5, 0x00)

FONT = "微软雅黑"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

def _font(run, size, color=TEXT, bold=False):
    run.font.size = Pt(size); run.font.bold = bold; run.font.color.rgb = color; run.font.name = FONT
    r = run._r; rPr = r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', FONT)

def rect(s, x, y, w, h, color, line=False):
    shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = color
    if line: shp.line.color.rgb = color; shp.line.width = Pt(1)
    else: shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def text(s, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False; p.alignment = align
        for t, sz, c, b in line:
            r = p.add_run(); r.text = t; _font(r, sz, c, b)
    return tb

def bullets(s, x, y, w, h, items, size=15, gap=8, color=TEXT, marker="▪ ", mc=ACCENT):
    tb = s.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False; p.space_after = Pt(gap); p.line_spacing = 1.15
        r0 = p.add_run(); r0.text = marker; _font(r0, size, mc, True)
        if isinstance(it, tuple):
            r1 = p.add_run(); r1.text = it[0]; _font(r1, size, color, it[2] if len(it) > 2 else False)
            if it[1]:
                r2 = p.add_run(); r2.text = "　" + it[1]; _font(r2, size-3, GRAY, False)
        else:
            r1 = p.add_run(); r1.text = it; _font(r1, size, color, False)
    return tb

def section(title, subtitle, num):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, DARK)
    rect(s, Inches(0.6), Inches(2.9), Inches(0.18), Inches(1.7), ACCENT)
    text(s, Inches(1.0), Inches(2.55), Inches(11.3), Inches(0.6), [[("0"+str(num), 20, ACCENT, True)]])
    text(s, Inches(1.0), Inches(3.0), Inches(11.3), Inches(1.4), [[(title, 42, WHITE, True)]])
    text(s, Inches(1.0), Inches(4.4), Inches(11.0), Inches(1.0), [[(subtitle, 18, RGBColor(0xBF,0xD6,0xD1), False)]])
    return s

def header(s, title, num):
    rect(s, 0, 0, SW, Inches(1.05), WHITE)
    rect(s, 0, Inches(1.05), SW, Pt(3), PRIMARY)
    rect(s, Inches(0.55), Inches(0.32), Inches(0.14), Inches(0.42), ACCENT)
    text(s, Inches(0.85), Inches(0.2), Inches(2.0), Inches(0.6), [[(f"{num:02d}", 15, ACCENT, True)]])
    text(s, Inches(0.85), Inches(0.42), Inches(12.0), Inches(0.6), [[(title, 24, DARK, True)]])

# ============ 1 封面 ============
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, DARK)
rect(s, 0, 0, SW, Inches(0.35), ACCENT)
rect(s, Inches(7.2), 0, Inches(6.13), SH, PRIMARY)
ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.4), Inches(1.1), Inches(3.6), Inches(3.6))
ring.fill.background(); ring.line.color.rgb = RGBColor(0x3E,0x9E,0x86); ring.line.width = Pt(3)
ring2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.0), Inches(1.7), Inches(2.4), Inches(2.4))
ring2.fill.background(); ring2.line.color.rgb = ACCENT; ring2.line.width = Pt(2)
text(s, Inches(0.9), Inches(1.4), Inches(6.0), Inches(0.5), [[("AI 减肥系统", 22, ACCENT, True)]])
text(s, Inches(0.9), Inches(2.1), Inches(6.3), Inches(2.4),
     [[("为爱吃的东西", 40, WHITE, True)],
      [("付出有意义努力", 40, WHITE, True)]])
text(s, Inches(0.9), Inches(4.5), Inches(6.2), Inches(1.6),
     [[("地图算消耗 · 拍照估卡 · 余卡货币 · 存钱基因排行榜", 15, RGBColor(0xBF,0xD6,0xD1), False)]])
text(s, Inches(8.2), Inches(4.6), Inches(4.6), Inches(1.8),
     [[("产品全案（更新版）", 20, WHITE, True)],
      [("", 8, WHITE, False)],
      [("不是饿 · 是馋 · 试过总放弃", 16, ACCENT, True)]])

# ============ 2 目录 ============
s = prs.slides.add_slide(BLANK)
header(s, "目录", 1)
toc = [
    ("02", "一句话定位", "四大支柱一句话讲清"),
    ("03", "用户画像", "馋 / 不会吃 / 想被看见"),
    ("04", "POV 观点", "用户 · 需求 · 洞察"),
    ("05", "HMW 重构", "五个关键问题"),
    ("06", "核心功能与逻辑链", "算卡地基 → 估卡 → 余卡 → 社交"),
    ("07", "算卡地基", "运动端精确 · 摄入端兜误差"),
    ("08", "拍照估卡 + 成功率", "不假装精确，管可信度"),
    ("09", "记账 / 存款 / 借款", "核心货币：余卡"),
    ("10", "食谱升级 + 地图选菜", "推荐来自地图，走路是余量"),
    ("11", "余卡排行榜 · 存钱基因", "蓄水池 + 存钱本能"),
    ("12", "社区 + 晒图", "社交引擎 + 活广告"),
    ("13", "差异化", "四类竞品的不同"),
    ("14", "电梯陈述", "30 秒讲清楚"),
]
for i, (n, t, d) in enumerate(toc):
    col = i % 2
    yy = Inches(1.5) + Inches((i // 2) * 0.78)
    x = Inches(0.7 + col * 6.3)
    rect(s, x, yy, Inches(5.9), Inches(0.68), LIGHT)
    rect(s, x, yy, Inches(0.09), Inches(0.68), ACCENT)
    text(s, x + Inches(0.25), yy + Inches(0.12), Inches(0.9), Inches(0.5), [[(n, 18, PRIMARY, True)]])
    text(s, x + Inches(1.15), yy + Inches(0.08), Inches(4.5), Inches(0.5), [[(t, 13, TEXT, True)]])
    text(s, x + Inches(1.15), yy + Inches(0.36), Inches(4.5), Inches(0.3), [[(d, 10, GRAY, False)]])

# ============ 3 一句话定位 ============
s = prs.slides.add_slide(BLANK)
header(s, "一句话定位", 2)
rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(2.0), PRIMARY)
text(s, Inches(1.1), Inches(2.0), Inches(11.1), Inches(1.5),
     [[("一个以「地图距离+时间精确计算运动消耗」为地基、用「拍照AI估卡+成功率兜住误差」管理摄入、以「余卡（存款/借款）」为货币、", 16, WHITE, True)],
      [("并借「存钱心理」驱动的排行榜与交流社区，让用户主动多攒少花、从而实现可持续减肥的 AI 系统。", 16, WHITE, True)]],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.7), Inches(3.95), Inches(11.9), Inches(0.6),
     [[("核心是帮「不是饿、是馋、试过总放弃」的人，把减肥重构成", 16, TEXT, False),
       ("「为爱吃的东西付出有意义努力」", 16, ACCENT, True)]])
cards = [
    ("算卡地基", "地图距离+时间+体重+配速\n→ 精确运动消耗"),
    ("拍照估卡", "拍照AI估卡 + 成功率\n→ 管 ±20~40% 摄入误差"),
    ("余卡货币", "消耗−摄入=余卡\n→ 存款/借款，影响后续"),
    ("存钱基因", "余卡排行榜+社区\n→ 数字变大=爽，舍不得花"),
]
for i, (t, d) in enumerate(cards):
    x = Inches(0.7 + i * 3.05)
    rect(s, x, Inches(4.75), Inches(2.85), Inches(1.7), LIGHT)
    rect(s, x, Inches(4.75), Inches(2.85), Inches(0.12), ACCENT)
    text(s, x + Inches(0.2), Inches(4.95), Inches(2.5), Inches(0.5), [[(t, 15, PRIMARY, True)]])
    text(s, x + Inches(0.2), Inches(5.45), Inches(2.5), Inches(1.0), [[(d, 11, GRAY, False)]])

# ============ 4 用户画像 ============
s = prs.slides.add_slide(BLANK)
header(s, "用户画像 · 主用户：想「算卡+地图运动+能吃爱吃」的人", 3)
profiles = [
    ("画像A｜大基数想瘦又馋（核心）", ACCENT,
     [("现状", "体重偏高，试过节食、纯记步；难的不是饿，是馋"),
      ("痛点", "App 逼人克制、食谱难吃；记步只给数字，不给反馈"),
      ("渴望", "「运动够了就能吃爱吃的」看得见的正反馈")]),
    ("画像B｜有运动、不会管饮食", PRIMARY,
     [("现状", "常走路/跑步/骑车，但对摄入没概念"),
      ("痛点", "不知道这顿能吃到什么程度不会胖"),
      ("渴望", "把已有运动变成「能吃爱吃的额度」")]),
    ("画像C｜想被看见、爱攒东西", GREEN2,
     [("现状", "一个人减肥容易放弃"),
      ("痛点", "缺竞争、缺认同、缺「看得见的积累」"),
      ("渴望", "排行榜、晒图、互赞，看着数字一点点变大")]),
]
for i, (t, c, rows) in enumerate(profiles):
    x = Inches(0.7 + i * 4.1)
    rect(s, x, Inches(1.7), Inches(3.85), Inches(5.2), LIGHT)
    rect(s, x, Inches(1.7), Inches(3.85), Inches(0.95), c)
    text(s, x + Inches(0.2), Inches(1.82), Inches(3.45), Inches(0.75), [[(t, 14, WHITE, True)]])
    yy = 2.95
    for tag, desc in rows:
        text(s, x + Inches(0.25), Inches(yy), Inches(1.0), Inches(0.4), [[(tag, 12, c, True)]])
        text(s, x + Inches(1.25), Inches(yy), Inches(2.55), Inches(0.9), [[(desc, 11, TEXT, False)]])
        yy += 1.05

# ============ 5 POV ============
s = prs.slides.add_slide(BLANK)
header(s, "POV · 观点陈述", 4)
rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(1.6), DARK)
text(s, Inches(1.1), Inches(1.85), Inches(11.1), Inches(1.3),
     [[("大基数、爱吃、试过减肥却总因『馋』和『难吃』放弃的人，需要把『运动额度』变成『吃点好的』的可兑现货币，", 14, WHITE, True)],
      [("因为『为心爱食物付出努力 + 看见数字积累』的过程，会唤起比纯计数更持久的满足感与攒积欲，让减肥第一次可持续。", 14, WHITE, True)]])
pov = [
    ("User（用户）", "大基数、爱吃、屡败屡战的减肥者"),
    ("Need（需求）", "运动换算成「能吃爱吃的额度」，吃得明白、可控"),
    ("Insight（洞察）", "坚持靠的是「这段路为了吃到想吃的」的目标感 +「数字越攒越大」的本能爽感，两者叠加比计步更持久"),
]
for i, (k, v) in enumerate(pov):
    x = Inches(0.7 + i * 4.1)
    rect(s, x, Inches(3.7), Inches(3.85), Inches(2.1), WHITE)
    rect(s, x, Inches(3.7), Inches(3.85), Pt(2), PRIMARY)
    text(s, x + Inches(0.2), Inches(3.9), Inches(3.4), Inches(0.5), [[(k, 15, PRIMARY, True)]])
    text(s, x + Inches(0.2), Inches(4.4), Inches(3.45), Inches(1.2), [[(v, 12, TEXT, False)]])

# ============ 6 HMW ============
s = prs.slides.add_slide(BLANK)
header(s, "HMW · 我们如何能（问题重构）", 5)
hmw = [
    "如何让「运动额度」变成用户看得懂、够得着的「能吃爱吃的预算」？",
    "如何让吃进去的卡路里在存在 ±误差 的现实下，仍安全驱动「升级/记账/借款」？",
    "如何让「为爱吃的付出运动」和「看着数字变大」本身，成为比计数更快乐的体验？",
    "如何用排行榜+社区+存钱心理，把「存余卡」变成被羡慕、被点赞、想炫耀、不想轻易花的成就？",
    "如何做到「误差端保守兜底、精确端算准」，兑现「敢吃还能瘦」的承诺？",
]
bullets(s, Inches(0.9), Inches(1.7), Inches(11.5), Inches(5.2), hmw, size=16, gap=20, marker="Q", mc=ACCENT)

# ============ 7 核心功能 ============
s = prs.slides.add_slide(BLANK)
header(s, "核心功能与逻辑链", 6)
flow = [
    ("1 算卡地基", "运动端精确\n距离+时长+体重+配速"),
    ("2 拍照估卡", "吃进端有误差\n拍照AI+成功率"),
    ("3 余卡货币", "消耗−摄入=余卡\n存款/借款"),
    ("4 食谱升级", "推荐来自地图\n走路=纯余量"),
    ("5 社交引擎", "余卡排行榜\n社区晒图+点赞"),
]
for i, (t, d) in enumerate(flow):
    x = Inches(0.55 + i * 2.5)
    rect(s, x, Inches(1.7), Inches(2.25), Inches(1.9), PRIMARY)
    text(s, x + Inches(0.15), Inches(1.9), Inches(1.95), Inches(0.6), [[(t, 14, WHITE, True)]])
    text(s, x + Inches(0.15), Inches(2.5), Inches(1.95), Inches(1.0), [[(d, 11, WHITE, False)]])
    if i < 4:
        ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + Inches(2.28), Inches(2.4), Inches(0.3), Inches(0.5))
        ar.fill.solid(); ar.fill.fore_color.rgb = ACCENT; ar.line.fill.background(); ar.shadow.inherit = False
text(s, Inches(0.55), Inches(4.0), Inches(12.0), Inches(2.9),
     [[("命门：区分两端", 15, ACCENT, True)],
      [("运动消耗端（支付端）——务必精确：靠地图距离+时长+体重+配速，误差最小、可控。", 13, TEXT, False)],
      [("吃进去端（被支付端）——必然有误差：菜量/油量/拍摄估卡存在 20~40% 误差，不回避。", 13, TEXT, False)],
      [("把最能控制的一端算准，把误差最大的一端用「成功率+保守低估」管理住。", 13, GRAY, True)]])

# ============ 8 算卡地基 ============
s = prs.slides.add_slide(BLANK)
header(s, "1 · 算卡地基（运动消耗端）", 7)
text(s, Inches(0.7), Inches(1.6), Inches(11.9), Inches(0.8),
     [[("不靠计步数，靠可验证的物理量算消耗，是整个系统最可靠的支柱。", 16, TEXT, False)]])
formula = [("距离", "地图路径"), ("时长", "出发→到店计时"), ("体重", "用户画像"), ("配速", "走路/跑步/骑车")]
for i, (k, v) in enumerate(formula):
    x = Inches(0.7 + i * 3.05)
    rect(s, x, Inches(2.6), Inches(2.85), Inches(1.5), LIGHT)
    rect(s, x, Inches(2.6), Inches(2.85), Inches(0.5), PRIMARY)
    text(s, x, Inches(2.68), Inches(2.85), Inches(0.4), [[(k, 15, WHITE, True)]], align=PP_ALIGN.CENTER)
    text(s, x, Inches(3.25), Inches(2.85), Inches(0.8), [[(v, 13, TEXT, False)]], align=PP_ALIGN.CENTER)
    if i < 3:
        text(s, x + Inches(2.85), Inches(2.95), Inches(0.35), Inches(0.6), [[("→", 22, ACCENT, True)]])
text(s, Inches(0.7), Inches(4.5), Inches(11.9), Inches(1.8),
     [[("消耗公式（准确、可复算）", 15, ACCENT, True)],
      [("消耗 = MET × 体重(kg) × 时长(h)；MET 由配速查表决定。距离+时间→配速→MET→消耗，每一步可追溯。", 13, TEXT, False)],
      [("这是「敢吃还能瘦」承诺的底气：给额度的是算得准的运动，不是粗放的步数估算。", 13, GRAY, True)]])

# ============ 9 拍照估卡 ============
s = prs.slides.add_slide(BLANK)
header(s, "2 · 拍照估卡 + 成功率（吃进去端）", 8)
text(s, Inches(0.7), Inches(1.55), Inches(11.9), Inches(0.8),
     [[("拍照AI分析菜品卡路里，数据上传沉淀、越用越准；误差客观存在，不假装精确，用「成功率」管理可信度。", 14, TEXT, False)]])
rates = [
    ("刚好卡点", "成功率低", "风险提示：刚好卡线，吃完可能超支", RED),
    ("有余量 30%", "成功率高", "余量足，放心吃", PRIMARY),
    ("余量更足", "成功率更高", "余量越多越有把握，越敢吃", GREEN2),
]
for i, (t, r, d, c) in enumerate(rates):
    x = Inches(0.7 + i * 4.1)
    rect(s, x, Inches(2.6), Inches(3.85), Inches(1.7), LIGHT)
    rect(s, x, Inches(2.6), Inches(3.85), Inches(0.55), c)
    text(s, x + Inches(0.2), Inches(2.68), Inches(3.45), Inches(0.4), [[(t, 14, WHITE, True)]])
    text(s, x + Inches(0.2), Inches(3.25), Inches(3.45), Inches(0.5), [[(r, 15, c, True)]])
    text(s, x + Inches(0.2), Inches(3.75), Inches(3.45), Inches(0.55), [[(d, 11, GRAY, False)]])
text(s, Inches(0.7), Inches(4.6), Inches(11.9), Inches(1.4),
     [[("记住", 14, ACCENT, True)],
      [("成功率不是替代算卡，而是在「算卡之上」管理「这个卡数多可信」——卡照算，成功率管可信度。", 13, TEXT, True)]])

# ============ 10 记账/存款/借款 ============
s = prs.slides.add_slide(BLANK)
header(s, "3 · 记账 / 存款 / 借款（核心货币：余卡）", 9)
rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(1.1), DARK)
text(s, Inches(1.0), Inches(1.85), Inches(11.3), Inches(0.8),
     [[("余卡 = 可吃额度 − 实际吃掉", 20, WHITE, True),
       ("（可吃额度 = 餐额 + 走到店 + 今日运动）", 14, ACCENT, True)]])
rules = [
    ("余卡为正 → 存款", "多动少花，攒下额度", PRIMARY),
    ("透支 → 借款", "吃超了记一笔借款", RED),
    ("借款影响后续", "欠卡后食谱/升级额度收紧", ACCENT),
    ("保守低估兜底", "余卡按最低值展示，标「至少余XX卡」", GREEN2),
]
for i, (t, d, c) in enumerate(rules):
    x = Inches(0.7 + i * 3.05)
    rect(s, x, Inches(3.1), Inches(2.85), Inches(2.0), LIGHT)
    rect(s, x, Inches(3.1), Inches(2.85), Inches(0.12), c)
    text(s, x + Inches(0.2), Inches(3.3), Inches(2.5), Inches(0.9), [[(t, 15, c, True)]])
    text(s, x + Inches(0.2), Inches(4.1), Inches(2.5), Inches(0.9), [[(d, 11, GRAY, False)]])
text(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.0),
     [[("为什么「敢吃还能瘦」不倒：", 14, ACCENT, True),
       ("余卡按保守值展示，宁可少算不吹，保住承诺的可信度。", 13, TEXT, False)]])

# ============ 11 食谱升级 + 地图选菜 ============
s = prs.slides.add_slide(BLANK)
header(s, "4 · 食谱升级 + 地图选菜 + 赚余量", 10)
text(s, Inches(0.7), Inches(1.55), Inches(11.9), Inches(0.7),
     [[("推荐菜直接来自地图：附近低卡美食 + 标出店面；走到店的消耗是纯「余量」加成，不占这顿饭额度。", 15, TEXT, False)]])
box = [
    ("地图选低卡", "AI 从附近店挑低卡菜推荐，标店面+距离", PRIMARY),
    ("走路=余量", "推荐 1km 兰州拉面 → 走到店消耗 100 千卡 → 就有 100 千卡余量", ACCENT),
    ("赚余量", "微信步数 / 跑步机 / 自由填运动，AI 按 MET 估消耗", GREEN2),
]
for i, (t, d, c) in enumerate(box):
    y = Inches(2.5) + Inches(i * 1.3)
    rect(s, Inches(0.7), y, Inches(2.8), Inches(1.1), c)
    text(s, Inches(0.9), y + Inches(0.3), Inches(2.5), Inches(0.6), [[(t, 14, WHITE, True)]])
    rect(s, Inches(3.55), y, Inches(9.05), Inches(1.1), LIGHT)
    text(s, Inches(3.8), y + Inches(0.25), Inches(8.6), Inches(0.7), [[(d, 12, TEXT, False)]], anchor=MSO_ANCHOR.MIDDLE)
text(s, Inches(0.7), Inches(6.5), Inches(11.9), Inches(0.8),
     [[("关键：计划不包含走到店的消耗——", 14, ACCENT, True),
       ("消耗是额外赚到的余量，走越远吃越好，还留了灵活口子。", 13, TEXT, False)]])

# ============ 12 余卡排行榜 · 存钱基因 ============
s = prs.slides.add_slide(BLANK)
header(s, "5 · 余卡排行榜 · 核心心理机制「存钱基因」", 11)
text(s, Inches(0.7), Inches(1.5), Inches(11.9), Inches(0.7),
     [[("在「过程灵活性」上的关键升级，两个机制叠加：", 15, TEXT, True)]])
rect(s, Inches(0.7), Inches(2.35), Inches(5.8), Inches(3.4), LIGHT)
rect(s, Inches(0.7), Inches(2.35), Inches(5.8), Inches(0.6), PRIMARY)
text(s, Inches(0.9), Inches(2.45), Inches(5.4), Inches(0.5), [[("机制一 · 余量 = 时间蓄水池", 15, WHITE, True)]])
bullets(s, Inches(0.95), Inches(3.15), Inches(5.3), Inches(2.5),
        ["不逼用户天天运动满额",
         "某天没时间运动？只要之前攒的余卡够，照样吃顿好的",
         "余卡 = 时间蓄水池 / 信用额度，更灵活、更包容"],
        size=12, gap=10)
rect(s, Inches(6.85), Inches(2.35), Inches(5.85), Inches(3.4), LIGHT)
rect(s, Inches(6.85), Inches(2.35), Inches(5.85), Inches(0.6), ACCENT)
text(s, Inches(7.05), Inches(2.45), Inches(5.5), Inches(0.5), [[("机制二 · 借「存钱」本能", 15, WHITE, True)]])
bullets(s, Inches(7.1), Inches(3.15), Inches(5.4), Inches(2.5),
        ["数字摆在眼前，就想让它变大、想存、舍不得花",
         "数字变大 = 成就感 = 爽；变小 = 舍不得 = 主动克制",
         "成功率是设计出来、被心理机制托高的，而非咬牙硬扛"],
        size=12, gap=10)
rect(s, Inches(0.7), Inches(6.0), Inches(11.9), Inches(1.0), DARK)
text(s, Inches(1.0), Inches(6.15), Inches(11.3), Inches(0.7),
     [[("落地：按体重/基础/目标分层防逼卷 · 加「超越昨天的自己」进步榜 · 冷数据零内容也能跑，适合先上线。", 13, WHITE, False)]])

# ============ 13 社区 + 晒图 ============
s = prs.slides.add_slide(BLANK)
header(s, "6 · 交流社区 + 拍照晒图（社交引擎 + 活广告）", 12)
left = [
    ("晒图主题", "核心「存余卡+成功吃上爱吃的却仍在预算内=余卡健康」，不是谁吃得爽"),
    ("炫耀模板", "「走了3800米，2520大卡，终于换了这顿烤肉，余卡还剩180」"),
    ("衔接存钱心理", "晒的正是「数字又变大了、吃上了想吃的还没乱花」，点赞激励多攒多晒"),
    ("正循环", "点赞互动 → 多存少花发内容 → 内容成为活广告"),
]
text(s, Inches(0.7), Inches(1.65), Inches(6.3), Inches(0.5), [[("社交引擎", 15, PRIMARY, True)]])
bullets(s, Inches(0.9), Inches(2.2), Inches(6.0), Inches(4.6), left, size=13, gap=12)
rect(s, Inches(7.3), Inches(1.7), Inches(5.4), Inches(4.9), DARK)
text(s, Inches(7.6), Inches(1.9), Inches(4.8), Inches(0.5), [[("内容 = 活广告", 16, ACCENT, True)]])
bullets(s, Inches(7.7), Inches(2.5), Inches(4.7), Inches(3.9),
        ["用户自发晒图传播，为产品拉新",
         "冷启动：先做排行榜（零内容能跑），社区轻功能先挂",
         "用打卡奖励制造初始内容",
         "低门槛晒图：运动轨迹图（走/跑地图打卡，无羞耻、高满足）"],
        size=12, gap=12, color=WHITE, mc=ACCENT)

# ============ 14 差异化 ============
s = prs.slides.add_slide(BLANK)
header(s, "差异化（纯产品视角）", 13)
diffs = [
    ("对纯计步/工具App", "不是记步，而是「距离+时间→精确运动消耗」，给出「能吃爱吃的额度」"),
    ("对纯卡路里App", "不假装精确，用「成功率管理误差+保守低估」，做到「敢吃还能瘦」"),
    ("对食谱/健身博主", "不逼挨饿，把减肥过程灵活化、个性化、可兑现"),
    ("对一般打卡/社区App", "不只靠自律和游戏化，用中国人天然的「存钱心理」驱动主动攒积——难复制、最贴本土的行为引擎"),
]
for i, (t, d) in enumerate(diffs):
    y = Inches(1.6) + Inches(i * 1.35)
    rect(s, Inches(0.7), y, Inches(3.2), Inches(1.15), PRIMARY)
    text(s, Inches(0.9), y + Inches(0.3), Inches(2.9), Inches(0.6), [[(t, 13, WHITE, True)]])
    rect(s, Inches(3.95), y, Inches(8.65), Inches(1.15), LIGHT)
    text(s, Inches(4.2), y + Inches(0.25), Inches(8.2), Inches(0.7), [[(d, 13, TEXT, False)]], anchor=MSO_ANCHOR.MIDDLE)

# ============ 15 电梯陈述 ============
s = prs.slides.add_slide(BLANK)
header(s, "一页纸 · 电梯陈述", 14)
rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(4.3), DARK)
text(s, Inches(1.15), Inches(2.0), Inches(11.0), Inches(3.8),
     [[("我们做的是把减肥重构成『为爱吃的东西付出有意义努力 + 看着数字积累变大的爽感』的 AI 系统。", 17, WHITE, True)],
      [("", 7, WHITE, False)],
      [("它以地图距离+时间精确算出运动消耗为地基，用拍照AI估卡+成功率管理误差管理吃进去的额度，把剩余额度变成可存可借的『余卡』——卡点提示风险、有余量放心吃。", 14, RGBColor(0xCF,0xE0,0xDB), False)],
      [("", 7, WHITE, False)],
      [("关键是『余卡排行榜』借了中国人存钱的本能：数字摆在那，你就想让它变大、舍不得花，还要晒给朋友看；没时间运动也没关系，攒下的余卡是你随时能吃顿好的底气。", 14, RGBColor(0xCF,0xE0,0xDB), False)],
      [("", 7, WHITE, False)],
      [("我们帮的是『不是饿、是馋、试过总放弃』的人——不承诺100%能瘦，而是让『攒着、看着、吃着、瘦着』变得自然而然。", 15, ACCENT, True)]])

# ============ 16 结尾 ============
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, DARK)
rect(s, 0, 0, SW, Inches(0.35), ACCENT)
text(s, Inches(1.0), Inches(2.6), Inches(11.3), Inches(1.2),
     [[("把减肥，变成一件值得坚持的事。", 40, WHITE, True)]], align=PP_ALIGN.CENTER)
text(s, Inches(1.0), Inches(4.0), Inches(11.3), Inches(0.8),
     [[("不是饿 · 是馋 · 但这次你「有把握」", 18, ACCENT, True)]], align=PP_ALIGN.CENTER)

desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
out = os.path.join(desktop, "AI减肥系统_产品全案_更新版.pptx")
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
