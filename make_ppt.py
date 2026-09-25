# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- 主题色 ----------
DARK    = RGBColor(0x0F, 0x2E, 0x2A)   # 深墨绿
PRIMARY = RGBColor(0x1F, 0x6E, 0x5C)   # 主绿
GREEN2  = RGBColor(0x2E, 0x8B, 0x74)
ACCENT  = RGBColor(0xFF, 0x7A, 0x45)   # 橙
LIGHT   = RGBColor(0xF4, 0xF7, 0xF6)   # 浅底
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
TEXT    = RGBColor(0x1A, 0x1A, 0x1A)
GRAY    = RGBColor(0x6B, 0x77, 0x74)

FONT = "微软雅黑"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

def _set_font(run, size, color=TEXT, bold=False, font=FONT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    r = run._r
    rPr = r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)

def add_rect(slide, x, y, w, h, color, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if line:
        shp.line.color.rgb = color
        shp.line.width = Pt(1)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def add_text(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """lines: list of list-of-(text,size,color,bold) OR list of dict runs"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if isinstance(line, dict):
            line = [line]
        for seg in line:
            txt, size, color, bold = seg
            r = p.add_run()
            r.text = txt
            _set_font(r, size, color, bold)
    return tb

def add_bullets(slide, x, y, w, h, items, size=15, gap=8, color=TEXT, marker="▪ ", marker_color=ACCENT, line_spacing=1.15):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        p.line_spacing = line_spacing
        r0 = p.add_run(); r0.text = marker
        _set_font(r0, size, marker_color, True)
        # it may be (text) or (text, sub, bold)
        if isinstance(it, tuple):
            main, sub = it[0], it[1]
            bold = it[2] if len(it) > 2 else False
            r1 = p.add_run(); r1.text = main
            _set_font(r1, size, color, bold)
            if sub:
                r2 = p.add_run(); r2.text = "　" + sub
                _set_font(r2, size - 3, GRAY, False)
        else:
            r1 = p.add_run(); r1.text = it
            _set_font(r1, size, color, False)
    return tb

def section_slide(title, subtitle, num):
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, DARK)
    add_rect(s, Inches(0.6), Inches(2.9), Inches(0.18), Inches(1.7), ACCENT)
    add_text(s, Inches(1.0), Inches(2.55), Inches(11.3), Inches(0.6),
             [[(f"0{num}", 20, ACCENT, True)]])
    add_text(s, Inches(1.0), Inches(3.0), Inches(11.3), Inches(1.4),
             [[(title, 44, WHITE, True)]])
    add_text(s, Inches(1.0), Inches(4.4), Inches(11.0), Inches(1.0),
             [[(subtitle, 18, RGBColor(0xBF, 0xD6, 0xD1), False)]])
    return s

def header(slide, title, num):
    add_rect(slide, 0, 0, SW, Inches(1.05), WHITE)
    add_rect(slide, 0, Inches(1.05), SW, Pt(3), PRIMARY)
    add_rect(slide, Inches(0.55), Inches(0.32), Inches(0.14), Inches(0.42), ACCENT)
    add_text(slide, Inches(0.85), Inches(0.22), Inches(2.0), Inches(0.6),
             [[(f"{num:02d}", 15, ACCENT, True)]])
    add_text(slide, Inches(0.85), Inches(0.42), Inches(12.0), Inches(0.6),
             [[(title, 24, DARK, True)]])

# ================= 1. 封面 =================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), ACCENT)
add_rect(s, Inches(7.2), 0, Inches(6.13), SH, PRIMARY)
add_rect(s, Inches(7.2), Inches(3.3), Inches(6.13), Inches(0.12), ACCENT)
# 装饰圆环
ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8.4), Inches(1.1), Inches(3.6), Inches(3.6))
ring.fill.background(); ring.line.color.rgb = RGBColor(0x3E, 0x9E, 0x86); ring.line.width = Pt(3)
ring2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.0), Inches(1.7), Inches(2.4), Inches(2.4))
ring2.fill.background(); ring2.line.color.rgb = ACCENT; ring2.line.width = Pt(2)

add_text(s, Inches(0.9), Inches(1.5), Inches(6.0), Inches(0.5),
         [[("AI 减肥操作系统", 22, ACCENT, True)]])
add_text(s, Inches(0.9), Inches(2.2), Inches(6.2), Inches(2.4),
         [[("为爱吃的东西", 40, WHITE, True)],
          [("付出有意义努力", 40, WHITE, True)]])
add_text(s, Inches(0.9), Inches(4.6), Inches(6.2), Inches(1.6),
         [[("地图距离+时间算消耗 · 拍照估卡+成功率管误差", 15, RGBColor(0xBF, 0xD6, 0xD1), False)],
          [("余卡当货币 · 排行榜+社区驱动坚持", 15, RGBColor(0xBF, 0xD6, 0xD1), False)]])

add_text(s, Inches(8.2), Inches(4.6), Inches(4.6), Inches(1.8),
         [[("产品全案", 20, WHITE, True)],
          [("（纯产品逻辑版）", 14, RGBColor(0xBF, 0xD6, 0xD1), False)],
          [("", 8, WHITE, False)],
          [("不是饿 · 是馋 · 试过总放弃", 16, ACCENT, True)]])

# ================= 2. 目录 =================
s = prs.slides.add_slide(BLANK)
header(s, "目录", 1)
toc = [
    ("02", "一句话定位", "把减肥重构为「为爱吃的东西付出有意义努力」"),
    ("03", "用户画像", "三类核心用户：馋、不会吃、想被看见"),
    ("04", "POV 观点", "用设计思维拆解用户·需求·洞察"),
    ("05", "HMW 重构", "五个关键问题，框定产品边界"),
    ("06", "核心功能与逻辑链", "算卡地基 → 拍照估卡 → 余卡货币 → 社交引擎"),
    ("07", "差异化", "对计步/卡路里/博主/地图大厂的不同"),
    ("08", "电梯陈述", "30 秒讲清楚我们在做什么"),
]
y = 1.45
for i, (n, t, d) in enumerate(toc):
    col = i % 2
    x = Inches(0.7 + col * 6.3)
    if i >= 2:
        yy = y + (i - 2) * 0.0
    yy = Inches(1.55) + Inches((i // 2) * 2.0)
    add_rect(s, Inches(0.7 + col * 6.3), yy, Inches(5.9), Inches(1.75), LIGHT)
    add_rect(s, Inches(0.7 + col * 6.3), yy, Inches(0.1), Inches(1.75), ACCENT)
    add_text(s, Inches(1.0 + col * 6.3), yy + Inches(0.28), Inches(0.9), Inches(0.6),
             [[(n, 26, PRIMARY, True)]])
    add_text(s, Inches(2.0 + col * 6.3), yy + Inches(0.25), Inches(4.4), Inches(0.6),
             [[(t, 17, TEXT, True)]])
    add_text(s, Inches(2.0 + col * 6.3), yy + Inches(0.85), Inches(4.4), Inches(0.7),
             [[(d, 12, GRAY, False)]])

# ================= 3. 一句话定位 =================
s = prs.slides.add_slide(BLANK)
header(s, "一句话定位", 2)
add_rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(2.0), PRIMARY)
add_text(s, Inches(1.1), Inches(2.0), Inches(11.1), Inches(1.5),
         [[("一个以「地图距离+时间精确计算运动消耗」为地基、用「拍照AI估卡+成功率兜住误差」管理摄入、", 19, WHITE, True)],
          [("以「余卡（存款/借款）」为货币、配排行榜与交流社区驱动坚持的 AI 减肥系统。", 19, WHITE, True)]],
         anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.7), Inches(3.95), Inches(11.9), Inches(0.6),
         [[("核心是帮「不是饿、是馋、试过总放弃」的人，把减肥重构成", 16, TEXT, False),
           ("「为爱吃的东西付出有意义努力」", 16, ACCENT, True)]])
add_text(s, Inches(0.7), Inches(4.65), Inches(11.9), Inches(2.3),
         [[("四大支柱", 14, GRAY, True)]], )
cards = [
    ("算卡地基", "地图距离+时长+体重+配速\n→ 精确运动消耗（可靠支柱）"),
    ("拍照估卡", "拍照AI估卡 + 成功率\n→ 管理±20~40%摄入误差"),
    ("余卡货币", "消耗−摄入=余卡\n→ 存款/借款，影响后续标准"),
    ("社交引擎", "余卡排行榜+社区晒图\n→ 把存卡变被羡慕的成就"),
]
for i, (t, d) in enumerate(cards):
    x = Inches(0.7 + i * 3.05)
    add_rect(s, x, Inches(5.15), Inches(2.85), Inches(1.7), LIGHT)
    add_rect(s, x, Inches(5.15), Inches(2.85), Inches(0.12), ACCENT)
    add_text(s, x + Inches(0.2), Inches(5.35), Inches(2.5), Inches(0.5), [[(t, 15, PRIMARY, True)]])
    add_text(s, x + Inches(0.2), Inches(5.85), Inches(2.5), Inches(0.95), [[(d, 11, GRAY, False)]])

# ================= 4. 用户画像 =================
s = prs.slides.add_slide(BLANK)
header(s, "用户画像 · 主用户：想「算卡+地图运动+能吃爱吃」的人", 3)
profiles = [
    ("画像A｜大基数想瘦又馋的人（最核心）", ACCENT,
     [("现状", "体重偏高（如100kg），试过节食、纯记步"),
      ("痛点", "不是饿，是馋；食谱难吃导致坚持不下去"),
      ("渴望", "想要「运动够了，就能吃顿爱吃的」的正反馈")]),
    ("画像B｜有运动习惯、但不会管饮食", PRIMARY,
     [("现状", "常走路/跑步/骑车，但对摄入没概念"),
      ("痛点", "不知道「这顿能吃到什么程度不会胖」"),
      ("渴望", "把已有运动，变成「能吃爱吃的额度」")]),
    ("画像C｜想被看见、想有人一起坚持", GREEN2,
     [("现状", "一个人减肥容易放弃"),
      ("痛点", "缺竞争和认同感"),
      ("渴望", "排行榜、晒图、社区互赞的激励")]),
]
for i, (t, c, rows) in enumerate(profiles):
    x = Inches(0.7 + i * 4.1)
    add_rect(s, x, Inches(1.7), Inches(3.85), Inches(5.1), LIGHT)
    add_rect(s, x, Inches(1.7), Inches(3.85), Inches(0.9), c)
    add_text(s, x + Inches(0.2), Inches(1.82), Inches(3.45), Inches(0.7), [[(t, 14, WHITE, True)]])
    yy = 2.85
    for tag, desc in rows:
        add_text(s, x + Inches(0.25), Inches(yy), Inches(1.0), Inches(0.4), [[(tag, 12, c, True)]])
        add_text(s, x + Inches(1.25), Inches(yy), Inches(2.5), Inches(0.9), [[(desc, 11, TEXT, False)]])
        yy += 0.95

# ================= 5. POV =================
s = prs.slides.add_slide(BLANK)
header(s, "POV · 观点陈述", 4)
add_rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(1.6), DARK)
add_text(s, Inches(1.1), Inches(1.85), Inches(11.1), Inches(1.3),
         [[("大基数、爱吃、试过减肥却总因『馋』和『难吃』放弃的人，需要把『运动额度』变成『吃点好的』的可兑现货币，", 15, WHITE, True)],
          [("因为『为心爱食物付出努力的过程』本身就会带来比纯计数更有意义的满足感，从而让减肥第一次变得可持续。", 15, WHITE, True)]])
pov = [
    ("User（用户）", "大基数、爱吃、屡败屡战的减肥者"),
    ("Need（需求）", "运动换算成「能吃爱吃的额度」，吃得明白、控制在预算内"),
    ("Insight（洞察）", "让人坚持的不是「记了多少」，而是「这段路是为了吃到想吃的」——目标感与延迟满足比单纯计步更持久"),
]
for i, (k, v) in enumerate(pov):
    x = Inches(0.7 + i * 4.1)
    add_rect(s, x, Inches(3.7), Inches(3.85), Inches(1.9), WHITE)
    add_rect(s, x, Inches(3.7), Inches(3.85), Pt(2), PRIMARY)
    add_text(s, x + Inches(0.2), Inches(3.9), Inches(3.4), Inches(0.5), [[(k, 15, PRIMARY, True)]])
    add_text(s, x + Inches(0.2), Inches(4.4), Inches(3.45), Inches(1.1), [[(v, 12, TEXT, False)]])

# ================= 6. HMW =================
s = prs.slides.add_slide(BLANK)
header(s, "HMW · 我们如何能（问题重构）", 5)
hmw = [
    "如何让「运动额度」变成用户看得懂、够得着的「能吃爱吃的预算」？",
    "如何让吃进去的卡路里，在「存在±误差」的现实下，仍能安全驱动「升级/记账/借款」？",
    "如何让「为爱吃的付出运动」本身，成为比单纯计数更快乐的体验？",
    "如何用排行榜+社区，把「存余卡」变成被羡慕、被点赞、想炫耀的成就？",
    "如何做到「误差端（吃进去）保守兜底、精确端（运动消耗）算准」，兑现「敢吃还能瘦」？",
]
add_bullets(s, Inches(0.9), Inches(1.7), Inches(11.5), Inches(5.2),
            hmw, size=16, gap=18, marker="Q", marker_color=ACCENT)

# ================= 7. 核心功能总览 =================
s = prs.slides.add_slide(BLANK)
header(s, "核心功能与逻辑链", 6)
flow = [
    ("1 · 算卡地基", "运动消耗端精确\n距离+时长+体重+配速"),
    ("2 · 拍照估卡", "吃进去端有误差\n拍照AI + 成功率"),
    ("3 · 余卡货币", "消耗−摄入=余卡\n存款 / 借款"),
    ("4 · 食谱升级", "AI定制低卡可入口\n升级依据=卡数"),
    ("5 · 社交引擎", "余卡排行榜\n社区晒图+点赞"),
]
for i, (t, d) in enumerate(flow):
    x = Inches(0.55 + i * 2.5)
    add_rect(s, x, Inches(1.7), Inches(2.25), Inches(1.9), PRIMARY)
    add_text(s, x + Inches(0.15), Inches(1.9), Inches(1.95), Inches(0.6), [[(t, 15, WHITE, True)]])
    add_text(s, x + Inches(0.15), Inches(2.55), Inches(1.95), Inches(1.0), [[(d, 11, WHITE, False)]])
    if i < 4:
        ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + Inches(2.28), Inches(2.4), Inches(0.3), Inches(0.5))
        ar.fill.solid(); ar.fill.fore_color.rgb = ACCENT; ar.line.fill.background(); ar.shadow.inherit = False
add_text(s, Inches(0.55), Inches(4.0), Inches(12.0), Inches(2.9),
         [[("关键原则：算卡是根基地，不可动摇", 15, ACCENT, True)],
          [("运动消耗端（支付端）——务必精确：不靠计步数，靠地图距离+时长+体重+配速（误差最小、可控）。", 13, TEXT, False)],
          [("吃进去端（被支付端）——必然有误差：菜量/油量/拍摄估卡存在20~40%误差，这是现实，不回避。", 13, TEXT, False)],
          [("区分两端是整个方案的命门：把最能控制的一端（运动消耗）算准，把误差最大的一端（吃进去）用「成功率+保守低估」管理住。", 13, GRAY, True)]])

# ================= 8. 算卡地基 =================
s = prs.slides.add_slide(BLANK)
header(s, "1 · 算卡地基（运动消耗端）", 7)
add_text(s, Inches(0.7), Inches(1.6), Inches(11.9), Inches(0.8),
         [[("不靠计步数，而是靠可验证的物理量算消耗，这是整个系统最可靠的支柱。", 16, TEXT, False)]])
formula = [
    ("距离", "地图路径距离"),
    ("时长", "开始计时→到店"),
    ("体重", "用户画像体重"),
    ("配速", "跑步/步行/速度"),
]
for i, (k, v) in enumerate(formula):
    x = Inches(0.7 + i * 3.05)
    add_rect(s, x, Inches(2.6), Inches(2.85), Inches(1.5), LIGHT)
    add_rect(s, x, Inches(2.6), Inches(2.85), Inches(0.5), PRIMARY)
    add_text(s, x, Inches(2.68), Inches(2.85), Inches(0.4), [[(k, 15, WHITE, True)]], align=PP_ALIGN.CENTER)
    add_text(s, x, Inches(3.25), Inches(2.85), Inches(0.8), [[(v, 13, TEXT, False)]], align=PP_ALIGN.CENTER)
    if i < 3:
        plus = s.shapes.add_textbox(x + Inches(2.85), Inches(2.9), Inches(0.35), Inches(0.6))
        pf = plus.text_frame; p = pf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = "→"; _set_font(r, 22, ACCENT, True)
add_text(s, Inches(0.7), Inches(4.5), Inches(11.9), Inches(1.6),
         [[("为什么它「敢给额度」", 15, ACCENT, True)],
          [("距离+时间→配速→代谢当量（MET）→消耗，每一步都可追溯、可复算；", 13, TEXT, False)],
          [("相比「步数×估算」的粗放，这套算法误差小、可控，是「敢吃还能瘦」承诺的底气。", 13, GRAY, True)]])

# ================= 9. 拍照估卡 + 成功率 =================
s = prs.slides.add_slide(BLANK)
header(s, "2 · 拍照估卡 + 成功率（吃进去端）", 8)
add_text(s, Inches(0.7), Inches(1.55), Inches(11.9), Inches(0.8),
         [[("拍照AI分析菜品卡路里，数据上传服务器持续沉淀、越用越准；误差客观存在，系统不假装精确，而是用「成功率」管理可信度。", 14, TEXT, False)]])
rates = [
    ("刚好卡点", "成功率低", "风险提示：这次刚好卡线，吃完可能超支", RGBColor(0xC0, 0x3A, 0x2B)),
    ("有余量 30%", "成功率高", "这次余量足，放心吃", PRIMARY),
    ("余量更足", "成功率更高", "余量越多，越有把握，越敢吃", GREEN2),
]
for i, (t, r, d, c) in enumerate(rates):
    x = Inches(0.7 + i * 4.1)
    add_rect(s, x, Inches(2.6), Inches(3.85), Inches(1.7), LIGHT)
    add_rect(s, x, Inches(2.6), Inches(3.85), Inches(0.55), c)
    add_text(s, x + Inches(0.2), Inches(2.68), Inches(3.45), Inches(0.4), [[(t, 14, WHITE, True)]])
    add_text(s, x + Inches(0.2), Inches(3.25), Inches(3.45), Inches(0.5), [[(r, 15, c, True)]])
    add_text(s, x + Inches(0.2), Inches(3.75), Inches(3.45), Inches(0.55), [[(d, 11, GRAY, False)]])
add_text(s, Inches(0.7), Inches(4.6), Inches(11.9), Inches(1.5),
         [[("记住", 14, ACCENT, True)],
          [("成功率不是替代算卡，而是在「算卡之上」管理「这个卡数多可信」——卡照算，成功率管可信度。", 13, TEXT, True)]])

# ================= 10. 记账 / 存款 / 借款 =================
s = prs.slides.add_slide(BLANK)
header(s, "3 · 记账 / 存款 / 借款（核心货币：余卡）", 9)
add_rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(1.1), DARK)
add_text(s, Inches(1.0), Inches(1.85), Inches(11.3), Inches(0.8),
         [[("余卡 = 运动消耗卡路里 − 吃掉卡路里", 20, WHITE, True),
           ("（支付剩下来的额度，就是「存款」）", 14, ACCENT, True)]])
rules = [
    ("余卡为正 → 存款", "多运动少花，攒下额度", PRIMARY),
    ("透支 → 借款", "吃超了记一笔借款", RGBColor(0xC0, 0x3A, 0x2B)),
    ("借款影响后续标准", "欠了卡，食谱/升级标准收紧，吃得保守", ACCENT),
    ("保守低估兜底", "余卡按最低值展示，晒图标「至少余XX卡」", GREEN2),
]
for i, (t, d, c) in enumerate(rules):
    x = Inches(0.7 + i * 3.05)
    add_rect(s, x, Inches(3.1), Inches(2.85), Inches(2.0), LIGHT)
    add_rect(s, x, Inches(3.1), Inches(2.85), Inches(0.12), c)
    add_text(s, x + Inches(0.2), Inches(3.3), Inches(2.5), Inches(0.9), [[(t, 15, c, True)]])
    add_text(s, x + Inches(0.2), Inches(4.1), Inches(2.5), Inches(0.9), [[(d, 11, GRAY, False)]])
add_text(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.0),
         [[("为什么「敢吃还能瘦」不倒：", 14, ACCENT, True),
           ("余卡按保守值展示，宁可少算不吹，保住承诺的可信度。", 13, TEXT, False)]])

# ================= 11. 食谱升级 =================
s = prs.slides.add_slide(BLANK)
header(s, "4 · 食谱升级（出发动机）", 10)
add_text(s, Inches(0.7), Inches(1.6), Inches(11.9), Inches(1.6),
         [[("AI 根据画像设计「低卡但可入口」的食谱，并可定制到「爱吃但想吃更好的」。", 15, TEXT, False)],
          [("升级的依据就是算出的卡路里——没有卡数，就无法定义「这顿是否超预算、能否升级、该不该借款」。", 14, GRAY, False)],
          [("所以算卡是升级的依据，不可动摇。", 15, ACCENT, True)]])
add_rect(s, Inches(0.7), Inches(3.4), Inches(11.9), Inches(2.5), LIGHT)
add_text(s, Inches(1.0), Inches(3.6), Inches(11.3), Inches(0.5), [[("AI 个性化维度", 15, PRIMARY, True)]])
dims = ["体重 & 目标", "偏好口味", "额定运动比例", "少吃比例", "健康风险检测"]
for i, d in enumerate(dims):
    x = Inches(1.0 + i * 2.35)
    add_rect(s, x, Inches(4.15), Inches(2.15), Inches(0.75), WHITE)
    add_text(s, x, Inches(4.3), Inches(2.15), Inches(0.45), [[(d, 12, TEXT, False)]], align=PP_ALIGN.CENTER)
add_text(s, Inches(1.0), Inches(5.1), Inches(11.3), Inches(0.8),
         [[("健康兜底：", 13, RGBColor(0xC0, 0x3A, 0x2B), True),
           ("系统检测计划是否过于激进（热量缺口过大/运动超负荷），一旦危害健康即拦截并给出安全方案。", 12, GRAY, False)]])

# ================= 12. 排行榜 + 社区 =================
s = prs.slides.add_slide(BLANK)
header(s, "5 · 余卡排行榜 + 交流社区（社交引擎）", 11)
left = [
    ("余卡排行榜", "鼓励多走多跑、少花、存余卡"),
    ("分层防挫败", "按体重/基础/目标分层，避免逼卷大基数用户"),
    ("进步榜", "「超越昨天的自己」，弱化横向比较"),
    ("晒图主题", "核心主题「存余卡+成功吃上爱吃的=余卡健康」，而非谁吃得爽"),
    ("炫耀模板", "「走了3800米，2520大卡，换了这顿烤肉，余卡还剩180」"),
    ("点赞互动", "存卡多+吃上爱吃的人收获点赞，激发更多人存卡发内容"),
]
add_text(s, Inches(0.7), Inches(1.65), Inches(6.3), Inches(0.5), [[("社交引擎", 15, PRIMARY, True)]])
add_bullets(s, Inches(0.9), Inches(2.2), Inches(6.0), Inches(4.6), left, size=13, gap=11)
add_rect(s, Inches(7.3), Inches(1.7), Inches(5.4), Inches(4.9), DARK)
add_text(s, Inches(7.6), Inches(1.9), Inches(4.8), Inches(0.5), [[("内容 = 活广告", 16, ACCENT, True)]])
add_bullets(s, Inches(7.7), Inches(2.5), Inches(4.7), Inches(3.9),
            ["用户自发晒图传播，为产品拉新",
             "冷启动：先做排行榜（系统自动算、零内容也能跑）",
             "社区作为轻功能先挂上，用打卡奖励制造初始内容",
             "低门槛晒图：运动轨迹图（走/跑地图打卡，无羞耻、高满足，不依赖「吃了什么」）"],
            size=12, gap=12, color=WHITE, marker_color=ACCENT)

# ================= 13. 差异化 =================
s = prs.slides.add_slide(BLANK)
header(s, "差异化", 12)
diffs = [
    ("对纯计步/工具App", "不是记步，而是「距离+时间→精确运动消耗」，给出「能吃爱吃的额度」反馈"),
    ("对纯卡路里App", "不假装精确，用「成功率管理误差+保守低估」，做到「敢吃还能瘦」的承诺"),
    ("对食谱/健身博主", "不逼挨饿，把减肥过程灵活化、个性化、可兑现"),
    ("对地图大厂（运动导航）", "别人告诉你「这食物多少卡」，我们给你「运动挣来的额度」+「敢不敢吃」的把握"),
]
for i, (t, d) in enumerate(diffs):
    y = Inches(1.6) + Inches(i * 1.35)
    add_rect(s, Inches(0.7), y, Inches(3.2), Inches(1.15), PRIMARY)
    add_text(s, Inches(0.9), y + Inches(0.3), Inches(2.9), Inches(0.6), [[(t, 13, WHITE, True)]])
    add_rect(s, Inches(3.95), y, Inches(8.65), Inches(1.15), LIGHT)
    add_text(s, Inches(4.2), y + Inches(0.25), Inches(8.2), Inches(0.7), [[(d, 13, TEXT, False)]],
             anchor=MSO_ANCHOR.MIDDLE)

# ================= 14. 电梯陈述 =================
s = prs.slides.add_slide(BLANK)
header(s, "一页纸 · 电梯陈述", 13)
add_rect(s, Inches(0.7), Inches(1.7), Inches(11.9), Inches(4.3), DARK)
add_text(s, Inches(1.15), Inches(2.0), Inches(11.0), Inches(3.8),
         [[("我们做的是把减肥重构成『为爱吃的东西付出有意义努力』的 AI 系统。", 18, WHITE, True)],
          [("", 8, WHITE, False)],
          [("它以地图距离+时间精确算出运动消耗为地基，用拍照AI估卡路里+成功率管理误差管理吃进去的额度，把剩余额度变成可存可借的『余卡』——卡点提示风险、有余量放心吃。", 15, RGBColor(0xCF, 0xE0, 0xDB), False)],
          [("", 8, WHITE, False)],
          [("再加上余卡排行榜和社区晒图，让存卡变成被羡慕的成就和活广告。", 15, RGBColor(0xCF, 0xE0, 0xDB), False)],
          [("", 8, WHITE, False)],
          [("我们帮的是那些『不是饿、是馋、试过总放弃』的人——", 15, RGBColor(0xCF, 0xE0, 0xDB), False)],
          [("我们承诺的不是100%能瘦，而是『你每次敢不敢吃，给你一个看得见的把握』。", 17, ACCENT, True)]])

# ================= 15. 结尾 =================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, DARK)
add_rect(s, 0, 0, SW, Inches(0.35), ACCENT)
add_text(s, Inches(1.0), Inches(2.6), Inches(11.3), Inches(1.2),
         [[("把减肥，变成一件值得坚持的事。", 40, WHITE, True)]], align=PP_ALIGN.CENTER)
add_text(s, Inches(1.0), Inches(4.0), Inches(11.3), Inches(0.8),
         [[("不是饿 · 是馋 · 但这次你「有把握」", 18, ACCENT, True)]], align=PP_ALIGN.CENTER)

out = "AI减肥操作系统_产品全案.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
