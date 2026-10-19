import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Executive Theme Colors
    c_navy = RGBColor(11, 59, 111)          # #0b3b6f
    c_navy_dark = RGBColor(7, 32, 64)       # #072040
    c_navy_card = RGBColor(13, 48, 88)      # #0d3058
    c_teal = RGBColor(18, 128, 106)         # #12806a
    c_teal_light = RGBColor(127, 240, 207)  # #7ff0cf
    c_teal_dark = RGBColor(13, 106, 88)     # #0d6a58
    c_blue2 = RGBColor(29, 90, 153)         # #1d5a99
    c_gold = RGBColor(200, 147, 12)         # #c8930c
    c_gold_dark = RGBColor(138, 101, 0)     # #8a6500
    c_ink = RGBColor(30, 47, 66)            # #1e2f42
    c_muted = RGBColor(91, 107, 124)        # #5b6b7c
    c_line = RGBColor(221, 231, 241)        # #dde7f1
    c_soft = RGBColor(244, 248, 252)        # #f4f8fc
    c_mint = RGBColor(233, 246, 241)        # #e9f6f1
    c_sand = RGBColor(251, 245, 230)        # #fbf5e6
    c_white = RGBColor(255, 255, 255)
    c_bg_gray = RGBColor(246, 249, 252)     # #f6f9fc

    font_main = 'Segoe UI'

    def set_rtl(p):
        p.alignment = PP_ALIGN.RIGHT
        pPr = p._p.get_or_add_pPr()
        pPr.set('rtl', '1')

    def add_card_with_badge(slide, pos_x, pos_y, w, h, num_str, title_str, bg_color, border_color, badge_bg, badge_fg, title_color=c_ink, title_size=12.5):
        # Card Background
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, pos_x, pos_y, w, h)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.2)

        # Number Badge (Right side in RTL)
        badge_w = Inches(0.68)
        badge_h = Inches(0.42)
        pad = Inches(0.12)
        badge_x = pos_x + w - badge_w - pad
        badge_y = pos_y + (h - badge_h) / 2

        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, badge_x, badge_y, badge_w, badge_h)
        badge.fill.solid()
        badge.fill.fore_color.rgb = badge_bg
        badge.line.color.rgb = badge_bg
        b_tf = badge.text_frame
        b_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        b_tf.margin_left = b_tf.margin_top = b_tf.margin_right = b_tf.margin_bottom = 0
        bp = b_tf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = num_str
        br.font.name = font_main
        br.font.size = Pt(10.5)
        br.font.bold = True
        br.font.color.rgb = badge_fg

        # Title text box (takes remaining width on the left)
        tb_x = pos_x + pad
        tb_w = w - badge_w - (pad * 2.8)
        tb_y = pos_y
        tb_h = h

        tb = slide.shapes.add_textbox(tb_x, tb_y, tb_w, tb_h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p = tf.paragraphs[0]
        set_rtl(p)
        r = p.add_run()
        r.text = title_str
        r.font.name = font_main
        r.font.size = Pt(title_size)
        r.font.bold = True
        r.font.color.rgb = title_color

    def add_page_header(slide, page_num, sec_num_str, title_str, subtitle_str, theme_color):
        # Top gradient line
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = theme_color
        top_bar.line.fill.background()

        # Section badge (circle/square on the right)
        badge_size = Inches(0.55)
        badge_x = Inches(13.333 - 0.8 - 0.55)
        badge_y = Inches(0.3)
        b_shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, badge_x, badge_y, badge_size, badge_size)
        b_shp.fill.solid()
        b_shp.fill.fore_color.rgb = theme_color
        b_shp.line.fill.background()
        btf = b_shp.text_frame
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        br = bp.add_run()
        br.text = sec_num_str
        br.font.name = font_main
        br.font.size = Pt(18)
        br.font.bold = True
        br.font.color.rgb = c_white

        # Header Titles
        h_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(11.0), Inches(0.75))
        tf = h_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p0 = tf.paragraphs[0]
        set_rtl(p0)
        p0.space_after = Pt(1)
        r0 = p0.add_run()
        r0.text = f"Movia Bus  •  {subtitle_str}"
        r0.font.name = font_main
        r0.font.size = Pt(10.5)
        r0.font.bold = True
        r0.font.color.rgb = theme_color

        p1 = tf.add_paragraph()
        set_rtl(p1)
        r1 = p1.add_run()
        r1.text = title_str
        r1.font.name = font_main
        r1.font.size = Pt(21)
        r1.font.bold = True
        r1.font.color.rgb = c_navy

        # Footer
        f_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.08), Inches(11.733), Inches(0.32))
        f_tf = f_box.text_frame
        f_tf.margin_left = f_tf.margin_top = f_tf.margin_right = f_tf.margin_bottom = 0
        fp = f_tf.paragraphs[0]
        set_rtl(fp)
        fr = fp.add_run()
        fr.text = f"Movia Bus — نظام إدارة وتتبع حافلات نقل الطلاب  |  الصفحة {page_num}"
        fr.font.name = font_main
        fr.font.size = Pt(9.5)
        fr.font.color.rgb = c_muted

    # =========================================================================
    # SLIDE 1: COVER
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)

    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = c_navy_dark
    bg1.line.fill.background()

    top_accent = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    top_accent.fill.solid()
    top_accent.fill.fore_color.rgb = c_teal
    top_accent.line.fill.background()

    # Right Content Area: Brand + Title (Centered vertically on the right half)
    c_box = s1.shapes.add_textbox(Inches(5.7), Inches(1.8), Inches(6.8), Inches(4.2))
    c_tf = c_box.text_frame
    c_tf.word_wrap = True
    c_tf.margin_left = c_tf.margin_top = c_tf.margin_right = c_tf.margin_bottom = 0

    p_title = c_tf.paragraphs[0]
    set_rtl(p_title)
    p_title.space_after = Pt(4)
    r_title = p_title.add_run()
    r_title.text = "Movia Bus"
    r_title.font.name = font_main
    r_title.font.size = Pt(54)
    r_title.font.bold = True
    r_title.font.color.rgb = c_white

    p_sub = c_tf.add_paragraph()
    set_rtl(p_sub)
    p_sub.space_after = Pt(20)
    r_sub = p_sub.add_run()
    r_sub.text = "ماذا يقدم التطبيق؟"
    r_sub.font.name = font_main
    r_sub.font.size = Pt(32)
    r_sub.font.bold = True
    r_sub.font.color.rgb = c_teal_light

    p_desc = c_tf.add_paragraph()
    set_rtl(p_desc)
    r_desc = p_desc.add_run()
    r_desc.text = "نظام متكامل يربط ثلاثة أطراف في تجربة نقل جامعي حديثة ومريحة: الطالب يتابع ويحجز بسهولة، السائق يشغّل ويرى الطلب لايف، والإدارة تتحكم وتتابع بالأرقام والتقارير الحية."
    r_desc.font.name = font_main
    r_desc.font.size = Pt(14)
    r_desc.font.color.rgb = RGBColor(220, 235, 248)

    # Left Area: Team Members Card (Properly Structured)
    t_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.1), Inches(4.5), Inches(5.3))
    t_card.fill.solid()
    t_card.fill.fore_color.rgb = c_navy_card
    t_card.line.color.rgb = RGBColor(25, 75, 125)
    t_card.line.width = Pt(1.5)

    # Header box for Team Card
    th_box = s1.shapes.add_textbox(Inches(1.0), Inches(1.3), Inches(4.1), Inches(0.85))
    th_tf = th_box.text_frame
    th_tf.margin_left = th_tf.margin_top = th_tf.margin_right = th_tf.margin_bottom = 0

    thp0 = th_tf.paragraphs[0]
    thp0.alignment = PP_ALIGN.CENTER
    thp0.space_after = Pt(2)
    thr0 = thp0.add_run()
    thr0.text = "فريق العمل"
    thr0.font.name = font_main
    thr0.font.size = Pt(19)
    thr0.font.bold = True
    thr0.font.color.rgb = c_teal_light

    thp1 = th_tf.add_paragraph()
    thp1.alignment = PP_ALIGN.CENTER
    thr1 = thp1.add_run()
    thr1.text = "مشروع نظام إدارة وتتبع حافلات الطلاب"
    thr1.font.name = font_main
    thr1.font.size = Pt(11)
    thr1.font.color.rgb = RGBColor(180, 210, 235)

    # 6 Member Pills (2 columns x 3 rows)
    team_members = [
        "حسن ممدوح",
        "محمد شحاته",
        "أحمد عصام",
        "عبدالله جمال",
        "مولر باسم",
        "كيرلس سمير"
    ]

    m_w = Inches(1.9)
    m_h = Inches(0.72)
    m_gap_x = Inches(0.18)
    m_gap_y = Inches(0.18)
    m_start_x = Inches(1.06)
    m_start_y = Inches(2.35)

    for i, mem in enumerate(team_members):
        col = i // 3 # 0 for first 3 (Right column in RTL), 1 for next 3 (Left column in RTL)
        row = i % 3  # 0, 1, 2 (Top, Middle, Bottom)
        mx = m_start_x + (1 - col) * (m_w + m_gap_x)
        my = m_start_y + row * (m_h + m_gap_y)

        m_shp = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, my, m_w, m_h)
        m_shp.fill.solid()
        m_shp.fill.fore_color.rgb = RGBColor(18, 58, 102)
        m_shp.line.color.rgb = RGBColor(35, 95, 155)
        m_shp.line.width = Pt(1.2)
        mtf = m_shp.text_frame
        mtf.vertical_anchor = MSO_ANCHOR.MIDDLE
        mtf.margin_left = mtf.margin_right = Inches(0.1)
        mp = mtf.paragraphs[0]
        mp.alignment = PP_ALIGN.CENTER
        mr = mp.add_run()
        mr.text = mem
        mr.font.name = font_main
        mr.font.size = Pt(13.5)
        mr.font.bold = True
        mr.font.color.rgb = c_white

    # University Footer in Team Card
    uf_box = s1.shapes.add_textbox(Inches(1.0), Inches(5.75), Inches(4.1), Inches(0.45))
    uf_tf = uf_box.text_frame
    ufp = uf_tf.paragraphs[0]
    ufp.alignment = PP_ALIGN.CENTER
    ufr = ufp.add_run()
    ufr.text = "جامعة اللوتس — المنيا الجديدة"
    ufr.font.name = font_main
    ufr.font.size = Pt(12)
    ufr.font.bold = True
    ufr.font.color.rgb = c_teal_light

    # =========================================================================
    # SLIDE 2: STUDENT (PAGE 2) - Titles Only!
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    bg2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = c_bg_gray
    bg2.line.fill.background()

    add_page_header(s2, 2, "1", "ماذا يقدم التطبيق للطالب؟", "تجربة الطالب", c_teal)

    # Sub-banner A: Exclusive Features
    banner_a = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.15), Inches(11.733), Inches(0.36))
    banner_a.fill.solid()
    banner_a.fill.fore_color.rgb = c_teal
    banner_a.line.fill.background()
    ba_tf = banner_a.text_frame
    ba_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ba_tf.margin_right = Inches(0.2)
    ba_p = ba_tf.paragraphs[0]
    set_rtl(ba_p)
    ba_r = ba_p.add_run()
    ba_r.text = "نقاط التميز الحصرية — موجودة في Movia Bus ومميزة عن أي تطبيق نقل"
    ba_r.font.name = font_main
    ba_r.font.size = Pt(11)
    ba_r.font.bold = True
    ba_r.font.color.rgb = c_white

    # 6 Exclusive Cards (Titles matching requirements.html)
    exclusive_cards = [
        ("1.1", "زر Ready"),
        ("1.2", "رحلة أقصر — محطات فاضية بتتعدّي"),
        ("1.3", "مساعد AI يحلل جدولك الدراسي"),
        ("1.4", "مشاركة الرحلة مع الأهل"),
        ("1.5", "توقع الازدحام بمساعدة الـ AI"),
        ("1.6", "فاتك الموعد؟ بديل فوري"),
    ]

    c_w = Inches(3.75)
    c_h = Inches(0.68)
    gx = Inches(0.24)
    gy = Inches(0.14)
    sy_a = Inches(1.60)

    for i, (num, title) in enumerate(exclusive_cards):
        col = i % 3
        row = i // 3
        px = Inches(0.8) + (2 - col) * (c_w + gx)
        py = sy_a + row * (c_h + gy)
        add_card_with_badge(
            slide=s2,
            pos_x=px, pos_y=py, w=c_w, h=c_h,
            num_str=num, title_str=title,
            bg_color=c_mint,
            border_color=RGBColor(180, 222, 208),
            badge_bg=c_teal,
            badge_fg=c_white,
            title_color=c_ink,
            title_size=12.5
        )

    # Sub-banner B: Core Features
    banner_b = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.30), Inches(11.733), Inches(0.36))
    banner_b.fill.solid()
    banner_b.fill.fore_color.rgb = c_blue2
    banner_b.line.fill.background()
    bb_tf = banner_b.text_frame
    bb_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bb_tf.margin_right = Inches(0.2)
    bb_p = bb_tf.paragraphs[0]
    set_rtl(bb_p)
    bb_r = bb_p.add_run()
    bb_r.text = "الميزات الأساسية المشتركة — كل ما يحتاجه الطالب يومياً"
    bb_r.font.name = font_main
    bb_r.font.size = Pt(11)
    bb_r.font.bold = True
    bb_r.font.color.rgb = c_white

    # 9 Core Cards (Titles matching requirements.html)
    core_cards = [
        ("1.7", "خريطة حية للحافلة"),
        ("1.8", "وقت الوصول المتوقع"),
        ("1.9", "حجز المقاعد وقائمة الانتظار"),
        ("1.10", "بطاقة ركوب QR"),
        ("1.11", "إشعارات موجهة — مش إزعاج عام"),
        ("1.12", "زر طوارئ SOS"),
        ("1.13", "دفع الاشتراك من التطبيق"),
        ("1.14", "تقييم وبلاغات"),
        ("1.15", "بيانات الحافلة والسائق"),
    ]

    sy_b = Inches(3.76)
    for i, (num, title) in enumerate(core_cards):
        col = i % 3
        row = i // 3
        px = Inches(0.8) + (2 - col) * (c_w + gx)
        py = sy_b + row * (c_h + gy)
        add_card_with_badge(
            slide=s2,
            pos_x=px, pos_y=py, w=c_w, h=c_h,
            num_str=num, title_str=title,
            bg_color=c_white,
            border_color=c_line,
            badge_bg=c_blue2,
            badge_fg=c_white,
            title_color=c_ink,
            title_size=12.5
        )

    # Bottom summary callout
    bot_callout = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.36), Inches(11.733), Inches(0.52))
    bot_callout.fill.solid()
    bot_callout.fill.fore_color.rgb = c_soft
    bot_callout.line.color.rgb = c_line
    bc_tf = bot_callout.text_frame
    bc_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bc_tf.margin_right = Inches(0.2)
    bc_p = bc_tf.paragraphs[0]
    set_rtl(bc_p)
    bcr1 = bc_p.add_run()
    bcr1.text = "إجمالي 15 ميزة متكاملة: "
    bcr1.font.name = font_main
    bcr1.font.size = Pt(11)
    bcr1.font.bold = True
    bcr1.font.color.rgb = c_navy
    bcr2 = bc_p.add_run()
    bcr2.text = "6 ميزات تفوق حصرية لحل مشاكل الانتظار والتأخير + 9 ميزات أساسية تضمن أقصى درجات الراحة والأمان."
    bcr2.font.name = font_main
    bcr2.font.size = Pt(10.5)
    bcr2.font.color.rgb = c_muted

    # =========================================================================
    # SLIDE 3: DRIVER (PAGE 3) - Titles Only!
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    bg3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg3.fill.solid()
    bg3.fill.fore_color.rgb = c_bg_gray
    bg3.line.fill.background()

    add_page_header(s3, 3, "2", "ماذا يقدم التطبيق للسائق؟", "تجربة السائق", c_teal_dark)

    # Role Hero Banner
    driver_hero = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.18), Inches(11.733), Inches(0.82))
    driver_hero.fill.solid()
    driver_hero.fill.fore_color.rgb = c_teal_dark
    driver_hero.line.fill.background()
    dh_tf = driver_hero.text_frame
    dh_tf.margin_top = Inches(0.12)
    dh_tf.margin_right = Inches(0.25)

    dhp0 = dh_tf.paragraphs[0]
    set_rtl(dhp0)
    dhr0 = dhp0.add_run()
    dhr0.text = "شغل أبسط — بلا اتصالات وبلا تخمين"
    dhr0.font.name = font_main
    dhr0.font.size = Pt(14.5)
    dhr0.font.bold = True
    dhr0.font.color.rgb = c_white

    dhp1 = dh_tf.add_paragraph()
    set_rtl(dhp1)
    dhp1.space_before = Pt(2)
    dhr1 = dhp1.add_run()
    dhr1.text = "واجهة تشغيل فائقة السهولة: السائق يدير رحلته بزرين فقط، ويشاهد أماكن انتظار الطلاب على الخريطة دون مكالمات تشتته."
    dhr1.font.name = font_main
    dhr1.font.size = Pt(11)
    dhr1.font.color.rgb = RGBColor(220, 248, 238)

    # 9 Driver Cards (Titles matching requirements.html)
    driver_cards = [
        ("2.1", "زرين بس: «ابدأ الرحلة» و«أنهِها»"),
        ("2.2", "بثّ الموقع تلقائياً"),
        ("2.3", "حالة الرحلة بتتغير لوحدها"),
        ("2.4", "خريطة الطلب على محطاته"),
        ("2.5", "تخطي المحطات الفاضية"),
        ("2.6", "مسح QR للركوب + عداد حي"),
        ("2.7", "كشف ركوب يدوي احتياطي"),
        ("2.8", "رحلاته مجدولة وواضحة"),
        ("2.9", "مفيش مكالمات — فيه شات جوه التطبيق"),
    ]

    dc_w = Inches(3.75)
    dc_h = Inches(1.02)
    d_gx = Inches(0.24)
    d_gy = Inches(0.22)
    start_y_d = Inches(2.26)

    for i, (num, title) in enumerate(driver_cards):
        col = i % 3
        row = i // 3
        px = Inches(0.8) + (2 - col) * (dc_w + d_gx)
        py = start_y_d + row * (dc_h + d_gy)
        add_card_with_badge(
            slide=s3,
            pos_x=px, pos_y=py, w=dc_w, h=dc_h,
            num_str=num, title_str=title,
            bg_color=c_white,
            border_color=RGBColor(190, 226, 215),
            badge_bg=c_teal,
            badge_fg=c_white,
            title_color=c_ink,
            title_size=13.5
        )

    # Driver Callout at bottom
    d_callout = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.68))
    d_callout.fill.solid()
    d_callout.fill.fore_color.rgb = c_mint
    d_callout.line.color.rgb = RGBColor(170, 222, 205)
    dc_tf = d_callout.text_frame
    dc_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    dc_tf.margin_right = Inches(0.25)
    dcp = dc_tf.paragraphs[0]
    set_rtl(dcp)
    dcr1 = dcp.add_run()
    dcr1.text = "الخلاصة للسائق: "
    dcr1.font.name = font_main
    dcr1.font.size = Pt(12)
    dcr1.font.bold = True
    dcr1.font.color.rgb = c_teal

    dcr2 = dcp.add_run()
    dcr2.text = "أقل تشتيت أثناء القيادة، بدون مكالمات هاتفية ترن وراء المقود، ومعرفة دقيقة بعدد الركاب — رحلة أسرع وأهدى وأأمن."
    dcr2.font.name = font_main
    dcr2.font.size = Pt(11)
    dcr2.font.color.rgb = c_ink

    # =========================================================================
    # SLIDE 4: ADMIN (PAGE 4) - Titles Only!
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    bg4 = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg4.fill.solid()
    bg4.fill.fore_color.rgb = c_bg_gray
    bg4.line.fill.background()

    add_page_header(s4, 4, "3", "ماذا يقدم التطبيق للإدارة؟", "لوحة الإدارة والتحكم", c_gold_dark)

    # Role Hero Banner
    admin_hero = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.18), Inches(11.733), Inches(0.82))
    admin_hero.fill.solid()
    admin_hero.fill.fore_color.rgb = c_gold_dark
    admin_hero.line.fill.background()
    ah_tf = admin_hero.text_frame
    ah_tf.margin_top = Inches(0.12)
    ah_tf.margin_right = Inches(0.25)

    ahp0 = ah_tf.paragraphs[0]
    set_rtl(ahp0)
    ahr0 = ahp0.add_run()
    ahr0.text = "تحكم كامل — من غير رسايل يدوية"
    ahr0.font.name = font_main
    ahr0.font.size = Pt(14.5)
    ahr0.font.bold = True
    ahr0.font.color.rgb = c_white

    ahp1 = ah_tf.add_paragraph()
    set_rtl(ahp1)
    ahp1.space_before = Pt(2)
    ahr1 = ahp1.add_run()
    ahr1.text = "نشر جدول الرحلات بضغطتين، متابعة الأسطول لحظياً على الخريطة، وقرارات تشغيلية مبنية على أرقام وتقارير موثقة لأول مرة."
    ahr1.font.name = font_main
    ahr1.font.size = Pt(11)
    ahr1.font.color.rgb = RGBColor(255, 245, 220)

    # 10 Admin Cards (Titles matching requirements.html)
    admin_cards = [
        ("3.1", "نشر جدول اليوم مع مراجعة قبل التأكيد"),
        ("3.2", "إضافة وتعديل وإلغاء الرحلات"),
        ("3.3", "لوحة طلب مباشرة"),
        ("3.4", "متابعة حية للتشغيل"),
        ("3.5", "إدارة الاشتراكات والتحصيل"),
        ("3.6", "تقارير قابلة للتصدير PDF/Excel"),
        ("3.7", "متابعة البلاغات"),
        ("3.8", "حماية دخل الاشتراكات"),
        ("3.9", "خصوصية الطلاب محفوظة"),
    ]
    card_10 = ("3.10", "سياسات مثبتة للجميع")

    ac_w = Inches(3.75)
    ac_h = Inches(0.82)
    a_gx = Inches(0.24)
    a_gy = Inches(0.14)
    start_y_admin = Inches(2.26)

    # First 9 cards in 3x3 grid
    for i, (num, title) in enumerate(admin_cards):
        col = i % 3
        row = i // 3
        px = Inches(0.8) + (2 - col) * (ac_w + a_gx)
        py = start_y_admin + row * (ac_h + a_gy)
        add_card_with_badge(
            slide=s4,
            pos_x=px, pos_y=py, w=ac_w, h=ac_h,
            num_str=num, title_str=title,
            bg_color=c_white,
            border_color=RGBColor(230, 212, 175),
            badge_bg=c_gold_dark,
            badge_fg=c_white,
            title_color=c_ink,
            title_size=12.5
        )

    # Card 10 placed prominently centered across the middle
    c10_w = Inches(5.8)
    c10_x = Inches(0.8) + (Inches(11.733) - c10_w) / 2
    c10_y = start_y_admin + 3 * (ac_h + a_gy)
    add_card_with_badge(
        slide=s4,
        pos_x=c10_x, pos_y=c10_y, w=c10_w, h=ac_h,
        num_str=card_10[0], title_str=card_10[1],
        bg_color=c_sand,
        border_color=RGBColor(225, 200, 150),
        badge_bg=c_gold_dark,
        badge_fg=c_white,
        title_color=c_ink,
        title_size=13
    )

    # Admin Callout at bottom
    a_callout = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.68))
    a_callout.fill.solid()
    a_callout.fill.fore_color.rgb = c_sand
    a_callout.line.color.rgb = RGBColor(225, 205, 160)
    ac_tf = a_callout.text_frame
    ac_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ac_tf.margin_right = Inches(0.25)
    acp = ac_tf.paragraphs[0]
    set_rtl(acp)
    acr1 = acp.add_run()
    acr1.text = "الخلاصة للإدارة: "
    acr1.font.name = font_main
    acr1.font.size = Pt(12)
    acr1.font.bold = True
    acr1.font.color.rgb = c_gold_dark

    acr2 = acp.add_run()
    acr2.text = "من «كتابة يدوية ومعلومة مفقودة» إلى «نشر بضغطة، متابعة لايف، وقرارات مبنية على أرقام حقيقية» مع حماية دخل الاشتراكات."
    acr2.font.name = font_main
    acr2.font.size = Pt(11)
    acr2.font.color.rgb = c_ink

    # Save
    out_desktop = r"C:\Users\Lenovo\Desktop\bus-app-presentation.pptx"
    out_repo = r"C:\Users\Lenovo\Desktop\pdf-requirements\bus-app-presentation.pptx"
    prs.save(out_desktop)
    prs.save(out_repo)
    print("V3 presentation generated and saved to Desktop and Repo!")

if __name__ == '__main__':
    build_presentation()
