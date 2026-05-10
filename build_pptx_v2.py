"""
大樹藥局 × MediSpan × NEOV.AI — PowerPoint v2
Clean Professional Medical Design
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

# ── Color palette ──────────────────────────────────────────────────────────
GRN_DARK  = RGBColor(0x1B, 0x6B, 0x3A)
GRN_MED   = RGBColor(0x2E, 0x8B, 0x57)
GRN_LIGHT = RGBColor(0xE8, 0xF5, 0xEE)
GRN_ACC   = RGBColor(0x43, 0xA0, 0x47)
GRN_PALE  = RGBColor(0xF7, 0xFB, 0xF8)
BLU_DARK  = RGBColor(0x15, 0x65, 0xC0)
BLU_MED   = RGBColor(0x19, 0x76, 0xD2)
BLU_LIGHT = RGBColor(0xE3, 0xF0, 0xFF)
TEAL      = RGBColor(0x00, 0x79, 0x6B)
TEAL_LT   = RGBColor(0xE0, 0xF2, 0xF1)
ORANGE    = RGBColor(0xE6, 0x51, 0x00)
ORG_LIGHT = RGBColor(0xFF, 0xF3, 0xE0)
RED_DARK  = RGBColor(0xC6, 0x28, 0x28)
RED_LIGHT = RGBColor(0xFF, 0xEB, 0xEE)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
TXT_DARK  = RGBColor(0x1A, 0x2E, 0x1A)
TXT_BODY  = RGBColor(0x3D, 0x4D, 0x3D)
LGRAY     = RGBColor(0xF5, 0xF5, 0xF5)
MGRAY     = RGBColor(0xBD, 0xBD, 0xBD)
DGRAY     = RGBColor(0x61, 0x61, 0x61)
BORDER    = RGBColor(0xC8, 0xE6, 0xC9)
NAVY      = RGBColor(0x1A, 0x23, 0x7E)
A5D6A7    = RGBColor(0xA5, 0xD6, 0xA7)

# ── Slide dimensions ───────────────────────────────────────────────────────
W = Inches(13.33)
H = Inches(7.5)

prs = Presentation()
prs.slide_width = W
prs.slide_height = H
BLANK = prs.slide_layouts[6]


# ── Low-level helpers ──────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line_color=None, line_width=Pt(1)):
    shape = slide.shapes.add_shape(1, l, t, w, h)
    if fill is not None:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color is not None:
        shape.line.color.rgb = line_color
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape


def txbox(slide, text, l, t, w, h,
          size=14, bold=False, color=None, align=PP_ALIGN.LEFT, italic=False,
          wrap=True):
    if color is None:
        color = TXT_DARK
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tb


def txbox_multi(slide, lines, l, t, w, h,
                size=13, color=None, bold_first=False, line_space=1.0):
    """Add a textbox with multiple lines (list of strings or (text, bold, size, color) tuples)."""
    if color is None:
        color = TXT_BODY
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        if isinstance(item, str):
            txt, b, sz, col = item, False, size, color
        else:
            txt = item[0]
            b   = item[1] if len(item) > 1 else False
            sz  = item[2] if len(item) > 2 else size
            col = item[3] if len(item) > 3 else color
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        run = p.add_run()
        run.text = txt
        run.font.size = Pt(sz)
        run.font.bold = b
        run.font.color.rgb = col
    return tb


def std_header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, W, H, fill=WHITE)
    add_rect(slide, 0, 0, W, Inches(1.2), fill=GRN_DARK)
    txbox(slide, title, Inches(0.5), Inches(0.1), Inches(12.3), Inches(0.75),
          size=26, bold=True, color=WHITE)
    if subtitle:
        txbox(slide, subtitle, Inches(0.5), Inches(0.82), Inches(12), Inches(0.35),
              size=13, color=A5D6A7)
    add_rect(slide, 0, Inches(1.2), W, Inches(0.05), fill=GRN_ACC)
    add_rect(slide, 0, H - Inches(0.32), W, Inches(0.32), fill=GRN_DARK)
    txbox(slide, "大樹藥局 × MediSpan × NEOV.AI  ·  數據藥局提案  ·  2026",
          Inches(0.3), H - Inches(0.3), Inches(12), Inches(0.28),
          size=9, color=A5D6A7)


def kpi_box(slide, label, value, l, t, w, h, bg=GRN_DARK, fg=WHITE, sub=None):
    add_rect(slide, l, t, w, h, fill=bg, line_color=None)
    txbox(slide, value, l, t + Inches(0.12), w, Inches(0.65),
          size=32, bold=True, color=fg, align=PP_ALIGN.CENTER)
    txbox(slide, label, l, t + Inches(0.72), w, Inches(0.38),
          size=11, bold=False, color=fg, align=PP_ALIGN.CENTER)
    if sub:
        txbox(slide, sub, l, t + Inches(1.05), w, Inches(0.28),
              size=9, color=fg, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ══════════════════════════════════════════════════════════════════════════
def slide01(prs):
    sl = prs.slides.add_slide(BLANK)

    # Dark green background
    add_rect(sl, 0, 0, W, H, fill=GRN_DARK)

    # Decorative overlapping rectangles (abstract "tree" shapes)
    # Bottom-right large light rectangle
    sh = sl.shapes.add_shape(1, Inches(8.5), Inches(2.5), Inches(5.5), Inches(5.5))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0x23, 0x7A, 0x47)
    sh.line.fill.background()

    sh2 = sl.shapes.add_shape(1, Inches(9.8), Inches(1.2), Inches(4.2), Inches(4.2))
    sh2.fill.solid(); sh2.fill.fore_color.rgb = RGBColor(0x2D, 0x88, 0x55)
    sh2.line.fill.background()

    sh3 = sl.shapes.add_shape(1, Inches(10.8), Inches(3.8), Inches(2.8), Inches(4.0))
    sh3.fill.solid(); sh3.fill.fore_color.rgb = RGBColor(0x1A, 0x5C, 0x32)
    sh3.line.fill.background()

    sh4 = sl.shapes.add_shape(1, Inches(7.5), Inches(5.0), Inches(3.0), Inches(2.8))
    sh4.fill.solid(); sh4.fill.fore_color.rgb = RGBColor(0x20, 0x72, 0x42)
    sh4.line.fill.background()

    # Thin accent bar
    add_rect(sl, Inches(0.5), Inches(1.6), Inches(3.5), Inches(0.06), fill=GRN_ACC)

    # Title
    txbox(sl, "大樹藥局 × MediSpan × NEOV.AI",
          Inches(0.5), Inches(1.8), Inches(10), Inches(1.0),
          size=34, bold=True, color=WHITE)

    # Subtitle
    txbox(sl, "打造數據藥局 · 以藥物交互作用驅動補充品銷售",
          Inches(0.5), Inches(2.85), Inches(10), Inches(0.6),
          size=20, color=A5D6A7)

    # Divider
    add_rect(sl, Inches(0.5), Inches(3.6), Inches(5), Inches(0.04), fill=GRN_MED)

    # Tagline
    txbox(sl, "從合規工具到銷售引擎 — 數據驅動的下一代藥局服務模式",
          Inches(0.5), Inches(3.75), Inches(10), Inches(0.45),
          size=14, color=A5D6A7)

    # Bottom authors
    add_rect(sl, 0, H - Inches(0.95), W, Inches(0.95), fill=RGBColor(0x12, 0x4A, 0x28))
    txbox(sl, "Angela Hsu  /  Frank Kao  /  Andy Lam  |  2026年5月",
          Inches(0.5), H - Inches(0.85), Inches(9), Inches(0.4),
          size=14, color=A5D6A7, bold=True)
    txbox(sl, "NEOV.AI  |  數位健康解決方案",
          Inches(0.5), H - Inches(0.5), Inches(9), Inches(0.35),
          size=11, color=RGBColor(0x81, 0xC7, 0x84))


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 2 — AGENDA
# ══════════════════════════════════════════════════════════════════════════
def slide02(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "簡報大綱")

    items = [
        ("01", "台灣藥局面臨的結構性危機",
         "電商競爭、AI衝擊——藥局傳統模式正面臨雙重夾擊", GRN_DARK, GRN_LIGHT),
        ("02", "MediSpan 產品介紹",
         "醫院等級藥物情報資料庫：150+專家、每日更新、0.4秒偵測", GRN_MED, GRN_LIGHT),
        ("03", "實驗數據：一般AI vs MediSpan",
         "147張真實處方嚴格測試——68.7%情況下一般AI漏報或低估風險", BLU_DARK, BLU_LIGHT),
        ("04", "DDI 驅動補充品銷售方案 + 收益試算",
         "每次警示即銷售機會——100家門市月增 NT$60萬–315萬潛力", TEAL, TEAL_LT),
    ]

    top = Inches(1.45)
    row_h = Inches(1.3)
    gap = Inches(0.12)

    for i, (num, title, desc, accent, bg) in enumerate(items):
        y = top + i * (row_h + gap)
        add_rect(sl, Inches(0.4), y, W - Inches(0.8), row_h, fill=bg,
                 line_color=BORDER, line_width=Pt(1))
        # Number circle (filled rect approximation)
        add_rect(sl, Inches(0.55), y + Inches(0.3), Inches(0.6), Inches(0.6), fill=accent)
        txbox(sl, num, Inches(0.55), y + Inches(0.28), Inches(0.6), Inches(0.6),
              size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, title, Inches(1.3), y + Inches(0.12), Inches(9), Inches(0.48),
              size=18, bold=True, color=accent)
        txbox(sl, desc, Inches(1.3), y + Inches(0.6), Inches(11.3), Inches(0.55),
              size=12, color=TXT_BODY)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 3 — 台灣藥局的困境
# ══════════════════════════════════════════════════════════════════════════
def slide03(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "台灣藥局正面臨兩股夾擊")

    cx = W / 2
    top = Inches(1.38)
    box_w = Inches(5.8)
    box_h = Inches(3.5)
    gap = Inches(0.25)

    # Left box — orange/commerce threat
    lx = Inches(0.4)
    add_rect(sl, lx, top, box_w, box_h, fill=ORG_LIGHT, line_color=ORANGE, line_width=Pt(2))
    add_rect(sl, lx, top, box_w, Inches(0.52), fill=ORANGE)
    txbox(sl, "🏪  電商 + 藥妝店競爭",
          lx + Inches(0.15), top + Inches(0.06), box_w - Inches(0.3), Inches(0.42),
          size=16, bold=True, color=WHITE)

    bullets_l = [
        "• 成藥、補充品可在電商 / 全聯 / 藥妝店買到",
        "  價格更低、配送更快、24小時可購",
        "",
        "• 消費者不清楚藥局藥師的專業加值",
        "  「藥局只是貴一點的全聯」觀感普遍",
        "",
        "• 2023年台灣電商健康品市場年增 18%",
        "  藥局補充品毛利持續受壓",
    ]
    txbox_multi(sl, bullets_l, lx + Inches(0.2), top + Inches(0.65),
                box_w - Inches(0.35), Inches(2.7), size=12, color=TXT_DARK)

    # Right box — AI threat
    rx = lx + box_w + gap
    add_rect(sl, rx, top, box_w, box_h, fill=BLU_LIGHT, line_color=BLU_DARK, line_width=Pt(2))
    add_rect(sl, rx, top, box_w, Inches(0.52), fill=BLU_DARK)
    txbox(sl, "🤖  AI 普及化衝擊",
          rx + Inches(0.15), top + Inches(0.06), box_w - Inches(0.3), Inches(0.42),
          size=16, bold=True, color=WHITE)

    bullets_r = [
        "• 患者拿手機問 ChatGPT 用藥問題",
        "  「AI說可以」直接挑戰藥師專業判斷",
        "",
        "• 一般AI在DDI偵測的準確率僅31.3%",
        "  但患者不知道——他們相信AI",
        "",
        "• 藥師難以即時反駁或提供科學依據",
        "  缺乏標準化、可引用的資料庫支撐",
    ]
    txbox_multi(sl, bullets_r, rx + Inches(0.2), top + Inches(0.65),
                box_w - Inches(0.35), Inches(2.7), size=12, color=TXT_DARK)

    # Bottom result banner
    by = top + box_h + Inches(0.2)
    add_rect(sl, Inches(0.4), by, W - Inches(0.8), Inches(0.72), fill=RED_DARK)
    txbox(sl, "結果 →  藥局正在淪為「離家近的領藥通路 / 藥妝店」，專業價值被忽視、客單價持續下滑",
          Inches(0.65), by + Inches(0.08), W - Inches(1.3), Inches(0.55),
          size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 4 — MEDISPAN PRODUCT INTRO
# ══════════════════════════════════════════════════════════════════════════
def slide04(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "Medi-Span® 是什麼？",
               "Medi-Span® Clinical — Wolters Kluwer Health 旗下藥物情報資料庫")

    top = Inches(1.38)
    content_h = H - top - Inches(0.45)

    # Left column (60%)
    lw = Inches(7.6)
    add_rect(sl, Inches(0.4), top, lw, content_h, fill=GRN_PALE, line_color=BORDER)

    left_content = [
        ("核心功能", True, 15, GRN_DARK),
        ("", False, 5, TXT_BODY),
        ("▶  藥物交互作用（DDI）即時篩查", False, 13, TXT_DARK),
        ("    嵌入臨床流程，0.4秒 / 張處方完成檢核", False, 12, TXT_BODY),
        ("", False, 5, TXT_BODY),
        ("▶  三層警示顯示系統", False, 13, TXT_DARK),
        ("    Icon警示 → 短訊說明 → 完整臨床指引", False, 12, TXT_BODY),
        ("", False, 5, TXT_BODY),
        ("▶  標準化嚴重性分級", False, 13, TXT_DARK),
        ("    極高 / 重大 / 中度 / 輕微 — 四級分類", False, 12, TXT_BODY),
        ("", False, 5, TXT_BODY),
        ("▶  Command Center", False, 13, TXT_DARK),
        ("    機構可自訂警示設定、閾值與工作流程", False, 12, TXT_BODY),
        ("", False, 5, TXT_BODY),
        ("▶  API 整合", False, 13, TXT_DARK),
        ("    無縫嵌入藥局 POS / 調劑系統，不中斷工作流程", False, 12, TXT_BODY),
    ]
    txbox_multi(sl, left_content,
                Inches(0.65), top + Inches(0.2), lw - Inches(0.5), content_h - Inches(0.3),
                size=13, color=TXT_BODY)

    # Right column — 4 KPI boxes
    rw = Inches(4.3)
    rx = Inches(8.6)
    kpi_data = [
        ("150+", "全職藥學 / 醫學內容專家", GRN_DARK),
        ("70+",  "持有進階學位藥師",         GRN_MED),
        ("260+", "外部顧問專家",              BLU_DARK),
        ("每日", "醫學文獻監測更新",          TEAL),
    ]
    kpi_h = Inches(1.3)
    kpi_gap = Inches(0.15)
    for idx, (val, lbl, bg) in enumerate(kpi_data):
        ky = top + idx * (kpi_h + kpi_gap)
        add_rect(sl, rx, ky, rw, kpi_h, fill=bg)
        txbox(sl, val, rx, ky + Inches(0.06), rw, Inches(0.72),
              size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, lbl, rx, ky + Inches(0.76), rw, Inches(0.4),
              size=12, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 5 — MEDISPAN CONTENT QUALITY
# ══════════════════════════════════════════════════════════════════════════
def slide05(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "為什麼 Medi-Span 資料是醫院等級？")

    sections = [
        (GRN_DARK, "內容編輯卓越性", [
            "• 每日監測產業動態與主要醫學文獻（FDA、EMA、台灣衛福部）",
            "• 涵蓋 Medication Guides、REMS、藥品短缺資訊",
            "• 製藥商公告、臨床試驗結果、上市後安全報告",
            "• 所有內容均有版本控制與審核記錄，可追溯",
        ]),
        (BLU_DARK, "跨專科團隊審核", [
            "• 150+ 全職內容專家（含 70+ 進階學位藥師 PharmD / PhD）",
            "• 260+ 外部顧問（含 40+ 美國以外國際市場專家）",
            "• 與 UpToDate 7,500+ 醫師作者網絡相互連結驗證",
            "• 所有條目均經同儕審閱（Peer Review），非AI生成",
        ]),
        (TEAL, "即時臨床決策支援", [
            "• 嵌入式解決方案，不中斷藥師 / 醫師工作流程",
            "• 標準化嚴重性分級：極高 / 重大 / 中度 / 輕微",
            "• 每月 11,694+ 筆劑量調整更新",
            "• 0.4 秒 / 張處方完成全庫比對，支援高流量環境",
        ]),
    ]

    top = Inches(1.38)
    sec_h = Inches(1.72)
    gap = Inches(0.12)

    for i, (color, title, bullets) in enumerate(sections):
        y = top + i * (sec_h + gap)
        # color sidebar
        add_rect(sl, Inches(0.4), y, Inches(0.3), sec_h, fill=color)
        # main box
        add_rect(sl, Inches(0.7), y, W - Inches(1.1), sec_h, fill=GRN_PALE,
                 line_color=BORDER, line_width=Pt(1))
        # section title
        add_rect(sl, Inches(0.7), y, W - Inches(1.1), Inches(0.42), fill=color)
        txbox(sl, title, Inches(0.9), y + Inches(0.06), W - Inches(1.5), Inches(0.34),
              size=15, bold=True, color=WHITE)
        # bullets
        txbox_multi(sl, bullets,
                    Inches(1.0), y + Inches(0.5), W - Inches(1.5), sec_h - Inches(0.58),
                    size=12, color=TXT_DARK)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 6 — MEDISPAN IN 大樹藥局 (Live data)
# ══════════════════════════════════════════════════════════════════════════
def slide06(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "Medi-Span 已在大樹藥局運行",
               "這不是概念驗證，這是已發生的真實數據")

    # 4 KPI boxes
    top = Inches(1.38)
    kpi_data = [
        ("17,972 張", "已掃描\n處方總數",      GRN_DARK),
        ("131 張",    "觸發高風險警示\n(含147個交互作用對)", RED_DARK),
        ("0.4 秒",    "平均每張處方\n檢查時間", BLU_DARK),
        ("151.7 張/分", "系統\n處理速度",       TEAL),
    ]
    kpi_w = Inches(2.98)
    kpi_h = Inches(2.1)
    gap   = Inches(0.18)

    for i, (val, lbl, bg) in enumerate(kpi_data):
        kx = Inches(0.4) + i * (kpi_w + gap)
        add_rect(sl, kx, top, kpi_w, kpi_h, fill=bg)
        txbox(sl, val, kx, top + Inches(0.18), kpi_w, Inches(0.85),
              size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, lbl, kx, top + Inches(1.0), kpi_w, Inches(0.85),
              size=12, color=A5D6A7, align=PP_ALIGN.CENTER)

    # Insight box
    iy = top + kpi_h + Inches(0.25)
    add_rect(sl, Inches(0.4), iy, W - Inches(0.8), Inches(1.55), fill=GRN_LIGHT,
             line_color=GRN_MED, line_width=Pt(2))
    txbox(sl, "每掃描 137 張處方，就有 1 張觸發高風險警示",
          Inches(0.7), iy + Inches(0.15), W - Inches(1.4), Inches(0.48),
          size=17, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER)
    txbox(sl, "→  各門市每 1–2 天就有 1 次高信任度 DDI 對話機會",
          Inches(0.7), iy + Inches(0.6), W - Inches(1.4), Inches(0.4),
          size=14, color=TXT_DARK, align=PP_ALIGN.CENTER)

    # Quote
    qy = iy + Inches(1.6)
    add_rect(sl, Inches(0.4), qy, W - Inches(0.8), Inches(0.65), fill=RGBColor(0xFD, 0xF0, 0xE5))
    add_rect(sl, Inches(0.4), qy, Inches(0.08), Inches(0.65), fill=ORANGE)
    txbox(sl, "「這個機會，目前只用來做合規記錄。我們能做更多。」",
          Inches(0.65), qy + Inches(0.1), W - Inches(1.2), Inches(0.45),
          size=14, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 7 — ANDY'S EXPERIMENT INTRO
# ══════════════════════════════════════════════════════════════════════════
def slide07(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "實驗設計：一般 AI vs Medi-Span 精確度對比",
               "NEOV.AI（Andy Lam）主導 — 台灣真實臨床處方嚴格測試")

    top = Inches(1.38)

    # Methodology box
    mby = top + Inches(0.15)
    add_rect(sl, Inches(0.4), mby, W - Inches(0.8), Inches(2.5),
             fill=GRN_PALE, line_color=GRN_MED, line_width=Pt(2))
    add_rect(sl, Inches(0.4), mby, W - Inches(0.8), Inches(0.46), fill=GRN_DARK)
    txbox(sl, "實驗方法",
          Inches(0.65), mby + Inches(0.06), Inches(4), Inches(0.38),
          size=16, bold=True, color=WHITE)

    method_lines = [
        ("取樣：  147 張真實台灣處方（來自大樹藥局實際數據）", False, 13, TXT_DARK),
        ("", False, 6, TXT_BODY),
        ("測試：  同一批處方同步輸入 Medi-Span DDI 系統 + 市面主流生成式AI", False, 13, TXT_DARK),
        ("", False, 6, TXT_BODY),
        ("比對：  以 Medi-Span 為黃金標準，評估 AI 的偵測準確度與風險分級正確性", False, 13, TXT_DARK),
        ("", False, 6, TXT_BODY),
        ("對象：  全部為 Medi-Span 評定「極高風險 Very High」的藥物交互作用案例", False, 13, TXT_DARK),
    ]
    txbox_multi(sl, method_lines,
                Inches(0.7), mby + Inches(0.56), W - Inches(1.3), Inches(1.8),
                size=13)

    # 3 definition boxes
    dy = mby + Inches(2.75)
    defs = [
        (GRN_MED,   "V  符合",   "兩者均偵測到極高 / 重大風險\n→ 系統一致，患者安全"),
        (ORANGE,    "P  部分符合", "Medi-Span: 極高風險\n一般AI: 僅評輕微或中度\n→ 嚴重低估，產生誤判"),
        (RED_DARK,  "X  不符合",  "Medi-Span偵測到交互作用\n一般AI完全未發現\n→ 漏報，患者暴露於風險"),
    ]
    dw = Inches(3.95)
    dgap = Inches(0.3)
    for i, (bg, label, desc) in enumerate(defs):
        dx = Inches(0.4) + i * (dw + dgap)
        add_rect(sl, dx, dy, dw, Inches(1.85), fill=bg)
        txbox(sl, label, dx, dy + Inches(0.1), dw, Inches(0.48),
              size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, desc, dx + Inches(0.15), dy + Inches(0.6), dw - Inches(0.3), Inches(1.1),
              size=12, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 8 — ANDY'S DATA RESULTS (KEY SLIDE)
# ══════════════════════════════════════════════════════════════════════════
def slide08(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "實驗結果：一般 AI 的真實表現",
               "147 張處方，全部為 Medi-Span 評定「極高風險（Very High）」")

    top = Inches(1.38)
    col_w = Inches(3.9)
    col_h = Inches(3.4)
    gap   = Inches(0.28)

    results = [
        (RED_DARK,  "X  不符合", "79 筆", "53.7%",
         "一般AI 完全未偵測到",    RED_LIGHT),
        (ORANGE,    "P  部分符合", "22 筆", "15.0%",
         "Medi-Span: 極高風險\n一般AI: 輕微或中度", ORG_LIGHT),
        (RGBColor(0x2E, 0x7D, 0x32), "V  符合", "46 筆", "31.3%",
         "兩者均偵測到\n極高 / 重大風險", GRN_LIGHT),
    ]

    for i, (accent, label, count, pct, desc, bg) in enumerate(results):
        cx = Inches(0.4) + i * (col_w + gap)
        # Main box
        add_rect(sl, cx, top, col_w, col_h, fill=bg,
                 line_color=accent, line_width=Pt(3))
        # Header bar
        add_rect(sl, cx, top, col_w, Inches(0.5), fill=accent)
        txbox(sl, label, cx, top + Inches(0.06), col_w, Inches(0.42),
              size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Big number
        txbox(sl, count, cx, top + Inches(0.6), col_w, Inches(1.0),
              size=44, bold=True, color=accent, align=PP_ALIGN.CENTER)
        # Percentage
        txbox(sl, pct, cx, top + Inches(1.55), col_w, Inches(0.5),
              size=24, bold=True, color=accent, align=PP_ALIGN.CENTER)
        # Description
        txbox(sl, desc, cx + Inches(0.15), top + Inches(2.1), col_w - Inches(0.3), Inches(1.1),
              size=13, color=TXT_DARK, align=PP_ALIGN.CENTER)

    # Bold conclusion banner
    by = top + col_h + Inches(0.2)
    add_rect(sl, Inches(0.4), by, W - Inches(0.8), Inches(0.88), fill=GRN_DARK)
    txbox(sl, "結論：68.7% 的情況下，一般 AI 要麼完全漏報，要麼嚴重低估風險等級",
          Inches(0.65), by + Inches(0.13), W - Inches(1.3), Inches(0.62),
          size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Small note
    ny = by + Inches(0.95)
    txbox(sl, "＊不符合定義：Medi-Span 偵測到交互作用而一般AI分析為無。部分符合：Medi-Span 評定 Very High 而一般AI分析為 Minor 或 Moderate。",
          Inches(0.5), ny, W - Inches(1.0), Inches(0.35),
          size=9, color=DGRAY, italic=True)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 9 — ACTUAL CASES
# ══════════════════════════════════════════════════════════════════════════
def slide09(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "實際案例：這些風險，一般 AI 完全看不見",
               "以下均為 Medi-Span 評定「極高風險」、一般AI漏報或低估的真實案例")

    cases = [
        (RED_DARK, "CYMBALTA (Duloxetine) + Trazodone",
         "Medi-Span:  ⚠ 極高風險",
         "臨床風險：血清素症候群（可致命）\n意識混亂、肌躍、高燒、自律神經失調",
         "一般AI評定：中度，建議注意"),
        (RED_DARK, "Statin + Amlodipine",
         "Medi-Span:  ⚠ 極高風險",
         "臨床風險：Simvastatin 血中濃度上升\n橫紋肌溶解症風險（可致急性腎衰竭）",
         "一般AI評定：輕微，無特別警告"),
        (ORANGE, "NSAIDs + SSRIs（如布洛芬 + 抗憂鬱藥）",
         "Medi-Span:  ⚠ 重大風險",
         "臨床風險：腸胃道出血風險上升 3–15 倍\n尤其合併使用時期超過 2 週",
         "一般AI評定：輕微，無特別說明"),
        (ORANGE, "ARBs + Spironolactone（如 Losartan + 螺旋內酯）",
         "Medi-Span:  ⚠ 重大風險",
         "臨床風險：高血鉀症 + QT間期延長 + 心律不整\n老年患者尤其高風險",
         "一般AI評定：輕微或完全未提及"),
    ]

    top  = Inches(1.38)
    cw   = Inches(6.0)
    ch   = Inches(2.6)
    hgap = Inches(0.35)
    vgap = Inches(0.2)

    for i, (border_col, drug, ms, risk, ai_label) in enumerate(cases):
        row = i // 2
        col = i % 2
        cx  = Inches(0.4) + col * (cw + hgap)
        cy  = top + row * (ch + vgap)
        bg  = RED_LIGHT if border_col == RED_DARK else ORG_LIGHT
        add_rect(sl, cx, cy, cw, ch, fill=bg, line_color=border_col, line_width=Pt(2))
        # Drug name
        add_rect(sl, cx, cy, cw, Inches(0.44), fill=border_col)
        txbox(sl, drug, cx + Inches(0.12), cy + Inches(0.06), cw - Inches(0.2), Inches(0.36),
              size=13, bold=True, color=WHITE)
        # Medi-Span label
        txbox(sl, ms, cx + Inches(0.15), cy + Inches(0.52), cw - Inches(0.2), Inches(0.32),
              size=12, bold=True, color=border_col)
        # Clinical risk
        txbox(sl, risk, cx + Inches(0.15), cy + Inches(0.82), cw - Inches(0.2), Inches(1.0),
              size=12, color=TXT_DARK)
        # AI label
        add_rect(sl, cx + Inches(0.15), cy + ch - Inches(0.46), cw - Inches(0.3), Inches(0.32),
                 fill=LGRAY, line_color=MGRAY, line_width=Pt(1))
        txbox(sl, ai_label, cx + Inches(0.2), cy + ch - Inches(0.45), cw - Inches(0.4), Inches(0.3),
              size=11, color=DGRAY, italic=True)

    # Bottom quote
    qy = top + 2 * (ch + vgap) + Inches(0.1)
    add_rect(sl, Inches(0.4), qy, W - Inches(0.8), Inches(0.55), fill=GRN_LIGHT,
             line_color=GRN_MED, line_width=Pt(1))
    txbox(sl, "「每一個被漏掉的交互作用，都是一個潛在的病患傷害，也是一個流失的信任機會。」",
          Inches(0.7), qy + Inches(0.08), W - Inches(1.4), Inches(0.4),
          size=13, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER, italic=True)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 10 — COMPARISON TABLE
# ══════════════════════════════════════════════════════════════════════════
def slide10(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "兩種工具，兩種等級",
               "不是「AI好不好」的問題，而是「哪個等級適合醫療場域」")

    rows = [
        ("比較維度",          "一般生成式 AI  🏪",           "Medi-Span  🏥",          True),
        ("內容品質",          "網路資料彙整，無追溯",        "150+ 全職藥學專家審核",   False),
        ("審查機制",          "無同儕審閱",                  "70+進階學位藥師+260+顧問", False),
        ("更新頻率",          "靜態訓練資料",                "每月 11,694+ 筆更新",     False),
        ("風險分級",          "模糊描述",                    "標準化四級分類",           False),
        ("147處方測試",       "68.7% 漏報或低報",            "零遺漏",                  False),
        ("工作流整合",        "無法嵌入系統",                "API整合+Command Center",  False),
        ("法律依據",          "若出事，藥師無據可循",        "有據可查的臨床決策工具",   False),
    ]

    top    = Inches(1.38)
    col_x  = [Inches(0.4), Inches(3.1), Inches(8.15)]
    col_w  = [Inches(2.65), Inches(5.0), Inches(5.0)]
    row_h  = Inches(0.5)

    for r, (dim, ai, ms, is_hdr) in enumerate(rows):
        ry = top + r * row_h
        bg_dim = GRN_DARK   if is_hdr else (LGRAY if r % 2 == 0 else WHITE)
        bg_ai  = GRN_DARK   if is_hdr else (RED_LIGHT if r % 2 == 0 else RGBColor(0xFF, 0xF0, 0xEE))
        bg_ms  = GRN_DARK   if is_hdr else (GRN_LIGHT if r % 2 == 0 else RGBColor(0xF0, 0xF9, 0xF2))
        fg     = WHITE if is_hdr else TXT_DARK
        sz     = 13 if is_hdr else 12
        bld    = is_hdr

        add_rect(sl, col_x[0], ry, col_w[0], row_h, fill=bg_dim,
                 line_color=BORDER, line_width=Pt(0.5))
        add_rect(sl, col_x[1], ry, col_w[1], row_h, fill=bg_ai,
                 line_color=BORDER, line_width=Pt(0.5))
        add_rect(sl, col_x[2], ry, col_w[2], row_h, fill=bg_ms,
                 line_color=BORDER, line_width=Pt(0.5))

        txbox(sl, dim, col_x[0] + Inches(0.1), ry + Inches(0.1), col_w[0] - Inches(0.15),
              row_h - Inches(0.08), size=sz, bold=bld, color=fg)
        txbox(sl, ai,  col_x[1] + Inches(0.1), ry + Inches(0.1), col_w[1] - Inches(0.15),
              row_h - Inches(0.08), size=sz, bold=bld,
              color=fg if is_hdr else RED_DARK)
        txbox(sl, ms,  col_x[2] + Inches(0.1), ry + Inches(0.1), col_w[2] - Inches(0.15),
              row_h - Inches(0.08), size=sz, bold=bld,
              color=fg if is_hdr else GRN_DARK)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 11 — DDI AS SALES ENGINE
# ══════════════════════════════════════════════════════════════════════════
def slide11(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "DDI 警示 → 銷售引擎",
               "每一個高風險警示，都是一次高信任度的補充品推薦機會")

    top = Inches(1.38)

    # Flow diagram
    steps = [
        ("MediSpan\n偵測", GRN_DARK),
        ("0.4秒\n觸發警示", GRN_MED),
        ("藥師\n開啟對話", BLU_DARK),
        ("科學依據\n推薦", TEAL),
        ("補充品銷售\nNT$450–1,800", ORANGE),
    ]
    sw = Inches(2.2)
    sh = Inches(1.4)
    sy = top + Inches(0.2)
    aw = Inches(0.45)
    total_w = len(steps) * sw + (len(steps) - 1) * aw
    sx_start = (W - total_w) / 2

    for i, (label, bg) in enumerate(steps):
        bx = sx_start + i * (sw + aw)
        add_rect(sl, bx, sy, sw, sh, fill=bg)
        txbox(sl, label, bx, sy + Inches(0.25), sw, sh - Inches(0.3),
              size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        if i < len(steps) - 1:
            ax = bx + sw
            add_rect(sl, ax, sy + sh / 2 - Inches(0.08), aw, Inches(0.16), fill=GRN_ACC)
            txbox(sl, "▶", ax, sy + sh / 2 - Inches(0.16), aw, Inches(0.35),
                  size=12, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER)

    # 3 columns — why this works
    wy = sy + sh + Inches(0.3)
    why = [
        (GRN_DARK, "病患最有感的時刻",
         "藥師主動說明用藥風險時，患者信任度最高\n→ 最佳推薦時機，不是廣告，是專業"),
        (BLU_DARK, "科學依據支撐",
         "不是推銷話術，是基於臨床數據的專業建議\n→ Medi-Span 提供可引用的科學依據"),
        (TEAL, "完全合規",
         "不改處方、不涉及診斷，屬藥師執業範疇\n→ 非藥物療法推薦，零法律風險"),
    ]
    ww = Inches(3.9)
    wgap = Inches(0.35)
    wx_start = Inches(0.55)

    for i, (accent, title, desc) in enumerate(why):
        wbx = wx_start + i * (ww + wgap)
        add_rect(sl, wbx, wy, ww, Inches(2.2), fill=GRN_PALE, line_color=accent, line_width=Pt(2))
        add_rect(sl, wbx, wy, ww, Inches(0.42), fill=accent)
        txbox(sl, title, wbx + Inches(0.1), wy + Inches(0.05), ww - Inches(0.2), Inches(0.36),
              size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        txbox(sl, desc, wbx + Inches(0.15), wy + Inches(0.52), ww - Inches(0.3), Inches(1.55),
              size=12, color=TXT_DARK, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDES 12–16 — 5 Sales Scenarios
# ══════════════════════════════════════════════════════════════════════════
def sales_scenario_slide(prs, num, title, subtitle, alert_text,
                          script_lines, products, price_range):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, f"情境 {num}：{title}", subtitle)

    top = Inches(1.38)

    # Scenario number circle
    add_rect(sl, Inches(0.4), top + Inches(0.1), Inches(0.65), Inches(0.65), fill=GRN_LIGHT,
             line_color=GRN_MED, line_width=Pt(2))
    txbox(sl, str(num), Inches(0.4), top + Inches(0.08), Inches(0.65), Inches(0.62),
          size=22, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER)

    # Alert box
    ay = top + Inches(0.05)
    add_rect(sl, Inches(1.2), ay, W - Inches(1.6), Inches(0.8), fill=RED_LIGHT,
             line_color=ORANGE, line_width=Pt(2))
    txbox(sl, "⚠  " + alert_text, Inches(1.35), ay + Inches(0.1), W - Inches(1.8), Inches(0.6),
          size=13, bold=True, color=RED_DARK)

    # Script box
    sy_box = ay + Inches(0.95)
    add_rect(sl, Inches(0.4), sy_box, W - Inches(0.8), Inches(1.9), fill=WHITE,
             line_color=GRN_MED, line_width=Pt(1))
    add_rect(sl, Inches(0.4), sy_box, Inches(0.12), Inches(1.9), fill=GRN_DARK)
    txbox(sl, "藥師話術", Inches(0.65), sy_box + Inches(0.08), Inches(3), Inches(0.3),
          size=12, bold=True, color=GRN_DARK)
    y_off = sy_box + Inches(0.4)
    for line in script_lines:
        txbox(sl, line, Inches(0.65), y_off, W - Inches(1.2), Inches(0.3),
              size=12, color=TXT_DARK)
        y_off += Inches(0.35)

    # Products
    py_box = sy_box + Inches(2.05)
    pw = Inches(2.8)
    ph = Inches(1.2)
    pgap = Inches(0.25)
    for i, (pname, pdesc) in enumerate(products):
        px = Inches(0.4) + i * (pw + pgap)
        add_rect(sl, px, py_box, pw, ph, fill=WHITE, line_color=BORDER, line_width=Pt(1))
        add_rect(sl, px, py_box, Inches(0.25), ph, fill=GRN_MED)
        txbox(sl, pname, px + Inches(0.35), py_box + Inches(0.1), pw - Inches(0.45), Inches(0.4),
              size=13, bold=True, color=GRN_DARK)
        txbox(sl, pdesc, px + Inches(0.35), py_box + Inches(0.5), pw - Inches(0.45), Inches(0.6),
              size=11, color=TXT_BODY)

    # Price badge
    bx = Inches(0.4) + len(products) * (pw + pgap)
    if bx + Inches(2.5) > W - Inches(0.4):
        bx = W - Inches(2.9)
    add_rect(sl, bx, py_box + Inches(0.25), Inches(2.4), Inches(0.7), fill=GRN_DARK)
    txbox(sl, price_range, bx, py_box + Inches(0.28), Inches(2.4), Inches(0.55),
          size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def slide12(prs):
    sales_scenario_slide(
        prs, 1,
        "Statin → CoQ10 + Omega-3",
        "降膽固醇藥物 × 肌肉保護補充品",
        "Medi-Span 偵測：Statin 可能耗竭體內 CoQ10，增加肌肉疼痛風險",
        [
            "「您這張處方含有 Statin（降血脂藥），研究顯示長期使用",
            " 可能降低體內 CoQ10 濃度，導致肌肉疲勞或疼痛。」",
            "「我們有醫院等級建議的補充組合，幫助維持肌肉功能。」",
        ],
        [
            ("CoQ10 輔酶", "100–200mg/天\n肌肉能量支持"),
            ("Omega-3 魚油", "EPA+DHA\n抗發炎保護"),
            ("維生素 D3", "骨骼 + 免疫支援\n配合 Statin"),
        ],
        "NT$ 450 – 800 / 月"
    )


def slide13(prs):
    sales_scenario_slide(
        prs, 2,
        "精神科用藥（QT延長）→ 鎂 + Omega-3",
        "抗憂鬱 / 抗精神病藥 × 心臟節律保護",
        "Medi-Span 偵測：多種精神科藥物可延長 QT 間期，與電解質失衡協同風險",
        [
            "「您使用的精神科藥物有一個需要注意的地方——」",
            "「它可能影響心臟節律，搭配鎂和 Omega-3 有助穩定心電傳導。」",
            "「這是我們根據藥物數據庫特別建議的組合，不是一般保健品。」",
        ],
        [
            ("鎂 Magnesium", "Glycinate形式\n心臟 + 神經支持"),
            ("Omega-3 魚油", "DHA心臟保護\n抗炎減QT風險"),
            ("褪黑激素", "改善睡眠品質\n精神科患者適用"),
        ],
        "NT$ 600 – 1,200 / 月"
    )


def slide14(prs):
    sales_scenario_slide(
        prs, 3,
        "NSAIDs + SSRIs → 胃腸道保護組合",
        "止痛藥 + 抗憂鬱藥 × 消化道出血風險",
        "Medi-Span 偵測：NSAIDs + SSRIs 合用，胃腸道出血風險上升 3–15 倍",
        [
            "「您同時使用止痛藥和抗憂鬱藥，這個組合有一個特別要注意的地方——」",
            "「研究顯示可能增加胃黏膜刺激，甚至出血風險。」",
            "「我建議加上益生菌和鋅，幫助保護胃腸道黏膜。」",
        ],
        [
            ("益生菌", "Lactobacillus族\n修復腸道黏膜"),
            ("鋅 Zinc", "胃黏膜修復\n免疫支持"),
            ("薑黃素", "天然抗炎\n減少胃刺激"),
        ],
        "NT$ 400 – 700 / 月"
    )


def slide15(prs):
    sales_scenario_slide(
        prs, 4,
        "心血管高風險 → 心臟三寶",
        "高血壓 / 心臟用藥 × 全方位心血管保護",
        "Medi-Span 偵測：ARB + 利尿劑組合，電解質流失 + 腎功能監測必要",
        [
            "「您的處方是心血管保護用藥組合，效果很好。」",
            "「長期服用可能影響體內鎂和鉀的平衡，建議補充。」",
            "「這三種補充品是心臟科建議搭配的黃金組合。」",
        ],
        [
            ("CoQ10", "心肌能量\n老年患者必備"),
            ("鎂 Magnesium", "血壓調節\n防心律不整"),
            ("Omega-3", "降三酸甘油酯\nAHA認可"),
        ],
        "NT$ 1,200 – 1,800 / 月"
    )


def slide16(prs):
    sales_scenario_slide(
        prs, 5,
        "鎮定安眠藥物 → 自然助眠",
        "Benzodiazepine / Z-drug × 減依賴自然轉銜",
        "Medi-Span 偵測：長期使用安眠藥有依賴風險，認知功能可能受影響",
        [
            "「您目前使用的安眠藥很有效，但長期使用要注意依賴問題。」",
            "「我們有幾種天然助眠成分，可以在醫師指導下逐步配合減量。」",
            "「褪黑激素 + 鎂的組合，在臨床上有助改善睡眠品質。」",
        ],
        [
            ("褪黑激素 3mg", "調節生理時鐘\n非成癮性"),
            ("鎂 Magnesium", "肌肉放鬆\n睡眠品質"),
            ("L-茶胺酸", "放鬆無嗜睡\n焦慮緩解"),
        ],
        "NT$ 500 – 900 / 月"
    )


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 17 — REVENUE MODEL
# ══════════════════════════════════════════════════════════════════════════
def slide17(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "收益試算",
               "以大樹藥局 100 家門市計算")

    top = Inches(1.38)
    cw  = Inches(5.6)
    ch  = Inches(4.1)
    gap = Inches(0.45)
    lx  = Inches(0.9)
    rx  = lx + cw + gap

    models = [
        (lx, GRN_DARK, GRN_LIGHT, "保守模型", [
            ("每店每日 DDI 對話",  "1 次"),
            ("補充品轉換率",       "40%"),
            ("平均加購金額",       "NT$500"),
            ("每店 / 月",          "NT$6,000"),
        ], "100店 / 月：NT$600,000"),
        (rx, BLU_DARK, BLU_LIGHT, "積極模型", [
            ("每店每日 DDI 對話",  "3 次"),
            ("補充品轉換率",       "50%"),
            ("平均加購金額",       "NT$700"),
            ("每店 / 月",          "NT$31,500"),
        ], "100店 / 月：NT$3,150,000"),
    ]

    for bx, accent, bg, label, items, total in models:
        add_rect(sl, bx, top, cw, ch, fill=bg, line_color=accent, line_width=Pt(2))
        add_rect(sl, bx, top, cw, Inches(0.5), fill=accent)
        txbox(sl, label, bx, top + Inches(0.07), cw, Inches(0.4),
              size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for j, (k, v) in enumerate(items):
            iy = top + Inches(0.65) + j * Inches(0.58)
            bg2 = WHITE if j % 2 == 0 else bg
            add_rect(sl, bx + Inches(0.2), iy, cw - Inches(0.4), Inches(0.48), fill=bg2,
                     line_color=BORDER, line_width=Pt(0.5))
            txbox(sl, k, bx + Inches(0.35), iy + Inches(0.08), Inches(2.8), Inches(0.35),
                  size=13, color=TXT_DARK)
            txbox(sl, v, bx + Inches(3.1), iy + Inches(0.08), Inches(2.1), Inches(0.35),
                  size=13, bold=True, color=accent, align=PP_ALIGN.RIGHT)
        # Total box
        ty = top + Inches(0.65) + 4 * Inches(0.58) + Inches(0.1)
        add_rect(sl, bx + Inches(0.2), ty, cw - Inches(0.4), Inches(0.55), fill=accent)
        txbox(sl, total, bx + Inches(0.2), ty + Inches(0.08), cw - Inches(0.4), Inches(0.42),
              size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # LTV note
    ly = top + ch + Inches(0.2)
    add_rect(sl, Inches(0.9), ly, W - Inches(1.8), Inches(0.8), fill=GRN_LIGHT,
             line_color=GRN_MED, line_width=Pt(1))
    txbox(sl, "LTV 倍增效應：建立科學推薦信任後，補充品客戶留存率提升 30–50%",
          Inches(1.1), ly + Inches(0.1), W - Inches(2.2), Inches(0.32),
          size=13, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER)
    txbox(sl, "→ 1位每月NT$700補充品客戶 × 24個月 = LTV NT$16,800 / 人",
          Inches(1.1), ly + Inches(0.42), W - Inches(2.2), Inches(0.3),
          size=12, color=TXT_DARK, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 18 — 3-PHASE IMPLEMENTATION
# ══════════════════════════════════════════════════════════════════════════
def slide18(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "三階段導入計畫")

    phases = [
        (GRN_DARK, "Phase 1", "1–3 個月", "數據奠基",
         ["Medi-Span DDI 系統全面整合",
          "藥師話術培訓 + 補充品知識庫",
          "POS 系統補充品推薦模組",
          "5–10家門市試點"],
         "試點完成 + 基礎銷售轉換 ≥20%"),
        (TEAL, "Phase 2", "4–6 個月", "規模化推廣",
         ["全台 100 家門市部署",
          "DDI→補充品推薦流程標準化",
          "月度績效儀表板",
          "補充品庫存智能優化"],
         "平均每店月增 NT$6,000+"),
        (BLU_DARK, "Phase 3", "7–12 個月", "Expert AI 上線",
         ["Medi-Span Expert AI 整合",
          "個人化推薦演算法",
          "患者用藥歷史分析",
          "競爭護城河建立"],
         "先發優勢鎖定 3–5 年"),
    ]

    top = Inches(1.38)
    pw  = Inches(3.9)
    ph  = Inches(4.65)
    gap = Inches(0.32)

    for i, (accent, phase, period, goal, items, kpi) in enumerate(phases):
        px = Inches(0.4) + i * (pw + gap)
        add_rect(sl, px, top, pw, ph, fill=GRN_PALE, line_color=accent, line_width=Pt(2))
        # Phase header
        add_rect(sl, px, top, pw, Inches(0.52), fill=accent)
        txbox(sl, phase, px, top + Inches(0.06), pw, Inches(0.4),
              size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Period
        add_rect(sl, px + Inches(0.5), top + Inches(0.6), pw - Inches(1.0), Inches(0.38),
                 fill=accent, line_color=None)
        txbox(sl, period, px, top + Inches(0.62), pw, Inches(0.34),
              size=13, color=WHITE, align=PP_ALIGN.CENTER)
        # Goal
        txbox(sl, goal, px, top + Inches(1.1), pw, Inches(0.38),
              size=16, bold=True, color=accent, align=PP_ALIGN.CENTER)
        # Items
        for j, item in enumerate(items):
            iy = top + Inches(1.58) + j * Inches(0.5)
            add_rect(sl, px + Inches(0.2), iy, pw - Inches(0.4), Inches(0.42), fill=WHITE,
                     line_color=BORDER, line_width=Pt(0.5))
            txbox(sl, "• " + item, px + Inches(0.35), iy + Inches(0.06),
                  pw - Inches(0.55), Inches(0.32), size=12, color=TXT_DARK)
        # KPI
        ky = top + ph - Inches(0.62)
        add_rect(sl, px + Inches(0.2), ky, pw - Inches(0.4), Inches(0.5), fill=accent)
        txbox(sl, kpi, px, ky + Inches(0.07), pw, Inches(0.38),
              size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 19 — EXPERT AI
# ══════════════════════════════════════════════════════════════════════════
def slide19(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "Medi-Span Expert AI — 2026年最新技術",
               "人工智慧 × 藥物情報 × 銷售引擎的三合一平台")

    top = Inches(1.38)

    # 5-step flow
    steps = [
        ("掃描處方", GRN_DARK),
        ("DDI偵測", GRN_MED),
        ("AI配對\n補充品", BLU_DARK),
        ("話術提示\n推送", TEAL),
        ("銷售 +\n學習回饋", ORANGE),
    ]
    sw = Inches(2.22)
    sh = Inches(1.35)
    sy = top + Inches(0.15)
    aw = Inches(0.3)
    total_flow = len(steps) * sw + (len(steps) - 1) * aw
    sx_start = (W - total_flow) / 2

    for i, (label, bg) in enumerate(steps):
        bx = sx_start + i * (sw + aw)
        add_rect(sl, bx, sy, sw, sh, fill=bg)
        txbox(sl, label, bx, sy + Inches(0.28), sw, sh - Inches(0.28),
              size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        if i < len(steps) - 1:
            ax = bx + sw
            txbox(sl, "▶", ax, sy + sh / 2 - Inches(0.18), aw, Inches(0.38),
                  size=13, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER)

    # 4 moat items (2×2)
    moats = [
        (GRN_DARK, "數據深度",  "147張真實處方 + 大樹持續累積\n→ 台灣本土DDI知識庫"),
        (BLU_DARK, "整合深度",  "POS + 調劑系統 + 話術AI\n→ 工作流完全嵌入"),
        (TEAL,     "時間優勢",  "先部署者建立數據飛輪\n→ 後進者難以複製"),
        (ORANGE,   "品牌護城河", "「數據藥局」品牌定位\n→ 患者信任不可量化"),
    ]
    mw = Inches(5.8)
    mh = Inches(1.25)
    mgap_x = Inches(0.3)
    mgap_y = Inches(0.18)
    my_start = sy + sh + Inches(0.3)

    for i, (accent, mtitle, mdesc) in enumerate(moats):
        row = i // 2
        col = i % 2
        mx  = Inches(0.4) + col * (mw + mgap_x)
        my  = my_start + row * (mh + mgap_y)
        add_rect(sl, mx, my, mw, mh, fill=WHITE, line_color=accent, line_width=Pt(2))
        add_rect(sl, mx, my, mw, Inches(0.06), fill=accent)
        txbox(sl, mtitle, mx + Inches(0.15), my + Inches(0.12), Inches(2), Inches(0.38),
              size=15, bold=True, color=accent)
        txbox(sl, mdesc, mx + Inches(0.15), my + Inches(0.5), mw - Inches(0.3), Inches(0.65),
              size=12, color=TXT_DARK)

    # Bottom statement
    by = my_start + 2 * (mh + mgap_y) + Inches(0.12)
    add_rect(sl, Inches(0.4), by, W - Inches(0.8), Inches(0.55), fill=GRN_DARK)
    txbox(sl, "保守估計：為大樹藥局建立 3–5 年先發競爭優勢，進入者壁壘持續加深",
          Inches(0.6), by + Inches(0.1), W - Inches(1.2), Inches(0.38),
          size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════
# SLIDE 20 — CALL TO ACTION
# ══════════════════════════════════════════════════════════════════════════
def slide20(prs):
    sl = prs.slides.add_slide(BLANK)
    std_header(sl, "下一步", "從決策到行動——兩種命運的選擇")

    top = Inches(1.38)
    cw  = Inches(5.9)
    ch  = Inches(1.65)
    gap = Inches(0.35)

    # Two fates
    lx = Inches(0.4)
    rx = lx + cw + gap

    # 命運一 — gray box
    add_rect(sl, lx, top, cw, ch, fill=LGRAY, line_color=MGRAY, line_width=Pt(1))
    txbox(sl, "命運一（現狀繼續）", lx + Inches(0.15), top + Inches(0.1),
          cw - Inches(0.3), Inches(0.35), size=14, bold=True, color=DGRAY)
    fate1 = [
        "• DDI數據只用於合規記錄，未轉換價值",
        "• 每月 100+ 次對話機會白白流失",
        "• 補充品銷售繼續依賴促銷折扣",
        "• 藥師專業無法被量化，信任難以積累",
    ]
    txbox_multi(sl, fate1, lx + Inches(0.15), top + Inches(0.5),
                cw - Inches(0.3), ch - Inches(0.55), size=12, color=DGRAY)

    # 命運二 — green box
    add_rect(sl, rx, top, cw, ch, fill=GRN_LIGHT, line_color=GRN_DARK, line_width=Pt(2))
    txbox(sl, "命運二（數據藥局）", rx + Inches(0.15), top + Inches(0.1),
          cw - Inches(0.3), Inches(0.35), size=14, bold=True, color=GRN_DARK)
    fate2 = [
        "• 每次警示 = 高信任度銷售對話",
        "• 月增 NT$60萬–315萬，低邊際成本",
        "• 藥師話術標準化，可複製可培訓",
        "• 品牌定位：值得信賴的醫療服務專家",
    ]
    txbox_multi(sl, fate2, rx + Inches(0.15), top + Inches(0.5),
                cw - Inches(0.3), ch - Inches(0.55), size=12, color=GRN_DARK)

    # Action table
    actions = [
        ("Week 1–2",  "確認 MediSpan API 整合規格 + 補充品品項清單"),
        ("Week 3–4",  "藥師話術培訓（首批5–10家門市）"),
        ("Month 2",   "試點數據收集：DDI對話次數 + 補充品轉換率"),
        ("Month 3",   "數據檢視會議 + 全台展開決策"),
        ("Month 4+",  "Expert AI 功能規劃 + 品牌數據藥局啟動"),
    ]

    at = top + ch + Inches(0.2)
    ah = Inches(0.46)
    col_w1 = Inches(2.0)
    col_w2 = W - Inches(0.8) - col_w1

    # Table header
    add_rect(sl, Inches(0.4), at, col_w1, ah, fill=GRN_DARK)
    add_rect(sl, Inches(0.4) + col_w1, at, col_w2, ah, fill=GRN_DARK)
    txbox(sl, "時程", Inches(0.55), at + Inches(0.08), col_w1 - Inches(0.15), Inches(0.3),
          size=13, bold=True, color=WHITE)
    txbox(sl, "行動項目", Inches(0.55) + col_w1, at + Inches(0.08), col_w2 - Inches(0.15),
          Inches(0.3), size=13, bold=True, color=WHITE)

    for i, (period, action) in enumerate(actions):
        ry = at + (i + 1) * ah
        bg2 = GRN_LIGHT if i % 2 == 0 else WHITE
        add_rect(sl, Inches(0.4), ry, col_w1, ah, fill=bg2, line_color=BORDER, line_width=Pt(0.5))
        add_rect(sl, Inches(0.4) + col_w1, ry, col_w2, ah, fill=bg2, line_color=BORDER, line_width=Pt(0.5))
        txbox(sl, period, Inches(0.55), ry + Inches(0.08), col_w1 - Inches(0.15), Inches(0.3),
              size=12, bold=True, color=GRN_DARK)
        txbox(sl, action, Inches(0.55) + col_w1, ry + Inches(0.08), col_w2 - Inches(0.15),
              Inches(0.3), size=12, color=TXT_DARK)

    # Closing quote
    qy = at + 6 * ah + Inches(0.12)
    add_rect(sl, Inches(0.4), qy, W - Inches(0.8), Inches(0.6), fill=GRN_LIGHT,
             line_color=GRN_MED, line_width=Pt(1))
    txbox(sl, "「從領藥的地方，到值得信賴的醫療服務專家。」",
          Inches(0.7), qy + Inches(0.1), W - Inches(1.4), Inches(0.42),
          size=16, bold=True, color=GRN_DARK, align=PP_ALIGN.CENTER, italic=True)


# ══════════════════════════════════════════════════════════════════════════
# BUILD ALL SLIDES
# ══════════════════════════════════════════════════════════════════════════
print("Building slides...")
slide01(prs)
print("  Slide 1 done")
slide02(prs)
print("  Slide 2 done")
slide03(prs)
print("  Slide 3 done")
slide04(prs)
print("  Slide 4 done")
slide05(prs)
print("  Slide 5 done")
slide06(prs)
print("  Slide 6 done")
slide07(prs)
print("  Slide 7 done")
slide08(prs)
print("  Slide 8 done")
slide09(prs)
print("  Slide 9 done")
slide10(prs)
print("  Slide 10 done")
slide11(prs)
print("  Slide 11 done")
slide12(prs)
print("  Slide 12 done")
slide13(prs)
print("  Slide 13 done")
slide14(prs)
print("  Slide 14 done")
slide15(prs)
print("  Slide 15 done")
slide16(prs)
print("  Slide 16 done")
slide17(prs)
print("  Slide 17 done")
slide18(prs)
print("  Slide 18 done")
slide19(prs)
print("  Slide 19 done")
slide20(prs)
print("  Slide 20 done")

out = "/home/user/adgo/daashu-medispan-v2.pptx"
prs.save(out)
print(f"\nSaved: {out}")
print(f"Slides: {len(prs.slides)}")
