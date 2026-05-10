from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Brand colours ──────────────────────────────────────────────
NAVY    = RGBColor(0x1A, 0x2E, 0x4A)   # deep navy – title bg
BLUE    = RGBColor(0x00, 0x6E, 0xB6)   # Wolters Kluwer blue
TEAL    = RGBColor(0x00, 0xA8, 0x9C)   # accent teal
ORANGE  = RGBColor(0xF0, 0x7D, 0x00)   # warning / highlight
RED     = RGBColor(0xC0, 0x39, 0x2B)   # danger
GREEN   = RGBColor(0x27, 0xAE, 0x60)   # positive
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY   = RGBColor(0xF4, 0xF6, 0xF8)   # slide bg
DGRAY   = RGBColor(0x4A, 0x4A, 0x4A)   # body text

W = Inches(13.33)   # widescreen 16:9
H = Inches(7.5)

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

BLANK = prs.slide_layouts[6]   # completely blank


# ── Helpers ────────────────────────────────────────────────────

def add_rect(slide, l, t, w, h, fill=None, line=None):
    shape = slide.shapes.add_shape(1, l, t, w, h)   # MSO_SHAPE_TYPE.RECTANGLE=1
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    else:
        shape.line.fill.background()
    return shape


def txbox(slide, text, l, t, w, h,
          size=18, bold=False, color=DGRAY, align=PP_ALIGN.LEFT,
          italic=False, wrap=True):
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


def slide_bg(slide, color=LGRAY):
    add_rect(slide, 0, 0, W, H, fill=color)


def title_bar(slide, title, subtitle=None):
    """Top navy bar with white title text."""
    bar_h = Inches(1.35)
    add_rect(slide, 0, 0, W, bar_h, fill=NAVY)
    txbox(slide, title, Inches(0.5), Inches(0.15), Inches(11), Inches(0.8),
          size=28, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        txbox(slide, subtitle, Inches(0.5), Inches(0.9), Inches(11), Inches(0.4),
              size=14, color=RGBColor(0xA8, 0xC8, 0xE8), align=PP_ALIGN.LEFT)


def footer(slide, text="大樹藥局 × MediSpan × NEOV.AI  |  數據藥局提案  |  2026"):
    add_rect(slide, 0, H - Inches(0.35), W, Inches(0.35), fill=NAVY)
    txbox(slide, text, Inches(0.3), H - Inches(0.32), Inches(12), Inches(0.28),
          size=9, color=RGBColor(0xA0, 0xB8, 0xD0), align=PP_ALIGN.LEFT)


def kpi_box(slide, label, value, l, t, w=Inches(2.8), h=Inches(1.4),
            bg=BLUE, val_size=32):
    add_rect(slide, l, t, w, h, fill=bg)
    txbox(slide, value, l + Inches(0.15), t + Inches(0.1), w - Inches(0.3), Inches(0.7),
          size=val_size, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txbox(slide, label, l + Inches(0.1), t + Inches(0.85), w - Inches(0.2), Inches(0.45),
          size=11, color=RGBColor(0xD0, 0xE8, 0xFF), align=PP_ALIGN.CENTER)


def bullet_box(slide, items, l, t, w, h, title=None,
               bg=WHITE, title_color=BLUE, bullet_color=DGRAY,
               title_size=14, bullet_size=13, gap=Inches(0.06)):
    add_rect(slide, l, t, w, h, fill=bg, line=RGBColor(0xCC, 0xCC, 0xCC))
    cy = t + Inches(0.15)
    if title:
        txbox(slide, title, l + Inches(0.15), cy, w - Inches(0.3), Inches(0.35),
              size=title_size, bold=True, color=title_color)
        cy += Inches(0.38)
    for item in items:
        txbox(slide, f"• {item}", l + Inches(0.2), cy, w - Inches(0.35), Inches(0.32),
              size=bullet_size, color=bullet_color)
        cy += Inches(0.32) + gap


# ══════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, W, H, fill=NAVY)
# accent stripe
add_rect(s, 0, Inches(3.5), Inches(0.12), Inches(2.2), fill=TEAL)
add_rect(s, Inches(0.12), Inches(3.5), Inches(0.06), Inches(2.2), fill=ORANGE)

txbox(s, "大樹藥局 × MediSpan × NEOV.AI",
      Inches(0.6), Inches(1.6), Inches(12), Inches(0.7),
      size=22, color=RGBColor(0xA8, 0xD0, 0xF0), bold=False)
txbox(s, "打造台灣第一家「數據藥局」",
      Inches(0.6), Inches(2.2), Inches(12), Inches(1.0),
      size=40, bold=True, color=WHITE)
txbox(s, "以藥物交互作用為核心，建立可複製的補充品銷售引擎",
      Inches(0.6), Inches(3.25), Inches(11), Inches(0.55),
      size=18, color=RGBColor(0xB8, 0xD8, 0xF0))

add_rect(s, Inches(0.6), Inches(4.0), Inches(11.5), Inches(0.03), fill=TEAL)

txbox(s, "Angela Hsu  |  Wolters Kluwer MediSpan",
      Inches(0.6), Inches(4.2), Inches(5.5), Inches(0.4), size=13,
      color=RGBColor(0x80, 0xA8, 0xC8))
txbox(s, "Frank Kao  |  Insight Software",
      Inches(0.6), Inches(4.6), Inches(5.5), Inches(0.4), size=13,
      color=RGBColor(0x80, 0xA8, 0xC8))
txbox(s, "Andy Lam  |  NEOV.AI",
      Inches(0.6), Inches(5.0), Inches(5.5), Inches(0.4), size=13,
      color=RGBColor(0x80, 0xA8, 0xC8))
txbox(s, "2026年5月", Inches(0.6), Inches(5.5), Inches(4), Inches(0.4),
      size=13, color=RGBColor(0x60, 0x90, 0xB8))


# ══════════════════════════════════════════════════════════════
# SLIDE 2 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "執行摘要", "核心命題：每一個 DDI 警示，都是一次高信任度的補充品銷售機會")
footer(s)

# Hook quote
add_rect(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(0.9), fill=TEAL)
txbox(s, "在 17,972 張處方中，有 131 個關鍵時刻——每一個時刻，都是藥師在病患最需要時提供精準建議的機會。",
      Inches(0.7), Inches(1.55), Inches(12.0), Inches(0.8),
      size=15, bold=True, color=WHITE)

# 3 outcome boxes
for i, (icon, title, desc, bg) in enumerate([
    ("01", "病患信任升級", "從「離家近的領藥點」→ 值得信賴的醫療夥伴\n回訪率預估提升 30–50%", BLUE),
    ("02", "補充品客單提升", "五大 DDI 場景話術\n每次推薦 NT$450–1,800 加購", TEAL),
    ("03", "可複製連鎖標準", "標準化流程，從試點門市\n擴展至全台所有據點", RGBColor(0x5B, 0x6A, 0xB0)),
]):
    x = Inches(0.5 + i * 4.15)
    add_rect(s, x, Inches(2.6), Inches(3.9), Inches(2.3), fill=bg)
    txbox(s, icon, x + Inches(0.2), Inches(2.7), Inches(0.6), Inches(0.5),
          size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF, ))
    txbox(s, title, x + Inches(0.2), Inches(3.1), Inches(3.5), Inches(0.45),
          size=16, bold=True, color=WHITE)
    txbox(s, desc, x + Inches(0.2), Inches(3.55), Inches(3.5), Inches(0.9),
          size=12, color=RGBColor(0xD0, 0xEC, 0xFF))

# Revenue highlight
add_rect(s, Inches(0.5), Inches(5.1), Inches(12.3), Inches(0.85), fill=ORANGE)
txbox(s, "預估收益規模：保守模型 NT$60萬/月　|　積極模型 NT$315萬/月　（以100家門市計算）",
      Inches(0.7), Inches(5.2), Inches(12.0), Inches(0.6),
      size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 3 — 菜市場 vs 精品店
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "菜市場 vs. 精品店：一個決定命運的選擇",
          "台灣藥局正在被便利商店化 + AI普及化兩股力量夾擊")
footer(s)

# Left panel – 菜市場
add_rect(s, Inches(0.4), Inches(1.5), Inches(5.9), Inches(5.6), fill=RGBColor(0xFF, 0xF3, 0xE0))
txbox(s, "菜市場", Inches(0.5), Inches(1.6), Inches(5.7), Inches(0.55),
      size=24, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
txbox(s, "一般生成式 AI（ChatGPT 等）",
      Inches(0.5), Inches(2.15), Inches(5.7), Inches(0.38),
      size=13, color=DGRAY, align=PP_ALIGN.CENTER)
add_rect(s, Inches(0.4), Inches(2.52), Inches(5.9), Inches(0.03), fill=ORANGE)

rows_left = [
    "網路資料彙整，無法追溯來源",
    "無同儕審閱，無標準流程",
    "靜態訓練資料，不持續更新",
    "53.7% 完全漏報高風險交互作用",
    "15.0% 嚴重低估風險等級",
    "人人都有，無法建立護城河",
    "藥師被 AI 取代",
]
cy = Inches(2.65)
for r in rows_left:
    txbox(s, f"✗  {r}", Inches(0.6), cy, Inches(5.5), Inches(0.37),
          size=12, color=RED)
    cy += Inches(0.4)

# Right panel – 精品店
add_rect(s, Inches(7.0), Inches(1.5), Inches(5.9), Inches(5.6), fill=RGBColor(0xE8, 0xF5, 0xFD))
txbox(s, "精品店", Inches(7.1), Inches(1.6), Inches(5.7), Inches(0.55),
      size=24, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
txbox(s, "MediSpan 藥物情報資料庫",
      Inches(7.1), Inches(2.15), Inches(5.7), Inches(0.38),
      size=13, color=DGRAY, align=PP_ALIGN.CENTER)
add_rect(s, Inches(7.0), Inches(2.52), Inches(5.9), Inches(0.03), fill=BLUE)

rows_right = [
    "150+ 全職藥學/醫學專家編寫",
    "70+ 進階學位藥師審核",
    "每月 11,694+ 筆更新；每日文獻監測",
    "高風險警示零遺漏（147處方測試）",
    "標準化嚴重性分級",
    "機構授權 + 自訂設定，不可複製",
    "藥師被 AI 武裝，成不可替代的專家",
]
cy = Inches(2.65)
for r in rows_right:
    txbox(s, f"✓  {r}", Inches(7.2), cy, Inches(5.5), Inches(0.37),
          size=12, color=GREEN)
    cy += Inches(0.4)

# VS badge
add_rect(s, Inches(6.07), Inches(3.9), Inches(0.85), Inches(0.85), fill=NAVY)
txbox(s, "VS", Inches(6.07), Inches(3.95), Inches(0.85), Inches(0.75),
      size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 4 — Andy's 147-prescription data
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "147 張處方的嚴格對比實驗",
          "NEOV.AI（Andy Lam）主導 — 以真實台灣處方同步測試 MediSpan vs 一般AI")
footer(s)

# 3 big KPI boxes
kpi_box(s, "完全漏報（AI完全未發現）", "53.7%\n79 筆",
        Inches(0.5), Inches(1.55), w=Inches(3.9), h=Inches(2.0), bg=RED, val_size=28)
kpi_box(s, "嚴重低估（Very High → Minor）", "15.0%\n22 筆",
        Inches(4.7), Inches(1.55), w=Inches(3.9), h=Inches(2.0), bg=ORANGE, val_size=28)
kpi_box(s, "一致正確（兩者均偵測）", "31.3%\n46 筆",
        Inches(8.9), Inches(1.55), w=Inches(3.9), h=Inches(2.0), bg=GREEN, val_size=28)

# Banner
add_rect(s, Inches(0.5), Inches(3.75), Inches(12.3), Inches(0.6), fill=NAVY)
txbox(s, "結論：68.7% 的情況下，一般 AI 要麼完全遺漏風險，要麼嚴重低估風險等級",
      Inches(0.7), Inches(3.82), Inches(12.0), Inches(0.45),
      size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Case examples
cases = [
    ("Statin + Amlodipine",   "AI：輕微", "MediSpan：極高風險（橫紋肌溶解）"),
    ("CYMBALTA + Trazodone",  "AI：中度", "MediSpan：極高風險（血清素症候群，可致命）"),
    ("NSAIDs + SSRIs",        "AI：輕微", "MediSpan：重大風險（腸胃道出血 ×3–15）"),
    ("ARBs + Spironolactone", "AI：未提及", "MediSpan：重大風險（高血鉀 + 心律不整）"),
]
for i, (drug, ai, ms) in enumerate(cases):
    x = Inches(0.5 + (i % 2) * 6.45)
    y = Inches(4.55 + (i // 2) * 1.1)
    add_rect(s, x, y, Inches(6.1), Inches(0.95), fill=WHITE,
             line=RGBColor(0xCC, 0xCC, 0xCC))
    txbox(s, drug, x + Inches(0.15), y + Inches(0.05), Inches(5.8), Inches(0.35),
          size=13, bold=True, color=NAVY)
    txbox(s, ai, x + Inches(0.15), y + Inches(0.42), Inches(2.5), Inches(0.3),
          size=11, color=RED)
    txbox(s, ms, x + Inches(2.9), y + Inches(0.42), Inches(3.1), Inches(0.3),
          size=11, color=GREEN)


# ══════════════════════════════════════════════════════════════
# SLIDE 5 — 大樹藥局 live data
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "大樹藥局實戰成果——數據已在你們手中",
          "MediSpan DDI 引擎已上線運行，現在是變現的時候")
footer(s)

kpis = [
    ("掃描處方總數", "17,972 張", BLUE),
    ("高風險警示處方", "131 張", RED),
    ("交互作用觸發率", "0.73%", ORANGE),
    ("平均檢查時間", "0.4 秒/張", TEAL),
]
for i, (lbl, val, bg) in enumerate(kpis):
    kpi_box(s, lbl, val, Inches(0.5 + i * 3.15), Inches(1.55),
            w=Inches(2.9), h=Inches(1.6), bg=bg, val_size=26)

# Insight box
add_rect(s, Inches(0.5), Inches(3.35), Inches(12.3), Inches(1.0), fill=RGBColor(0xE8, 0xF5, 0xFD))
txbox(s, "💡  每掃描 137 張處方，就有 1 張觸發高風險警示\n"
         "   → 各門市每 1–2 天就有 1 次高信任度 DDI 對話機會",
      Inches(0.7), Inches(3.4), Inches(12.0), Inches(0.85),
      size=15, bold=True, color=NAVY)

# Flow arrow
add_rect(s, Inches(0.5), Inches(4.55), Inches(12.3), Inches(0.65), fill=NAVY)
txbox(s, "這個機會，目前只用來做合規記錄。現在，我們要把它變成銷售引擎。",
      Inches(0.7), Inches(4.62), Inches(12.0), Inches(0.5),
      size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# 3 actions
actions = [
    ("STEP 1", "DDI 警示觸發", "0.4 秒 MediSpan 偵測"),
    ("STEP 2", "藥師話術啟動", "科學依據 + 情境對話"),
    ("STEP 3", "補充品推薦", "NT$450–1,800 加購"),
]
for i, (step, title, desc) in enumerate(actions):
    x = Inches(0.8 + i * 4.2)
    add_rect(s, x, Inches(5.35), Inches(3.7), Inches(1.55),
             fill=RGBColor(0xF0, 0xF8, 0xFF), line=BLUE)
    txbox(s, step, x + Inches(0.15), Inches(5.4), Inches(1.2), Inches(0.35),
          size=11, bold=True, color=BLUE)
    txbox(s, title, x + Inches(0.15), Inches(5.75), Inches(3.4), Inches(0.4),
          size=15, bold=True, color=NAVY)
    txbox(s, desc, x + Inches(0.15), Inches(6.15), Inches(3.4), Inches(0.35),
          size=12, color=DGRAY)


# ══════════════════════════════════════════════════════════════
# SLIDES 6–10 — Five Sales Scenarios
# ══════════════════════════════════════════════════════════════
scenarios = [
    {
        "no": "場景一",
        "drug": "Statin 用藥",
        "trigger": "Statin + Amlodipine / 葡萄柚汁",
        "alert": "Simvastatin 濃度上升 → 橫紋肌溶解風險\nStatin 抑制 HMG-CoA 還原酶，同時影響 CoQ10 合成",
        "script": "「您服用的降血脂藥長期使用下，會影響體內輔酶 Q10 的合成。"
                  "Q10 是心肌細胞能量代謝的輔因子，補充 CoQ10 可有效預防肌肉疼痛副作用。"
                  "搭配 Omega-3 魚油對三酸甘油脂的控制也有額外幫助。」",
        "products": ["CoQ10 100mg  — 肌肉保護、心肌能量", "Omega-3 魚油 1,000mg  — 輔助降三酸甘油脂"],
        "price": "NT$450–800",
        "bg": RGBColor(0xE3, 0xF2, 0xFD),
        "accent": BLUE,
    },
    {
        "no": "場景二",
        "drug": "精神科用藥",
        "trigger": "Quetiapine + Escitalopram / Flupentixol（QT 延長警示）",
        "alert": "多種精神科藥物協同延長心臟 QT 間期\n增加心律不整風險；常見副作用：睡眠品質下降",
        "script": "「您的用藥組合對心臟電氣傳導有影響，建議補充鎂——天然心臟節律穩定劑。"
                  "Omega-3 DHA 對心臟膜健康有額外支持。"
                  "若有入睡困難，褪黑激素比增加鎮靜藥劑量更溫和。」",
        "products": ["甘胺酸鎂 300mg  — 心律穩定、情緒", "Omega-3 魚油  — 心臟電位保護",
                     "褪黑激素 0.5–1mg  — 改善入睡", "L-茶胺酸  — 放鬆、減少焦慮"],
        "price": "NT$600–1,200",
        "bg": RGBColor(0xF3, 0xE5, 0xF5),
        "accent": RGBColor(0x7B, 0x1F, 0xA2),
    },
    {
        "no": "場景三",
        "drug": "NSAIDs + SSRIs",
        "trigger": "消炎止痛藥 + 抗憂鬱藥",
        "alert": "SSRIs 影響血小板凝集 + NSAIDs 損傷胃黏膜\n合用腸胃道出血風險上升 3–15 倍",
        "script": "「這兩種藥合用會輕微增加腸胃道出血風險。"
                  "建議用藥期間補充益生菌維持腸道菌相；Omega-3 天然抗炎；"
                  "維生素 C 有助黏膜修復。注意大便顏色，如變黑立即告知醫師。」",
        "products": ["益生菌複合配方  — 腸道黏膜保護", "Omega-3 魚油  — 抗炎、減少黏膜刺激",
                     "維生素 C 500mg  — 黏膜修復", "鎂 200mg  — 減少腸道平滑肌痙攣"],
        "price": "NT$400–700",
        "bg": RGBColor(0xE8, 0xF5, 0xE9),
        "accent": GREEN,
    },
    {
        "no": "場景四",
        "drug": "心血管高風險「心臟三寶」",
        "trigger": "ARBs + Spironolactone / 多重 QT 延長藥物",
        "alert": "高血鉀症風險 + 心臟 QT 間期延長\n需要整體的心臟微量元素支持",
        "script": "「心臟照護已有很好的藥物基礎，我們建議三個強化方向：\n"
                  "鎂——天然鈣離子拮抗劑，穩定心律；\n"
                  "CoQ10——心肌核心能量物質；\n"
                  "Omega-3——大量研究支持的心血管整體保護。\n我們稱為心臟三寶。」",
        "products": ["CoQ10 200mg  — 心肌能量", "甘胺酸鎂 300mg  — 心律穩定",
                     "Omega-3 魚油 2,000mg  — 心血管整體保護"],
        "price": "NT$1,200–1,800",
        "bg": RGBColor(0xFF, 0xF3, 0xE0),
        "accent": ORANGE,
    },
    {
        "no": "場景五",
        "drug": "鎮定安眠藥物",
        "trigger": "BZD + 抗憂鬱藥（CNS 抑制協同作用）",
        "alert": "協同加強中樞神經抑制 → 日間嗜睡、跌倒、認知損害風險\n長期 BZD 有依賴性問題",
        "script": "「合用鎮靜效果可能更強，早晨起床特別注意頭暈，防止跌倒。"
                  "若想漸進式減少安眠藥依賴：褪黑激素建立規律睡眠週期不成癮；"
                  "L-茶胺酸減少入睡前思緒；維生素 B 群修復神經系統。」",
        "products": ["褪黑激素 0.5–1mg  — 調節睡眠周期", "L-茶胺酸 200mg  — 放鬆、無嗜睡",
                     "維生素 B 複合物  — 神經系統支持"],
        "price": "NT$500–900",
        "bg": RGBColor(0xE0, 0xF7, 0xFA),
        "accent": TEAL,
    },
]

for sc in scenarios:
    s = prs.slides.add_slide(BLANK)
    slide_bg(s, sc["bg"])
    title_bar(s, f"銷售場景 {sc['no']}：{sc['drug']}",
              f"觸發條件：{sc['trigger']}")
    footer(s)

    # Alert box
    add_rect(s, Inches(0.5), Inches(1.5), Inches(12.3), Inches(0.85),
             fill=sc["accent"])
    txbox(s, f"MediSpan 警示：{sc['alert']}",
          Inches(0.7), Inches(1.55), Inches(12.0), Inches(0.75),
          size=12, color=WHITE)

    # Script
    add_rect(s, Inches(0.5), Inches(2.5), Inches(7.5), Inches(2.7), fill=WHITE,
             line=sc["accent"])
    txbox(s, "藥師話術", Inches(0.65), Inches(2.55), Inches(3.0), Inches(0.35),
          size=13, bold=True, color=sc["accent"])
    txbox(s, sc["script"], Inches(0.65), Inches(2.92), Inches(7.2), Inches(2.2),
          size=12, color=DGRAY)

    # Products
    add_rect(s, Inches(8.2), Inches(2.5), Inches(4.6), Inches(2.7), fill=NAVY)
    txbox(s, "推薦補充品組合", Inches(8.35), Inches(2.55), Inches(4.3), Inches(0.38),
          size=13, bold=True, color=WHITE)
    cy2 = Inches(2.98)
    for prod in sc["products"]:
        txbox(s, f"▸  {prod}", Inches(8.4), cy2, Inches(4.2), Inches(0.38),
              size=11, color=RGBColor(0xC8, 0xE8, 0xFF))
        cy2 += Inches(0.42)

    # Price badge
    add_rect(s, Inches(8.2), Inches(5.4), Inches(4.6), Inches(0.75), fill=sc["accent"])
    txbox(s, f"預估加購客單價：{sc['price']}",
          Inches(8.35), Inches(5.5), Inches(4.3), Inches(0.55),
          size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Summary arrow
    add_rect(s, Inches(0.5), Inches(5.4), Inches(7.5), Inches(0.75),
             fill=RGBColor(0xF0, 0xF0, 0xF0), line=sc["accent"])
    txbox(s, f"DDI 警示 → 藥師對話 → 補充品銷售  →  {sc['price']} / 次",
          Inches(0.65), Inches(5.5), Inches(7.2), Inches(0.55),
          size=14, bold=True, color=sc["accent"])


# ══════════════════════════════════════════════════════════════
# SLIDE 11 — Revenue Model
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "商業模式與收益試算",
          "每家門市每1–2天1次 DDI 對話 → 系統性補充品收益")
footer(s)

# Conservative
add_rect(s, Inches(0.4), Inches(1.5), Inches(5.8), Inches(4.8), fill=RGBColor(0xE8, 0xF5, 0xFD))
txbox(s, "保守模型", Inches(0.5), Inches(1.6), Inches(5.6), Inches(0.5),
      size=20, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
rows_c = [
    ("每店每日 DDI 對話", "1 次"),
    ("轉換率", "40%"),
    ("平均加購金額", "NT$500"),
    ("每店每日額外收益", "NT$200"),
    ("每店每月額外收益", "NT$6,000"),
]
cy3 = Inches(2.2)
for lbl, val in rows_c:
    txbox(s, lbl, Inches(0.6), cy3, Inches(3.5), Inches(0.38), size=12, color=DGRAY)
    txbox(s, val, Inches(4.1), cy3, Inches(1.8), Inches(0.38), size=12, bold=True,
          color=NAVY, align=PP_ALIGN.RIGHT)
    cy3 += Inches(0.42)
add_rect(s, Inches(0.4), Inches(4.45), Inches(5.8), Inches(0.75), fill=BLUE)
txbox(s, "100 店 / 月：NT$600,000", Inches(0.5), Inches(4.52), Inches(5.6), Inches(0.6),
      size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Aggressive
add_rect(s, Inches(6.5), Inches(1.5), Inches(6.4), Inches(4.8), fill=RGBColor(0xFF, 0xF3, 0xE0))
txbox(s, "積極模型", Inches(6.6), Inches(1.6), Inches(6.2), Inches(0.5),
      size=20, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)
rows_a = [
    ("每店每日 DDI 對話", "3 次"),
    ("轉換率", "50%"),
    ("平均加購金額", "NT$700"),
    ("每店每日額外收益", "NT$1,050"),
    ("每店每月額外收益", "NT$31,500"),
]
cy4 = Inches(2.2)
for lbl, val in rows_a:
    txbox(s, lbl, Inches(6.7), cy4, Inches(3.8), Inches(0.38), size=12, color=DGRAY)
    txbox(s, val, Inches(10.5), cy4, Inches(2.0), Inches(0.38), size=12, bold=True,
          color=NAVY, align=PP_ALIGN.RIGHT)
    cy4 += Inches(0.42)
add_rect(s, Inches(6.5), Inches(4.45), Inches(6.4), Inches(0.75), fill=ORANGE)
txbox(s, "100 店 / 月：NT$3,150,000", Inches(6.6), Inches(4.52), Inches(6.2), Inches(0.6),
      size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# LTV note
add_rect(s, Inches(0.4), Inches(5.45), Inches(12.5), Inches(0.75),
         fill=RGBColor(0xE8, 0xF8, 0xF1), line=GREEN)
txbox(s, "超越交易的更大價值：每次 DDI 專業對話，使患者回訪率預估提升 30–50%，口碑轉介效應無法量化",
      Inches(0.6), Inches(5.55), Inches(12.2), Inches(0.55),
      size=13, color=GREEN, bold=True)


# ══════════════════════════════════════════════════════════════
# SLIDE 12 — Implementation Plan
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "三階段導入計畫", "從試點驗證到全鏈標準化")
footer(s)

phases = [
    {
        "phase": "第一階段", "period": "第 1–2 個月",
        "goal": "數據驗證",
        "bg": BLUE,
        "items": [
            "選定 3–5 家試點門市（高處方量優先）",
            "確認 MediSpan DDI 警示數據完整記錄",
            "建立「補充品推薦追蹤表」",
            "藥師初階話術培訓（五大場景）",
        ],
        "kpi": "DDI 記錄完整率 ≥ 95%\n推薦率 ≥ 60%",
    },
    {
        "phase": "第二階段", "period": "第 3–4 個月",
        "goal": "話術與銷售培訓",
        "bg": TEAL,
        "items": [
            "分析試點數據，優化高轉換率場景",
            "Angela 提供藥學專業深化培訓",
            "Andy 建立 DDI→推薦自動化提示系統",
            "Frank 整合 POS 追蹤 DDI 觸發銷售",
        ],
        "kpi": "轉換率目標 ≥ 40%\nPOS 追蹤系統上線",
    },
    {
        "phase": "第三階段", "period": "第 5–6 個月",
        "goal": "擴張與標準化",
        "bg": RGBColor(0x5B, 0x6A, 0xB0),
        "items": [
            "全鏈推廣，納入新人訓練課程",
            "啟動 Medi-Span Expert AI 整合開發",
            "每月三方數據回顧會議",
            "迭代補充品組合與定價策略",
        ],
        "kpi": "覆蓋率 ≥ 80% 門市\n月收益達 NT$600,000",
    },
]

for i, ph in enumerate(phases):
    x = Inches(0.4 + i * 4.3)
    add_rect(s, x, Inches(1.5), Inches(4.0), Inches(0.7), fill=ph["bg"])
    txbox(s, f"{ph['phase']}  {ph['period']}", x + Inches(0.15), Inches(1.55),
          Inches(3.7), Inches(0.55), size=14, bold=True, color=WHITE)
    add_rect(s, x, Inches(2.2), Inches(4.0), Inches(0.5), fill=RGBColor(0xF0, 0xF4, 0xFF))
    txbox(s, f"目標：{ph['goal']}", x + Inches(0.15), Inches(2.25), Inches(3.7), Inches(0.38),
          size=13, bold=True, color=ph["bg"])
    add_rect(s, x, Inches(2.72), Inches(4.0), Inches(2.85),
             fill=WHITE, line=RGBColor(0xCC, 0xCC, 0xCC))
    cy5 = Inches(2.8)
    for item in ph["items"]:
        txbox(s, f"• {item}", x + Inches(0.15), cy5, Inches(3.7), Inches(0.38),
              size=11, color=DGRAY)
        cy5 += Inches(0.42)
    add_rect(s, x, Inches(5.6), Inches(4.0), Inches(0.9), fill=ph["bg"])
    txbox(s, ph["kpi"], x + Inches(0.1), Inches(5.65), Inches(3.8), Inches(0.8),
          size=11, color=WHITE)

# Arrow connectors (simple)
for ax in [Inches(4.4), Inches(8.7)]:
    add_rect(s, ax, Inches(3.3), Inches(0.4), Inches(0.08), fill=NAVY)
    txbox(s, "▶", ax, Inches(3.22), Inches(0.4), Inches(0.38),
          size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 13 — Medi-Span Expert AI
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
slide_bg(s)
title_bar(s, "Medi-Span Expert AI — 未來護城河",
          "2026 年 2 月上線 | MCP 架構 | NEOV.AI 整合開發")
footer(s)

add_rect(s, Inches(0.4), Inches(1.5), Inches(12.5), Inches(0.75), fill=NAVY)
txbox(s, "業界第一個基於 MCP（Model Context Protocol）架構的藥物情報 AI 整合框架",
      Inches(0.6), Inches(1.58), Inches(12.2), Inches(0.55),
      size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Flow diagram
flow = [
    ("藥師掃描處方", BLUE),
    ("MediSpan DDI\n引擎分析\n（0.4 秒）", RED),
    ("AI Agent\n自動配對\n補充品邏輯", TEAL),
    ("個性化話術\n提示卡顯示", GREEN),
    ("銷售成交\n+記錄優化", ORANGE),
]
for i, (label, bg) in enumerate(flow):
    x = Inches(0.5 + i * 2.55)
    add_rect(s, x, Inches(2.5), Inches(2.2), Inches(1.4), fill=bg)
    txbox(s, label, x + Inches(0.1), Inches(2.6), Inches(2.0), Inches(1.2),
          size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if i < len(flow) - 1:
        txbox(s, "→", x + Inches(2.2), Inches(2.9), Inches(0.35), Inches(0.6),
              size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# Moat points
moat_items = [
    ("數據深度", "MediSpan 150+ 專家知識庫，爬蟲無法複製"),
    ("整合深度", "嵌入大樹作業系統的定制化配置，每家藥局獨特"),
    ("時間優勢", "導入到系統成熟需 6–12 個月學習週期"),
    ("品牌護城河", "「醫院等級藥物安全系統」品牌認知，口頭說明無法取代"),
]
txbox(s, "競爭護城河分析", Inches(0.5), Inches(4.15), Inches(5), Inches(0.38),
      size=15, bold=True, color=NAVY)
for i, (title, desc) in enumerate(moat_items):
    y = Inches(4.6 + i * 0.55)
    add_rect(s, Inches(0.5), y, Inches(2.2), Inches(0.45), fill=NAVY)
    txbox(s, title, Inches(0.55), y + Inches(0.05), Inches(2.1), Inches(0.35),
          size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    txbox(s, desc, Inches(2.85), y + Inches(0.05), Inches(9.7), Inches(0.35),
          size=12, color=DGRAY)

add_rect(s, Inches(0.5), Inches(6.85), Inches(12.5), Inches(0.45), fill=TEAL)
txbox(s, "保守估計：此架構為大樹藥局建立 3–5 年先發競爭優勢",
      Inches(0.7), Inches(6.9), Inches(12.2), Inches(0.38),
      size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════
# SLIDE 14 — Call to Action
# ══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, W, H, fill=NAVY)

txbox(s, "一張處方，兩種命運", Inches(1), Inches(0.8), Inches(11.3), Inches(0.8),
      size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(s, Inches(2), Inches(1.6), Inches(9.33), Inches(0.04), fill=TEAL)

# Destiny 1
add_rect(s, Inches(0.5), Inches(1.8), Inches(5.8), Inches(2.1),
         fill=RGBColor(0x2C, 0x3E, 0x55))
txbox(s, "命運一（現狀）", Inches(0.65), Inches(1.88), Inches(5.5), Inches(0.42),
      size=14, bold=True, color=ORANGE)
txbox(s, "患者取藥 → 簡短說明 → 離開\nDDI 警示留在系統，無對話\n無補充品推薦，無回訪理由",
      Inches(0.65), Inches(2.3), Inches(5.5), Inches(1.4),
      size=13, color=RGBColor(0xA0, 0xB8, 0xD0))

# Destiny 2
add_rect(s, Inches(7.0), Inches(1.8), Inches(5.8), Inches(2.1),
         fill=RGBColor(0x1A, 0x4A, 0x3A))
txbox(s, "命運二（數據藥局）", Inches(7.15), Inches(1.88), Inches(5.5), Inches(0.42),
      size=14, bold=True, color=TEAL)
txbox(s, "MediSpan 0.4 秒偵測 → 話術啟動\n科學推薦補充品 → 患者信任建立\n成為回頭客 + 口碑轉介",
      Inches(7.15), Inches(2.3), Inches(5.5), Inches(1.4),
      size=13, color=RGBColor(0xA0, 0xD8, 0xC0))

txbox(s, "▶▶▶", Inches(6.07), Inches(2.45), Inches(0.9), Inches(0.6),
      size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

# Action table
add_rect(s, Inches(0.5), Inches(4.1), Inches(12.3), Inches(0.42), fill=BLUE)
for txt, x in [("步驟", Inches(0.7)), ("負責人", Inches(3.5)), ("時間目標", Inches(8.0))]:
    txbox(s, txt, x, Inches(4.15), Inches(2.5), Inches(0.35),
          size=13, bold=True, color=WHITE)

action_rows = [
    ("確認試點門市（3–5家）",     "大樹藥局管理團隊",          "2 週內"),
    ("MediSpan 警示數據盤點",    "Angela Hsu (WK)",           "2 週內"),
    ("藥師話術培訓排程",         "Angela + 大樹培訓團隊",      "第 1 個月"),
    ("DDI→補充品 AI 工作流開發", "Andy Lam (NEOV.AI)",        "第 1–2 個月"),
    ("POS 追蹤模組整合",         "Frank Kao (Insight SW)",    "第 2 個月"),
]
for i, (step, owner, time) in enumerate(action_rows):
    y = Inches(4.55 + i * 0.42)
    bg = RGBColor(0x1E, 0x3A, 0x5F) if i % 2 == 0 else RGBColor(0x16, 0x2D, 0x4A)
    add_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.4), fill=bg)
    txbox(s, step, Inches(0.7), y + Inches(0.04), Inches(2.7), Inches(0.35),
          size=11, color=WHITE)
    txbox(s, owner, Inches(3.5), y + Inches(0.04), Inches(4.4), Inches(0.35),
          size=11, color=RGBColor(0xA8, 0xD0, 0xF0))
    txbox(s, time, Inches(8.0), y + Inches(0.04), Inches(2.5), Inches(0.35),
          size=11, color=TEAL)

txbox(s, "「從領藥的地方，到值得信賴的醫療服務專家。」",
      Inches(1), Inches(7.0), Inches(11.3), Inches(0.42),
      size=14, bold=True, color=RGBColor(0x80, 0xD8, 0xC0), align=PP_ALIGN.CENTER)


# ── Save ───────────────────────────────────────────────────────
out = "/home/user/adgo/daashu-medispan-proposal.pptx"
prs.save(out)
print(f"Saved: {out}")
