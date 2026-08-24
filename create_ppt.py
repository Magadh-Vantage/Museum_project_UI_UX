import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def build_presentation():
    prs = Presentation()
    # Set slide dimensions to 16:9 Widescreen (13.333 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    # Color Palette Definitions
    COLOR_BG_BASE = RGBColor(15, 23, 42)       # #0F172A Deep Slate
    COLOR_BG_CARD = RGBColor(30, 41, 59)       # #1E293B Surface Slate
    COLOR_GOLD = RGBColor(197, 155, 39)        # #C59B27 Royal Gold
    COLOR_BLUE = RGBColor(56, 189, 248)        # #38BDF8 Electric Blue
    COLOR_WHITE = RGBColor(248, 250, 252)      # #F8FAFC Crisp White
    COLOR_SLATE_TEXT = RGBColor(148, 163, 184) # #94A3B8 Subtitle Slate
    COLOR_DARK_TEXT = RGBColor(15, 23, 42)     # Dark text on gold

    # 1. Full Screen Background Shape
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_BG_BASE
    bg.line.fill.background()

    # 2. Top Header Bar (Fixed 60px height approx 0.65 in)
    header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.65))
    header.fill.solid()
    header.fill.fore_color.rgb = RGBColor(11, 17, 32) # #0B1120 Inset Dark
    header.line.color.rgb = COLOR_GOLD
    header.line.width = Pt(1)

    tf_hdr = header.text_frame
    tf_hdr.word_wrap = True
    p_hdr = tf_hdr.paragraphs[0]
    p_hdr.text = "🏛️ GOVERNMENT MUSEUM DIGITAL EXHIBIT   |   55\" INTERACTIVE TOUCH KIOSK"
    p_hdr.font.size = Pt(11)
    p_hdr.font.bold = True
    p_hdr.font.color.rgb = COLOR_GOLD
    p_hdr.font.name = "Outfit"

    # Add right side header controls info (Language / Audio)
    hdr_right = slide.shapes.add_textbox(Inches(9.5), Inches(0.08), Inches(3.6), Inches(0.5))
    tf_hr = hdr_right.text_frame
    p_hr = tf_hr.paragraphs[0]
    p_hr.alignment = PP_ALIGN.RIGHT
    p_hr.text = "[ 🌐 HINDI / ENGLISH ]   [ 🔊 AUDIO 80% ]   [ 🔍 A+ / A- ]"
    p_hr.font.size = Pt(10)
    p_hr.font.bold = True
    p_hr.font.color.rgb = COLOR_BLUE
    p_hr.font.name = "Inter"

    # 3. Main Title & Subtitle Card (Left Top Section)
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.85), Inches(7.5), Inches(1.8))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    
    # Badge
    p_badge = tf_t.paragraphs[0]
    p_badge.text = "SLIDE 1: TITLE & WELCOME SCREEN (LEVEL 0 / LEVEL 1)"
    p_badge.font.size = Pt(11)
    p_badge.font.bold = True
    p_badge.font.color.rgb = COLOR_BLUE
    p_badge.font.name = "Inter"

    # Title
    p_title = tf_t.add_paragraph()
    p_title.text = "Pre-Presidential Journey of\nDr. Rajendra Prasad"
    p_title.font.size = Pt(26)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_WHITE
    p_title.font.name = "Outfit"

    # Subtitle
    p_sub = tf_t.add_paragraph()
    p_sub.text = "\"From Legal Luminary to the First President of India (1950)\""
    p_sub.font.size = Pt(13)
    p_sub.font.italic = True
    p_sub.font.color.rgb = COLOR_GOLD
    p_sub.font.name = "Inter"

    # 4. Hero Visual Asset (Right Top Section)
    img_path = r"C:\Users\hp\.gemini\antigravity\brain\c6935483-6f7c-4a9f-b91a-5d7d7e563d75\dr_rajendra_prasad_portrait_hero_1785527107082.jpg"
    if os.path.exists(img_path):
        # Frame background for portrait
        img_frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.5), Inches(0.85), Inches(4.3), Inches(2.5))
        img_frame.fill.solid()
        img_frame.fill.fore_color.rgb = COLOR_BG_CARD
        img_frame.line.color.rgb = COLOR_GOLD
        img_frame.line.width = Pt(2)
        
        slide.shapes.add_picture(img_path, Inches(8.6), Inches(0.95), width=Inches(4.1), height=Inches(2.3))

    # 5. Attract Mode Call-To-Action Banner (Center Interactive Prompt)
    cta = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(2.75), Inches(12.133), Inches(0.75))
    cta.fill.solid()
    cta.fill.fore_color.rgb = COLOR_GOLD
    cta.line.color.rgb = COLOR_WHITE
    cta.line.width = Pt(1.5)

    tf_cta = cta.text_frame
    tf_cta.word_wrap = True
    p_cta = tf_cta.paragraphs[0]
    p_cta.alignment = PP_ALIGN.CENTER
    p_cta.text = "👇 TOUCH SCREEN ANYWHERE TO BEGIN  |  अन्वेषण के लिए स्क्रीन को स्पर्श करें"
    p_cta.font.size = Pt(15)
    p_cta.font.bold = True
    p_cta.font.color.rgb = COLOR_DARK_TEXT
    p_cta.font.name = "Outfit"

    # 6. Level 1: 9-Phase Journey Quick-Jump Grid (3x3 Grid)
    grid_title = slide.shapes.add_textbox(Inches(0.6), Inches(3.55), Inches(12.133), Inches(0.35))
    tf_gt = grid_title.text_frame
    p_gt = tf_gt.paragraphs[0]
    p_gt.text = "EXHIBIT JOURNEY ARCHITECTURE (9 INTERACTIVE MILESTONES)"
    p_gt.font.size = Pt(12)
    p_gt.font.bold = True
    p_gt.font.color.rgb = COLOR_SLATE_TEXT
    p_gt.font.name = "Outfit"

    milestones = [
        ("P-01: Welcome & Intro", "Digital Storytelling Overview (1884–1950)"),
        ("P-02: The Advocate", "Calcutta & Patna High Court Practice (1911-16)"),
        ("P-03: The Journalist", "Editorial Leadership in Searchlight & Desh"),
        ("P-04: The Humanitarian", "1934 Bihar Earthquake Relief Operations"),
        ("P-05: The Educationist", "Co-founding Bihar Vidyapeeth (1921)"),
        ("P-06: Champaran Leader", "1917 Satyagraha & Farmers' Letter Explorer"),
        ("P-07: Agriculture Minister", "1946 Interim Govt & Grow More Food Campaign"),
        ("P-08: Constitution Maker", "Presiding President of Constituent Assembly"),
        ("P-09: First President", "Unanimous Election & 1950 Swearing-in Oath")
    ]

    start_x = 0.6
    start_y = 3.9
    card_w = 3.85
    card_h = 0.85
    gap_x = 0.28
    gap_y = 0.15

    for idx, (m_title, m_desc) in enumerate(milestones):
        row = idx // 3
        col = idx % 3
        cx = start_x + col * (card_w + gap_x)
        cy = start_y + row * (card_h + gap_y)

        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(cy), Inches(card_w), Inches(card_h))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_BG_CARD
        card.line.color.rgb = COLOR_GOLD if idx == 0 else COLOR_SLATE_TEXT
        card.line.width = Pt(1.5 if idx == 0 else 1)

        tf_c = card.text_frame
        tf_c.word_wrap = True
        
        p1 = tf_c.paragraphs[0]
        p1.text = m_title
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_GOLD if idx == 0 else COLOR_WHITE
        p1.font.name = "Outfit"

        p2 = tf_c.add_paragraph()
        p2.text = m_desc
        p2.font.size = Pt(9)
        p2.font.color.rgb = COLOR_SLATE_TEXT
        p2.font.name = "Inter"

    # 7. Persistent Bottom Timeline Stepper Bar (70px height)
    timeline_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(6.85), Inches(13.333), Inches(0.65))
    timeline_bg.fill.solid()
    timeline_bg.fill.fore_color.rgb = RGBColor(11, 17, 32)
    timeline_bg.line.color.rgb = COLOR_GOLD
    timeline_bg.line.width = Pt(1)

    tf_tl = timeline_bg.text_frame
    tf_tl.word_wrap = True
    p_tl = tf_tl.paragraphs[0]
    p_tl.alignment = PP_ALIGN.CENTER
    p_tl.text = "TIMELINE NAVIGATOR:   (1) Intro ★   ➜   (2) Advocate   ➜   (3) Journalist   ➜   (4) Relief   ➜   (5) Education   ➜   (6) Champaran   ➜   (7) Agri   ➜   (8) Constitution   ➜   (9) President (1950)"
    p_tl.font.size = Pt(10)
    p_tl.font.bold = True
    p_tl.font.color.rgb = COLOR_GOLD
    p_tl.font.name = "Inter"

    # Save presentation
    output_pptx = r"d:\Museum project\Slide1_Dr_Rajendra_Prasad_Kiosk.pptx"
    prs.save(output_pptx)
    print(f"Presentation saved successfully to: {output_pptx}")

if __name__ == "__main__":
    build_presentation()
