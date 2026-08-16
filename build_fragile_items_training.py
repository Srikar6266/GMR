#!/usr/bin/env python3
"""Build professional Fragile Items training PPT + Word for Civil Team."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import nsmap
from pptx.util import Inches, Pt, Emu

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches as DocInches, Pt as DocPt, RGBColor as DocRGBColor, Cm

OUT_DIR = Path("/workspace")
PPT_PATH = OUT_DIR / "Fragile_Items_Safe_Handling_Storage_Training.pptx"
DOC_PATH = OUT_DIR / "Fragile_Items_Safe_Handling_Storage_Training.docx"

# Professional safety palette (navy / slate / amber accent)
NAVY = RGBColor(0x0F, 0x2A, 0x44)
SLATE = RGBColor(0x1F, 0x3A, 0x56)
TEAL = RGBColor(0x1A, 0x6B, 0x6B)
AMBER = RGBColor(0xC4, 0x7B, 0x2D)
LIGHT = RGBColor(0xF4, 0xF6, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5A, 0x6A, 0x7A)
GREEN = RGBColor(0x2E, 0x7D, 0x4F)
RED = RGBColor(0xA8, 0x3A, 0x32)
LINE = RGBColor(0xD0, 0xD7, 0xDE)


def set_run(run, size=18, bold=False, color=DARK, font="Calibri"):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def add_bg(slide, color):
    fill = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
    )
    fill.fill.solid()
    fill.fill.fore_color.rgb = color
    fill.line.fill.background()
    # send to back
    spTree = slide.shapes._spTree
    sp = fill._element
    spTree.remove(sp)
    spTree.insert(2, sp)
    return fill


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(1)
    return shape


def add_text_box(slide, left, top, width, height, text, size=18, bold=False,
                 color=DARK, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def add_bullets(slide, left, top, width, height, items, size=16, color=DARK, bullet="•"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        p.level = 0
        run = p.add_run()
        run.text = f"{bullet}  {item}"
        set_run(run, size=size, color=color)
    return box


def add_image_placeholder(slide, left, top, width, height, caption="Add site photo here"):
    shape = add_rect(slide, left, top, width, height, LIGHT, LINE)
    # dashed look via inner label
    add_text_box(
        slide,
        left,
        top + height / 2 - Inches(0.35),
        width,
        Inches(0.7),
        f"[ Image Placeholder ]\n{caption}",
        size=12,
        color=MUTED,
        align=PP_ALIGN.CENTER,
    )
    return shape


def header_bar(slide, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(1.05), NAVY)
    add_rect(slide, Inches(0), Inches(1.05), Inches(13.333), Inches(0.08), AMBER)
    add_text_box(slide, Inches(0.5), Inches(0.22), Inches(12), Inches(0.5),
                 title, size=26, bold=True, color=WHITE)
    if subtitle:
        add_text_box(slide, Inches(0.5), Inches(0.65), Inches(12), Inches(0.35),
                     subtitle, size=12, color=RGBColor(0xC8, 0xD5, 0xE0))


def footer(slide, page, total=14):
    add_text_box(
        slide, Inches(0.5), Inches(7.1), Inches(10), Inches(0.3),
        "Fragile Items – Safe Handling & Storage  |  Civil Team Toolbox Training",
        size=10, color=MUTED,
    )
    add_text_box(
        slide, Inches(11.5), Inches(7.1), Inches(1.5), Inches(0.3),
        f"{page} / {total}", size=10, color=MUTED, align=PP_ALIGN.RIGHT,
    )


def build_ppt():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    total = 14

    # ---- 1 Title ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    add_rect(s, Inches(0), Inches(0), Inches(0.35), Inches(7.5), NAVY)
    add_rect(s, Inches(0.35), Inches(0), Inches(0.12), Inches(7.5), AMBER)
    add_text_box(s, Inches(1.0), Inches(1.6), Inches(11), Inches(0.4),
                 "SITE SAFETY / TOOLBOX TRAINING", size=14, bold=True, color=TEAL)
    add_text_box(s, Inches(1.0), Inches(2.1), Inches(11), Inches(1.2),
                 "Fragile Items – Safe Handling\n& Storage Training",
                 size=36, bold=True, color=NAVY)
    add_rect(s, Inches(1.0), Inches(3.55), Inches(2.2), Inches(0.07), AMBER)
    meta = [
        "Target Team: Civil Team",
        "Duration: 30–45 Minutes",
        "Training Type: Toolbox / Site Safety Training",
    ]
    add_bullets(s, Inches(1.0), Inches(3.9), Inches(8), Inches(1.5), meta, size=16, color=SLATE, bullet="▸")
    add_image_placeholder(s, Inches(9.2), Inches(4.3), Inches(3.5), Inches(2.3), "Civil site / materials photo")
    add_text_box(s, Inches(1.0), Inches(6.8), Inches(8), Inches(0.3),
                 "Insert company logo / project name in header area when editing",
                 size=11, color=MUTED)

    # ---- 2 Agenda ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Training Agenda", "30–45 minutes structured session")
    agenda = [
        ("01", "Introduction & Objective", "5 Min"),
        ("02", "What are Fragile Items?", "5 Min"),
        ("03", "Identification of Fragile Materials at Site", "5 Min"),
        ("04", "Correct Handling & Lifting Methods", "10 Min"),
        ("05", "Transportation & Movement of Fragile Items", "5 Min"),
        ("06", "Storage & Stacking Requirements", "5 Min"),
        ("07", "Do's & Don'ts", "5 Min"),
        ("08", "Practical Demonstration", "5 Min"),
        ("09", "Questions & Feedback", "5 Min"),
    ]
    y = 1.35
    for num, topic, dur in agenda:
        add_rect(s, Inches(0.6), Inches(y), Inches(0.55), Inches(0.42), NAVY)
        add_text_box(s, Inches(0.6), Inches(y + 0.05), Inches(0.55), Inches(0.35),
                     num, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, Inches(1.35), Inches(y + 0.05), Inches(9), Inches(0.35),
                     topic, size=15, color=DARK)
        add_text_box(s, Inches(11.2), Inches(y + 0.05), Inches(1.4), Inches(0.35),
                     dur, size=14, bold=True, color=TEAL, align=PP_ALIGN.RIGHT)
        y += 0.55
    footer(s, 2, total)

    # ---- 3 Objectives ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Key Training Objectives", "What we will achieve today")
    objectives = [
        "Prevent damage to fragile materials and finished works.",
        "Prevent injuries during lifting, shifting and transportation.",
        "Ensure proper storage and protection of fragile items.",
        "Reduce material wastage and replacement costs.",
        "Make the Civil Team responsible for safe material handling.",
        "Maintain good housekeeping and organized storage areas.",
    ]
    # two columns of cards
    for i, obj in enumerate(objectives):
        col = i % 2
        row = i // 2
        left = Inches(0.55 + col * 6.3)
        top = Inches(1.4 + row * 1.55)
        add_rect(s, left, top, Inches(6.0), Inches(1.35), LIGHT, LINE)
        add_rect(s, left, top, Inches(0.12), Inches(1.35), TEAL if i % 2 == 0 else AMBER)
        add_text_box(s, left + Inches(0.35), top + Inches(0.15), Inches(5.4), Inches(0.35),
                     f"Objective {i+1:02d}", size=12, bold=True, color=TEAL)
        add_text_box(s, left + Inches(0.35), top + Inches(0.5), Inches(5.4), Inches(0.7),
                     obj, size=15, color=DARK)
    footer(s, 3, total)

    # ---- 4 What are Fragile Items ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "What are Fragile Items?", "Materials that break, crack, chip or get damaged easily")
    add_text_box(
        s, Inches(0.6), Inches(1.4), Inches(7.5), Inches(1.8),
        "Fragile items are materials that can be easily damaged by impact, "
        "pressure, improper lifting, vibration, stacking, moisture or rough handling. "
        "Once damaged, they often cannot be repaired and must be replaced — causing "
        "delay, cost and safety risk.",
        size=17, color=DARK,
    )
    points = [
        "Easily damaged by impact or pressure",
        "Often expensive / long lead-time to replace",
        "Can cause injury if broken during handling",
        "Usually marked FRAGILE / HANDLE WITH CARE",
    ]
    add_bullets(s, Inches(0.6), Inches(3.4), Inches(7.2), Inches(2.5), points, size=16)
    add_image_placeholder(s, Inches(8.5), Inches(1.5), Inches(4.3), Inches(4.8),
                          "Photo: fragile materials on site")
    footer(s, 4, total)

    # ---- 5 Identification ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Identification of Fragile Materials at Site", "Know what needs special care")
    items = [
        "Glass and glass panels",
        "Tiles and marble",
        "Wash basins & sanitary fixtures",
        "Ceramic items",
        "Doors and door frames",
        "False ceiling materials",
        "Lighting fixtures",
        "Finished civil surfaces",
        "ACP / decorative panels",
        "Mirrors",
        "Pre-finished materials",
        "Any item marked FRAGILE / HANDLE WITH CARE",
    ]
    for i, item in enumerate(items):
        col = i // 6
        row = i % 6
        left = Inches(0.55 + col * 4.15)
        top = Inches(1.35 + row * 0.75)
        add_rect(s, left, top, Inches(3.95), Inches(0.62), LIGHT, LINE)
        add_text_box(s, left + Inches(0.2), top + Inches(0.12), Inches(3.55), Inches(0.4),
                     f"•  {item}", size=13, color=DARK)
    add_image_placeholder(s, Inches(8.85), Inches(1.35), Inches(3.95), Inches(4.7),
                          "Collage: glass / tiles / sanitary / panels")
    footer(s, 5, total)

    # ---- 6 Safe Handling ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Correct Handling & Lifting Methods", "10 minutes — core practical topic")
    left_items = [
        "Check material condition before moving.",
        "Use correct number of workers for weight & size.",
        "Wear required PPE at all times.",
        "Use proper lifting posture (bend knees, keep back straight).",
    ]
    right_items = [
        "Do not drag fragile materials on the floor.",
        "Do not throw, drop or slide materials.",
        "Hold from designated / supporting points.",
        "Use suitable trolleys or lifting equipment.",
    ]
    add_rect(s, Inches(0.5), Inches(1.35), Inches(6.0), Inches(4.5), LIGHT, LINE)
    add_text_box(s, Inches(0.75), Inches(1.5), Inches(5.5), Inches(0.4),
                 "SAFE PRACTICE", size=14, bold=True, color=GREEN)
    add_bullets(s, Inches(0.75), Inches(2.05), Inches(5.5), Inches(3.5), left_items, size=15)
    add_rect(s, Inches(6.8), Inches(1.35), Inches(6.0), Inches(4.5), LIGHT, LINE)
    add_text_box(s, Inches(7.05), Inches(1.5), Inches(5.5), Inches(0.4),
                 "AVOID THESE ACTIONS", size=14, bold=True, color=RED)
    add_bullets(s, Inches(7.05), Inches(2.05), Inches(5.5), Inches(3.5), right_items, size=15)
    add_image_placeholder(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(0.9),
                          "Optional: add lifting / PPE demonstration photo strip here")
    footer(s, 6, total)

    # ---- 7 Transportation ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Transportation & Movement of Fragile Items", "Plan the route — protect the material")
    steps = [
        ("1", "Plan", "Plan the movement route before shifting."),
        ("2", "Clear", "Keep the route clear of obstacles."),
        ("3", "Secure", "Secure materials during transportation."),
        ("4", "Limit", "Do not overload trolleys."),
        ("5", "Steady", "Avoid sudden movements and impacts."),
        ("6", "Protect", "Protect corners and finished surfaces."),
        ("7", "Stack Rule", "Never place heavy materials over fragile items."),
    ]
    for i, (num, title, desc) in enumerate(steps):
        y = 1.3 + i * 0.7
        add_rect(s, Inches(0.55), Inches(y), Inches(0.55), Inches(0.55), NAVY)
        add_text_box(s, Inches(0.55), Inches(y + 0.08), Inches(0.55), Inches(0.4),
                     num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, Inches(1.3), Inches(y + 0.02), Inches(2.2), Inches(0.35),
                     title, size=15, bold=True, color=TEAL)
        add_text_box(s, Inches(3.5), Inches(y + 0.02), Inches(5.5), Inches(0.5),
                     desc, size=14, color=DARK)
    add_image_placeholder(s, Inches(9.3), Inches(1.3), Inches(3.5), Inches(5.2),
                          "Photo: trolley movement on site")
    footer(s, 7, total)

    # ---- 8 Storage ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Storage & Stacking Requirements", "Store safe — retrieve safe")
    storage = [
        "Store materials in designated areas only.",
        "Keep fragile items on stable and level surfaces.",
        "Use racks, pallets, supports or protective frames where required.",
        "Avoid unstable stacking.",
        "Keep heavy materials at the bottom.",
        "Protect from water, dust and construction activities.",
        "Maintain clear access around stored materials.",
        "Follow manufacturer's storage instructions.",
    ]
    add_bullets(s, Inches(0.6), Inches(1.4), Inches(7.3), Inches(4.8), storage, size=16)
    add_image_placeholder(s, Inches(8.3), Inches(1.4), Inches(4.5), Inches(4.8),
                          "Photo: correct rack / pallet storage")
    footer(s, 8, total)

    # ---- 9 Do's ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Do's — Safe Practices", "Follow every time")
    dos = [
        "Inspect before handling.",
        "Use proper PPE.",
        "Use proper tools / equipment.",
        "Follow lifting instructions.",
        "Use sufficient manpower.",
        "Communicate clearly during movement.",
        "Protect materials after installation.",
        "Report any damage immediately.",
    ]
    for i, item in enumerate(dos):
        col = i % 2
        row = i // 2
        left = Inches(0.55 + col * 6.35)
        top = Inches(1.4 + row * 1.2)
        add_rect(s, left, top, Inches(6.05), Inches(1.0), RGBColor(0xE8, 0xF5, 0xEE), RGBColor(0x2E, 0x7D, 0x4F))
        add_text_box(s, left + Inches(0.25), top + Inches(0.28), Inches(5.5), Inches(0.45),
                     f"✓   {item}", size=16, bold=True, color=RGBColor(0x1B, 0x5E, 0x3A))
    footer(s, 9, total)

    # ---- 10 Don'ts ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Don'ts — Unsafe Practices", "Stop these habits")
    donts = [
        "Don't throw or drop materials.",
        "Don't drag fragile items.",
        "Don't stand fragile items in an unstable position.",
        "Don't stack heavy materials on fragile materials.",
        "Don't use damaged lifting equipment.",
        "Don't carry oversized materials without assistance.",
        "Don't leave fragile materials exposed to unnecessary site activity.",
        "Don't continue using damaged material without inspection.",
    ]
    for i, item in enumerate(donts):
        col = i % 2
        row = i // 2
        left = Inches(0.55 + col * 6.35)
        top = Inches(1.4 + row * 1.2)
        add_rect(s, left, top, Inches(6.05), Inches(1.0), RGBColor(0xFB, 0xEE, 0xEC), RGBColor(0xA8, 0x3A, 0x32))
        add_text_box(s, left + Inches(0.25), top + Inches(0.28), Inches(5.5), Inches(0.45),
                     f"✗   {item}", size=15, bold=True, color=RGBColor(0x7A, 0x28, 0x22))
    footer(s, 10, total)

    # ---- 11 Practical Demonstration ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Practical Demonstration", "Trainer to demonstrate live on site")
    demos = [
        "Correct lifting of a fragile item",
        "Correct use of a trolley",
        "Safe movement through a work area",
        "Correct placement on a pallet / rack",
        "Proper protection using sheets, foam, cardboard or other suitable protection",
        "Correct method of reporting damaged material",
    ]
    for i, d in enumerate(demos):
        y = 1.35 + i * 0.8
        add_rect(s, Inches(0.55), Inches(y), Inches(0.7), Inches(0.6), AMBER)
        add_text_box(s, Inches(0.55), Inches(y + 0.1), Inches(0.7), Inches(0.4),
                     str(i + 1), size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(s, Inches(1.4), Inches(y), Inches(7.2), Inches(0.6), LIGHT, LINE)
        add_text_box(s, Inches(1.6), Inches(y + 0.12), Inches(6.8), Inches(0.4),
                     d, size=15, color=DARK)
    add_image_placeholder(s, Inches(8.9), Inches(1.35), Inches(3.9), Inches(4.9),
                          "Photo during live demonstration")
    footer(s, 11, total)

    # ---- 12 Key Takeaways ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Key Takeaways", "Remember before you leave")
    takes = [
        ("Identify", "Know which materials are fragile before you touch them."),
        ("Handle", "Inspect, use PPE, lift correctly — never drag or drop."),
        ("Move", "Plan the route, secure the load, move steadily."),
        ("Store", "Designated area, stable base, heavy at bottom, protect from damage."),
        ("Report", "Any damage must be reported immediately — do not hide it."),
    ]
    for i, (t, d) in enumerate(takes):
        y = 1.35 + i * 0.95
        add_rect(s, Inches(0.55), Inches(y), Inches(2.2), Inches(0.75), NAVY)
        add_text_box(s, Inches(0.55), Inches(y + 0.18), Inches(2.2), Inches(0.45),
                     t, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, Inches(3.0), Inches(y + 0.18), Inches(9.5), Inches(0.5),
                     d, size=16, color=DARK)
    footer(s, 12, total)

    # ---- 13 Closing ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    add_rect(s, Inches(0.8), Inches(1.8), Inches(11.7), Inches(3.6), SLATE)
    add_text_box(s, Inches(1.2), Inches(2.1), Inches(10.9), Inches(0.4),
                 "CLOSING MESSAGE", size=14, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    add_text_box(
        s, Inches(1.3), Inches(2.7), Inches(10.7), Inches(2.2),
        "Handle every fragile item as if it is already installed.\n\n"
        "One careless movement can cause material damage,\nrework, additional cost and safety risks.",
        size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
    )
    add_text_box(s, Inches(1.0), Inches(6.2), Inches(11.3), Inches(0.4),
                 "Civil Team  •  Toolbox Training  •  Safe Handling & Storage",
                 size=14, color=RGBColor(0xC8, 0xD5, 0xE0), align=PP_ALIGN.CENTER)

    # ---- 14 Q&A ----
    s = prs.slides.add_slide(blank)
    add_bg(s, WHITE)
    header_bar(s, "Questions & Feedback", "5 minutes")
    add_text_box(s, Inches(0.8), Inches(2.2), Inches(11.5), Inches(1.0),
                 "Any questions from the Civil Team?",
                 size=28, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text_box(s, Inches(1.5), Inches(3.4), Inches(10), Inches(1.2),
                 "Share site examples, near-misses or concerns related to fragile material handling.\n"
                 "Trainer to record attendance and feedback.",
                 size=16, color=SLATE, align=PP_ALIGN.CENTER)
    add_image_placeholder(s, Inches(3.5), Inches(4.8), Inches(6.3), Inches(1.7),
                          "Optional: attendance / team photo")
    footer(s, 14, total)

    prs.save(PPT_PATH)
    print("Saved", PPT_PATH)


# ---------------- Word document ----------------

def shade_cell(cell, hex_color):
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, color=None, size=11, center=False):
    cell.text = ""
    p = cell.paragraphs[0]
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = DocPt(size)
    run.font.name = "Calibri"
    if color:
        run.font.color.rgb = DocRGBColor(
            int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        )


def add_heading_styled(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = DocRGBColor(0x0F, 0x2A, 0x44)
        run.font.name = "Calibri"
    return h


def build_docx():
    doc = Document()

    # margins
    for section in doc.sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    # Title block
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("FRAGILE ITEMS – SAFE HANDLING & STORAGE TRAINING")
    r.bold = True
    r.font.size = DocPt(20)
    r.font.color.rgb = DocRGBColor(0x0F, 0x2A, 0x44)
    r.font.name = "Calibri"

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("Toolbox / Site Safety Training  |  Civil Team  |  Duration: 30–45 Minutes")
    r.font.size = DocPt(12)
    r.font.color.rgb = DocRGBColor(0x1A, 0x6B, 0x6B)
    r.font.name = "Calibri"

    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = note.add_run("[ Insert Company Logo / Project Name / Date here ]")
    r.italic = True
    r.font.size = DocPt(10)
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)

    doc.add_paragraph()

    # Meta table
    meta = doc.add_table(rows=4, cols=2)
    meta.style = "Table Grid"
    meta_data = [
        ("Target Team", "Civil Team"),
        ("Duration", "30–45 Minutes"),
        ("Training Type", "Toolbox Training / Site Safety Training"),
        ("Document Use", "Trainer guide + participant handout (edit as needed)"),
    ]
    for i, (k, v) in enumerate(meta_data):
        set_cell_text(meta.rows[i].cells[0], k, bold=True, color="0F2A44", size=11)
        set_cell_text(meta.rows[i].cells[1], v, size=11)
        shade_cell(meta.rows[i].cells[0], "E8EEF3")

    doc.add_paragraph()
    add_heading_styled(doc, "1. Training Agenda", 1)

    agenda = doc.add_table(rows=10, cols=3)
    agenda.style = "Table Grid"
    agenda.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Sl. No.", "Topic", "Duration"]
    for i, h in enumerate(headers):
        set_cell_text(agenda.rows[0].cells[i], h, bold=True, color="FFFFFF", size=11, center=True)
        shade_cell(agenda.rows[0].cells[i], "0F2A44")

    agenda_rows = [
        ("1", "Introduction & Objective", "5 Min"),
        ("2", "What are Fragile Items?", "5 Min"),
        ("3", "Identification of Fragile Materials at Site", "5 Min"),
        ("4", "Correct Handling & Lifting Methods", "10 Min"),
        ("5", "Transportation & Movement of Fragile Items", "5 Min"),
        ("6", "Storage & Stacking Requirements", "5 Min"),
        ("7", "Do's & Don'ts", "5 Min"),
        ("8", "Practical Demonstration", "5 Min"),
        ("9", "Questions & Feedback", "5 Min"),
    ]
    for i, row in enumerate(agenda_rows, start=1):
        for j, val in enumerate(row):
            set_cell_text(agenda.rows[i].cells[j], val, size=11, center=(j != 1))
            if i % 2 == 0:
                shade_cell(agenda.rows[i].cells[j], "F4F6F8")

    add_heading_styled(doc, "2. Key Training Objectives", 1)
    objectives = [
        "Prevent damage to fragile materials and finished works.",
        "Prevent injuries during lifting, shifting and transportation.",
        "Ensure proper storage and protection of fragile items.",
        "Reduce material wastage and replacement costs.",
        "Make the Civil Team responsible for safe material handling.",
        "Maintain good housekeeping and organized storage areas.",
    ]
    for o in objectives:
        p = doc.add_paragraph(o, style="List Bullet")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)

    add_heading_styled(doc, "3. What are Fragile Items?", 1)
    p = doc.add_paragraph(
        "Fragile items are materials that can be easily damaged by impact, pressure, "
        "improper lifting, vibration, stacking, moisture or rough handling. Damaged items "
        "often cannot be repaired and must be replaced — causing delay, extra cost and safety risk."
    )
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = DocPt(11)

    p = doc.add_paragraph()
    r = p.add_run("[ Image Placeholder — Add photo of fragile materials at site ]")
    r.italic = True
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    add_heading_styled(doc, "4. Identification of Fragile Materials at Site", 1)
    id_items = [
        "Glass and glass panels",
        "Tiles and marble",
        "Wash basins and sanitary fixtures",
        "Ceramic items",
        "Doors and door frames",
        "False ceiling materials",
        "Lighting fixtures",
        "Finished civil surfaces",
        "ACP / decorative panels",
        "Mirrors",
        "Pre-finished materials",
        "Any material marked FRAGILE / HANDLE WITH CARE",
    ]
    for item in id_items:
        p = doc.add_paragraph(item, style="List Bullet")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)

    p = doc.add_paragraph()
    r = p.add_run("[ Image Placeholder — Collage: glass / tiles / sanitary / panels ]")
    r.italic = True
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    add_heading_styled(doc, "5. Safe Handling & Lifting Methods", 1)
    handling = [
        "Check the material condition before moving it.",
        "Use the correct number of workers for the item's weight and size.",
        "Wear required PPE.",
        "Use proper lifting posture.",
        "Do not drag fragile materials on the floor.",
        "Do not throw, drop or slide materials.",
        "Hold materials firmly from designated/supporting points.",
        "Use suitable trolleys or lifting equipment wherever required.",
    ]
    for item in handling:
        p = doc.add_paragraph(item, style="List Bullet")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)

    p = doc.add_paragraph()
    r = p.add_run("[ Image Placeholder — Correct lifting / PPE demonstration ]")
    r.italic = True
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    add_heading_styled(doc, "6. Transportation & Movement", 1)
    transport = [
        "Plan the movement route before shifting.",
        "Keep the route clear of obstacles.",
        "Secure materials during transportation.",
        "Do not overload trolleys.",
        "Avoid sudden movements and impacts.",
        "Protect corners and finished surfaces.",
        "Never place heavy materials over fragile items.",
    ]
    for item in transport:
        p = doc.add_paragraph(item, style="List Bullet")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)

    p = doc.add_paragraph()
    r = p.add_run("[ Image Placeholder — Trolley movement on site ]")
    r.italic = True
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    add_heading_styled(doc, "7. Storage & Stacking Requirements", 1)
    storage = [
        "Store materials in designated areas.",
        "Keep fragile items on stable and level surfaces.",
        "Use racks, pallets, supports or protective frames where required.",
        "Avoid unstable stacking.",
        "Keep heavy materials at the bottom.",
        "Protect materials from water, dust and construction activities.",
        "Maintain clear access around stored materials.",
        "Follow manufacturer's storage instructions.",
    ]
    for item in storage:
        p = doc.add_paragraph(item, style="List Bullet")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)

    p = doc.add_paragraph()
    r = p.add_run("[ Image Placeholder — Correct rack / pallet storage ]")
    r.italic = True
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    add_heading_styled(doc, "8. Do's", 1)
    dos = [
        "Inspect before handling.",
        "Use proper PPE.",
        "Use proper tools/equipment.",
        "Follow lifting instructions.",
        "Use sufficient manpower.",
        "Communicate clearly during movement.",
        "Protect materials after installation.",
        "Report any damage immediately.",
    ]
    for item in dos:
        p = doc.add_paragraph(f"✓  {item}")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)
            run.font.color.rgb = DocRGBColor(0x2E, 0x7D, 0x4F)

    add_heading_styled(doc, "9. Don'ts", 1)
    donts = [
        "Don't throw or drop materials.",
        "Don't drag fragile items.",
        "Don't stand fragile items in an unstable position.",
        "Don't stack heavy materials on fragile materials.",
        "Don't use damaged lifting equipment.",
        "Don't carry oversized materials without assistance.",
        "Don't leave fragile materials exposed to unnecessary site activity.",
        "Don't continue using damaged material without inspection.",
    ]
    for item in donts:
        p = doc.add_paragraph(f"✗  {item}")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)
            run.font.color.rgb = DocRGBColor(0xA8, 0x3A, 0x32)

    add_heading_styled(doc, "10. Practical Demonstration Checklist", 1)
    p = doc.add_paragraph("The trainer should demonstrate:")
    for run in p.runs:
        run.font.name = "Calibri"
    demos = [
        "Correct lifting of a fragile item.",
        "Correct use of a trolley.",
        "Safe movement through a work area.",
        "Correct placement on a pallet/rack.",
        "Proper protection using sheets, foam, cardboard or other suitable protection.",
        "Correct method of reporting damaged material.",
    ]
    for i, d in enumerate(demos, 1):
        p = doc.add_paragraph(f"{i}.  {d}     ☐ Demonstrated")
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = DocPt(11)

    p = doc.add_paragraph()
    r = p.add_run("[ Image Placeholder — Live demonstration photo ]")
    r.italic = True
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    add_heading_styled(doc, "11. Closing Message", 1)
    quote = doc.add_paragraph()
    quote.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = quote.add_run(
        '"Handle every fragile item as if it is already installed. '
        'One careless movement can cause material damage, rework, additional cost and safety risks."'
    )
    r.bold = True
    r.italic = True
    r.font.size = DocPt(13)
    r.font.color.rgb = DocRGBColor(0x0F, 0x2A, 0x44)
    r.font.name = "Calibri"

    add_heading_styled(doc, "12. Attendance & Feedback (for trainer use)", 1)
    att = doc.add_table(rows=8, cols=4)
    att.style = "Table Grid"
    for i, h in enumerate(["S.No.", "Name", "Designation / Trade", "Signature"]):
        set_cell_text(att.rows[0].cells[i], h, bold=True, color="FFFFFF", size=10, center=True)
        shade_cell(att.rows[0].cells[i], "0F2A44")
    for i in range(1, 8):
        set_cell_text(att.rows[i].cells[0], str(i), center=True, size=10)

    doc.add_paragraph()
    fb = doc.add_paragraph()
    r = fb.add_run("Trainer Name: ______________________    Date: ______________    Location: ____________________")
    r.font.name = "Calibri"
    r.font.size = DocPt(11)

    doc.add_paragraph()
    end = doc.add_paragraph()
    end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = end.add_run("— End of Training Pack —")
    r.font.color.rgb = DocRGBColor(0x5A, 0x6A, 0x7A)
    r.font.size = DocPt(10)

    doc.save(DOC_PATH)
    print("Saved", DOC_PATH)


if __name__ == "__main__":
    build_ppt()
    build_docx()
