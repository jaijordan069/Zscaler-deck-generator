# generate_ppt.py
"""
Optimized PPT generator using an existing Zscaler template.

Usage (standalone):
    python generate_ppt.py

Or with Streamlit (if you want UI):
    streamlit run generate_ppt.py

Place your template file at the default path or pass a custom path to `TEMPLATE_PATH`.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE
import io
import os
import requests

# ---------- CONFIG ----------
TEMPLATE_PATH = "/mnt/data/TEMPLATE - PS Project Transition Slides.pptx"
OUTPUT_PATH = "OUTPUT_Transition_Deck.pptx"

# ---------- Helper utilities ----------
def load_presentation(template_path=TEMPLATE_PATH):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found at: {template_path}")
    return Presentation(template_path)

def safe_layout(prs, idx, fallback=0):
    """Return a slide_layout from prs safely (falls back to layout 0)."""
    try:
        return prs.slide_layouts[idx]
    except Exception:
        return prs.slide_layouts[fallback]

def add_footer_and_page_number(slide, page_num=None, footer_text=None):
    """
    Add/update a footer and page number in the bottom area.
    If the template already contains a footer placeholder, prefer that.
    """
    # Footer text (non-intrusive)
    if footer_text:
        try:
            footer_box = slide.shapes.add_textbox(Inches(0.3), Inches(6.6), Inches(8.5), Inches(0.3))
            tf = footer_box.text_frame
            tf.text = footer_text
            p = tf.paragraphs[0]
            p.font.size = Pt(8)
            p.font.color.rgb = RGBColor(128, 128, 128)
            p.alignment = PP_ALIGN.LEFT
        except Exception:
            pass

    if page_num is not None:
        try:
            num_box = slide.shapes.add_textbox(Inches(9.3), Inches(6.55), Inches(0.7), Inches(0.3))
            ntf = num_box.text_frame
            ntf.text = str(page_num)
            p = ntf.paragraphs[0]
            p.font.size = Pt(11)
            p.font.color.rgb = RGBColor(128, 128, 128)
            p.alignment = PP_ALIGN.RIGHT
        except Exception:
            pass

# ---------- Content helper functions ----------
def add_title_slide(prs, title_text, subtitle_text=None, layout_idx=0, page_num=None):
    layout = safe_layout(prs, layout_idx)
    slide = prs.slides.add_slide(layout)
    # Try to set title placeholder if exists
    try:
        if slide.shapes.title:
            slide.shapes.title.text = title_text
            slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(44)
    except Exception:
        # fallback: manual textbox
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(8.5), Inches(1.5))
        tf = tb.text_frame
        tf.text = title_text
        tf.paragraphs[0].font.size = Pt(44)
        tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    if subtitle_text:
        # try placeholder index 1
        try:
            placeholder = slide.placeholders[1]
            placeholder.text = subtitle_text
            placeholder.text_frame.paragraphs[0].font.size = Pt(28)
        except Exception:
            sub_tb = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(8.5), Inches(1.0))
            stf = sub_tb.text_frame
            stf.text = subtitle_text
            stf.paragraphs[0].font.size = Pt(28)
            stf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

    add_footer_and_page_number(slide, page_num=page_num, footer_text="Zscaler, Inc. All rights reserved. © 2025")
    return slide

def add_agenda(prs, bullets, layout_idx=1, page_num=None):
    layout = safe_layout(prs, layout_idx)
    slide = prs.slides.add_slide(layout)
    # title
    try:
        slide.shapes.title.text = "Meeting Agenda"
        slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(28)
    except Exception:
        pass

    # content placeholder
    try:
        content = slide.placeholders[1]
        tf = content.text_frame
        tf.clear()
        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"• {b}"
            p.level = 0
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(255, 255, 255)
    except Exception:
        # fallback: manual textbox
        tb = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(4))
        tf = tb.text_frame
        tf.clear()
        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"• {b}"
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(255, 255, 255)

    add_footer_and_page_number(slide, page_num=page_num)
    return slide

def add_table_slide(prs, title, headers, rows, layout_idx=5, page_num=None):
    """
    Add a table slide. headers: list of header strings. rows: list of lists.
    """
    layout = safe_layout(prs, layout_idx)
    slide = prs.slides.add_slide(layout)

    # Title
    try:
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(9), Inches(0.5))
        tf = title_box.text_frame
        tf.text = title
        tf.paragraphs[0].font.size = Pt(28)
        tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    except Exception:
        pass

    rows_count = max(1, len(rows)) + 1
    cols_count = max(1, len(headers))
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9.0)
    height = Inches(4.0)

    table = slide.shapes.add_table(rows_count, cols_count, left, top, width, height).table

    # Headers
    for c_idx, h in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.text = h
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(12)
        cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0, 102, 204)

    # Data rows
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(value)
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11)
            p.font.color.rgb = RGBColor(0, 0, 0)
        # alternate background for readability
        if r_idx % 2 == 1:
            for c in range(cols_count):
                cell = table.cell(r_idx, c)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(242, 242, 242)

    add_footer_and_page_number(slide, page_num=page_num)
    return slide

def add_deliverables_with_checks(prs, title, deliverables, layout_idx=5, start_top_inch=1.5, page_num=None):
    """
    deliverables: list of tuples/lists: (name, date)
    Adds a table and draws a small green oval with a white tick for each delivered item.
    """
    headers = ["Deliverable", "Date delivered"]
    rows = [[d[0], d[1]] for d in deliverables]
    slide = add_table_slide(prs, title, headers, rows, layout_idx=layout_idx, page_num=page_num)

    # Add check marks as small green ovals with ✔ to the left of the table rows
    base_left = Inches(0.25)
    row_height = 0.32  # inches per row - approximate; adjust if your template uses different row heights
    for i, _ in enumerate(rows, start=1):
        top = Inches(start_top_inch) + Inches(row_height) * (i - 1)
        try:
            check_shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                base_left,
                top,
                Inches(0.26),
                Inches(0.26)
            )
            check_shape.fill.solid()
            check_shape.fill.fore_color.rgb = RGBColor(0, 176, 80)
            # Add white tick text centered
            tf = check_shape.text_frame
            tf.text = "✔"
            p = tf.paragraphs[0]
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(255, 255, 255)
            p.alignment = PP_ALIGN.CENTER
        except Exception:
            # if shapes can't be added, continue gracefully
            continue

    return slide

def add_zia_diagram(prs, layout_idx=5, page_num=None):
    """
    Adds a simple ZIA architecture diagram using shapes.
    Keep it generic — the more complex/professional diagram should live in the design doc.
    """
    layout = safe_layout(prs, layout_idx)
    slide = prs.slides.add_slide(layout)

    # Title
    try:
        tbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(9), Inches(0.5))
        tbox_tf = tbox.text_frame
        tbox_tf.text = "Deployed ZIA Architecture"
        tbox_tf.paragraphs[0].font.size = Pt(28)
        tbox_tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    except Exception:
        pass

    # Central Authority
    central = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.0), Inches(2.0), Inches(2.0), Inches(0.9))
    central.fill.solid()
    central.fill.fore_color.rgb = RGBColor(0, 102, 204)
    ct = central.text_frame
    ct.text = "Central Authority"
    ct.paragraphs[0].font.size = Pt(12)
    ct.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    ct.paragraphs[0].alignment = PP_ALIGN.CENTER

    tunnels = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0), Inches(2.0), Inches(2.0), Inches(0.9))
    tunnels.fill.solid()
    tunnels.fill.fore_color.rgb = RGBColor(0, 102, 204)
    tt = tunnels.text_frame
    tt.text = "Z-Tunnels"
    tt.paragraphs[0].font.size = Pt(12)
    tt.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    tt.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Connector arrow
    connector = slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT, Inches(4.0), Inches(2.45), Inches(5.0), Inches(2.45))
    try:
        connector.line.color.rgb = RGBColor(0, 102, 204)
        connector.line.width = Pt(2)
    except Exception:
        pass

    add_footer_and_page_number(slide, page_num=page_num)
    return slide

# ---------- Main generator example ----------
def build_deck_from_data(data, template_path=TEMPLATE_PATH, output_path=OUTPUT_PATH):
    """
    data is a dict with keys like:
    {
      'customer_name': 'Pixartprinting',
      'today_date': '14/11/2025',
      'milestones': [...],
      'deliverables': [('Doc', '01/11/2025'), ...],
      'short_term': [...],
      'long_term': [...],
      'open_items': [...],
      'contacts': {'pm':..., 'consultant':..., 'primary':..., 'secondary':...}
    }
    """
    prs = load_presentation(template_path)
    page = 1

    # Title slide
    title = f"Professional Services Transition Meeting - {data.get('customer_name', '')}"
    subtitle = data.get('today_date', '')
    add_title_slide(prs, title, subtitle, layout_idx=0, page_num=page)
    page += 1

    # Agenda
    agenda_items = data.get('agenda', ["Project Summary", "Technical Summary", "Recommended Next Steps"])
    add_agenda(prs, agenda_items, layout_idx=1, page_num=page)
    page += 1

    # Project summary (title)
    add_title_slide(prs, "Project Summary", "", layout_idx=0, page_num=page)
    page += 1

    # Project dates table
    dates_headers = ["Today's Date", "Start Date", "End Date"]
    date_row = [[data.get('today_date', ''), data.get('start_date', ''), data.get('end_date', '')]]
    add_table_slide(prs, f"Final Project Status Report – {data.get('customer_name','')}", dates_headers, date_row, layout_idx=5, page_num=page)
    page += 1

    # Deliverables
    add_deliverables_with_checks(prs, "Deliverables", data.get('deliverables', []), layout_idx=5, page_num=page)
    page += 1

    # Technical summary (title + diagram)
    add_title_slide(prs, "Technical Summary", "", layout_idx=0, page_num=page)
    page += 1
    add_zia_diagram(prs, layout_idx=5, page_num=page)
    page += 1

    # Open items table
    open_items = data.get('open_items', [])
    open_headers = ["Task/ Description", "Date", "Owner", "Transition Plan/ Next Steps"]
    open_rows = [[oi.get('task',''), oi.get('date',''), oi.get('owner',''), oi.get('steps','')] for oi in open_items]
    add_table_slide(prs, "Open Items", open_headers, open_rows, layout_idx=5, page_num=page)
    page += 1

    # Recommended Next Steps
    next_slide = add_title_slide(prs, "Recommended Next Steps", "", layout_idx=0, page_num=page)
    # Create two-column boxes for short/long term using shapes:
    short = data.get('short_term', [])
    longt = data.get('long_term', [])
    # Short box
    sbox = next_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.6), Inches(4.5), Inches(3.8))
    sbox.fill.solid()
    sbox.fill.fore_color.rgb = RGBColor(0, 176, 80)
    stf = sbox.text_frame
    stf.text = "Short Term Activities"
    stf.paragraphs[0].font.size = Pt(16)
    stf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    for it in short:
        p = stf.add_paragraph()
        p.text = "• " + it
        p.font.size = Pt(12)
        p.level = 0

    # Long box
    lbox = next_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.0), Inches(1.6), Inches(4.5), Inches(3.8))
    lbox.fill.solid()
    lbox.fill.fore_color.rgb = RGBColor(0, 102, 204)
    ltf = lbox.text_frame
    ltf.text = "Long Term Activities"
    ltf.paragraphs[0].font.size = Pt(16)
    ltf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
    for it in longt:
        p = ltf.add_paragraph()
        p.text = "• " + it
        p.font.size = Pt(12)

    add_footer_and_page_number(next_slide, page_num=page)
    page += 1

    # Thank you slide (use template's thank you layout if present)
    add_title_slide(prs, "Thank you", "", layout_idx=1, page_num=page)
    page += 1

    # Save
    prs.save(output_path)
    return output_path

# ---------- Example usage ----------
if __name__ == "__main__":
    # Example data you can customize or pass programmatically or via a small UI.
    demo_data = {
        "customer_name": "Pixartprinting",
        "today_date": "14/11/2025",
        "start_date": "01/06/2025",
        "end_date": "14/11/2025",
        "deliverables": [
            ("Design Document", "01/09/2025"),
            ("Deployment Checklist", "05/10/2025"),
            ("User Guide", "10/11/2025")
        ],
        "open_items": [
            {"task": "Finalize RBAC", "date": "20/11/2025", "owner": "IT Team", "steps": "Create roles and test"},
            {"task": "DLP Tune", "date": "22/11/2025", "owner": "Security Team", "steps": "Review policies"}
        ],
        "short_term": ["Finish Production rollout", "Tighten Firewall policies"],
        "long_term": ["Consider ZPA adoption", "Upgrade Sandbox license"]
    }

    out = build_deck_from_data(demo_data, template_path=TEMPLATE_PATH, output_path=OUTPUT_PATH)
    print(f"Generated PPTX: {out}")
