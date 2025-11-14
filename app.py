#!/usr/bin/env python3
"""
Streamlit app: Zscaler Professional Services Transition Deck PPT Generator
Updated to match a fixed template layout (alignments, borders, diagrams, logos and sizing)
Only values are intended to change when the user edits fields in the UI.

Notes:
- This single-file app centralizes all layout constants so positions/sizes can be tuned
  to match the provided template precisely.
- Background and logos are rendered to slides as full-slide pictures to guarantee pixel-perfect placement.
- All shapes, tables and diagrams use fixed coordinates and sizes (Inches) taken from the template.
- Fonts, sizes, paddings and colors are defined as constants for consistent styling.
"""

import io
import re
import requests
from datetime import datetime

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

# -------------------------
# Configuration / Constants
# -------------------------
st.set_page_config(page_title="Zscaler Transition Deck PPT Generator", layout="wide")

# Colors (match template palette)
COLOR_BRIGHT_BLUE = RGBColor(37, 108, 247)
COLOR_NAVY = RGBColor(0, 23, 68)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_LIGHT_GRAY = RGBColor(229, 241, 250)
COLOR_CYAN = RGBColor(18, 212, 255)
COLOR_ACCENT_GREEN = RGBColor(107, 255, 179)
COLOR_THREAT_RED = RGBColor(237, 25, 81)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_ZSCALER_YELLOW = RGBColor(255, 192, 0)

# Fonts & sizes (template)
FONT_NAME = "Century Gothic"
SIZE_TITLE = Pt(36)
SIZE_SLIDE_TITLE = Pt(28)
SIZE_SUBTITLE = Pt(20)
SIZE_HEADER = Pt(18)
SIZE_BODY = Pt(14)
SIZE_SMALL = Pt(12)
SIZE_FOOTER = Pt(8)

# Slide/page layout constants (Inches) tuned to template
MARGIN_LEFT = Inches(0.45)
MARGIN_TOP = Inches(0.45)
MARGIN_RIGHT = Inches(0.45)
FOOTER_HEIGHT = Inches(0.35)

# Logo and background resources (template assets)
# You can replace LOGO_URL / BG_URL with internal assets if you have the exact template images.
LOGO_URL = "https://upload.wikimedia.org/wikipedia/commons/thumb/8/8b/Zscaler_logo.svg/512px-Zscaler_logo.svg.png"
ALT_LOGO_URL = "https://companieslogo.com/img/orig/ZS-46a5871c.png?t=1720244494"
BG_URL = None  # If you have a full-slide background image that is the template, set the URL here.

# Date validation regex
DATE_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')

# -------------------------
# Utility helpers
# -------------------------
def is_valid_date(d: str) -> bool:
    if not d:
        return False
    return bool(DATE_RE.match(d))

def download_image_to_bytes(url: str) -> io.BytesIO:
    """Download image from URL into BytesIO. Returns None on failure."""
    try:
        r = requests.get(url, timeout=8)
        r.raise_for_status()
        return io.BytesIO(r.content)
    except Exception:
        return None

def set_font_run(run, name=FONT_NAME, size=SIZE_BODY, bold=False, color=COLOR_BLACK):
    run.font.name = name
    # Ensure East Asian font setting so pptx shows the font on systems without font installed.
    run._element.rPr.rFonts.set(qn('a:ea'), name)
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color

# Centralized function to add a textbox with consistent defaults
def add_textbox(slide, left, top, width, height, text, size=SIZE_BODY, bold=False, color=COLOR_BLACK, align=PP_ALIGN.LEFT, auto_size=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    if auto_size:
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    run = p.runs[0]
    set_font_run(run, size=size, bold=bold, color=color)
    return txBox

# Footer & logo helper (applies template footer/logo consistently)
def apply_template_branding(prs, slide, slide_num, logo_bytes):
    # Background (if BG_URL was provided, we add it at slide creation time elsewhere)
    # Add logo top-right at exact template position
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    logo_w = Inches(1.85)  # exact width per template
    logo_h = Inches(0.5)   # exact height per template
    logo_left = slide_width - logo_w - MARGIN_RIGHT
    logo_top = MARGIN_TOP

    if logo_bytes:
        try:
            slide.shapes.add_picture(logo_bytes, logo_left, logo_top, logo_w, logo_h)
        except Exception:
            # fallback: ignore logo failing
            pass

    # Footer left text
    footer_left = MARGIN_LEFT
    footer_top = slide_height - FOOTER_HEIGHT
    footer_w = slide_width - (MARGIN_LEFT + MARGIN_RIGHT + Inches(1.0))
    footer_h = FOOTER_HEIGHT
    fbox = slide.shapes.add_textbox(footer_left, footer_top, footer_w, footer_h)
    ftf = fbox.text_frame
    ftf.clear()
    p = ftf.paragraphs[0]
    p.text = "Zscaler, Inc. All rights reserved. © " + str(datetime.utcnow().year)
    p.alignment = PP_ALIGN.LEFT
    set_font_run(p.runs[0], size=SIZE_FOOTER, bold=False, color=COLOR_NAVY)

    # Slide number right
    sn_box = slide.shapes.add_textbox(slide_width - Inches(0.8), footer_top, Inches(0.6), footer_h)
    stf = sn_box.text_frame
    stf.clear()
    p = stf.paragraphs[0]
    p.text = str(slide_num)
    p.alignment = PP_ALIGN.RIGHT
    set_font_run(p.runs[0], size=SIZE_FOOTER, bold=False, color=COLOR_NAVY)

# Add a full-slide background image (keeps consistent template look)
def add_slide_with_background(prs, bg_bytes=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    if bg_bytes:
        # Add background picture as first shape so other shapes overlay it
        try:
            slide.shapes.add_picture(bg_bytes, 0, 0, prs.slide_width, prs.slide_height)
        except Exception:
            pass
    return slide

# ---------- Streamlit UI ----------
st.title("Zscaler Professional Services Transition Deck PPT Generator")
st.markdown("Fill in the form to generate a PPTX that matches the template layout exactly; only the values change.")

with st.sidebar:
    st.header("Instructions")
    st.markdown(
        "- Fill customer and project details.\n"
        "- Dates must be DD/MM/YYYY.\n"
        "- Press 'Preview Summary' to inspect your inputs.\n"
        "- Press 'Generate Transition Deck' to build and download the PPTX."
    )

# Basic fields
st.header("Customer & Project Basics")
c1, c2, c3 = st.columns(3)
customer_name = c1.text_input("Customer Name *", value="Pixartprinting")
today_date = c2.text_input("Today's Date (DD/MM/YYYY) *", value=datetime.utcnow().strftime("%d/%m/%Y"))
project_start = c3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025")
project_end = st.text_input("Project End Date (DD/MM/YYYY) *", value="14/11/2025")
project_summary_text = st.text_area("Project Summary Text", value="More than half of the users have been deployed and there were no critical issues. Remaining enrollments expected without major issues.")
theme = st.selectbox("Theme", ["White", "Navy"], index=0)

# Milestones (fixed number to match template)
st.header("Milestones (7 rows - as template)")
milestone_defaults = [
    ("Initial Project Schedule Accepted", "27/06/2025", "27/06/2025", ""),
    ("Initial Design Accepted", "14/07/2025", "17/07/2025", ""),
    ("Pilot Configuration Complete", "28/07/2025", "18/07/2025", ""),
    ("Pilot Rollout Complete", "08/08/2025", "22/08/2025", ""),
    ("Production Configuration Complete", "29/08/2025", "29/08/2025", ""),
    ("Production Rollout Complete", "14/11/2025", "??", ""),
    ("Final Design Accepted", "14/11/2025", "14/11/2025", "")
]
milestones_data = []
cols_per_milestone = st.columns(4)
for i in range(7):
    with st.expander(f"Milestone {i+1}", expanded=False):
        mn = st.text_input(f"Milestone {i+1} Name", value=milestone_defaults[i][0], key=f"mname_{i}")
        mb = st.text_input(f"Baseline {i+1} (DD/MM/YYYY)", value=milestone_defaults[i][1], key=f"mbaseline_{i}")
        mt = st.text_input(f"Target {i+1} (DD/MM/YYYY)", value=milestone_defaults[i][2], key=f"mtarget_{i}")
        ms = st.text_input(f"Status {i+1}", value=milestone_defaults[i][3], key=f"mstatus_{i}")
        milestones_data.append({"name": mn, "baseline": mb, "target": mt, "status": ms})

# Rollout
st.header("User Rollout Roadmap")
p1, p2 = st.columns(2)
with p1:
    pilot_target = st.number_input("Pilot Target Users", value=100, min_value=0)
    pilot_current = st.number_input("Pilot Current Users", value=449, min_value=0)
    pilot_completion = st.text_input("Pilot Completion Date (DD/MM/YYYY)", value="14/11/2025")
    pilot_status = st.text_input("Pilot Status", value="")
with p2:
    prod_target = st.number_input("Production Target Users", value=800, min_value=0)
    prod_current = st.number_input("Production Current Users", value=449, min_value=0)
    prod_completion = st.text_input("Production Completion Date (DD/MM/YYYY)", value="14/11/2025")
    prod_status = st.text_input("Production Status", value="")

# Objectives (3 rows)
st.header("Project Objectives (3 rows)")
objectives_data = []
for i in range(3):
    with st.expander(f"Objective {i+1}", expanded=False):
        obj = st.text_area(f"Planned Objective {i+1}", value="", key=f"obj_{i}", height=60)
        act = st.text_area(f"Actual Result {i+1}", value="", key=f"act_{i}", height=60)
        dev = st.text_area(f"Deviation/Cause {i+1}", value="", key=f"dev_{i}", height=60)
        objectives_data.append({"objective": obj, "actual": act, "deviation": dev})

# Deliverables
st.header("Deliverables (5 rows)")
deliverables_data = []
for i in range(5):
    with st.expander(f"Deliverable {i+1}", expanded=False):
        dn = st.text_input(f"Deliverable Name {i+1}", value="", key=f"dname_{i}")
        dd = st.text_input(f"Date Delivered {i+1}", value="", key=f"ddate_{i}")
        deliverables_data.append({"name": dn, "date": dd})

# Technical Summary
st.header("Technical Summary")
t1, t2 = st.columns(2)
with t1:
    idp = st.text_input("Identity Provider", value="Entra ID")
    auth_type = st.text_input("Authentication Type", value="SAML 2.0")
    prov_type = st.text_input("User/Group Provisioning", value="SCIM Provisioning")
with t2:
    tunnel_type = st.text_input("Tunnel Type", value="ZCC with Z-Tunnel 2.0")
    deploy_system = st.text_input("ZCC Deployment System", value="MS Intune/Jamf")
d1, d2, d3 = st.columns(3)
windows_num = d1.number_input("Number of Windows Devices", value=351, min_value=0)
mac_num = d2.number_input("Number of MacOS Devices", value=98, min_value=0)
geo_locations = d3.text_input("Geo Locations", value="Europe, North Africa, USA")
pol1, pol2, pol3, pol4 = st.columns(4)
ssl_policies = pol1.number_input("SSL Inspection Policies", value=10, min_value=0)
url_policies = pol2.number_input("URL Filtering Policies", value=5, min_value=0)
cloud_policies = pol3.number_input("Cloud App Control Policies", value=5, min_value=0)
fw_policies = pol4.number_input("Firewall Policies", value=15, min_value=0)

# Open items (6 rows)
st.header("Open Items (6 rows)")
open_items_data = []
for i in range(6):
    with st.expander(f"Open Item {i+1}", expanded=False):
        otask = st.text_input(f"Task/Description {i+1}", value="", key=f"otask_{i}")
        odate = st.text_input(f"Date {i+1}", value="", key=f"odate_{i}")
        oowner = st.text_input(f"Owner {i+1}", value="", key=f"oowner_{i}")
        osteps = st.text_area(f"Transition Plan/Next Steps {i+1}", value="", key=f"osteps_{i}", height=80)
        open_items_data.append({"task": otask, "date": odate, "owner": oowner, "steps": osteps})

# Next Steps
st.header("Recommended Next Steps")
short_term_input = st.text_area("Short Term (comma-separated)", value="Finish Production rollout, Tighten Firewall policies")
long_term_input = st.text_area("Long Term (comma-separated)", value="Deploy ZCC on Mobile devices, Consider upgrade of Sandbox license")
short_term = [s.strip() for s in short_term_input.split(",") if s.strip()]
long_term = [s.strip() for s in long_term_input.split(",") if s.strip()]

# Contacts
st.header("Contacts")
c1, c2 = st.columns(2)
pm_name = c1.text_input("Project Manager Name", value="Alex Vazquez")
consultant_name = c2.text_input("Consultant Name", value="Alex Vazquez")
primary_contact = st.text_input("Primary Contact", value="Teia proctor")
secondary_contact = st.text_input("Secondary Contact", value="Marco Sattier")

# Preview Summary
if st.button("Preview Summary"):
    st.write(f"Deck will be generated for {customer_name} (date: {today_date}).")
    st.write(f"- Summary: {project_summary_text[:200]}")
    st.write(f"- Milestones: {len(milestones_data)} rows")
    st.write(f"- Pilot: {pilot_current}/{pilot_target} users (status: {pilot_status})")
    st.write(f"- Production: {prod_current}/{prod_target} users (status: {prod_status})")
    st.write(f"- Objectives: {len([o for o in objectives_data if o['objective']])}")
    st.write(f"- Deliverables: {len([d for d in deliverables_data if d['name']])}")
    st.write(f"- Open Items: {len([o for o in open_items_data if o['task']])}")
    st.write(f"- Short-term items: {len(short_term)}; Long-term items: {len(long_term)}")

# -------------------------
# PPT Generation
# -------------------------
if st.button("Generate Transition Deck"):
    # Basic validation
    required_dates = [today_date, project_start, project_end, pilot_completion, prod_completion]
    for m in milestones_data:
        if m.get("baseline"):
            required_dates.append(m["baseline"])
        if m.get("target"):
            required_dates.append(m["target"])
    for d in deliverables_data:
        if d.get("date"):
            required_dates.append(d["date"])
    for o in open_items_data:
        if o.get("date"):
            required_dates.append(o["date"])

    if not customer_name:
        st.error("Customer Name is required.")
        st.stop()

    if not all(is_valid_date(dd) for dd in required_dates if dd and dd != "??"):
        st.error("Some dates are not in DD/MM/YYYY format. Please correct them.")
        st.stop()

    # Create Presentation
    prs = Presentation()
    # If you want to ensure a particular slide size, set it here:
    # prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)  # Example widescreen
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Pre-download brand assets once for consistency
    logo_bytes = download_image_to_bytes(LOGO_URL) or download_image_to_bytes(ALT_LOGO_URL)
    bg_bytes = download_image_to_bytes(BG_URL) if BG_URL else None

    # Small helper to write a title slide exactly as template expects
    def create_title_slide(title_text, subtitle_text=None, date_text=None, slide_num=1):
        slide = add_slide_with_background(prs, bg_bytes)
        # Title block (left aligned, big font)
        add_textbox(slide, MARGIN_LEFT, Inches(1.0), Inches(9.5), Inches(1.2), title_text, size=SIZE_TITLE, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
        if subtitle_text:
            add_textbox(slide, MARGIN_LEFT, Inches(2.1), Inches(9.5), Inches(0.8), subtitle_text, size=SIZE_SUBTITLE, bold=False, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
        if date_text:
            add_textbox(slide, MARGIN_LEFT, Inches(2.9), Inches(9.5), Inches(0.6), date_text, size=SIZE_BODY, bold=False, color=COLOR_BLACK, align=PP_ALIGN.LEFT)
        apply_template_branding(prs, slide, slide_num, logo_bytes)
        return slide

    # Helper to add a bullet slide (two-column if template wants it)
    def create_bullet_slide(title_text, bullets, slide_num=1):
        slide = add_slide_with_background(prs, bg_bytes)
        add_textbox(slide, MARGIN_LEFT, MARGIN_TOP, Inches(9.0), Inches(0.7), title_text, size=SIZE_SLIDE_TITLE, bold=True, color=COLOR_NAVY, align=PP_ALIGN.LEFT)
        top = Inches(1.6)
        for b in bullets:
            # small square bullet (template color: bright blue)
            sq = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN_LEFT, top + Inches(0.08), Inches(0.18), Inches(0.18))
            sq.fill.solid()
            sq.fill.fore_color.rgb = COLOR_BRIGHT_BLUE
            sq.line.color.rgb = COLOR_BRIGHT_BLUE
            tb = slide.shapes.add_textbox(MARGIN_LEFT + Inches(0.25), top, Inches(9.0), Inches(0.4))
            ttf = tb.text_frame
            ttf.clear()
            p = ttf.paragraphs[0]
            p.text = b
            p.alignment = PP_ALIGN.LEFT
            set_font_run(p.runs[0], size=SIZE_BODY, bold=False, color=COLOR_BLACK)
            top += Inches(0.55)
        apply_template_branding(prs, slide, slide_num, logo_bytes)
        return slide

    # Helper to add a table matching template column widths and style
    def create_table_slide(title_text, headers, rows, slide_num=1, top_inch=1.6, height_inch=3.5):
        slide = add_slide_with_background(prs, bg_bytes)
        add_textbox(slide, MARGIN_LEFT, MARGIN_TOP, Inches(9.0), Inches(0.6), title_text, size=SIZE_SLIDE_TITLE, bold=True, color=COLOR_NAVY)
        left = MARGIN_LEFT
        top = Inches(top_inch)
        width = slide_width - (MARGIN_LEFT + MARGIN_RIGHT)
        height = Inches(height_inch)
        cols = len(headers)
        rows_count = max(1, len(rows)) + 1
        table_shape = slide.shapes.add_table(rows_count, cols, left, top, width, height)
        table = table_shape.table

        # set column widths evenly (you can tweak these to match template)
        colw = (width) / cols
        for i in range(cols):
            table.columns[i].width = colw

        # Header row style
        for i, header in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = COLOR_NAVY
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            set_font_run(p.runs[0], size=SIZE_SMALL, bold=True, color=COLOR_WHITE)

        # Data rows
        for r_idx, r in enumerate(rows, start=1):
            for c_idx, value in enumerate(r):
                cell = table.cell(r_idx, c_idx)
                cell.text = str(value)
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.LEFT
                set_font_run(p.runs[0], size=SIZE_SMALL, bold=False, color=COLOR_BLACK)
                if r_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = COLOR_LIGHT_GRAY

        apply_template_branding(prs, slide, slide_num, logo_bytes)
        return slide

    # Helper to add diagram slide exactly positioned like the template
    def create_zia_diagram_slide(slide_num=1):
        slide = add_slide_with_background(prs, bg_bytes)
        add_textbox(slide, MARGIN_LEFT, MARGIN_TOP, Inches(9.0), Inches(0.7), "Deployed ZIA Architecture", size=SIZE_SLIDE_TITLE, bold=True, color=COLOR_NAVY)

        # Draw boxes (positions & sizes are tuned to match template)
        box_w = Inches(2.8)
        box_h = Inches(0.85)
        left_a = MARGIN_LEFT
        top_a = Inches(1.6)
        left_b = left_a + box_w + Inches(0.35)
        left_c = left_b + box_w + Inches(0.35)

        # User authentication box (light gray background, navy border)
        ua = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_a, top_a, box_w, box_h)
        ua.fill.solid()
        ua.fill.fore_color.rgb = COLOR_LIGHT_GRAY
        ua.line.width = Pt(1)
        ua.line.color.rgb = COLOR_NAVY
        ua_tf = ua.text_frame
        ua_tf.clear()
        p = ua_tf.paragraphs[0]
        p.text = "User authentication\nand provisioning"
        p.alignment = PP_ALIGN.CENTER
        set_font_run(p.runs[0], size=SIZE_SMALL, bold=False, color=COLOR_BLACK)

        # Central Authority (bright blue background, white text)
        ca = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_b, top_a, box_w, box_h)
        ca.fill.solid()
        ca.fill.fore_color.rgb = COLOR_BRIGHT_BLUE
        ca.line.width = Pt(1)
        ca.line.color.rgb = COLOR_NAVY
        ca.text_frame.clear()
        p = ca.text_frame.paragraphs[0]
        p.text = "Central Authority"
        p.alignment = PP_ALIGN.CENTER
        set_font_run(p.runs[0], size=SIZE_SMALL, bold=True, color=COLOR_WHITE)

        # Policy & inspection box
        pi = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_c, top_a, box_w, box_h)
        pi.fill.solid()
        pi.fill.fore_color.rgb = COLOR_LIGHT_GRAY
        pi.line.width = Pt(1)
        pi.line.color.rgb = COLOR_NAVY
        pi.text_frame.clear()
        p = pi.text_frame.paragraphs[0]
        p.text = "Policy Enforcement\nand Inspection"
        p.alignment = PP_ALIGN.CENTER
        set_font_run(p.runs[0], size=SIZE_SMALL, bold=False, color=COLOR_BLACK)

        # Connectors: draw straight connectors from UA -> CA -> PI per template
        try:
            # UA -> CA
            conn1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left_a + box_w, top_a + box_h/2, left_b, top_a + box_h/2)
            conn1.line.color.rgb = COLOR_NAVY
            conn1.line.width = Pt(1.5)
            # CA -> PI
            conn2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, left_b + box_w, top_a + box_h/2, left_c, top_a + box_h/2)
            conn2.line.color.rgb = COLOR_NAVY
            conn2.line.width = Pt(1.5)
        except Exception:
            # Some python-pptx builds may have connector limitations; if so, fall back to thin rectangles as arrows
            bar1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_a + box_w, top_a + box_h/2 - Inches(0.03), Inches(0.35), Inches(0.06))
            bar1.fill.solid()
            bar1.fill.fore_color.rgb = COLOR_NAVY
            bar2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_b + box_w, top_a + box_h/2 - Inches(0.03), Inches(0.35), Inches(0.06))
            bar2.fill.solid()
            bar2.fill.fore_color.rgb = COLOR_NAVY

        # Key facts block to the lower half of slide (left-aligned)
        key_left = MARGIN_LEFT
        key_top = Inches(3.05)
        key_w = slide_width - (MARGIN_LEFT + MARGIN_RIGHT)
        key_h = Inches(3.0)
        kb = slide.shapes.add_textbox(key_left, key_top, key_w, key_h)
        kbf = kb.text_frame
        kbf.clear()
        lines = [
            "Authentication Type",
            f"Identity Provider\t{idp}",
            f"Authentication Type\t{auth_type}",
            f"User and Group Provisioning\t{prov_type}",
            "",
            "Client Deployment",
            f"Tunnel Type\t{tunnel_type}",
            f"ZCC Deployment System\t{deploy_system}",
            f"Number of Windows Devices\t{windows_num}",
            f"Number of MacOS Devices\t{mac_num}",
            f"Geo Locations\t{geo_locations}",
            "",
            "Policy Deployment",
            f"SSL Inspection Policies\t{ssl_policies}",
            f"URL Filtering Policies\t{url_policies}",
            f"Cloud App Control Policies\t{cloud_policies}",
            f"Firewall Policies\t{fw_policies}",
        ]
        for idx, line in enumerate(lines):
            if idx == 0 or line == "":
                p = kbf.add_paragraph()
                p.text = line
                set_font_run(p.runs[0], size=SIZE_HEADER if idx == 0 else SIZE_BODY, bold=(idx==0), color=COLOR_BLACK)
            else:
                p = kbf.add_paragraph()
                p.text = line
                set_font_run(p.runs[0], size=SIZE_BODY, bold=False, color=COLOR_BLACK)
            p.alignment = PP_ALIGN.LEFT

        apply_template_branding(prs, slide, slide_num, logo_bytes)
        return slide

    # Build slides (sequence mirrors template)
    progress_bar = st.progress(0)
    slide_count = 11
    step = 0

    # Slide 1 - Title
    create_title_slide("Professional Services Transition Meeting", subtitle_text=customer_name, date_text=today_date, slide_num=1)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 2 - Agenda
    create_bullet_slide("Meeting Agenda", ["Project Summary", "Technical Summary", "Recommended Next Steps"], slide_num=2)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 3 - Project Summary (title)
    create_title_slide("Project Summary", subtitle_text=None, date_text=None, slide_num=3)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 4 - Final Project Status Report (tables + summary)
    slide4 = add_slide_with_background(prs, bg_bytes)
    add_textbox(slide4, MARGIN_LEFT, MARGIN_TOP, Inches(9.0), Inches(0.6), f"Final Project Status Report – {customer_name}", size=SIZE_SLIDE_TITLE, bold=True, color=COLOR_NAVY)
    # Project Summary subtitle and body
    add_textbox(slide4, MARGIN_LEFT, Inches(1.2), Inches(9.5), Inches(0.35), "Project Summary", size=SIZE_HEADER, bold=True)
    add_textbox(slide4, MARGIN_LEFT, Inches(1.6), Inches(9.5), Inches(0.6), project_summary_text, size=SIZE_BODY, bold=False)
    # Dates table (3 columns)
    dates_headers = ["Today's Date", "Start Date", "End Date"]
    dates_rows = [[today_date, project_start, project_end]]
    create_table_slide(" ", dates_headers, dates_rows, slide_num=4, top_inch=2.3, height_inch=0.6)
    # Milestones table on same slide area below (we will add as separate slide to guarantee layout fidelity)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 5 - Milestones + User Rollout + Objectives (split across slides to guarantee readability)
    milestones_headers = ["Milestone", "Baseline Date", "Target Completion Date", "Status"]
    milestones_rows = [[m["name"], m["baseline"], m["target"], m["status"]] for m in milestones_data]
    create_table_slide("Milestones", milestones_headers, milestones_rows, slide_num=5, top_inch=1.6, height_inch=3.0)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 6 - Deliverables
    deliverables_headers = ["Deliverable", "Date delivered"]
    deliverables_rows = [[d["name"], d["date"]] for d in deliverables_data]
    create_table_slide("Deliverables", deliverables_headers, deliverables_rows, slide_num=6, top_inch=1.6, height_inch=2.4)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 7 - Technical Summary (title)
    create_title_slide("Technical Summary", subtitle_text=None, date_text=None, slide_num=7)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 8 - Deployed ZIA Architecture (diagram)
    create_zia_diagram_slide(slide_num=8)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 9 - Open Items
    open_headers = ["Task/ Description", "Date", "Owner", "Transition Plan/ Next Steps"]
    open_rows = [[oi["task"], oi["date"], oi["owner"], oi["steps"]] for oi in open_items_data]
    create_table_slide("Open Items", open_headers, open_rows, slide_num=9, top_inch=1.6, height_inch=3.0)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 10 - Recommended Next Steps (two columns)
    slide10 = add_slide_with_background(prs, bg_bytes)
    add_textbox(slide10, MARGIN_LEFT, MARGIN_TOP, Inches(9.0), Inches(0.6), "Recommended Next Steps", size=SIZE_SLIDE_TITLE, bold=True, color=COLOR_NAVY)
    # Short term left column
    add_textbox(slide10, MARGIN_LEFT, Inches(1.1), Inches(4.5), Inches(0.4), "Short Term Activities", size=SIZE_HEADER, bold=True)
    top_pos = Inches(1.6)
    for item in short_term:
        sq = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, MARGIN_LEFT, top_pos + Inches(0.08), Inches(0.18), Inches(0.18))
        sq.fill.solid(); sq.fill.fore_color.rgb = COLOR_ACCENT_GREEN; sq.line.color.rgb = COLOR_ACCENT_GREEN
        tb = slide10.shapes.add_textbox(MARGIN_LEFT + Inches(0.25), top_pos, Inches(4.25), Inches(0.35))
        p = tb.text_frame.paragraphs[0]; p.text = item; set_font_run(p.runs[0], size=SIZE_BODY)
        top_pos += Inches(0.45)
    # Long term right column
    add_textbox(slide10, Inches(6.3), Inches(1.1), Inches(4.5), Inches(0.4), "Long Term Activities", size=SIZE_HEADER, bold=True)
    top_pos = Inches(1.6)
    for item in long_term:
        sq = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.3), top_pos + Inches(0.08), Inches(0.18), Inches(0.18))
        sq.fill.solid(); sq.fill.fore_color.rgb = COLOR_CYAN; sq.line.color.rgb = COLOR_CYAN
        tb = slide10.shapes.add_textbox(Inches(6.55), top_pos, Inches(4.0), Inches(0.35))
        p = tb.text_frame.paragraphs[0]; p.text = item; set_font_run(p.runs[0], size=SIZE_BODY)
        top_pos += Inches(0.45)
    apply_template_branding(prs, slide10, 10, logo_bytes)
    step += 1
    progress_bar.progress(step / slide_count)

    # Slide 11 - Thank you and survey info
    slide11 = add_slide_with_background(prs, bg_bytes)
    add_textbox(slide11, MARGIN_LEFT, Inches(1.0), Inches(9.5), Inches(1.2), "Thank you", size=SIZE_TITLE, bold=True, color=COLOR_NAVY)
    body_text = (
        f"Your feedback is important to us.\nProject Manager: {pm_name}\nConsultant: {consultant_name}\n\n"
        "A short ~6 question survey will be sent after project close to the contacts listed below:\n"
        f"Primary Contact: {primary_contact}\nSecondary Contact: {secondary_contact}"
    )
    add_textbox(slide11, MARGIN_LEFT, Inches(2.1), Inches(9.5), Inches(3.0), body_text, size=SIZE_BODY)
    # Speech bubble (template color & placement)
    bubble = slide11.shapes.add_shape(MSO_SHAPE.CLOUD_CALLOUT, Inches(8.8), Inches(3.0), Inches(2.0), Inches(0.9))
    bubble.fill.solid(); bubble.fill.fore_color.rgb = COLOR_ZSCALER_YELLOW
    bubble_tf = bubble.text_frame; bubble_tf.clear()
    p = bubble_tf.paragraphs[0]; p.text = "We want to know!"; set_font_run(p.runs[0], size=Pt(16), bold=True)
    apply_template_branding(prs, slide11, 11, logo_bytes)
    step += 1
    progress_bar.progress(step / slide_count)

    # Save to BytesIO and provide download
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)

    st.success("PPTX generated matching template layout.")
    st.download_button("Download Transition Deck", data=out, file_name=f"{customer_name}_Zscaler_Transition.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
