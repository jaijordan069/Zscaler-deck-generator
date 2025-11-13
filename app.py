import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.table import MSO_TABLE_STYLE
from datetime import datetime
import io

# Page config
st.set_page_config(page_title="Zscaler Transition Deck PPT Generator", layout="wide")

st.title("Zscaler Professional Services Transition Deck PPT Generator")
st.markdown("Fill in details to generate a customized PowerPoint transition meeting deck based on the provided template.")

# Sidebar for instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    - Enter customer-specific details in the form.
    - For lists/tables (e.g., milestones, open items), use comma-separated values or multiple inputs where prompted.
    - Dates should be in DD/MM/YYYY format.
    - Click 'Generate Deck' to create and download the PPTX file.
    """)

# Customer & Project Basics
st.header("Customer & Project Basics")
col1, col2, col3 = st.columns(3)
customer_name = col1.text_input("Customer Name", value="Pixartprinting")
today_date = col2.text_input("Today's Date (DD/MM/YYYY)", value="14/11/2025")
project_start = col3.text_input("Project Start Date (DD/MM/YYYY)", value="01/06/2025")
project_end = st.text_input("Project End Date (DD/MM/YYYY)", value="14/11/2025")

project_summary_text = st.text_area("Project Summary Text", 
    value="More than half of the users have been deployed and there were not any critical issues. Not expected issues during enrollment of remaining users")

# Milestones
st.header("Milestones")
milestones_data = []
for i in range(7):
    with st.expander(f"Milestone {i+1}"):
        name = st.text_input(f"Milestone Name {i+1}", key=f"mname_{i}")
        baseline = st.text_input(f"Baseline Date {i+1} (DD/MM/YYYY)", key=f"mbaseline_{i}")
        target = st.text_input(f"Target Completion {i+1} (DD/MM/YYYY)", key=f"mtarget_{i}")
        status = st.text_input(f"Status {i+1} (e.g., Completed)", key=f"mstatus_{i}")
        if name:
            milestones_data.append({"name": name, "baseline": baseline, "target": target, "status": status})

# User Rollout Roadmap
st.header("User Rollout Roadmap")
col_p1, col_p2 = st.columns(2)
with col_p1:
    st.subheader("Pilot")
    pilot_target = st.number_input("Pilot Target Users", value=100)
    pilot_current = st.number_input("Pilot Current Users", value=449)
    pilot_completion = st.text_input("Pilot Completion Date", value="14/11/2025")
    pilot_status = st.text_input("Pilot Status", value="Completed")
with col_p2:
    st.subheader("Production")
    prod_target = st.number_input("Production Target Users", value=800)
    prod_current = st.number_input("Production Current Users", value=449)
    prod_completion = st.text_input("Production Completion Date", value="14/11/2025")
    prod_status = st.text_input("Production Status", value="In Progress")

# Project Objectives
st.header("Project Objectives")
objectives_data = []
for i in range(3):
    with st.expander(f"Objective {i+1}"):
        objective = st.text_area(f"Planned Objective {i+1}", key=f"obj_{i}", height=50)
        actual = st.text_area(f"Actual Result {i+1}", key=f"act_{i}", height=50)
        deviation = st.text_area(f"Deviation/Cause {i+1}", key=f"dev_{i}", height=50)
        if objective:
            objectives_data.append({"objective": objective, "actual": actual, "deviation": deviation})

# Deliverables
st.header("Deliverables")
deliverables_data = []
for i in range(5):
    with st.expander(f"Deliverable {i+1}"):
        name = st.text_input(f"Deliverable Name {i+1}", key=f"dname_{i}")
        date_del = st.text_input(f"Date Delivered {i+1}", key=f"ddate_{i}")
        if name:
            deliverables_data.append({"name": name, "date": date_del})

# Technical Summary
st.header("Technical Summary")
col_t1, col_t2 = st.columns(2)
with col_t1:
    st.subheader("Authentication & Provisioning")
    idp = st.text_input("Identity Provider", value="Entra ID")
    auth_type = st.text_input("Authentication Type", value="SAML 2.0")
    prov_type = st.text_input("User/Group Provisioning", value="SCIM Provisioning")
with col_t2:
    st.subheader("Client Deployment")
    tunnel_type = st.text_input("Tunnel Type", value="ZCC with Z-Tunnel 2.0")
    deploy_system = st.text_input("ZCC Deployment System", value="MS Intune/Jamf")

col_d1, col_d2, col_d3 = st.columns(3)
windows_num = col_d1.number_input("Number of Windows Devices", value=351)
mac_num = col_d2.number_input("Number of MacOS Devices", value=98)
geo_locations = col_d3.text_input("Geo Locations", value="Europe, North Africa, USA")

col_pol1, col_pol2, col_pol3, col_pol4 = st.columns(4)
ssl_policies = col_pol1.number_input("SSL Inspection Policies", value=10)
url_policies = col_pol2.number_input("URL Filtering Policies", value=5)
cloud_policies = col_pol3.number_input("Cloud App Control Policies", value=5)
fw_policies = col_pol4.number_input("Firewall Policies", value=15)

# Open Items
st.header("Open Items")
open_items_data = []
for i in range(6):
    with st.expander(f"Open Item {i+1}"):
        task = st.text_input(f"Task/Description {i+1}", key=f"otask_{i}")
        o_date = st.text_input(f"Date {i+1}", key=f"odate_{i}")
        owner = st.text_input(f"Owner {i+1}", key=f"oowner_{i}")
        steps = st.text_area(f"Transition Plan/Next Steps {i+1}", key=f"osteps_{i}", height=50)
        if task:
            open_items_data.append({"task": task, "date": o_date, "owner": owner, "steps": steps})

# Recommended Next Steps
st.header("Recommended Next Steps")
st.subheader("Short Term Activities")
short_term_input = st.text_area("Short Term (comma-separated)", value="Finish Production rollout, Tighten Firewall policies, Tighten Cloud App Control Policies, Fine tune SSL Inspection policies, Configure Role Based Access Control (RBAC), Configure DLP policies")
short_term = [item.strip() for item in short_term_input.split(",") if item.strip()]

st.subheader("Long Term Activities")
long_term_input = st.text_area("Long Term (comma-separated)", value="Deploy ZCC on Mobile devices, Consider an upgrade of Sandbox license to have better antimalware protection, Consider an upgrade of the Firewall License to be able to apply policies based on user groups and network applications, Adopt additional Zscaler solutions like Zscaler Private Access (ZPA) or Zscaler Digital experience (ZDX), Consider using ZCC Client when the users are on-prem for a more consistent user experience, Integrate ZIA with 3rd party SIEM")
long_term = [item.strip() for item in long_term_input.split(",") if item.strip()]

# Contacts
st.header("Contacts")
col_c1, col_c2 = st.columns(2)
pm_name = col_c1.text_input("Project Manager Name", value="Alex Vazquez")
consultant_name = col_c2.text_input("Consultant Name", value="Alex Vazquez")
primary_contact = st.text_input("Primary Contact", value="Teia proctor")
secondary_contact = st.text_input("Secondary Contact", value="Marco Sattier")

# Generate button
if st.button("Generate Transition Deck"):
    # Create PPTX
    prs = Presentation()

    # Helper function to add title slide
    def add_title_slide(title, subtitle=None):
        slide_layout = prs.slide_layouts[0]  # Title slide
        slide = prs.slides.add_slide(slide_layout)
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
        if subtitle:
            subtitle_placeholder = slide.placeholders[1]
            subtitle_placeholder.text = subtitle

    # Helper for bullet slide
    def add_bullet_slide(title, bullets):
        slide_layout = prs.slide_layouts[1]  # Title and content
        slide = prs.slides.add_slide(slide_layout)
        title_placeholder = slide.shapes.title
        title_placeholder.text = title
        content_placeholder = slide.placeholders[1]
        tf = content_placeholder.text_frame
        tf.clear()
        for bullet in bullets:
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 0

    # Helper for table slide
    def add_table_slide(title, rows, cols, data):
        slide_layout = prs.slide_layouts[5]  # Blank
        slide = prs.slides.add_slide(slide_layout)
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(5))
        tf = txBox.text_frame
        tf.text = title
        tf.paragraphs[0].font.size = Pt(24)

        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(9)
        height = Inches(4)
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table

        # Headers
        for i, header in enumerate(data[0]):
            table.cell(0, i).text = header
            table.cell(0, i).fill.solid()
            table.cell(0, i).fill.fore_color.rgb = RGBColor(0, 112, 192)  # Blue
            table.cell(0, i).text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

        # Data
        for row_idx, row in enumerate(data[1:], 1):
            for col_idx, cell_text in enumerate(row):
                table.cell(row_idx, col_idx).text = str(cell_text)

    # Slide 1: Title
    add_title_slide("Professional Services Transition Meeting", f"{customer_name}\n{today_date}")

    # Slide 2: Agenda
    agenda_bullets = ["Project Summary", "Technical Summary", "Recommended Next Steps"]
    add_bullet_slide("Meeting Agenda", agenda_bullets)

    # Slide 3: Project Summary Title
    add_title_slide("Project Summary")

    # Slide 4: Project Status Report
    # Milestones table
    milestones_headers = ["Milestone", "Baseline Date", "Target Completion Date", "Status"]
    milestones_rows = [[m["name"], m["baseline"], m["target"], m["status"]] for m in milestones_data]
    add_table_slide("Final Project Status Report – " + customer_name, len(milestones_rows) + 1, 4, [milestones_headers] + milestones_rows)

    # User Rollout table
    rollout_headers = ["Milestone", "Target Users", "Current Users", "Target Completion", "Status"]
    rollout_rows = [
        ["Pilot", pilot_target, pilot_current, pilot_completion, pilot_status],
        ["Production", prod_target, prod_current, prod_completion, prod_status]
    ]
    add_table_slide("User Rollout Roadmap", 3, 5, [rollout_headers] + rollout_rows)

    # Objectives table
    objectives_headers = ["Planned Project Objective (Target)", "Actual Project Result (Actual)", "Deviation/Cause"]
    objectives_rows = [[o["objective"], o["actual"], o["deviation"]] for o in objectives_data]
    add_table_slide("Project Objectives", len(objectives_rows) + 1, 3, [objectives_headers] + objectives_rows)

    # Slide 5: Deliverables
    deliverables_headers = ["Deliverable", "Date delivered"]
    deliverables_rows = [[d["name"], d["date"]] for d in deliverables_data]
    add_table_slide("Deliverables", len(deliverables_rows) + 1, 2, [deliverables_headers] + deliverables_rows)

    # Slide 6: Technical Summary Title
    add_title_slide("Technical Summary")

    # Slide 7: Deployed ZIA Architecture - Bullets for simplicity
    tech_bullets = [
        f"Identity Provider: {idp}",
        f"Authentication Type: {auth_type}",
        f"User and Group Provisioning: {prov_type}",
        f"Tunnel Type: {tunnel_type}",
        f"ZCC Deployment System: {deploy_system}",
        f"Number of Windows and MacOS Devices: {windows_num} Windows Devices, {mac_num} MacOS Devices",
        f"Geo Locations: {geo_locations}",
        f"SSL Inspection Policies: {ssl_policies} Policies",
        f"URL Filtering Policies: {url_policies} Policies",
        f"Cloud App Control Policies: {cloud_policies} Policies",
        f"Firewall Policies: {fw_policies} Policies"
    ]
    add_bullet_slide("Deployed ZIA Architecture", tech_bullets)

    # Slide 8: Open Items
    open_items_headers = ["Task/Description", "Date", "Owner", "Transition Plan/Next Steps"]
    open_items_rows = [[oi["task"], oi["date"], oi["owner"], oi["steps"]] for oi in open_items_data]
    add_table_slide("Open Items", len(open_items_rows) + 1, 4, [open_items_headers] + open_items_rows)

    # Slide 9: Recommended Next Steps
    short_term_slide = prs.slides.add_slide(prs.slide_layouts[1])
    short_title = short_term_slide.shapes.title
    short_title.text = "Recommended Next Steps - Short Term Activities"
    short_content = short_term_slide.placeholders[1]
    short_tf = short_content.text_frame
    short_tf.clear()
    for item in short_term:
        p = short_tf.add_paragraph()
        p.text = "• " + item

    long_term_slide = prs.slides.add_slide(prs.slide_layouts[1])
    long_title = long_term_slide.shapes.title
    long_title.text = "Recommended Next Steps - Long Term Activities"
    long_content = long_term_slide.placeholders[1]
    long_tf = long_content.text_frame
    long_tf.clear()
    for item in long_term:
        p = long_tf.add_paragraph()
        p.text = "• " + item

    # Slide 10: Thank You
    thank_slide = prs.slides.add_slide(prs.slide_layouts[1])
    thank_title = thank_slide.shapes.title
    thank_title.text = "Thank you"
    content = thank_slide.placeholders[1]
    tf = content.text_frame
    tf.clear()
    p = tf.add_paragraph()
    p.text = f"Your feedback on our project and Professional Services team is important to us.\n\nProject Manager: {pm_name}\nConsultant: {consultant_name}\n\nA short ~6 question survey... (Primary Contact: {primary_contact}, Secondary Contact: {secondary_contact})"

    # Slide 11: Final Thank You
    add_title_slide("Thank you")

    # Add Zscaler footer to all slides
    for slide in prs.slides:
        footer = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(0.5))
        f_tf = footer.text_frame
        f_tf.text = "Zscaler, Inc. All rights reserved. © 2025"
        f_tf.paragraphs[0].font.size = Pt(8)

    # Save to bytes
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)

    st.success("Deck generated! Download below.")
    st.download_button(
        label="Download PPTX",
        data=bio.getvalue(),
        file_name=f"{customer_name}_Transition_Deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
