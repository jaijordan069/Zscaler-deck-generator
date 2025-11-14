import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_CONNECTOR
import io
import re
import requests

# Page config
st.set_page_config(page_title="Zscaler Transition Deck PPT Generator", layout="wide")

# Add Zscaler design background to Streamlit app
st.markdown("""
<style>
.stApp {
    background: linear-gradient(to bottom, #0066CC, white);
    background-size: cover;
}
</style>
""", unsafe_allow_html=True)

# Add Zscaler logo
st.image("https://companieslogo.com/img/orig/ZS-46a5871c.png?t=1720244494", width=200)

st.title("Zscaler Professional Services Transition Deck PPT Generator")

st.markdown("Fill in details to generate a customized PowerPoint transition meeting deck based on the provided template.")

# Sidebar for instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    - Enter customer-specific details in the form.
    - Required fields are marked with *.
    - Dates must be in DD/MM/YYYY format.
    - Use the 'Preview Summary' button to review your data before generating.
    - Click 'Generate Deck' to create and download the PPTX file.
    """)

# Validation function for dates
def is_valid_date(date_str):
    return bool(re.match(r'^\d{2}/\d{2}/\d{4}$', date_str))

# Customer & Project Basics
st.header("Customer & Project Basics")
col1, col2, col3 = st.columns(3)
customer_name = col1.text_input("Customer Name *", value="Pixartprinting", help="Enter the customer's name, e.g., Pixartprinting")
today_date = col2.text_input("Today's Date (DD/MM/YYYY) *", value="19/09/2025", help="Enter today's date in DD/MM/YYYY format")
project_start = col3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025", help="Enter project start date in DD/MM/YYYY format")
project_end = st.text_input("Project End Date (DD/MM/YYYY) *", value="19/09/2025", help="Enter project end date in DD/MM/YYYY format")
project_summary_text = st.text_area("Project Summary Text",
    value="More than half of the users have been deployed and there were not any critical issues. Not expected issues during enrollment of remaining users",
    help="Provide a brief project summary")

# Milestones
st.header("Milestones")
milestones_data = []
for i in range(7):
    with st.expander(f"Milestone {i+1}", expanded=i < 1): # Expand first by default
        name = st.text_input(f"Milestone Name {i+1}", key=f"mname_{i}", help="Enter milestone name")
        baseline = st.text_input(f"Baseline Date {i+1} (DD/MM/YYYY)", key=f"mbaseline_{i}", help="Enter baseline date in DD/MM/YYYY format")
        target = st.text_input(f"Target Completion {i+1} (DD/MM/YYYY)", key=f"mtarget_{i}", help="Enter target completion date in DD/MM/YYYY format")
        status = st.text_input(f"Status {i+1} (e.g., Completed)", key=f"mstatus_{i}", help="Enter status, e.g., Completed or In Progress")
        if name:
            milestones_data.append({"name": name, "baseline": baseline, "target": target, "status": status})

# User Rollout Roadmap
st.header("User Rollout Roadmap")
col_p1, col_p2 = st.columns(2)
with col_p1:
    st.subheader("Pilot")
    pilot_target = st.number_input("Pilot Target Users", value=100, help="Enter target number of users for pilot")
    pilot_current = st.number_input("Pilot Current Users", value=449, help="Enter current number of users in pilot")
    pilot_completion = st.text_input("Pilot Completion Date", value="19/09/2025", help="Enter pilot completion date in DD/MM/YYYY format")
    pilot_status = st.text_input("Pilot Status", value="Completed", help="Enter pilot status, e.g., Completed")
with col_p2:
    st.subheader("Production")
    prod_target = st.number_input("Production Target Users", value=800, help="Enter target number of users for production")
    prod_current = st.number_input("Production Current Users", value=449, help="Enter current number of users in production")
    prod_completion = st.text_input("Production Completion Date", value="19/09/2025", help="Enter production completion date in DD/MM/YYYY format")
    prod_status = st.text_input("Production Status", value="In Progress", help="Enter production status, e.g., In Progress")

# Project Objectives
st.header("Project Objectives")
objectives_data = []
for i in range(3):
    with st.expander(f"Objective {i+1}", expanded=i < 1):
        objective = st.text_area(f"Planned Objective {i+1}", key=f"obj_{i}", height=50, help="Enter the planned objective")
        actual = st.text_area(f"Actual Result {i+1}", key=f"act_{i}", height=50, help="Enter the actual result")
        deviation = st.text_area(f"Deviation/Cause {i+1}", key=f"dev_{i}", height=50, help="Enter any deviation or cause")
        if objective:
            objectives_data.append({"objective": objective, "actual": actual, "deviation": deviation})

# Deliverables
st.header("Deliverables")
deliverables_data = []
for i in range(5):
    with st.expander(f"Deliverable {i+1}", expanded=i < 1):
        name = st.text_input(f"Deliverable Name {i+1}", key=f"dname_{i}", help="Enter deliverable name")
        date_del = st.text_input(f"Date Delivered {i+1}", key=f"ddate_{i}", help="Enter date delivered in DD/MM/YYYY format")
        if name:
            deliverables_data.append({"name": name, "date": date_del})

# Technical Summary
st.header("Technical Summary")
col_t1, col_t2 = st.columns(2)
with col_t1:
    st.subheader("Authentication & Provisioning")
    idp = st.text_input("Identity Provider", value="Entra ID", help="Enter identity provider, e.g., Entra ID")
    auth_type = st.text_input("Authentication Type", value="SAML 2.0", help="Enter authentication type, e.g., SAML 2.0")
    prov_type = st.text_input("User/Group Provisioning", value="SCIM Provisioning", help="Enter provisioning type, e.g., SCIM Provisioning")
with col_t2:
    st.subheader("Client Deployment")
    tunnel_type = st.text_input("Tunnel Type", value="ZCC with Z-Tunnel 2.0", help="Enter tunnel type, e.g., ZCC with Z-Tunnel 2.0")
    deploy_system = st.text_input("ZCC Deployment System", value="MS Intune/Jamf", help="Enter deployment system, e.g., MS Intune/Jamf")
col_d1, col_d2, col_d3 = st.columns(3)
windows_num = col_d1.number_input("Number of Windows Devices", value=351, help="Enter number of Windows devices")
mac_num = col_d2.number_input("Number of MacOS Devices", value=98, help="Enter number of MacOS devices")
geo_locations = col_d3.text_input("Geo Locations", value="Europe, North Africa, USA", help="Enter geo locations, comma-separated")
col_pol1, col_pol2, col_pol3, col_pol4 = st.columns(4)
ssl_policies = col_pol1.number_input("SSL Inspection Policies", value=10, help="Enter number of SSL inspection policies")
url_policies = col_pol2.number_input("URL Filtering Policies", value=5, help="Enter number of URL filtering policies")
cloud_policies = col_pol3.number_input("Cloud App Control Policies", value=5, help="Enter number of cloud app control policies")
fw_policies = col_pol4.number_input("Firewall Policies", value=15, help="Enter number of firewall policies")

# Open Items
st.header("Open Items")
open_items_data = []
for i in range(6):
    with st.expander(f"Open Item {i+1}", expanded=i < 1):
        task = st.text_input(f"Task/Description {i+1}", key=f"otask_{i}", help="Enter task description")
        o_date = st.text_input(f"Date {i+1}", key=f"odate_{i}", help="Enter date, e.g., October 2025")
        owner = st.text_input(f"Owner {i+1}", key=f"oowner_{i}", help="Enter owner, e.g., Pixartprinting")
        steps = st.text_area(f"Transition Plan/Next Steps {i+1}", key=f"osteps_{i}", height=50, help="Enter next steps")
        if task:
            open_items_data.append({"task": task, "date": o_date, "owner": owner, "steps": steps})

# Recommended Next Steps
st.header("Recommended Next Steps")
st.subheader("Short Term Activities")
short_term_input = st.text_area("Short Term (comma-separated)", value="Finish Production rollout, Tighten Firewall policies, Tighten Cloud App Control Policies, Fine tune SSL Inspection policies, Configure Role Based Access Control (RBAC), Configure DLP policies", help="Enter short term activities, comma-separated")
short_term = [item.strip() for item in short_term_input.split(",") if item.strip()]
st.subheader("Long Term Activities")
long_term_input = st.text_area("Long Term (comma-separated)", value="Deploy ZCC on Mobile devices, Consider an upgrade of Sandbox license to have better antimalware protection, Consider an upgrade of the Firewall License to be able to apply policies based on user groups and network applications, Adopt additional Zscaler solutions like Zscaler Private Access (ZPA) or Zscaler Digital experience (ZDX), Consider using ZCC Client when the users are on-prem for a more consistent user experience, Integrate ZIA with 3rd party SIEM", help="Enter long term activities, comma-separated")
long_term = [item.strip() for item in long_term_input.split(",") if item.strip()]

# Contacts
st.header("Contacts")
col_c1, col_c2 = st.columns(2)
pm_name = col_c1.text_input("Project Manager Name", value="Alex Vazquez", help="Enter project manager name")
consultant_name = col_c2.text_input("Consultant Name", value="Alex Vazquez", help="Enter consultant name")
primary_contact = st.text_input("Primary Contact", value="Teia proctor", help="Enter primary contact")
secondary_contact = st.text_input("Secondary Contact", value="Marco Sattier", help="Enter secondary contact")

# Preview Summary
if st.button("Preview Summary"):
    st.write(f"Deck for {customer_name} on {today_date}:")
    st.write(f"- Project Summary: {project_summary_text[:100]}...")
    st.write(f"- {len(milestones_data)} Milestones (e.g., {milestones_data[0]['name'] if milestones_data else 'None'})")
    st.write(f"- Pilot Rollout: {pilot_current}/{pilot_target} users, Status: {pilot_status}")
    st.write(f"- Production Rollout: {prod_current}/{prod_target} users, Status: {prod_status}")
    st.write(f"- {len(objectives_data)} Objectives (e.g., {objectives_data[0]['objective'] if objectives_data else 'None'})")
    st.write(f"- {len(deliverables_data)} Deliverables (e.g., {deliverables_data[0]['name'] if deliverables_data else 'None'})")
    st.write(f"- Technical: {windows_num} Windows, {mac_num} MacOS, Geo: {geo_locations}, Policies: SSL {ssl_policies}, URL {url_policies}, Cloud {cloud_policies}, FW {fw_policies}")
    st.write(f"- {len(open_items_data)} Open Items (e.g., {open_items_data[0]['task'] if open_items_data else 'None'})")
    st.write(f"- Short Term: {len(short_term)} items (e.g., {short_term[0] if short_term else 'None'})")
    st.write(f"- Long Term: {len(long_term)} items (e.g., {long_term[0] if long_term else 'None'})")
    st.write(f"- Contacts: PM {pm_name}, Consultant {consultant_name}, Primary {primary_contact}, Secondary {secondary_contact}")

# Generate button
if st.button("Generate Transition Deck"):
    # Validation
    if not customer_name:
        st.error("Customer Name is required.")
    elif not all(is_valid_date(d) for d in [today_date, project_start, project_end, pilot_completion, prod_completion]):
        st.error("All dates must be in DD/MM/YYYY format.")
    else:
        # Create PPTX
        prs = Presentation()
        # Helper to set background based on theme
        def set_background(slide, theme):
            fill = slide.background.fill
            fill.solid()
            if theme == "Navy":
                fill.fore_color.rgb = NAVY
            else:
                fill.fore_color.rgb = WHITE

        # Helper to add logo and footer (consistent across all slides)
        def add_logo_footer(slide, theme):
            # Logo in top right
            txBox = slide.shapes.add_textbox(Inches(10.5), Inches(0.1), Inches(2), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Zscaler"
            p.alignment = PP_ALIGN.RIGHT
            p.font.name = 'Century Gothic'
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = WHITE if theme == "Navy" else NAVY

            # Footer in bottom right
            txBox = slide.shapes.add_textbox(Inches(10), Inches(7), Inches(3), Inches(0.3))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "2025 Zscaler, Inc. All rights reserved"
            p.alignment = PP_ALIGN.RIGHT
            p.font.name = 'Century Gothic'
            p.font.size = Pt(8)
            p.font.color.rgb = WHITE if theme == "Navy" else NAVY

        # Blank layout for custom building
        blank_layout = prs.slide_layouts[6]

        # Cover Slide (based on Cover A layout)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(2))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = f"{customer_name} Zscaler Transition Plan".title()  # Title case
        p.font.name = 'Century Gothic'
        p.font.size = Pt(44)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = BRIGHT_BLUE if theme == "White" else WHITE
        # Subtitle
        txBox = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = f"From {start_date} to {end_date}".capitalize()  # Sentence case
        p.font.name = 'Century Gothic'
        p.font.size = Pt(24)
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = NAVY if theme == "White" else CYAN

        # Agenda Slide (based on Agenda layout)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Agenda".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Items with square bullets
        agenda_items = ["Current State", "Proposed Architecture", "Transition Milestones", "Benefits", "Team", "Next Steps"]
        top = Inches(2.5)
        for item in agenda_items:
            # Square bullet
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), top + Inches(0.1), Inches(0.2), Inches(0.2))
            shape.fill.solid()
            shape.fill.fore_color.rgb = BRIGHT_BLUE if theme == "White" else ACCENT_GREEN
            shape.line.color.rgb = shape.fill.fore_color.rgb
            # Text
            txBox = slide.shapes.add_textbox(Inches(1.5), top, Inches(10), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = item.capitalize()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(20)
            p.font.color.rgb = NAVY if theme == "White" else LIGHT_GRAY
            top += Inches(0.6)

        # Current State Slide (content slide with columns)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Current State".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Body (example text, can customize)
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(5), Inches(4))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Legacy architecture with VPNs and firewalls.".capitalize()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(18)
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Second column
        txBox = slide.shapes.add_textbox(Inches(6.5), Inches(2.5), Inches(5), Inches(4))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Challenges: High latency, security gaps.".capitalize()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(18)
        p.font.color.rgb = NAVY if theme == "White" else WHITE

        # Proposed Architecture Slide (simple diagram based on template specs)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Proposed Zscaler Architecture".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Diagram: Cloud with arrows (using template specs: 1pt stroke, arrows)
        # Cloud shape (filled, Navy stroke)
        cloud = slide.shapes.add_shape(MSO_SHAPE.CLOUD, Inches(4), Inches(3), Inches(4), Inches(2))
        cloud.fill.solid()
        cloud.fill.fore_color.rgb = LIGHT_GRAY
        cloud.line.color.rgb = NAVY
        cloud.line.width = Pt(1)
        # Text in shape
        tf = cloud.text_frame
        p = tf.add_paragraph()
        p.text = "Zscaler Zero Trust Exchange".capitalize()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(18)
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = BRIGHT_BLUE
        # Arrow to cloud
        arrow = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2), Inches(4), Inches(4), Inches(4))
        arrow.line.color.rgb = BRIGHT_BLUE
        arrow.line.width = Pt(1.25)
        # User icon (simple rectangle for example)
        user = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(3.5), Inches(1), Inches(1))
        user.fill.solid()
        user.fill.fore_color.rgb = CYAN
        user.line.color.rgb = NAVY

        # Transition Milestones Slide (timeline based on template)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Transition Milestones".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Timeline line
        line = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(1), Inches(4), Inches(11), Inches(0))
        line.line.color.rgb = BRIGHT_BLUE
        line.line.width = Pt(1.25)
        # Parse milestones
        milestone_list = [m.strip() for m in milestones.split('\n') if m.strip()]
        spacing = 11 / max(1, len(milestone_list))
        left = Inches(1)
        for m in milestone_list:
            date, desc = m.split(':', 1) if ':' in m else ("Date", m)
            # Point on line
            point = slide.shapes.add_shape(MSO_SHAPE.OVAL, left - Inches(0.1), Inches(3.9), Inches(0.2), Inches(0.2))
            point.fill.solid()
            point.fill.fore_color.rgb = ACCENT_GREEN
            point.line.color.rgb = NAVY
            # Date above
            txBox = slide.shapes.add_textbox(left - Inches(0.5), Inches(3), Inches(1.5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = date.strip()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(14)
            p.font.color.rgb = NAVY if theme == "White" else WHITE
            # Desc below
            txBox = slide.shapes.add_textbox(left - Inches(1), Inches(4.5), Inches(3), Inches(1))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = desc.strip().capitalize()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(16)
            p.font.color.rgb = NAVY if theme == "White" else WHITE
            left += Inches(spacing)

        # Benefits Slide (stats with checkmarks)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Benefits".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Objectives as benefits with checkmarks
        obj_list = [o.strip().capitalize() for o in objectives.split('\n') if o.strip()]
        top = Inches(2.5)
        for obj in obj_list:
            # Checkmark using textbox with Unicode
            txBox = slide.shapes.add_textbox(Inches(1), top, Inches(0.5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = '✓'
            p.font.name = 'Century Gothic'
            p.font.size = Pt(20)
            p.font.color.rgb = ACCENT_GREEN
            p.alignment = PP_ALIGN.CENTER
            # Text
            txBox = slide.shapes.add_textbox(Inches(1.6), top, Inches(10), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = obj
            p.font.name = 'Century Gothic'
            p.font.size = Pt(20)
            p.font.color.rgb = NAVY if theme == "White" else WHITE
            top += Inches(0.6)

        # Team Slide (speakers layout)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Team".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Team members
        team_list = [t.strip() for t in team_members.split('\n') if t.strip()]
        top = Inches(2.5)
        for t in team_list:
            name, position = t.split(',', 1) if ',' in t else (t, "Position")
            txBox = slide.shapes.add_textbox(Inches(1), top, Inches(5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = name.strip().title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = BRIGHT_BLUE if theme == "White" else CYAN
            txBox = slide.shapes.add_textbox(Inches(1), top + Inches(0.5), Inches(5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = position.strip().capitalize()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(18)
            p.font.color.rgb = NAVY if theme == "White" else LIGHT_GRAY
            top += Inches(1.5)

        # Next Steps / Thanks Slide (quote layout)
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Next Steps".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Body
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(2))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Schedule kickoff meeting. Review architecture. Begin implementation.".capitalize()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(18)
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Thanks
        txBox = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Thanks".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = BRIGHT_BLUE if theme == "White" else CYAN

        # Save to buffer and provide download
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        st.download_button("Download Transition Deck", bio, file_name=f"{customer_name}_Zscaler_Transition.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
