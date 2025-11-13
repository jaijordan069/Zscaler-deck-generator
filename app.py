import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL, MSO_PATTERN, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR_TYPE
from datetime import datetime
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
    background-image: radial-gradient(circle, rgba(255,255,255,0.2) 1px, transparent 1px);
    background-size: 10px 10px;
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
today_date = col2.text_input("Today's Date (DD/MM/YYYY) *", value="14/11/2025", help="Enter today's date in DD/MM/YYYY format")
project_start = col3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025", help="Enter project start date in DD/MM/YYYY format")
project_end = st.text_input("Project End Date (DD/MM/YYYY) *", value="14/11/2025", help="Enter project end date in DD/MM/YYYY format")

project_summary_text = st.text_area("Project Summary Text", 
    value="More than half of the users have been deployed and there were not any critical issues. Not expected issues during enrollment of remaining users",
    help="Provide a brief project summary")

# Milestones
st.header("Milestones")
milestones_data = []
for i in range(7):
    with st.expander(f"Milestone {i+1}", expanded=i < 1):  # Expand first by default
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
    pilot_completion = st.text_input("Pilot Completion Date", value="14/11/2025", help="Enter pilot completion date in DD/MM/YYYY format")
    pilot_status = st.text_input("Pilot Status", value="Completed", help="Enter pilot status, e.g., Completed")
with col_p2:
    st.subheader("Production")
    prod_target = st.number_input("Production Target Users", value=800, help="Enter target number of users for production")
    prod_current = st.number_input("Production Current Users", value=449, help="Enter current number of users in production")
    prod_completion = st.text_input("Production Completion Date", value="14/11/2025", help="Enter production completion date in DD/MM/YYYY format")
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

        # Set slide master for background, header, footer
        master = prs.slide_masters[0]
        master_slide = master.slide_layouts[0]

        # Background gradient with dot pattern
        background = master.background
        fill = background.fill
        fill.gradient()
        fill.gradient_stops[0].color.rgb = RGBColor(0, 102, 204)  # Dark blue
        fill.gradient_stops[1].color.rgb = RGBColor(255, 255, 255)  # White
        fill.gradient_angle = 90  # Vertical fade
        fill.patterned()
        fill.pattern = MSO_PATTERN.DOTTED_GRID
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        fill.back_color.rgb = RGBColor(255, 255, 255)

        # Helper function to add title slide
        def add_title_slide(title, subtitle=None):
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            title_placeholder = slide.shapes.title
            title_placeholder.text = title
            title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
            title_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            title_placeholder.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            if subtitle:
                subtitle_placeholder = slide.placeholders[1]
                subtitle_placeholder.text = subtitle
                subtitle_placeholder.text_frame.paragraphs[0].font.size = Pt(32)
                subtitle_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)  # Red for customer
            add_header_footer_number(slide, str(len(prs.slides)))
            return slide

        # Helper for bullet slide
        def add_bullet_slide(title, bullets):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            title_placeholder = slide.shapes.title
            title_placeholder.text = title
            title_placeholder.text_frame.paragraphs[0].font.size = Pt(28)
            title_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
            content_placeholder = slide.placeholders[1]
            tf = content_placeholder.text_frame
            tf.clear()
            for bullet in bullets:
                p = tf.add_paragraph()
                p.text = bullet
                p.level = 0
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.bullet.color.rgb = RGBColor(0, 102, 204)  # Blue bullets
                p.alignment = PP_ALIGN.LEFT
            add_header_footer_number(slide, str(len(prs.slides)))
            return slide

        # Helper for table slide
        def add_table_slide(title, rows, cols, data):
            slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank
            # Title
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(0.5))
            tf = txBox.text_frame
            tf.text = title
            tf.paragraphs[0].font.size = Pt(28)
            tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

            # Table
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(4)
            table = slide.shapes.add_table(rows, cols, left, top, width, height).table

            # Headers
            for i, header in enumerate(data[0]):
                cell = table.cell(0, i)
                cell.text = header
                fill = cell.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(0, 102, 204)  # Blue
                tf = cell.text_frame
                p = tf.paragraphs[0]
                p.font.color.rgb = RGBColor(255, 255, 255)
                p.font.bold = True
                p.font.size = Pt(14)
                p.alignment = PP_ALIGN.LEFT
                # Borders
                for side in ['left', 'top', 'right', 'bottom']:
                    line = getattr(cell, f"{side}_line")
                    line.color.rgb = RGBColor(255, 255, 255)
                    line.width = Pt(1)

            # Data
            for row_idx, row in enumerate(data[1:], 1):
               
