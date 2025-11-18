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
from docx import Document
from docx.shared import Inches as DocInches, Pt as DocPt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF

# Page config
st.set_page_config(page_title="Zscaler Deck Generator", layout="wide")

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

st.title("Zscaler Deck Generator")

st.markdown("Fill in details to generate a customized deck. Choose between Transition and Kick-Off Deck.")

# Sidebar for instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    - Enter customer-specific details in the form.
    - Required fields are marked with *.
    - Dates must be in DD/MM/YYYY format.
    - Use the 'Preview Summary' button to review your data before generating.
    - Click 'Generate Deck' to create and download the file in selected format.
    """)

# Color palette from the 2025 template
BRIGHT_BLUE = RGBColor(37, 108, 247)
NAVY = RGBColor(0, 23, 68)
WHITE = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(229, 241, 250)
CYAN = RGBColor(18, 212, 255)
ACCENT_GREEN = RGBColor(107, 255, 179)
PINK = RGBColor(54, 0, 226)
THREAT_RED = RGBColor(237, 25, 81)
BLACK = RGBColor(0, 0, 0)

# Logo URL
LOGO_URL = "https://upload.wikimedia.org/wikipedia/commons/thumb/8/8b/Zscaler_logo.svg/512px-Zscaler_logo.svg.png"

# Validation function for dates
def is_valid_date(date_str):
    return bool(re.match(r'^\d{2}/\d{2}/\d{4}$', date_str))

# Tabs for deck types
tab1, tab2 = st.tabs(["Transition Deck", "Kick-Off Deck"])

with tab1:
    st.header("Transition Deck Inputs")
    col1, col2, col3 = st.columns(3)
    customer_name = col1.text_input("Customer Name *", value="Pixartprinting", key="customer_name_trans")
    today_date = col2.text_input("Today's Date (DD/MM/YYYY) *", value="19/09/2025", key="today_date_trans")
    project_start = col3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025", key="project_start_trans")
    project_end = st.text_input("Project End Date (DD/MM/YYYY) *", value="19/09/2025", key="project_end_trans")
    project_summary_text = st.text_area("Project Summary Text", value="More than half of the users have been deployed and there were not any critical issues. Not expected issues during enrollment of remaining users", key="project_summary_text_trans")
    theme = st.selectbox("Theme", ["White", "Navy"], key="theme_trans")

    # Milestones
    st.header("Milestones")
    milestones_data = []
    milestone_defaults = [
        {"name": "Initial Project Schedule Accepted", "baseline": "27/06/2025", "target": "27/06/2025", "status": ""},
        {"name": "Initial Design Accepted", "baseline": "14/07/2025", "target": "17/07/2025", "status": ""},
        {"name": "Pilot Configuration Complete", "baseline": "28/07/2025", "target": "18/07/2025", "status": ""},
        {"name": "Pilot Rollout Complete", "baseline": "08/08/2025", "target": "22/08/2025", "status": ""},
        {"name": "Production Configuration Complete", "baseline": "29/08/2025", "target": "29/08/2025", "status": ""},
        {"name": "Production Rollout Complete", "baseline": "19/09/2025", "target": "??", "status": ""},
        {"name": "Final Design Accepted", "baseline": "19/09/2025", "target": "19/09/2025", "status": ""}
    ]
    for i in range(7):
        with st.expander(f"Milestone {i+1}", expanded=True):
            name = st.text_input(f"Milestone Name {i+1}", key=f"mname_{i}", value=milestone_defaults[i]["name"])
            baseline = st.text_input(f"Baseline Date {i+1} (DD/MM/YYYY)", key=f"mbaseline_{i}", value=milestone_defaults[i]["baseline"])
            target = st.text_input(f"Target Completion {i+1} (DD/MM/YYYY)", key=f"mtarget_{i}", value=milestone_defaults[i]["target"])
            status = st.text_input(f"Status {i+1} (e.g., Completed)", key=f"mstatus_{i}", value=milestone_defaults[i]["status"])
            if name:
                milestones_data.append({"name": name, "baseline": baseline, "target": target, "status": status})

    # User Rollout Roadmap
    st.header("User Rollout Roadmap")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.subheader("Pilot")
        pilot_target = st.number_input("Pilot Target Users", value=100)
        pilot_current = st.number_input("Pilot Current Users", value=449)
        pilot_completion = st.text_input("Pilot Completion Date", value="19/09/2025")
        pilot_status = st.text_input("Pilot Status", value="")
    with col_p2:
        st.subheader("Production")
        prod_target = st.number_input("Production Target Users", value=800)
        prod_current = st.number_input("Production Current Users", value=449)
        prod_completion = st.text_input("Production Completion Date", value="19/09/2025")
        prod_status = st.text_input("Production Status", value="")

    # Project Objectives
    st.header("Project Objectives")
    objectives_data = []
    objective_defaults = [
        {"objective": "Protect and Secure Internet Access for Users", "actual": "More than half of the users have Zscaler Client Connector deployed and are fully protected when they are outside of the corporate office", "deviation": "Not enough time to deploy ZCC in all users but deployment is on track to be finished by Pixartprinting and no critical issues are expected."},
        {"objective": "Complete user posture", "actual": "Users and devices are identified, and policies can be applied based on this criteria", "deviation": "No deviations"},
        {"objective": "Comprehensive Web filtering", "actual": "Web filtering based on reputation and dynamic categorization rather than simply categories.", "deviation": "No deviations"}
    ]
    for i in range(3):
        with st.expander(f"Objective {i+1}", expanded=True):
            objective = st.text_area(f"Planned Objective {i+1}", key=f"obj_{i}", height=50, value=objective_defaults[i]["objective"])
            actual = st.text_area(f"Actual Result {i+1}", key=f"act_{i}", height=50, value=objective_defaults[i]["actual"])
            deviation = st.text_area(f"Deviation/Cause {i+1}", key=f"dev_{i}", height=50, value=objective_defaults[i]["deviation"])
            if objective:
                objectives_data.append({"objective": objective, "actual": actual, "deviation": deviation})

    # Deliverables
    st.header("Deliverables")
    deliverables_data = []
    deliverable_defaults = [
        {"name": "Kick-Off Meeting and Slides", "date": "27/06/2025"},
        {"name": "Design and Configuration of Zscaler Platform (per scope)", "date": "30/06/2025 – 11/07/2025"},
        {"name": "Troubleshooting Guide(s)", "date": "18/07/2025"},
        {"name": "Initial & Final Design Document", "date": "17/07/2025 – 17/09/2025"},
        {"name": "Transition Meeting Slides", "date": "19/09/2025"}
    ]
    for i in range(5):
        with st.expander(f"Deliverable {i+1}", expanded=True):
            name = st.text_input(f"Deliverable Name {i+1}", key=f"dname_{i}", value=deliverable_defaults[i]["name"])
            date_del = st.text_input(f"Date Delivered {i+1}", key=f"ddate_{i}", value=deliverable_defaults[i]["date"])
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
    open_defaults = [
        {"task": "Finish Production rollout", "date": "October 2025", "owner": "Pixartprinting", "steps": "Onboard remaining users from all departments including Developers."},
        {"task": "Tighten Firewall policies", "date": "October 2025", "owner": "Pixartprinting", "steps": "Change the default Firewall rule from Allow All to Block All after configuring all the required exceptions."},
        {"task": "Tighten Cloud App Control Policies", "date": "October 2025", "owner": "Pixartprinting", "steps": "Configure block policies for high risk applications in all categories."},
        {"task": "Fine tune SSL Inspection policies", "date": "November 2025", "owner": "Pixartprinting", "steps": "Continue adjusting and adding exclusions to SSL Inspection policies as required."},
        {"task": "Configure DLP policies", "date": "December 2025", "owner": "Pixartprinting", "steps": "Configure DLP policies to control sensitive data and avoid potential data leaks."},
        {"task": "Deploy ZCC on Mobile devices", "date": "January 2026", "owner": "Pixartprinting", "steps": "Expand the deployment of Zscaler Client Connector to Mobile devices."}
    ]
    for i in range(6):
        with st.expander(f"Open Item {i+1}", expanded=True):
            task = st.text_input(f"Task/Description {i+1}", key=f"otask_{i}", value=open_defaults[i]["task"])
            o_date = st.text_input(f"Date {i+1}", key=f"odate_{i}", value=open_defaults[i]["date"])
            owner = st.text_input(f"Owner {i+1}", key=f"oowner_{i}", value=open_defaults[i]["owner"])
            steps = st.text_area(f"Transition Plan/Next Steps {i+1}", key=f"osteps_{i}", height=50, value=open_defaults[i]["steps"])
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

team_members = st.text_area("Team Members (Name, Position, one per line)", value="John Doe, Project Manager\nJane Smith, Security Architect")
objectives = st.text_area("Key Objectives (one per line)", value="Objective 1\nObjective 2\nObjective 3")
milestones = st.text_area("Milestones (YYYY-MM-DD: Description, one per line)", value="2025-01-01: Kickoff\n2025-03-01: Implementation\n2025-06-01: Go-Live")

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
    elif not all(is_valid_date(d) for d in [today_date, project_start, project_end] + [m["baseline"] for m in milestones_data if m["baseline"]] + [m["target"] for m in milestones_data if m["target"]] + [pilot_completion, prod_completion] + [d["date"] for d in deliverables_data if d["date"]] + [oi["date"] for oi in open_items_data if oi["date"]]):
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

            # Footer in bottom left
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(7), Inches(3), Inches(0.3))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "2025 Zscaler, Inc. All rights reserved"
            p.alignment = PP_ALIGN.LEFT
            p.font.name = 'Century Gothic'
            p.font.size = Pt(8)
            p.font.color.rgb = WHITE if theme == "White" else NAVY

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
        p.text = f"From {project_start} to {project_end}".capitalize()  # Sentence case
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
        agenda_items = ["Project Summary", "Technical Summary", "Recommended Next Steps"]
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
            p.alignment = PP_ALIGN.LEFT
            top += Inches(0.6)

        # Project Summary Title Slide
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Project Summary".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE

        # Final Project Status Report Slide
        slide = prs.slides.add_slide(blank_layout)
        set_background(slide, theme)
        add_logo_footer(slide, theme)
        # Title
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = f"Final Project Status Report – {customer_name}".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = NAVY if theme == "White" else WHITE
        # Project Summary subtitle
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(11), Inches(0.5))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = "Project Summary".title()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = BLACK if theme == "White" else WHITE
        p.alignment = PP_ALIGN.LEFT
        # Project Summary text
        txBox = slide.shapes.add_textbox(Inches(1), Inches(2.3), Inches(11), Inches(0.5))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = project_summary_text.capitalize()
        p.font.name = 'Century Gothic'
        p.font.size = Pt(16)
        p.font.color.rgb = BLACK if theme == "White" else WHITE
        p.alignment = PP_ALIGN.LEFT
        # Dates table
        table = slide.shapes.add_table(2, 3, Inches(1), Inches(3), Inches(11), Inches(0.5)).table
        table.cell(0,0).text = "Today's Date"
        table.cell(0,1).text = "Start Date"
        table.cell(0,2).text = "End Date"
        table.cell(1,0).text = today_date
        table.cell(1,1).text = project_start
        table.cell(1,2).text = project_end
        for cell in table.iter_cells():
            tf = cell.text_frame
            p = tf.paragraphs[0]
            p.font.name = 'Century Gothic'
            p.font.size = Pt(14)
            p.alignment = PP_ALIGN.LEFT
            if cell in table.rows[0].cells:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                p.font.color.rgb = WHITE
                p.font.bold = True
            else:
                p.font.color.rgb = BLACK if theme == "White" else WHITE
        # Milestones table
        milestones_headers = ["Milestone", "Baseline Date", "Target Completion Date", "Status"]
        milestones_rows = [[m["name"], m["baseline"], m["target"], m["status"]] for m in milestones_data]
        table = slide.shapes.add_table(len(milestones_rows) + 1, 4, Inches(1), Inches(3.8), Inches(11), Inches(2)).table
        for i, header in enumerate(milestones_headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            tf = cell.text_frame
            p = tf.paragraphs[0]
            p.font.name = 'Century Gothic'
            p.font.color.rgb = WHITE
            p.font.bold = True
            p.font.size = Pt(14)
            p.alignment = PP_ALIGN.LEFT
        for row_idx, row in enumerate(milestones_rows, 1):
            for col_idx, text in enumerate(row):
                cell = table.cell(row_idx, col_idx)
                cell.text = text
                tf = cell.text_frame
                p = tf.paragraphs[0]
                p.font.name = 'Century Gothic'
                p.font.size = Pt(12)
                p.font.color.rgb = BLACK if theme == "White" else WHITE
                p.alignment = PP_ALIGN.LEFT
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT_GRAY
        # User Rollout table
        rollout_headers = ["Milestone", "Target Users", "Current Users", "Target Completion", "Status"]
        rollout_rows = [
            ["Pilot", str(pilot_target), str(pilot_current), pilot_completion, pilot_status],
            ["Production", str(prod_target), str(prod_current), prod_completion, prod_status]
        ]
        table = slide.shapes.add_table(3, 5, Inches(1), Inches(6), Inches(11), Inches(1)).table
        for i, header in enumerate(rollout_headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            tf = cell.text_frame
            p = tf.paragraphs[0]
            p.font.name = 'Century Gothic'
            p.font.color.rgb = WHITE
            p.font.bold = True
            p.font.size = Pt(14)
            p.alignment = PP_ALIGN.LEFT
        for row_idx, row in enumerate(rollout_rows, 1):
            for col_idx, text in enumerate(row):
                cell = table.cell(row_idx, col_idx)
                cell.text = text
                tf = cell.text_frame
                p = tf.paragraphs[0]
                p.font.name = 'Century Gothic'
                p.font.size = Pt(12)
                p.font.color.rgb = BLACK if theme == "White" else WHITE
                p.alignment = PP_ALIGN.LEFT
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT_GRAY
        # Objectives table
        objectives_headers = ["Planned Project Objective (Target)", "Actual Project Result (Actual)", "Deviation/ Cause"]
        objectives_rows = [[o["objective"], o["actual"], o["deviation"]] for o in objectives_data]
        table = slide.shapes.add_table(len(objectives_rows) + 1, 3, Inches(1), Inches(7.1), Inches(11), Inches(1.5)).table
        for i, header in enumerate(objectives_headers):
            cell = table.cell(0, i)
            cell.text = header
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
            tf = cell.text_frame
            p = tf.paragraphs[0]
            p.font.name = 'Century Gothic'
            p.font.color.rgb = WHITE
            p.font.bold = True
            p.font.size = Pt(14)
            p.alignment = PP_ALIGN.LEFT
        for row_idx, row in enumerate(objectives_rows, 1):
            for col_idx, text in enumerate(row):
                cell = table.cell(row_idx, col_idx)
                cell.text = text
                tf = cell.text_frame
                p = tf.paragraphs[0]
                p.font.name = 'Century Gothic'
                p.font.size = Pt(12)
                p.font.color.rgb = BLACK if theme == "White" else WHITE
                p.alignment = PP_ALIGN.LEFT
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT_GRAY
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 5: Deliverables
        deliverables_headers = ["Deliverable", "Date delivered"]
        deliverables_rows = [[d["name"], d["date"]] for d in deliverables_data]
        deliverables_slide = add_table_slide("Deliverables", len(deliverables_rows) + 1, 2, [deliverables_headers] + deliverables_rows, top_inch=1, height_inch=2)
        # Add checkmarks
        for row_idx in range(1, len(deliverables_rows) + 1):
            txBox = deliverables_slide.shapes.add_textbox(Inches(0.3), Inches(1) + Inches(0.3) * (row_idx-1), Inches(0.3), Inches(0.3))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = '✓'
            p.font.name = 'Century Gothic'
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(0, 176, 80)  # Green
            p.alignment = PP_ALIGN.CENTER
        # RAG Status Key
        rag_box = deliverables_slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(1.5))
        rag_tf = rag_box.text_frame
        rag_tf.text = "Who: External & Internal Project Team \nWhat: Project Status Report\nWhen: Weekly\nWhy: Keeps project stakeholders informed on a weekly basis on critical aspects of the project such as scope, schedule, risks, issues, and next steps. \nMandatory: Yes (all projects)\n\nRAG Status Key:\nRed - Not On Track\nAmber - At Risk\nGreen - On Track\nBlue - Complete\nGray - Not Started"
        for para in rag_tf.paragraphs:
            para.font.name = 'Century Gothic'
            para.font.size = Pt(14)
            para.font.color.rgb = BLACK
            para.alignment = PP_ALIGN.LEFT
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 6: Technical Summary Title
        technical_slide = add_title_slide("Technical Summary")
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 7: Deployed ZIA Architecture
        zia_slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
        set_background(zia_slide)
        # Title
        txBox = zia_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.5))
        tf = txBox.text_frame
        tf.text = "Deployed ZIA Architecture".title()
        tf.paragraphs[0].font.name = 'Century Gothic'
        tf.paragraphs[0].font.size = Pt(28)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = NAVY
        # Diagram
        # User authentication 
        user_auth = zia_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(2.5), Inches(1))
        user_auth.fill.solid()
        user_auth.fill.fore_color.rgb = LIGHT_GRAY
        user_auth.line.color.rgb = NAVY
        user_auth.text_frame.text = "User authentication \nand provisioning"
        user_auth.text_frame.paragraphs[0].font.name = 'Century Gothic'
        user_auth.text_frame.paragraphs[0].font.size = Pt(12)
        user_auth.text_frame.paragraphs[0].font.color.rgb = BLACK
        # Central Authority
        central = zia_slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.5), Inches(1.5), Inches(2.5), Inches(1))
        central.fill.solid()
        central.fill.fore_color.rgb = BRIGHT_BLUE
        central.line.color.rgb = NAVY
        central.text_frame.text = "Central Authority"
        central.text_frame.paragraphs[0].font.name = 'Century Gothic'
        central.text_frame.paragraphs[0].font.size = Pt(12)
        central.text_frame.paragraphs[0].font.color.rgb = WHITE
        # Add more as needed...
        # Key facts
        key_box = zia_slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(12), Inches(3))
        key_tf = key_box.text_frame
        key_tf.text = f"Authentication Type\nIdentity Provider\t{idp}\nAuthentication Type\t{auth_type}\nUser and Group Provisioning\t{prov_type}\n\nClient Deployment\nTunnel Type\t{tunnel_type}\nZCC Deployment System\t{deploy_system}\nNumber of Windows and MacOS Devices\t{windows_num} Windows Devices\n\t{mac_num} MacOS Devices\nGeo Locations\t{geo_locations}\n\nPolicy Deployment\nSSL Inspection Policies\t{ssl_policies}\nURL Filtering Policies\t{url_policies}\nCloud App Control Policies\t{cloud_policies}\nFirewall Policies\t{fw_policies}"
        for p in key_tf.paragraphs:
            p.font.name = 'Century Gothic'
            p.font.size = Pt(12)
            p.font.color.rgb = BLACK
            p.alignment = PP_ALIGN.LEFT
        # Overview
        overview_box = zia_slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(12), Inches(0.5))
        overview_tf = overview_box.text_frame
        overview_tf.text = "An overview of the deployed architecture and key facts - diagram stays generic (custom diagram will be in design document) Numbers on the diagram help to orient the conversation,"
        overview_tf.paragraphs[0].font.name = 'Century Gothic'
        overview_tf.paragraphs[0].font.size = Pt(12)
        overview_tf.paragraphs[0].font.color.rgb = BLACK
        add_logo_footer_number(zia_slide, len(prs.slides))
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 8: Open Items
        open_items_headers = ["Task/ Description", "Date", "Owner", "Transition Plan/ Next Steps"]
        open_items_rows = [[oi["task"], oi["date"], oi["owner"], oi["steps"]] for oi in open_items_data]
        open_items_slide = add_table_slide("Open Items", len(open_items_rows) + 1, 4, [open_items_headers] + open_items_rows)
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 9: Recommended Next Steps
        next_steps_slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
        set_background(next_steps_slide)
        # Title
        txBox = next_steps_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.5))
        tf = txBox.text_frame
        tf.text = "Recommended Next Steps".title()
        tf.paragraphs[0].font.name = 'Century Gothic'
        tf.paragraphs[0].font.size = Pt(28)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = NAVY
        # Short Term
        short_title = next_steps_slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(6), Inches(0.5))
        short_tf = short_title.text_frame
        short_tf.text = "Short Term Activities"
        short_tf.paragraphs[0].font.name = 'Century Gothic'
        short_tf.paragraphs[0].font.size = Pt(18)
        short_tf.paragraphs[0].font.bold = True
        short_tf.paragraphs[0].font.color.rgb = NAVY
        top = Inches(1.5)
        for item in short_term:
            # Square bullet
            shape = next_steps_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), top + Inches(0.1), Inches(0.2), Inches(0.2))
            shape.fill.solid()
            shape.fill.fore_color.rgb = ACCENT_GREEN
            shape.line.color.rgb = ACCENT_GREEN
            # Text
            txBox = next_steps_slide.shapes.add_textbox(Inches(0.8), top, Inches(5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = item + "."
            p.font.name = 'Century Gothic'
            p.font.size = Pt(14)
            p.font.color.rgb = BLACK
            p.alignment = PP_ALIGN.LEFT
            top += Inches(0.4)
        # Long Term
        long_title = next_steps_slide.shapes.add_textbox(Inches(6.5), Inches(1), Inches(6), Inches(0.5))
        long_tf = long_title.text_frame
        long_tf.text = "Long Term Activities"
        long_tf.paragraphs[0].font.name = 'Century Gothic'
        long_tf.paragraphs[0].font.size = Pt(18)
        long_tf.paragraphs[0].font.bold = True
        long_tf.paragraphs[0].font.color.rgb = NAVY
        top = Inches(1.5)
        for item in long_term:
            # Square bullet
            shape = next_steps_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.5), top + Inches(0.1), Inches(0.2), Inches(0.2))
            shape.fill.solid()
            shape.fill.fore_color.rgb = CYAN
            shape.line.color.rgb = CYAN
            # Text
            txBox = next_steps_slide.shapes.add_textbox(Inches(6.8), top, Inches(5), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = item + "."
            p.font.name = 'Century Gothic'
            p.font.size = Pt(14)
            p.font.color.rgb = BLACK
            p.alignment = PP_ALIGN.LEFT
            top += Inches(0.4)
        # Note text
        note_box = next_steps_slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(9), Inches(0.5))
        note_tf = note_box.text_frame
        note_tf.text = "Next Short- and Long-Term Activities\nIf additional resources and/or expertise are required to complete any of the recommendations above, customer should consider engaging Zscaler Professional Services to assist with this effort."
        note_tf.paragraphs[0].font.name = 'Century Gothic'
        note_tf.paragraphs[0].font.size = Pt(14)
        note_tf.paragraphs[0].font.color.rgb = BLACK
        note_tf.paragraphs[0].alignment = PP_ALIGN.LEFT
        add_logo_footer_number(next_steps_slide, len(prs.slides))
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 10: Thank You with feedback
        thank_slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
        set_background(thank_slide)
        # Title
        txBox = thank_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.5))
        tf = txBox.text_frame
        tf.text = "Thank you".title()
        tf.paragraphs[0].font.name = 'Century Gothic'
        tf.paragraphs[0].font.size = Pt(36)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = NAVY
        # Body
        body_box = thank_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(3))
        body_tf = body_box.text_frame
        body_tf.text = f"Your feedback on our project and Professional Services team is important to us. \nProject Manager: {pm_name}\nConsultant: {consultant_name}\n\nA short ~6 question survey on how your Professional Services team did will be automatically sent after the project has closed. The following people will receive the survey via email:\nPrimary Contact: {primary_contact}\nSecondary Contact: {secondary_contact}\nWe appreciate any insights you can provide to help us improve our processes and ensure we provide the best possible service in future projects."
        for para in body_tf.paragraphs:
            para.font.name = 'Century Gothic'
            para.font.size = Pt(14)
            para.font.color.rgb = BLACK
            para.alignment = PP_ALIGN.LEFT
        # Speech bubble
        bubble = thank_slide.shapes.add_shape(MSO_SHAPE.CLOUD_CALLOUT, Inches(7), Inches(3), Inches(2), Inches(1))
        bubble.fill.solid()
        bubble.fill.fore_color.rgb = RGBColor(255, 192, 0) # Yellow
        bubble_tf = bubble.text_frame
        bubble_tf.text = "We want to know!"
        bubble_tf.paragraphs[0].font.name = 'Century Gothic'
        bubble_tf.paragraphs[0].font.size = Pt(18)
        bubble_tf.paragraphs[0].font.color.rgb = BLACK
        bubble_tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        add_logo_footer_number(thank_slide, len(prs.slides))
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Slide 11: Final Thank You
        final_thank_slide = add_title_slide("Thank you")
        current_slide += 1
        progress.progress(current_slide / total_slides)
        # Save to buffer and provide download
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        st.download_button("Download Transition Deck", bio, file_name=f"{customer_name}_Zscaler_Transition.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
