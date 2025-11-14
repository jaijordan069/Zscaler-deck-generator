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
    # Transition deck inputs (from previous code)
    col1, col2, col3 = st.columns(3)
    customer_name_trans = col1.text_input("Customer Name *", value="Pixartprinting", key="customer_name_trans")
    today_date_trans = col2.text_input("Today's Date (DD/MM/YYYY) *", value="14/11/2025", key="today_date_trans")
    project_start_trans = col3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025", key="project_start_trans")
    project_end_trans = st.text_input("Project End Date (DD/MM/YYYY) *", value="14/11/2025", key="project_end_trans")
    project_summary_text_trans = st.text_area("Project Summary Text", value="More than half of the users have been deployed and there were not any critical issues. Not expected issues during enrollment of remaining users", key="project_summary_text_trans")
    theme_trans = st.selectbox("Theme", ["White", "Navy"], key="theme_trans")

    # Milestones (transition)
    st.header("Milestones")
    milestones_data_trans = []
    milestone_defaults_trans = [
        {"name": "Initial Project Schedule Accepted", "baseline": "27/06/2025", "target": "27/06/2025", "status": ""},
        {"name": "Initial Design Accepted", "baseline": "14/07/2025", "target": "17/07/2025", "status": ""},
        {"name": "Pilot Configuration Complete", "baseline": "28/07/2025", "target": "18/07/2025", "status": ""},
        {"name": "Pilot Rollout Complete", "baseline": "08/08/2025", "target": "22/08/2025", "status": ""},
        {"name": "Production Configuration Complete", "baseline": "29/08/2025", "target": "29/08/2025", "status": ""},
        {"name": "Production Rollout Complete", "baseline": "14/11/2025", "target": "??", "status": ""},
        {"name": "Final Design Accepted", "baseline": "14/11/2025", "target": "14/11/2025", "status": ""}
    ]
    for i in range(7):
        with st.expander(f"Milestone {i+1}", expanded=True):
            name = st.text_input(f"Milestone Name {i+1}", key=f"trans_mname_{i}", value=milestone_defaults_trans[i]["name"])
            baseline = st.text_input(f"Baseline Date {i+1} (DD/MM/YYYY)", key=f"trans_mbaseline_{i}", value=milestone_defaults_trans[i]["baseline"])
            target = st.text_input(f"Target Completion {i+1} (DD/MM/YYYY)", key=f"trans_mtarget_{i}", value=milestone_defaults_trans[i]["target"])
            status = st.text_input(f"Status {i+1} (e.g., Completed)", key=f"trans_mstatus_{i}", value=milestone_defaults_trans[i]["status"])
            if name:
                milestones_data_trans.append({"name": name, "baseline": baseline, "target": target, "status": status})

    # User Rollout Roadmap (transition)
    st.header("User Rollout Roadmap")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.subheader("Pilot")
        pilot_target_trans = st.number_input("Pilot Target Users", value=100, key="pilot_target_trans")
        pilot_current_trans = st.number_input("Pilot Current Users", value=449, key="pilot_current_trans")
        pilot_completion_trans = st.text_input("Pilot Completion Date", value="14/11/2025", key="pilot_completion_trans")
        pilot_status_trans = st.text_input("Pilot Status", value="", key="pilot_status_trans")
    with col_p2:
        st.subheader("Production")
        prod_target_trans = st.number_input("Production Target Users", value=800, key="prod_target_trans")
        prod_current_trans = st.number_input("Production Current Users", value=449, key="prod_current_trans")
        prod_completion_trans = st.text_input("Production Completion Date", value="14/11/2025", key="prod_completion_trans")
        prod_status_trans = st.text_input("Production Status", value="", key="prod_status_trans")

    # Project Objectives (transition)
    st.header("Project Objectives")
    objectives_data_trans = []
    objective_defaults_trans = [
        {"objective": "Protect and Secure Internet Access for Users", "actual": "More than half of the users have Zscaler Client Connector deployed and are fully protected when they are outside of the corporate office", "deviation": "Not enough time to deploy ZCC in all users but deployment is on track to be finished by Pixartprinting and no critical issues are expected."},
        {"objective": "Complete user posture", "actual": "Users and devices are identified, and policies can be applied based on this criteria", "deviation": "No deviations"},
        {"objective": "Comprehensive Web filtering", "actual": "Web filtering based on reputation and dynamic categorization rather than simply categories.", "deviation": "No deviations"}
    ]
    for i in range(3):
        with st.expander(f"Objective {i+1}", expanded=True):
            objective = st.text_area(f"Planned Objective {i+1}", key=f"trans_obj_{i}", height=50, value=objective_defaults_trans[i]["objective"])
            actual = st.text_area(f"Actual Result {i+1}", key=f"trans_act_{i}", height=50, value=objective_defaults_trans[i]["actual"])
            deviation = st.text_area(f"Deviation/Cause {i+1}", key=f"trans_dev_{i}", height=50, value=objective_defaults_trans[i]["deviation"])
            if objective:
                objectives_data_trans.append({"objective": objective, "actual": actual, "deviation": deviation})

    # Deliverables (transition)
    st.header("Deliverables")
    deliverables_data_trans = []
    deliverable_defaults_trans = [
        {"name": "Kick-Off Meeting and Slides", "date": "27/06/2025"},
        {"name": "Design and Configuration of Zscaler Platform (per scope)", "date": "30/06/2025 – 11/07/2025"},
        {"name": "Troubleshooting Guide(s)", "date": "18/07/2025"},
        {"name": "Initial & Final Design Document", "date": "17/07/2025 – 17/09/2025"},
        {"name": "Transition Meeting Slides", "date": "19/09/2025"}
    ]
    for i in range(5):
        with st.expander(f"Deliverable {i+1}", expanded=True):
            name = st.text_input(f"Deliverable Name {i+1}", key=f"trans_dname_{i}", value=deliverable_defaults_trans[i]["name"])
            date_del = st.text_input(f"Date Delivered {i+1}", key=f"trans_ddate_{i}", value=deliverable_defaults_trans[i]["date"])
            if name:
                deliverables_data_trans.append({"name": name, "date": date_del})

    # Technical Summary (transition)
    st.header("Technical Summary")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.subheader("Authentication & Provisioning")
        idp_trans = st.text_input("Identity Provider", value="Entra ID", key="idp_trans")
        auth_type_trans = st.text_input("Authentication Type", value="SAML 2.0", key="auth_type_trans")
        prov_type_trans = st.text_input("User/Group Provisioning", value="SCIM Provisioning", key="prov_type_trans")
    with col_t2:
        st.subheader("Client Deployment")
        tunnel_type_trans = st.text_input("Tunnel Type", value="ZCC with Z-Tunnel 2.0", key="tunnel_type_trans")
        deploy_system_trans = st.text_input("ZCC Deployment System", value="MS Intune/Jamf", key="deploy_system_trans")
    col_d1, col_d2, col_d3 = st.columns(3)
    windows_num_trans = col_d1.number_input("Number of Windows Devices", value=351, key="windows_num_trans")
    mac_num_trans = col_d2.number_input("Number of MacOS Devices", value=98, key="mac_num_trans")
    geo_locations_trans = col_d3.text_input("Geo Locations", value="Europe, North Africa, USA", key="geo_locations_trans")
    col_pol1, col_pol2, col_pol3, col_pol4 = st.columns(4)
    ssl_policies_trans = col_pol1.number_input("SSL Inspection Policies", value=10, key="ssl_policies_trans")
    url_policies_trans = col_pol2.number_input("URL Filtering Policies", value=5, key="url_policies_trans")
    cloud_policies_trans = col_pol3.number_input("Cloud App Control Policies", value=5, key="cloud_policies_trans")
    fw_policies_trans = col_pol4.number_input("Firewall Policies", value=15, key="fw_policies_trans")

    # Open Items (transition)
    st.header("Open Items")
    open_items_data_trans = []
    open_defaults_trans = [
        {"task": "Finish Production rollout", "date": "October 2025", "owner": "Pixartprinting", "steps": "Onboard remaining users from all departments including Developers."},
        {"task": "Tighten Firewall policies", "date": "October 2025", "owner": "Pixartprinting", "steps": "Change the default Firewall rule from Allow All to Block All after configuring all the required exceptions."},
        {"task": "Tighten Cloud App Control Policies", "date": "October 2025", "owner": "Pixartprinting", "steps": "Configure block policies for high risk applications in all categories."},
        {"task": "Fine tune SSL Inspection policies", "date": "November 2025", "owner": "Pixartprinting", "steps": "Continue adjusting and adding exclusions to SSL Inspection policies as required."},
        {"task": "Configure DLP policies", "date": "December 2025", "owner": "Pixartprinting", "steps": "Configure DLP policies to control sensitive data and avoid potential data leaks."},
        {"task": "Deploy ZCC on Mobile devices", "date": "January 2026", "owner": "Pixartprinting", "steps": "Expand the deployment of Zscaler Client Connector to Mobile devices."}
    ]
    for i in range(6):
        with st.expander(f"Open Item {i+1}", expanded=True):
            task = st.text_input(f"Task/Description {i+1}", key=f"trans_otask_{i}", value=open_defaults_trans[i]["task"])
            o_date = st.text_input(f"Date {i+1}", key=f"trans_odate_{i}", value=open_defaults_trans[i]["date"])
            owner = st.text_input(f"Owner {i+1}", key=f"trans_oowner_{i}", value=open_defaults_trans[i]["owner"])
            steps = st.text_area(f"Transition Plan/Next Steps {i+1}", key=f"trans_osteps_{i}", height=50, value=open_defaults_trans[i]["steps"])
            if task:
                open_items_data_trans.append({"task": task, "date": o_date, "owner": owner, "steps": steps})

    # Recommended Next Steps (transition)
    st.header("Recommended Next Steps")
    st.subheader("Short Term Activities")
    short_term_input_trans = st.text_area("Short Term (comma-separated)", value="Finish Production rollout, Tighten Firewall policies, Tighten Cloud App Control Policies, Fine tune SSL Inspection policies, Configure Role Based Access Control (RBAC), Configure DLP policies", key="short_term_input_trans")
    short_term_trans = [item.strip() for item in short_term_input_trans.split(",") if item.strip()]
    st.subheader("Long Term Activities")
    long_term_input_trans = st.text_area("Long Term (comma-separated)", value="Deploy ZCC on Mobile devices, Consider an upgrade of Sandbox license to have better antimalware protection, Consider an upgrade of the Firewall License to be able to apply policies based on user groups and network applications, Adopt additional Zscaler solutions like Zscaler Private Access (ZPA) or Zscaler Digital experience (ZDX), Consider using ZCC Client when the users are on-prem for a more consistent user experience, Integrate ZIA with 3rd party SIEM", key="long_term_input_trans")
    long_term_trans = [item.strip() for item in long_term_input_trans.split(",") if item.strip()]

    # Contacts (transition)
    st.header("Contacts")
    col_c1, col_c2 = st.columns(2)
    pm_name_trans = col_c1.text_input("Project Manager Name", value="Alex Vazquez", key="pm_name_trans")
    consultant_name_trans = col_c2.text_input("Consultant Name", value="Alex Vazquez", key="consultant_name_trans")
    primary_contact_trans = st.text_input("Primary Contact", value="Teia proctor", key="primary_contact_trans")
    secondary_contact_trans = st.text_input("Secondary Contact", value="Marco Sattier", key="secondary_contact_trans")

with tab2:
    st.header("Kick-Off Deck Inputs")
    # Kick-Off deck inputs based on template
    customer_name_kick = st.text_input("Company Name *", value="ACME Corp", key="customer_name_kick")
    pm_name_kick = st.text_input("PM Name *", value="PM Name", key="pm_name_kick")
    psc_name_kick = st.text_input("PSC Name *", value="PSC Name", key="psc_name_kick")
    date_kick = st.text_input("Date (DD/MM/YYYY) *", value="14/11/2025", key="date_kick")
    theme_kick = st.selectbox("Theme", ["White", "Navy"], key="theme_kick")

    # Team Contacts (kick-off)
    st.header("Team Contacts")
    team_contacts_kick = []
    team_defaults_kick = [
        {"name": "Name", "role": "Professional Services Project Manager", "responsibility": "Main point of contact during project\nManages project to success\nCoordinate between ACME Corp and Zscaler"},
        {"name": "Name", "role": "Professional Services Consultant", "responsibility": "Review scope and timelines\nLead the project technical delivery"},
        {"name": "Name", "role": "Technical Success Manager", "responsibility": "Drive product adoption and use case execution\nEnsure positive customer experience, deliverable execution and operational best practices."},
        {"name": "Name", "role": "Account Team", "responsibility": "Validate customer use cases from pre-sales activities\nAccount management, Zscaler sponsor and escalation point"},
        {"name": "Name", "role": "Project Sponsor", "responsibility": "Responsible for the success and benefits realization of a project\nEnsures the project aligns with business goals, strategy, and objectives\nSupport the Project Manager to manage risks as they arise"},
        {"name": "Name", "role": "Project Manager", "responsibility": "Oversees the project, addresses blockers and is the initial escalation point for the team"},
        {"name": "Name", "role": "Network Engineer, IT", "responsibility": "Establishes and maintains network performance\nBuilds network configurations and connections\nTroubleshoots network problems"},
        {"name": "Name", "role": "Security Architect, IT", "responsibility": "Designs, implements, and manages security measures"},
        {"name": "Name", "role": "Application Owners", "responsibility": "Owns and manages applications"},
        {"name": "Name", "role": "Director, IT Infrastructure", "responsibility": "Focuses on security, infrastructure scalability, and employee productivity.\nLead functions,\nGlobal IT Support, Production services and DevOps, Data center and Cloud operations"}
    ]
    for i in range(10):
        with st.expander(f"Team Contact {i+1}", expanded=True):
            name = st.text_input(f"Name {i+1}", key=f"kick_name_{i}", value=team_defaults_kick[i]["name"])
            role = st.text_input(f"Role {i+1}", key=f"kick_role_{i}", value=team_defaults_kick[i]["role"])
            responsibility = st.text_area(f"Responsibility {i+1}", key=f"kick_resp_{i}", value=team_defaults_kick[i]["responsibility"])
            if name:
                team_contacts_kick.append({"name": name, "role": role, "responsibility": responsibility})

    # Project Overview (kick-off)
    st.header("Project Overview")
    scope_kick = st.text_area("Scope", value="To protect and secure internet access and mission critical applications with visibility into performance metrics by implementing Zscaler licensed products which include: \nZscaler Internet Access (ZIA)\nZscaler Private Access (ZPA)\nZscaler Digital Experience (ZDX) \nData Protection (DLP)")
    deliverables_kick = st.text_area("Deliverables", value="Kick-Off \nProject Kick-Off Deck\nPrerequisites Document(s)\n\nDesign and Configure\nInstallation Guides (Zscaler Client Connector and App Connector)\nDesign Document(s)\n\nPilot  \nPilot Plan(s) - (Test and Rollout)\nCommunications Guide\nTroubleshooting Guide(s)\n\nProject Closure\nTransition with recommended Next Steps")
    overview_kick = st.text_area("Overview", value="Scope alignment, review the current environment, and provides a blueprint of the Design \nProvide task-level guidance for each phase of the project in alignment with Zscalers leading practices\nConfiguration Sessions for each licensed product in scope\nStandardized templates provide security policies configuration guidance \nPilot roll-out support up to the first 50 users and initial Production roll-out support (up to 2 weeks) within the project duration.\nProject transition")
    zscaler_resources_kick = st.text_area("Zscaler Resources", value="Primary PS Consultant\nProviding advisory and leading practices for the design of your Zscaler solution")
    engagement_kick = st.text_area("Engagement", value="Duration: up to 90 days  or completion of project tasks. 	All resources and services are remote")
    out_of_scope_kick = st.text_area("Out of Scope", value="Advisory, deployment, and consulting activities not\nexplicitly stated in the project definition are deemed out of scope. Additional PS can be purchased to increase scope.")
    ps_package_kick = st.text_area("Professional Services Package + Terms & Conditions", value="")

    # Success Criteria (kick-off)
    st.header("Success Criteria")
    success_criteria_kick = st.text_area("Success Criteria", value="Approve baseline project schedule \nDesign Workshops\nInitial Design\nConfigure & Testing\nPilot\nProduction Support\nTransition")

    # Project Goals & Objectives (kick-off)
    st.header("Project Goals & Objectives")
    goals_kick = st.text_area("Goals & Objectives", value="Cyber Threat Protection\nPrevent compromise & lateral movement\nUse Case: Secure Workforce/Remote Users\n<Details> be specific - what is their pain point\n\nData Protection \nPrevent data loss, inline and API\nUse Case: Prevent Inline Data Loss … \n<Details> -  be specific - what is their pain point\n\nOperationalization\nTroubleshoot and operations\nUse Case: Enable Help Desk to Troubleshoot connectivity issues to rapidly fix application, network, and device issues\n<Details> -  be specific - what is their pain point\n\nZero Trust Networking\nConnect to apps, not networks \nUse Case: Secure Mission Critical Apps - SFDC/Workday\n<Details> be specific - what is their pain point\nPriority 1: Protect and Secure Internet Access for Users \n\nPriority 2: Protect Mission Critical Application Access\n\nPriority 3: Visibility into performance metrics")

    # Project Timeline (kick-off)
    st.header("Project Timeline")
    project_timeline_kick = st.text_area("Project Timeline", value="Initiate\nToday\n\nPlan\nFROM-TO DATES\n\nConfigure and Testing\nFROM-TO DATES\n\nPilot\nFROM-TO DATES\n\nProduction Support\nFROM-TO DATES\n\nTransition\nFROM-TO DATES\n\nAssessment\nAuthentication\nTraffic Forwarding - ZCC/App Connectors\nSecure Internet\nDesign Workshops\nProject Planning\nInline DLP\nLogging/Telemetry\nSecure Application Access\nProduction\nPilot\nTransition\nPilot & Production Rollout Planning\nProject Kick-Off")

    # Change Management & Key Milestones (kick-off)
    st.header("Change Management & Key Milestones")
    change_mgmt_kick = st.text_area("Change Management & Key Milestones", value="Key Milestones\nApprove baseline project schedule 	Proposed Completion Date\nDesign Workshops	Proposed Completion Date\nInitial Design	Proposed Completion Date\nConfigure & Testing	Proposed Completion Date\nPilot	Proposed Completion Date\nProduction Support	Proposed Completion Date\nTransition	Proposed Completion Date")

    # Next Steps & Follow-up Items (kick-off)
    st.header("Next Steps & Follow-up Items")
    next_steps_kick = st.text_area("Next Steps & Follow-up Items", value="")

# Output format
st.header("Output Format")
output_format = st.selectbox("Select Output Format", ["PPT", "DOCX", "PDF"])

# Preview Summary
if st.button("Preview Summary"):
    st.write(f"Deck for {customer_name if deck_type == "Transition Deck" else customer_name_kick} on {today_date if deck_type == "Transition Deck" else date_kick}:")
    # Add preview logic for both decks
    # ... (similar to previous)

# Generate button
if st.button("Generate Deck"):
    # Validation (for both decks)
    if deck_type == "Transition Deck":
        if not customer_name:
            st.error("Customer Name is required.")
        elif not all(is_valid_date(d) for d in [today_date, project_start, project_end] + [m["baseline"] for m in milestones_data if m["baseline"]] + [m["target"] for m in milestones_data if m["target"]] + [pilot_completion, prod_completion] + [d["date"] for d in deliverables_data if d["date"]] + [oi["date"] for oi in open_items_data if oi["date"]]):
            st.error("All dates must be in DD/MM/YYYY format.")
        else:
            # Generate Transition Deck PPT
            prs = Presentation()
            # ... (existing transition deck generation logic from previous code, without instructional slides)
            # Remove unnecessary slides (none in transition)
            bio = io.BytesIO()
            prs.save(bio)
            bio.seek(0)
            if output_format == "PPT":
                st.download_button("Download Transition Deck", bio, file_name=f"{customer_name}_Transition_Deck.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            elif output_format == "DOCX":
                doc = Document()
                # Convert PPT to DOCX (simplified - add sections for each slide)
                doc.add_heading("Transition Deck", 0)
                # Add content from slides to doc
                doc.save(bio)
                st.download_button("Download Transition Deck", bio, file_name=f"{customer_name}_Transition_Deck.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            elif output_format == "PDF":
                pdf = FPDF()
                # Convert to PDF (simplified - add text from slides)
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt="Transition Deck", ln=1, align='C')
                # Add content
                pdf.output(bio)
                st.download_button("Download Transition Deck", bio, file_name=f"{customer_name}_Transition_Deck.pdf", mime="application/pdf")
    else:
        if not customer_name_kick or not pm_name_kick or not psc_name_kick:
            st.error("Required fields are missing.")
        elif not is_valid_date(date_kick):
            st.error("Date must be in DD/MM/YYYY format.")
        else:
            # Generate Kick-Off Deck PPT
            prs = Presentation()
            # ... (new kick-off deck generation logic based on template)
            # Remove instructional slides (e.g., Guidelines, any reference slides)
            # Slide 1: Title
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(slide, theme_kick)
            add_logo_footer(slide, theme_kick)
            txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(11), Inches(2))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Professional Services Project Kick-Off Meeting".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = NAVY
            p.alignment = PP_ALIGN.LEFT
            # Subtitle
            txBox = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(11), Inches(1))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = customer_name_kick.capitalize()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.color.rgb = NAVY
            p.alignment = PP_ALIGN.LEFT
            # PM/PSC
            txBox = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = f"{pm_name_kick}\n{psc_name_kick}".capitalize()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(20)
            p.font.color.rgb = NAVY
            p.alignment = PP_ALIGN.LEFT
            # Date
            txBox = slide.shapes.add_textbox(Inches(1), Inches(5.5), Inches(11), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = date_kick
            p.font.name = 'Century Gothic'
            p.font.size = Pt(18)
            p.font.color.rgb = NAVY
            p.alignment = PP_ALIGN.LEFT

            # Slide 2: Agenda
            add_bullet_slide("Agenda", ["Introductions", "Project Overview", "Success Criteria", "Project Methodology & Timeline", "Change Management & Key Milestones", "Next Steps & Follow-up Items"])

            # Slide 3: Team Contacts
            team_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(team_slide)
            add_logo_footer(team_slide, theme_kick)
            txBox = team_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Team Contacts".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            table = team_slide.shapes.add_table(len(team_contacts_kick) + 1, 3, Inches(0.5), Inches(1.5), Inches(12), Inches(3)).table
            table.cell(0,0).text = "Name"
            table.cell(0,1).text = "Role"
            table.cell(0,2).text = "Responsibility"
            for row_idx, contact in enumerate(team_contacts_kick, 1):
                table.cell(row_idx,0).text = contact["name"]
                table.cell(row_idx,1).text = contact["role"]
                table.cell(row_idx,2).text = contact["responsibility"]
            for cell in table.iter_cells():
                tf = cell.text_frame
                p = tf.paragraphs[0]
                p.font.name = 'Century Gothic'
                p.font.size = Pt(12)
                p.alignment = PP_ALIGN.LEFT
                if cell in table.rows[0].cells:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = NAVY
                    p.font.color.rgb = WHITE
                    p.font.bold = True
                else:
                    p.font.color.rgb = BLACK
                if row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = LIGHT_GRAY

            # Slide 4: Project Overview
            overview_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(overview_slide)
            add_logo_footer(overview_slide, theme_kick)
            txBox = overview_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Project Overview - Essential Services".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Scope
            scope_box = overview_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4), Inches(2))
            scope_tf = scope_box.text_frame
            scope_tf.text = "Scope\n" + scope_kick
            for para in scope_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT
            # Deliverables
            del_box = overview_slide.shapes.add_textbox(Inches(4.5), Inches(1.5), Inches(4), Inches(2))
            del_tf = del_box.text_frame
            del_tf.text = "Deliverables\n" + deliverables_kick
            for para in del_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT
            # Overview
            overview_box = overview_slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(4), Inches(2))
            overview_tf = overview_box.text_frame
            overview_tf.text = "Overview\n" + overview_kick
            for para in overview_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT
            # Zscaler Resources
            res_box = overview_slide.shapes.add_textbox(Inches(4.5), Inches(3.5), Inches(4), Inches(2))
            res_tf = res_box.text_frame
            res_tf.text = "Zscaler Resources\n" + zscaler_resources_kick
            for para in res_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT
            # Engagement
            eng_box = overview_slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(4), Inches(1))
            eng_tf = eng_box.text_frame
            eng_tf.text = "Engagement\n" + engagement_kick
            for para in eng_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT
            # Out of Scope
            out_box = overview_slide.shapes.add_textbox(Inches(4.5), Inches(5.5), Inches(4), Inches(1))
            out_tf = out_box.text_frame
            out_tf.text = "Out of Scope\n" + out_of_scope_kick
            for para in out_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT
            # PS Package
            ps_box = overview_slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(8), Inches(0.5))
            ps_tf = ps_box.text_frame
            ps_tf.text = "Professional Services Package + Terms & Conditions\n" + ps_package_kick
            ps_tf.paragraphs[0].font.name = 'Century Gothic'
            ps_tf.paragraphs[0].font.size = Pt(14)
            ps_tf.paragraphs[0].font.color.rgb = BLACK
            ps_tf.paragraphs[0].alignment = PP_ALIGN.LEFT
            # Add more slides for kick-off template (Success Criteria, Goals, Methodology, Timeline, Change Management)
            # Slide 5: Success Criteria
            success_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(success_slide)
            add_logo_footer(success_slide, theme_kick)
            txBox = success_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = f"{customer_name_kick} Project Success Criteria".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Success criteria text
            sc_box = success_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(5))
            sc_tf = sc_box.text_frame
            sc_tf.text = success_criteria_kick
            for para in sc_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT

            # Slide 6: Project Goals & Objectives
            goals_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(goals_slide)
            add_logo_footer(goals_slide, theme_kick)
            txBox = goals_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Project Goals & Objectives".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Goals text
            goals_box = goals_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(5))
            goals_tf = goals_box.text_frame
            goals_tf.text = goals_kick
            for para in goals_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT

            # Slide 7: Deployment Methodology
            meth_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(meth_slide)
            add_logo_footer(meth_slide, theme_kick)
            txBox = meth_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Zscaler Deployment Methodology".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Methodology text/table (simplified)
            meth_box = meth_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(5))
            meth_tf = meth_box.text_frame
            meth_tf.text = "Initiate\nPlan\nConfigure\nPilot\nProduction\nTransition"
            for para in meth_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT

            # Slide 8: Project Timeline
            timeline_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(timeline_slide)
            add_logo_footer(timeline_slide, theme_kick)
            txBox = timeline_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Project Timeline".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Timeline text
            timeline_box = timeline_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(5))
            timeline_tf = timeline_box.text_frame
            timeline_tf.text = project_timeline_kick
            for para in timeline_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT

            # Slide 9: Change Management & Key Milestones
            change_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(change_slide)
            add_logo_footer(change_slide, theme_kick)
            txBox = change_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Change Management impact on Project Key Milestones".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Change mgmt text
            change_box = change_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(5))
            change_tf = change_box.text_frame
            change_tf.text = change_mgmt_kick
            for para in change_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT

            # Slide 10: Next Steps & Follow-up Items
            next_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(next_slide)
            add_logo_footer(next_slide, theme_kick)
            txBox = next_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(0.5))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Next Steps & Follow-up Items".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(28)
            p.font.bold = True
            p.font.color.rgb = NAVY
            # Next steps text
            next_box = next_slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(8), Inches(5))
            next_tf = next_box.text_frame
            next_tf.text = next_steps_kick
            for para in next_tf.paragraphs:
                para.font.name = 'Century Gothic'
                para.font.size = Pt(14)
                para.font.color.rgb = BLACK
                para.alignment = PP_ALIGN.LEFT

            # Slide 11: Thank You
            thank_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_background(thank_slide)
            add_logo_footer(thank_slide, theme_kick)
            txBox = thank_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = "Thank you".title()
            p.font.name = 'Century Gothic'
            p.font.size = Pt(36)
            p.font.bold = True
            p.font.color.rgb = NAVY
            p.alignment = PP_ALIGN.LEFT

            bio = io.BytesIO()
            prs.save(bio)
            bio.seek(0)
            if output_format == "PPT":
                st.download_button("Download Kick-Off Deck", bio, file_name=f"{customer_name_kick}_Kickoff_Deck.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            elif output_format == "DOCX":
                doc = Document()
                doc.add_heading("Kick-Off Deck", 0)
                # Add content from slides to doc (simplified)
                doc.add_paragraph("Professional Services Project Kick-Off Meeting")
                doc.add_paragraph(customer_name_kick)
                doc.add_paragraph(pm_name_kick)
                doc.add_paragraph(psc_name_kick)
                doc.add_paragraph(date_kick)
                # Add agenda, team, etc. as paragraphs/tables
                doc.add_heading("Agenda", level=1)
                for item in ["Introductions", "Project Overview", "Success Criteria", "Project Methodology & Timeline", "Change Management & Key Milestones", "Next Steps & Follow-up Items"]:
                    doc.add_paragraph(item, style='List Bullet')
                # Team table
                doc.add_heading("Team Contacts", level=1)
                team_table = doc.add_table(rows=1, cols=3)
                hdr_cells = team_table.rows[0].cells
                hdr_cells[0].text = "Name"
                hdr_cells[1].text = "Role"
                hdr_cells[2].text = "Responsibility"
                for contact in team_contacts_kick:
                    row_cells = team_table.add_row().cells
                    row_cells[0].text = contact["name"]
                    row_cells[1].text = contact["role"]
                    row_cells[2].text = contact["responsibility"]
                # Add other sections similarly
                doc.save(bio)
                st.download_button("Download Kick-Off Deck", bio, file_name=f"{customer_name_kick}_Kickoff_Deck.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            elif output_format = "PDF":
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt="Kick-Off Deck", ln=1, align='C')
                # Add content as text
                pdf.multi_cell(0, 10, "Professional Services Project Kick-Off Meeting\n" + customer_name_kick + "\n" + pm_name_kick + "\n" + psc_name_kick + "\n" + date_kick)
                # Add agenda, team, etc.
                pdf.output(bio)
                st.download_button("Download Kick-Off Deck", bio, file_name=f"{customer_name_kick}_Kickoff_Deck.pdf", mime="application/pdf")

# Generate button for both tabs (shared)
output_format = st.selectbox("Select Output Format", ["PPT", "DOCX", "PDF"])

# Preview Summary (shared for simplicity)
if st.button("Preview Summary"):
    # Preview logic for selected tab
    if deck_type == "Transition Deck":
        # Transition preview
        st.write(f"Transition Deck for {customer_name_trans} on {today_date_trans}:")
        # ... (add preview details)
    else:
        # Kick-Off preview
        st.write(f"Kick-Off Deck for {customer_name_kick} on {date_kick}:")
        # ... (add preview details)

# Generate button
if st.button("Generate Deck"):
    # Validation and generation based on tab
    if deck_type == "Transition Deck":
        # Transition validation and generation (from previous code)
        # ... (insert transition logic)
        pass
    else:
        # Kick-Off validation and generation
        # ... (insert kick-off logic as above)
        pass
