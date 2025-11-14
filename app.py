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
theme = st.selectbox("Theme", ["White", "Navy"])

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
    {"task": "Tighten Firewall policies",
