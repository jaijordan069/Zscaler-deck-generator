import streamlit as st
import openai
import requests
import json
from jsonschema import validate, ValidationError
from tenacity import retry, wait_exponential, stop_after_attempt
from datetime import datetime
import os
import time  # For any delays if needed

# Page config
st.set_page_config(page_title="Zscaler Transition Deck Figma Generator", layout="wide")

# API Keys (use secrets in production)
openai.api_key = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY"))
FIGMA_TOKEN = st.secrets.get("FIGMA_TOKEN", os.getenv("FIGMA_TOKEN"))
FIGMA_BASE = "https://api.figma.com/v1"

st.title("Zscaler Professional Services Transition Deck Figma Generator")
st.markdown("Fill in details to generate a Figma file based on the template. Outputs editable frames for each slide.")

# Sidebar instructions
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    - Enter customer data as before.
    - Generation uses AI for layout JSON—review for tweaks.
    - Download: Figma file URL provided.
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

# Figma Schema Validator
figma_schema = {
    "type": "object",
    "properties": {
        "id": {"type": "string"},
        "type": {"enum": ["DOCUMENT", "CANVAS", "FRAME"]},
        "children": {"type": "array"}
    },
    "required": ["id", "type", "children"]
}

# Prompt Generator for Slide JSON with JSON mode
def generate_slide_json(slide_desc, content):
    client = openai.OpenAI()
    prompt = f"""
    Generate valid Figma JSON for a slide frame based on this Zscaler template description: {slide_desc}.
    Inject content: {content}.
    Use blue theme (#0066CC fills, white text), tables for data, bullets for lists.
    Output ONLY JSON: {{"id": "slide_id", "type": "FRAME", "name": "Slide Title", "children": [...]}}
    Dimensions: 1920x1080. Include text nodes, rectangles, auto-layout where possible.
    """
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "Output valid Figma JSON schema only. Use RGB colors (0-1 scale)."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.2,
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)

# Create Figma File as JSON (no API upload)
def create_figma_file(children, customer_name):
    json_data = {
        "document": {
            "id": "0:0",
            "type": "DOCUMENT",
            "name": f"{customer_name} Transition Deck",
            "children": children
        }
    }
    return json.dumps(json_data, indent=2)

# Generate Button
if st.button("Generate Figma Deck") and openai.api_key:
    with st.spinner("Generating slides..."):
        # Prepare dynamic content for prompts
        milestones_str = "; ".join([f"{m['name']}: {m['baseline']}, {m['target']}, {m['status']}" for m in milestones_data])
        objectives_str = "; ".join([f"{o['objective']}: {o['actual']}, {o['deviation']}" for o in objectives_data])
        deliverables_str = "; ".join([f"{d['name']}: {d['date']}" for d in deliverables_data])
        open_items_str = "; ".join([f"{oi['task']}: {oi['date']}, {oi['owner']}, {oi['steps']}" for oi in open_items_data])
        tech_str = f"Auth: {idp}, {auth_type}; Prov: {prov_type}; Tunnel: {tunnel_type}; Deploy: {deploy_system}; Devices: {windows_num} Win, {mac_num} Mac; Geo: {geo_locations}; Policies: SSL {ssl_policies}, URL {url_policies}, Cloud {cloud_policies}, FW {fw_policies}"
        short_str = "; ".join(short_term)
        long_str = "; ".join(long_term)
        contacts_str = f"PM: {pm_name}; Consultant: {consultant_name}; Primary: {primary_contact}; Secondary: {secondary_contact}"

        # Define slides with descriptions and content injection
        slides = [
            {"desc": "Title slide: Professional Services Transition Meeting with customer name and date. Office background.", "content": f"{customer_name}, {today_date}"},
            {"desc": "Agenda slide: Bullets for Project Summary, Technical Summary, Recommended Next Steps.", "content": "Standard agenda items"},
            {"desc": "Project Summary title slide.", "content": ""},
            {"desc": "Status report: Tables for milestones, rollout roadmap, objectives. Include project summary text.", "content": f"Summary: {project_summary_text}; Milestones: {milestones_str}; Rollout: Pilot {pilot_target}/{pilot_current} {pilot_status}, Prod {prod_target}/{prod_current} {prod_status}; Objectives: {objectives_str}"},
            {"desc": "Deliverables table.", "content": f"Deliverables: {deliverables_str}"},
            {"desc": "Technical Summary title.", "content": ""},
            {"desc": "ZIA Architecture diagram: Boxes for auth, tunnels, policies. Use rectangles and lines.", "content": tech_str},
            {"desc": "Open Items table.", "content": f"Open Items: {open_items_str}"},
            {"desc": "Next Steps: Two columns for short/long term bullets.", "content": f"Short: {short_str}; Long: {long_str}"},
            {"desc": "Thank You slide with contacts and survey note.", "content": contacts_str},
            {"desc": "Final Thank You.", "content": ""}
        ]
        
        children = []
        for i, slide in enumerate(slides):
            try:
                slide_json = generate_slide_json(slide["desc"], slide["content"])
                validate(instance=slide_json, schema=figma_schema)
                children.append(slide_json)
            except (ValidationError, json.JSONDecodeError) as e:
                st.warning(f"Slide {i+1} generation failed: {str(e)[:100]}... Skipping.")
                # Fallback simple frame
                fallback = {"id": f"slide_{i}", "type": "FRAME", "name": f"Slide {i+1}", "children": []}
                children.append(fallback)
        
        # Generate JSON output and download
        json_output = create_figma_file(children, customer_name)
        st.download_button(
            label="Download Figma JSON (Import Manually)",
            data=json_output,
            file_name=f"{customer_name}_Transition_Deck.json",
            mime="application/json"
        )
        st.info("Download the JSON above and import it into Figma using a plugin like 'JSON to Figma' for your editable deck. Update dates like 14/11/2025 as needed.")
        st.success("Deck generated successfully! Import the JSON to Figma to view/edit the 11 slides.")
        st.balloons()
else:
    st.warning("Add your OPENAI_API_KEY in app settings > Secrets to enable generation.")
    st.info("Click 'Generate Figma Deck' once the key is set.")
