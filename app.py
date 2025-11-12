import streamlit as st
import openai
import requests
import json
from jsonschema import validate, ValidationError
from tenacity import retry, wait_exponential, stop_after_attempt
from datetime import datetime
import os

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

# Reuse inputs from original code (abbreviated for brevity; expand as needed)
customer_name = st.text_input("Customer Name", value="Pixartprinting")
today_date = st.text_input("Today's Date (DD/MM/YYYY)", value="19/09/2025")
# ... (Add other inputs: milestones, objectives, etc., as in original)

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

# Prompt Generator for Slide JSON
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
        temperature=0.2
    )
    return json.loads(response.choices[0].message.content)

# Create Figma File
@retry(wait=wait_exponential(min=1, max=10), stop=stop_after_attempt(3))
def create_figma_file(children):
    headers = {"X-Figma-Token": FIGMA_TOKEN}
    payload = {
        "document": {
            "id": "0:0",
            "type": "DOCUMENT",
            "name": f"{customer_name} Transition Deck",
            "children": children
        }
    }
    response = requests.post(f"{FIGMA_BASE}/files", json=payload, headers=headers)
    response.raise_for_status()
    return response.json()["key"]

# Generate Button
if st.button("Generate Figma Deck") and FIGMA_TOKEN and openai.api_key:
    with st.spinner("Generating slides..."):
        # Define slides with descriptions and content injection
        slides = [
            {"desc": "Title slide: Professional Services Transition Meeting with customer name and date. Office background.", "content": f"{customer_name}, {today_date}"},
            {"desc": "Agenda slide: Bullets for Project Summary, Technical Summary, Recommended Next Steps.", "content": "Standard agenda items"},
            {"desc": "Project Summary title slide.", "content": ""},
            {"desc": "Status report: Tables for milestones, rollout roadmap, objectives.", "content": f"Milestones: [insert data]; Objectives: [insert data]"},  # Inject dynamic data here
            {"desc": "Deliverables table.", "content": "List of deliverables with dates"},
            {"desc": "Technical Summary title.", "content": ""},
            {"desc": "ZIA Architecture diagram: Boxes for auth, tunnels, policies. Use rectangles and lines.", "content": "Auth: Entra ID, SAML; Policies: 10 SSL, etc."},
            {"desc": "Open Items table.", "content": "Tasks with dates, owners, steps"},
            {"desc": "Next Steps: Two columns for short/long term bullets.", "content": "Short: Finish rollout, etc.; Long: Deploy mobile, etc."},
            {"desc": "Thank You slide with contacts and survey note.", "content": f"PM: Alex Vazquez; Contacts: Teia Proctor, Marco Sattler"},
            {"desc": "Final Thank You.", "content": ""}
        ]
        
        children = []
        for i, slide in enumerate(slides):
            try:
                slide_json = generate_slide_json(slide["desc"], slide["content"])
                validate(instance=slide_json, schema=figma_schema)
                children.append(slide_json)
            except ValidationError as e:
                st.warning(f"Slide {i+1} validation failed: {e.message}")
        
        file_key = create_figma_file(children)
        figma_url = f"https://www.figma.com/file/{file_key}"
        st.success(f"Figma file created! Edit here: {figma_url}")
        st.balloons()
else:
    st.info("Add API keys to secrets.toml or env vars to proceed.")
