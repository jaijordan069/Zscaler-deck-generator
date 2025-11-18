#!/usr/bin/env python3
"""
Streamlit app: Fill your exact PPTX template (content.pptx) while preserving
the template master (backgrounds, headers/footers, logos).

What this version does:
- Prefer a local template file ./content.pptx (place your committed template in the app folder)
- If not present, attempt to download the exact file from the repo commit raw URL (private repo supported via GITHUB_TOKEN)
- Fallback to Streamlit file_uploader if automatic retrieval fails
- Scans the template for all placeholder tokens of the form {{TOKEN}} and shows them to you
- Automatically replaces common placeholders with the UI values you enter
- Heuristically fills tables that appear to be Milestones, Deliverables or Open Items by matching header text
- Replaces a shape whose text is exactly {{LOGO}} with an uploaded logo while preserving the shape bounding box
- Does not recreate slides or add branding programmatically — the template's masters, headers/footers and background remain intact

How to use:
1. Put the exact template PPTX in the same folder as this app and name it content.pptx (recommended).
2. Alternatively, make sure the TEMPLATE_RAW_URL below points at the raw GitHub URL for your template commit (or set GITHUB_TOKEN env var for private repos).
3. Add placeholders in your template (e.g. {{CUSTOMER_NAME}}, {{TODAY_DATE}}, {{M1_NAME}}, {{DELIVERABLE_1}} or put a shape with text {{LOGO}} where the logo should go).
4. Run: streamlit run app.py ; fill the UI and press Generate Transition Deck.

Notes:
- python-pptx does not embed fonts; for exact font rendering ensure the target environment has the same fonts installed.
- Table auto-fill overwrites existing table rows (it will not add new rows if the table in the template has fewer rows than the data).
"""

from __future__ import annotations

import io
import os
import re
import requests
from datetime import datetime
from typing import Dict, List, Optional, Set

import streamlit as st
from pptx import Presentation
from pptx.util import Pt

# ------------------------
# Config - adjust as needed
# ------------------------
st.set_page_config(page_title="Zscaler Template Filler (Exact PPTX)", layout="wide")

# Local filename to prefer (recommended: commit content.pptx to repo and this app runs from repo root)
LOCAL_TEMPLATE_PATH = "content.pptx"

# Raw URL to the exact template commit (fallback). Update if you have a different commit/path.
TEMPLATE_RAW_URL = (
    "https://raw.githubusercontent.com/jaijordan069/Zscaler-deck-generator/"
    "db10c36303a97f12497f797108dcc58b3cb4a327/content.pptx"
)

# Date pattern
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")

# ------------------------
# Utility helpers
# ------------------------
def is_valid_date(d: str) -> bool:
    return bool(d and DATE_RE.match(d))


def download_bytes_from_github(raw_url: str, timeout: int = 20) -> Optional[io.BytesIO]:
    headers = {"Accept": "application/octet-stream"}
    token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
    if token:
        headers["Authorization"] = f"token {token}"
    try:
        r = requests.get(raw_url, headers=headers, timeout=timeout)
        r.raise_for_status()
        return io.BytesIO(r.content)
    except Exception:
        return None


# Placeholder extraction helpers
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Z0-9_\-]+)\s*\}\}")

def find_placeholders_in_text(text: str) -> Set[str]:
    return {f"{{{{{m.group(1)}}}}}" for m in PLACEHOLDER_RE.finditer(text or "")}

def scan_presentation_for_placeholders(prs: Presentation) -> Set[str]:
    found: Set[str] = set()
    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                if hasattr(shape, "text_frame") and shape.text_frame is not None:
                    txt = shape.text_frame.text or ""
                    found.update(find_placeholders_in_text(txt))
            except Exception:
                pass
            try:
                if getattr(shape, "has_table", False):
                    tbl = shape.table
                    for row in tbl.rows:
                        for cell in row.cells:
                            try:
                                ct = cell.text or ""
                                found.update(find_placeholders_in_text(ct))
                            except Exception:
                                pass
            except Exception:
                pass
    return found


def replace_text_in_shape(shape, mapping: Dict[str, str]):
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return
    tf = shape.text_frame
    full_text = tf.text or ""
    new_text = full_text
    for k, v in mapping.items():
        if k in new_text:
            new_text = new_text.replace(k, v)
    if new_text != full_text:
        try:
            first_para = tf.paragraphs[0]
            font_props = {}
            if first_para.runs:
                r0 = first_para.runs[0]
                try:
                    font_props["name"] = r0.font.name
                    font_props["size"] = r0.font.size
                    font_props["bold"] = r0.font.bold
                    font_props["italic"] = r0.font.italic
                    font_props["color"] = r0.font.color.rgb if r0.font.color else None
                except Exception:
                    font_props = {}
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = new_text
            if font_props:
                try:
                    if font_props.get("name"):
                        run.font.name = font_props["name"]
                    if font_props.get("size"):
                        run.font.size = font_props["size"]
                    if "bold" in font_props:
                        run.font.bold = font_props["bold"]
                    if "italic" in font_props:
                        run.font.italic = font_props["italic"]
                    if font_props.get("color"):
                        run.font.color.rgb = font_props["color"]
                except Exception:
                    pass
        except Exception:
            try:
                tf.clear()
                tf.paragraphs[0].text = new_text
            except Exception:
                pass
    else:
        for para in tf.paragraphs:
            for run in list(para.runs):
                rt = run.text
                new_rt = rt
                for k, v in mapping.items():
                    if k in new_rt:
                        new_rt = new_rt.replace(k, v)
                if new_rt != rt:
                    try:
                        run.text = new_rt
                    except Exception:
                        pass


def replace_texts_in_table(table, mapping: Dict[str, str]):
    for row in table.rows:
        for cell in row.cells:
            try:
                text_before = cell.text or ""
            except Exception:
                continue
            new_text = text_before
            for k, v in mapping.items():
                if k in new_text:
                    new_text = new_text.replace(k, v)
            if new_text != text_before:
                try:
                    cell.text = new_text
                except Exception:
                    try:
                        cell.text_frame.clear()
                        cell.text_frame.paragraphs[0].text = new_text
                    except Exception:
                        pass


def replace_placeholders_in_presentation(prs: Presentation, mapping: Dict[str, str], logo_bytes: Optional[io.BytesIO] = None):
    shapes_removed = []
    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                if getattr(shape, "has_table", False):
                    replace_texts_in_table(shape.table, mapping)
                    continue
            except Exception:
                pass
            try:
                if hasattr(shape, "text_frame") and shape.text_frame is not None:
                    txt = shape.text_frame.text.strip()
                    if txt == "{{LOGO}}" and logo_bytes:
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        shapes_removed.append((slide, shape))
                        try:
                            logo_bytes.seek(0)
                            slide.shapes.add_picture(logo_bytes, left, top, width=width, height=height)
                        except Exception:
                            try:
                                logo_bytes.seek(0)
                                slide.shapes.add_picture(logo_bytes, left, top, width=width)
                            except Exception:
                                try:
                                    logo_bytes.seek(0)
                                    slide.shapes.add_picture(logo_bytes, left, top)
                                except Exception:
                                    pass
                        continue
            except Exception:
                pass
            try:
                if hasattr(shape, "text_frame") and shape.text_frame is not None:
                    replace_text_in_shape(shape, mapping)
            except Exception:
                pass
    for slide, shape in shapes_removed:
        try:
            el = shape._element
            el.getparent().remove(el)
        except Exception:
            pass
    return prs


def table_header_texts(table) -> List[str]:
    headers = []
    try:
        if not table.rows:
            return headers
        top_row = table.rows[0]
        for cell in top_row.cells:
            try:
                headers.append((cell.text or "").strip().lower())
            except Exception:
                headers.append("")
    except Exception:
        pass
    return headers


def fill_table_rows_with_data(table, data_rows: List[List[str]]):
    max_rows = len(table.rows) - 1
    rows_to_write = min(max_rows, len(data_rows))
    for r_idx in range(rows_to_write):
        data = data_rows[r_idx]
        for c_idx in range(len(table.columns)):
            try:
                cell = table.cell(r_idx + 1, c_idx)
                value = data[c_idx] if c_idx < len(data) else ""
                cell.text = str(value or "")
            except Exception:
                pass


def heuristically_fill_known_tables(prs: Presentation, milestones_rows: List[List[str]], deliverables_rows: List[List[str]], open_rows: List[List[str]]):
    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                if getattr(shape, "has_table", False):
                    tbl = shape.table
                    headers = table_header_texts(tbl)
                    header_concat = " ".join(headers)
                    if any("milestone" in h for h in headers) or ("baseline" in header_concat and "target" in header_concat):
                        if milestones_rows:
                            fill_table_rows_with_data(tbl, milestones_rows)
                            continue
                    if any("deliver" in h for h in headers) or ("date" in header_concat and "deliver" in header_concat):
                        if deliverables_rows:
                            fill_table_rows_with_data(tbl, deliverables_rows)
                            continue
                    if any("open" in h for h in headers) or any("task" in h or "owner" in h or "transition" in h for h in headers):
                        if open_rows:
                            fill_table_rows_with_data(tbl, open_rows)
                            continue
            except Exception:
                pass


st.title("Zscaler Transition Deck — Exact Template Filler")
st.markdown("This app edits your exact template (content.pptx). It will preserve backgrounds, headers, footers and masters.")

with st.sidebar:
    st.header("Template loading options")
    st.markdown("- Place content.pptx in the app folder (preferred) or upload it below.\n- If the template lives in a private GitHub repo, set GITHUB_TOKEN in the environment where the app runs.\n- Put placeholders in the template (like {{CUSTOMER_NAME}}). Use a textbox with text {{LOGO}} to mark where a logo should go.") 

st.header("Customer & Project Basics")
c1, c2, c3 = st.columns(3)
customer_name = c1.text_input("Customer Name *", value="Pixartprinting")
today_date = c2.text_input("Today's Date (DD/MM/YYYY) *", value=datetime.utcnow().strftime("%d/%m/%Y"))
project_start = c3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025")
project_end = st.text_input("Project End Date (DD/MM/YYYY) *", value="14/11/2025")
project_summary_text = st.text_area("Project Summary Text", value="More than half of the users have been deployed and there were no critical issues. Remaining enrollments expected without major issues.")
theme = st.selectbox("Theme", ["White", "Navy"], index=0)

st.header("Milestones (7 rows - as template)")
milestone_defaults = [
    ("Initial Project Schedule Accepted", "27/06/2025", "27/06/2025", ""),
    ("Initial Design Accepted", "14/07/2025", "17/07/2025", ""),
    ("Pilot Configuration Complete", "28/07/2025", "18/07/2025", ""),
    ("Pilot Rollout Complete", "08/08/2025", "22/08/2025", ""),
    ("Production Configuration Complete", "29/08/2025", "29/08/2025", ""),
    ("Production Rollout Complete", "14/11/2025", "??", ""),
    ("Final Design
