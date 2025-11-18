#!/usr/bin/env python3
"""
Streamlit app: Fill your exact PPTX template (content.pptx) while preserving
the template master (backgrounds, headers/footers, logos).

How it works:
 - Prefer a local content.pptx file (place it next to this app) so masters/backgrounds/footers are preserved.
 - If the local file is missing, attempt to download the exact template from the repo raw URL.
 - Fallback to a user upload via Streamlit.
 - Replace {{TOKEN}} placeholders anywhere in text shapes or table cells.
 - Replace a shape whose text is exactly {{LOGO}} with an uploaded image, preserving the bounding box.
 - Heuristically fill tables for milestones/deliverables/open items if the template includes such tables.
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
# Configuration
# ------------------------
st.set_page_config(page_title="Zscaler Template Filler", layout="wide")

LOCAL_TEMPLATE_PATH = "content.pptx"
TEMPLATE_RAW_URL = (
    "https://raw.githubusercontent.com/jaijordan069/Zscaler-deck-generator/"
    "db10c36303a97f12497f797108dcc58b3cb4a327/content.pptx"
)

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
PLACEHOLDER_RE = re.compile(r"\{\{\s*([A-Z0-9_\-]+)\s*\}\}")

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


def find_placeholders_in_text(text: str) -> Set[str]:
    return {f"{{{{{m.group(1)}}}}}" for m in PLACEHOLDER_RE.finditer(text or "")}


def scan_presentation_for_placeholders(prs: Presentation) -> Set[str]:
    found: Set[str] = set()
    for slide in prs.slides:
        for shape in list(slide.shapes):
            try:
                if hasattr(shape, "text_frame") and shape.text_frame is not None:
                    found.update(find_placeholders_in_text(shape.text_frame.text or ""))
            except Exception:
                pass
            try:
                if getattr(shape, "has_table", False):
                    tbl = shape.table
                    for row in tbl.rows:
                        for cell in row.cells:
                            try:
                                found.update(find_placeholders_in_text(cell.text or ""))
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
            # attempt to preserve first-run font basics
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
        # per-run replace (preserve run formatting)
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
                before = cell.text or ""
            except Exception:
                continue
            after = before
            for k, v in mapping.items():
                if k in after:
                    after = after.replace(k, v)
            if after != before:
                try:
                    cell.text = after
                except Exception:
                    try:
                        cell.text_frame.clear()
                        cell.text_frame.paragraphs[0].text = after
                    except Exception:
                        pass


def replace_placeholders_in_presentation(prs: Presentation, mapping: Dict[str, str], logo_bytes: Optional[io.BytesIO] = None):
    shapes_to_remove = []
    for slide in prs.slides:
        for shape in list(slide.shapes):
            # tables
            try:
                if getattr(shape, "has_table", False):
                    replace_texts_in_table(shape.table, mapping)
                    continue
            except Exception:
                pass

            # logo placeholder (exact match)
            try:
                if hasattr(shape, "text_frame") and shape.text_frame is not None:
                    txt = (shape.text_frame.text or "").strip()
                    if txt == "{{LOGO}}" and logo_bytes:
                        left, top, width, height = shape.left, shape.top, shape.width, shape.height
                        shapes_to_remove.append((slide, shape))
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

            # generic text shapes
            try:
                if hasattr(shape, "text_frame") and shape.text_frame is not None:
                    replace_text_in_shape(shape, mapping)
            except Exception:
                pass

    for slide, shape in shapes_to_remove:
        try:
            el = shape._element
            el.getparent().remove(el)
        except Exception:
            pass

    return prs


def table_header_texts(table) -> List[str]:
    headers: List[str] = []
    try:
        if not table.rows:
            return headers
        for cell in table.rows[0].cells:
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
                    # milestones heuristics
                    if any("milestone" in h for h in headers) or ("baseline" in header_concat and "target" in header_concat):
                        if milestones_rows:
                            fill_table_rows_with_data(tbl, milestones_rows)
                            continue
                    # deliverables heuristics
                    if any("deliver" in h for h in headers) or ("date" in header_concat and "deliver" in header_concat):
                        if deliverables_rows:
                            fill_table_rows_with_data(tbl, deliverables_rows)
                            continue
                    # open items heuristics
                    if any("open" in h for h in headers) or any(("task" in h or "owner" in h or "transition" in h) for h in headers):
                        if open_rows:
                            fill_table_rows_with_data(tbl, open_rows)
                            continue
            except Exception:
                pass


# ------------------------
# Streamlit UI
# ------------------------
st.title("Zscaler Transition Deck — Template Filler")
st.markdown("This app edits your exact PPTX template (content.pptx) and preserves masters/backgrounds/footers.")

with st.sidebar:
    st.header("Instructions")
    st.markdown(
        "- Put placeholders in your template like {{CUSTOMER_NAME}} and {{PROJECT_SUMMARY}}.\n"
        "- Place a small textbox whose content is exactly {{LOGO}} where the logo should go; upload a logo here to replace it.\n"
        "- Place content.pptx alongside this app or upload it. If the repo is private, set GITHUB_TOKEN in the environment."
    )

# Basic fields
st.header("Customer & Project Basics")
c1, c2, c3 = st.columns(3)
customer_name = c1.text_input("Customer Name *", value="Pixartprinting")
today_date = c2.text_input("Today's Date (DD/MM/YYYY) *", value=datetime.utcnow().strftime("%d/%m/%Y"))
project_start = c3.text_input("Project Start Date (DD/MM/YYYY) *", value="01/06/2025")
project_end = st.text_input("Project End Date (DD/MM/YYYY) *", value="14/11/2025")
project_summary_text = st.text_area("Project Summary Text", value="More than half of the users have been deployed and there were no critical issues.")
theme = st.selectbox("Theme", ["White", "Navy"], index=0)

# Milestones (7)
st.header("Milestones (7 rows)")
milestone_defaults = [
    ("Initial Project Schedule Accepted", "27/06/2025", "27/06/2025", ""),
    ("Initial Design Accepted", "14/07/2025", "17/07/2025", ""),
    ("Pilot Configuration Complete", "28/07/2025", "18/07/2025", ""),
    ("Pilot Rollout Complete", "08/08/2025", "22/08/2025", ""),
    ("Production Configuration Complete", "29/08/2025", "29/08/2025", ""),
    ("Production Rollout Complete", "14/11/2025", "??", ""),
    ("Final Design Accepted", "14/11/2025", "14/11/2025", ""),
]
milestones_data = []
for i in range(7):
    with st.expander(f"Milestone {i+1}", expanded=False):
        mn = st.text_input(f"Milestone {i+1} Name", value=milestone_defaults[i][0], key=f"mname_{i}")
        mb = st.text_input(f"Baseline {i+1} (DD/MM/YYYY)", value=milestone_defaults[i][1], key=f"mbaseline_{i}")
        mt = st.text_input(f"Target {i+1} (DD/MM/YYYY)", value=milestone_defaults[i][2], key=f"mtarget_{i}")
        ms = st.text_input(f"Status {i+1}", value=milestone_defaults[i][3], key=f"mstatus_{i}")
        milestones_data.append({"name": mn, "baseline": mb, "target": mt, "status": ms})

# Deliverables (5)
st.header("Deliverables (5 rows)")
deliverables_data = []
for i in range(5):
    with st.expander(f"Deliverable {i+1}", expanded=False):
        dn = st.text_input(f"Deliverable Name {i+1}", value="", key=f"dname_{i}")
        dd = st.text_input(f"Date Delivered {i+1}", value="", key=f"ddate_{i}")
        deliverables_data.append({"name": dn, "date": dd})

# Open items (6)
st.header("Open Items (6 rows)")
open_items_data = []
for i in range(6):
    with st.expander(f"Open Item {i+1}", expanded=False):
        otask = st.text_input(f"Task/Description {i+1}", value="", key=f"otask_{i}")
        odate = st.text_input(f"Date {i+1}", value="", key=f"odate_{i}")
        oowner = st.text_input(f"Owner {i+1}", value="", key=f"oowner_{i}")
        osteps = st.text_area(f"Transition Plan/Next Steps {i+1}", value="", key=f"osteps_{i}", height=80)
        open_items_data.append({"task": otask, "date": odate, "owner": oowner, "steps": osteps})

# Contacts
st.header("Contacts")
c1, c2 = st.columns(2)
pm_name = c1.text_input("Project Manager Name", value="Alex Vazquez")
consultant_name = c2.text_input("Consultant Name", value="Alex Vazquez")
primary_contact = st.text_input("Primary Contact", value="Teia proctor")
secondary_contact = st.text_input("Secondary Contact", value="Marco Sattier")

# Uploaders
st.markdown("Optional: upload logo to replace {{LOGO}} in the template.")
logo_upload = st.file_uploader("Upload logo (png/jpg)", type=["png", "jpg", "jpeg"])
st.markdown("Optional: upload template PPTX if content.pptx is not present.")
uploaded_template = st.file_uploader("Upload template PPTX", type=["pptx"])

# Preview placeholders
if st.button("Preview placeholders in template"):
    tpl_bytes = None
    if os.path.exists(LOCAL_TEMPLATE_PATH):
        try:
            with open(LOCAL_TEMPLATE_PATH, "rb") as f:
                tpl_bytes = io.BytesIO(f.read())
            st.success("Loaded local template.")
        except Exception as e:
            st.warning(f"Failed to read local template: {e}")
    if tpl_bytes is None:
        tpl_bytes = download_bytes_from_github(TEMPLATE_RAW_URL)
        if tpl_bytes:
            st.success("Downloaded template from GitHub raw URL.")
    if tpl_bytes is None and uploaded_template:
        tpl_bytes = io.BytesIO(uploaded_template.read())
        st.success("Loaded uploaded template.")
    if tpl_bytes is None:
        st.warning("Template not found. Place content.pptx in app folder or upload it.")
    else:
        try:
            prs = Presentation(tpl_bytes)
            placeholders = sorted(scan_presentation_for_placeholders(prs))
            st.write("Placeholders found:")
            st.json(placeholders)
        except Exception as e:
            st.error(f"Failed to open PPTX: {e}")

# Generate deck
if st.button("Generate Transition Deck"):
    # basic validation
    if not customer_name:
        st.error("Customer Name is required.")
        st.stop()
    for d in (today_date, project_start, project_end):
        if d and not is_valid_date(d):
            st.error(f"Date '{d}' must be DD/MM/YYYY format.")
            st.stop()

    # load template (local -> github raw -> upload)
    tpl_bytes = None
    if os.path.exists(LOCAL_TEMPLATE_PATH):
        try:
            with open(LOCAL_TEMPLATE_PATH, "rb") as f:
                tpl_bytes = io.BytesIO(f.read())
            st.info("Loaded local template.")
        except Exception:
            tpl_bytes = None

    if tpl_bytes is None:
        tpl_bytes = download_bytes_from_github(TEMPLATE_RAW_URL)
        if tpl_bytes:
            st.info("Downloaded template from GitHub raw URL.")

    if tpl_bytes is None:
        if uploaded_template is None:
            st.error("Template not available. Upload or place content.pptx in app folder.")
            st.stop()
        tpl_bytes = io.BytesIO(uploaded_template.read())
        st.info("Loaded uploaded template.")

    try:
        prs = Presentation(tpl_bytes)
    except Exception as e:
        st.error(f"Failed to load template: {e}")
        st.stop()

    # mapping placeholders
    mapping: Dict[str, str] = {
        "{{CUSTOMER_NAME}}": customer_name,
        "{{TODAY_DATE}}": today_date,
        "{{PROJECT_START}}": project_start,
        "{{PROJECT_END}}": project_end,
        "{{PROJECT_SUMMARY}}": project_summary_text,
        "{{PM_NAME}}": pm_name,
        "{{CONSULTANT_NAME}}": consultant_name,
        "{{PRIMARY_CONTACT}}": primary_contact,
        "{{SECONDARY_CONTACT}}": secondary_contact,
    }

    # present extra placeholders for manual mapping
    placeholders_found = scan_presentation_for_placeholders(prs)
    extra = sorted([p for p in placeholders_found if p not in mapping])
    if extra:
        st.markdown("Provide values for additional placeholders detected in the template:")
        for token in extra:
            val = st.text_input(f"Value for {token}", key=f"ph_{token}")
            if val:
                mapping[token] = val

    # prepare table data rows
    milestones_rows = [[m["name"], m["baseline"], m["target"], m["status"]] for m in milestones_data]
    deliverables_rows = [[d["name"], d["date"]] for d in deliverables_data]
    open_rows = [[o["task"], o["date"], o["owner"], o["steps"]] for o in open_items_data]

    # attempt heuristic table fill
    try:
        heuristically_fill_known_tables(prs, milestones_rows, deliverables_rows, open_rows)
    except Exception:
        pass

    # prepare logo bytes
    logo_bytes = None
    if logo_upload:
        try:
            logo_bytes = io.BytesIO(logo_upload.read())
        except Exception:
            logo_bytes = None

    # replace placeholders
    try:
        prs_filled = replace_placeholders_in_presentation(prs, mapping, logo_bytes=logo_bytes)
    except Exception as e:
        st.error(f"Failed to replace placeholders: {e}")
        st.stop()

    # save and provide download
    out = io.BytesIO()
    try:
        prs_filled.save(out)
        out.seek(0)
    except Exception as e:
        st.error(f"Failed to save PPTX: {e}")
        st.stop()

    st.success("Generated PPTX from template.")
    st.download_button("Download Transition Deck", data=out, file_name=f"{customer_name}_Transition_Filled.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
