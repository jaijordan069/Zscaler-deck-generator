import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL, MSO_PATTERN, MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR_TYPE
import io, re, requests

st.set_page_config(page_title="Zscaler Transition Deck Generator", layout="wide")

# ----- UI styling -------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {background: linear-gradient(to bottom, #0066CC, white);}
    </style>
    """,
    unsafe_allow_html=True,
)
st.image("https://companieslogo.com/img/orig/ZS-46a5871c.png?t=1720244494", width=200)

st.title("Zscaler Professional Services Transition Deck Generator")
st.markdown("Fill the form → Preview → Generate → Download PPTX")

# ----- Helper: date validation ------------------------------------
def is_valid_date(s):
    return bool(re.match(r'^\d{2}/\d{2}/\d{4}$', s))

# ----- FORM -------------------------------------------------------
st.header("Customer & Project Basics")
c1, c2, c3 = st.columns(3)
customer_name = c1.text_input("Customer Name *", "Pixartprinting")
today_date    = c2.text_input("Today's Date (DD/MM/YYYY) *", "14/11/2025")
project_start = c3.text_input("Project Start Date (DD/MM/YYYY) *", "01/06/2025")
project_end   = st.text_input("Project End Date (DD/MM/YYYY) *", "14/11/2025")
project_summary_text = st.text_area(
    "Project Summary Text",
    "More than half of the users have been deployed and there were not any critical issues. Not expected issues during enrollment of remaining users",
)

# ----- Milestones -------------------------------------------------
st.header("Milestones")
milestones_data = []
for i in range(7):
    with st.expander(f"Milestone {i+1}", expanded=i==0):
        name = st.text_input(f"Name {i+1}", key=f"mname_{i}")
        baseline = st.text_input(f"Baseline {i+1}", key=f"mbaseline_{i}")
        target   = st.text_input(f"Target {i+1}", key=f"mtarget_{i}")
        status   = st.text_input(f"Status {i+1}", key=f"mstatus_{i}")
        if name:
            milestones_data.append({"name": name, "baseline": baseline,
                                   "target": target, "status": status})

# ----- Rollout ----------------------------------------------------
st.header("User Rollout Roadmap")
p1, p2 = st.columns(2)
with p1:
    st.subheader("Pilot")
    pilot_target = st.number_input("Target Users", 100)
    pilot_current = st.number_input("Current Users", 449)
    pilot_completion = st.text_input("Completion Date", "14/11/2025")
    pilot_status = st.text_input("Status", "Completed")
with p2:
    st.subheader("Production")
    prod_target = st.number_input("Target Users ", 800)
    prod_current = st.number_input("Current Users ", 449)
    prod_completion = st.text_input("Completion Date ", "14/11/2025")
    prod_status = st.text_input("Status ", "In Progress")

# ----- Objectives -------------------------------------------------
st.header("Project Objectives")
objectives_data = []
for i in range(3):
    with st.expander(f"Objective {i+1}", expanded=i==0):
        obj = st.text_area(f"Planned {i+1}", key=f"obj_{i}", height=50)
        act = st.text_area(f"Actual {i+1}", key=f"act_{i}", height=50)
        dev = st.text_area(f"Deviation {i+1}", key=f"dev_{i}", height=50)
        if obj:
            objectives_data.append({"objective": obj, "actual": act, "deviation": dev})

# ----- Deliverables ------------------------------------------------
st.header("Deliverables")
deliverables_data = []
for i in range(5):
    with st.expander(f"Deliverable {i+1}", expanded=i==0):
        name = st.text_input(f"Name {i+1}", key=f"dname_{i}")
        date = st.text_input(f"Date {i+1}", key=f"ddate_{i}")
        if name:
            deliverables_data.append({"name": name, "date": date})

# ----- Technical --------------------------------------------------
st.header("Technical Summary")
t1, t2 = st.columns(2)
with t1:
    idp       = st.text_input("Identity Provider", "Entra ID")
    auth_type = st.text_input("Authentication Type", "SAML 2.0")
    prov_type = st.text_input("Provisioning", "SCIM Provisioning")
with t2:
    tunnel_type = st.text_input("Tunnel Type", "ZCC with Z-Tunnel 2.0")
    deploy_system = st.text_input("Deployment System", "MS Intune/Jamf")
d1, d2, d3 = st.columns(3)
windows_num = d1.number_input("Windows Devices", 351)
mac_num     = d2.number_input("MacOS Devices", 98)
geo_locations = d3.text_input("Geo Locations", "Europe, North Africa, USA")
p1, p2, p3, p4 = st.columns(4)
ssl_policies = p1.number_input("SSL Policies", 10)
url_policies = p2.number_input("URL Policies", 5)
cloud_policies = p3.number_input("Cloud App Control", 5)
fw_policies   = p4.number_input("Firewall Policies", 15)

# ----- Open Items -------------------------------------------------
st.header("Open Items")
open_items_data = []
for i in range(6):
    with st.expander(f"Open Item {i+1}", expanded=i==0):
        task = st.text_input(f"Task {i+1}", key=f"otask_{i}")
        odate = st.text_input(f"Date {i+1}", key=f"odate_{i}")
        owner = st.text_input(f"Owner {i+1}", key=f"oowner_{i}")
        steps = st.text_area(f"Steps {i+1}", key=f"osteps_{i}", height=50)
        if task:
            open_items_data.append({"task": task, "date": odate,
                                   "owner": owner, "steps": steps})

# ----- Next Steps -------------------------------------------------
st.header("Recommended Next Steps")
st.subheader("Short Term")
short_term_input = st.text_area(
    "Short Term (comma-separated)",
    "Finish Production rollout, Tighten Firewall policies, Tighten Cloud App Control Policies, Fine tune SSL Inspection policies, Configure Role Based Access Control (RBAC), Configure DLP policies",
)
short_term = [s.strip() for s in short_term_input.split(",") if s.strip()]

st.subheader("Long Term")
long_term_input = st.text_area(
    "Long Term (comma-separated)",
    "Deploy ZCC on Mobile devices, Consider an upgrade of Sandbox license to have better antimalware protection, Consider an upgrade of the Firewall License to be able to apply policies based on user groups and network applications, Adopt additional Zscaler solutions like Zscaler Private Access (ZPA) or Zscaler Digital experience (ZDX), Consider using ZCC Client when the users are on-prem for a more consistent user experience, Integrate ZIA with 3rd party SIEM",
)
long_term = [s.strip() for s in long_term_input.split(",") if s.strip()]

# ----- Contacts ---------------------------------------------------
st.header("Contacts")
c1, c2 = st.columns(2)
pm_name = c1.text_input("Project Manager", "Alex Vazquez")
consultant_name = c2.text_input("Consultant", "Alex Vazquez")
primary_contact = st.text_input("Primary Contact", "Teia proctor")
secondary_contact = st.text_input("Secondary Contact", "Marco Sattier")

# ----- Preview ----------------------------------------------------
if st.button("Preview Summary"):
    st.write(f"**Deck for {customer_name} – {today_date}**")
    st.write(f"- Summary: {project_summary_text[:100]}…")
    st.write(f"- Milestones: {len(milestones_data)}")
    st.write(f"- Pilot: {pilot_current}/{pilot_target} ({pilot_status})")
    st.write(f"- Production: {prod_current}/{prod_target} ({prod_status})")
    st.write(f"- Objectives: {len(objectives_data)}")
    st.write(f"- Deliverables: {len(deliverables_data)}")
    st.write(f"- Open Items: {len(open_items_data)}")
    st.write(f"- Short-term: {len(short_term)}")
    st.write(f"- Long-term: {len(long_term)}")
    st.write(f"- PM: {pm_name} | Consultant: {consultant_name}")

# ----- GENERATE PPTX -----------------------------------------------
if st.button("Generate Transition Deck"):
    # ---- validation ------------------------------------------------
    if not customer_name:
        st.error("Customer Name required")
    elif not all(is_valid_date(d) for d in [today_date, project_start,
                                          project_end, pilot_completion,
                                          prod_completion]):
        st.error("All dates must be DD/MM/YYYY")
    else:
        prs = Presentation()

        # ---- master background (gradient + dots) -------------------
        master = prs.slide_masters[0]
        fill = master.background.fill
        fill.gradient()
        fill.gradient_stops[0].color.rgb = RGBColor(0, 102, 204)   # dark blue
        fill.gradient_stops[1].color.rgb = RGBColor(255, 255, 255) # white
        fill.gradient_angle = 90
        fill.patterned()
        fill.pattern = MSO_PATTERN.DOTTED_GRID
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
        fill.back_color.rgb = RGBColor(255, 255, 255)

        # ---- helper: header / footer / slide number ---------------
        def add_header_footer_number(slide, num):
            # PROSERVE header
            h = slide.shapes.add_textbox(Inches(8), Inches(0), Inches(2), Inches(0.5))
            h.text_frame.text = "PROSERVE"
            h.text_frame.paragraphs[0].font.size = Pt(32)
            h.text_frame.paragraphs[0].font.bold = True
            h.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
            h.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

            # Footer
            f = slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(9), Inches(0.5))
            f.text_frame.text = "Zscaler, Inc. All rights reserved. © 2025"
            f.text_frame.paragraphs[0].font.size = Pt(8)
            f.text_frame.paragraphs[0].font.color.rgb = RGBColor(128,128,128)
            f.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

            # Slide number
            n = slide.shapes.add_textbox(Inches(9.5), Inches(6.5), Inches(0.5), Inches(0.5))
            n.text_frame.text = num
            n.text_frame.paragraphs[0].font.size = Pt(12)
            n.text_frame.paragraphs[0].font.color.rgb = RGBColor(128,128,128)
            n.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

        # ---- slide helpers -----------------------------------------
        def add_title_slide(title, subtitle=None):
            s = prs.slides.add_slide(prs.slide_layouts[0])
            s.shapes.title.text = title
            s.shapes.title.text_frame.paragraphs[0].font.size = Pt(44)
            s.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
            if subtitle:
                s.placeholders[1].text = subtitle
                s.placeholders[1].text_frame.paragraphs[0].font.size = Pt(32)
                s.placeholders[1].text_frame.paragraphs[0].font.color.rgb = RGBColor(255,0,0)
            add_header_footer_number(s, str(len(prs.slides)))
            return s

        def add_bullet_slide(title, bullets):
            s = prs.slides.add_slide(prs.slide_layouts[1])
            s.shapes.title.text = title
            s.shapes.title.text_frame.paragraphs[0].font.size = Pt(28)
            s.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
            tf = s.placeholders[1].text_frame
            tf.clear()
            for b in bullets:
                p = tf.add_paragraph()
                p.text = b
                p.level = 0
                p.font.size = Pt(18)
                p.font.color.rgb = RGBColor(255,255,255)
                p.bullet.color.rgb = RGBColor(0,102,204)
            add_header_footer_number(s, str(len(prs.slides)))
            return s

        def add_table_slide(title, rows, cols, data):
            s = prs.slides.add_slide(prs.slide_layouts[5])
            # title
            tb = s.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(0.5))
            tb.text_frame.text = title
            tb.text_frame.paragraphs[0].font.size = Pt(28)
            tb.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)

            # table
            tbl = s.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5),
                                    Inches(9), Inches(4)).table
            # headers
            for i, h in enumerate(data[0]):
                c = tbl.cell(0, i)
                c.text = h
                c.fill.solid()
                c.fill.fore_color.rgb = RGBColor(0,102,204)
                c.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
                c.text_frame.paragraphs[0].font.bold = True
                c.text_frame.paragraphs[0].font.size = Pt(14)
            # data rows
            for r, row in enumerate(data[1:], 1):
                for c, val in enumerate(row):
                    cell = tbl.cell(r, c)
                    cell.text = str(val)
                    cell.text_frame.paragraphs[0].font.size = Pt(12)
                    cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(0,0,0)
                    if r % 2 == 1:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(242,242,242)
            add_header_footer_number(s, str(len(prs.slides)))
            return s

        # ---- progress ------------------------------------------------
        prog = st.progress(0)
        total = 11
        cur = 0

        # 1 – Title
        title_slide = add_title_slide(
            "Professional Services Transition Meeting",
            f"{customer_name}\n{today_date}"
        )
        # office background picture
        img = requests.get(
            "https://thumbs.dreamstime.com/b/large-empty-office-many-people-their-desks-busy-working-spacious-unoccupied-space-numerous-individuals-diligently-379931527.jpg"
        )
        if img.status_code == 200:
            pic = io.BytesIO(img.content)
            title_slide.shapes.add_picture(pic, Inches(0), Inches(0),
                                          prs.slide_width, prs.slide_height)
        cur += 1; prog.progress(cur/total)

        # 2 – Agenda
        add_bullet_slide("Meeting Agenda",
                         ["Project Summary", "Technical Summary",
                          "Recommended Next Steps"])
        cur += 1; prog.progress(cur/total)

        # 3 – Project Summary title
        add_title_slide("Project Summary")
        cur += 1; prog.progress(cur/total)

        # 4 – Status report (multiple tables)
        add_table_slide("Final Project Status Report – " + customer_name,
                        2, 3, [["Today's Date","Start Date","End Date"],
                               [today_date, project_start, project_end]])
        add_table_slide("Milestones",
                        len(milestones_data)+1, 4,
                        [["Milestone","Baseline Date","Target Completion Date","Status"]] +
                        [[m["name"],m["baseline"],m["target"],m["status"]] for m in milestones_data])
        add_table_slide("User Rollout Roadmap",
                        3, 5,
                        [["Milestone","Target Users","Current Users","Target Completion","Status"],
                         ["Pilot",pilot_target,pilot_current,pilot_completion,pilot_status],
                         ["Production",prod_target,prod_current,prod_completion,prod_status]])
        add_table_slide("Project Objectives",
                        len(objectives_data)+1, 3,
                        [["Planned Project Objective (Target)","Actual Project Result (Actual)","Deviation/Cause"]] +
                        [[o["objective"],o["actual"],o["deviation"]] for o in objectives_data])
        # summary text under the last table
        last = prs.slides[-1]
        txt = last.shapes.add_textbox(Inches(0.5), Inches(5), Inches(9), Inches(1))
        txt.text_frame.text = project_summary_text
        txt.text_frame.paragraphs[0].font.size = Pt(14)
        txt.text_frame.paragraphs[0].font.color.rgb = RGBColor(0,0,0)
        cur += 1; prog.progress(cur/total)

        # 5 – Deliverables
        del_slide = add_table_slide("Deliverables",
                                    len(deliverables_data)+1, 2,
                                    [["Deliverable","Date delivered"]] +
                                    [[d["name"],d["date"]] for d in deliverables_data])
        # ---- checkmarks (fixed indentation) ----
        for row_idx in range(1, len(deliverables_data) + 1):
            check = del_slide.shapes.add_shape(
                MSO_SHAPE.CHECKMARK,
                Inches(0.3),
                Inches(1.5) + Inches(0.3) * (row_idx - 1),
                Inches(0.3),
                Inches(0.3)
            )
            check.fill.solid()
            check.fill.fore_color.rgb = RGBColor(0, 176, 80)   # green
        cur += 1; prog.progress(cur/total)

        # 6 – Technical Summary title
        add_title_slide("Technical Summary")
        cur += 1; prog.progress(cur/total)

        # 7 – ZIA Architecture
        zia = prs.slides.add_slide(prs.slide_layouts[5])
        tbox = zia.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(0.5))
        tbox.text_frame.text = "Deployed ZIA Architecture"
        tbox.text_frame.paragraphs[0].font.size = Pt(28)
        tbox.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)

        # simple diagram (central → Z-Tunnels)
        ca = zia.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2,2), Inches(2,1))
        ca.fill.solid(); ca.fill.fore_color.rgb = RGBColor(0,102,204)
        ca.text_frame.text = "Central Authority"
        ca.text_frame.paragraphs[0].font.size = Pt(12)
        ca.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
        ca.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        tun = zia.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5,2), Inches(2,1))
        tun.fill.solid(); tun.fill.fore_color.rgb = RGBColor(0,102,204)
        tun.text_frame.text = "Z-Tunnels"
        tun.text_frame.paragraphs[0].font.size = Pt(12)
        tun.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
        tun.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

        conn = zia.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,
                                        Inches(4,2.5), Inches(5,2.5))
        conn.line.color.rgb = RGBColor(0,102,204)
        conn.line.width = Pt(2)

        # tech details
        tech = zia.shapes.add_textbox(Inches(0.5,3), Inches(4,3))
        tech.text_frame.text = (
            f"Identity Provider: {idp}\nAuthentication Type: {auth_type}\n"
            f"User/Group Provisioning: {prov_type}\nTunnel Type: {tunnel_type}\n"
            f"ZCC Deployment System: {deploy_system}\nWindows Devices: {windows_num}\n"
            f"MacOS Devices: {mac_num}\nGeo Locations: {geo_locations}\n"
            f"SSL Inspection Policies: {ssl_policies}\nURL Filtering Policies: {url_policies}\n"
            f"Cloud App Control Policies: {cloud_policies}\nFirewall Policies: {fw_policies}"
        )
        for p in tech.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(255,255,255)
        add_header_footer_number(zia, "7")
        cur += 1; prog.progress(cur/total)

        # 8 – Open Items
        oi_headers = ["Task/ Description","Date","Owner","Transition Plan/ Next Steps"]
        oi_rows = [[oi["task"],oi["date"],oi["owner"],oi["steps"]] for oi in open_items_data]
        add_table_slide("Open Items", len(oi_rows)+1, 4, [oi_headers] + oi_rows)
        cur += 1; prog.progress(cur/total)

        # 9 – Next Steps (two columns)
        ns = prs.slides.add_slide(prs.slide_layouts[5])
        ns_title = ns.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(0.5))
        ns_title.text_frame.text = "Recommended Next Steps"
        ns_title.text_frame.paragraphs[0].font.size = Pt(28)
        ns_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)

        # short term (green)
        short = ns.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5,1.5), Inches(4.5,4))
        short.fill.solid(); short.fill.fore_color.rgb = RGBColor(0,176,80)
        stf = short.text_frame
        stf.text = "Short Term Activities"
        stf.paragraphs[0].font.size = Pt(18)
        stf.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
        stf.paragraphs[0].alignment = PP_ALIGN.CENTER
        for it in short_term:
            p = stf.add_paragraph()
            p.text = "• " + it
            p.level = 0
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(0,0,0)

        # long term (blue)
        long = ns.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5,1.5), Inches(4.5,4))
        long.fill.solid(); long.fill.fore_color.rgb = RGBColor(0,102,204)
        ltf = long.text_frame
        ltf.text = "Long Term Activities"
        ltf.paragraphs[0].font.size = Pt(18)
        ltf.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
        ltf.paragraphs[0].alignment = PP_ALIGN.CENTER
        for it in long_term:
            p = ltf.add_paragraph()
            p.text = "• " + it
            p.level = 0
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(255,255,255)
        add_header_footer_number(ns, "9")
        cur += 1; prog.progress(cur/total)

        # 10 – Thank you
        th = prs.slides.add_slide(prs.slide_layouts[1])
        th.shapes.title.text = "Thank you"
        th.shapes.title.text_frame.paragraphs[0].font.size = Pt(44)
        th.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
        tf = th.placeholders[1].text_frame
        tf.clear()
        p = tf.add_paragraph()
        p.text = (
            f"Your feedback on our project and Professional Services team is important to us.\n\n"
            f"Project Manager: {pm_name}\nConsultant: {consultant_name}\n\n"
            f"A short ~6 question survey on how your Professional Services team did will be automatically sent after the project has closed. "
            f"The following people will receive the survey via email:\n\n"
            f"Primary Contact: {primary_contact}\nSecondary Contact: {secondary_contact}\n"
            f"We appreciate any insights you can provide to help us improve our processes and ensure we provide the best possible service in future projects."
        )
        for para in tf.paragraphs:
            para.font.size = Pt(14)
            para.font.color.rgb = RGBColor(0,0,0)

        bubble = th.shapes.add_shape(MSO_SHAPE.CLOUD_CALL_OUT, Inches(7,3), Inches(2,1))
        bubble.fill.solid(); bubble.fill.fore_color.rgb = RGBColor(255,192,0)
        bubble.text_frame.text = "We want to know!"
        bubble.text_frame.paragraphs[0].font.size = Pt(18)
        bubble.text_frame.paragraphs[0].font.color.rgb = RGBColor(0,0,0)
        add_header_footer_number(th, "10")
        cur += 1; prog.progress(cur/total)

        # 11 – Final Thank you
        add_title_slide("Thank you")
        cur += 1; prog.progress(cur/total)

        # ----- save -------------------------------------------------
        bio = io.BytesIO()
        prs.save(bio)
        bio.seek(0)
        st.success("Deck ready!")
        st.download_button(
            label="Download PPTX",
            data=bio.getvalue(),
            file_name=f"{customer_name}_Transition_Deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
