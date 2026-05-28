import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Aerobond Dashboard", layout="wide", initial_sidebar_state="expanded")

PRIMARY_RED = "#D71920"
CHARCOAL = "#2B2B2B"
BG = "#F7F8FA"
CARD = "#FFFFFF"
GRAY = "#7A7F87"
GREEN = "#2E8B57"
AMBER = "#D99A00"
RED = "#C62828"

st.markdown(
    f"""
    <style>
    .stApp {{
        background: {BG};
        color: {CHARCOAL};
    }}
    section[data-testid="stSidebar"] {{
        background: #111111;
    }}
    section[data-testid="stSidebar"] * {{
        color: white;
    }}
    .brand-box {{
        padding: 1rem 1rem 0.5rem 1rem;
        border-radius: 14px;
        background: linear-gradient(135deg, {CARD} 0%, #fff 100%);
        border: 1px solid #e8e8e8;
        margin-bottom: 1rem;
    }}
    .kpi-card {{
        background: {CARD};
        border: 1px solid #eaeaea;
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
        min-height: 118px;
    }}
    .kpi-title {{
        font-size: 0.85rem;
        color: {GRAY};
        margin-bottom: 0.35rem;
    }}
    .kpi-value {{
        font-size: 2rem;
        font-weight: 700;
        color: {CHARCOAL};
        line-height: 1.1;
    }}
    .kpi-delta {{
        font-size: 0.85rem;
        color: {GRAY};
        margin-top: 0.3rem;
    }}
    .section-card {{
        background: {CARD};
        border: 1px solid #eaeaea;
        border-radius: 16px;
        padding: 1rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.04);
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

def init_state():
    if "auth" not in st.session_state:
        st.session_state.auth = False
    if "snapshots" not in st.session_state:
        st.session_state.snapshots = []
    if "uploaded_names" not in st.session_state:
        st.session_state.uploaded_names = set()

def check_password():
    secret_pw = st.secrets.get("password", "")
    with st.sidebar:
        st.markdown("### Login")
        username = st.text_input("Username", value="", key="username")
        password = st.text_input("Password", type="password", key="password")
        if st.button("Sign in", use_container_width=True):
            if password == secret_pw and secret_pw != "":
                st.session_state.auth = True
                st.success("Signed in")
            else:
                st.error("Invalid password")
    return st.session_state.auth

def read_workbook(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xls.sheet_names:
        sheets[name] = pd.read_excel(uploaded_file, sheet_name=name)
    return sheets

def safe_first(df, col, default=None):
    if df is None or df.empty or col not in df.columns:
        return default
    vals = df[col].dropna()
    if len(vals) == 0:
        return default
    return vals.iloc[0]

def parse_snapshot(uploaded_file):
    sheets = read_workbook(uploaded_file)
    md = sheets.get("Meeting Details", pd.DataFrame())
    summary = sheets.get("Summary", pd.DataFrame())
    actions = sheets.get("Actions Items", pd.DataFrame())
    risks = sheets.get("Risk Review", pd.DataFrame())
    trans = sheets.get("Document Transmittal", pd.DataFrame())
    next_steps = sheets.get("Next Steps", pd.DataFrame())
    production = sheets.get("Summary", pd.DataFrame())
    snapshot = {
        "filename": uploaded_file.name,
        "project_name": safe_first(md, "Project Name", "Unknown"),
        "project_no": safe_first(md, "Project No.", ""),
        "report_date": pd.to_datetime(safe_first(md, "Date", pd.Timestamp.today()), errors="coerce"),
        "reporting_staff": safe_first(md, "Reporting Staff", ""),
        "summary": summary,
        "actions": actions,
        "risks": risks,
        "transmittal": trans,
        "next_steps": next_steps,
        "production": production,
        "sheets": sheets,
    }
    return snapshot

def compute_kpis(snapshot, previous=None):
    actions = snapshot["actions"]
    risks = snapshot["risks"]
    summary = snapshot["summary"]
    report_date = snapshot["report_date"]

    open_actions = 0
    overdue_actions = 0
    due_next_7_days = 0

    if not actions.empty:
        if "Status" in actions.columns:
            status = actions["Status"].astype(str).str.lower()
            open_actions = int((status != "closed").sum())
            if "Due Date" in actions.columns:
                due = pd.to_datetime(actions["Due Date"], errors="coerce")
                overdue_actions = int(((due < report_date) & (status != "closed")).sum())
                due_next_7_days = int(((due >= report_date) & (due <= report_date + pd.Timedelta(days=7))).sum())

    open_risks = 0
    high_risks = 0
    if not risks.empty:
        if "Status" in risks.columns:
            open_risks = int((risks["Status"].astype(str).str.lower() != "closed").sum())
        if "Risk Level" in risks.columns:
            high_risks = int((risks["Risk Level"].astype(str).str.lower() == "high").sum())

    progress_value = None
    if not summary.empty:
        for col in ["Progress ex.gst", "Progress ex GST", "Progress"]:
            if col in summary.columns:
                progress_value = pd.to_numeric(summary[col], errors="coerce").fillna(0).sum()
                break

    previous_progress = previous.get("weekly_progress") if previous else None
    weekly_delta = None
    if progress_value is not None and previous_progress is not None:
        weekly_delta = progress_value - previous_progress

    scheduled_progress = None
    actual_progress = progress_value
    schedule_variance = None
    if actual_progress is not None and scheduled_progress is not None:
        schedule_variance = actual_progress - scheduled_progress

    program_health = "Green"
    if overdue_actions > 3 or high_risks >= 1:
        program_health = "Red"
    elif overdue_actions >= 1 or open_risks >= 1:
        program_health = "Amber"

    return {
        "program_health": program_health,
        "open_actions": open_actions,
        "overdue_actions": overdue_actions,
        "open_risks": open_risks,
        "high_risks": high_risks,
        "weekly_progress": progress_value,
        "weekly_delta": weekly_delta,
        "actual_progress": actual_progress,
        "scheduled_progress": scheduled_progress,
        "schedule_variance": schedule_variance,
        "due_next_7_days": due_next_7_days,
    }

@dataclass
class Snapshot:
    filename: str
    project_name: str
    project_no: str
    report_date: object
    reporting_staff: str
    summary: pd.DataFrame
    actions: pd.DataFrame
    risks: pd.DataFrame
    transmittal: pd.DataFrame
    next_steps: pd.DataFrame
    production: pd.DataFrame
    sheets: dict
    kpis: dict

def add_snapshot(parsed, kpis):
    st.session_state.snapshots.append(
        Snapshot(
            filename=parsed["filename"],
            project_name=parsed["project_name"],
            project_no=parsed["project_no"],
            report_date=parsed["report_date"],
            reporting_staff=parsed["reporting_staff"],
            summary=parsed["summary"],
            actions=parsed["actions"],
            risks=parsed["risks"],
            transmittal=parsed["transmittal"],
            next_steps=parsed["next_steps"],
            production=parsed["production"],
            sheets=parsed["sheets"],
            kpis=kpis,
        )
    )
    st.session_state.uploaded_names.add(parsed["filename"])

def latest_snapshot():
    if not st.session_state.snapshots:
        return None
    return st.session_state.snapshots[-1]

def previous_snapshot():
    if len(st.session_state.snapshots) < 2:
        return None
    return st.session_state.snapshots[-2]

def render_header():
    snap = latest_snapshot()
    logo_cols = st.columns([1, 3, 1])
    with logo_cols[0]:
        st.image("AEROBOND-logo_tagline-AV-DEF-SP_RGB-2.png", width=220)
    with logo_cols[1]:
        st.markdown(
            f"""
            <div class="brand-box">
            <div style="font-size: 1.15rem; font-weight: 700; color: {PRIMARY_RED};">Aerobond Defence Program Dashboard</div>
            <div style="color: {GRAY};">Password-protected weekly project reporting and KPI tracking</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with logo_cols[2]:
        if snap:
            st.markdown(
                f"""
                <div class="brand-box">
                <div style="font-size: 0.85rem; color: {GRAY};">Current snapshot</div>
                <div style="font-weight: 700;">{snap.filename}</div>
                <div style="color: {GRAY};">{snap.project_name}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

def kpi_card(title, value, delta=None, accent=PRIMARY_RED):
    delta_text = "" if delta is None else f"<div class='kpi-delta'>{delta}</div>"
    st.markdown(
        f"""
        <div class="kpi-card" style="border-top: 4px solid {accent};">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            {delta_text}
        </div>
        """,
        unsafe_allow_html=True,
    )

def overview_page():
    snap = latest_snapshot()
    if not snap:
        st.info("Upload one or more Excel reports to start.")
        return
    k = snap.kpis
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1: kpi_card("Program Health", k["program_health"], accent=GREEN if k["program_health"] == "Green" else AMBER if k["program_health"] == "Amber" else RED)
    with c2: kpi_card("Actual vs Scheduled", k["schedule_variance"] if k["schedule_variance"] is not None else "—", accent=PRIMARY_RED)
    with c3: kpi_card("Weekly Progress", k["weekly_progress"] if k["weekly_progress"] is not None else "—", f"Δ {k['weekly_delta']}" if k["weekly_delta"] is not None else None)
    with c4: kpi_card("Open Actions", k["open_actions"], f"Overdue: {k['overdue_actions']}", accent=AMBER if k["overdue_actions"] else GREEN)
    with c5: kpi_card("Open Risks", k["open_risks"], f"High: {k['high_risks']}", accent=RED if k["high_risks"] else GREEN)
    with c6: kpi_card("Due Next 7 Days", k["due_next_7_days"], accent=PRIMARY_RED)

    left, right = st.columns([2.4, 1])
    with left:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Weekly trend")
        hist = pd.DataFrame([
            {"date": s.report_date, "progress": s.kpis.get("weekly_progress"), "name": s.filename}
            for s in st.session_state.snapshots
        ]).dropna(subset=["date"])
        if len(hist) >= 2:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=hist["date"], y=hist["progress"], mode="lines+markers", name="Progress", line=dict(color=PRIMARY_RED, width=3)))
            fig.update_layout(height=360, margin=dict(l=10, r=10, t=20, b=10), plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Add at least two uploads to see trends.")
        st.markdown("</div>", unsafe_allow_html=True)
    with right:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("RAG summary")
        st.write(f"**Program health:** {k['program_health']}")
        st.write(f"**Open actions:** {k['open_actions']}")
        st.write(f"**Overdue actions:** {k['overdue_actions']}")
        st.write(f"**Open risks:** {k['open_risks']}")
        st.write(f"**High risks:** {k['high_risks']}")
        st.write(f"**Due next 7 days:** {k['due_next_7_days']}")
        st.markdown("</div>", unsafe_allow_html=True)

    a, r, p = st.columns(3)
    with a:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Actions")
        st.dataframe(snap.actions, use_container_width=True, height=320)
        st.markdown("</div>", unsafe_allow_html=True)
    with r:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Risks")
        st.dataframe(snap.risks, use_container_width=True, height=320)
        st.markdown("</div>", unsafe_allow_html=True)
    with p:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        st.subheader("Production / Delivery")
        st.write("WIP, QA Ready, Quarantine and completion metrics should be connected from the Summary sheet.")
        st.markdown("</div>", unsafe_allow_html=True)

def uploads_page():
    st.subheader("Upload reports")
    uploaded = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)
    if uploaded:
        for f in uploaded:
            if f.name in st.session_state.uploaded_names:
                st.warning(f"{f.name} already uploaded in this session.")
                continue
            try:
                parsed = parse_snapshot(f)
                prev = latest_snapshot()
                kpis = compute_kpis(parsed, previous=prev.kpis if prev else None)
                add_snapshot(parsed, kpis)
                st.success(f"Loaded {f.name}")
            except Exception as e:
                st.error(f"Failed to read {f.name}: {e}")

    if st.session_state.snapshots:
        rows = []
        for s in st.session_state.snapshots:
            rows.append({
                "filename": s.filename,
                "project_name": s.project_name,
                "project_no": s.project_no,
                "report_date": s.report_date,
                "program_health": s.kpis.get("program_health"),
                "open_actions": s.kpis.get("open_actions"),
                "open_risks": s.kpis.get("open_risks"),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

def schedule_page():
    snap = latest_snapshot()
    if not snap:
        st.info("Upload a report first.")
        return
    st.subheader("Schedule & progress")
    st.write("This page is ready for milestone/Gantt logic from the transmittal sheet.")
    st.dataframe(snap.transmittal, use_container_width=True)

def actions_page():
    snap = latest_snapshot()
    if not snap:
        st.info("Upload a report first.")
        return
    st.subheader("Actions & tasks")
    st.dataframe(snap.actions, use_container_width=True, height=600)

def risks_page():
    snap = latest_snapshot()
    if not snap:
        st.info("Upload a report first.")
        return
    st.subheader("Risks")
    st.dataframe(snap.risks, use_container_width=True, height=600)

def production_page():
    snap = latest_snapshot()
    if not snap:
        st.info("Upload a report first.")
        return
    st.subheader("Production / delivery")
    st.write("This page should be linked to production summary fields.")
    st.dataframe(snap.summary, use_container_width=True, height=600)

def history_page():
    if not st.session_state.snapshots:
        st.info("No snapshots yet.")
        return
    st.subheader("Snapshot history")
    rows = []
    for s in st.session_state.snapshots:
        rows.append({
            "filename": s.filename,
            "project_name": s.project_name,
            "report_date": s.report_date,
            "weekly_progress": s.kpis.get("weekly_progress"),
            "open_actions": s.kpis.get("open_actions"),
            "open_risks": s.kpis.get("open_risks"),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True)

def admin_page():
    st.subheader("Admin settings")
    st.write("Add user management, thresholds, and retention controls here.")
    st.write("Password is stored in Streamlit secrets.")

def main():
    init_state()
    with st.sidebar:
        st.image("AEROBOND-logo_tagline-AV-DEF-SP_RGB-2.png", use_container_width=True)
        st.markdown("### Aerobond Dashboard")
        if not st.session_state.auth:
            check_password()
        else:
            st.success("Authenticated")
        st.divider()
        page = st.radio(
            "Navigate",
            ["Overview", "Uploads", "Schedule", "Actions", "Risks", "Production", "History", "Admin"],
        )
        st.divider()
        if st.button("Sign out"):
            st.session_state.auth = False
            st.rerun()

    if not st.session_state.auth:
        st.title("Aerobond Defence Dashboard")
        st.info("Enter the password in the sidebar to continue.")
        return

    render_header()
    if page == "Overview":
        overview_page()
    elif page == "Uploads":
        uploads_page()
    elif page == "Schedule":
        schedule_page()
    elif page == "Actions":
        actions_page()
    elif page == "Risks":
        risks_page()
    elif page == "Production":
        production_page()
    elif page == "History":
        history_page()
    else:
        admin_page()

if __name__ == "__main__":
    main()
