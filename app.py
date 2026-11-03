import html
from pathlib import Path
import pandas as pd
import streamlit as st

# ============================================================
# BLUEPATH
# Blue Economy × Lean Construction
# Interactive Research Decision-Support Platform
# ============================================================

st.set_page_config(
    page_title="BluePath | Blue Economy × Lean Construction",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------
BASE_DIR = Path(__file__).parent
EXCEL_CANDIDATES = [BASE_DIR / "portal.xlsx", BASE_DIR / "Portal.xlsx"]
EXCEL_FILE = next((f for f in EXCEL_CANDIDATES if f.exists()), EXCEL_CANDIDATES[0])

IMPACT_COL = "Impact of Marine Construction"
PRACTICE_COL = "Blue Economy Practice"
CHALLENGE_COL = "Challenges to Implement Blue Economy Practice"
LEAN_COL = "Lean Principles"
REQUIRED_COLUMNS = [IMPACT_COL, PRACTICE_COL, CHALLENGE_COL, LEAN_COL]
EXPECTED_LEAN_COUNT = 5

# ------------------------------------------------------------
# ADVANCED CUSTOM DESIGN & PROFESSIONAL STYLING
# ------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

.stApp {
    background: linear-gradient(135deg, #F4F7F6 0%, #FAFDFC 100%);
    color: #1E293B;
    font-family: 'Inter', system-ui, sans-serif;
}

#MainMenu, footer, header { visibility: hidden; }

.block-container {
    max-width: 1200px;
    padding-top: 1.5rem;
    padding-bottom: 4rem;
}

/* ---------- Sidebar Styling ---------- */
[data-testid="stSidebar"] {
    background-color: #0F2D33;
    color: #F8FAFC;
    padding: 20px 10px;
}
[data-testid="stSidebar"] .sidebar-brand {
    font-size: 1.5rem;
    font-weight: 800;
    color: #38BDF8;
    letter-spacing: -0.5px;
    margin-bottom: 4px;
}
[data-testid="stSidebar"] .sidebar-sub {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 2px;
    color: #94A3B8;
    text-transform: uppercase;
    margin-bottom: 24px;
}
[data-testid="stSidebar"] hr {
    border-color: #1E3A40;
    margin: 20px 0;
}

/* ---------- Typography & Headers ---------- */
.brand {
    font-size: 2rem;
    font-weight: 800;
    letter-spacing: -0.8px;
    color: #0A6E74;
    margin-bottom: 2px;
}
.brand-subtitle {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 3px;
    color: #5A7A84;
    text-transform: uppercase;
}

.hero {
    background: linear-gradient(135deg, #0A6A72 0%, #0D8A8F 55%, #1F9D8A 100%);
    border-radius: 20px;
    padding: 45px 40px;
    margin: 15px 0 30px 0;
    color: white;
    box-shadow: 0 14px 32px rgba(10, 110, 116, 0.16);
}
.hero h1 {
    font-size: 2.5rem;
    font-weight: 800;
    line-height: 1.2;
    margin: 8px 0 14px 0;
    color: white;
}
.hero p {
    font-size: 1.05rem;
    line-height: 1.75;
    color: #E8F7F7;
    max-width: 800px;
    margin: 0;
}
.hero-label {
    font-size: 0.75rem;
    letter-spacing: 2.8px;
    font-weight: 700;
    color: #C8EFEF;
    text-transform: uppercase;
}

.section-title {
    font-size: 1.6rem;
    font-weight: 800;
    color: #0D5F66;
    margin: 6px 0 8px 0;
}
.section-description {
    font-size: 1rem;
    line-height: 1.65;
    color: #4A6A74;
    margin-bottom: 20px;
}
.step-label {
    font-size: 0.75rem;
    letter-spacing: 2.5px;
    font-weight: 750;
    color: #0D8A8A;
    text-transform: uppercase;
    margin-bottom: 6px;
}

/* ---------- Panels & Cards ---------- */
.content-panel {
    background: rgba(255, 255, 255, 0.95);
    border: 1px solid #D5E6E8;
    border-radius: 18px;
    padding: 28px;
    margin-bottom: 20px;
    box-shadow: 0 6px 20px rgba(30, 80, 90, 0.04);
}

.info-card {
    background: linear-gradient(135deg, #FFFFFF 0%, #FAFCFC 100%);
    border: 1px solid #D5E6E8;
    border-left: 5px solid #0A6E74;
    border-radius: 14px;
    padding: 18px 20px;
    margin-bottom: 14px;
    box-shadow: 0 2px 10px rgba(30, 80, 90, 0.03);
    transition: all 0.25s ease;
}
.info-card:hover {
    transform: translateY(-3px);
    border-left-color: #E05A5A;
    box-shadow: 0 8px 22px rgba(224, 90, 90, 0.12);
    background: #FFF9F9;
}

.pathway-card {
    background: #FFFFFF;
    border: 1px solid #D5E6E8;
    border-left: 5px solid #0A6E74;
    border-radius: 14px;
    padding: 18px 20px;
    margin-bottom: 14px;
    transition: all 0.25s ease;
}
.pathway-card:hover {
    border-left-color: #E05A5A;
    background: #FFF8F8;
    transform: translateX(4px);
    box-shadow: 0 6px 18px rgba(224, 90, 90, 0.12);
}

.number-circle {
    display: inline-flex;
    justify-content: center;
    align-items: center;
    width: 34px;
    height: 34px;
    border-radius: 50%;
    background: #E4F3F1;
    color: #0A6E74;
    font-size: 0.9rem;
    font-weight: 750;
    margin-right: 14px;
    flex-shrink: 0;
}

/* ---------- Lean Principle Cards ---------- */
.lean-card {
    background: linear-gradient(120deg, #EDF7F1 0%, #F6FBF7 100%);
    border: 1px solid #C5E0CF;
    border-left: 5px solid #1E7A4A;
    border-radius: 14px;
    padding: 18px 20px;
    margin-bottom: 14px;
    transition: all 0.22s ease;
}
.lean-card:hover {
    transform: translateX(5px);
    box-shadow: 0 6px 16px rgba(40, 120, 80, 0.10);
}
.priority-label {
    font-size: 0.7rem;
    font-weight: 750;
    letter-spacing: 1.2px;
    color: #1E7A4A;
    text-transform: uppercase;
    margin-bottom: 4px;
}

/* ---------- Notice Banner ---------- */
.notice {
    background: linear-gradient(135deg, #EAF5F5 0%, #E2F0F0 100%);
    border-left: 5px solid #0D8A8A;
    border-radius: 12px;
    padding: 16px 20px;
    margin: 16px 0 22px 0;
    color: #2A4A52;
    line-height: 1.65;
    font-size: 0.98rem;
    box-shadow: 0 2px 8px rgba(13, 138, 138, 0.05);
}

/* ---------- Buttons ---------- */
.stButton > button {
    min-height: 46px;
    border-radius: 12px;
    border: 1px solid #C5DCD9;
    background: #FFFFFF;
    color: #0A6E74;
    font-weight: 650;
    transition: all 0.22s ease;
}
.stButton > button:hover {
    background: #FFF1F1;
    border-color: #E05A5A;
    color: #C0392B;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(224, 90, 90, 0.15);
}
.stButton > button[kind="primary"] {
    background: linear-gradient(100deg, #0A6E74, #1A9A90);
    color: white;
    border: none;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(100deg, #C0392B, #E05A5A);
    color: white;
    box-shadow: 0 4px 12px rgba(224, 90, 90, 0.25);
}

/* ---------- Footer ---------- */
.footer {
    text-align: center;
    color: #6A848C;
    font-size: 0.82rem;
    border-top: 1px solid #D8E8EA;
    padding-top: 20px;
    margin-top: 50px;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# LOAD DATA
# ------------------------------------------------------------
@st.cache_data
def load_database(file_path, file_modified):
    if not Path(file_path).exists():
        raise FileNotFoundError(
            "Excel database not found. Place portal.xlsx (or Portal.xlsx) next to app.py."
        )
    data = pd.read_excel(file_path, engine="openpyxl")
    data.columns = data.columns.astype(str).str.strip()
    missing = [c for c in REQUIRED_COLUMNS if c not in data.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    data = data[REQUIRED_COLUMNS].copy()
    for col in REQUIRED_COLUMNS:
        data[col] = data[col].fillna("").astype(str).str.strip()
    data = data.replace("", pd.NA).dropna(subset=REQUIRED_COLUMNS)
    return data.reset_index(drop=True)

try:
    df = load_database(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0
    )
except Exception as e:
    st.error(f"Database loading error: {e}")
    st.stop()

# ------------------------------------------------------------
# SESSION STATE
# ------------------------------------------------------------
DEFAULTS = {
    "page": "welcome",
    "username": "",
    "selected_impact": None,
    "selected_practice": None,
    "selected_challenge": None
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_app():
    st.session_state.page = "welcome"
    st.session_state.username = ""
    st.session_state.selected_impact = None
    st.session_state.selected_practice = None
    st.session_state.selected_challenge = None

def go_home():
    st.session_state.page = "impacts"
    st.session_state.selected_impact = None
    st.session_state.selected_practice = None
    st.session_state.selected_challenge = None

# ------------------------------------------------------------
# SIDEBAR NAVIGATION & APP METRICS
# ------------------------------------------------------------
with st.sidebar:
    st.markdown('<div class="sidebar-brand">🌊 BluePath</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Decision Support Matrix</div>', unsafe_allow_html=True)
    
    if st.session_state.page != "welcome":
        if st.button("⌂ Return to Impacts Hub", use_container_width=True):
            go_home()
            st.rerun()
        if st.button("🔄 Restart Decision Pathway", use_container_width=True):
            reset_app()
            st.rerun()
        st.markdown("---")
    
    st.markdown("### 📊 Database Metrics")
    st.markdown(f"* **Total Records:** {len(df)}")
    st.markdown(f"* **Unique Impacts:** {df[IMPACT_COL].nunique()}")
    st.markdown(f"* **Blue Practices:** {df[PRACTICE_COL].nunique()}")
    st.markdown(f"* **Challenges Mapped:** {df[CHALLENGE_COL].nunique()}")
    
    st.markdown("---")
    st.markdown("### ℹ️ Research Guide")
    st.caption("BluePath bridges marine engineering construction impacts with Blue Economy sustainability frameworks through prioritized Lean construction solutions.")

# ------------------------------------------------------------
# HEADER PROGRESS
# ------------------------------------------------------------
PAGES = {"welcome": 0, "impacts": 1, "practices": 2, "challenges": 3, "lean": 4}
if st.session_state.page != "welcome":
    step = PAGES[st.session_state.page]
    st.progress(step / 4)
    st.caption(f"Decision Pathway Progress  →  Step {step} of 4")

# ============================================================
# 8. WELCOME & INTRODUCTION PAGE
# ============================================================
if st.session_state.page == "welcome":

    st.markdown("""
        <div class="hero">
            <div class="hero-label">Research Decision-Support Platform</div>
            <h1>BluePath: Blue Economy × Lean Construction</h1>
            <p>
                An interactive framework designed to assess marine construction impacts, 
                explore robust Blue Economy responses, and implement prioritized Lean principles 
                for sustainable project delivery.
            </p>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-title">Welcome & Getting Started</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-description">'
        'Follow our structured 4-step decision pathway. Each step dynamically filters database records '
        'to match your specific marine engineering context.'
        '</div>',
        unsafe_allow_html=True
    )

    username = st.text_input(
        "Your Name or Professional Title (Optional)",
        placeholder="e.g., Dr. Silva / Senior Researcher",
        max_chars=60
    )

    st.markdown("#### 🧭 Core Decision Pathway")

    steps = [
        ("01", "Select Marine Construction Impact",
         "Identify environmental, social, or structural impacts caused by marine activities."),
        ("02", "Select Blue Economy Practice",
         "Discover suitable sustainability practices tailored to mitigate the selected impact."),
        ("03", "Select Implementation Challenge",
         "Examine specific operational and logistical barriers in marine environments."),
        ("04", "Review Prioritized Lean Principles",
         "Analyze the exact 5-tier Lean priority ranking derived from empirical research records.")
    ]

    for num, title, desc in steps:
        st.markdown(f"""
            <div class="pathway-card">
                <span class="number-circle">{num}</span>
                <strong>{title}</strong>
                <p style="color:#4A6A74; margin:7px 0 0 48px; line-height:1.5; font-size:0.94rem;">
                    {desc}
                </p>
            </div>
        """, unsafe_allow_html=True)

    st.write("")
    if st.button("🚀 Launch Decision Pathway", type="primary", use_container_width=True):
        st.session_state.username = username.strip()
        st.session_state.page = "impacts"
        st.rerun()

# ============================================================
# 9. STEP 1: IMPACTS SELECTION
# ============================================================
elif st.session_state.page == "impacts":

    name = html.escape(st.session_state.username) if st.session_state.username else "Researcher"

    st.markdown(f'<div class="section-title">Welcome, {name}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-description">'
        'Step 1: Choose the marine construction impact you wish to analyze and manage.'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown("""
        <div class="notice">
            <strong>Step 1 Guidance</strong><br>
            Impacts are listed in original database sequence. Click any card below to filter associated Blue Economy practices.
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)

    impacts = df[IMPACT_COL].drop_duplicates().tolist()
    st.markdown(f"**{len(impacts)} construction impacts available in database**")
    st.write("")

    cols = st.columns(2)
    for i, impact in enumerate(impacts):
        with cols[i % 2]:
            st.markdown(f"""
                <div class="info-card">
                    <span class="number-circle">{i+1}</span>
                    <strong>{html.escape(impact)}</strong>
                </div>
            """, unsafe_allow_html=True)
            if st.button(f"Explore Impact {i+1}  →", key=f"impact_{i}", use_container_width=True):
                st.session_state.selected_impact = impact
                st.session_state.selected_practice = None
                st.session_state.selected_challenge = None
                st.session_state.page = "practices"
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# 10. STEP 2: PRACTICES SELECTION
# ============================================================
elif st.session_state.page == "practices":

    impact = st.session_state.selected_impact
    if impact is None:
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 2  →  Blue Economy Practice</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">Suitable Blue Economy Practices</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="notice">
            <strong>Active Impact Selection:</strong><br>
            {html.escape(impact)}
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)

    practices = (
        df.loc[df[IMPACT_COL] == impact, PRACTICE_COL]
        .drop_duplicates()
        .tolist()
    )
    st.markdown(f"**{len(practices)} practices linked to this impact**")
    st.write("")

    for i, practice in enumerate(practices):
        st.markdown(f"""
            <div class="info-card">
                <span class="number-circle">{i+1}</span>
                <strong>{html.escape(practice)}</strong>
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"Select Practice {i+1}  →", key=f"practice_{i}", use_container_width=True):
            st.session_state.selected_practice = practice
            st.session_state.selected_challenge = None
            st.session_state.page = "challenges"
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("←  Back to Impacts", use_container_width=True):
            st.session_state.page = "impacts"
            st.rerun()
    with col2:
        if st.button("🔄 Start Again", use_container_width=True):
            reset_app()
            st.rerun()

# ============================================================
# 11. STEP 3: CHALLENGES SELECTION
# ============================================================
elif st.session_state.page == "challenges":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    if impact is None or practice is None:
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 3  →  Implementation Challenge</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">Implementation Challenges</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="notice">
            <strong>Selected Impact:</strong> {html.escape(impact)}<br>
            <strong>Selected Practice:</strong> {html.escape(practice)}
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)

    filtered = df[(df[IMPACT_COL] == impact) & (df[PRACTICE_COL] == practice)]
    challenges = filtered[CHALLENGE_COL].drop_duplicates().tolist()
    st.markdown(f"**{len(challenges)} implementation challenges identified**")
    st.write("")

    for i, challenge in enumerate(challenges):
        st.markdown(f"""
            <div class="info-card">
                <span class="number-circle">{i+1}</span>
                <strong>{html.escape(challenge)}</strong>
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"View Lean Solutions {i+1}  →", key=f"challenge_{i}", use_container_width=True):
            st.session_state.selected_challenge = challenge
            st.session_state.page = "lean"
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("←  Back to Practices", use_container_width=True):
            st.session_state.page = "practices"
            st.rerun()
    with col2:
        if st.button("🔄 Start Again", use_container_width=True):
            reset_app()
            st.rerun()

# ============================================================
# 12. STEP 4: LEAN PRINCIPLES & SUMMARY CHARTS
# ============================================================
elif st.session_state.page == "lean":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    challenge = st.session_state.selected_challenge
    if not all([impact, practice, challenge]):
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 4  →  Prioritized Lean Principles & Summary Analytics</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">Lean Solutions & Analytical Summary</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="notice">
            <strong>Impact:</strong> {html.escape(impact)}<br>
            <strong>Practice:</strong> {html.escape(practice)}<br>
            <strong>Challenge:</strong> {html.escape(challenge)}
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)

    filtered = df[
        (df[IMPACT_COL] == impact) &
        (df[PRACTICE_COL] == practice) &
        (df[CHALLENGE_COL] == challenge)
    ]
    lean_principles = filtered[LEAN_COL].tolist()

    if not lean_principles:
        st.warning("No Lean principles found for this combination.")
    else:
        st.markdown(f"**{len(lean_principles)} Prioritized Lean Principles (Database Order)**")
        st.write("")
        for idx, principle in enumerate(lean_principles, start=1):
            st.markdown(f"""
                <div class="lean-card">
                    <div class="priority-label">Priority Rank {idx} of {len(lean_principles)}</div>
                    <strong style="font-size:1.05rem; color:#1A4A3A;">
                        {html.escape(principle)}
                    </strong>
                </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📈 Summary Analytics & Distribution Overview")
    st.caption("Visual distribution of practices and challenges across the selected impact domain.")

    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        st.markdown("##### Practices Distribution per Impact")
        impact_practice_counts = df[df[IMPACT_COL] == impact][PRACTICE_COL].value_counts()
        st.bar_chart(impact_practice_counts)

    with col_chart2:
        st.markdown("##### Challenges Breakdown for Practice")
        practice_challenge_counts = df[df[PRACTICE_COL] == practice][CHALLENGE_COL].value_counts()
        st.bar_chart(practice_challenge_counts)

    st.markdown("---")
    st.markdown("### 📥 Executive Report Generation")
    st.caption("Download a structured summary report of your current decision pathway.")

    report_content = f"""BLUEPATH EXECUTIVE DECISION REPORT
==================================================
Generated on: Research Session
User / Role: {st.session_state.username if st.session_state.username else 'General Researcher'}

1. MARINE CONSTRUCTION IMPACT:
{impact}

2. SELECTED BLUE ECONOMY PRACTICE:
{practice}

3. IMPLEMENTATION CHALLENGE:
{challenge}

4. PRIORITIZED LEAN PRINCIPLES:
"""
    for idx, p in enumerate(lean_principles, 1):
        report_content += f"   [{idx}] {p}\n"

    report_content += "\n==================================================\nBluePath Platform · Sustainable Marine Construction"

    st.download_button(
        label="📄 Download Executive Summary Report (.txt)",
        data=report_content,
        file_name="BluePath_Decision_Summary.txt",
        mime="text/plain",
        use_container_width=True
    )

    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("←  Back to Challenges", use_container_width=True):
            st.session_state.page = "challenges"
            st.rerun()
    with col2:
        if st.button("⌂  Impacts Hub", use_container_width=True):
            go_home()
            st.rerun()
    with col3:
        if st.button("🔄 Start Again", type="primary", use_container_width=True):
            reset_app()
            st.rerun()

# ------------------------------------------------------------
# FOOTER
# ------------------------------------------------------------
st.markdown("""
    <div class="footer">
        BluePath · Research Decision-Support Platform<br>
        Integrating Blue Economy Practices & Lean Construction Principles for Sustainable Marine Infrastructure
    </div>
""", unsafe_allow_html=True)
