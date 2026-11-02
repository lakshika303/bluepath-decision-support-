import html
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================
# BLUEPATH
# Blue Economy × Lean Construction
# Interactive Research Decision-Support Platform
# ============================================================

# ------------------------------------------------------------
# 1. PAGE CONFIGURATION
# ------------------------------------------------------------
st.set_page_config(
    page_title="BluePath | Blue Economy × Lean Construction",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ------------------------------------------------------------
# 2. APPLICATION CONFIGURATION
# ------------------------------------------------------------
BASE_DIR = Path(__file__).parent

EXCEL_CANDIDATES = [
    BASE_DIR / "portal.xlsx",
    BASE_DIR / "Portal.xlsx"
]

EXCEL_FILE = next(
    (f for f in EXCEL_CANDIDATES if f.exists()),
    EXCEL_CANDIDATES[0]
)

IMPACT_COL = "Impact of Marine Construction"
PRACTICE_COL = "Blue Economy Practice"
CHALLENGE_COL = "Challenges to Implement Blue Economy Practice"
LEAN_COL = "Lean Principles"

REQUIRED_COLUMNS = [IMPACT_COL, PRACTICE_COL, CHALLENGE_COL, LEAN_COL]
EXPECTED_LEAN_COUNT = 5

# ------------------------------------------------------------
# 3. CUSTOM VISUAL DESIGN
# ------------------------------------------------------------
st.markdown("""
<style>
/* ---------- Base ---------- */
.stApp {
    background: linear-gradient(145deg, #F4F9F9 0%, #F7FBFC 50%, #F2F8F5 100%);
    color: #1F3A44;
    font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
}

#MainMenu, footer, header { visibility: hidden; }

.block-container {
    max-width: 1140px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}

/* ---------- Branding ---------- */
.brand {
    font-size: 1.85rem;
    font-weight: 800;
    letter-spacing: -0.6px;
    color: #0A6E74;
    margin-bottom: 2px;
}
.brand-subtitle {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 2.8px;
    color: #5A7A84;
    text-transform: uppercase;
}

/* ---------- Hero ---------- */
.hero {
    background: linear-gradient(125deg, #0A6A72 0%, #0D8A8F 55%, #2BA88A 100%);
    border-radius: 18px;
    padding: 42px 40px;
    margin: 18px 0 28px 0;
    color: white;
    box-shadow: 0 12px 28px rgba(10, 110, 116, 0.18);
}
.hero h1 {
    font-size: 2.35rem;
    font-weight: 800;
    line-height: 1.25;
    margin: 8px 0 14px 0;
    color: white;
}
.hero p {
    font-size: 1.02rem;
    line-height: 1.75;
    color: #E8F7F7;
    max-width: 780px;
    margin: 0;
}
.hero-label {
    font-size: 0.72rem;
    letter-spacing: 2.5px;
    font-weight: 700;
    color: #C8EFEF;
    text-transform: uppercase;
}

/* ---------- Section headings ---------- */
.section-title {
    font-size: 1.55rem;
    font-weight: 750;
    color: #0D5F66;
    margin: 6px 0 6px 0;
}
.section-description {
    font-size: 0.98rem;
    line-height: 1.65;
    color: #4A6A74;
    margin-bottom: 18px;
}
.step-label {
    font-size: 0.72rem;
    letter-spacing: 2.2px;
    font-weight: 750;
    color: #0D8A8A;
    text-transform: uppercase;
    margin-bottom: 4px;
}

/* ---------- Cards ---------- */
.info-card {
    background: rgba(255,255,255,0.97);
    border: 1px solid #D5E6E8;
    border-radius: 14px;
    padding: 18px 18px 16px 18px;
    margin-bottom: 14px;
    box-shadow: 0 3px 12px rgba(30, 80, 90, 0.05);
    transition: all 0.22s ease;
}
.info-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 22px rgba(30, 80, 90, 0.09);
    border-color: #B8D8DA;
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
    font-size: 0.88rem;
    font-weight: 750;
    margin-right: 12px;
    flex-shrink: 0;
}

/* ---------- Lean principle cards ---------- */
.lean-card {
    background: linear-gradient(120deg, #EDF7F1 0%, #F6FBF7 100%);
    border: 1px solid #C5E0CF;
    border-radius: 13px;
    padding: 18px 18px 16px 18px;
    margin-bottom: 12px;
    transition: all 0.22s ease;
}
.lean-card:hover {
    transform: translateX(4px);
    box-shadow: 0 6px 16px rgba(40, 120, 80, 0.08);
}
.priority-label {
    font-size: 0.70rem;
    font-weight: 750;
    letter-spacing: 1.1px;
    color: #1E7A4A;
    text-transform: uppercase;
    margin-bottom: 4px;
}

/* ---------- Notice / selection summary ---------- */
.notice {
    background: #EAF5F5;
    border-left: 4px solid #0D8A8A;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 16px 0 20px 0;
    color: #2A4A52;
    line-height: 1.65;
    font-size: 0.95rem;
}

/* ---------- Buttons ---------- */
.stButton > button {
    min-height: 42px;
    border-radius: 10px;
    border: 1px solid #C5DCD9;
    background: #FFFFFF;
    color: #0A6E74;
    font-weight: 650;
    transition: all 0.2s ease;
}
.stButton > button:hover {
    background: #E8F4F2;
    border-color: #0D8A8A;
    color: #064E52;
    transform: translateY(-1px);
}
.stButton > button[kind="primary"] {
    background: linear-gradient(100deg, #0A6E74, #1A9A90);
    color: white;
    border: none;
}
.stButton > button[kind="primary"]:hover {
    background: #085A5F;
    color: white;
}

/* ---------- Progress ---------- */
.stProgress > div > div > div {
    background: linear-gradient(90deg, #0A6E74, #3AA88A);
}

/* ---------- Footer ---------- */
.footer {
    text-align: center;
    color: #6A848C;
    font-size: 0.80rem;
    border-top: 1px solid #D8E8EA;
    padding-top: 20px;
    margin-top: 48px;
}

/* ---------- Responsive ---------- */
@media (max-width: 700px) {
    .hero { padding: 28px 22px; }
    .hero h1 { font-size: 1.85rem; }
    .section-title { font-size: 1.35rem; }
    .brand { font-size: 1.55rem; }
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# 4. LOAD AND VALIDATE EXCEL DATABASE
# ------------------------------------------------------------
@st.cache_data
def load_database(file_path, file_modified):
    if not Path(file_path).exists():
        raise FileNotFoundError(
            "Excel database not found. Please place portal.xlsx "
            "(or Portal.xlsx) in the same folder as app.py."
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

    # CRITICAL: preserve original Excel row order – never sort Lean principles
    return data.reset_index(drop=True)

try:
    df = load_database(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0
    )
except Exception as e:
    st.error(f"Database loading error: {e}")
    st.info("Ensure portal.xlsx is present and contains the four required columns.")
    st.stop()

# ------------------------------------------------------------
# 5. SESSION STATE
# ------------------------------------------------------------
DEFAULTS = {
    "page": "welcome",
    "username": "",
    "selected_impact": None,
    "selected_practice": None,
    "selected_challenge": None
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value

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
# 6. HEADER
# ------------------------------------------------------------
col_logo, col_nav = st.columns([4, 1])

with col_logo:
    st.markdown("""
        <div class="brand">BluePath</div>
        <div class="brand-subtitle">Blue Economy × Lean Construction</div>
    """, unsafe_allow_html=True)

with col_nav:
    if st.session_state.page != "welcome":
        if st.button("Home", use_container_width=True):
            go_home()
            st.rerun()

# ------------------------------------------------------------
# 7. PROGRESS INDICATOR
# ------------------------------------------------------------
PAGES = {
    "welcome": 0,
    "impacts": 1,
    "practices": 2,
    "challenges": 3,
    "lean": 4
}

if st.session_state.page != "welcome":
    step = PAGES[st.session_state.page]
    st.progress(step / 4)
    st.caption(f"Decision pathway · Step {step} of 4")

# ============================================================
# 8. WELCOME PAGE
# ============================================================
if st.session_state.page == "welcome":

    st.markdown("""
        <div class="hero">
            <div class="hero-label">Research Decision-Support Platform</div>
            <h1>BluePath</h1>
            <p>
                An interactive platform that connects marine construction impacts
                with relevant Blue Economy practices and prioritised Lean principles
                to support sustainable implementation.
            </p>
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="section-title">Begin</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-description">'
        'Enter your name (optional) and start the structured decision pathway.'
        '</div>',
        unsafe_allow_html=True
    )

    username = st.text_input(
        "Your name",
        placeholder="Optional – used only for personalisation on this session",
        max_chars=60
    )

    st.markdown("#### Decision pathway")
    steps = [
        ("01", "Select marine construction impact",
         "Choose the impact you wish to address."),
        ("02", "Select Blue Economy practice",
         "View practices linked to the selected impact."),
        ("03", "Select implementation challenge",
         "Identify challenges associated with the practice."),
        ("04", "Review prioritised Lean principles",
         "See the five Lean principles in the exact priority order recorded in the research database.")
    ]

    for num, title, desc in steps:
        st.markdown(f"""
            <div class="info-card">
                <span class="number-circle">{num}</span>
                <strong>{title}</strong>
                <p style="color:#4A6A74; margin:8px 0 0 46px; line-height:1.55; font-size:0.93rem;">
                    {desc}
                </p>
            </div>
        """, unsafe_allow_html=True)

    st.write("")
    if st.button("Start decision pathway", type="primary", use_container_width=True):
        st.session_state.username = username.strip()
        st.session_state.page = "impacts"
        st.rerun()

# ============================================================
# 9. IMPACT SELECTION
# ============================================================
elif st.session_state.page == "impacts":

    name = html.escape(st.session_state.username) if st.session_state.username else "Researcher"

    st.markdown(f'<div class="section-title">Hello, {name}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-description">'
        'Select the marine construction impact you wish to manage.'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown("""
        <div class="notice">
            <strong>Step 1 · Marine Construction Impact</strong><br>
            Impacts appear in the order recorded in the research database.
        </div>
    """, unsafe_allow_html=True)

    impacts = df[IMPACT_COL].drop_duplicates().tolist()
    st.markdown(f"**{len(impacts)} impacts available**")

    cols = st.columns(2)
    for i, impact in enumerate(impacts):
        with cols[i % 2]:
            st.markdown(f"""
                <div class="info-card">
                    <span class="number-circle">{i+1}</span>
                    <strong>{html.escape(impact)}</strong>
                </div>
            """, unsafe_allow_html=True)

            if st.button(f"Select impact {i+1}", key=f"impact_{i}", use_container_width=True):
                st.session_state.selected_impact = impact
                st.session_state.selected_practice = None
                st.session_state.selected_challenge = None
                st.session_state.page = "practices"
                st.rerun()

# ============================================================
# 10. PRACTICE SELECTION
# ============================================================
elif st.session_state.page == "practices":

    impact = st.session_state.selected_impact
    if impact is None:
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 2 · Blue Economy Practice</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">Suitable Blue Economy practices</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="notice">
            <strong>Selected impact</strong><br>
            {html.escape(impact)}
        </div>
    """, unsafe_allow_html=True)

    practices = (
        df.loc[df[IMPACT_COL] == impact, PRACTICE_COL]
        .drop_duplicates()
        .tolist()
    )

    st.markdown(f"**{len(practices)} practices linked to this impact**")

    for i, practice in enumerate(practices):
        st.markdown(f"""
            <div class="info-card">
                <span class="number-circle">{i+1}</span>
                <strong>{html.escape(practice)}</strong>
            </div>
        """, unsafe_allow_html=True)

        if st.button("Select this practice", key=f"practice_{i}", use_container_width=True):
            st.session_state.selected_practice = practice
            st.session_state.selected_challenge = None
            st.session_state.page = "challenges"
            st.rerun()

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to impacts", use_container_width=True):
            st.session_state.page = "impacts"
            st.rerun()
    with col2:
        if st.button("Start again", use_container_width=True):
            reset_app()
            st.rerun()

# ============================================================
# 11. CHALLENGE SELECTION
# ============================================================
elif st.session_state.page == "challenges":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice

    if impact is None or practice is None:
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 3 · Implementation Challenge</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">Challenges to implementing the selected practice</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="notice">
            <strong>Selected impact</strong><br>
            {html.escape(impact)}<br><br>
            <strong>Selected Blue Economy practice</strong><br>
            {html.escape(practice)}
        </div>
    """, unsafe_allow_html=True)

    filtered = df[
        (df[IMPACT_COL] == impact) &
        (df[PRACTICE_COL] == practice)
    ]

    challenges = filtered[CHALLENGE_COL].drop_duplicates().tolist()

    st.markdown(f"**{len(challenges)} challenges identified**")

    for i, challenge in enumerate(challenges):
        st.markdown(f"""
            <div class="info-card">
                <span class="number-circle">{i+1}</span>
                <strong>{html.escape(challenge)}</strong>
            </div>
        """, unsafe_allow_html=True)

        if st.button("View Lean principles", key=f"challenge_{i}", use_container_width=True):
            st.session_state.selected_challenge = challenge
            st.session_state.page = "lean"
            st.rerun()

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to practices", use_container_width=True):
            st.session_state.page = "practices"
            st.rerun()
    with col2:
        if st.button("Start again", use_container_width=True):
            reset_app()
            st.rerun()

# ============================================================
# 12. LEAN PRINCIPLES (PRIORITY ORDER FROM EXCEL)
# ============================================================
elif st.session_state.page == "lean":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    challenge = st.session_state.selected_challenge

    if not all([impact, practice, challenge]):
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 4 · Prioritised Lean Principles</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-title">Lean principles to address the selected challenge</div>',
        unsafe_allow_html=True
    )

    st.markdown(f"""
        <div class="notice">
            <strong>Selected impact</strong><br>
            {html.escape(impact)}<br><br>
            <strong>Selected Blue Economy practice</strong><br>
            {html.escape(practice)}<br><br>
            <strong>Selected implementation challenge</strong><br>
            {html.escape(challenge)}
        </div>
    """, unsafe_allow_html=True)

    # Preserve exact order from Excel – do not sort
    filtered = df[
        (df[IMPACT_COL] == impact) &
        (df[PRACTICE_COL] == practice) &
        (df[CHALLENGE_COL] == challenge)
    ]

    lean_principles = filtered[LEAN_COL].tolist()

    if not lean_principles:
        st.warning("No Lean principles found for this combination.")
    else:
        st.markdown(
            f"**{len(lean_principles)} Lean principles** "
            "(shown in the priority order recorded in the research database)"
        )

        for idx, principle in enumerate(lean_principles, start=1):
            st.markdown(f"""
                <div class="lean-card">
                    <div class="priority-label">Priority {idx}</div>
                    <strong style="font-size:1.05rem; color:#1A4A3A;">
                        {html.escape(principle)}
                    </strong>
                </div>
            """, unsafe_allow_html=True)

        if len(lean_principles) != EXPECTED_LEAN_COUNT:
            st.caption(
                f"Note: Expected {EXPECTED_LEAN_COUNT} principles. "
                f"Found {len(lean_principles)} for this combination."
            )

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("← Back to challenges", use_container_width=True):
            st.session_state.page = "challenges"
            st.rerun()
    with col2:
        if st.button("Home", use_container_width=True):
            go_home()
            st.rerun()
    with col3:
        if st.button("Start again", type="primary", use_container_width=True):
            reset_app()
            st.rerun()

# ------------------------------------------------------------
# FOOTER
# ------------------------------------------------------------
st.markdown("""
    <div class="footer">
        BluePath · Research decision-support platform<br>
        Blue Economy practices × Lean principles for marine construction
    </div>
""", unsafe_allow_html=True)
