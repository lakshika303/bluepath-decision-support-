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
    initial_sidebar_state="collapsed"
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
# CUSTOM DESIGN
# ------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

.stApp {
    background: linear-gradient(145deg, #F4F9F9 0%, #F7FBFC 50%, #F2F8F5 100%);
    color: #1F3A44;
    font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
}
#MainMenu, footer, header { visibility: hidden; }
.block-container {
    max-width: 1140px;
    padding-top: 1.2rem;
    padding-bottom: 2.5rem;
}

/* Branding */
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

/* Typography */
.page-title {
    font-size: 1.55rem;
    font-weight: 750;
    color: #0D5F66;
    margin: 4px 0 6px 0;
}
.section-title {
    font-size: 1.25rem;
    font-weight: 700;
    color: #0D5F66;
    margin: 12px 0 6px 0;
}
.section-description {
    font-size: 0.96rem;
    line-height: 1.6;
    color: #4A6A74;
    margin-bottom: 14px;
}
.step-label {
    font-size: 0.72rem;
    letter-spacing: 2.2px;
    font-weight: 750;
    color: #0D8A8A;
    text-transform: uppercase;
    margin-bottom: 4px;
}
.caption-text {
    font-size: 0.82rem;
    color: #6A848C;
}

/* Hero */
.hero {
    background: linear-gradient(125deg, #0A6A72 0%, #0D8A8F 55%, #2BA88A 100%);
    border-radius: 18px;
    padding: 36px 34px;
    margin: 12px 0 22px 0;
    color: white;
    box-shadow: 0 12px 28px rgba(10, 110, 116, 0.18);
}
.hero h1 {
    font-size: 2.15rem;
    font-weight: 800;
    line-height: 1.25;
    margin: 6px 0 12px 0;
    color: white;
}
.hero p {
    font-size: 1.02rem;
    line-height: 1.7;
    color: #E8F7F7;
    max-width: 740px;
    margin: 0;
}
.hero-label {
    font-size: 0.72rem;
    letter-spacing: 2.5px;
    font-weight: 700;
    color: #C8EFEF;
    text-transform: uppercase;
}

/* Welcome highlight box */
.welcome-box {
    background: linear-gradient(135deg, #E8F6F3 0%, #F0F9F7 100%);
    border-radius: 16px;
    padding: 24px 26px;
    margin: 18px 0;
    border: 1px solid #C5E0D8;
}

/* Cards */
.content-panel {
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid #D5E6E8;
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 16px;
    box-shadow: 0 4px 16px rgba(30, 80, 90, 0.05);
}
.info-card {
    background: linear-gradient(135deg, #FFFFFF 0%, #FAFCFC 100%);
    border: 1px solid #D5E6E8;
    border-left: 5px solid #0A6E74;
    border-radius: 12px;
    padding: 14px 16px;
    margin-bottom: 11px;
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
    border-radius: 12px;
    padding: 14px 16px;
    margin-bottom: 10px;
    transition: all 0.25s ease;
}
.pathway-card:hover {
    border-left-color: #E05A5A;
    background: #FFF8F8;
    transform: translateX(4px);
}
.number-circle {
    display: inline-flex;
    justify-content: center;
    align-items: center;
    width: 30px;
    height: 30px;
    border-radius: 50%;
    background: #E4F3F1;
    color: #0A6E74;
    font-size: 0.84rem;
    font-weight: 750;
    margin-right: 11px;
}

/* Lean cards */
.lean-card {
    background: linear-gradient(120deg, #EDF7F1 0%, #F6FBF7 100%);
    border: 1px solid #C5E0CF;
    border-left: 5px solid #1E7A4A;
    border-radius: 12px;
    padding: 14px 16px;
    margin-bottom: 10px;
}
.priority-label {
    font-size: 0.68rem;
    font-weight: 750;
    letter-spacing: 1.1px;
    color: #1E7A4A;
    text-transform: uppercase;
    margin-bottom: 3px;
}

/* Notice */
.notice {
    background: linear-gradient(135deg, #EAF5F5 0%, #E2F0F0 100%);
    border-left: 5px solid #0D8A8A;
    border-radius: 10px;
    padding: 13px 16px;
    margin: 12px 0 16px 0;
    color: #2A4A52;
    line-height: 1.6;
    font-size: 0.94rem;
}

/* Summary box */
.summary-box {
    background: linear-gradient(135deg, #F0F9F7 0%, #E8F5F2 100%);
    border: 1px solid #B8DCD4;
    border-radius: 14px;
    padding: 18px 20px;
    margin: 16px 0;
}
.process-step {
    display: inline-block;
    background: white;
    border: 1px solid #C5DCD9;
    border-radius: 8px;
    padding: 7px 12px;
    margin: 3px 4px 3px 0;
    font-size: 0.86rem;
    font-weight: 600;
    color: #0A6E74;
}
.process-arrow {
    color: #0D8A8A;
    font-weight: 700;
    margin: 0 3px;
}

/* KPI / Metric cards */
.metric-card {
    background: white;
    border: 1px solid #D5E6E8;
    border-radius: 12px;
    padding: 18px 16px;
    text-align: center;
    box-shadow: 0 2px 10px rgba(30, 80, 90, 0.04);
}
.metric-value {
    font-size: 1.7rem;
    font-weight: 800;
    color: #0A6E74;
    margin: 4px 0;
}
.metric-label {
    font-size: 0.78rem;
    color: #5A7A84;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
}

/* Buttons */
.stButton > button {
    min-height: 42px;
    border-radius: 10px;
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
}
.stButton > button[kind="primary"] {
    background: linear-gradient(100deg, #0A6E74, #1A9A90);
    color: white;
    border: none;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(100deg, #C0392B, #E05A5A);
    color: white;
}

/* Progress */
.stProgress > div > div > div {
    background: linear-gradient(90deg, #0A6E74, #3AA88A);
}

/* Footer */
.footer {
    text-align: center;
    color: #6A848C;
    font-size: 0.80rem;
    border-top: 1px solid #D8E8EA;
    padding-top: 16px;
    margin-top: 36px;
}

@media (max-width: 700px) {
    .hero { padding: 24px 18px; }
    .hero h1 { font-size: 1.75rem; }
    .page-title { font-size: 1.35rem; }
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
    "selected_challenge": None,
    "feedback_submitted": False
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
    st.session_state.feedback_submitted = False

def go_home():
    st.session_state.page = "impacts"
    st.session_state.selected_impact = None
    st.session_state.selected_practice = None
    st.session_state.selected_challenge = None

# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------
col_logo, col_nav = st.columns([4, 1])
with col_logo:
    st.markdown("""
        <div class="brand">BluePath</div>
        <div class="brand-subtitle">Blue Economy × Lean Construction</div>
    """, unsafe_allow_html=True)
with col_nav:
    if st.session_state.page != "welcome":
        if st.button("⌂  Home", use_container_width=True):
            go_home()
            st.rerun()

# ------------------------------------------------------------
# PROGRESS
# ------------------------------------------------------------
PAGES = {"welcome": 0, "impacts": 1, "practices": 2, "challenges": 3, "lean": 4}
if st.session_state.page != "welcome":
    step = PAGES[st.session_state.page]
    st.progress(step / 4)
    st.caption(f"Decision pathway  →  Step {step} of 4")

# ============================================================
# 8. WELCOME PAGE
# ============================================================
if st.session_state.page == "welcome":

    st.markdown("""
        <div class="hero">
            <div class="hero-label">Research Decision-Support Platform</div>
            <h1>BluePath</h1>
            <p>
                Connect marine construction impacts with Blue Economy practices
                and prioritised Lean principles to support sustainable implementation.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Friendly welcome box (inspired by the image you liked)
    st.markdown("""
        <div class="welcome-box">
            <div style="font-size:1.25rem; font-weight:700; color:#0D5F66; margin-bottom:8px;">
                We're glad you're here
            </div>
            <div style="font-size:0.97rem; color:#3A5A64; line-height:1.65;">
                BluePath helps researchers and practitioners systematically explore how Blue Economy practices
                can address the impacts of marine construction, identify implementation challenges,
                and discover prioritised Lean principles that support successful delivery.
            </div>
        </div>
    """, unsafe_allow_html=True)

    # KPI cards
    n_impacts = df[IMPACT_COL].nunique()
    n_practices = df[PRACTICE_COL].nunique()
    n_challenges = df[CHALLENGE_COL].nunique()

    st.markdown('<div class="section-title">Database overview</div>', unsafe_allow_html=True)
    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{n_impacts}</div>
                <div class="metric-label">Marine Impacts</div>
            </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{n_practices}</div>
                <div class="metric-label">Blue Economy Practices</div>
            </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{n_challenges}</div>
                <div class="metric-label">Implementation Challenges</div>
            </div>
        """, unsafe_allow_html=True)

    # Process steps
    st.markdown('<div class="section-title">How the decision pathway works</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-value">01</div>
                <div class="metric-label">Impact</div>
                <div class="caption-text" style="margin-top:6px;">Select marine construction impact</div>
            </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-value">02</div>
                <div class="metric-label">Practice</div>
                <div class="caption-text" style="margin-top:6px;">Choose linked Blue Economy practice</div>
            </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-value">03</div>
                <div class="metric-label">Challenge</div>
                <div class="caption-text" style="margin-top:6px;">Identify implementation barriers</div>
            </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown("""
            <div class="metric-card">
                <div class="metric-value">04</div>
                <div class="metric-label">Lean</div>
                <div class="caption-text" style="margin-top:6px;">Review prioritised principles</div>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("""
        <div style="text-align:center; margin:16px 0 22px 0; color:#0D8A8A; font-weight:600;">
            Impact  →  Practice  →  Challenge  →  Prioritised Lean Principles
        </div>
    """, unsafe_allow_html=True)

    username = st.text_input(
        "Your name (optional)",
        placeholder="Used only for personalisation in this session",
        max_chars=60
    )

    if st.button("Start decision pathway  →", type="primary", use_container_width=True):
        st.session_state.username = username.strip()
        st.session_state.page = "impacts"
        st.rerun()

# ============================================================
# 9. IMPACTS
# ============================================================
elif st.session_state.page == "impacts":

    name = html.escape(st.session_state.username) if st.session_state.username else "Researcher"
    st.markdown(f'<div class="page-title">Hello, {name}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-description">'
        'Select the marine construction impact you wish to manage. '
        'Impacts appear in the order recorded in the research database.'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown("""
        <div class="notice">
            <strong>Step 1  →  Marine Construction Impact</strong><br>
            Choose one impact to explore the associated Blue Economy practices.
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)
    impacts = df[IMPACT_COL].drop_duplicates().tolist()
    st.markdown(f"**{len(impacts)} impacts available**")
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
            if st.button(f"Select  →", key=f"impact_{i}", use_container_width=True):
                st.session_state.selected_impact = impact
                st.session_state.selected_practice = None
                st.session_state.selected_challenge = None
                st.session_state.page = "practices"
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# 10. PRACTICES
# ============================================================
elif st.session_state.page == "practices":

    impact = st.session_state.selected_impact
    if impact is None:
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 2  →  Blue Economy Practice</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-title">Suitable Blue Economy practices</div>', unsafe_allow_html=True)

    st.markdown(f"""
        <div class="notice">
            <strong>Selected Impact</strong><br>
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
        if st.button(f"Select practice  →", key=f"practice_{i}", use_container_width=True):
            st.session_state.selected_practice = practice
            st.session_state.selected_challenge = None
            st.session_state.page = "challenges"
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("←  Back to impacts", use_container_width=True):
            st.session_state.page = "impacts"
            st.rerun()
    with col2:
        if st.button("Start again", use_container_width=True):
            reset_app()
            st.rerun()

# ============================================================
# 11. CHALLENGES
# ============================================================
elif st.session_state.page == "challenges":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    if impact is None or practice is None:
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 3  →  Implementation Challenge</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-title">Challenges to implementing the selected practice</div>', unsafe_allow_html=True)

    st.markdown(f"""
        <div class="notice">
            <strong>Selected Impact</strong><br>
            {html.escape(impact)}<br><br>
            <strong>Selected Blue Economy Practice</strong><br>
            {html.escape(practice)}
        </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)
    filtered = df[(df[IMPACT_COL] == impact) & (df[PRACTICE_COL] == practice)]
    challenges = filtered[CHALLENGE_COL].drop_duplicates().tolist()
    st.markdown(f"**{len(challenges)} challenges identified**")
    st.write("")

    for i, challenge in enumerate(challenges):
        st.markdown(f"""
            <div class="info-card">
                <span class="number-circle">{i+1}</span>
                <strong>{html.escape(challenge)}</strong>
            </div>
        """, unsafe_allow_html=True)
        if st.button(f"View Lean principles  →", key=f"challenge_{i}", use_container_width=True):
            st.session_state.selected_challenge = challenge
            st.session_state.page = "lean"
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("←  Back to practices", use_container_width=True):
            st.session_state.page = "practices"
            st.rerun()
    with col2:
        if st.button("Start again", use_container_width=True):
            reset_app()
            st.rerun()

# ============================================================
# 12. LEAN PRINCIPLES + VISUAL SUMMARY + FEEDBACK
# ============================================================
elif st.session_state.page == "lean":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    challenge = st.session_state.selected_challenge
    if not all([impact, practice, challenge]):
        go_home()
        st.rerun()

    st.markdown('<div class="step-label">Step 4  →  Prioritised Lean Principles</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-title">Lean principles to address the selected challenge</div>', unsafe_allow_html=True)

    st.markdown(f"""
        <div class="notice">
            <strong>Selected Impact</strong><br>
            {html.escape(impact)}<br><br>
            <strong>Selected Blue Economy Practice</strong><br>
            {html.escape(practice)}<br><br>
            <strong>Selected Implementation Challenge</strong><br>
            {html.escape(challenge)}
        </div>
    """, unsafe_allow_html=True)

    # Lean principles (Excel order preserved)
    filtered = df[
        (df[IMPACT_COL] == impact) &
        (df[PRACTICE_COL] == practice) &
        (df[CHALLENGE_COL] == challenge)
    ]
    lean_principles = filtered[LEAN_COL].tolist()

    st.markdown('<div class="content-panel">', unsafe_allow_html=True)
    if not lean_principles:
        st.warning("No Lean principles found for this combination.")
    else:
        st.markdown(
            f"**{len(lean_principles)} Lean principles** "
            "(shown in the exact priority order from the research database)"
        )
        st.write("")
        for idx, principle in enumerate(lean_principles, start=1):
            st.markdown(f"""
                <div class="lean-card">
                    <div class="priority-label">Priority Rank {idx} of {len(lean_principles)}</div>
                    <strong style="font-size:1.03rem; color:#1A4A3A;">
                        {html.escape(principle)}
                    </strong>
                </div>
            """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ---------- VISUAL SUMMARY ----------
    st.markdown('<div class="section-title">Decision path summary</div>', unsafe_allow_html=True)

    # Broader context for the whole Impact
    impact_df = df[df[IMPACT_COL] == impact]
    total_practices = impact_df[PRACTICE_COL].nunique()
    total_challenges = impact_df[CHALLENGE_COL].nunique()
    total_lean = len(impact_df)

    st.markdown(f"""
        <div class="summary-box">
            <div style="font-size:0.75rem; font-weight:700; letter-spacing:1.5px; color:#0D8A8A; text-transform:uppercase; margin-bottom:10px;">
                Selected path + Impact context
            </div>
            <div style="line-height:2.05;">
                <span class="process-step">① Impact</span>
                <span class="process-arrow">→</span>
                <span class="process-step">{html.escape(impact[:40] + ('…' if len(impact) > 40 else ''))}</span>
                <br>
                <span class="process-step">② Practice</span>
                <span class="process-arrow">→</span>
                <span class="process-step">{html.escape(practice[:40] + ('…' if len(practice) > 40 else ''))}</span>
                <br>
                <span class="process-step">③ Challenge</span>
                <span class="process-arrow">→</span>
                <span class="process-step">{html.escape(challenge[:40] + ('…' if len(challenge) > 40 else ''))}</span>
                <br>
                <span class="process-step">④ Lean priorities</span>
                <span class="process-arrow">→</span>
                <span class="process-step">{len(lean_principles)} principles (Excel order)</span>
            </div>
            <div style="margin-top:14px; font-size:0.92rem; color:#2A4A52;">
                <strong>Full Impact context:</strong> This impact is linked to 
                <strong>{total_practices}</strong> Blue Economy practices and 
                <strong>{total_challenges}</strong> unique challenges in the research database.
            </div>
        </div>
    """, unsafe_allow_html=True)

    # Small KPI row for the selected path
    s1, s2, s3 = st.columns(3)
    with s1:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{total_practices}</div>
                <div class="metric-label">Practices under this Impact</div>
            </div>
        """, unsafe_allow_html=True)
    with s2:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{total_challenges}</div>
                <div class="metric-label">Challenges under this Impact</div>
            </div>
        """, unsafe_allow_html=True)
    with s3:
        st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{len(lean_principles)}</div>
                <div class="metric-label">Lean principles (this path)</div>
            </div>
        """, unsafe_allow_html=True)

    # ---------- FEEDBACK ----------
    st.markdown('<div class="section-title">Your feedback</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-description">'
        'How useful was this decision pathway for your research or practice?'
        '</div>',
        unsafe_allow_html=True
    )

    if not st.session_state.feedback_submitted:
        rating = st.slider("Rating (1 = Poor, 5 = Excellent)", 1, 5, 4)
        comment = st.text_area("Optional comment", placeholder="Any suggestions or observations...")
        if st.button("Submit feedback", type="primary"):
            st.session_state.feedback_submitted = True
            st.success("Thank you for your feedback!")
            st.rerun()
    else:
        st.success("Thank you! Your feedback has been recorded for this session.")

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("←  Back to challenges", use_container_width=True):
            st.session_state.page = "challenges"
            st.rerun()
    with col2:
        if st.button("⌂  Home", use_container_width=True):
            go_home()
            st.rerun()
    with col3:
        if st.button("Start again  →", type="primary", use_container_width=True):
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
