import html
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# BLUEPATH
# BLUE ECONOMY x LEAN CONSTRUCTION
# Interactive Decision-Support Application
# ============================================================


# ------------------------------------------------------------
# 1. PAGE CONFIGURATION
# ------------------------------------------------------------

st.set_page_config(
    page_title="BluePath | Blue Economy and Lean Construction",
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
    (file for file in EXCEL_CANDIDATES if file.exists()),
    EXCEL_CANDIDATES[0]
)

IMPACT_COL = "Impact of Marine Construction"
PRACTICE_COL = "Blue Economy Practice"
CHALLENGE_COL = "Challenges to Implement Blue Economy Practice"
LEAN_COL = "Lean Principles"

REQUIRED_COLUMNS = [
    IMPACT_COL,
    PRACTICE_COL,
    CHALLENGE_COL,
    LEAN_COL
]


# ------------------------------------------------------------
# 3. PREMIUM MODERN DESIGN & SMOOTH MICRO-INTERACTIONS
# ------------------------------------------------------------

st.markdown("""
<style>

/* ==========================================
  GLOBAL LAYOUT & TYPOGRAPHY
========================================== */

.stApp {
    background: #F8FAFC;
    color: #1E293B;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

#MainMenu, footer, header {
    visibility: hidden;
}

.block-container {
    max-width: 1200px;
    padding-top: 2rem;
    padding-bottom: 4rem;
}


/* ==========================================
  SLIDE-BY-SLIDE FLUID ANIMATIONS
========================================== */

@keyframes slideUpFade {
    0% {
        opacity: 0;
        transform: translateY(24px);
    }
    100% {
        opacity: 1;
        transform: translateY(0);
    }
}

.slide-container {
    animation: slideUpFade 0.45s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}


/* ==========================================
  BRAND HEADER
========================================== */

.brand-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1px solid #E2E8F0;
    padding-bottom: 1.25rem;
    margin-bottom: 2rem;
}

.brand-logo {
    font-size: 1.6rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    color: #0F766E;
    text-decoration: none;
}

.brand-sub {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 2px;
    color: #64748B;
    text-transform: uppercase;
}


/* ==========================================
  HERO BANNER
========================================== */

.hero-card {
    background: linear-gradient(135deg, #0F766E 0%, #115E59 100%);
    border-radius: 24px;
    padding: 48px 40px;
    color: white;
    box-shadow: 0 20px 25px -5px rgba(15, 118, 110, 0.15), 0 10px 10px -5px rgba(15, 118, 110, 0.04);
    margin-bottom: 30px;
}

.hero-card h1 {
    font-size: 2.5rem;
    font-weight: 800;
    margin-bottom: 1rem;
    color: white;
    letter-spacing: -1px;
}

.hero-card p {
    font-size: 1.05rem;
    line-height: 1.7;
    color: #CCFBF1;
    max-width: 800px;
    margin-bottom: 0;
}


/* ==========================================
  SELECTION CARDS (INSTAGRAM-LIKE TILE DESIGN)
========================================== */

.choice-tile {
    background: #FFFFFF;
    border: 2px solid #E2E8F0;
    border-radius: 16px;
    padding: 22px 24px;
    margin-bottom: 14px;
    cursor: pointer;
    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.choice-tile:hover {
    border-color: #0F766E;
    transform: translateY(-2px);
    box-shadow: 0 10px 15px -3px rgba(15, 118, 110, 0.08);
}

/* Selected State with distinct highlight color */
.choice-tile-selected {
    background: #F0FDFA;
    border: 2px solid #0D9488;
    border-radius: 16px;
    padding: 22px 24px;
    margin-bottom: 14px;
    box-shadow: 0 4px 12px rgba(13, 148, 136, 0.15);
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.tile-number {
    width: 34px;
    height: 34px;
    background: #CCFBF1;
    color: #0F766E;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    font-size: 0.9rem;
    flex-shrink: 0;
    margin-right: 16px;
}

.tile-number-selected {
    width: 34px;
    height: 34px;
    background: #0D9488;
    color: white;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    font-size: 0.9rem;
    flex-shrink: 0;
    margin-right: 16px;
}

.tile-text {
    font-size: 1.02rem;
    font-weight: 600;
    color: #1E293B;
    flex-grow: 1;
}

.tile-text-selected {
    font-size: 1.02rem;
    font-weight: 700;
    color: #0F766E;
    flex-grow: 1;
}


/* ==========================================
  LEAN SOLUTION PRIORITY CARDS
========================================== */

.lean-tile {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-left: 5px solid #0D9488;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    transition: all 0.2s ease;
}

.lean-tile:hover {
    transform: translateX(4px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}

.priority-badge {
    font-size: 0.7rem;
    font-weight: 800;
    letter-spacing: 1.5px;
    color: #0D9488;
    text-transform: uppercase;
    margin-bottom: 6px;
}


/* ==========================================
  BREADCRUMB HEADER TRAIL
========================================== */

.breadcrumb-box {
    background: #F1F5F9;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 24px;
    font-size: 0.9rem;
    color: #475569;
    border: 1px solid #E2E8F0;
}

.breadcrumb-label {
    font-weight: 700;
    color: #0F766E;
    text-transform: uppercase;
    font-size: 0.72rem;
    letter-spacing: 1px;
    display: block;
    margin-bottom: 4px;
}


/* ==========================================
  CUSTOM BUTTON OVERRIDES
========================================== */

.stButton > button {
    min-height: 44px;
    border-radius: 10px;
    border: 1px solid #CBD5E1;
    background: #FFFFFF;
    color: #0F766E;
    font-weight: 700;
    font-size: 0.92rem;
    transition: all 0.2s ease;
    box-shadow: 0 1px 2px rgba(0,0,0,0.02);
}

.stButton > button:hover {
    background: #F8FAFC;
    border-color: #0D9488;
    color: #115E59;
    transform: translateY(-1px);
    box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
}

.stButton > button[kind="primary"] {
    background: #0D9488;
    color: white;
    border: none;
    box-shadow: 0 4px 10px rgba(13, 148, 136, 0.25);
}

.stButton > button[kind="primary"]:hover {
    background: #0F766E;
    box-shadow: 0 6px 15px rgba(15, 118, 110, 0.35);
}

</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------
# 4. LOAD & VALIDATE DATABASE
# ------------------------------------------------------------

@st.cache_data
def load_database(file_path):
    if not Path(file_path).exists():
        raise FileNotFoundError("portal.xlsx database file could not be found.")

    data = pd.read_excel(file_path, engine="openpyxl")
    data.columns = data.columns.astype(str).str.strip()

    missing = [c for c in REQUIRED_COLUMNS if c not in data.columns]
    if missing:
        raise ValueError(f"Missing columns in Excel: {missing}")

    data = data[REQUIRED_COLUMNS].copy()
    for col in REQUIRED_COLUMNS:
        data[col] = data[col].fillna("").astype(str).str.strip()

    data = data.replace("", pd.NA).dropna(subset=REQUIRED_COLUMNS)
    return data.reset_index(drop=True)


try:
    df = load_database(str(EXCEL_FILE))
except Exception as e:
    st.error(f"Database Error: {e}")
    st.stop()


# ------------------------------------------------------------
# 5. SESSION STATE MANAGEMENT
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
# 6. HEADER BRANDING & NAVIGATION BAR
# ------------------------------------------------------------

st.markdown("""
<div class="brand-header">
    <div>
        <div class="brand-logo">🌊 BluePath</div>
        <div class="brand-sub">Blue Economy × Lean Construction</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ------------------------------------------------------------
# 7. WELCOME SCREEN
# ------------------------------------------------------------

if st.session_state.page == "welcome":
    st.markdown("""
    <div class="slide-container">
        <div class="hero-card">
            <h1>Intelligence for Sustainable Marine Construction</h1>
            <p>
                Explore how integrated Blue Economy practices mitigate marine construction impacts,
                powered by Lean construction principles to streamline real-world project execution.
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### Get Started")
    username = st.text_input("Your Name", placeholder="Enter your name or professional title", max_chars=60)

    if st.button("Start Decision Journey →", type="primary", use_container_width=True):
        if not username.strip():
            st.warning("Please enter your name to continue.")
        else:
            st.session_state.username = username.strip()
            st.session_state.page = "impacts"
            st.rerun()


# ------------------------------------------------------------
# 8. STEP 1: IMPACT SELECTION
# ------------------------------------------------------------

elif st.session_state.page == "impacts":
    name = html.escape(st.session_state.username)

    st.markdown(f"""
    <div class="slide-container">
        <h2 style="color: #0F766E; font-weight: 800; margin-bottom: 4px;">Welcome, {name}</h2>
        <p style="color: #64748B; margin-bottom: 24px;">Step 1 of 4: Select a marine construction impact to manage.</p>
    </div>
    """, unsafe_allow_html=True)

    impacts = df[IMPACT_COL].drop_duplicates().tolist()

    for i, impact in enumerate(impacts):
        is_selected = (st.session_state.selected_impact == impact)
        
        tile_class = "choice-tile-selected" if is_selected else "choice-tile"
        num_class = "tile-number-selected" if is_selected else "tile-number"
        txt_class = "tile-text-selected" if is_selected else "tile-text"

        st.markdown(f"""
        <div class="{tile_class}">
            <div style="display: flex; align-items: center; width: 100%;">
                <div class="{num_class}">{i+1}</div>
                <div class="{txt_class}">{html.escape(impact)}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"Select Impact {i+1} →", key=f"imp_{i}", use_container_width=True):
            st.session_state.selected_impact = impact
            st.session_state.selected_practice = None
            st.session_state.selected_challenge = None
            st.session_state.page = "practices"
            st.rerun()


# ------------------------------------------------------------
# 9. STEP 2: PRACTICE SELECTION
# ------------------------------------------------------------

elif st.session_state.page == "practices":
    impact = st.session_state.selected_impact
    if not impact:
        go_home()
        st.rerun()

    st.markdown(f"""
    <div class="slide-container">
        <div class="breadcrumb-box">
            <span class="breadcrumb-label">Selected Marine Impact</span>
            <strong>{html.escape(impact)}</strong>
        </div>
        <h3 style="color: #0F766E; font-weight: 800; margin-bottom: 4px;">Step 2: Suitable Blue Economy Practices</h3>
        <p style="color: #64748B; margin-bottom: 24px;">Choose a practice designed to address this specific impact.</p>
    </div>
    """, unsafe_allow_html=True)

    practices = df.loc[df[IMPACT_COL] == impact, PRACTICE_COL].drop_duplicates().tolist()

    for i, practice in enumerate(practices):
        is_selected = (st.session_state.selected_practice == practice)
        tile_class = "choice-tile-selected" if is_selected else "choice-tile"
        num_class = "tile-number-selected" if is_selected else "tile-number"
        txt_class = "tile-text-selected" if is_selected else "tile-text"

        st.markdown(f"""
        <div class="{tile_class}">
            <div style="display: flex; align-items: center; width: 100%;">
                <div class="{num_class}">{i+1}</div>
                <div class="{txt_class}">{html.escape(practice)}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"Select Practice {i+1} →", key=f"prac_{i}", use_container_width=True):
            st.session_state.selected_practice = practice
            st.session_state.selected_challenge = None
            st.session_state.page = "challenges"
            st.rerun()

    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
    if st.button("← Back to Impacts", use_container_width=True):
        st.session_state.page = "impacts"
        st.rerun()


# ------------------------------------------------------------
# 10. STEP 3: CHALLENGE SELECTION
# ------------------------------------------------------------

elif st.session_state.page == "challenges":
    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    if not impact or not practice:
        go_home()
        st.rerun()

    st.markdown(f"""
    <div class="slide-container">
        <div class="breadcrumb-box">
            <span class="breadcrumb-label">Active Pathway</span>
            <div><strong>Impact:</strong> {html.escape(impact)}</div>
            <div style="margin-top: 4px;"><strong>Practice:</strong> {html.escape(practice)}</div>
        </div>
        <h3 style="color: #0F766E; font-weight: 800; margin-bottom: 4px;">Step 3: Implementation Challenges</h3>
        <p style="color: #64748B; margin-bottom: 24px;">Select an implementation challenge to reveal prioritized Lean solutions.</p>
    </div>
    """, unsafe_allow_html=True)

    filtered = df[(df[IMPACT_COL] == impact) & (df[PRACTICE_COL] == practice)]
    challenges = filtered[CHALLENGE_COL].drop_duplicates().tolist()

    for i, challenge in enumerate(challenges):
        is_selected = (st.session_state.selected_challenge == challenge)
        tile_class = "choice-tile-selected" if is_selected else "choice-tile"
        num_class = "tile-number-selected" if is_selected else "tile-number"
        txt_class = "tile-text-selected" if is_selected else "tile-text"

        st.markdown(f"""
        <div class="{tile_class}">
            <div style="display: flex; align-items: center; width: 100%;">
                <div class="{num_class}">{i+1}</div>
                <div class="{txt_class}">{html.escape(challenge)}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button(f"View Lean Solutions {i+1} →", key=f"chal_{i}", use_container_width=True):
            st.session_state.selected_challenge = challenge
            st.session_state.page = "lean"
            st.rerun()

    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
    if st.button("← Back to Practices", use_container_width=True):
        st.session_state.page = "practices"
        st.rerun()


# ------------------------------------------------------------
# 11. STEP 4: LEAN PRINCIPLES RESULT VIEW
# ------------------------------------------------------------

elif st.session_state.page == "lean":
    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    challenge = st.session_state.selected_challenge

    if not all([impact, practice, challenge]):
        go_home()
        st.rerun()

    st.markdown(f"""
    <div class="slide-container">
        <div class="breadcrumb-box">
            <span class="breadcrumb-label">Decision Summary</span>
            <div><strong>Impact:</strong> {html.escape(impact)}</div>
            <div style="margin-top: 4px;"><strong>Practice:</strong> {html.escape(practice)}</div>
            <div style="margin-top: 4px;"><strong>Challenge:</strong> {html.escape(challenge)}</div>
        </div>
        <h3 style="color: #0F766E; font-weight: 800; margin-bottom: 4px;">Step 4: Prioritized Lean Principles</h3>
        <p style="color: #64748B; margin-bottom: 24px;">Exact priority ranking from database records to overcome this challenge.</p>
    </div>
    """, unsafe_allow_html=True)

    lean_rows = df[
        (df[IMPACT_COL] == impact) &
        (df[PRACTICE_COL] == practice) &
        (df[CHALLENGE_COL] == challenge)
    ]
    lean_principles = lean_rows[LEAN_COL].tolist()

    for i, principle in enumerate(lean_principles):
        st.markdown(f"""
        <div class="lean-tile slide-container">
            <div class="priority-badge">Priority Rank {i+1} of {len(lean_principles)}</div>
            <div style="font-size: 1.05rem; font-weight: 700; color: #1E293B;">
                {html.escape(principle)}
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='margin-top: 24px;'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("← Back to Challenges", use_container_width=True):
            st.session_state.page = "challenges"
            st.rerun()
    with col2:
        if st.button("🔄 Start Again", type="primary", use_container_width=True):
            reset_app()
            st.rerun()

# Footer
st.markdown("""
<div style="text-align: center; color: #94A3B8; font-size: 0.8rem; margin-top: 50px; border-top: 1px solid #E2E8F0; padding-top: 20px;">
    BluePath Research Platform · Professional Decision Support Framework
</div>
""", unsafe_allow_html=True)
