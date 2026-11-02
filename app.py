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

# Support both possible Excel filenames
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

EXPECTED_LEAN_COUNT = 5


# ------------------------------------------------------------
# 3. CUSTOM VISUAL DESIGN
# ------------------------------------------------------------

st.markdown("""
<style>

/* ==========================================
  GENERAL APPEARANCE
========================================== */

.stApp {
    background:
        linear-gradient(
            140deg,
            #F0FAFA 0%,
            #F6FAFF 48%,
            #EFF8F3 100%
        );

    color: #294B5C;
    font-family: 'Inter', 'Segoe UI', sans-serif;
}

#MainMenu,
footer,
header {
    visibility: hidden;
}

.block-container {
    max-width: 1250px;
    padding-top: 1.5rem;
    padding-bottom: 3rem;
}


/* ==========================================
  BRANDING
========================================== */

.brand {
    font-size: 2rem;
    font-weight: 850;
    letter-spacing: -1px;
    color: #087F83;
}

.brand-subtitle {
    font-size: 0.72rem;
    font-weight: 750;
    letter-spacing: 3px;
    color: #68858E;
}


/* ==========================================
  HERO SECTION
========================================== */

.hero {
    background:
        linear-gradient(
            120deg,
            #076D78 0%,
            #0B9397 55%,
            #45B89B 100%
        );

    border-radius: 25px;
    padding: 55px 45px;
    margin-top: 25px;
    margin-bottom: 30px;

    box-shadow:
        0 15px 35px rgba(8, 127, 131, 0.15);

    color: white;
    animation: fadeIn 0.8s ease;
}

.hero h1 {
    font-size: 2.8rem;
    font-weight: 850;
    line-height: 1.2;
    color: white;
}

.hero p {
    font-size: 1.04rem;
    line-height: 1.9;
    color: #EDFFFF;
    max-width: 820px;
}

.hero-label {
    font-size: 0.78rem;
    letter-spacing: 3px;
    font-weight: 800;
    color: #D4FFFF;
}


/* ==========================================
  SECTION HEADINGS
========================================== */

.section-title {
    font-size: 1.85rem;
    font-weight: 850;
    color: #126F78;
    margin-bottom: 8px;
}

.section-description {
    font-size: 1rem;
    line-height: 1.75;
    color: #647F8C;
}

.step-label {
    font-size: 0.76rem;
    letter-spacing: 2.5px;
    font-weight: 850;
    color: #159B91;
    text-transform: uppercase;
}


/* ==========================================
  INFORMATION CARDS
========================================== */

.info-card {
    background: rgba(255,255,255,0.96);
    border: 1px solid #DCEBEC;
    border-radius: 17px;

    padding: 21px;
    margin-bottom: 12px;

    box-shadow:
        0 5px 17px rgba(34,91,104,0.05);

    transition: all 0.25s ease;
}

.info-card:hover {
    transform: translateY(-4px);

    box-shadow:
        0 11px 27px rgba(34,91,104,0.11);
}


/* ==========================================
  NUMBER CIRCLES
========================================== */

.number-circle {
    display: inline-flex;
    justify-content: center;
    align-items: center;

    width: 39px;
    height: 39px;

    border-radius: 50%;

    background: #E0F3EF;
    color: #087F83;

    font-size: 0.95rem;
    font-weight: 850;

    margin-right: 12px;
    flex-shrink: 0;
}


/* ==========================================
  INFORMATION BANNERS
========================================== */

.notice {
    background: #E8F5F5;
    border-left: 5px solid #159D9B;

    border-radius: 9px;
    padding: 17px 20px;

    margin: 18px 0;

    color: #365E68;
    line-height: 1.75;
}


/* ==========================================
  LEAN PRINCIPLE CARDS
========================================== */

.lean-card {
    background:
        linear-gradient(
            120deg,
            #E8F6ED,
            #F5FBF6
        );

    border: 1px solid #CBE7D2;
    border-radius: 16px;

    padding: 20px;
    margin-bottom: 13px;

    transition: all 0.25s ease;
}

.lean-card:hover {
    transform: translateX(5px);

    box-shadow:
        0 8px 20px rgba(48,127,83,0.09);
}

.priority-label {
    font-size: 0.73rem;
    font-weight: 850;
    letter-spacing: 1.2px;
    color: #24834E;
    text-transform: uppercase;
}


/* ==========================================
  NAVIGATION BUTTONS
========================================== */

.stButton > button {
    min-height: 45px;

    border-radius: 11px;
    border: 1px solid #C9E2E0;

    background: #FFFFFF;
    color: #167D80;

    font-weight: 750;

    padding: 0.65rem 1rem;

    transition: all 0.25s ease;
}

.stButton > button:hover {
    background: #E5F5F1;
    border-color: #159D9B;
    color: #075E64;

    transform: translateY(-2px);
}

.stButton > button[kind="primary"] {
    background:
        linear-gradient(
            100deg,
            #087F83,
            #21A89D
        );

    color: white;
    border: none;
}

.stButton > button[kind="primary"]:hover {
    background: #076C72;
    color: white;
}


/* ==========================================
  PROGRESS BAR
========================================== */

.stProgress > div > div > div {
    background:
        linear-gradient(
            90deg,
            #087F83,
            #58B99A
        );
}


/* ==========================================
  ANIMATIONS
========================================== */

@keyframes fadeIn {

    from {
        opacity: 0;
        transform: translateY(16px);
    }

    to {
        opacity: 1;
        transform: translateY(0);
    }

}

.animate {
    animation: fadeIn 0.65s ease;
}


/* ==========================================
  FOOTER
========================================== */

.footer {
    text-align: center;

    color: #77929A;
    font-size: 0.83rem;

    border-top: 1px solid #DCEBEC;

    padding-top: 24px;
    margin-top: 55px;
}


/* ==========================================
  INPUT FIELDS
========================================== */

.stTextInput input {
    border-radius: 10px;
    border-color: #C9E2E0;
    background: white;
}

.stTextInput input:focus {
    border-color: #159D9B;
}


/* ==========================================
  MOBILE RESPONSIVENESS
========================================== */

@media (max-width: 700px) {

    .hero {
        padding: 30px 22px;
    }

    .hero h1 {
        font-size: 2rem;
    }

    .section-title {
        font-size: 1.45rem;
    }

    .brand {
        font-size: 1.65rem;
    }

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
            "The Excel database could not be found. "
            "Please upload portal.xlsx to the repository."
        )

    data = pd.read_excel(
        file_path,
        engine="openpyxl"
    )

    # Remove unnecessary spaces from column headings
    data.columns = (
        data.columns
        .astype(str)
        .str.strip()
    )

    missing_columns = [
        col for col in REQUIRED_COLUMNS
        if col not in data.columns
    ]

    if missing_columns:

        raise ValueError(
            f"Missing Excel columns: {missing_columns}"
        )

    # Retain only the required columns
    data = data[REQUIRED_COLUMNS].copy()

    # Clean text without changing the row order
    for col in REQUIRED_COLUMNS:

        data[col] = (
            data[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # Remove incomplete records
    data = data.replace("", pd.NA)

    data = data.dropna(
        subset=REQUIRED_COLUMNS
    )

    # IMPORTANT:
    # Original Excel row order is preserved.
    # No sorting is applied to Lean principles.

    return data.reset_index(drop=True)


try:

    df = load_database(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0
    )

except Exception as error:

    st.error(f"Database loading error: {error}")

    st.info(
        "Check that the Excel file is named portal.xlsx "
        "or Portal.xlsx and is stored alongside app.py."
    )

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
# 6. HEADER AND BRANDING
# ------------------------------------------------------------

col_logo, col_home = st.columns([3, 1])

with col_logo:

    st.markdown("""
        <div class="brand">🌊 BluePath</div>

        <div class="brand-subtitle">
            BLUE ECONOMY × LEAN CONSTRUCTION
        </div>
    """, unsafe_allow_html=True)


with col_home:

    if st.session_state.page != "welcome":

        if st.button(
            "🏠 Home",
            use_container_width=True
        ):

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

    current_step = PAGES[st.session_state.page]

    st.progress(current_step / 4)

    st.caption(
        f"YOUR DECISION-MAKING JOURNEY | "
        f"STEP {current_step} OF 4"
    )


# ============================================================
# 8. WELCOME INTERFACE
# ============================================================

if st.session_state.page == "welcome":

    st.markdown("""
        <div class="hero animate">

            <div style="font-size:48px;">
                🌊 🚢 🌱
            </div>

            <div class="hero-label">
                MARINE CONSTRUCTION DECISION SUPPORT
            </div>

            <h1>
                Welcome to BluePath
            </h1>

            <p>
                Explore how Blue Economy practices can address
                the environmental, social and economic impacts
                of marine construction.

                Discover how Lean principles can support the
                successful implementation of these practices
                and contribute to sustainable marine development.
            </p>

            <div style="
                margin-top:22px;
                font-size:14px;
                font-weight:750;
                color:#E1FFFF;
                letter-spacing:1px;
            ">
                IDENTIFY IMPACTS · EXPLORE PRACTICES ·
                ADDRESS CHALLENGES · DISCOVER LEAN SOLUTIONS
            </div>

        </div>
    """, unsafe_allow_html=True)


    st.markdown(
        '<div class="section-title">Start your journey</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-description">'
        'Welcome! Enter your name to begin exploring '
        'the BluePath decision-support platform.'
        '</div>',
        unsafe_allow_html=True
    )


    username = st.text_input(

        "Your name",

        placeholder="Enter your name here",

        max_chars=60

    )


    st.markdown("### 📘 How to use BluePath")


    instructions = [

        (
            "01",
            "Identify Marine Construction Impacts",

            "Explore the impacts associated with marine "
            "construction and select an impact you wish to manage."
        ),

        (
            "02",
            "Explore Blue Economy Practices",

            "Discover the suitable Blue Economy practices "
            "associated with the selected marine construction impact."
        ),

        (
            "03",
            "Identify Implementation Challenges",

            "Explore the challenges associated with implementing "
            "each selected Blue Economy practice."
        ),

        (
            "04",
            "Discover Relevant Lean Principles",

            "Identify all five Lean principles associated with "
            "each challenge, presented according to their "
            "individual priority rankings in the research database."
        )

    ]


    for number, heading, description in instructions:

        st.markdown(
            f"""
            <div class="info-card animate">

                <span class="number-circle">
                    {number}
                </span>

                <strong>
                    {heading}
                </strong>

                <p style="
                    color:#617D89;
                    margin:10px 0 0 50px;
                    line-height:1.7;
                ">
                    {description}
                </p>

            </div>
            """,
            unsafe_allow_html=True
        )


    st.markdown("")


    if st.button(

        "🚀 Start Exploring BluePath",

        type="primary",

        use_container_width=True

    ):

        if not username.strip():

            st.warning(
                "Please enter your name before continuing."
            )

        else:

            st.session_state.username = username.strip()

            st.session_state.page = "impacts"

            st.rerun()


# ============================================================
# 9. MARINE CONSTRUCTION IMPACT SELECTION
# ============================================================

elif st.session_state.page == "impacts":

    name = html.escape(
        st.session_state.username
    )


    st.markdown(
        f"""
        <div class="section-title animate">
            Hello, {name}! 👋
        </div>

        <div class="section-description">
            Welcome to your BluePath journey.

            Begin by identifying the marine construction
            impact that you wish to manage.
        </div>
        """,
        unsafe_allow_html=True
    )


    st.markdown("""
        <div class="notice">

            <strong>
                STEP 1 | Select a Marine Construction Impact
            </strong>

            <br>

            Explore the impacts listed below and select
            one to discover suitable Blue Economy practices.

        </div>
    """, unsafe_allow_html=True)


    impacts = (
        df[IMPACT_COL]
        .drop_duplicates()
        .tolist()
    )


    st.markdown(
        f"**{len(impacts)} impacts identified in the database**"
    )


    # Two-column interactive impact cards
    cols = st.columns(2)


    for i, impact in enumerate(impacts):

        with cols[i % 2]:

            st.markdown(
                f"""
                <div class="info-card animate">

                    <span class="number-circle">
                        {i+1}
                    </span>

                    <strong>
                        {html.escape(impact)}
                    </strong>

                </div>
                """,
                unsafe_allow_html=True
            )


            if st.button(

                f"Explore impact {i+1} →",

                key=f"impact_{i}",

                use_container_width=True

            ):

                st.session_state.selected_impact = impact

                st.session_state.selected_practice = None

                st.session_state.selected_challenge = None

                st.session_state.page = "practices"

                st.rerun()


# ============================================================
# 10. BLUE ECONOMY PRACTICE SELECTION
# ============================================================

elif st.session_state.page == "practices":

    impact = st.session_state.selected_impact


    if impact is None:

        go_home()

        st.rerun()


    st.markdown(
        '<div class="step-label">'
        'STEP 2 | BLUE ECONOMY RESPONSE'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        '<div class="section-title">'
        'Suitable Blue Economy Practices to Manage the Impact'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        f"""
        <div class="notice">

            <strong>Selected Marine Construction Impact:</strong>

            <br>

            {html.escape(impact)}

        </div>
        """,
        unsafe_allow_html=True
    )


    practices = (
        df.loc[
            df[IMPACT_COL] == impact,
            PRACTICE_COL
        ]
        .drop_duplicates()
        .tolist()
    )


    st.write(
        "The following Blue Economy practices are associated "
        "with the selected impact."
    )


    st.markdown(
        f"**{len(practices)} suitable practices identified**"
    )


    for i, practice in enumerate(practices):

        st.markdown(
            f"""
            <div class="info-card animate">

                <span class="number-circle">
                    {i+1}
                </span>

                <strong>
                    {html.escape(practice)}
                </strong>

            </div>
            """,
            unsafe_allow_html=True
        )


        if st.button(

            "Explore implementation challenges →",

            key=f"practice_{i}",

            use_container_width=True

        ):

            st.session_state.selected_practice = practice

            st.session_state.selected_challenge = None

            st.session_state.page = "challenges"

            st.rerun()


    st.divider()


    if st.button("← Back to Marine Construction Impacts"):

        st.session_state.page = "impacts"

        st.rerun()


# ============================================================
# 11. BLUE ECONOMY IMPLEMENTATION CHALLENGES
# ============================================================

elif st.session_state.page == "challenges":

    impact = st.session_state.selected_impact

    practice = st.session_state.selected_practice


    if impact is None or practice is None:

        go_home()

        st.rerun()


    st.markdown(
        '<div class="step-label">'
        'STEP 3 | IMPLEMENTATION CHALLENGES'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        '<div class="section-title">'
        'Challenges in Implementing Each Blue Economy Practice '
        'in Marine Construction'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        f"""
        <div class="notice">

            <strong>Selected Impact:</strong>

            <br>

            {html.escape(impact)}

            <br><br>

            <strong>Selected Blue Economy Practice:</strong>

            <br>

            {html.escape(practice)}

        </div>
        """,
        unsafe_allow_html=True
    )


    filtered = df[

        (df[IMPACT_COL] == impact) &

        (df[PRACTICE_COL] == practice)

    ]


    challenges = (
        filtered[CHALLENGE_COL]
        .drop_duplicates()
        .tolist()
    )


    st.write(
        "Select a challenge to discover the Lean principles "
        "that can help address it."
    )


    st.markdown(
        f"**{len(challenges)} implementation challenges identified**"
    )


    for i, challenge in enumerate(challenges):

        st.markdown(
            f"""
            <div class="info-card animate">

                <span class="number-circle">
                    {i+1}
                </span>

                <strong>
                    {html.escape(challenge)}
                </strong>

            </div>
            """,
            unsafe_allow_html=True
        )


        if st.button(

            "Discover relevant Lean principles →",

            key=f"challenge_{i}",

            use_container_width=True

        ):

            st.session_state.selected_challenge = challenge

            st.session_state.page = "lean"

            st.rerun()


    st.divider()


    if st.button("← Back to Blue Economy Practices"):

        st.session_state.page = "practices"

        st.rerun()


# ============================================================
# 12. LEAN PRINCIPLE PRIORITY IDENTIFICATION
# ============================================================

elif st.session_state.page == "lean":

    impact = st.session_state.selected_impact

    practice = st.session_state.selected_practice

    challenge = st.session_state.selected_challenge


    if not all([impact, practice, challenge]):

        go_home()

        st.rerun()


    st.markdown(
        '<div class="step-label">'
        'STEP 4 | LEAN-ENABLED SOLUTIONS'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        '<div class="section-title">'
        'Pertinent Lean Principles to Minimise Blue Economy '
        'Challenges in Marine Construction'
        '</div>',
        unsafe_allow_html=True
    )


    st.markdown(
        f"""
        <div class="notice">

            <strong>Selected Impact:</strong>

            <br>

            {html.escape(impact)}

            <br><br>

            <strong>Selected Practice:</strong>

            <br>

            {html.escape(practice)}

            <br><br>

            <strong>Selected Challenge:</strong>

            <br>

            {html.escape(challenge)}

        </div>
        """,
        unsafe_allow_html=True
    )


    # Filter records matching all criteria to retrieve the exact 5 Lean principles in original database row order
    lean_rows = df[

        (df[IMPACT_COL] == impact) &

        (df[PRACTICE_COL] == practice) &

        (df[CHALLENGE_COL] == challenge)

    ]


    lean_principles = lean_rows[LEAN_COL].tolist()


    st.write(
        "The following five Lean principles are prioritized "
        "according to your research records to address this specific challenge."
    )


    st.markdown(
        f"**{len(lean_principles)} prioritized Lean principles identified**"
    )


    for i, principle in enumerate(lean_principles):

        st.markdown(
            f"""
            <div class="lean-card animate">

                <div class="priority-label">
                    PRIORITY RANK {i+1} OF {len(lean_principles)}
                </div>

                <div style="font-size: 1.1rem; font-weight: 750; color: #165B33; margin-top: 6px;">
                    {html.escape(principle)}
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


    st.divider()


    col_back, col_restart = st.columns(2)


    with col_back:

        if st.button("← Back to Implementation Challenges", use_container_width=True):

            st.session_state.page = "challenges"

            st.rerun()


    with col_restart:

        if st.button("🔄 Start Again", use_container_width=True, type="primary"):

            reset_app()

            st.rerun()


# ============================================================
# 13. FOOTER
# ============================================================

st.markdown("""
    <div class="footer">
        BluePath Research Platform · Powered by Streamlit & Python
    </div>
""", unsafe_allow_html=True)
