
import html
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# BLUEPATH
# Blue Economy x Lean Construction
# Interactive Research Decision-Support Platform
# ============================================================

st.set_page_config(
    page_title="BluePath | Decision Support",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# 1. CONFIGURATION
# ============================================================

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

EXPECTED_LEAN_COUNT = 5


# ============================================================
# 2. CUSTOM DESIGN
# ============================================================

st.markdown("""
<style>

/* ==========================================================
   GLOBAL DESIGN
   ========================================================== */

.stApp {
    background: #F2F6F7;
    color: #203B46;
    font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
}

#MainMenu,
footer,
header {
    visibility: hidden;
}

.block-container {
    max-width: 1250px;
    padding-top: 1.2rem;
    padding-bottom: 3rem;
}

/* ==========================================================
   BRAND HEADER
   ========================================================== */

.brand {
    font-size: 2rem;
    font-weight: 800;
    letter-spacing: -0.7px;
    color: #0A6E74;
    margin-bottom: 0;
}

.brand-subtitle {
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    color: #66828A;
    text-transform: uppercase;
}

.header-line {
    height: 3px;
    background: linear-gradient(
        90deg,
        #0A6E74 0%,
        #249B8D 70%,
        #DDE9E9 100%
    );
    margin-top: 14px;
    margin-bottom: 25px;
}

/* ==========================================================
   WELCOME HERO
   ========================================================== */

.hero {
    background:
        linear-gradient(
            115deg,
            rgba(8, 74, 83, 0.97),
            rgba(10, 110, 116, 0.94),
            rgba(31, 143, 126, 0.91)
        );

    border-radius: 5px;
    padding: 42px 42px;
    margin: 14px 0 30px 0;
    color: white;
    border-left: 7px solid #E35B57;
}

.hero-label {
    color: #BFE6E0;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 2.5px;
    text-transform: uppercase;
}

.hero h1 {
    color: white;
    font-size: 2.6rem;
    font-weight: 800;
    margin: 10px 0;
    letter-spacing: -1px;
}

.hero p {
    color: #E6F3F2;
    font-size: 1rem;
    line-height: 1.8;
    max-width: 780px;
    margin-bottom: 0;
}

/* ==========================================================
   SECTION HEADINGS
   ========================================================== */

.section-title {
    color: #0D5F66;
    font-size: 1.55rem;
    font-weight: 750;
    margin-top: 5px;
    margin-bottom: 7px;
}

.section-description {
    color: #5B727B;
    font-size: 0.94rem;
    line-height: 1.7;
    margin-bottom: 20px;
}

.step-label {
    color: #0D8A8A;
    font-size: 0.72rem;
    font-weight: 800;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 6px;
}

.small-label {
    color: #66818A;
    font-size: 0.68rem;
    font-weight: 800;
    letter-spacing: 1.4px;
    text-transform: uppercase;
    margin-bottom: 6px;
}

/* ==========================================================
   DECISION PATHWAY NAVIGATION
   ========================================================== */

.pathway-container {
    background: #FFFFFF;
    border: 1px solid #D6E3E5;
    padding: 17px 20px;
    margin-bottom: 25px;
}

.pathway-flow {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
}

.pathway-step {
    flex: 1;
    padding: 11px 8px;
    text-align: center;
    background: #F0F4F5;
    border: 1px solid #DCE5E7;
    color: #75858B;
    font-size: 0.75rem;
    font-weight: 700;
}

.pathway-step.active {
    background: #E3F2EF;
    border: 1px solid #0A6E74;
    color: #0A6E74;
    border-bottom: 3px solid #0A6E74;
}

.pathway-step.completed {
    background: #FCEDEC;
    border: 1px solid #E35B57;
    color: #B33C39;
}

.pathway-arrow {
    color: #E35B57;
    font-size: 1.15rem;
    font-weight: 800;
}

/* ==========================================================
   STAGE BACKGROUNDS
   ========================================================== */

.stage-impact {
    background: #E8F1F5;
    border-left: 5px solid #287B93;
}

.stage-practice {
    background: #E7F3F0;
    border-left: 5px solid #0A8B80;
}

.stage-challenge {
    background: #F7EEEA;
    border-left: 5px solid #D16B56;
}

.stage-lean {
    background: #EDF3E9;
    border-left: 5px solid #53865D;
}

.stage-banner {
    padding: 19px 23px;
    margin: 15px 0 20px 0;
    border-top: 1px solid rgba(50, 80, 90, 0.10);
    border-bottom: 1px solid rgba(50, 80, 90, 0.10);
}

.stage-banner-title {
    color: #203B46;
    font-size: 1rem;
    font-weight: 800;
    margin-bottom: 5px;
}

.stage-banner-text {
    color: #526B75;
    font-size: 0.88rem;
    line-height: 1.6;
}

/* ==========================================================
   CLICKABLE SELECTION CARDS
   ========================================================== */

.card-intro {
    color: #627982;
    font-size: 0.83rem;
    margin-bottom: 12px;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 3px;
}

/* Selection buttons */
.stButton > button {
    width: 100%;
    min-height: 52px;
    border-radius: 3px;
    border: 1px solid #D0DFE1;
    background: #FFFFFF;
    color: #24444E;
    font-size: 0.87rem;
    font-weight: 650;
    text-align: left;
    padding: 10px 15px;
    transition: all 0.18s ease;
}

/* Hover: red feedback */
.stButton > button:hover {
    background: #FFF1F0;
    border: 1px solid #E35B57;
    color: #B53C38;
    box-shadow: 0 3px 12px rgba(190, 65, 60, 0.10);
}

/* Primary action */
.stButton > button[kind="primary"] {
    background: #0A6E74;
    color: #FFFFFF;
    border: 1px solid #0A6E74;
    border-radius: 3px;
    text-align: center;
    font-weight: 750;
}

.stButton > button[kind="primary"]:hover {
    background: #E35B57;
    border-color: #E35B57;
    color: #FFFFFF;
}

/* ==========================================================
   INFORMATION PANELS
   ========================================================== */

.info-panel {
    background: #FFFFFF;
    border: 1px solid #D8E4E6;
    padding: 18px 20px;
    margin: 10px 0 15px 0;
}

.notice {
    background: #E9F3F3;
    border-left: 4px solid #0A6E74;
    padding: 14px 18px;
    margin: 14px 0 20px 0;
    color: #34545E;
    font-size: 0.88rem;
    line-height: 1.7;
}

.metric-box {
    background: #FFFFFF;
    border: 1px solid #D9E4E6;
    padding: 16px 18px;
    min-height: 90px;
}

.metric-value {
    color: #0A6E74;
    font-size: 1.65rem;
    font-weight: 800;
}

.metric-caption {
    color: #70838A;
    font-size: 0.75rem;
    margin-top: 4px;
}

/* ==========================================================
   LEAN PRINCIPLE CARDS
   ========================================================== */

.lean-card {
    background: #FFFFFF;
    border: 1px solid #D7E3D8;
    border-left: 5px solid #47815A;
    padding: 19px 22px;
    margin-bottom: 13px;
    transition: all 0.18s ease;
}

.lean-card:hover {
    border-left-color: #E35B57;
    background: #FFF8F6;
    box-shadow: 0 4px 14px rgba(40, 80, 60, 0.08);
}

.lean-priority {
    color: #4D8059;
    font-size: 0.68rem;
    font-weight: 800;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin-bottom: 6px;
}

.lean-name {
    color: #214A36;
    font-size: 1rem;
    font-weight: 750;
    line-height: 1.6;
}

/* ==========================================================
   SUMMARY
   ========================================================== */

.summary-panel {
    background: #FFFFFF;
    border: 1px solid #D8E4E6;
    padding: 20px;
    margin-top: 20px;
}

/* ==========================================================
   PROGRESS
   ========================================================== */

.stProgress > div > div > div {
    background: #0A6E74;
}

[data-testid="stProgressBar"] {
    background: #DCE8E9;
}

/* ==========================================================
   FOOTER
   ========================================================== */

.footer {
    border-top: 1px solid #D5E2E4;
    margin-top: 45px;
    padding-top: 17px;
    color: #74878D;
    text-align: center;
    font-size: 0.76rem;
    line-height: 1.8;
}

/* ==========================================================
   RESPONSIVE
   ========================================================== */

@media (max-width: 700px) {

    .hero {
        padding: 25px 20px;
    }

    .hero h1 {
        font-size: 2rem;
    }

    .section-title {
        font-size: 1.3rem;
    }

    .pathway-flow {
        flex-wrap: wrap;
    }

    .pathway-step {
        min-width: 40%;
    }

}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 3. LOAD DATABASE
# ============================================================

@st.cache_data
def load_database(file_path, file_modified):

    if not Path(file_path).exists():
        raise FileNotFoundError(
            "Database not found. Place portal.xlsx next to app.py."
        )

    data = pd.read_excel(
        file_path,
        engine="openpyxl"
    )

    data.columns = data.columns.astype(str).str.strip()

    missing = [
        column for column in REQUIRED_COLUMNS
        if column not in data.columns
    ]

    if missing:
        raise ValueError(
            f"Missing required Excel columns: {missing}"
        )

    data = data[REQUIRED_COLUMNS].copy()

    for column in REQUIRED_COLUMNS:
        data[column] = (
            data[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    data = data[
        (data[IMPACT_COL] != "")
        & (data[PRACTICE_COL] != "")
        & (data[CHALLENGE_COL] != "")
        & (data[LEAN_COL] != "")
    ].copy()

    return data.reset_index(drop=True)


try:

    df = load_database(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime if EXCEL_FILE.exists() else 0
    )

except Exception as error:

    st.error(f"Database loading error: {error}")
    st.stop()


if df.empty:
    st.error("The Excel database contains no valid records.")
    st.stop()


# ============================================================
# 4. SESSION STATE
# ============================================================

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


# ============================================================
# 5. HELPER FUNCTIONS
# ============================================================

def safe_text(value):
    return html.escape(str(value))


def show_stage_banner(title, description, stage_class):

    st.markdown(
        f"""
        <div class="stage-banner {stage_class}">
            <div class="stage-banner-title">
                {safe_text(title)}
            </div>
            <div class="stage-banner-text">
                {safe_text(description)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_pathway(current_step):

    labels = [
        "01  Impacts",
        "02  Practices",
        "03  Challenges",
        "04  Lean Principles"
    ]

    elements = []

    for index, label in enumerate(labels, start=1):

        if index < current_step:
            css_class = "completed"

        elif index == current_step:
            css_class = "active"

        else:
            css_class = ""

        elements.append(
            f'<div class="pathway-step {css_class}">'
            f'{label}</div>'
        )

        if index < 4:
            elements.append(
                '<div class="pathway-arrow">→</div>'
            )

    st.markdown(
        f"""
        <div class="pathway-container">
            <div class="small-label">Decision pathway</div>
            <div class="pathway-flow">
                {''.join(elements)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_selection_summary(impact=None, practice=None, challenge=None):

    parts = []

    if impact:
        parts.append(
            f"<strong>Impact:</strong> {safe_text(impact)}"
        )

    if practice:
        parts.append(
            f"<strong>Blue Economy practice:</strong> "
            f"{safe_text(practice)}"
        )

    if challenge:
        parts.append(
            f"<strong>Challenge:</strong> {safe_text(challenge)}"
        )

    st.markdown(
        '<div class="info-panel">'
        + "<br><br>".join(parts)
        + "</div>",
        unsafe_allow_html=True
    )


def show_back_and_reset(back_label, back_page):

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            f"←  {back_label}",
            use_container_width=True
        ):

            st.session_state.page = back_page
            st.rerun()

    with col2:

        if st.button(
            "Start again",
            use_container_width=True
        ):

            reset_app()
            st.rerun()


# ============================================================
# 6. HEADER
# ============================================================

col_logo, col_nav = st.columns([4, 1])

with col_logo:

    st.markdown(
        """
        <div class="brand">BluePath</div>
        <div class="brand-subtitle">
            Blue Economy x Lean Construction
        </div>
        """,
        unsafe_allow_html=True
    )

with col_nav:

    if st.session_state.page != "welcome":

        if st.button("Home", use_container_width=True):

            go_home()
            st.rerun()


st.markdown(
    '<div class="header-line"></div>',
    unsafe_allow_html=True
)


# ============================================================
# 7. PROGRESS
# ============================================================

PAGES = {
    "welcome": 0,
    "impacts": 1,
    "practices": 2,
    "challenges": 3,
    "lean": 4
}

current_page = st.session_state.page
step = PAGES[current_page]

if current_page != "welcome":

    show_pathway(step)

    st.progress(step / 4)

    st.caption(
        f"Decision pathway | Step {step} of 4"
    )


# ============================================================
# 8. WELCOME PAGE
# ============================================================

if current_page == "welcome":

    st.markdown(
        """
        <div class="hero">
            <div class="hero-label">
                Research Decision-Support Platform
            </div>

            <h1>BluePath</h1>

            <p>
                Connect marine construction impacts with relevant
                Blue Economy practices and prioritised Lean principles
                to support sustainable implementation in maritime
                infrastructure projects.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">'
        'Start your decision pathway'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Follow the four stages to identify suitable Blue Economy
            practices and corresponding Lean principles for a selected
            marine construction impact.
        </div>
        """,
        unsafe_allow_html=True
    )

    username = st.text_input(
        "Researcher name (optional)",
        placeholder="Enter your name",
        max_chars=60
    )

    st.markdown("### Decision pathway")

    steps = [
        (
            "01",
            "Marine construction impacts",
            "Identify the impact that requires attention."
        ),
        (
            "02",
            "Blue Economy practices",
            "Identify practices associated with the selected impact."
        ),
        (
            "03",
            "Implementation challenges",
            "Identify challenges associated with the selected practice."
        ),
        (
            "04",
            "Lean principles",
            "Review the five principles in the research database priority order."
        )
    ]

    for number, title, description in steps:

        st.markdown(
            f"""
            <div class="pathway-card">
                <div style="display:flex;align-items:center;gap:14px;">
                    <div style="
                        min-width:38px;
                        height:38px;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        background:#E4F3F1;
                        color:#0A6E74;
                        font-weight:800;
                        border-left:3px solid #E35B57;
                    ">
                        {number}
                    </div>

                    <div>
                        <div style="
                            font-weight:750;
                            color:#24444E;
                        ">
                            {safe_text(title)}
                        </div>

                        <div style="
                            color:#71848B;
                            font-size:0.83rem;
                            margin-top:5px;
                        ">
                            {safe_text(description)}
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    if st.button(
        "Begin decision process  →",
        type="primary",
        use_container_width=True
    ):

        st.session_state.username = username.strip()
        st.session_state.page = "impacts"
        st.rerun()


# ============================================================
# 9. MARINE CONSTRUCTION IMPACTS
# ============================================================

elif current_page == "impacts":

    researcher = (
        safe_text(st.session_state.username)
        if st.session_state.username
        else "Researcher"
    )

    st.markdown(
        '<div class="step-label">Stage 01</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">'
        'Marine Construction Impacts'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Select the marine construction impact that you intend
            to address through Blue Economy practices.
        </div>
        """,
        unsafe_allow_html=True
    )

    show_stage_banner(
        "Impact identification",
        "Select an impact to explore the relevant Blue Economy "
        "practices and implementation pathways.",
        "stage-impact"
    )

    impacts = df[IMPACT_COL].drop_duplicates().tolist()

    st.markdown(
        f'<div class="card-intro">'
        f'{len(impacts)} impact categories available in the database.'
        f'</div>',
        unsafe_allow_html=True
    )

    search = st.text_input(
        "Search impacts",
        placeholder="Type a keyword to filter the available impacts..."
    )

    filtered_impacts = [
        impact for impact in impacts
        if search.lower() in impact.lower()
    ]

    if not filtered_impacts:

        st.warning("No matching impacts were found.")

    else:

        cols = st.columns(2)

        for index, impact in enumerate(filtered_impacts):

            with cols[index % 2]:

                st.markdown(
                    f"""
                    <div style="
                        background:#FFFFFF;
                        border:1px solid #D5E2E6;
                        border-left:4px solid #287B93;
                        padding:12px 14px;
                        margin-top:8px;
                        margin-bottom:5px;
                        min-height:65px;
                    ">
                        <span style="
                            color:#287B93;
                            font-weight:800;
                            font-size:0.72rem;
                        ">
                            IMPACT {index + 1:02d}
                        </span>

                        <div style="
                            color:#24444E;
                            font-weight:700;
                            font-size:0.88rem;
                            margin-top:5px;
                        ">
                            {safe_text(impact)}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if st.button(
                    "Select impact  →",
                    key=f"impact_{index}",
                    use_container_width=True
                ):

                    st.session_state.selected_impact = impact
                    st.session_state.selected_practice = None
                    st.session_state.selected_challenge = None
                    st.session_state.page = "practices"
                    st.rerun()


# ============================================================
# 10. BLUE ECONOMY PRACTICES
# ============================================================

elif current_page == "practices":

    impact = st.session_state.selected_impact

    if not impact:
        go_home()
        st.rerun()

    st.markdown(
        '<div class="step-label">Stage 02</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">'
        'Blue Economy Practices'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Select a Blue Economy practice relevant to the previously
            selected marine construction impact.
        </div>
        """,
        unsafe_allow_html=True
    )

    show_stage_banner(
        "Selected marine construction impact",
        impact,
        "stage-practice"
    )

    practices = (
        df.loc[
            df[IMPACT_COL] == impact,
            PRACTICE_COL
        ]
        .drop_duplicates()
        .tolist()
    )

    st.markdown(
        f'<div class="card-intro">'
        f'{len(practices)} relevant Blue Economy practices identified.'
        f'</div>',
        unsafe_allow_html=True
    )

    search = st.text_input(
        "Search Blue Economy practices",
        placeholder="Type a keyword to filter the practices..."
    )

    filtered_practices = [
        practice for practice in practices
        if search.lower() in practice.lower()
    ]

    if not filtered_practices:

        st.warning("No matching practices were found.")

    else:

        for index, practice in enumerate(filtered_practices):

            st.markdown(
                f"""
                <div style="
                    background:#FFFFFF;
                    border:1px solid #D6E5E1;
                    border-left:4px solid #0A8B80;
                    padding:13px 17px;
                    margin-top:9px;
                    margin-bottom:5px;
                ">
                    <span style="
                        color:#0A8B80;
                        font-weight:800;
                        font-size:0.72rem;
                    ">
                        PRACTICE {index + 1:02d}
                    </span>

                    <div style="
                        color:#24444E;
                        font-weight:700;
                        font-size:0.90rem;
                        margin-top:5px;
                    ">
                        {safe_text(practice)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            if st.button(
                "Select practice  →",
                key=f"practice_{index}",
                use_container_width=True
            ):

                st.session_state.selected_practice = practice
                st.session_state.selected_challenge = None
                st.session_state.page = "challenges"
                st.rerun()

    show_back_and_reset(
        "Back to impacts",
        "impacts"
    )


# ============================================================
# 11. IMPLEMENTATION CHALLENGES
# ============================================================

elif current_page == "challenges":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice

    if not impact or not practice:
        go_home()
        st.rerun()

    st.markdown(
        '<div class="step-label">Stage 03</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">'
        'Blue Economy Implementation Challenges'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Identify the implementation challenge associated with
            the selected Blue Economy practice and marine construction
            impact.
        </div>
        """,
        unsafe_allow_html=True
    )

    show_selection_summary(
        impact=impact,
        practice=practice
    )

    filtered = df[
        (df[IMPACT_COL] == impact)
        & (df[PRACTICE_COL] == practice)
    ]

    challenges = (
        filtered[CHALLENGE_COL]
        .drop_duplicates()
        .tolist()
    )

    show_stage_banner(
        "Implementation challenge identification",
        "Select a challenge to identify the corresponding Lean "
        "principles in the research database.",
        "stage-challenge"
    )

    st.markdown(
        f'<div class="card-intro">'
        f'{len(challenges)} associated challenges identified.'
        f'</div>',
        unsafe_allow_html=True
    )

    search = st.text_input(
        "Search implementation challenges",
        placeholder="Type a keyword to filter the challenges..."
    )

    filtered_challenges = [
        challenge for challenge in challenges
        if search.lower() in challenge.lower()
    ]

    if not filtered_challenges:

        st.warning("No matching challenges were found.")

    else:

        for index, challenge in enumerate(filtered_challenges):

            st.markdown(
                f"""
                <div style="
                    background:#FFFFFF;
                    border:1px solid #E8DCD6;
                    border-left:4px solid #D16B56;
                    padding:13px 17px;
                    margin-top:9px;
                    margin-bottom:5px;
                ">
                    <span style="
                        color:#C35C49;
                        font-weight:800;
                        font-size:0.72rem;
                    ">
                        CHALLENGE {index + 1:02d}
                    </span>

                    <div style="
                        color:#24444E;
                        font-weight:700;
                        font-size:0.90rem;
                        margin-top:5px;
                    ">
                        {safe_text(challenge)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            if st.button(
                "View Lean principles  →",
                key=f"challenge_{index}",
                use_container_width=True
            ):

                st.session_state.selected_challenge = challenge
                st.session_state.page = "lean"
                st.rerun()

    show_back_and_reset(
        "Back to practices",
        "practices"
    )


# ============================================================
# 12. LEAN PRINCIPLES
# ============================================================

elif current_page == "lean":

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    challenge = st.session_state.selected_challenge

    if not all([impact, practice, challenge]):
        go_home()
        st.rerun()

    st.markdown(
        '<div class="step-label">Stage 04</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-title">'
        'Prioritised Lean Principles'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Review the Lean principles associated with the selected
            marine construction impact, Blue Economy practice and
            implementation challenge.
        </div>
        """,
        unsafe_allow_html=True
    )

    show_selection_summary(
        impact=impact,
        practice=practice,
        challenge=challenge
    )

    show_stage_banner(
        "Lean principle selection",
        "The principles below follow the exact sequence recorded "
        "in the research database. The sequence is retained as the "
        "challenge-specific priority order.",
        "stage-lean"
    )

    filtered = df[
        (df[IMPACT_COL] == impact)
        & (df[PRACTICE_COL] == practice)
        & (df[CHALLENGE_COL] == challenge)
    ].copy()

    # Remove duplicate principles while retaining Excel order
    filtered = filtered.drop_duplicates(
        subset=[LEAN_COL],
        keep="first"
    ).reset_index(drop=True)

    lean_principles = filtered[LEAN_COL].tolist()

    if not lean_principles:

        st.warning(
            "No Lean principles were found for this selection."
        )

    else:

        if len(lean_principles) == EXPECTED_LEAN_COUNT:

            st.success(
                "Five Lean principles identified for this decision pathway."
            )

        else:

            st.warning(
                f"The database contains {len(lean_principles)} "
                f"unique Lean principles rather than the expected "
                f"{EXPECTED_LEAN_COUNT}. All available records "
                "are displayed without adding missing principles."
            )

        for index, principle in enumerate(
            lean_principles,
            start=1
        ):

            st.markdown(
                f"""
                <div class="lean-card">
                    <div class="lean-priority">
                        PRIORITY {index:02d}
                    </div>

                    <div class="lean-name">
                        {safe_text(principle)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

    # --------------------------------------------------------
    # DECISION SUMMARY
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title" style="margin-top:30px;">'
        'Decision Summary'
        '</div>',
        unsafe_allow_html=True
    )

    summary = pd.DataFrame([
        {
            "Decision component": "Marine construction impact",
            "Selected item": impact
        },
        {
            "Decision component": "Blue Economy practice",
            "Selected item": practice
        },
        {
            "Decision component": "Implementation challenge",
            "Selected item": challenge
        },
        {
            "Decision component": "Number of Lean principles",
            "Selected item": str(len(lean_principles))
        }
    ])

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # DOWNLOAD RESULTS
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title" style="margin-top:28px;">'
        'Export Decision Results'
        '</div>',
        unsafe_allow_html=True
    )

    if not filtered.empty:

        export_data = filtered.copy()

        export_data.insert(
            0,
            "Selected Marine Construction Impact",
            impact
        )

        export_data.insert(
            1,
            "Selected Blue Economy Practice",
            practice
        )

        export_data.insert(
            2,
            "Selected Implementation Challenge",
            challenge
        )

        csv_data = export_data.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "Download decision results (CSV)",
            data=csv_data,
            file_name="bluepath_decision_results.csv",
            mime="text/csv",
            use_container_width=True
        )

    # --------------------------------------------------------
    # FINAL NAVIGATION
    # --------------------------------------------------------

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "← Back to challenges",
            use_container_width=True
        ):

            st.session_state.page = "challenges"
            st.rerun()

    with col2:

        if st.button(
            "Home",
            use_container_width=True
        ):

            go_home()
            st.rerun()

    with col3:

        if st.button(
            "Start new decision",
            type="primary",
            use_container_width=True
        ):

            reset_app()
            st.rerun()


# ============================================================
# 13. FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        <strong>BluePath</strong><br>
        Blue Economy x Lean Construction<br>
        Research decision-support platform for marine infrastructure
    </div>
    """,
    unsafe_allow_html=True
)
