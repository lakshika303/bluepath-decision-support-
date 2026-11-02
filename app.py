from pathlib import Path
import html

import pandas as pd
import streamlit as st


# ============================================================
# BLUEPATH | BLUE ECONOMY DECISION SUPPORT SYSTEM
# ============================================================

st.set_page_config(
    page_title="BluePath | Decision Support System",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_TITLE = "BLUEPATH"
APP_SUBTITLE = "Marine Construction Decision Support System"

EXCEL_FILE = Path(__file__).parent / "portal.xlsx"

COL_IMPACT = "Impact of Marine Construction"
COL_PRACTICE = "Blue Economy Practice"
COL_CHALLENGE = "Challenges to Implement Blue Economy Practice"
COL_LEAN = "Lean Principles"

REQUIRED_COLUMNS = [
    COL_IMPACT,
    COL_PRACTICE,
    COL_CHALLENGE,
    COL_LEAN
]


# ============================================================
# PROFESSIONAL INTERFACE STYLING
# ============================================================

st.markdown(
    """
    <style>

    /* General application */
    .stApp {
        background-color: #F4F6F8;
        color: #263746;
        font-family: Arial, Helvetica, sans-serif;
    }

    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header {
        background-color: transparent;
    }

    .block-container {
        padding-top: 1.3rem;
        padding-bottom: 3rem;
        max-width: 1350px;
    }

    /* Main header */
    .brand-header {
        background-color: #142F43;
        padding: 24px 30px;
        border-bottom: 4px solid #159A91;
        margin-bottom: 22px;
    }

    .brand-name {
        color: #FFFFFF;
        font-size: 25px;
        font-weight: 800;
        letter-spacing: 2px;
        margin: 0;
    }

    .brand-subtitle {
        color: #D5E1E8;
        font-size: 13px;
        margin-top: 7px;
        letter-spacing: 0.5px;
    }

    .brand-tag {
        color: #8DD6CC;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-bottom: 5px;
    }

    /* Section headings */
    .section-heading {
        color: #142F43;
        font-size: 21px;
        font-weight: 700;
        margin-top: 8px;
        margin-bottom: 5px;
    }

    .section-description {
        color: #657786;
        font-size: 13px;
        line-height: 1.6;
        margin-bottom: 20px;
    }

    .small-label {
        color: #617383;
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 5px;
    }

    /* Progress */
    .progress-wrapper {
        background-color: #FFFFFF;
        border: 1px solid #DCE3E8;
        padding: 16px 20px;
        margin-bottom: 24px;
    }

    .progress-title {
        font-size: 12px;
        font-weight: 700;
        color: #526575;
        margin-bottom: 10px;
        letter-spacing: 0.5px;
    }

    .progress-track {
        height: 6px;
        width: 100%;
        background-color: #E6ECEF;
    }

    .progress-fill {
        height: 6px;
        background-color: #159A91;
    }

    .progress-steps {
        display: flex;
        justify-content: space-between;
        margin-top: 9px;
        font-size: 11px;
        color: #71808C;
    }

    .progress-active {
        color: #087F78;
        font-weight: 700;
    }

    /* Selection panels */
    .selection-panel {
        background-color: #FFFFFF;
        border: 1px solid #DCE3E8;
        border-left: 4px solid #159A91;
        padding: 19px 22px;
        margin-bottom: 15px;
    }

    .selection-title {
        font-size: 14px;
        font-weight: 700;
        color: #18384D;
        margin-bottom: 8px;
    }

    .selection-text {
        font-size: 13px;
        color: #526575;
        line-height: 1.6;
    }

    /* Information cards */
    .metric-panel {
        background-color: #FFFFFF;
        border: 1px solid #DCE3E8;
        padding: 17px 20px;
        min-height: 100px;
    }

    .metric-number {
        font-size: 25px;
        font-weight: 700;
        color: #142F43;
    }

    .metric-caption {
        font-size: 12px;
        color: #71808C;
        margin-top: 4px;
    }

    /* Result panels */
    .result-panel {
        background-color: #FFFFFF;
        border: 1px solid #DCE3E8;
        padding: 20px 22px;
        margin-bottom: 12px;
    }

    .result-number {
        display: inline-block;
        background-color: #E7F3F1;
        color: #087F78;
        font-size: 12px;
        font-weight: 700;
        padding: 5px 9px;
        margin-right: 10px;
        min-width: 32px;
        text-align: center;
    }

    .result-name {
        color: #18384D;
        font-size: 14px;
        font-weight: 700;
        line-height: 1.5;
    }

    .result-description {
        color: #71808C;
        font-size: 12px;
        margin-top: 6px;
    }

    /* Buttons */
    .stButton > button {
        border-radius: 2px;
        border: 1px solid #C8D3DA;
        min-height: 40px;
        font-weight: 600;
        font-size: 13px;
        background-color: #FFFFFF;
        color: #18384D;
    }

    .stButton > button:hover {
        border-color: #159A91;
        color: #087F78;
        background-color: #F0F8F7;
    }

    /* Primary buttons */
    .stButton > button[kind="primary"] {
        background-color: #142F43;
        color: #FFFFFF;
        border: 1px solid #142F43;
    }

    .stButton > button[kind="primary"]:hover {
        background-color: #087F78;
        border-color: #087F78;
        color: #FFFFFF;
    }

    /* Input labels */
    .stSelectbox label,
    .stTextInput label {
        color: #18384D !important;
        font-size: 13px !important;
        font-weight: 700 !important;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] {
        border: 1px solid #DCE3E8;
    }

    /* Footer */
    .app-footer {
        border-top: 1px solid #DCE3E8;
        margin-top: 45px;
        padding-top: 15px;
        color: #84919B;
        font-size: 11px;
        text-align: center;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOAD AND VALIDATE EXCEL DATA
# ============================================================

@st.cache_data
def load_data(file_path, modified_time):

    data = pd.read_excel(file_path)

    # Remove unnecessary spaces from column names
    data.columns = data.columns.astype(str).str.strip()

    missing_columns = [
        column for column in REQUIRED_COLUMNS
        if column not in data.columns
    ]

    if missing_columns:
        raise ValueError(
            "The Excel file is missing the following columns: "
            + ", ".join(missing_columns)
        )

    # Retain only the required columns
    data = data[REQUIRED_COLUMNS].copy()

    # Clean text values without changing row order
    for column in REQUIRED_COLUMNS:
        data[column] = (
            data[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # Remove rows without essential information
    data = data[
        (data[COL_IMPACT] != "")
        & (data[COL_PRACTICE] != "")
        & (data[COL_CHALLENGE] != "")
        & (data[COL_LEAN] != "")
    ].copy()

    # Reset index while preserving Excel order
    data.reset_index(drop=True, inplace=True)

    return data


if not EXCEL_FILE.exists():

    st.error(
        "The Excel data file 'portal.xlsx' could not be found. "
        "Please place portal.xlsx in the same folder as app.py."
    )

    st.stop()


try:

    df = load_data(
        str(EXCEL_FILE),
        EXCEL_FILE.stat().st_mtime
    )

except Exception as error:

    st.error(f"Unable to load the Excel data: {error}")
    st.stop()


if df.empty:

    st.error("No valid records were found in the Excel file.")
    st.stop()


# ============================================================
# SESSION STATE
# ============================================================

if "current_page" not in st.session_state:
    st.session_state.current_page = 1

if "selected_impact" not in st.session_state:
    st.session_state.selected_impact = None

if "selected_practice" not in st.session_state:
    st.session_state.selected_practice = None

if "selected_challenge" not in st.session_state:
    st.session_state.selected_challenge = None


def reset_application():

    st.session_state.current_page = 1
    st.session_state.selected_impact = None
    st.session_state.selected_practice = None
    st.session_state.selected_challenge = None


def go_back():

    if st.session_state.current_page > 1:
        st.session_state.current_page -= 1


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="brand-header">
        <div class="brand-tag">Research Decision Support Platform</div>
        <div class="brand-name">BLUEPATH</div>
        <div class="brand-subtitle">
            Blue Economy and Lean Principles for Marine Construction
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PROGRESS INDICATOR
# ============================================================

page = st.session_state.current_page

progress_values = {
    1: 25,
    2: 50,
    3: 75,
    4: 100
}

progress = progress_values[page]

step_labels = [
    "Marine Impact",
    "Blue Economy Practice",
    "Implementation Challenge",
    "Lean Principles"
]

step_html = ""

for index, label in enumerate(step_labels, start=1):

    css_class = (
        "progress-active"
        if index <= page
        else ""
    )

    step_html += (
        f'<span class="{css_class}">{label}</span>'
    )


st.markdown(
    f"""
    <div class="progress-wrapper">
        <div class="progress-title">
            DECISION PROCESS
        </div>

        <div class="progress-track">
            <div class="progress-fill"
                 style="width: {progress}%;">
            </div>
        </div>

        <div class="progress-steps">
            {step_html}
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PAGE 1: SELECT MARINE CONSTRUCTION IMPACT
# ============================================================

if page == 1:

    st.markdown(
        '<div class="section-heading">'
        '01 | Marine Construction Impact'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Select the environmental, social or economic impact
            associated with the marine construction project.
            The system will identify the relevant Blue Economy practices.
        </div>
        """,
        unsafe_allow_html=True
    )

    impacts = df[COL_IMPACT].drop_duplicates().tolist()

    search_impact = st.text_input(
        "Search marine construction impacts",
        placeholder="Enter a keyword to filter the impact list..."
    )

    filtered_impacts = [
        impact for impact in impacts
        if search_impact.lower() in impact.lower()
    ]

    if not filtered_impacts:

        st.warning("No matching marine construction impacts found.")

    else:

        selected = st.selectbox(
            "Available marine construction impacts",
            options=filtered_impacts,
            index=None,
            placeholder="Select a marine construction impact"
        )

        st.caption(
            f"{len(filtered_impacts)} impact categories available"
        )

        if selected:

            impact_records = df[
                df[COL_IMPACT] == selected
            ]

            practices_count = (
                impact_records[COL_PRACTICE]
                .nunique()
            )

            challenges_count = (
                impact_records[COL_CHALLENGE]
                .nunique()
            )

            st.markdown(
                f"""
                <div class="selection-panel">
                    <div class="selection-title">
                        Selected marine construction impact
                    </div>
                    <div class="selection-text">
                        {html.escape(selected)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            col1, col2 = st.columns(2)

            with col1:

                st.markdown(
                    f"""
                    <div class="metric-panel">
                        <div class="metric-number">
                            {practices_count}
                        </div>
                        <div class="metric-caption">
                            Associated Blue Economy practices
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            with col2:

                st.markdown(
                    f"""
                    <div class="metric-panel">
                        <div class="metric-number">
                            {challenges_count}
                        </div>
                        <div class="metric-caption">
                            Associated implementation challenges
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            st.write("")

            if st.button(
                "Continue to Blue Economy Practices",
                type="primary",
                use_container_width=True
            ):

                st.session_state.selected_impact = selected
                st.session_state.selected_practice = None
                st.session_state.selected_challenge = None
                st.session_state.current_page = 2

                st.rerun()


# ============================================================
# PAGE 2: SELECT BLUE ECONOMY PRACTICE
# ============================================================

elif page == 2:

    impact = st.session_state.selected_impact

    if not impact:

        reset_application()
        st.rerun()

    st.markdown(
        '<div class="section-heading">'
        '02 | Blue Economy Practice'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Identify the Blue Economy practice that can address
            the selected marine construction impact.
            Only practices associated with the selected impact
            are displayed.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="selection-panel">
            <div class="small-label">Selected impact</div>
            <div class="selection-text">
                {html.escape(impact)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    impact_data = df[
        df[COL_IMPACT] == impact
    ]

    practices = (
        impact_data[COL_PRACTICE]
        .drop_duplicates()
        .tolist()
    )

    search_practice = st.text_input(
        "Search Blue Economy practices",
        placeholder="Enter a practice or keyword..."
    )

    filtered_practices = [
        practice for practice in practices
        if search_practice.lower() in practice.lower()
    ]

    st.caption(
        f"{len(filtered_practices)} relevant practices identified"
    )

    if not filtered_practices:

        st.warning("No matching practices found.")

    else:

        selected = st.selectbox(
            "Available Blue Economy practices",
            options=filtered_practices,
            index=None,
            placeholder="Select a Blue Economy practice"
        )

        if selected:

            practice_data = impact_data[
                impact_data[COL_PRACTICE] == selected
            ]

            challenges_count = (
                practice_data[COL_CHALLENGE].nunique()
            )

            st.markdown(
                f"""
                <div class="selection-panel">
                    <div class="selection-title">
                        Selected Blue Economy practice
                    </div>
                    <div class="selection-text">
                        {html.escape(selected)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.info(
                f"{challenges_count} implementation challenges "
                "are associated with this practice and impact."
            )

            if st.button(
                "Continue to Implementation Challenges",
                type="primary",
                use_container_width=True
            ):

                st.session_state.selected_practice = selected
                st.session_state.selected_challenge = None
                st.session_state.current_page = 3

                st.rerun()

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        if st.button("Back to Marine Impacts"):
            st.session_state.current_page = 1
            st.rerun()

    with col2:

        if st.button("Start Again"):
            reset_application()
            st.rerun()


# ============================================================
# PAGE 3: SELECT IMPLEMENTATION CHALLENGE
# ============================================================

elif page == 3:

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice

    if not impact or not practice:

        reset_application()
        st.rerun()

    st.markdown(
        '<div class="section-heading">'
        '03 | Implementation Challenge'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            Select the challenge that affects the implementation
            of the chosen Blue Economy practice.
            The available challenges are filtered using both
            previous selections.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="selection-panel">
            <div class="small-label">Selected marine construction impact</div>
            <div class="selection-text">
                {html.escape(impact)}
            </div>
            <br>
            <div class="small-label">Selected Blue Economy practice</div>
            <div class="selection-text">
                {html.escape(practice)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    practice_data = df[
        (df[COL_IMPACT] == impact)
        & (df[COL_PRACTICE] == practice)
    ]

    challenges = (
        practice_data[COL_CHALLENGE]
        .drop_duplicates()
        .tolist()
    )

    search_challenge = st.text_input(
        "Search implementation challenges",
        placeholder="Enter a challenge or keyword..."
    )

    filtered_challenges = [
        challenge for challenge in challenges
        if search_challenge.lower() in challenge.lower()
    ]

    st.caption(
        f"{len(filtered_challenges)} relevant challenges identified"
    )

    if not filtered_challenges:

        st.warning("No matching implementation challenges found.")

    else:

        selected = st.selectbox(
            "Available implementation challenges",
            options=filtered_challenges,
            index=None,
            placeholder="Select an implementation challenge"
        )

        if selected:

            challenge_data = practice_data[
                practice_data[COL_CHALLENGE] == selected
            ]

            lean_count = (
                challenge_data[COL_LEAN]
                .drop_duplicates()
                .shape[0]
            )

            st.markdown(
                f"""
                <div class="selection-panel">
                    <div class="selection-title">
                        Selected implementation challenge
                    </div>
                    <div class="selection-text">
                        {html.escape(selected)}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.info(
                f"{lean_count} unique Lean principle records "
                "are associated with this selection."
            )

            if st.button(
                "Generate Lean Principle Recommendations",
                type="primary",
                use_container_width=True
            ):

                st.session_state.selected_challenge = selected
                st.session_state.current_page = 4

                st.rerun()

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        if st.button("Back to Blue Economy Practices"):
            st.session_state.current_page = 2
            st.rerun()

    with col2:

        if st.button("Start Again"):
            reset_application()
            st.rerun()


# ============================================================
# PAGE 4: LEAN PRINCIPLE RECOMMENDATIONS
# ============================================================

elif page == 4:

    impact = st.session_state.selected_impact
    practice = st.session_state.selected_practice
    challenge = st.session_state.selected_challenge

    if not impact or not practice or not challenge:

        reset_application()
        st.rerun()

    st.markdown(
        '<div class="section-heading">'
        '04 | Lean Principle Recommendations'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-description">
            The following Lean principles are associated with
            the selected marine construction impact, Blue Economy
            practice and implementation challenge.
            The displayed sequence follows the order of records
            in the original Excel dataset.
        </div>
        """,
        unsafe_allow_html=True
    )

    # Filter using all three selections
    result = df[
        (df[COL_IMPACT] == impact)
        & (df[COL_PRACTICE] == practice)
        & (df[COL_CHALLENGE] == challenge)
    ].copy()

    # Preserve Excel sequence and remove duplicate Lean principles
    result = result.drop_duplicates(
        subset=[COL_LEAN],
        keep="first"
    ).reset_index(drop=True)

    lean_principles = result[COL_LEAN].tolist()

    # Summary of selected criteria
    st.markdown(
        f"""
        <div class="selection-panel">
            <div class="small-label">Marine construction impact</div>
            <div class="selection-text">
                {html.escape(impact)}
            </div>
            <br>
            <div class="small-label">Blue Economy practice</div>
            <div class="selection-text">
                {html.escape(practice)}
            </div>
            <br>
            <div class="small-label">Implementation challenge</div>
            <div class="selection-text">
                {html.escape(challenge)}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-heading">'
        'Identified Lean Principles'
        '</div>',
        unsafe_allow_html=True
    )

    st.caption(
        "Principles are presented in their original dataset order."
    )

    if len(lean_principles) == 5:

        st.success(
            "Five unique Lean principles have been identified "
            "for the selected decision pathway."
        )

    else:

        st.warning(
            f"The dataset contains {len(lean_principles)} "
            "unique Lean principle(s) for this selection, "
            "rather than five. All available records are displayed "
            "without generating additional recommendations."
        )

    if not lean_principles:

        st.error(
            "No Lean principles were identified for this selection."
        )

    else:

        for index, principle in enumerate(
            lean_principles,
            start=1
        ):

            st.markdown(
                f"""
                <div class="result-panel">
                    <span class="result-number">
                        {index:02d}
                    </span>
                    <span class="result-name">
                        {html.escape(principle)}
                    </span>
                </div>
                """,
                unsafe_allow_html=True
            )

    # ========================================================
    # DECISION SUMMARY
    # ========================================================

    st.markdown(
        '<div class="section-heading">'
        'Decision Summary'
        '</div>',
        unsafe_allow_html=True
    )

    summary_data = pd.DataFrame(
        [
            {
                "Decision criterion": "Marine construction impact",
                "Selected option": impact
            },
            {
                "Decision criterion": "Blue Economy practice",
                "Selected option": practice
            },
            {
                "Decision criterion": "Implementation challenge",
                "Selected option": challenge
            },
            {
                "Decision criterion": "Number of Lean principles",
                "Selected option": str(len(lean_principles))
            }
        ]
    )

    st.dataframe(
        summary_data,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # EXPORT RESULTS
    # ========================================================

    st.markdown(
        '<div class="section-heading">'
        'Export Decision Results'
        '</div>',
        unsafe_allow_html=True
    )

    if not result.empty:

        export_data = result.copy()

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
            label="Download Decision Results (CSV)",
            data=csv_data,
            file_name="bluepath_decision_results.csv",
            mime="text/csv",
            use_container_width=True
        )

    st.divider()

    # ========================================================
    # NAVIGATION
    # ========================================================

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button(
            "Back to Challenges",
            use_container_width=True
        ):

            st.session_state.current_page = 3
            st.rerun()

    with col2:

        if st.button(
            "New Decision",
            type="primary",
            use_container_width=True
        ):

            reset_application()
            st.rerun()

    with col3:

        if st.button(
            "Change Blue Economy Practice",
            use_container_width=True
        ):

            st.session_state.selected_challenge = None
            st.session_state.current_page = 2
            st.rerun()


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="app-footer">
        BLUEPATH | Blue Economy and Lean Construction
        Decision Support System
        <br>
        Research-based decision support for marine infrastructure projects
    </div>
    """,
    unsafe_allow_html=True
)
