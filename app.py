import html
import io
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Safe import for ReportLab (PDF Generation)
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# ============================================================
# BLUEPATH: MARINE CONSTRUCTION DECISION SUPPORT PLATFORM
# Blue Economy × Lean Construction
# Enhanced Multi-Page Decision Support System (v2 - Cloud Resilient)
# ============================================================

st.set_page_config(
    page_title="BluePath | Marine Construction Decision Support",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------------------------------------------
# CONFIGURATION & DATA SOURCE
# ------------------------------------------------------------
BASE_DIR = Path(__file__).parent if "__file__" in globals() else Path.cwd()
EXCEL_CANDIDATES = [BASE_DIR / "portal.xlsx", BASE_DIR / "Portal.xlsx"]
EXCEL_FILE = next((f for f in EXCEL_CANDIDATES if f.exists()), EXCEL_CANDIDATES[0])

IMPACT_COL = "Impact of Marine Construction"
PRACTICE_COL = "Blue Economy Practice"
CHALLENGE_COL = "Challenges to Implement Blue Economy Practice"
LEAN_COL = "Lean Principles"
REQUIRED_COLUMNS = [IMPACT_COL, PRACTICE_COL, CHALLENGE_COL, LEAN_COL]

EXPECTED_LEAN_COUNT = 5

# ------------------------------------------------------------
# MOCK DATASET FALLBACK (If Excel File is Missing)
# ------------------------------------------------------------
MOCK_DATA = [
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "High upfront material cost & specialized installation equipment",
        LEAN_COL: "Eliminate Waste"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "High upfront material cost & specialized installation equipment",
        LEAN_COL: "Optimize the Whole"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "High upfront material cost & specialized installation equipment",
        LEAN_COL: "Build Quality In"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "High upfront material cost & specialized installation equipment",
        LEAN_COL: "Create Continuous Flow"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "High upfront material cost & specialized installation equipment",
        LEAN_COL: "Respect People & Marine Environment"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "Regulatory approval delays & lack of standardized eco-testing",
        LEAN_COL: "Focus on Customer/Stakeholder Value"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Living Shorelines & Mangrove/Seagrass Restoration",
        CHALLENGE_COL: "High early-stage mortality under extreme storm events",
        LEAN_COL: "Build Quality In"
    },
    {
        IMPACT_COL: "Turbidity & Sediment Resuspension",
        PRACTICE_COL: "Real-Time Turbidity Monitoring & Smart Silt Curtains",
        CHALLENGE_COL: "Sensor fouling & telemetry connectivity in offshore conditions",
        LEAN_COL: "Visual Management"
    },
    {
        IMPACT_COL: "Turbidity & Sediment Resuspension",
        PRACTICE_COL: "Closed-Bucket Biodegradable Dredging Systems",
        CHALLENGE_COL: "Lower production rate compared to conventional open buckets",
        LEAN_COL: "Eliminate Waste"
    },
    {
        IMPACT_COL: "Noise Pollution & Underwater Acoustics",
        PRACTICE_COL: "Bubble Curtains & Hydro-Sound Dampers for Pile Driving",
        CHALLENGE_COL: "High compressed air energy demand and vessel footprint",
        LEAN_COL: "Optimize Logistics & Layout"
    }
]

# ------------------------------------------------------------
# CUSTOM CSS & VISUAL DESIGN SYSTEM
# Inspired by Analytics Dashboard, Risk Heatmap & Modern UI
# ------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', system-ui, -apple-system, sans-serif;
}

.stApp {
    background: linear-gradient(135deg, #F0F6F7 0%, #F7FAFA 50%, #EEF4F5 100%);
    color: #1F3A44;
}

#MainMenu, footer, header { visibility: hidden; }

.block-container {
    max-width: 1200px;
    padding-top: 1.2rem;
    padding-bottom: 2.5rem;
}

h1, h2, h3, h4 {
    font-family: 'Inter', sans-serif;
    color: #0A5C63;
    font-weight: 750;
    letter-spacing: -0.4px;
}

/* Header & Navbar */
.brand-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 14px 22px;
    background: #FFFFFF;
    border-radius: 14px;
    border: 1px solid #D6E6E8;
    box-shadow: 0 4px 14px rgba(10, 110, 116, 0.05);
    margin-bottom: 20px;
}

.brand-title {
    font-size: 1.6rem;
    font-weight: 800;
    color: #0A6E74;
    letter-spacing: -0.5px;
    display: flex;
    align-items: center;
    gap: 8px;
}

.brand-subtitle {
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 2px;
    color: #5A7A84;
    text-transform: uppercase;
}

/* Welcome Hero */
.welcome-hero {
    background: linear-gradient(125deg, #095D64 0%, #0D8A8F 50%, #2BA88A 100%);
    border-radius: 18px;
    padding: 36px 36px;
    color: white;
    box-shadow: 0 12px 30px rgba(10, 110, 116, 0.18);
    position: relative;
    overflow: hidden;
    margin-bottom: 24px;
}

.welcome-hero h1 {
    font-size: 2.2rem;
    font-weight: 800;
    color: #FFFFFF;
    margin-top: 6px;
    margin-bottom: 12px;
    line-height: 1.2;
}

.welcome-hero p {
    font-size: 1.02rem;
    line-height: 1.65;
    color: #E2F5F5;
    max-width: 780px;
    margin: 0;
}

.badge-tag {
    display: inline-block;
    background: rgba(255, 255, 255, 0.18);
    backdrop-filter: blur(8px);
    border: 1px solid rgba(255, 255, 255, 0.3);
    padding: 5px 12px;
    border-radius: 30px;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: #FFFFFF;
}

/* KPI Cards */
.kpi-container {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
}

.kpi-card {
    background: #FFFFFF;
    border: 1px solid #D6E6E8;
    border-radius: 14px;
    padding: 20px 18px;
    box-shadow: 0 4px 14px rgba(30, 80, 90, 0.04);
    position: relative;
    transition: all 0.25s ease;
}

.kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 22px rgba(10, 110, 116, 0.12);
    border-color: #0A6E74;
}

.kpi-label {
    font-size: 0.72rem;
    font-weight: 750;
    letter-spacing: 1.2px;
    color: #6A8790;
    text-transform: uppercase;
    margin-bottom: 6px;
}

.kpi-value {
    font-size: 1.85rem;
    font-weight: 800;
    color: #095D64;
    line-height: 1.1;
}

.kpi-badge {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    font-size: 0.75rem;
    font-weight: 700;
    padding: 3px 8px;
    border-radius: 20px;
    margin-top: 8px;
}

.kpi-badge-positive { background: #E6F7F0; color: #1E7A4A; }
.kpi-badge-purple { background: #F0EBFB; color: #6C5CE7; }

/* Step Cards */
.step-card {
    background: #FFFFFF;
    border: 1px solid #D6E6E8;
    border-left: 5px solid #0A6E74;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 14px;
    box-shadow: 0 3px 12px rgba(30, 80, 90, 0.04);
    transition: all 0.22s ease;
}

.step-card:hover {
    transform: translateX(4px);
    border-left-color: #E05A5A;
    box-shadow: 0 6px 18px rgba(224, 90, 90, 0.12);
}

.number-badge {
    display: inline-flex;
    justify-content: center;
    align-items: center;
    width: 32px;
    height: 32px;
    border-radius: 50%;
    background: #E0F2F1;
    color: #0A6E74;
    font-weight: 800;
    font-size: 0.88rem;
    margin-right: 10px;
}

.lean-card {
    background: linear-gradient(120deg, #F0F9F5 0%, #F8FCFA 100%);
    border: 1px solid #C4E5D3;
    border-left: 6px solid #1E7A4A;
    border-radius: 12px;
    padding: 18px 20px;
    margin-bottom: 14px;
    box-shadow: 0 3px 10px rgba(30, 120, 70, 0.05);
}

.priority-rank {
    font-size: 0.70rem;
    font-weight: 800;
    letter-spacing: 1.5px;
    color: #1E7A4A;
    text-transform: uppercase;
    margin-bottom: 4px;
}

.notice-box {
    background: linear-gradient(135deg, #E6F4F4 0%, #EDF7F7 100%);
    border-left: 5px solid #0D8A8F;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 20px;
    color: #1E3A42;
    font-size: 0.95rem;
    line-height: 1.6;
}

.custom-footer {
    text-align: center;
    color: #6A8790;
    font-size: 0.82rem;
    border-top: 1px solid #D8E8EA;
    padding-top: 20px;
    margin-top: 40px;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# DATA LOADING FUNCTION
# ------------------------------------------------------------
@st.cache_data
def load_database():
    if EXCEL_FILE.exists():
        try:
            data = pd.read_excel(EXCEL_FILE, engine="openpyxl")
            data.columns = data.columns.astype(str).str.strip()
            missing = [c for c in REQUIRED_COLUMNS if c not in data.columns]
            if not missing:
                data = data[REQUIRED_COLUMNS].copy()
                for col in REQUIRED_COLUMNS:
                    data[col] = data[col].fillna("").astype(str).str.strip()
                data = data.replace("", pd.NA).dropna(subset=REQUIRED_COLUMNS)
                if not data.empty:
                    return data.reset_index(drop=True)
        except Exception:
            pass
    return pd.DataFrame(MOCK_DATA).reset_index(drop=True)

df = load_database()

# ------------------------------------------------------------
# SESSION STATE INITIALIZATION
# ------------------------------------------------------------
DEFAULTS = {
    "page": "welcome",
    "username": "",
    "organization": "",
    "user_role": "Marine Contractor",
    "selected_impact": None,
    "selected_practice": None,
    "selected_challenge": None,
    "feedback_list": []
}

for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_pathway():
    st.session_state.selected_impact = None
    st.session_state.selected_practice = None
    st.session_state.selected_challenge = None

# ------------------------------------------------------------
# SIDEBAR NAVIGATION
# ------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🌊 BluePath Navigation")
    st.caption("Marine Construction Decision Support")
    st.write("")
    
    page_selection = st.radio(
        "Select Portal Module:",
        options=[
            "🏠 Overview & Welcome",
            "🗺️ Decision Pathway (4 Steps)",
            "📊 Total Impact Analytics & Heatmap",
            "📄 Report Generation Center",
            "⭐ Feedback & Evaluation"
        ],
        index=0 if st.session_state.page == "welcome"
        else 1 if st.session_state.page in ["impacts", "practices", "challenges", "lean"]
        else 2 if st.session_state.page == "analytics"
        else 3 if st.session_state.page == "reports"
        else 4
    )
    
    st.divider()
    
    if "Overview" in page_selection:
        st.session_state.page = "welcome"
    elif "Decision Pathway" in page_selection:
        if st.session_state.page in ["welcome", "analytics", "reports", "feedback"]:
            st.session_state.page = "impacts" if not st.session_state.selected_impact else st.session_state.page
    elif "Analytics" in page_selection:
        st.session_state.page = "analytics"
    elif "Report" in page_selection:
        st.session_state.page = "reports"
    elif "Feedback" in page_selection:
        st.session_state.page = "feedback"

    st.markdown("#### 📌 Current Selection")
    if st.session_state.selected_impact:
        st.markdown(f"**Impact:** `{st.session_state.selected_impact}`")
    else:
        st.caption("No impact selected yet.")
        
    if st.button("🔄 Reset Pathway", use_container_width=True):
        reset_pathway()
        st.session_state.page = "impacts"
        st.rerun()

# ------------------------------------------------------------
# BRAND HEADER
# ------------------------------------------------------------
st.markdown("""
<div class="brand-header">
    <div>
        <div class="brand-title">🌊 BluePath</div>
        <div class="brand-subtitle">Blue Economy × Lean Construction Platform</div>
    </div>
    <div style="text-align: right;">
        <span class="badge-tag" style="background:#0A6E74; color:white;">Research Edition</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# MODULE 1: WELCOME & OVERVIEW
# ============================================================
if st.session_state.page == "welcome":
    st.markdown("""
    <div class="welcome-hero">
        <span class="badge-tag">Welcome to BluePath</span>
        <h1>We're glad you're here!</h1>
        <p>
            Connect marine construction environmental impacts with Blue Economy practices
            and prioritised Lean principles to support sustainable project implementation.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    total_impacts = df[IMPACT_COL].nunique()
    total_practices = df[PRACTICE_COL].nunique()
    total_challenges = df[CHALLENGE_COL].nunique()
    total_lean = df[LEAN_COL].nunique()
    
    st.markdown("""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-label">Marine Construction Impacts</div>
            <div class="kpi-value">{impacts}</div>
            <div class="kpi-badge kpi-badge-positive">✓ Mapped</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Blue Economy Practices</div>
            <div class="kpi-value">{practices}</div>
            <div class="kpi-badge kpi-badge-purple">⚡ Sustainable Solutions</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Implementation Challenges</div>
            <div class="kpi-value">{challenges}</div>
            <div class="kpi-badge kpi-badge-positive">🔍 Risk Evaluated</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Lean Principles Mapped</div>
            <div class="kpi-value">{lean}</div>
            <div class="kpi-badge kpi-badge-purple">🎯 Priority Ranked</div>
        </div>
    </div>
    """.format(impacts=total_impacts, practices=total_practices, challenges=total_challenges, lean=total_lean), unsafe_allow_html=True)
    
    col_left, col_right = st.columns([1.5, 1])
    with col_left:
        st.markdown("### 🚀 Decision Pathway Steps")
        steps_info = [
            ("01", "Select Marine Construction Impact", "Choose environmental stressors to address."),
            ("02", "Explore Blue Economy Practices", "View practices linked to the impact."),
            ("03", "Identify Implementation Challenges", "Examine operational & regulatory bottlenecks."),
            ("04", "Deploy Prioritised Lean Principles", "Access 5 Lean principles ranked by priority.")
        ]
        for num, title, desc in steps_info:
            st.markdown(f"""
            <div class="step-card">
                <span class="number-badge">{num}</span>
                <strong>{title}</strong>
                <p style="color:#4A6A74; margin:6px 0 0 42px; font-size:0.92rem; line-height:1.5;">{desc}</p>
            </div>
            """, unsafe_allow_html=True)
            
    with col_right:
        st.markdown("### 👤 User Onboarding")
        with st.form("onboarding_form"):
            user_name = st.text_input("Your Name / ID", value=st.session_state.username, placeholder="e.g. Dr. Alex Morgan")
            org = st.text_input("Organization", value=st.session_state.organization, placeholder="e.g. Coastal Infrastructure Ltd")
            role = st.selectbox("Role", ["Marine Contractor", "Environmental Engineer", "Lean Consultant", "Project Manager"])
            if st.form_submit_button("Start Decision Pathway →", type="primary", use_container_width=True):
                st.session_state.username = user_name.strip()
                st.session_state.organization = org.strip()
                st.session_state.user_role = role
                st.session_state.page = "impacts"
                st.rerun()

# ============================================================
# MODULE 2: DECISION PATHWAY (STEPS 1 - 4)
# ============================================================
elif st.session_state.page in ["impacts", "practices", "challenges", "lean"]:
    PAGES_MAP = {"impacts": 1, "practices": 2, "challenges": 3, "lean": 4}
    current_step = PAGES_MAP[st.session_state.page]
    
    st.progress(current_step / 4)
    st.caption(f"Decision Pathway → Step {current_step} of 4")
    
    if st.session_state.page == "impacts":
        st.markdown("## Step 1: Select Marine Construction Impact")
        impacts = df[IMPACT_COL].drop_duplicates().tolist()
        cols = st.columns(2)
        for i, impact in enumerate(impacts):
            with cols[i % 2]:
                st.markdown(f"""
                <div class="step-card">
                    <span class="number-badge">{i+1}</span>
                    <strong style="font-size:1.05rem;">{html.escape(impact)}</strong>
                </div>
                """, unsafe_allow_html=True)
                if st.button(f"Explore Impact {i+1} →", key=f"btn_imp_{i}", use_container_width=True):
                    st.session_state.selected_impact = impact
                    st.session_state.selected_practice = None
                    st.session_state.selected_challenge = None
                    st.session_state.page = "practices"
                    st.rerun()
                    
    elif st.session_state.page == "practices":
        impact = st.session_state.selected_impact
        st.markdown(f"## Step 2: Select Blue Economy Practice")
        st.markdown(f"""
        <div class="notice-box">
            <strong>Selected Impact:</strong> {html.escape(str(impact))}
        </div>
        """, unsafe_allow_html=True)
        practices = df.loc[df[IMPACT_COL] == impact, PRACTICE_COL].drop_duplicates().tolist()
        for i, practice in enumerate(practices):
            st.markdown(f"""
            <div class="step-card">
                <span class="number-badge">{i+1}</span>
                <strong>{html.escape(practice)}</strong>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"Select Practice {i+1} →", key=f"btn_prac_{i}", use_container_width=True):
                st.session_state.selected_practice = practice
                st.session_state.selected_challenge = None
                st.session_state.page = "challenges"
                st.rerun()
                
    elif st.session_state.page == "challenges":
        impact = st.session_state.selected_impact
        practice = st.session_state.selected_practice
        st.markdown("## Step 3: Identify Implementation Challenge")
        filtered = df[(df[IMPACT_COL] == impact) & (df[PRACTICE_COL] == practice)]
        challenges = filtered[CHALLENGE_COL].drop_duplicates().tolist()
        for i, challenge in enumerate(challenges):
            st.markdown(f"""
            <div class="step-card" style="border-left-color: #E67E22;">
                <span class="number-badge" style="background:#FDEBD0; color:#E67E22;">{i+1}</span>
                <strong>{html.escape(challenge)}</strong>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"View Prioritised Lean Principles {i+1} →", key=f"btn_chal_{i}", use_container_width=True):
                st.session_state.selected_challenge = challenge
                st.session_state.page = "lean"
                st.rerun()
                
    elif st.session_state.page == "lean":
        impact = st.session_state.selected_impact
        practice = st.session_state.selected_practice
        challenge = st.session_state.selected_challenge
        st.markdown("## Step 4: Prioritised Lean Principles")
        filtered = df[(df[IMPACT_COL] == impact) & (df[PRACTICE_COL] == practice) & (df[CHALLENGE_COL] == challenge)]
        lean_principles = filtered[LEAN_COL].tolist()
        
        for idx, principle in enumerate(lean_principles, start=1):
            st.markdown(f"""
            <div class="lean-card">
                <div class="priority-rank">Priority Rank #{idx}</div>
                <strong style="font-size:1.1rem; color:#095D64;">{html.escape(principle)}</strong>
            </div>
            """, unsafe_allow_html=True)
            
        col_actions = st.columns(2)
        with col_actions:
            if st.button("📊 View Total Impact Analytics & Summary", type="primary", use_container_width=True):
                st.session_state.page = "analytics"
                st.rerun()
        with col_actions:
            if st.button("📄 Export Pathway Report", use_container_width=True):
                st.session_state.page = "reports"
                st.rerun()

# ============================================================
# MODULE 3: TOTAL IMPACT ANALYTICS & HEATMAP
# ============================================================
elif st.session_state.page == "analytics":
    st.markdown("## 📊 Total Impact Analytics & Decision Heatmap")
    impact_options = df[IMPACT_COL].unique().tolist()
    default_idx = impact_options.index(st.session_state.selected_impact) if st.session_state.selected_impact in impact_options else 0
    selected_analytics_impact = st.selectbox("Select Marine Construction Impact to Evaluate:", options=impact_options, index=default_idx)
    impact_df = df[df[IMPACT_COL] == selected_analytics_impact]
    
    st.markdown("### 🔥 Challenge vs. Lean Principle Priority Decision Matrix")
    ct = pd.crosstab(impact_df[CHALLENGE_COL], impact_df[LEAN_COL])
    
    fig_heatmap = px.imshow(
        ct,
        labels=dict(x="Lean Principle", y="Implementation Challenge", color="Frequency"),
        x=ct.columns, y=ct.index,
        color_continuous_scale=[[0.0, "#EBF5FB"], [0.25, "#2BA88A"], [0.5, "#F4B41A"], [0.75, "#E67E22"], [1.0, "#E74C3C"]],
        aspect="auto"
    )
    fig_heatmap.update_layout(font_family="Inter", height=400, margin=dict(l=40, r=40, t=40, b=40))
    st.plotly_chart(fig_heatmap, use_container_width=True)
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown("#### 📈 Challenges per Practice")
        chal_by_prac = impact_df.groupby(PRACTICE_COL)[CHALLENGE_COL].nunique().reset_index()
        fig_bar = px.bar(chal_by_prac, x=CHALLENGE_COL, y=PRACTICE_COL, orientation="h", color_discrete_sequence=["#0A6E74"])
        fig_bar.update_layout(height=300, font_family="Inter")
        st.plotly_chart(fig_bar, use_container_width=True)
    with col_c2:
        st.markdown("#### 🍩 Lean Distribution")
        lean_counts = impact_df[LEAN_COL].value_counts().reset_index()
        lean_counts.columns = ["Lean Principle", "Count"]
        fig_donut = px.pie(lean_counts, values="Count", names="Lean Principle", hole=0.45, color_discrete_sequence=["#0A6E74", "#2BA88A", "#6C5CE7", "#F4B41A"])
        fig_donut.update_layout(height=300, font_family="Inter")
        st.plotly_chart(fig_donut, use_container_width=True)

# ============================================================
# MODULE 4 & 5: REPORTS & FEEDBACK
# ============================================================
elif st.session_state.page == "reports":
    st.markdown("## 📄 Report Generation Center")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.markdown("### 📥 Executive PDF Report")
        
        def generate_pdf():
            if not REPORTLAB_AVAILABLE:
                return None
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
            styles = getSampleStyleSheet()
            elements = [
                Paragraph("BluePath Executive Decision Report", styles['Heading1']),
                HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0A6E74'), spaceAfter=15),
                Paragraph("Mapped Decision Matrix Summary", styles['Heading2'])
            ]
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        pdf_bytes = generate_pdf()
        if pdf_bytes:
            st.download_button("📥 Download PDF Report", data=pdf_bytes, file_name="BluePath_Report.pdf", mime="application/pdf", type="primary", use_container_width=True)
        else:
            st.warning("⚠️ `reportlab` library is not installed in your Python environment.")
            st.info("To enable PDF generation on Streamlit Cloud, add `reportlab` to your repository's `requirements.txt` file.")
            
    with col_r2:
        st.markdown("### 📊 Complete Excel Export")
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='BluePath_Matrix')
        buffer_excel.seek(0)
        st.download_button("📊 Download Excel Dataset", data=buffer_excel.getvalue(), file_name="BluePath_Matrix.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

elif st.session_state.page == "feedback":
    st.markdown("## ⭐ Application Feedback & Evaluation")
    with st.form("feedback_form"):
        rating = st.slider("Rating (1-5)", 1, 5, 5)
        comments = st.text_area("Your Comments / Suggestions")
        if st.form_submit_button("Submit Feedback", type="primary"):
            st.session_state.feedback_list.append({"rating": rating, "comments": comments})
            st.success("Thank you! Feedback recorded.")

st.markdown("""
<div class="custom-footer">
    <strong>BluePath</strong> · Marine Construction Decision Support Platform<br>
    Integrating Blue Economy Practices with Prioritised Lean Construction Principles
</div>
""", unsafe_allow_html=True)
