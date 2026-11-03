import io
import html
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ============================================================
# BLUEPATH: MARINE CONSTRUCTION DECISION SUPPORT PLATFORM
# Blue Economy × Lean Construction
# Enhanced Multi-Page Decision Support System
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
    # Impact 1: Habitat Destruction & Coastal Erosion
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
        PRACTICE_COL: "Eco-Engineered Concrete Armor Units (Artificial Reefs)",
        CHALLENGE_COL: "Regulatory approval delays & lack of standardized eco-testing",
        LEAN_COL: "Create Continuous Flow"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Living Shorelines & Mangrove/Seagrass Restoration",
        CHALLENGE_COL: "High early-stage mortality under extreme storm events",
        LEAN_COL: "Build Quality In"
    },
    {
        IMPACT_COL: "Habitat Destruction & Coastal Erosion",
        PRACTICE_COL: "Living Shorelines & Mangrove/Seagrass Restoration",
        CHALLENGE_COL: "High early-stage mortality under extreme storm events",
        LEAN_COL: "Continuous Improvement (Kaizen)"
    },
    # Impact 2: Turbidity & Sediment Resuspension
    {
        IMPACT_COL: "Turbidity & Sediment Resuspension",
        PRACTICE_COL: "Real-Time Turbidity Monitoring & Smart Silt Curtains",
        CHALLENGE_COL: "Sensor fouling & telemetry connectivity in offshore conditions",
        LEAN_COL: "Visual Management"
    },
    {
        IMPACT_COL: "Turbidity & Sediment Resuspension",
        PRACTICE_COL: "Real-Time Turbidity Monitoring & Smart Silt Curtains",
        CHALLENGE_COL: "Sensor fouling & telemetry connectivity in offshore conditions",
        LEAN_COL: "Build Quality In"
    },
    {
        IMPACT_COL: "Turbidity & Sediment Resuspension",
        PRACTICE_COL: "Closed-Bucket Biodegradable Dredging Systems",
        CHALLENGE_COL: "Lower production rate compared to conventional open buckets",
        LEAN_COL: "Eliminate Waste"
    },
    {
        IMPACT_COL: "Turbidity & Sediment Resuspension",
        PRACTICE_COL: "Closed-Bucket Biodegradable Dredging Systems",
        CHALLENGE_COL: "Lower production rate compared to conventional open buckets",
        LEAN_COL: "Optimize the Whole"
    },
    # Impact 3: Noise Pollution & Underwater Acoustics
    {
        IMPACT_COL: "Noise Pollution & Underwater Acoustics",
        PRACTICE_COL: "Bubble Curtains & Hydro-Sound Dampers for Pile Driving",
        CHALLENGE_COL: "High compressed air energy demand and vessel footprint",
        LEAN_COL: "Eliminate Energy Waste"
    },
    {
        IMPACT_COL: "Noise Pollution & Underwater Acoustics",
        PRACTICE_COL: "Bubble Curtains & Hydro-Sound Dampers for Pile Driving",
        CHALLENGE_COL: "High compressed air energy demand and vessel footprint",
        LEAN_COL: "Optimize Logistics & Layout"
    },
    # Impact 4: Carbon Emissions & Fuel Consumption
    {
        IMPACT_COL: "Carbon Emissions & Fuel Consumption",
        PRACTICE_COL: "Hybrid-Electric Offshore Workboats & Battery Energy Storage",
        CHALLENGE_COL: "Limited shore power charging infrastructure at remote marine bases",
        LEAN_COL: "Create Continuous Flow"
    },
    {
        IMPACT_COL: "Carbon Emissions & Fuel Consumption",
        PRACTICE_COL: "Hybrid-Electric Offshore Workboats & Battery Energy Storage",
        CHALLENGE_COL: "Limited shore power charging infrastructure at remote marine bases",
        LEAN_COL: "Optimize the Whole"
    }
]

# ------------------------------------------------------------
# CUSTOM CSS & VISUAL DESIGN SYSTEM
# Inspired by Analytics Dashboard, Risk Heatmap & Modern UI
# ------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ---------- Base Setup ---------- */
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

/* ---------- Typography Hierarchy ---------- */
h1, h2, h3, h4 {
    font-family: 'Inter', sans-serif;
    color: #0A5C63;
    font-weight: 750;
    letter-spacing: -0.4px;
}

/* ---------- Top Navbar / Brand Header ---------- */
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

/* ---------- Welcome Banner (Inspired by Image 3) ---------- */
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

/* ---------- KPI Metric Cards (Inspired by Analytics Dashboard Image 2) ---------- */
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

.kpi-badge-positive {
    background: #E6F7F0;
    color: #1E7A4A;
}

.kpi-badge-purple {
    background: #F0EBFB;
    color: #6C5CE7;
}

/* ---------- Step Cards & Info Cards ---------- */
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

/* ---------- Lean Priority Cards ---------- */
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

/* ---------- Notice Box ---------- */
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

/* ---------- Button Polish ---------- */
.stButton > button {
    border-radius: 10px;
    font-weight: 650;
    transition: all 0.2s ease;
    min-height: 44px;
}

/* ---------- Footer ---------- */
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
    # Fallback to mock dataset if Excel not available or invalid
    df_mock = pd.DataFrame(MOCK_DATA)
    return df_mock.reset_index(drop=True)

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
    st.caption("Marine Construction Decision Support Platform")
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
    
    # Quick Navigation Mapper
    if "Overview" in page_selection:
        st.session_state.page = "welcome"
    elif "Decision Pathway" in page_selection:
        if st.session_state.page == "welcome" or st.session_state.page in ["analytics", "reports", "feedback"]:
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
        
    if st.session_state.selected_practice:
        st.markdown(f"**Practice:** `{st.session_state.selected_practice}`")
    if st.session_state.selected_challenge:
        st.markdown(f"**Challenge:** `{st.session_state.selected_challenge}`")
        
    if st.button("🔄 Reset Decision Pathway", use_container_width=True):
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
        <div class="brand-subtitle">Blue Economy × Lean Construction Decision Platform</div>
    </div>
    <div style="text-align: right;">
        <span class="badge-tag" style="background:#0A6E74; color:white;">Research Edition v2.5</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# MODULE 1: WELCOME & OVERVIEW (Inspired by Image 3)
# ============================================================
if st.session_state.page == "welcome":
    st.markdown("""
    <div class="welcome-hero">
        <span class="badge-tag">Welcome to BluePath</span>
        <h1>We're glad you're here!</h1>
        <p>
            BluePath empowers marine construction decision-makers to systematically connect marine environmental impacts
            with actionable Blue Economy practices and prioritised Lean principles. Mitigate ecological risk while optimizing construction efficiency.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # KPI Metrics Bar (Inspired by Image 2 - Analytics Dashboard)
    total_impacts = df[IMPACT_COL].nunique()
    total_practices = df[PRACTICE_COL].nunique()
    total_challenges = df[CHALLENGE_COL].nunique()
    total_lean = df[LEAN_COL].nunique()
    
    st.markdown("""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-label">Marine Construction Impacts</div>
            <div class="kpi-value">{impacts}</div>
            <div class="kpi-badge kpi-badge-positive">✓ Mapped in Database</div>
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
    """.format(
        impacts=total_impacts,
        practices=total_practices,
        challenges=total_challenges,
        lean=total_lean
    ), unsafe_allow_html=True)
    
    col_left, col_right = st.columns([1.5, 1])
    
    with col_left:
        st.markdown("### 🚀 Interactive Decision Pathway Overview")
        st.markdown("""
        Follow the 4-step framework to evaluate risk and identify top Lean interventions:
        """)
        
        steps_info = [
            ("01", "Select Marine Construction Impact", "Choose specific environmental stressors (e.g. Habitat Destruction, Turbidity, Acoustics)."),
            ("02", "Explore Blue Economy Practices", "View sustainable marine practices engineered to mitigate the selected impact."),
            ("03", "Identify Implementation Challenges", "Examine operational, financial, or regulatory bottlenecks."),
            ("04", "Deploy Prioritised Lean Principles", "Access the 5 Lean principles ranked by research database priority to ensure smooth execution.")
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
            user_name = st.text_input("Your Name / Researcher ID", value=st.session_state.username, placeholder="e.g. Dr. Alex Morgan")
            org = st.text_input("Organization / Project Name", value=st.session_state.organization, placeholder="e.g. Coastal Infrastructure Offshore")
            role = st.selectbox("Primary Role", ["Marine Contractor", "Environmental Engineer", "Lean Consultant", "Project Manager", "Policy Researcher"])
            
            submitted = st.form_submit_button("Start Decision Pathway →", type="primary", use_container_width=True)
            if submitted:
                st.session_state.username = user_name.strip()
                st.session_state.organization = org.strip()
                st.session_state.user_role = role
                st.session_state.page = "impacts"
                st.rerun()

# ============================================================
# MODULE 2: DECISION PATHWAY (STEPS 1 - 4)
# ============================================================
elif st.session_state.page in ["impacts", "practices", "challenges", "lean"]:
    
    # Progress Bar Calculation
    PAGES_MAP = {"impacts": 1, "practices": 2, "challenges": 3, "lean": 4}
    current_step = PAGES_MAP[st.session_state.page]
    
    st.progress(current_step / 4)
    st.caption(f"Decision Pathway → Step {current_step} of 4")
    
    # STEP 1: IMPACTS
    if st.session_state.page == "impacts":
        user_disp = html.escape(st.session_state.username) if st.session_state.username else "Researcher"
        st.markdown(f"## Step 1: Select Marine Construction Impact")
        st.markdown(f"Welcome, **{user_disp}**. Select the environmental impact topic to investigate:")
        
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
                    
    # STEP 2: PRACTICES
    elif st.session_state.page == "practices":
        impact = st.session_state.selected_impact
        if not impact:
            st.session_state.page = "impacts"
            st.rerun()
            
        st.markdown("## Step 2: Select Blue Economy Practice")
        st.markdown(f"""
        <div class="notice-box">
            <strong>Selected Marine Construction Impact:</strong><br>
            <span style="font-size:1.1rem; font-weight:700; color:#0A6E74;">{html.escape(impact)}</span>
        </div>
        """, unsafe_allow_html=True)
        
        practices = df.loc[df[IMPACT_COL] == impact, PRACTICE_COL].drop_duplicates().tolist()
        
        st.markdown(f"### Mapped Practices ({len(practices)} available)")
        for i, practice in enumerate(practices):
            st.markdown(f"""
            <div class="step-card">
                <span class="number-badge">{i+1}</span>
                <strong style="font-size:1.02rem;">{html.escape(practice)}</strong>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"Select Practice {i+1} →", key=f"btn_prac_{i}", use_container_width=True):
                st.session_state.selected_practice = practice
                st.session_state.selected_challenge = None
                st.session_state.page = "challenges"
                st.rerun()
                
        st.divider()
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("← Back to Impact Selection", use_container_width=True):
                st.session_state.page = "impacts"
                st.rerun()
                
    # STEP 3: CHALLENGES
    elif st.session_state.page == "challenges":
        impact = st.session_state.selected_impact
        practice = st.session_state.selected_practice
        if not impact or not practice:
            st.session_state.page = "impacts"
            st.rerun()
            
        st.markdown("## Step 3: Identify Implementation Challenge")
        st.markdown(f"""
        <div class="notice-box">
            <strong>Selected Impact:</strong> {html.escape(impact)}<br>
            <strong>Selected Practice:</strong> {html.escape(practice)}
        </div>
        """, unsafe_allow_html=True)
        
        filtered = df[(df[IMPACT_COL] == impact) & (df[PRACTICE_COL] == practice)]
        challenges = filtered[CHALLENGE_COL].drop_duplicates().tolist()
        
        for i, challenge in enumerate(challenges):
            st.markdown(f"""
            <div class="step-card" style="border-left-color: #E67E22;">
                <span class="number-badge" style="background:#FDEBD0; color:#E67E22;">{i+1}</span>
                <strong style="font-size:1.02rem;">{html.escape(challenge)}</strong>
            </div>
            """, unsafe_allow_html=True)
            if st.button(f"View Prioritised Lean Principles {i+1} →", key=f"btn_chal_{i}", use_container_width=True):
                st.session_state.selected_challenge = challenge
                st.session_state.page = "lean"
                st.rerun()
                
        st.divider()
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("← Back to Practice Selection", use_container_width=True):
                st.session_state.page = "practices"
                st.rerun()
                
    # STEP 4: PRIORITISED LEAN PRINCIPLES
    elif st.session_state.page == "lean":
        impact = st.session_state.selected_impact
        practice = st.session_state.selected_practice
        challenge = st.session_state.selected_challenge
        
        st.markdown("## Step 4: Prioritised Lean Principles")
        st.markdown(f"""
        <div class="notice-box">
            <strong>Impact:</strong> {html.escape(str(impact))}<br>
            <strong>Practice:</strong> {html.escape(str(practice))}<br>
            <strong>Challenge:</strong> {html.escape(str(challenge))}
        </div>
        """, unsafe_allow_html=True)
        
        filtered = df[
            (df[IMPACT_COL] == impact) &
            (df[PRACTICE_COL] == practice) &
            (df[CHALLENGE_COL] == challenge)
        ]
        
        lean_principles = filtered[LEAN_COL].tolist()
        
        if not lean_principles:
            st.warning("No specific Lean principles found for this exact path.")
        else:
            st.markdown(f"### Recommended Priority Interventions ({len(lean_principles)} principles)")
            for idx, principle in enumerate(lean_principles, start=1):
                st.markdown(f"""
                <div class="lean-card">
                    <div class="priority-rank">Priority Rank #{idx}</div>
                    <strong style="font-size:1.1rem; color:#095D64;">{html.escape(principle)}</strong>
                    <p style="margin:6px 0 0 0; color:#3A5A64; font-size:0.92rem;">
                        Targeted lean construction strategy to streamline workflow and mitigate eco-risk.
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
        st.divider()
        col_actions = st.columns(3)
        with col_actions[0]:
            if st.button("← Back to Challenges", use_container_width=True):
                st.session_state.page = "challenges"
                st.rerun()
        with col_actions[1]:
            if st.button("📊 View Full Impact Summary & Heatmap", type="primary", use_container_width=True):
                st.session_state.page = "analytics"
                st.rerun()
        with col_actions[2]:
            if st.button("📄 Export Pathway Report", use_container_width=True):
                st.session_state.page = "reports"
                st.rerun()

# ============================================================
# MODULE 3: TOTAL IMPACT ANALYTICS & HEATMAP (Inspired by Image 4)
# ============================================================
elif st.session_state.page == "analytics":
    st.markdown("## 📊 Total Impact Analytics & Decision Heatmap")
    st.markdown("Analyze the **entire selected marine impact** across all linked practices, challenges, and prioritised Lean principles.")
    
    impact_options = df[IMPACT_COL].unique().tolist()
    default_idx = impact_options.index(st.session_state.selected_impact) if st.session_state.selected_impact in impact_options else 0
    
    selected_analytics_impact = st.selectbox(
        "Select Marine Construction Impact to Evaluate:",
        options=impact_options,
        index=default_idx
    )
    
    impact_df = df[df[IMPACT_COL] == selected_analytics_impact]
    
    # Summary Metrics
    prac_cnt = impact_df[PRACTICE_COL].nunique()
    chal_cnt = impact_df[CHALLENGE_COL].nunique()
    lean_cnt = impact_df[LEAN_COL].nunique()
    total_mappings = len(impact_df)
    
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card">
            <div class="kpi-label">Selected Impact</div>
            <div class="kpi-value" style="font-size:1.2rem;">{selected_analytics_impact}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Total Practices</div>
            <div class="kpi-value">{prac_cnt}</div>
            <div class="kpi-badge kpi-badge-positive">Blue Economy Solutions</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Total Challenges</div>
            <div class="kpi-value">{chal_cnt}</div>
            <div class="kpi-badge kpi-badge-purple">Identified Bottlenecks</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Lean Interventions</div>
            <div class="kpi-value">{total_mappings}</div>
            <div class="kpi-badge kpi-badge-positive">Priority Alignments</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # ------------------------------------------------------------
    # DECISION HEATMAP MATRIX (Inspired by Image 4 Risk Heatmap)
    # Exact colors: Green #2BA88A, Yellow #F4B41A, Orange #E67E22, Red #E74C3C, Critical #C0392B
    # ------------------------------------------------------------
    st.markdown("### 🔥 Challenge vs. Lean Principle Priority Decision Matrix")
    st.caption("Visual risk & alignment density matrix for all combinations under this impact.")
    
    # Cross-tabulation of Challenge vs Lean Principle
    ct = pd.crosstab(impact_df[CHALLENGE_COL], impact_df[LEAN_COL])
    
    fig_heatmap = px.imshow(
        ct,
        labels=dict(x="Lean Principle", y="Implementation Challenge", color="Frequency"),
        x=ct.columns,
        y=ct.index,
        color_continuous_scale=[
            [0.0, "#EBF5FB"],
            [0.2, "#2BA88A"],  # Green
            [0.5, "#F4B41A"],  # Amber
            [0.8, "#E67E22"],  # Orange
            [1.0, "#E74C3C"]   # Red
        ],
        aspect="auto"
    )
    
    fig_heatmap.update_layout(
        font_family="Inter",
        title_font_size=16,
        margin=dict(l=40, r=40, t=40, b=40),
        height=400,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
    )
    
    st.plotly_chart(fig_heatmap, use_container_width=True)
    
    # ------------------------------------------------------------
    # CHARTS: Practices & Lean Distribution
    # ------------------------------------------------------------
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("#### 📈 Challenges per Blue Economy Practice")
        chal_by_prac = impact_df.groupby(PRACTICE_COL)[CHALLENGE_COL].nunique().reset_index()
        fig_bar = px.bar(
            chal_by_prac,
            x=CHALLENGE_COL,
            y=PRACTICE_COL,
            orientation="h",
            color_discrete_sequence=["#0A6E74"],
            labels={CHALLENGE_COL: "Number of Challenges", PRACTICE_COL: "Practice"}
        )
        fig_bar.update_layout(height=300, font_family="Inter", margin=dict(l=20, r=20, t=20, b=20))
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with col_c2:
        st.markdown("#### 🍩 Lean Principle Priority Distribution")
        lean_counts = impact_df[LEAN_COL].value_counts().reset_index()
        lean_counts.columns = ["Lean Principle", "Count"]
        fig_donut = px.pie(
            lean_counts,
            values="Count",
            names="Lean Principle",
            hole=0.45,
            color_discrete_sequence=["#0A6E74", "#2BA88A", "#6C5CE7", "#F4B41A", "#E67E22"]
        )
        fig_donut.update_layout(height=300, font_family="Inter", margin=dict(l=20, r=20, t=20, b=20))
        st.plotly_chart(fig_donut, use_container_width=True)
        
    # Process Flow Diagram
    st.markdown("### 🔄 Integrated Process Relationship Flow")
    st.info(f"**Impact:** {selected_analytics_impact} ➔ **{prac_cnt} Practices** ➔ **{chal_cnt} Challenges** ➔ **{total_mappings} Mapped Lean Priorities**")

# ============================================================
# MODULE 4: REPORT GENERATION CENTER
# ============================================================
elif st.session_state.page == "reports":
    st.markdown("## 📄 Report Generation Center")
    st.markdown("Generate executive summary reports in PDF or Excel format for stakeholders.")
    
    col_r1, col_r2 = st.columns(2)
    
    with col_r1:
        st.markdown("""
        <div class="kpi-card">
            <h3 style="margin-top:0;">📥 Export PDF Executive Report</h3>
            <p style="color:#5A7A84; font-size:0.92rem;">
                Includes selected pathway inputs, priority rankings, heatmap summary, and Lean recommendations.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        def generate_pdf():
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
            styles = getSampleStyleSheet()
            
            title_style = ParagraphStyle(
                'DocTitle',
                parent=styles['Heading1'],
                fontName='Helvetica-Bold',
                fontSize=20,
                leading=24,
                textColor=colors.HexColor('#0A6E74')
            )
            
            body_style = ParagraphStyle(
                'DocBody',
                parent=styles['Normal'],
                fontName='Helvetica',
                fontSize=10,
                leading=14,
                textColor=colors.HexColor('#2C3E50')
            )
            
            elements = []
            elements.append(Paragraph("BluePath Executive Decision Report", title_style))
            elements.append(Paragraph("Blue Economy × Lean Construction Decision Support", body_style))
            elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#0A6E74'), spaceAfter=15))
            
            # User Meta
            meta_data = [
                ["Researcher Name:", st.session_state.username or "N/A"],
                ["Organization:", st.session_state.organization or "N/A"],
                ["Role:", st.session_state.user_role],
                ["Selected Impact:", st.session_state.selected_impact or "All Impacts"]
            ]
            t_meta = Table(meta_data, colWidths=[130, 350])
            t_meta.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F4F9F9')),
                ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#0A6E74')),
                ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ]))
            elements.append(t_meta)
            elements.append(Spacer(1, 15))
            
            elements.append(Paragraph("Mapped Pathways Summary", styles['Heading2']))
            
            # Data table
            headers = ["Impact", "Practice", "Challenge", "Lean Principle"]
            table_rows = [headers]
            
            rep_df = df if not st.session_state.selected_impact else df[df[IMPACT_COL] == st.session_state.selected_impact]
            for _, row in rep_df.head(15).iterrows():
                table_rows.append([
                    row[IMPACT_COL][:25] + "..." if len(str(row[IMPACT_COL])) > 25 else str(row[IMPACT_COL]),
                    row[PRACTICE_COL][:25] + "..." if len(str(row[PRACTICE_COL])) > 25 else str(row[PRACTICE_COL]),
                    row[CHALLENGE_COL][:25] + "..." if len(str(row[CHALLENGE_COL])) > 25 else str(row[CHALLENGE_COL]),
                    str(row[LEAN_COL])
                ])
                
            t_data = Table(table_rows, colWidths=[120, 130, 140, 120])
            t_data.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0A6E74')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D6E6E8')),
            ]))
            elements.append(t_data)
            
            doc.build(elements)
            buffer.seek(0)
            return buffer.getvalue()

        pdf_bytes = generate_pdf()
        st.download_button(
            label="📥 Download Executive PDF Report",
            data=pdf_bytes,
            file_name="BluePath_Executive_Report.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
        
    with col_r2:
        st.markdown("""
        <div class="kpi-card">
            <h3 style="margin-top:0;">📊 Export Excel Database</h3>
            <p style="color:#5A7A84; font-size:0.92rem;">
                Download raw structured dataset containing full impact-practice-lean matrix mappings.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='BluePath_Matrix')
        buffer_excel.seek(0)
        
        st.download_button(
            label="📊 Download Complete Excel Dataset",
            data=buffer_excel,
            file_name="BluePath_Decision_Matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ============================================================
# MODULE 5: FEEDBACK & RATING
# ============================================================
elif st.session_state.page == "feedback":
    st.markdown("## ⭐ Application Feedback & Evaluation")
    st.markdown("Help us refine the BluePath decision support platform by sharing your feedback:")
    
    with st.form("feedback_form"):
        rating = st.slider("Rating (1 = Needs Work, 5 = Excellent)", 1, 5, 5)
        category = st.selectbox("Feedback Category", ["Usability & Interface", "Decision Matrix Accuracy", "Lean Principles Alignment", "Feature Request"])
        comments = st.text_area("Your Comments / Suggestions", placeholder="Describe your experience or suggested enhancements...")
        
        fb_submit = st.form_submit_button("Submit Feedback →", type="primary")
        if fb_submit:
            st.session_state.feedback_list.append({
                "rating": rating,
                "category": category,
                "comments": comments,
                "user": st.session_state.username or "Anonymous"
            })
            st.success("Thank you for your feedback! Your evaluation has been recorded.")
            
    if st.session_state.feedback_list:
        st.markdown("### 📝 Recent Feedback Submitted")
        for fb in reversed(st.session_state.feedback_list):
            st.markdown(f"""
            <div class="step-card">
                <strong>{"⭐" * fb['rating']} ({fb['category']})</strong> by <em>{fb['user']}</em>
                <p style="margin:4px 0 0 0; color:#4A6A74;">{html.escape(fb['comments'])}</p>
            </div>
            """, unsafe_allow_html=True)

# ------------------------------------------------------------
# FOOTER
# ------------------------------------------------------------
st.markdown("""
<div class="custom-footer">
    <strong>BluePath</strong> · Marine Construction Decision Support Platform<br>
    Integrating Blue Economy Practices with Prioritised Lean Construction Principles
</div>
""", unsafe_allow_html=True)
