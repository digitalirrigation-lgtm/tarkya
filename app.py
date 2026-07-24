# ============================================================
# LOCAL SQL CAREER DATA MINER – SEMANTIC SEARCH ENGINE
# ============================================================
# This app runs locally on your laptop.
# No data is sent to the internet.
# Uses sentence-transformers for semantic matching.
# ============================================================

import streamlit as st
import pandas as pd
import sqlite3
import os
import re
from datetime import datetime, timedelta
import json
import base64
from io import BytesIO

# Document generation
try:
    from docx import Document
    from docx.shared import Inches, Pt
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# Semantic Search (Local AI)
try:
    from sentence_transformers import SentenceTransformer, util
    SEMANTIC_AVAILABLE = True
except ImportError:
    SEMANTIC_AVAILABLE = False

# Language checking
try:
    import language_tool_python
    LT_AVAILABLE = True
except ImportError:
    LT_AVAILABLE = False

# ============================================================
# CONFIGURATION
# ============================================================
DB_PATH = "career_vault.db"
MODEL_NAME = "all-MiniLM-L6-v2"  # Small, fast, free, local
DEFAULT_PROFILE = {
    "name": "ZEDAGIM TESFAYE TANTU",
    "email": "zedagim100@gmail.com",
    "phone": "+251-924-700-390",
    "location": "Jigjiga, Ethiopia",
    "linkedin": "linkedin.com/in/zed10",
    "github": "digitalirrigation-lgtm.github.io/Zedagim10"
}

# ============================================================
# DATABASE INITIALIZATION
# ============================================================
def get_db():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    """Create database and tables if they don't exist"""
    conn = get_db()
    c = conn.cursor()
    
    # Profile_Data table (your career data with category codes)
    c.execute('''CREATE TABLE IF NOT EXISTS Profile_Data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_code INTEGER,
        title TEXT,
        content TEXT,
        tags TEXT,
        date_updated TEXT
    )''')
    
    # Job_History table (track applications)
    c.execute('''CREATE TABLE IF NOT EXISTS Job_History (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job_title TEXT,
        company_name TEXT,
        job_description TEXT,
        applied_date TEXT,
        status TEXT,
        generated_cv TEXT,
        generated_cover_letter TEXT,
        generated_motivation_letter TEXT
    )''')
    
    conn.commit()
    conn.close()

def insert_sample_data():
    """Insert your complete master profile data with codes"""
    conn = get_db()
    c = conn.cursor()
    
    # Check if data already exists
    c.execute("SELECT COUNT(*) FROM Profile_Data")
    if c.fetchone()[0] > 0:
        conn.close()
        return
    
    data = [
        # Code 100: Personal Information
        (100, 'Personal Information', 
         'ZEDAGIM TESFAYE TANTU | Jigjiga, Ethiopia | Phone: +251-924-700-390 | Email: zedagim100@gmail.com | LinkedIn: linkedin.com/in/zed10 | GitHub: digitalirrigation-lgtm.github.io/Zedagim10',
         'contact, ethiopia, linkedin, github', datetime.now().isoformat()),
        
        # Code 500: Education
        (500, 'Bachelor of Engineering – Water Resource & Irrigation Engineering',
         'Jigjiga University, Jigjiga, Ethiopia. Graduated: July 2022. Cumulative GPA: 3.87 out of 4.00. Rank: Top 1% of Engineering Faculty. Core Courses: GIS and Remote Sensing, Irrigation Water Management, Drip Irrigation, Sprinkler Irrigation, Irrigation Structures 1 & 2, Flood Management, Drought Management, Water Resource Planning and Management, Integrated River Basin Management, Groundwater Hydrology, Foundation Engineering, Water Supply Engineering, Water Quality Engineering, Water Well Engineering.',
         'engineering, water resource, irrigation, gpa 3.87, top 1%, remote sensing, gis', datetime.now().isoformat()),
        
        # Code 700: Kaizen Certification
        (700, 'Kaizen / Japanese Improvement System – Full Training',
         'Mastered 5S (Sort, Set, Shine, Standardize, Sustain). Muda (waste elimination). Mura (evening workload). Muri (reducing overburden). PDCA (Plan-Do-Check-Act). Kanban (visual system). JIT (Just-in-Time). Poka-yoke (mistake-proofing). QCC (Quality Control Circle). Lean Management. TQM (Total Quality Management). TPS (Toyota Production System). Jidoka (stop and fix problems). Hoshin Kanri (policy deployment). Gemba (go and see). 5 Why Analysis. Value Stream Mapping. Six Sigma. Andon System. Heijunka. Kanri.',
         'kaizen, lean, tqm, quality control, japanese system, 5s, pdca', datetime.now().isoformat()),
        
        # Code 700: Debrief Interview
        (700, 'Debrief Interview Techniques – British Home Office',
         'Trained to ask open, closed, and probing questions to find root causes. Used this to interview trafficking victims.',
         'interview, investigation, root cause, debrief, british home office', datetime.now().isoformat()),
        
        # Code 700: Python for Data Science
        (700, 'Python for Data Science – Certified',
         'Certified on June 20, 2026. Proficient in Python for data analysis and machine learning.',
         'python, data science, certification', datetime.now().isoformat()),
        
        # Code 700: Data Analysis Fundamentals
        (700, 'Data Analysis Fundamentals – Udacity',
         'Completed August 2024. Built foundational skills in analyzing and visualizing data.',
         'data analysis, udacity, fundamentals', datetime.now().isoformat()),
        
        # Code 700: AI Operations
        (700, 'Artificial Intelligence Operations – FutureLearn',
         'Academic Score: 97%. Focused on AI operational frameworks.',
         'ai, operations, futurelearn, 97%', datetime.now().isoformat()),
        
        # Code 700: Project Management
        (700, 'Project Management and Infrastructure Execution – Saylor.org',
         'Completed July 2024. Learned execution strategies for infrastructure projects.',
         'project management, infrastructure, saylor', datetime.now().isoformat()),
        
        # Code 620: Field Experience - Anti-Trafficking
        (620, 'Field Coordinator – Love Justice Ethiopia (Anti-Trafficking)',
         'Location: Jigjiga, Somali Region, Ethiopia. Period: 2019 – 2022. Mapped trafficking routes in Somali region. Interviewed 200+ potential victims. Saved 500+ potential victims from dangerous migration. 95% of victims cited drought as root cause. Led a team of 4 members using Kaizen methods. Worked with stakeholders from UK, South Africa, Uganda, Kenya, Ethiopian Federal Police, Ethiopian Justice Office. Presented to Country Director, East Africa Directors, UK stakeholders. Produced daily, weekly, monthly, quarterly, yearly reports with visuals.',
         'human trafficking, mapping, fieldwork, leadership, kaizen, somali, anti-trafficking, drought', datetime.now().isoformat()),
        
        # Code 610: Maritime Projects (4 items)
        (610, 'Maritime GeoAI – 4D Ship Tracking & Ocean Intelligence System',
         'Designed a 4D system: Latitude (X), Longitude (Y), Depth of Water (Z), Time (Present, Past, Future). Real-time ship tracking on map. Checks weather, speed, temperature, pressure. Calculates risk and predicts arrival time. Detects borders and tides automatically. Haversine formula for distance on sphere. Open Meteo API for free, real-time global weather. Updates every 0.05 seconds. Risk is calculated, not estimated.',
         'maritime, geoai, ship tracking, haversine, real-time, ocean, navigation, api', datetime.now().isoformat()),
        
        (610, 'Maritime – Vessel Route Optimization System',
         'Calculates shortest and safest routes using weather data to avoid storms.',
         'maritime, optimization, routing, vessel, safety', datetime.now().isoformat()),
        
        (610, 'Maritime – Border Detection System',
         'Detects when ships cross into different zones automatically.',
         'maritime, border detection, zones, real-time', datetime.now().isoformat()),
        
        (610, 'Maritime – Real-time Dashboard for Captains',
         'All information in one place. Satellite view of ships from space.',
         'maritime, dashboard, visualization, real-time', datetime.now().isoformat()),
        
        # Code 611: Agriculture Projects (4 items)
        (611, 'Digital Irrigation Decision-Support System',
         'Built using Python, Streamlit, Sentinel-2. NDVI thresholds: >0.6 healthy, 0.35-0.5 stressed, <0.25 critical. FAO56 Penman-Monteith evapotranspiration. SCS Curve Number for runoff. SPI for drought monitoring. CHIRPS rainfall data. CMIP6 climate prediction. Replaced ArcGIS, QGIS, SWAT, HEC-HMS, CropWat, AquaCrop with free Python code. Deployed 4 apps across 4 countries. Cut water waste 50%, fuel costs 35%.',
         'irrigation, agriculture, ndvi, sentinel-2, fa056, spi, chirps, cmip6, python', datetime.now().isoformat()),
        
        (611, 'Smart Irrigation Scheduling System',
         'Real-time soil moisture monitoring. Automated irrigation scheduling based on crop needs.',
         'irrigation, scheduling, automation, soil moisture', datetime.now().isoformat()),
        
        (611, 'Drought Early Warning System',
         'Uses SPI and CHIRPS data to predict drought conditions before they happen.',
         'drought, early warning, spi, chirps, prediction', datetime.now().isoformat()),
        
        (611, 'Crop Health Monitoring Dashboard',
         'NDVI-based crop health visualization. Monitors thousands of hectares at once.',
         'crop health, ndvi, dashboard, monitoring', datetime.now().isoformat()),
        
        # Code 400: Work Experience
        (400, 'Part-Time Kaizen Trainer',
         'Locations: Horana College, Universal College Jigjiga, Universal College Karebaya Branch. Period: 2021 – 2023. Trained 1000+ students in Kaizen techniques. Taught: 5S, Muda, Mura, Muri, PDCA, Kanban, JIT, Poka-yoke, QCC, Lean Management, TQM. Focused on workplace safety, efficiency, and continuous improvement.',
         'training, kaizen, capacity building, education', datetime.now().isoformat()),
        
        (400, 'Academic Administration – Board Member',
         'Universal College, Jigjiga. Period: 2021 – 2023. Helped establish and improve a college in remote area. Focused on Business and Medical departments. Applied QCC for education standards, training, evaluation. Applied 5S for workplace safety. Motto: "Education is a foundation for development."',
         'administration, education, quality control, board member, college', datetime.now().isoformat()),
        
        # Code 300: Technical Skills
        (300, 'General Technical Skills – Cloud & Automation',
         'Streamlit Cloud Hosting. n8n Automation Engine. API Integrations. Git Version Control. Python Operations: Pandas, NumPy, Matplotlib, Scikit-Learn.',
         'python, streamlit, cloud, automation, api, machine learning', datetime.now().isoformat()),
        
        (310, 'Technical Skills – Maritime Specific',
         'Haversine Formula. Real-time API Integration (Open Meteo). 4D Data Visualization. Border Detection Algorithms. Tide Detection Algorithms. Vessel Tracking Systems.',
         'maritime, naval, api, tracking, haversine, real-time', datetime.now().isoformat()),
        
        (311, 'Technical Skills – Agriculture / Hydrology Specific',
         'FAO56 Penman-Monteith. SCS Curve Number. NDVI (Sentinel-2). SPI. CHIRPS. CMIP6. Drought and Flood Risk Assessment.',
         'agriculture, hydrology, remote sensing, drought, irrigation, ndvi, fao56', datetime.now().isoformat()),
        
        # Code 800: Volunteering
        (800, 'Library Circulation Support (Night Shift)',
         'Jijiga University. Period: June 2021 – June 2022. Volunteered during a staffing shortage. Supported continuous library operations during off-hours. Responded to urgent, real-time needs. Received official recommendation letter for reliability.',
         'volunteer, library, night shift, reliability', datetime.now().isoformat()),
        
        # Code 200: General Skills
        (200, 'Leadership & Management Skills',
         'Kaizen / Lean Management. 5S. Muda. Mura. Muri. PDCA. QCC. TQM. Hoshin Kanri. JIT. Poka-yoke. Team Leadership. Cross-Cultural Communication (UK, South Africa, Uganda, Kenya). Stakeholder Collaboration. Project Management. Report Writing. Visual Communication. Debrief Interview Techniques.',
         'leadership, management, kaizen, project management, communication, cross-cultural', datetime.now().isoformat()),
    ]
    
    for item in data:
        c.execute(
            "INSERT INTO Profile_Data (category_code, title, content, tags, date_updated) VALUES (?, ?, ?, ?, ?)",
            item
        )
    
    conn.commit()
    conn.close()
    print("✅ Sample data inserted!")

# Initialize database
if not os.path.exists(DB_PATH):
    init_db()
    insert_sample_data()
else:
    init_db()
    # Check if data exists, if not insert
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM Profile_Data")
    if c.fetchone()[0] == 0:
        insert_sample_data()
    conn.close()

# ============================================================
# SEMANTIC SEARCH ENGINE
# ============================================================
def load_semantic_model():
    """Load the sentence-transformer model for semantic search"""
    if not SEMANTIC_AVAILABLE:
        return None
    try:
        with st.spinner("🧠 Loading semantic model (first time may take a moment)..."):
            model = SentenceTransformer(MODEL_NAME)
        return model
    except Exception as e:
        st.error(f"Could not load semantic model: {e}")
        return None

def semantic_search(job_description, model, top_k=5, prioritize_geoai=True):
    """Search the database for top matching rows using semantic similarity"""
    if model is None:
        return []
    
    conn = get_db()
    cursor = conn.cursor()
    
    # Get all profile data
    cursor.execute("SELECT id, category_code, title, content, tags FROM Profile_Data")
    rows = cursor.fetchall()
    conn.close()
    
    if not rows:
        return []
    
    # Prepare texts for embedding
    texts = []
    row_data = []
    for row in rows:
        # Combine title, content, and tags for better matching
        combined = f"{row[2]} {row[3]} {row[4]}"
        texts.append(combined)
        row_data.append({"id": row[0], "category_code": row[1], "title": row[2], "content": row[3], "tags": row[4]})
    
    # Encode job description and all texts
    job_embedding = model.encode(job_description, convert_to_tensor=True)
    text_embeddings = model.encode(texts, convert_to_tensor=True)
    
    # Compute similarity scores
    scores = util.cos_sim(job_embedding, text_embeddings)[0]
    
    # Get top matches
    top_results = []
    for idx in scores.argsort(descending=True):
        if len(top_results) >= top_k:
            break
        score = scores[idx].item()
        if score > 0.2:  # Only include if similarity is meaningful
            row = row_data[idx]
            row["score"] = score
            # Prioritize category 300 (GeoAI) by boosting score
            if prioritize_geoai and row["category_code"] in [300, 310, 311, 610, 611]:
                row["score"] = score * 1.2  # Boost GeoAI matches
            top_results.append(row)
    
    # Sort by final score
    top_results.sort(key=lambda x: x["score"], reverse=True)
    return top_results

# ============================================================
# DOCUMENT GENERATOR (6-STEP FORMULA)
# ============================================================
def generate_cv_from_matches(profile, matches, job_title=None):
    """Generate a CV using the matched profile data"""
    if not matches:
        return "No matching data found in your profile."
    
    # Extract content from matches
    personal_info = next((m for m in matches if m["category_code"] == 100), None)
    education = next((m for m in matches if m["category_code"] == 500), None)
    experience = [m for m in matches if m["category_code"] in [400, 620]]
    skills = [m for m in matches if m["category_code"] in [300, 310, 311]]
    projects = [m for m in matches if m["category_code"] in [610, 611]]
    certifications = [m for m in matches if m["category_code"] == 700]
    
    # Build CV
    cv = []
    
    # Personal Information
    if personal_info:
        cv.append("=" * 50)
        cv.append(personal_info["content"])
        cv.append("=" * 50)
    
    cv.append("")
    
    # Professional Summary
    cv.append("**PROFESSIONAL SUMMARY**")
    cv.append("Water Resource and Irrigation Engineer with a GPA of 3.87/4.00. Specialized in Python-driven digital irrigation and maritime intelligence systems using satellite data and local AI. Proven track record of deploying 4 apps across 4 countries. Strong field experience leading teams in remote areas and collaborating across cultures.")
    cv.append("")
    
    # Education
    if education:
        cv.append("**EDUCATION**")
        cv.append(education["content"])
        cv.append("")
    
    # Experience
    if experience:
        cv.append("**EXPERIENCE**")
        for exp in experience:
            cv.append(f"• {exp['title']}")
            cv.append(f"  {exp['content'][:300]}...")
        cv.append("")
    
    # Projects
    if projects:
        cv.append("**KEY PROJECTS**")
        for proj in projects:
            cv.append(f"• {proj['title']}")
            cv.append(f"  {proj['content'][:200]}...")
        cv.append("")
    
    # Skills
    if skills:
        cv.append("**SKILLS**")
        skill_texts = [s["content"] for s in skills]
        cv.append("; ".join(skill_texts))
        cv.append("")
    
    # Certifications
    if certifications:
        cv.append("**CERTIFICATIONS**")
        for cert in certifications:
            cv.append(f"• {cert['title']}")
        cv.append("")
    
    return "\n".join(cv)

def generate_cover_letter_from_matches(profile, matches, job_title=None, company=None):
    """Generate a cover letter using the 6-step formula"""
    if not matches:
        return "No matching data found in your profile."
    
    # Extract content from matches
    personal = next((m for m in matches if m["category_code"] == 100), None)
    achievements = [m for m in matches if m["category_code"] in [610, 611, 620]]
    skills = [m for m in matches if m["category_code"] in [300, 310, 311]]
    
    # Build letter using 6-step formula
    letter = []
    
    # Step 1: HOOK
    if achievements:
        hook_achievement = achievements[0] if achievements else None
        if hook_achievement:
            letter.append(f"Your work on {hook_achievement['title']} caught my attention.")
        else:
            letter.append("I am writing to express my interest in your opportunity.")
    
    letter.append("")
    
    # Step 2: PIVOT
    letter.append("My background in Water Resource Engineering and GeoAI aligns well with your needs.")
    letter.append("")
    
    # Step 3: EVIDENCE (extract from matches)
    letter.append("**My Experience:**")
    for ach in achievements[:2]:
        letter.append(f"• {ach['title']}: {ach['content'][:150]}...")
    letter.append("")
    
    # Step 4: RESULTS (numbers from matches)
    letter.append("**Key Results:**")
    result_items = [
        "• Deployed 4 production applications across 4 countries",
        "• Reduced irrigation water waste by approximately 50%",
        "• Reduced operational fuel costs by about 35%",
        "• Interviewed 200+ victims, saved 500+ from migration",
        "• Trained 1000+ students in Kaizen"
    ]
    for item in result_items:
        letter.append(item)
    letter.append("")
    
    # Step 5: VALUE
    letter.append("**What I Offer:**")
    letter.append("I combine strong technical skills with real field experience. I make complex data simple and useful for decision-makers. I have proven I can deliver results with zero budget.")
    letter.append("")
    
    # Step 6: CALL TO ACTION
    letter.append("I am ready to discuss my potential alignment with your goals.")
    if personal:
        letter.append(f"You can reach me at {personal['content']}")
    letter.append("I respond within hours.")
    
    return "\n".join(letter)

def generate_motivation_letter_from_matches(profile, matches, program_name=None):
    """Generate a motivation letter using the 6-step formula"""
    if not matches:
        return "No matching data found in your profile."
    
    # Extract content from matches
    personal = next((m for m in matches if m["category_code"] == 100), None)
    narrative = next((m for m in matches if m["category_code"] == 620), None)
    achievements = [m for m in matches if m["category_code"] in [610, 611]]
    
    # Build letter
    letter = []
    
    # Step 1: HOOK - Personal Story
    letter.append("I watched people lose everything to drought and be forced into dangerous migration.")
    letter.append("")
    
    # Step 2: PIVOT
    letter.append("That experience changed me. I realized that data is not just numbers – it is survival.")
    letter.append("")
    
    # Step 3: EVIDENCE
    if narrative:
        letter.append(narrative["content"][:300] + "...")
    letter.append("")
    
    # Step 4: RESULTS
    letter.append("I built working prototypes, deployed 4 applications across 4 countries, and achieved:")
    result_items = [
        "• 50% water savings",
        "• 35% fuel cost reduction",
        "• Early detection of crop stress before it is visible"
    ]
    for item in result_items:
        letter.append(item)
    letter.append("")
    
    # Step 5: VALUE
    letter.append("I make complex technical data simple and useful. I combine technical skills with field experience. I solve problems with limited resources.")
    letter.append("")
    
    # Step 6: CALL TO ACTION
    letter.append("I am ready to contribute my skills, experience, and passion to your program.")
    if personal:
        letter.append(f"Email: zedagim100@gmail.com | Phone: +251-924-700-390")
    
    return "\n".join(letter)

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(layout="wide", page_title="🚀 Local Career Data Miner", page_icon="🚀")

# Theme
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(145deg, #d4e6f1 0%, #f0f0f0 50%, #e8e8e8 100%);
    }
    .stButton button {
        background: linear-gradient(145deg, #FFD700, #B8860B) !important;
        color: #1a1a2e !important;
        border-radius: 30px !important;
        font-weight: bold !important;
    }
    .golden-text {
        color: #b8860b;
        text-shadow: 0 0 8px rgba(184, 134, 11, 0.3);
    }
</style>
""", unsafe_allow_html=True)

st.title("🚀 Local Career Data Miner")
st.markdown("<p class='golden-text'>Semantic Search • No Internet • Your Data Stays Local</p>", unsafe_allow_html=True)

# ---------- SIDEBAR ----------
st.sidebar.title("📊 Dashboard")
st.sidebar.info("💡 This app uses local AI to find the best parts of your profile for each job application.")

# Check semantic model status
if SEMANTIC_AVAILABLE:
    st.sidebar.success("✅ Semantic search available")
else:
    st.sidebar.warning("⚠️ Install sentence-transformers for semantic search")

# ---------- TABS ----------
tab1, tab2, tab3 = st.tabs(["📚 Manage Data", "🎯 Apply", "📄 Export"])

# ============================================================
# TAB 1: MANAGE DATA
# ============================================================
with tab1:
    st.subheader("📚 Your Career Database")
    
    conn = get_db()
    df = pd.read_sql("SELECT id, category_code, title, content, tags, date_updated FROM Profile_Data ORDER BY category_code", conn)
    conn.close()
    
    st.dataframe(df, use_container_width=True)
    
    st.markdown("---")
    st.subheader("📊 Category Summary")
    
    cat_counts = df.groupby('category_code').size().reset_index(name='count')
    cat_counts.columns = ['Category Code', 'Count']
    st.dataframe(cat_counts, use_container_width=True)
    
    st.markdown("---")
    st.subheader("➕ Add New Entry")
    
    with st.form("add_entry_form"):
        col1, col2 = st.columns(2)
        with col1:
            category_code = st.selectbox("Category Code", [100, 200, 300, 310, 311, 400, 500, 610, 611, 620, 700, 800])
            title = st.text_input("Title")
        with col2:
            tags = st.text_input("Tags (comma separated)")
        content = st.text_area("Content", height=150)
        
        if st.form_submit_button("Add Entry"):
            if title and content:
                conn = get_db()
                c = conn.cursor()
                c.execute(
                    "INSERT INTO Profile_Data (category_code, title, content, tags, date_updated) VALUES (?, ?, ?, ?, ?)",
                    (category_code, title, content, tags, datetime.now().isoformat())
                )
                conn.commit()
                conn.close()
                st.success("✅ Entry added!")
                st.rerun()
            else:
                st.warning("Title and content are required.")

# ============================================================
# TAB 2: APPLY
# ============================================================
with tab2:
    st.subheader("🎯 Paste Job or Scholarship Description")
    
    # Job description input
    job_description = st.text_area(
        "Paste the job or scholarship description here",
        height=200,
        placeholder="Paste the job description from the company website..."
    )
    
    col1, col2 = st.columns(2)
    with col1:
        job_title = st.text_input("Job Title (optional)")
    with col2:
        company_name = st.text_input("Company/Organization (optional)")
    
    # Load semantic model
    model = None
    if SEMANTIC_AVAILABLE and job_description:
        model = load_semantic_model()
    
    col_run, col_clear = st.columns(2)
    with col_run:
        if st.button("🔍 Run Semantic Search", use_container_width=True):
            if not job_description:
                st.warning("Please paste a job description first.")
            elif model is None:
                st.error("Semantic model not available. Please install sentence-transformers.")
            else:
                with st.spinner("🔍 Searching your database..."):
                    matches = semantic_search(job_description, model, top_k=8)
                    
                    if matches:
                        st.success(f"✅ Found {len(matches)} matching items!")
                        
                        # Display matches
                        st.subheader("📌 Top Matches")
                        for m in matches:
                            score = m.get("score", 0)
                            st.markdown(f"**{m['title']}** (Category {m['category_code']}) - Score: {score:.3f}")
                            st.caption(m['content'][:200] + "...")
                            st.divider()
                        
                        # Store matches in session state
                        st.session_state['matches'] = matches
                        st.session_state['job_description'] = job_description
                        st.session_state['job_title'] = job_title
                        st.session_state['company_name'] = company_name
                    else:
                        st.warning("No matching data found in your profile.")
    
    with col_clear:
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state.pop('matches', None)
            st.session_state.pop('job_description', None)
            st.rerun()
    
    # If we have matches, show document generation
    if 'matches' in st.session_state and st.session_state['matches']:
        st.subheader("📄 Generate Documents")
        
        col_gen1, col_gen2, col_gen3 = st.columns(3)
        
        with col_gen1:
            if st.button("📝 Generate CV", use_container_width=True):
                cv = generate_cv_from_matches(
                    DEFAULT_PROFILE,
                    st.session_state['matches'],
                    st.session_state.get('job_title')
                )
                st.session_state['generated_cv'] = cv
                st.success("✅ CV generated!")
        
        with col_gen2:
            if st.button("✉️ Generate Cover Letter", use_container_width=True):
                cl = generate_cover_letter_from_matches(
                    DEFAULT_PROFILE,
                    st.session_state['matches'],
                    st.session_state.get('job_title'),
                    st.session_state.get('company_name')
                )
                st.session_state['generated_cl'] = cl
                st.success("✅ Cover Letter generated!")
        
        with col_gen3:
            if st.button("📨 Generate Motivation Letter", use_container_width=True):
                ml = generate_motivation_letter_from_matches(
                    DEFAULT_PROFILE,
                    st.session_state['matches'],
                    st.session_state.get('company_name')
                )
                st.session_state['generated_ml'] = ml
                st.success("✅ Motivation Letter generated!")

# ============================================================
# TAB 3: EXPORT
# ============================================================
with tab3:
    st.subheader("📄 Preview and Export Documents")
    
    # Display generated documents
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if 'generated_cv' in st.session_state:
            st.text_area("📄 CV", st.session_state['generated_cv'], height=300)
            st.download_button(
                "⬇️ Download CV",
                data=st.session_state['generated_cv'],
                file_name=f"CV_{datetime.now().strftime('%Y%m%d')}.txt"
            )
    
    with col2:
        if 'generated_cl' in st.session_state:
            st.text_area("✉️ Cover Letter", st.session_state['generated_cl'], height=300)
            st.download_button(
                "⬇️ Download Cover Letter",
                data=st.session_state['generated_cl'],
                file_name=f"Cover_Letter_{datetime.now().strftime('%Y%m%d')}.txt"
            )
    
    with col3:
        if 'generated_ml' in st.session_state:
            st.text_area("📨 Motivation Letter", st.session_state['generated_ml'], height=300)
            st.download_button(
                "⬇️ Download Motivation Letter",
                data=st.session_state['generated_ml'],
                file_name=f"Motivation_Letter_{datetime.now().strftime('%Y%m%d')}.txt"
            )
    
    # Save to Job History
    if 'matches' in st.session_state and st.session_state['matches']:
        st.markdown("---")
        st.subheader("💾 Save to Job History")
        
        if st.button("Save This Application"):
            conn = get_db()
            c = conn.cursor()
            c.execute("""
                INSERT INTO Job_History (job_title, company_name, job_description, applied_date, status, generated_cv, generated_cover_letter, generated_motivation_letter)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                st.session_state.get('job_title', ''),
                st.session_state.get('company_name', ''),
                st.session_state.get('job_description', ''),
                datetime.now().isoformat(),
                'Draft',
                st.session_state.get('generated_cv', ''),
                st.session_state.get('generated_cl', ''),
                st.session_state.get('generated_ml', '')
            ))
            conn.commit()
            conn.close()
            st.success("✅ Application saved to history!")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption(f"⚡ Data stored in {DB_PATH} | Semantic Model: {MODEL_NAME} | All AI runs locally")
