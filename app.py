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

# Document generation
try:
    from docx import Document
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
MODEL_NAME = "all-MiniLM-L6-v2"
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
    conn = get_db()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS Profile_Data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category_code INTEGER,
        title TEXT,
        content TEXT,
        tags TEXT,
        date_updated TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS Job_History (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job_title TEXT,
        company_name TEXT,
        job_description TEXT,
        applied_date TEXT,
        deadline_date TEXT,
        status TEXT,
        generated_cv TEXT,
        generated_cover_letter TEXT,
        generated_motivation_letter TEXT
    )''')
    conn.commit()
    conn.close()

def ensure_job_history_schema():
    conn = get_db()
    c = conn.cursor()
    c.execute("PRAGMA table_info(Job_History)")
    columns = [col[1] for col in c.fetchall()]
    if 'deadline_date' not in columns:
        c.execute("ALTER TABLE Job_History ADD COLUMN deadline_date TEXT")
    if 'status' not in columns:
        c.execute("ALTER TABLE Job_History ADD COLUMN status TEXT DEFAULT 'Saved'")
    conn.commit()
    conn.close()

def insert_sample_data():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM Profile_Data")
    if c.fetchone()[0] > 0:
        conn.close()
        return
    data = [
        (100, 'Personal Information', 'ZEDAGIM TESFAYE TANTU | Jigjiga, Ethiopia | Phone: +251-924-700-390 | Email: zedagim100@gmail.com | LinkedIn: linkedin.com/in/zed10 | GitHub: digitalirrigation-lgtm.github.io/Zedagim10', 'contact, ethiopia, linkedin, github', datetime.now().isoformat()),
        (500, 'Bachelor of Engineering – Water Resource & Irrigation Engineering', 'Jigjiga University, Jigjiga, Ethiopia. Graduated: July 2022. Cumulative GPA: 3.87 out of 4.00. Rank: Top 1% of Engineering Faculty. Core Courses: GIS and Remote Sensing, Irrigation Water Management, Drip Irrigation, Sprinkler Irrigation, Irrigation Structures 1 & 2, Flood Management, Drought Management, Water Resource Planning and Management, Integrated River Basin Management, Groundwater Hydrology, Foundation Engineering, Water Supply Engineering, Water Quality Engineering, Water Well Engineering.', 'engineering, water resource, irrigation, gpa 3.87, top 1%, remote sensing, gis', datetime.now().isoformat()),
        (700, 'Kaizen / Japanese Improvement System – Full Training', 'Mastered 5S (Sort, Set, Shine, Standardize, Sustain). Muda (waste elimination). Mura (evening workload). Muri (reducing overburden). PDCA (Plan-Do-Check-Act). Kanban (visual system). JIT (Just-in-Time). Poka-yoke (mistake-proofing). QCC (Quality Control Circle). Lean Management. TQM (Total Quality Management). TPS (Toyota Production System). Jidoka (stop and fix problems). Hoshin Kanri (policy deployment). Gemba (go and see). 5 Why Analysis. Value Stream Mapping. Six Sigma. Andon System. Heijunka. Kanri.', 'kaizen, lean, tqm, quality control, japanese system, 5s, pdca', datetime.now().isoformat()),
        (700, 'Debrief Interview Techniques – British Home Office', 'Trained to ask open, closed, and probing questions to find root causes.', 'interview, investigation, root cause, debrief', datetime.now().isoformat()),
        (700, 'Python for Data Science – Certified', 'Certified on June 20, 2026. Proficient in Python for data analysis and machine learning.', 'python, data science, certification', datetime.now().isoformat()),
        (700, 'Data Analysis Fundamentals – Udacity', 'Completed August 2024.', 'data analysis, udacity', datetime.now().isoformat()),
        (700, 'Artificial Intelligence Operations – FutureLearn', 'Academic Score: 97%.', 'ai, operations, futurelearn', datetime.now().isoformat()),
        (700, 'Project Management and Infrastructure Execution – Saylor.org', 'Completed July 2024.', 'project management, infrastructure', datetime.now().isoformat()),
        (620, 'Field Coordinator – Love Justice Ethiopia (Anti-Trafficking)', 'Location: Jigjiga, Somali Region, Ethiopia. Period: 2019 – 2022. Mapped trafficking routes. Interviewed 200+ victims. Saved 500+ from dangerous migration. 95% cited drought as root cause. Led a team of 4 using Kaizen methods. Collaborated with UK, South Africa, Uganda, Kenya, Ethiopian Police, Justice Office. Presented to Country Director, UK stakeholders. Produced daily, weekly, monthly, yearly reports with visuals.', 'human trafficking, mapping, fieldwork, leadership, kaizen, somali, anti-trafficking, drought', datetime.now().isoformat()),
        (610, 'Maritime GeoAI – 4D Ship Tracking & Ocean Intelligence System', 'Designed a 4D system: Latitude, Longitude, Depth, Time. Real-time tracking, weather, speed, pressure, risk calculation, arrival prediction, border/tide detection. Haversine formula for distance. Open Meteo API for live weather. Updates every 0.05 seconds. Risk is calculated, not estimated.', 'maritime, geoai, ship tracking, haversine, real-time, ocean, navigation, api', datetime.now().isoformat()),
        (610, 'Maritime – Vessel Route Optimization System', 'Calculates shortest and safest routes using weather data to avoid storms.', 'maritime, optimization, routing, vessel, safety', datetime.now().isoformat()),
        (610, 'Maritime – Border Detection System', 'Detects when ships cross into different zones automatically.', 'maritime, border detection, zones', datetime.now().isoformat()),
        (610, 'Maritime – Real-time Dashboard for Captains', 'All information in one place. Satellite view of ships from space.', 'maritime, dashboard, visualization', datetime.now().isoformat()),
        (611, 'Digital Irrigation Decision-Support System', 'Built using Python, Streamlit, Sentinel-2. NDVI thresholds: >0.6 healthy, 0.35-0.5 stressed, <0.25 critical. FAO56 Penman-Monteith evapotranspiration. SCS Curve Number for runoff. SPI for drought monitoring. CHIRPS rainfall data. CMIP6 climate prediction. Replaced ArcGIS, QGIS, SWAT, HEC-HMS, CropWat, AquaCrop with free Python. Deployed 4 apps across 4 countries. Cut water waste 50%, fuel costs 35%.', 'irrigation, agriculture, ndvi, sentinel-2, fa056, spi, chirps, cmip6, python', datetime.now().isoformat()),
        (611, 'Smart Irrigation Scheduling System', 'Real-time soil moisture monitoring. Automated irrigation scheduling.', 'irrigation, scheduling, automation, soil moisture', datetime.now().isoformat()),
        (611, 'Drought Early Warning System', 'Uses SPI and CHIRPS data to predict drought conditions.', 'drought, early warning, spi, chirps', datetime.now().isoformat()),
        (611, 'Crop Health Monitoring Dashboard', 'NDVI-based crop health visualization. Monitors thousands of hectares.', 'crop health, ndvi, dashboard', datetime.now().isoformat()),
        (400, 'Part-Time Kaizen Trainer', 'Trained 1000+ students at Horana College, Universal College Jigjiga, Karebaya Branch. Taught: 5S, Muda, Mura, Muri, PDCA, Kanban, JIT, Poka-yoke, QCC, Lean, TQM.', 'training, kaizen, capacity building', datetime.now().isoformat()),
        (400, 'Academic Administration – Board Member', 'Universal College, Jigjiga. Helped establish college in remote area. Applied QCC for education standards. Applied 5S for workplace safety. Motto: "Education is a foundation for development."', 'administration, education, quality control, board member', datetime.now().isoformat()),
        (300, 'General Technical Skills – Cloud & Automation', 'Streamlit Cloud Hosting, n8n, API Integrations, Git, Python: Pandas, NumPy, Matplotlib, Scikit-Learn.', 'python, streamlit, cloud, automation, api, machine learning', datetime.now().isoformat()),
        (310, 'Technical Skills – Maritime Specific', 'Haversine Formula, Real-time API Integration (Open Meteo), 4D Data Visualization, Border/Tide Detection, Vessel Tracking.', 'maritime, naval, api, tracking, haversine', datetime.now().isoformat()),
        (311, 'Technical Skills – Agriculture / Hydrology Specific', 'FAO56, SCS Curve Number, NDVI (Sentinel-2), SPI, CHIRPS, CMIP6, Drought/Flood Assessment.', 'agriculture, hydrology, remote sensing, drought, irrigation, ndvi, fao56', datetime.now().isoformat()),
        (800, 'Library Circulation Support (Night Shift)', 'Jijiga University. Volunteered during staffing shortage. Supported continuous library operations. Received recommendation letter for reliability.', 'volunteer, library, night shift', datetime.now().isoformat()),
        (200, 'Leadership & Management Skills', 'Kaizen / Lean Management, 5S, Muda, Mura, Muri, PDCA, QCC, TQM, Hoshin Kanri, JIT, Poka-yoke, Team Leadership, Cross-Cultural Communication (UK, South Africa, Uganda, Kenya), Stakeholder Collaboration, Project Management, Report Writing, Visual Communication, Debrief Interview Techniques.', 'leadership, management, kaizen, project management, communication', datetime.now().isoformat()),
    ]
    for item in data:
        c.execute("INSERT INTO Profile_Data (category_code, title, content, tags, date_updated) VALUES (?, ?, ?, ?, ?)", item)
    conn.commit()
    conn.close()

# Initialize database
if not os.path.exists(DB_PATH):
    init_db()
    insert_sample_data()
else:
    init_db()
    ensure_job_history_schema()
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
    if not SEMANTIC_AVAILABLE:
        return None
    try:
        with st.spinner("🧠 Loading semantic model (first time may take a moment)..."):
            model = SentenceTransformer(MODEL_NAME)
        return model
    except Exception as e:
        st.error(f"Could not load semantic model: {e}")
        return None

def semantic_search(job_description, model, top_k=10, prioritize_geoai=True):
    if model is None:
        return []
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT id, category_code, title, content, tags FROM Profile_Data")
    rows = cursor.fetchall()
    conn.close()
    if not rows:
        return []
    texts = []
    row_data = []
    for row in rows:
        combined = f"{row[2]} {row[3]} {row[4]}"
        texts.append(combined)
        row_data.append({"id": row[0], "category_code": row[1], "title": row[2], "content": row[3], "tags": row[4]})
    job_embedding = model.encode(job_description, convert_to_tensor=True)
    text_embeddings = model.encode(texts, convert_to_tensor=True)
    scores = util.cos_sim(job_embedding, text_embeddings)[0]
    top_results = []
    for idx in scores.argsort(descending=True):
        if len(top_results) >= top_k:
            break
        score = scores[idx].item()
        if score > 0.25:
            row = row_data[idx]
            row["score"] = score
            if prioritize_geoai and row["category_code"] in [300, 310, 311, 610, 611]:
                row["score"] = score * 1.3
            top_results.append(row)
    top_results.sort(key=lambda x: x["score"], reverse=True)
    return top_results

# ============================================================
# DYNAMIC DOCUMENT GENERATORS
# ============================================================
def extract_job_requirements(job_description):
    tech_keywords = ['python', 'gis', 'remote sensing', 'satellite', 'ndvi', 'irrigation', 
                     'water', 'hydrology', 'drought', 'climate', 'data analysis', 'machine learning',
                     'streamlit', 'api', 'geoai', 'maritime', 'ship tracking', 'haversine',
                     'fao56', 'spi', 'chirps', 'cmip6', 'sentinel']
    extracted = []
    for word in tech_keywords:
        if word in job_description.lower():
            extracted.append(word)
    return extracted

def generate_dynamic_cv(matches, job_title, job_description):
    if not matches:
        return "No matching data found in your profile. Please add more data."
    
    personal_info = next((m for m in matches if m["category_code"] == 100), None)
    education = next((m for m in matches if m["category_code"] == 500), None)
    experience = [m for m in matches if m["category_code"] in [400, 620]]
    skills = [m for m in matches if m["category_code"] in [300, 310, 311]]
    projects = [m for m in matches if m["category_code"] in [610, 611]]
    certifications = [m for m in matches if m["category_code"] == 700]
    job_req = extract_job_requirements(job_description)
    
    cv = []
    if personal_info:
        cv.append("=" * 60)
        cv.append(personal_info["content"])
        cv.append("=" * 60)
    cv.append("")
    
    cv.append("**PROFESSIONAL SUMMARY**")
    summary = "Water Resource and Irrigation Engineer with a GPA of 3.87/4.00. "
    if any(t in job_req for t in ['maritime', 'ship', 'navigation']):
        summary += "Specialized in Maritime GeoAI and real-time ship tracking using Python and satellite data. "
    if any(t in job_req for t in ['irrigation', 'agriculture', 'drought']):
        summary += "Specialized in Digital Irrigation Systems using Python, satellite data, and hydrological modeling. "
    summary += "Deployed 4 applications across 4 countries. Strong field experience leading teams in remote areas."
    cv.append(summary)
    cv.append("")
    
    if education:
        cv.append("**EDUCATION**")
        cv.append(education["content"])
        cv.append("")
    
    if experience:
        cv.append("**EXPERIENCE**")
        for exp in experience:
            cv.append(f"• {exp['title']}")
            content = exp['content']
            if job_req:
                matched = False
                for req in job_req[:3]:
                    if req in content.lower():
                        cv.append(f"  {content[:250]}...")
                        matched = True
                        break
                if not matched:
                    cv.append(f"  {content[:200]}...")
            else:
                cv.append(f"  {content[:200]}...")
        cv.append("")
    
    cv.append("**KEY PROJECTS**")
    for proj in projects:
        cv.append(f"• {proj['title']}")
        if job_req and any(req in proj['content'].lower() for req in job_req):
            cv.append(f"  ✅ {proj['content'][:150]}...")
        else:
            cv.append(f"  {proj['content'][:150]}...")
    cv.append("")
    
    if skills:
        cv.append("**SKILLS**")
        skill_texts = []
        for s in skills:
            if job_req and any(req in s['content'].lower() for req in job_req):
                skill_texts.append(f"⭐ {s['content']}")
            else:
                skill_texts.append(s['content'])
        cv.append("; ".join(skill_texts))
        cv.append("")
    
    if certifications:
        cv.append("**CERTIFICATIONS**")
        for cert in certifications:
            cv.append(f"• {cert['title']}")
    
    return "\n".join(cv)

def generate_dynamic_cover_letter(matches, job_title, company, job_description):
    if not matches:
        return "No matching data found in your profile. Please add more data."
    
    personal = next((m for m in matches if m["category_code"] == 100), None)
    achievements = [m for m in matches if m["category_code"] in [610, 611, 620]]
    job_req = extract_job_requirements(job_description)
    
    letter = []
    if achievements:
        best_ach = max(achievements[:3], key=lambda x: x.get('score', 0))
        letter.append(f"Your work on {best_ach['title']} caught my attention. I've been building similar solutions that deliver measurable results.")
    else:
        letter.append("I am writing to express my strong interest in this opportunity.")
    letter.append("")
    
    if job_req:
        letter.append(f"My background in Water Resource Engineering and GeoAI, particularly in {', '.join(job_req[:3])}, aligns with your needs.")
    else:
        letter.append("My background in Water Resource Engineering and GeoAI aligns with your needs.")
    letter.append("")
    
    letter.append("**What I Have Done:**")
    for ach in achievements[:3]:
        letter.append(f"• {ach['title']}")
        content = ach['content'][:200]
        if job_req:
            for req in job_req[:2]:
                if req in content.lower():
                    content = content.replace(req, f"**{req}**")
        letter.append(f"  {content}...")
    letter.append("")
    
    letter.append("**My Results:**")
    results = [
        "• Deployed 4 production applications across 4 countries",
        "• Reduced irrigation water waste by approximately 50%",
        "• Reduced operational fuel costs by about 35%",
        "• Interviewed 200+ victims, saved 500+ from migration",
        "• Trained 1000+ students in Kaizen"
    ]
    for r in results:
        letter.append(r)
    letter.append("")
    
    letter.append("**How I Will Contribute:**")
    if job_req:
        value = f"I will apply my expertise in {', '.join(job_req[:3])} to "
        if 'maritime' in job_req or 'ship' in job_req:
            value += "optimize your maritime operations, reduce costs, and improve safety."
        elif 'irrigation' in job_req or 'agriculture' in job_req:
            value += "optimize your irrigation systems, reduce water waste, and increase crop yield."
        else:
            value += "solve your water and climate challenges with data-driven solutions."
    else:
        value = "I will solve your water and climate challenges with data-driven solutions."
    letter.append(value)
    letter.append("")
    
    letter.append("**I am ready to discuss how I can contribute to your team.**")
    if personal:
        contact = personal["content"].split("|")
        for c in contact:
            if "Email" in c or "Phone" in c:
                letter.append(c.strip())
    letter.append("I respond within hours and am available for a conversation at your convenience.")
    
    return "\n".join(letter)

def generate_dynamic_motivation_letter(matches, program_name, job_description):
    if not matches:
        return "No matching data found in your profile. Please add more data."
    
    personal = next((m for m in matches if m["category_code"] == 100), None)
    narrative = next((m for m in matches if m["category_code"] == 620), None)
    achievements = [m for m in matches if m["category_code"] in [610, 611]]
    
    letter = []
    if narrative:
        letter.append("I watched people lose everything to drought and be forced into dangerous migration.")
    else:
        letter.append("My journey in water engineering began with a simple observation: farmers need data to survive.")
    letter.append("")
    
    letter.append("That experience changed me. I realized that data is not just numbers – it is survival.")
    letter.append("")
    
    if narrative:
        letter.append(narrative["content"][:250] + "...")
    letter.append("")
    
    letter.append("I built working prototypes, deployed 4 applications across 4 countries, and achieved:")
    results = ["• 50% water savings", "• 35% fuel cost reduction", "• Early detection of crop stress"]
    for r in results:
        letter.append(r)
    letter.append("")
    
    letter.append("I make complex technical data simple and useful. I combine technical skills with field experience.")
    if narrative and "UK" in narrative["content"]:
        letter.append("I have collaborated across cultures – UK, South Africa, Uganda, Kenya – and understand how to deliver results globally.")
    letter.append("")
    
    letter.append("I am ready to contribute my skills, experience, and passion to your program.")
    if personal:
        letter.append("Email: zedagim100@gmail.com | Phone: +251-924-700-390")
    letter.append("")
    letter.append("I look forward to the opportunity to make a difference.")
    
    return "\n".join(letter)

# ============================================================
# JOB HISTORY FUNCTIONS
# ============================================================
def save_application(job_title, company_name, job_description, deadline_date, cv, cl, ml):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO Job_History (job_title, company_name, job_description, applied_date, deadline_date, status, generated_cv, generated_cover_letter, generated_motivation_letter)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        job_title, company_name, job_description,
        datetime.now().isoformat(),
        deadline_date,
        'Saved',
        cv, cl, ml
    ))
    conn.commit()
    conn.close()

def get_job_history():
    conn = get_db()
    df = pd.read_sql("SELECT * FROM Job_History ORDER BY id DESC", conn)
    conn.close()
    return df

def get_latest_application():
    conn = get_db()
    df = pd.read_sql("SELECT * FROM Job_History ORDER BY id DESC LIMIT 1", conn)
    conn.close()
    if df.empty:
        return None
    return df.iloc[0].to_dict()

def delete_application(app_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM Job_History WHERE id = ?", (app_id,))
    conn.commit()
    conn.close()

def update_status(app_id, new_status):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE Job_History SET status = ? WHERE id = ?", (new_status, app_id))
    conn.commit()
    conn.close()

# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(layout="wide", page_title="🚀 AI Career Data Miner", page_icon="🚀")

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
    .evidence-box {
        background: rgba(255,255,255,0.8);
        border-radius: 10px;
        padding: 15px;
        border-left: 4px solid #b8860b;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

st.title("🚀 AI Career Data Miner")
st.markdown("<p class='golden-text'>Semantic Search • Dynamic Generation • Auto-Save to History</p>", unsafe_allow_html=True)

st.sidebar.title("📊 Dashboard")
st.sidebar.info("💡 Paste a job description, AI will extract matching profile data and generate tailored documents.")

if SEMANTIC_AVAILABLE:
    st.sidebar.success("✅ Semantic search available")
else:
    st.sidebar.warning("⚠️ Install sentence-transformers")

# ============================================================
# TABS
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📚 Manage Data", "🎯 Apply", "📄 Export", "📝 Templates", "📅 Job History"])

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
    st.markdown("_The AI will analyze, extract requirements, and generate tailored documents. After generation, documents are saved to history and cleared from this tab._")
    
    job_description = st.text_area(
        "Paste the job or scholarship description here",
        height=250,
        placeholder="Paste the complete job description from the company website..."
    )
    
    col1, col2, col3 = st.columns(3)
    with col1:
        job_title = st.text_input("Job Title")
    with col2:
        company_name = st.text_input("Company/Organization")
    with col3:
        deadline_date = st.date_input("Application Deadline", value=datetime.now().date() + timedelta(days=30))
    
    model = None
    if SEMANTIC_AVAILABLE and job_description:
        model = load_semantic_model()
    
    col_run, col_clear = st.columns(2)
    with col_run:
        if st.button("🔍 Analyze Job", use_container_width=True):
            if not job_description:
                st.warning("Please paste a job description first.")
            elif model is None:
                st.error("Semantic model not available. Please install sentence-transformers.")
            else:
                with st.spinner("🔍 Analyzing job description and matching your profile..."):
                    matches = semantic_search(job_description, model, top_k=10)
                    if matches:
                        st.success(f"✅ Found {len(matches)} matching items!")
                        job_req = extract_job_requirements(job_description)
                        if job_req:
                            st.subheader("🔑 Extracted Key Requirements")
                            st.markdown(f"**Detected skills:** {', '.join(job_req)}")
                        
                        with st.expander("📌 View Top Matches"):
                            for m in matches[:5]:
                                score = m.get("score", 0)
                                st.markdown(f"**{m['title']}** (Category {m['category_code']}) - Score: {score:.3f}")
                                st.caption(m['content'][:200] + "...")
                                st.divider()
                        
                        # Store matches and job info in session state
                        st.session_state['matches'] = matches
                        st.session_state['job_description'] = job_description
                        st.session_state['job_title'] = job_title
                        st.session_state['company_name'] = company_name
                        st.session_state['deadline_date'] = deadline_date.strftime("%Y-%m-%d")
                    else:
                        st.warning("No matching data found in your profile.")
    
    with col_clear:
        if st.button("🗑️ Clear All", use_container_width=True):
            for key in ['matches', 'job_description', 'job_title', 'company_name', 'deadline_date']:
                st.session_state.pop(key, None)
            st.rerun()
    
    # If we have matches, show "Generate & Save" button
    if 'matches' in st.session_state and st.session_state['matches']:
        st.markdown("---")
        st.subheader("📄 Generate & Save All Documents")
        st.markdown("_This will generate your CV, Cover Letter, and Motivation Letter, save them to history, and clear the UI for the next job._")
        
        if st.button("🚀 Generate All & Save to History", use_container_width=True):
            with st.spinner("🔄 Generating documents..."):
                cv = generate_dynamic_cv(
                    st.session_state['matches'],
                    st.session_state.get('job_title', ''),
                    st.session_state.get('job_description', '')
                )
                cl = generate_dynamic_cover_letter(
                    st.session_state['matches'],
                    st.session_state.get('job_title', ''),
                    st.session_state.get('company_name', ''),
                    st.session_state.get('job_description', '')
                )
                ml = generate_dynamic_motivation_letter(
                    st.session_state['matches'],
                    st.session_state.get('company_name', ''),
                    st.session_state.get('job_description', '')
                )
                
                # Save to history
                save_application(
                    st.session_state.get('job_title', ''),
                    st.session_state.get('company_name', ''),
                    st.session_state.get('job_description', ''),
                    st.session_state.get('deadline_date', datetime.now().strftime("%Y-%m-%d")),
                    cv, cl, ml
                )
                
                # Clear session state to remove documents from UI
                for key in ['matches', 'job_description', 'job_title', 'company_name', 'deadline_date']:
                    st.session_state.pop(key, None)
                
                st.success("✅ All documents generated and saved to history! You can view them in the 'Export' or 'Job History' tabs.")
                st.rerun()

# ============================================================
# TAB 3: EXPORT
# ============================================================
with tab3:
    st.subheader("📄 Preview & Download Saved Documents")
    
    # Get latest application from history
    latest = get_latest_application()
    
    if latest is None:
        st.info("No saved applications found. Generate documents in the 'Apply' tab first.")
    else:
        st.markdown(f"### 📌 Latest Application: {latest['job_title']} - {latest['company_name']}")
        st.write(f"**Applied Date:** {latest['applied_date']}")
        st.write(f"**Deadline:** {latest['deadline_date']}")
        st.write(f"**Status:** {latest['status']}")
        
        st.markdown("---")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if latest['generated_cv']:
                st.text_area("📄 CV", latest['generated_cv'], height=300, key="cv_export")
                st.download_button(
                    "⬇️ Download CV",
                    data=latest['generated_cv'],
                    file_name=f"CV_{latest['job_title']}_{datetime.now().strftime('%Y%m%d')}.txt"
                )
        
        with col2:
            if latest['generated_cover_letter']:
                st.text_area("✉️ Cover Letter", latest['generated_cover_letter'], height=300, key="cl_export")
                st.download_button(
                    "⬇️ Download Cover Letter",
                    data=latest['generated_cover_letter'],
                    file_name=f"Cover_Letter_{latest['job_title']}_{datetime.now().strftime('%Y%m%d')}.txt"
                )
        
        with col3:
            if latest['generated_motivation_letter']:
                st.text_area("📨 Motivation Letter", latest['generated_motivation_letter'], height=300, key="ml_export")
                st.download_button(
                    "⬇️ Download Motivation Letter",
                    data=latest['generated_motivation_letter'],
                    file_name=f"Motivation_Letter_{latest['job_title']}_{datetime.now().strftime('%Y%m%d')}.txt"
                )
        
        st.markdown("---")
        st.info("💡 To view previous applications, go to the 'Job History' tab.")

# ============================================================
# TAB 4: TEMPLATES
# ============================================================
with tab4:
    st.subheader("📝 Ready-to-Use Outreach Templates")
    templates_dir = os.path.join(os.path.dirname(__file__), "templates")
    if os.path.exists(templates_dir):
        template_files = {
            "LinkedIn Message": "linkedin.txt",
            "Networking Email": "email.txt",
            "Cover Letter": "cover_letter.txt",
            "Motivation Letter": "motivation_letter.txt"
        }
        for display_name, filename in template_files.items():
            file_path = os.path.join(templates_dir, filename)
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                with st.expander(f"📄 {display_name}", expanded=False):
                    st.text_area(f"{display_name}", content, height=150, key=f"template_{filename}")
                    st.download_button(
                        label=f"⬇️ Download {display_name}",
                        data=content,
                        file_name=filename,
                        key=f"download_{filename}"
                    )
            else:
                st.warning(f"File {filename} not found.")
    else:
        st.error("❌ Templates folder not found.")
        st.info("Create a folder called `templates` with the .txt files.")

# ============================================================
# TAB 5: JOB HISTORY
# ============================================================
with tab5:
    st.subheader("📅 Job Application History")
    df_history = get_job_history()
    if df_history.empty:
        st.info("No applications saved yet. Generate documents in the 'Apply' tab.")
    else:
        total = len(df_history)
        saved = len(df_history[df_history['status'] == 'Saved'])
        applied = len(df_history[df_history['status'] == 'Applied'])
        col1, col2, col3 = st.columns(3)
        col1.metric("📌 Total", total)
        col2.metric("💾 Saved", saved)
        col3.metric("✅ Applied", applied)
        st.markdown("---")
        
        st.subheader("📋 All Applications")
        if 'deadline_date' in df_history.columns:
            df_history['deadline_date'] = pd.to_datetime(df_history['deadline_date']).dt.date
            df_history['applied_date_display'] = pd.to_datetime(df_history['applied_date']).dt.strftime("%Y-%m-%d")
        
        display_cols = ['id', 'job_title', 'company_name', 'applied_date', 'deadline_date', 'status']
        st.dataframe(df_history[display_cols], use_container_width=True)
        
        selected_id = st.selectbox("Select Application ID to manage", df_history['id'].tolist())
        if selected_id:
            row = df_history[df_history['id'] == selected_id].iloc[0]
            with st.expander(f"📄 {row['job_title']} - {row['company_name']}", expanded=True):
                st.write(f"**Applied Date:** {row['applied_date']}")
                st.write(f"**Deadline:** {row['deadline_date']}")
                st.write(f"**Status:** {row['status']}")
                
                if 'generated_cv' in row and row['generated_cv']:
                    with st.expander("📄 View CV"):
                        st.text(row['generated_cv'][:500] + "...")
                
                if 'generated_cover_letter' in row and row['generated_cover_letter']:
                    with st.expander("✉️ View Cover Letter"):
                        st.text(row['generated_cover_letter'][:500] + "...")
                
                if 'generated_motivation_letter' in row and row['generated_motivation_letter']:
                    with st.expander("📨 View Motivation Letter"):
                        st.text(row['generated_motivation_letter'][:500] + "...")
                
                st.markdown("---")
                col_action1, col_action2, col_action3 = st.columns(3)
                with col_action1:
                    new_status = st.selectbox("Update Status", ["Saved", "Applied", "Rejected", "Interview"], key=f"status_{selected_id}")
                    if st.button("Update Status", key=f"update_{selected_id}"):
                        update_status(selected_id, new_status)
                        st.success("✅ Status updated!")
                        st.rerun()
                with col_action2:
                    if st.button("🗑️ Delete Application", key=f"delete_{selected_id}"):
                        delete_application(selected_id)
                        st.success("✅ Application deleted!")
                        st.rerun()
                with col_action3:
                    if st.button("📅 Set Reminder", key=f"reminder_{selected_id}"):
                        st.info(f"⏰ Reminder set for {row['deadline_date']}")

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption(f"⚡ Data stored in {DB_PATH} | Dynamic AI Generation | Auto-save to history")
