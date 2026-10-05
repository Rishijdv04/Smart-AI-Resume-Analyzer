"""
Smart Resume AI - Main Application
"""
import io
import os
import re
import time
import json
import base64
import traceback
from datetime import datetime, timedelta

import requests
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from PIL import Image
from streamlit_lottie import st_lottie
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

from jobs.job_search import render_job_search
from ui_components import (
    apply_modern_styles, hero_section, feature_card, about_section,
    page_header, render_analytics_section, render_activity_section,
    render_suggestions_section
)
from feedback.feedback import FeedbackManager
from dashboard.dashboard import DashboardManager
from config.courses import (
    COURSES_BY_CATEGORY, RESUME_VIDEOS, INTERVIEW_VIDEOS,
    get_courses_for_role, get_category_for_role
)
from config.job_roles import JOB_ROLES
from config.database import (
    get_database_connection, save_resume_data, save_analysis_data,
    init_database, verify_admin, log_admin_action, save_ai_analysis_data,
    get_ai_analysis_stats, reset_ai_analysis_stats, get_detailed_ai_analysis_stats
)
from utils.ai_resume_analyzer import AIResumeAnalyzer
from utils.resume_builder import ResumeBuilder
from utils.resume_analyzer import ResumeAnalyzer

# Set page config at the very beginning
st.set_page_config(
    page_title="Smart Resume AI",
    page_icon="🚀",
    layout="wide"
)

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def score_status(score):
    """Text label for a 0-100 score."""
    if score >= 80:
        return "Excellent"
    if score >= 60:
        return "Good"
    return "Needs Improvement"


def score_color(score):
    """Hex colour for a 0-100 score."""
    if score >= 80:
        return "#4CAF50"
    if score >= 60:
        return "#FFA500"
    return "#FF4444"


def gauge_figure(value, title):
    """Gauge chart used in the AI analysis report."""
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={'text': title, 'font': {'size': 16}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': score_color(value)},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "gray",
            'steps': [
                {'range': [0, 40], 'color': 'rgba(255, 68, 68, 0.2)'},
                {'range': [40, 60], 'color': 'rgba(255, 165, 0, 0.2)'},
                {'range': [60, 80], 'color': 'rgba(255, 214, 0, 0.2)'},
                {'range': [80, 100], 'color': 'rgba(76, 175, 80, 0.2)'}
            ],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.75,
                'value': 60
            }
        }
    ))
    fig.update_layout(height=250, margin=dict(l=20, r=20, t=50, b=20))
    return fig


def style_dark_chart(fig, height, title=None):
    """Apply the shared dark/transparent look to a plotly figure."""
    layout = dict(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color="#ffffff", size=14),
        height=height,
        margin=dict(l=20, r=20, t=50, b=30),
    )
    if title:
        layout['title'] = {
            'text': title, 'y': 0.95, 'x': 0.5,
            'xanchor': 'center', 'yanchor': 'top',
            'font': {'size': 18, 'color': 'white'}
        }
    fig.update_layout(**layout)


def suggestion_block(title, lis_html):
    """Render one titled list of suggestions as a single HTML block."""
    st.markdown(
        f"<div style='background-color: #1e1e1e; padding: 15px; border-radius: 10px; margin: 10px 0;'>"
        f"<h3 style='color: #4CAF50; margin-bottom: 10px;'>{title}</h3>"
        f"<ul style='list-style-type: none; padding-left: 0;'>{lis_html}</ul></div>",
        unsafe_allow_html=True
    )


def li_items(items):
    return "".join(f"<li style='margin-bottom: 8px;'>✓ {s}</li>" for s in items)


# Section header styles for the AI analysis report: header -> (icon, colour1, colour2)
SECTION_STYLES = {
    "## Overall Assessment": ("fa-chart-line", "#1e3a8a", "#3b82f6"),
    "## Professional Profile Analysis": ("fa-user-tie", "#047857", "#10b981"),
    "## Skills Analysis": ("fa-tools", "#4f46e5", "#818cf8"),
    "## Experience Analysis": ("fa-briefcase", "#9f1239", "#e11d48"),
    "## Education Analysis": ("fa-graduation-cap", "#854d0e", "#eab308"),
    "## Key Strengths": ("fa-check-circle", "#166534", "#22c55e"),
    "## Areas for Improvement": ("fa-exclamation-circle", "#9f1239", "#fb7185"),
    "## ATS Optimization Assessment": ("fa-robot", "#0e7490", "#06b6d4"),
    "## Recommended Courses": ("fa-book", "#5b21b6", "#8b5cf6"),
    "## Resume Score": ("fa-star", "#0369a1", "#0ea5e9"),
    "## Role Alignment Analysis": ("fa-bullseye", "#7c2d12", "#ea580c"),
    "## Job Match Analysis": ("fa-handshake", "#4d7c0f", "#84cc16"),
}


def format_analysis_report(text):
    """Wrap known '## Section' headings of the AI report in styled HTML blocks."""
    out = []
    section_open = False
    for line in text.split("\n"):
        key = line.strip()
        if key in SECTION_STYLES:
            if section_open:
                out.extend(["", "</div></div>"])
            icon, c1, c2 = SECTION_STYLES[key]
            title = key[3:]
            out.append(
                '<div class="report-section">'
                f'<h3 style="background: linear-gradient(90deg, {c1}, {c2}); color: white; '
                f'padding: 10px; border-radius: 5px;"><i class="fas {icon}"></i> {title}</h3>'
                '<div class="section-content">'
            )
            out.append("")
            section_open = True
        else:
            out.append(line)
    if section_open:
        out.extend(["", "</div></div>"])
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------
class ResumeApp:
    def __init__(self):
        """Initialize the application"""
        if 'form_data' not in st.session_state:
            st.session_state.form_data = {
                'personal_info': {
                    'full_name': '',
                    'email': '',
                    'phone': '',
                    'location': '',
                    'linkedin': '',
                    'portfolio': ''
                },
                'summary': '',
                'experiences': [],
                'education': [],
                'projects': [],
                'skills_categories': {
                    'technical': [],
                    'soft': [],
                    'languages': [],
                    'tools': []
                }
            }

        # Initialize navigation state
        if 'page' not in st.session_state:
            st.session_state.page = 'home'

        # Initialize admin state
        if 'is_admin' not in st.session_state:
            st.session_state.is_admin = False

        self.pages = {
            "🏠 HOME": self.render_home,
            "🔍 RESUME ANALYZER": self.render_analyzer,
            "📝 RESUME BUILDER": self.render_builder,
            "📊 DASHBOARD": self.render_dashboard,
            "🎯 JOB SEARCH": self.render_job_search,
            "💬 FEEDBACK": self.render_feedback_page,
            "ℹ️ ABOUT": self.render_about
        }

        # Initialize dashboard manager
        self.dashboard_manager = DashboardManager()

        self.analyzer = ResumeAnalyzer()
        self.ai_analyzer = AIResumeAnalyzer()
        self.builder = ResumeBuilder()
        self.job_roles = JOB_ROLES

        # Initialize session state
        if 'user_id' not in st.session_state:
            st.session_state.user_id = 'default_user'
        if 'selected_role' not in st.session_state:
            st.session_state.selected_role = None

        # Initialize database
        init_database()

        # Load external CSS
        with open('style/style.css') as f:
            st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

        # Load Google Fonts
        st.markdown("""
            <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        """, unsafe_allow_html=True)

        if 'resume_data' not in st.session_state:
            st.session_state.resume_data = []
        if 'ai_analysis_stats' not in st.session_state:
            st.session_state.ai_analysis_stats = {
                'score_distribution': {},
                'total_analyses': 0,
                'average_score': 0
            }

    # ------------------------------------------------------------------
    # Utilities
    # ------------------------------------------------------------------
    def load_lottie_url(self, url: str):
        """Load Lottie animation from URL (returns None if offline/unavailable)"""
        try:
            r = requests.get(url, timeout=5)
            if r.status_code != 200:
                return None
            return r.json()
        except Exception as e:
            print(f"Could not load lottie animation: {e}")
            return None

    def apply_global_styles(self):
        st.markdown("""
        <style>
        /* Custom Scrollbar */
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: #1a1a1a; border-radius: 4px; }
        ::-webkit-scrollbar-thumb { background: #4CAF50; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #45a049; }

        /* Global Styles */
        .main-header {
            background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
            padding: 2rem; border-radius: 15px; margin-bottom: 2rem;
            box-shadow: 0 10px 20px rgba(0,0,0,0.2); text-align: center;
            position: relative; overflow: hidden;
        }
        .main-header::before {
            content: ''; position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            background: linear-gradient(45deg, transparent 0%, rgba(255,255,255,0.1) 100%);
            z-index: 1;
        }
        .main-header h1 {
            color: white; font-size: 2.5rem; font-weight: 600; margin: 0;
            position: relative; z-index: 2;
        }

        /* Template Card Styles */
        .template-container {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
            gap: 2rem; padding: 1rem;
        }
        .template-card {
            background: rgba(45, 45, 45, 0.9); border-radius: 20px; padding: 2rem;
            position: relative; overflow: hidden; backdrop-filter: blur(10px);
            border: 1px solid rgba(255,255,255,0.1);
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .template-card:hover {
            transform: translateY(-10px); box-shadow: 0 20px 40px rgba(0,0,0,0.3);
            border-color: #4CAF50;
        }
        .template-card::before {
            content: ''; position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            background: linear-gradient(45deg, transparent 0%, rgba(76,175,80,0.1) 100%);
            z-index: 1;
        }
        .template-icon { font-size: 3rem; color: #4CAF50; margin-bottom: 1.5rem; position: relative; z-index: 2; }
        .template-title { font-size: 1.8rem; font-weight: 600; color: white; margin-bottom: 1rem; position: relative; z-index: 2; }
        .template-description { color: #aaa; margin-bottom: 1.5rem; position: relative; z-index: 2; line-height: 1.6; }

        /* Feature List Styles */
        .feature-list { list-style: none; padding: 0; margin: 1.5rem 0; position: relative; z-index: 2; }
        .feature-item { display: flex; align-items: center; margin-bottom: 1rem; color: #ddd; font-size: 0.95rem; }
        .feature-icon { color: #4CAF50; margin-right: 0.8rem; font-size: 1.1rem; }

        /* Button Styles */
        .action-button {
            background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
            color: white; padding: 1rem 2rem; border-radius: 50px; border: none;
            font-weight: 500; cursor: pointer; width: 100%; text-align: center;
            position: relative; overflow: hidden; z-index: 2;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .action-button:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(76,175,80,0.3); }
        .action-button::before {
            content: ''; position: absolute; top: 0; left: -100%; width: 100%; height: 100%;
            background: linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.2) 50%, transparent 100%);
            transition: all 0.6s ease;
        }
        .action-button:hover::before { left: 100%; }

        /* Form Section Styles */
        .form-section {
            background: rgba(45, 45, 45, 0.9); border-radius: 20px; padding: 2rem;
            margin: 2rem 0; backdrop-filter: blur(10px); border: 1px solid rgba(255,255,255,0.1);
        }
        .form-section-title {
            font-size: 1.8rem; font-weight: 600; color: white; margin-bottom: 1.5rem;
            padding-bottom: 0.8rem; border-bottom: 2px solid #4CAF50;
        }
        .form-group { margin-bottom: 1.5rem; }
        .form-label { color: #ddd; font-weight: 500; margin-bottom: 0.8rem; display: block; }
        .form-input {
            width: 100%; padding: 1rem; border-radius: 10px;
            border: 1px solid rgba(255,255,255,0.1); background: rgba(30, 30, 30, 0.9);
            color: white; transition: all 0.3s ease;
        }
        .form-input:focus { border-color: #4CAF50; box-shadow: 0 0 0 2px rgba(76,175,80,0.2); outline: none; }

        /* Skill Tags */
        .skill-tag-container { display: flex; flex-wrap: wrap; gap: 0.8rem; margin-top: 1rem; }
        .skill-tag {
            background: rgba(76,175,80,0.1); color: #4CAF50; padding: 0.6rem 1.2rem;
            border-radius: 50px; border: 1px solid #4CAF50; font-size: 0.9rem;
            transition: all 0.3s ease; cursor: pointer;
        }
        .skill-tag:hover {
            background: #4CAF50; color: white; transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(76,175,80,0.2);
        }

        /* Progress Circle */
        .progress-container { position: relative; width: 150px; height: 150px; margin: 2rem auto; }
        .progress-circle { transform: rotate(-90deg); width: 100%; height: 100%; }
        .progress-circle circle {
            fill: none; stroke-width: 8; stroke-linecap: round; stroke: #4CAF50;
            transform-origin: 50% 50%; transition: all 0.3s ease;
        }
        .progress-text {
            position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
            font-size: 1.5rem; font-weight: 600; color: white;
        }
        .main .block-container { padding-top: 2rem; padding-bottom: 2rem; }
        .feature-card {
            background-color: #1e1e1e; border-radius: 10px; padding: 20px;
            margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }

        /* Animations */
        @keyframes slideIn {
            from { opacity: 0; transform: translateY(30px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .animate-slide-in { animation: slideIn 0.6s cubic-bezier(0.4, 0, 0.2, 1) forwards; }

        /* Responsive Design */
        @media (max-width: 768px) {
            .template-container { grid-template-columns: 1fr; }
            .main-header { padding: 1.5rem; }
            .main-header h1 { font-size: 2rem; }
            .template-card { padding: 1.5rem; }
            .action-button { padding: 0.8rem 1.6rem; }
        }
        </style>
        """, unsafe_allow_html=True)

    def load_image(self, image_name):
        """Load an image from the local assets directory as a base64 data URI"""
        try:
            image_path = os.path.join(os.path.dirname(__file__), "assets", image_name)
            with open(image_path, "rb") as f:
                image_bytes = f.read()
            encoded = base64.b64encode(image_bytes).decode()
            return f"data:image/png;base64,{encoded}"
        except Exception as e:
            print(f"Error loading image {image_name}: {e}")
            return None

    def export_to_excel(self):
        """Export resume data to Excel"""
        conn = get_database_connection()

        # Get resume data with analysis
        query = """
            SELECT
                rd.name, rd.email, rd.phone, rd.linkedin, rd.github, rd.portfolio,
                rd.summary, rd.target_role, rd.target_category,
                rd.education, rd.experience, rd.projects, rd.skills,
                ra.ats_score, ra.keyword_match_score, ra.format_score, ra.section_score,
                ra.missing_skills, ra.recommendations,
                rd.created_at
            FROM resume_data rd
            LEFT JOIN resume_analysis ra ON rd.id = ra.resume_id
        """

        try:
            df = pd.read_sql_query(query, conn)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Resume Data')
            return output.getvalue()
        except Exception as e:
            print(f"Error exporting to Excel: {str(e)}")
            return None
        finally:
            conn.close()

    def render_dashboard(self):
        """Render the dashboard page"""
        self.dashboard_manager.render_dashboard()

        st.toast("Check out these repositories: [Awesome Hacking](https://github.com/Hunterdii/Awesome-Hacking)", icon="ℹ️")

    def render_empty_state(self, icon, message):
        """Render an empty state with icon and message"""
        return f"""
            <div style='text-align: center; padding: 2rem; color: #666;'>
                <i class='{icon}' style='font-size: 2rem; margin-bottom: 1rem; color: #00bfa5;'></i>
                <p style='margin: 0;'>{message}</p>
            </div>
        """

    def extract_resume_text(self, uploaded_file, analyzer=None):
        """Extract text from an uploaded PDF / DOCX / text file"""
        analyzer = analyzer or self.analyzer
        if uploaded_file.type == "application/pdf":
            return analyzer.extract_text_from_pdf(uploaded_file)
        if uploaded_file.type == DOCX_MIME:
            return analyzer.extract_text_from_docx(uploaded_file)
        return uploaded_file.getvalue().decode('utf-8', errors='ignore')

    def analyze_resume(self, resume_text):
        """Analyze resume and store results"""
        analytics = self.analyzer.analyze_resume(resume_text)
        st.session_state.analytics_data = analytics
        return analytics

    def handle_resume_upload(self):
        """Handle resume upload and analysis"""
        uploaded_file = st.file_uploader("Upload your resume", type=['pdf', 'docx'])

        if uploaded_file is not None:
            try:
                # Extract text from resume (FIX: use the analyzer's methods)
                resume_text = self.extract_resume_text(uploaded_file)

                # Store resume data
                st.session_state.resume_data = {
                    'filename': uploaded_file.name,
                    'content': resume_text,
                    'upload_time': datetime.now().isoformat()
                }

                # Analyze resume
                self.analyze_resume(resume_text)
                return True
            except Exception as e:
                st.error(f"Error processing resume: {str(e)}")
                return False
        return False

    # ------------------------------------------------------------------
    # Resume builder
    # ------------------------------------------------------------------
    def render_builder(self):
        st.title("Resume Builder 📝")
        st.write("Create your professional resume")

        # Template selection
        template_options = ["Modern", "Professional", "Minimal", "Creative"]
        selected_template = st.selectbox("Select Resume Template", template_options)
        st.success(f"🎨 Currently using: {selected_template} Template")

        # Personal Information
        st.subheader("Personal Information")

        col1, col2 = st.columns(2)
        with col1:
            existing_name = st.session_state.form_data['personal_info']['full_name']
            existing_email = st.session_state.form_data['personal_info']['email']
            existing_phone = st.session_state.form_data['personal_info']['phone']

            full_name = st.text_input("Full Name", value=existing_name)
            email = st.text_input("Email", value=existing_email, key="email_input")
            phone = st.text_input("Phone", value=existing_phone)

            # Immediately update session state after email input
            if 'email_input' in st.session_state:
                st.session_state.form_data['personal_info']['email'] = st.session_state.email_input

        with col2:
            existing_location = st.session_state.form_data['personal_info']['location']
            existing_linkedin = st.session_state.form_data['personal_info']['linkedin']
            existing_portfolio = st.session_state.form_data['personal_info']['portfolio']

            location = st.text_input("Location", value=existing_location)
            linkedin = st.text_input("LinkedIn URL", value=existing_linkedin)
            portfolio = st.text_input("Portfolio Website", value=existing_portfolio)

        # Update personal info in session state
        st.session_state.form_data['personal_info'] = {
            'full_name': full_name,
            'email': email,
            'phone': phone,
            'location': location,
            'linkedin': linkedin,
            'portfolio': portfolio
        }

        # Professional Summary
        st.subheader("Professional Summary")
        summary = st.text_area(
            "Professional Summary",
            value=st.session_state.form_data.get('summary', ''),
            height=150,
            help="Write a brief summary highlighting your key skills and experience"
        )

        # Experience Section
        st.subheader("Work Experience")
        if 'experiences' not in st.session_state.form_data:
            st.session_state.form_data['experiences'] = []

        if st.button("Add Experience"):
            st.session_state.form_data['experiences'].append({
                'company': '',
                'position': '',
                'start_date': '',
                'end_date': '',
                'description': '',
                'responsibilities': [],
                'achievements': []
            })

        for idx, exp in enumerate(st.session_state.form_data['experiences']):
            with st.expander(f"Experience {idx + 1}", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    exp['company'] = st.text_input(
                        "Company Name", key=f"company_{idx}", value=exp.get('company', ''))
                    exp['position'] = st.text_input(
                        "Position", key=f"position_{idx}", value=exp.get('position', ''))
                with col2:
                    exp['start_date'] = st.text_input(
                        "Start Date", key=f"start_date_{idx}", value=exp.get('start_date', ''))
                    exp['end_date'] = st.text_input(
                        "End Date", key=f"end_date_{idx}", value=exp.get('end_date', ''))

                exp['description'] = st.text_area(
                    "Role Overview", key=f"desc_{idx}",
                    value=exp.get('description', ''),
                    help="Brief overview of your role and impact")

                # Responsibilities
                st.markdown("##### Key Responsibilities")
                resp_text = st.text_area(
                    "Enter responsibilities (one per line)",
                    key=f"resp_{idx}",
                    value='\n'.join(exp.get('responsibilities', [])),
                    height=100,
                    help="List your main responsibilities, one per line")
                exp['responsibilities'] = [r.strip() for r in resp_text.split('\n') if r.strip()]

                # Achievements
                st.markdown("##### Key Achievements")
                achv_text = st.text_area(
                    "Enter achievements (one per line)",
                    key=f"achv_{idx}",
                    value='\n'.join(exp.get('achievements', [])),
                    height=100,
                    help="List your notable achievements, one per line")
                exp['achievements'] = [a.strip() for a in achv_text.split('\n') if a.strip()]

                if st.button("Remove Experience", key=f"remove_exp_{idx}"):
                    st.session_state.form_data['experiences'].pop(idx)
                    st.rerun()

        # Projects Section
        st.subheader("Projects")
        if 'projects' not in st.session_state.form_data:
            st.session_state.form_data['projects'] = []

        if st.button("Add Project"):
            st.session_state.form_data['projects'].append({
                'name': '',
                'technologies': '',
                'description': '',
                'responsibilities': [],
                'achievements': [],
                'link': ''
            })

        for idx, proj in enumerate(st.session_state.form_data['projects']):
            with st.expander(f"Project {idx + 1}", expanded=True):
                proj['name'] = st.text_input(
                    "Project Name", key=f"proj_name_{idx}", value=proj.get('name', ''))
                proj['technologies'] = st.text_input(
                    "Technologies Used", key=f"proj_tech_{idx}",
                    value=proj.get('technologies', ''),
                    help="List the main technologies, frameworks, and tools used")

                proj['description'] = st.text_area(
                    "Project Overview", key=f"proj_desc_{idx}",
                    value=proj.get('description', ''),
                    help="Brief overview of the project and its goals")

                # Project Responsibilities
                st.markdown("##### Key Responsibilities")
                proj_resp_text = st.text_area(
                    "Enter responsibilities (one per line)",
                    key=f"proj_resp_{idx}",
                    value='\n'.join(proj.get('responsibilities', [])),
                    height=100,
                    help="List your main responsibilities in the project")
                proj['responsibilities'] = [r.strip() for r in proj_resp_text.split('\n') if r.strip()]

                # Project Achievements
                st.markdown("##### Key Achievements")
                proj_achv_text = st.text_area(
                    "Enter achievements (one per line)",
                    key=f"proj_achv_{idx}",
                    value='\n'.join(proj.get('achievements', [])),
                    height=100,
                    help="List the project's key achievements and your contributions")
                proj['achievements'] = [a.strip() for a in proj_achv_text.split('\n') if a.strip()]

                proj['link'] = st.text_input(
                    "Project Link (optional)", key=f"proj_link_{idx}",
                    value=proj.get('link', ''),
                    help="Link to the project repository, demo, or documentation")

                if st.button("Remove Project", key=f"remove_proj_{idx}"):
                    st.session_state.form_data['projects'].pop(idx)
                    st.rerun()

        # Education Section
        st.subheader("Education")
        if 'education' not in st.session_state.form_data:
            st.session_state.form_data['education'] = []

        if st.button("Add Education"):
            st.session_state.form_data['education'].append({
                'school': '',
                'degree': '',
                'field': '',
                'graduation_date': '',
                'gpa': '',
                'achievements': []
            })

        for idx, edu in enumerate(st.session_state.form_data['education']):
            with st.expander(f"Education {idx + 1}", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    edu['school'] = st.text_input(
                        "School/University", key=f"school_{idx}", value=edu.get('school', ''))
                    edu['degree'] = st.text_input(
                        "Degree", key=f"degree_{idx}", value=edu.get('degree', ''))
                with col2:
                    edu['field'] = st.text_input(
                        "Field of Study", key=f"field_{idx}", value=edu.get('field', ''))
                    edu['graduation_date'] = st.text_input(
                        "Graduation Date", key=f"grad_date_{idx}",
                        value=edu.get('graduation_date', ''))

                edu['gpa'] = st.text_input(
                    "GPA (optional)", key=f"gpa_{idx}", value=edu.get('gpa', ''))

                # Educational Achievements
                st.markdown("##### Achievements & Activities")
                edu_achv_text = st.text_area(
                    "Enter achievements (one per line)",
                    key=f"edu_achv_{idx}",
                    value='\n'.join(edu.get('achievements', [])),
                    height=100,
                    help="List academic achievements, relevant coursework, or activities")
                edu['achievements'] = [a.strip() for a in edu_achv_text.split('\n') if a.strip()]

                if st.button("Remove Education", key=f"remove_edu_{idx}"):
                    st.session_state.form_data['education'].pop(idx)
                    st.rerun()

        # Skills Section
        st.subheader("Skills")
        if 'skills_categories' not in st.session_state.form_data:
            st.session_state.form_data['skills_categories'] = {
                'technical': [],
                'soft': [],
                'languages': [],
                'tools': []
            }

        skills = st.session_state.form_data['skills_categories']

        col1, col2 = st.columns(2)
        with col1:
            tech_skills = st.text_area(
                "Technical Skills (one per line)",
                value='\n'.join(skills['technical']),
                height=150,
                help="Programming languages, frameworks, databases, etc.")
            skills['technical'] = [s.strip() for s in tech_skills.split('\n') if s.strip()]

            soft_skills = st.text_area(
                "Soft Skills (one per line)",
                value='\n'.join(skills['soft']),
                height=150,
                help="Leadership, communication, problem-solving, etc.")
            skills['soft'] = [s.strip() for s in soft_skills.split('\n') if s.strip()]

        with col2:
            languages = st.text_area(
                "Languages (one per line)",
                value='\n'.join(skills['languages']),
                height=150,
                help="Programming or human languages with proficiency level")
            skills['languages'] = [x.strip() for x in languages.split('\n') if x.strip()]

            tools = st.text_area(
                "Tools & Technologies (one per line)",
                value='\n'.join(skills['tools']),
                height=150,
                help="Development tools, software, platforms, etc.")
            skills['tools'] = [t.strip() for t in tools.split('\n') if t.strip()]

        # Update form data in session state
        st.session_state.form_data.update({'summary': summary})

        # Generate Resume button
        if st.button("Generate Resume 📄", type="primary"):
            print("Validating form data...")
            print(f"Session state form data: {st.session_state.form_data}")
            print(f"Email input value: {st.session_state.get('email_input', '')}")

            # Get the current values from form
            current_name = st.session_state.form_data['personal_info']['full_name'].strip()
            current_email = st.session_state.email_input if 'email_input' in st.session_state else ''

            print(f"Current name: {current_name}")
            print(f"Current email: {current_email}")

            # Validate required fields
            if not current_name:
                st.error("⚠️ Please enter your full name.")
                return

            if not current_email:
                st.error("⚠️ Please enter your email address.")
                return

            # Update email in form data one final time
            st.session_state.form_data['personal_info']['email'] = current_email

            try:
                print("Preparing resume data...")
                resume_data = {
                    "personal_info": st.session_state.form_data['personal_info'],
                    "summary": st.session_state.form_data.get('summary', '').strip(),
                    "experience": st.session_state.form_data.get('experiences', []),
                    "education": st.session_state.form_data.get('education', []),
                    "projects": st.session_state.form_data.get('projects', []),
                    "skills": st.session_state.form_data.get('skills_categories', {
                        'technical': [],
                        'soft': [],
                        'languages': [],
                        'tools': []
                    }),
                    "template": selected_template
                }

                print(f"Resume data prepared: {resume_data}")

                try:
                    resume_buffer = self.builder.generate_resume(resume_data)
                    if resume_buffer:
                        # FIX: build the file name outside the f-string (multi-line
                        # f-strings are a SyntaxError on Python < 3.12)
                        safe_name = current_name.replace(' ', '_')
                        download_name = f"{safe_name}_resume.docx"

                        try:
                            # Save resume data to database
                            save_resume_data(resume_data)

                            st.success("✅ Resume generated successfully!")
                            st.snow()
                        except Exception as db_error:
                            print(f"Warning: Failed to save to database: {str(db_error)}")
                            # Still allow download even if database save fails
                            st.warning("⚠️ Resume generated but couldn't be saved to database")
                            st.balloons()

                        st.download_button(
                            label="Download Resume 📥",
                            data=resume_buffer,
                            file_name=download_name,
                            mime=DOCX_MIME,
                            on_click=lambda: st.balloons()
                        )
                    else:
                        st.error("❌ Failed to generate resume. Please try again.")
                        print("Resume buffer was None")
                except Exception as gen_error:
                    print(f"Error during resume generation: {str(gen_error)}")
                    print(f"Full traceback: {traceback.format_exc()}")
                    st.error(f"❌ Error generating resume: {str(gen_error)}")

            except Exception as e:
                print(f"Error preparing resume data: {str(e)}")
                print(f"Full traceback: {traceback.format_exc()}")
                st.error(f"❌ Error preparing resume data: {str(e)}")

        st.toast("Check out these repositories: [30-Days-Of-Rust](https://github.com/Hunterdii/30-Days-Of-Rust)", icon="ℹ️")

    # ------------------------------------------------------------------
    # About page
    # ------------------------------------------------------------------
    def render_about(self):
        """Render the about page"""

        def get_image_as_base64(file_path):
            try:
                with open(file_path, "rb") as image_file:
                    encoded = base64.b64encode(image_file.read()).decode()
                    return f"data:image/jpeg;base64,{encoded}"
            except Exception:
                return None

        image_path = os.path.join(os.path.dirname(__file__), "assets", "124852522.jpeg")
        image_base64 = get_image_as_base64(image_path)

        apply_modern_styles()

        # Add Font Awesome icons and custom CSS
        st.markdown("""
            <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
            <style>
                .profile-section, .vision-section, .feature-card {
                    text-align: center; padding: 2rem; background: rgba(45, 45, 45, 0.9);
                    border-radius: 20px; margin: 2rem auto; max-width: 800px;
                }
                .profile-image {
                    width: 200px; height: 200px; border-radius: 50%; margin: 0 auto 1.5rem;
                    display: block; object-fit: cover; border: 4px solid #4CAF50;
                }
                .profile-name { font-size: 2.5rem; color: white; margin-bottom: 0.5rem; }
                .profile-title { font-size: 1.2rem; color: #4CAF50; margin-bottom: 1.5rem; }
                .social-links { display: flex; justify-content: center; gap: 1.5rem; margin: 2rem 0; }
                .social-link {
                    font-size: 2rem; color: #4CAF50; transition: all 0.3s ease; padding: 0.5rem;
                    border-radius: 50%; background: rgba(76, 175, 80, 0.1); width: 60px; height: 60px;
                    display: flex; align-items: center; justify-content: center; text-decoration: none;
                }
                .social-link:hover {
                    transform: translateY(-5px); background: #4CAF50; color: white;
                    box-shadow: 0 5px 15px rgba(76, 175, 80, 0.3);
                }
                .bio-text { color: #ddd; line-height: 1.8; font-size: 1.1rem; margin-top: 2rem; text-align: left; }
                .vision-text {
                    color: #ddd; line-height: 1.8; font-size: 1.1rem; font-style: italic;
                    margin: 1.5rem 0; text-align: left;
                }
                .vision-icon { font-size: 2.5rem; color: #4CAF50; margin-bottom: 1rem; }
                .vision-title { font-size: 2rem; color: white; margin-bottom: 1rem; }
                .features-grid {
                    display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                    gap: 2rem; margin: 2rem auto; max-width: 1200px;
                }
                .feature-card { padding: 2rem; margin: 0; }
                .feature-icon { font-size: 2.5rem; color: #4CAF50; margin-bottom: 1rem; }
                .feature-title { font-size: 1.5rem; color: white; margin: 1rem 0; }
                .feature-description { color: #ddd; line-height: 1.6; }
            </style>
        """, unsafe_allow_html=True)

        # Hero Section
        st.markdown("""
            <div class="hero-section">
                <h1 class="hero-title">About Smart Resume AI</h1>
                <p class="hero-subtitle">A powerful AI-driven platform for optimizing your resume</p>
            </div>
        """, unsafe_allow_html=True)

        # Profile Section
        profile_img = image_base64 if image_base64 else 'https://avatars.githubusercontent.com/Hunterdii'
        st.markdown(f"""
            <div class="profile-section">
                <img src="{profile_img}"
                     alt="Het Patel"
                     class="profile-image"
                     onerror="this.onerror=null; this.src='https://avatars.githubusercontent.com/Hunterdii';">
                <h2 class="profile-name">Het Patel (Hunterdii)</h2>
                <p class="profile-title">Full Stack Developer & AI/ML Enthusiast</p>
                <div class="social-links">
                    <a href="https://github.com/Hunterdii" class="social-link" target="_blank">
                        <i class="fab fa-github"></i>
                    </a>
                    <a href="https://www.linkedin.com/in/patel-hetkumar-sandipbhai-8b110525a/" class="social-link" target="_blank">
                        <i class="fab fa-linkedin"></i>
                    </a>
                    <a href="mailto:hunterdii9879@gmail.com" class="social-link" target="_blank">
                        <i class="fas fa-envelope"></i>
                    </a>
                </div>
                <p class="bio-text">
                    Hello! I'm a passionate Full Stack Developer with expertise in AI and Machine Learning.
                    I created Smart Resume AI to revolutionize how job seekers approach their career journey.
                    With my background in both software development and AI, I've designed this platform to
                    provide intelligent, data-driven insights for resume optimization.
                </p>
            </div>
        """, unsafe_allow_html=True)

        # Vision Section
        st.markdown("""
            <div class="vision-section">
                <i class="fas fa-lightbulb vision-icon"></i>
                <h2 class="vision-title">Our Vision</h2>
                <p class="vision-text">
                    "Smart Resume AI represents my vision of democratizing career advancement through technology.
                    By combining cutting-edge AI with intuitive design, this platform empowers job seekers at
                    every career stage to showcase their true potential and stand out in today's competitive job market."
                </p>
            </div>
        """, unsafe_allow_html=True)

        # Features Section
        st.markdown("""
            <div class="features-grid">
                <div class="feature-card">
                    <i class="fas fa-robot feature-icon"></i>
                    <h3 class="feature-title">AI-Powered Analysis</h3>
                    <p class="feature-description">
                        Advanced AI algorithms provide detailed insights and suggestions to optimize your resume for maximum impact.
                    </p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-chart-line feature-icon"></i>
                    <h3 class="feature-title">Data-Driven Insights</h3>
                    <p class="feature-description">
                        Make informed decisions with our analytics-based recommendations and industry insights.
                    </p>
                </div>
                <div class="feature-card">
                    <i class="fas fa-shield-alt feature-icon"></i>
                    <h3 class="feature-title">Privacy First</h3>
                    <p class="feature-description">
                        Your data security is our priority. We ensure your information is always protected and private.
                    </p>
                </div>
            </div>
            <div style="text-align: center; margin: 3rem 0;">
                <a href="?page=analyzer" class="cta-button">
                    Start Your Journey
                    <i class="fas fa-arrow-right" style="margin-left: 10px;"></i>
                </a>
            </div>
        """, unsafe_allow_html=True)

        st.toast("Check out these repositories: [Iriswise](https://github.com/Hunterdii/Iriswise)", icon="ℹ️")

    # ------------------------------------------------------------------
    # Resume analyzer page
    # ------------------------------------------------------------------
    def render_analyzer(self):
        """Render the resume analyzer page"""
        apply_modern_styles()

        page_header(
            "Resume Analyzer",
            "Get instant AI-powered feedback to optimize your resume"
        )

        analyzer_tabs = st.tabs(["Standard Analyzer", "AI Analyzer"])

        with analyzer_tabs[0]:
            self.render_standard_analyzer()

        with analyzer_tabs[1]:
            self.render_ai_analyzer()

        st.toast("Check out these repositories: [Awesome Java](https://github.com/Hunterdii/Awesome-Java)", icon="ℹ️")

    def render_role_selector(self, key_prefix):
        """Job category/role selector + info card. Returns (category, role, role_info)"""
        categories = list(self.job_roles.keys())
        selected_category = st.selectbox("Job Category", categories, key=f"{key_prefix}_category")

        roles = list(self.job_roles[selected_category].keys())
        selected_role = st.selectbox("Specific Role", roles, key=f"{key_prefix}_role")

        role_info = self.job_roles[selected_category][selected_role]

        required = ', '.join(role_info['required_skills'])
        st.markdown(f"""
        <div style='background-color: #1e1e1e; padding: 20px; border-radius: 10px; margin: 10px 0;'>
            <h3>{selected_role}</h3>
            <p>{role_info['description']}</p>
            <h4>Required Skills:</h4>
            <p>{required}</p>
        </div>
        """, unsafe_allow_html=True)
        return selected_category, selected_role, role_info

    def render_standard_analyzer(self):
        selected_category, selected_role, role_info = self.render_role_selector("standard")

        uploaded_file = st.file_uploader(
            "Upload your resume", type=['pdf', 'docx'], key="standard_file")

        if not uploaded_file:
            st.markdown(
                self.render_empty_state(
                    "fas fa-cloud-upload-alt",
                    "Upload your resume to get started with standard analysis"
                ),
                unsafe_allow_html=True
            )
            return

        analyze_standard = st.button(
            "🔍 Analyze My Resume", type="primary", width='stretch',
            key="analyze_standard_button")

        if not analyze_standard:
            return

        with st.spinner("Analyzing your document..."):
            # Get file content
            text = ""
            try:
                if uploaded_file.type == "application/pdf":
                    try:
                        text = self.analyzer.extract_text_from_pdf(uploaded_file)
                    except Exception as pdf_error:
                        st.error(f"PDF extraction failed: {str(pdf_error)}")
                        st.info("Trying alternative PDF extraction method...")
                        try:
                            text = self.ai_analyzer.extract_text_from_pdf(uploaded_file)
                        except Exception as backup_error:
                            st.error(f"All PDF extraction methods failed: {str(backup_error)}")
                            return
                elif uploaded_file.type == DOCX_MIME:
                    try:
                        text = self.analyzer.extract_text_from_docx(uploaded_file)
                    except Exception as docx_error:
                        st.error(f"DOCX extraction failed: {str(docx_error)}")
                        try:
                            text = self.ai_analyzer.extract_text_from_docx(uploaded_file)
                        except Exception as backup_error:
                            st.error(f"All DOCX extraction methods failed: {str(backup_error)}")
                            return
                else:
                    text = uploaded_file.getvalue().decode()

                if not text or text.strip() == "":
                    st.error("Could not extract any text from the uploaded file. Please try a different file.")
                    return
            except Exception as e:
                st.error(f"Error reading file: {str(e)}")
                return

            # Analyze the document
            analysis = self.analyzer.analyze_resume({'raw_text': text}, role_info)

            if 'error' in analysis:
                st.error(analysis['error'])
                return

            st.snow()

            # Save resume data to database
            resume_data = {
                'personal_info': {
                    'name': analysis.get('name', ''),
                    'email': analysis.get('email', ''),
                    'phone': analysis.get('phone', ''),
                    'linkedin': analysis.get('linkedin', ''),
                    'github': analysis.get('github', ''),
                    'portfolio': analysis.get('portfolio', '')
                },
                'summary': analysis.get('summary', ''),
                'target_role': selected_role,
                'target_category': selected_category,
                'education': analysis.get('education', []),
                'experience': analysis.get('experience', []),
                'projects': analysis.get('projects', []),
                'skills': analysis.get('skills', []),
                'template': ''
            }

            try:
                resume_id = save_resume_data(resume_data)

                analysis_data = {
                    'resume_id': resume_id,
                    'ats_score': analysis['ats_score'],
                    'keyword_match_score': analysis['keyword_match']['score'],
                    'format_score': analysis['format_score'],
                    'section_score': analysis['section_score'],
                    'missing_skills': ','.join(analysis['keyword_match']['missing_skills']),
                    'recommendations': ','.join(analysis['suggestions'])
                }
                save_analysis_data(resume_id, analysis_data)
                st.success("Resume data saved successfully!")
            except Exception as e:
                st.error(f"Error saving to database: {str(e)}")
                print(f"Database error: {e}")

            # Show results only for resumes
            if analysis.get('document_type') != 'resume':
                doc_type = analysis.get('document_type')
                st.error(f"⚠️ This appears to be a {doc_type} document, not a resume!")
                st.warning("Please upload a proper resume for ATS analysis.")
                return

        # ---- Results ----
        ats = analysis['ats_score']
        ats_color = score_color(ats)
        ats_status = score_status(ats)
        missing_skills = analysis['keyword_match']['missing_skills']

        col1, col2 = st.columns(2)

        with col1:
            # ATS Score card with circular progress
            st.markdown(f"""
            <div class="feature-card">
                <h2>ATS Score</h2>
                <div style="position: relative; width: 150px; height: 150px; margin: 0 auto;">
                    <div style="position: absolute; width: 150px; height: 150px; border-radius: 50%;
                                background: conic-gradient(#4CAF50 0% {ats}%, #2c2c2c {ats}% 100%);
                                display: flex; align-items: center; justify-content: center;">
                        <div style="width: 120px; height: 120px; background: #1a1a1a; border-radius: 50%;
                                    display: flex; align-items: center; justify-content: center;
                                    font-size: 24px; font-weight: bold; color: {ats_color};">
                            {ats}
                        </div>
                    </div>
                </div>
                <div style="text-align: center; margin-top: 10px;">
                    <span style="font-size: 1.2em; color: {ats_color}; font-weight: bold;">{ats_status}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Skills Match card
            st.markdown('<div class="feature-card"><h2>Skills Match</h2></div>', unsafe_allow_html=True)
            keyword_score = int(analysis.get('keyword_match', {}).get('score', 0))
            st.metric("Keyword Match", f"{keyword_score}%")

            if missing_skills:
                st.markdown("#### Missing Skills:")
                for skill in missing_skills:
                    st.markdown(f"- {skill}")

        with col2:
            # Format Score card
            st.markdown('<div class="feature-card"><h2>Format Analysis</h2></div>', unsafe_allow_html=True)
            st.metric("Format Score", f"{int(analysis.get('format_score', 0))}%")
            st.metric("Section Score", f"{int(analysis.get('section_score', 0))}%")

            # Suggestions
            st.markdown(
                '<div class="feature-card"><h2>📋 Resume Improvement Suggestions</h2></div>',
                unsafe_allow_html=True)

            if analysis.get('contact_suggestions'):
                suggestion_block("📞 Contact Information", li_items(analysis['contact_suggestions']))

            if analysis.get('summary_suggestions'):
                suggestion_block("📝 Professional Summary", li_items(analysis['summary_suggestions']))

            if analysis.get('skills_suggestions') or missing_skills:
                lis = li_items(analysis.get('skills_suggestions', []))
                if missing_skills:
                    lis += "<li style='margin-bottom: 8px;'>✓ Consider adding these relevant skills:</li>"
                    lis += "".join(
                        f"<li style='margin-left: 20px; margin-bottom: 4px;'>• {s}</li>"
                        for s in missing_skills)
                suggestion_block("🎯 Skills", lis)

            if analysis.get('experience_suggestions'):
                suggestion_block("💼 Work Experience", li_items(analysis['experience_suggestions']))

            if analysis.get('education_suggestions'):
                suggestion_block("🎓 Education", li_items(analysis['education_suggestions']))

            if analysis.get('format_suggestions'):
                suggestion_block("📄 Formatting", li_items(analysis['format_suggestions']))

        # Course Recommendations
        st.markdown('<div class="feature-card"><h2>📚 Recommended Courses</h2></div>', unsafe_allow_html=True)

        courses = get_courses_for_role(selected_role)
        if not courses:
            category = get_category_for_role(selected_role)
            courses = COURSES_BY_CATEGORY.get(category, {}).get(selected_role, [])

        cols = st.columns(2)
        for i, course in enumerate(courses[:6]):  # Show top 6 courses
            with cols[i % 2]:
                st.markdown(f"""
                <div style='background-color: #1e1e1e; padding: 15px; border-radius: 10px; margin: 10px 0;'>
                    <h4>{course[0]}</h4>
                    <a href='{course[1]}' target='_blank'>View Course</a>
                </div>
                """, unsafe_allow_html=True)

        # Learning Resources
        st.markdown('<div class="feature-card"><h2>📺 Helpful Videos</h2></div>', unsafe_allow_html=True)

        tab1, tab2 = st.tabs(["Resume Tips", "Interview Tips"])

        with tab1:
            for category, videos in RESUME_VIDEOS.items():
                st.subheader(category)
                vcols = st.columns(2)
                for i, video in enumerate(videos):
                    with vcols[i % 2]:
                        st.video(video[1])

        with tab2:
            for category, videos in INTERVIEW_VIDEOS.items():
                st.subheader(category)
                vcols = st.columns(2)
                for i, video in enumerate(videos):
                    with vcols[i % 2]:
                        st.video(video[1])

    # ------------------------------------------------------------------
    # AI analyzer tab
    # ------------------------------------------------------------------
    def render_ai_stats(self):
        """AI analyzer statistics expander"""
        with st.expander("📊 AI Analyzer Statistics", expanded=False):
            try:
                # Reset button for admin users
                if st.session_state.get('is_admin', False):
                    if st.button("🔄 Reset AI Analysis Statistics", type="secondary",
                                 key="reset_ai_stats_button_2"):
                        result = reset_ai_analysis_stats()
                        if result["success"]:
                            st.success(result["message"])
                        else:
                            st.error(result["message"])
                        st.rerun()  # FIX: st.experimental_rerun() was removed from Streamlit

                ai_stats = get_detailed_ai_analysis_stats()

                if ai_stats["total_analyses"] == 0:
                    st.info("No AI analysis data available yet. Upload and analyze resumes to see statistics here.")
                    return

                st.markdown("""
                <style>
                .stats-card {
                    background: linear-gradient(135deg, #1e3c72, #2a5298); border-radius: 10px;
                    padding: 15px; margin-bottom: 15px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                    text-align: center;
                }
                .stats-value { font-size: 28px; font-weight: bold; color: white; margin: 10px 0; }
                .stats-label {
                    font-size: 14px; color: rgba(255, 255, 255, 0.8);
                    text-transform: uppercase; letter-spacing: 1px;
                }
                </style>
                """, unsafe_allow_html=True)

                avg = ai_stats["average_score"]
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.markdown(f"""
                    <div class="stats-card">
                        <div class="stats-label">Total AI Analyses</div>
                        <div class="stats-value">{ai_stats["total_analyses"]}</div>
                    </div>
                    """, unsafe_allow_html=True)

                avg_color = "#38ef7d" if avg >= 80 else "#FFEB3B" if avg >= 60 else "#FF5252"

                with col2:
                    st.markdown(f"""
                    <div class="stats-card" style="background: linear-gradient(135deg, #2c3e50, {avg_color});">
                        <div class="stats-label">Average Resume Score</div>
                        <div class="stats-value">{avg}/100</div>
                    </div>
                    """, unsafe_allow_html=True)

                with col3:
                    fig = go.Figure(go.Indicator(
                        mode="gauge+number",
                        value=avg,
                        domain={'x': [0, 1], 'y': [0, 1]},
                        title={'text': "Score", 'font': {'size': 14, 'color': 'white'}},
                        gauge={
                            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "white"},
                            'bar': {'color': avg_color},
                            'bgcolor': "rgba(0,0,0,0)",
                            'borderwidth': 2,
                            'bordercolor': "white",
                            'steps': [
                                {'range': [0, 40], 'color': 'rgba(255, 82, 82, 0.3)'},
                                {'range': [40, 70], 'color': 'rgba(255, 235, 59, 0.3)'},
                                {'range': [70, 100], 'color': 'rgba(56, 239, 125, 0.3)'}
                            ],
                        }
                    ))
                    fig.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        font={'color': "white"}, height=150,
                        margin=dict(l=10, r=10, t=30, b=10))
                    st.plotly_chart(fig, width='stretch')

                # Model usage
                if ai_stats["model_usage"]:
                    st.markdown("### 🤖 Model Usage")
                    model_data = pd.DataFrame(ai_stats["model_usage"])
                    fig = px.pie(
                        model_data, values="count", names="model",
                        color_discrete_sequence=px.colors.qualitative.Bold, hole=0.4)
                    fig.update_traces(
                        textposition='inside', textinfo='percent+label',
                        marker=dict(line=dict(color='#000000', width=1.5)))
                    style_dark_chart(fig, 300, 'AI Model Distribution')
                    fig.update_layout(
                        margin=dict(l=20, r=20, t=30, b=20),
                        legend=dict(orientation="h", yanchor="bottom", y=-0.1,
                                    xanchor="center", x=0.5))
                    st.plotly_chart(fig, width='stretch')

                # Top job roles + trend
                if ai_stats["top_job_roles"]:
                    st.markdown("### 🎯 Top Job Roles")
                    roles_data = pd.DataFrame(ai_stats["top_job_roles"])
                    fig = px.bar(
                        roles_data, x="role", y="count", color="count",
                        color_continuous_scale=px.colors.sequential.Viridis,
                        labels={"role": "Job Role", "count": "Number of Analyses"})
                    fig.update_traces(marker_line_width=1.5, marker_line_color="white", opacity=0.9)
                    style_dark_chart(fig, 350, 'Most Analyzed Job Roles')
                    fig.update_layout(
                        xaxis=dict(title="", tickangle=-45, tickfont=dict(size=12)),
                        yaxis=dict(title="Number of Analyses", gridcolor="rgba(255, 255, 255, 0.1)"),
                        coloraxis_showscale=False)
                    st.plotly_chart(fig, width='stretch')

                    # Analysis trend (illustrative data)
                    st.markdown("### 📈 Analysis Trend")
                    st.info("This is a conceptual visualization. To implement actual time-based "
                            "analysis, additional data collection would be needed.")

                    today = datetime.now()
                    dates = [(today - timedelta(days=i)).strftime('%Y-%m-%d') for i in range(7)]
                    dates.reverse()

                    total = ai_stats["total_analyses"]
                    if total > 7:
                        values = np.random.dirichlet(np.ones(7)) * total
                        values = [int(round(v)) for v in values]
                        values[-1] += total - sum(values)
                    else:
                        values = [0] * 7
                        for i in range(total):
                            values[-(i % 7) - 1] += 1

                    trend_data = pd.DataFrame({'Date': dates, 'Analyses': values})
                    fig = px.line(
                        trend_data, x='Date', y='Analyses', markers=True,
                        line_shape='spline', color_discrete_sequence=["#38ef7d"])
                    fig.update_traces(
                        line=dict(width=3),
                        marker=dict(size=8, line=dict(width=2, color='white')))
                    style_dark_chart(fig, 300, 'Analysis Activity (Last 7 Days)')
                    fig.update_layout(
                        xaxis=dict(title="", gridcolor="rgba(255, 255, 255, 0.1)"),
                        yaxis=dict(title="Number of Analyses", gridcolor="rgba(255, 255, 255, 0.1)"))
                    st.plotly_chart(fig, width='stretch')

                # Score distribution
                if ai_stats["score_distribution"]:
                    st.markdown("""
                    <h3 style='text-align: center; margin-bottom: 20px; background: linear-gradient(90deg, #4b6cb7, #182848); padding: 15px; border-radius: 10px; color: white; box-shadow: 0 4px 10px rgba(0,0,0,0.2);'>
                        📊 Score Distribution Analysis
                    </h3>
                    """, unsafe_allow_html=True)

                    score_data = pd.DataFrame(ai_stats["score_distribution"])
                    fig = px.bar(
                        score_data, x="range", y="count", color="range",
                        color_discrete_map={
                            "0-20": "#FF5252", "21-40": "#FF7043", "41-60": "#FFEB3B",
                            "61-80": "#8BC34A", "81-100": "#38ef7d"
                        },
                        labels={"range": "Score Range", "count": "Number of Resumes"},
                        text="count")
                    fig.update_traces(
                        marker_line_width=2, marker_line_color="white", opacity=0.9,
                        textposition='outside',
                        textfont=dict(color="white", size=14, family="Arial, sans-serif"),
                        hovertemplate="<b>Score Range:</b> %{x}<br><b>Number of Resumes:</b> %{y}<extra></extra>")
                    style_dark_chart(fig, 400)
                    fig.update_layout(
                        xaxis=dict(
                            title=dict(text="Score Range", font=dict(size=16, color="white")),
                            categoryorder="array",
                            categoryarray=["0-20", "21-40", "41-60", "61-80", "81-100"],
                            tickfont=dict(size=14, color="white"),
                            gridcolor="rgba(255, 255, 255, 0.1)"),
                        yaxis=dict(
                            title=dict(text="Number of Resumes", font=dict(size=16, color="white")),
                            tickfont=dict(size=14, color="white"),
                            gridcolor="rgba(255, 255, 255, 0.1)", zeroline=False),
                        showlegend=False,
                        bargap=0.2)

                    if len(score_data) > 0:
                        max_range = score_data.loc[score_data["count"].idxmax(), "range"]
                        fig.add_annotation(
                            x=0.5, y=1.12, xref="paper", yref="paper",
                            text=f"Most resumes fall in the {max_range} score range",
                            showarrow=False, font=dict(size=14, color="#FFEB3B"),
                            bgcolor="rgba(0,0,0,0.5)", bordercolor="#FFEB3B",
                            borderwidth=1, borderpad=4, opacity=0.8)

                    st.plotly_chart(fig, width='stretch')
                    st.caption("This chart shows the distribution of resume scores across "
                               "different ranges, helping identify common performance levels.")

                # Recent analyses
                if ai_stats["recent_analyses"]:
                    st.markdown("""
                    <h3 style='text-align: center; margin-bottom: 20px; background: linear-gradient(90deg, #4b6cb7, #182848); padding: 15px; border-radius: 10px; color: white; box-shadow: 0 4px 10px rgba(0,0,0,0.2);'>
                        🕒 Recent Resume Analyses
                    </h3>
                    """, unsafe_allow_html=True)

                    rows = ""
                    for item in ai_stats["recent_analyses"]:
                        score = item["score"]
                        score_class = "score-high" if score >= 80 else "score-medium" if score >= 60 else "score-low"
                        model_name = item["model"]
                        model_class = ("model-gemini" if "Gemini" in model_name
                                       else "model-claude" if "Claude" in model_name else "")
                        try:
                            date_obj = datetime.strptime(item["date"], "%Y-%m-%d %H:%M:%S")
                            formatted_date = date_obj.strftime("%b %d, %Y")
                        except Exception:
                            formatted_date = item["date"]

                        rows += (
                            "<tr>"
                            f'<td><div class="model-badge {model_class}">{model_name}</div></td>'
                            f'<td><div class="score-pill {score_class}">{score}/100</div></td>'
                            f'<td><div class="role-badge">{item["job_role"]}</div></td>'
                            f'<td><div class="date-badge">{formatted_date}</div></td>'
                            "</tr>"
                        )

                    st.markdown("""
                    <style>
                    .modern-analyses-table { width: 100%; border-collapse: separate; border-spacing: 0 8px; font-family: Arial, sans-serif; }
                    .modern-analyses-table th {
                        background: linear-gradient(135deg, #1e3c72, #2a5298); color: white; padding: 15px;
                        text-align: left; font-weight: bold; font-size: 14px; text-transform: uppercase;
                        letter-spacing: 1px;
                    }
                    .modern-analyses-table td {
                        padding: 15px; background-color: rgba(30, 30, 30, 0.7); color: white;
                        border-top: 1px solid rgba(255, 255, 255, 0.05);
                    }
                    .model-badge, .score-pill, .date-badge, .role-badge {
                        display: inline-block; padding: 6px 12px; border-radius: 20px;
                        font-weight: bold; text-align: center; font-size: 12px;
                    }
                    .model-gemini { background: linear-gradient(135deg, #4e54c8, #8f94fb); color: white; }
                    .model-claude { background: linear-gradient(135deg, #834d9b, #d04ed6); color: white; }
                    .score-high { background: linear-gradient(135deg, #11998e, #38ef7d); color: white; }
                    .score-medium { background: linear-gradient(135deg, #f2994a, #f2c94c); color: white; }
                    .score-low { background: linear-gradient(135deg, #cb2d3e, #ef473a); color: white; }
                    .date-badge { background-color: rgba(255, 255, 255, 0.1); color: #e0e0e0; font-weight: normal; }
                    .role-badge {
                        background-color: rgba(33, 150, 243, 0.2); color: #90caf9; border-radius: 8px;
                        font-weight: normal; max-width: 200px; white-space: nowrap;
                        overflow: hidden; text-overflow: ellipsis;
                    }
                    </style>
                    """, unsafe_allow_html=True)

                    st.markdown(
                        "<table class='modern-analyses-table'>"
                        "<tr><th>AI Model</th><th>Score</th><th>Job Role</th><th>Date</th></tr>"
                        f"{rows}</table>",
                        unsafe_allow_html=True)
                    st.caption("These are the most recent resume analyses performed by our AI models.")

            except Exception as e:
                st.error(f"Error loading AI analysis statistics: {str(e)}")

    def render_ai_analyzer(self):
        st.markdown("""
        <div style='background-color: #1e1e1e; padding: 20px; border-radius: 10px; margin: 10px 0;'>
            <h3>AI-Powered Resume Analysis</h3>
            <p>Get detailed insights from advanced AI models that analyze your resume and provide personalized recommendations.</p>
            <p><strong>Upload your resume to get AI-powered analysis and recommendations.</strong></p>
        </div>
        """, unsafe_allow_html=True)

        ai_model = st.selectbox(
            "Select AI Model", ["Google Gemini"],
            help="Choose the AI model to analyze your resume")

        use_custom_job_desc = st.checkbox(
            "Use custom job description", value=False,
            help="Enable this to provide a specific job description for more targeted analysis")

        custom_job_description = ""
        if use_custom_job_desc:
            custom_job_description = st.text_area(
                "Paste the job description here",
                height=200,
                placeholder="Paste the full job description from the company here for more targeted analysis...",
                help="Providing the actual job description will help the AI analyze your resume specifically for this position")

            st.markdown("""
            <div style='background-color: #2e7d32; padding: 15px; border-radius: 10px; margin: 10px 0;'>
                <p><i class="fas fa-lightbulb"></i> <strong>Pro Tip:</strong> Including the actual job description significantly improves the accuracy of the analysis and provides more relevant recommendations tailored to the specific position.</p>
            </div>
            """, unsafe_allow_html=True)

        self.render_ai_stats()

        selected_category, selected_role, role_info = self.render_role_selector("ai")

        uploaded_file = st.file_uploader(
            "Upload your resume", type=['pdf', 'docx'], key="ai_file")

        if not uploaded_file:
            st.markdown(
                self.render_empty_state(
                    "fas fa-robot",
                    "Upload your resume to get AI-powered analysis and recommendations"
                ),
                unsafe_allow_html=True
            )
            return

        analyze_ai = st.button(
            "🤖 Analyze with AI", type="primary", width='stretch',
            key="analyze_ai_button")

        if not analyze_ai:
            return

        with st.spinner(f"Analyzing your resume with {ai_model}..."):
            try:
                progress_bar = st.progress(0)
                selected_model = "Google Gemini"
                progress_bar.progress(10)

                # Extract text from the resume (single extraction)
                analyzer = AIResumeAnalyzer()
                try:
                    resume_text = self.extract_resume_text(uploaded_file, analyzer)
                except Exception as e:
                    st.error(f"Error reading file: {str(e)}")
                    return
                progress_bar.progress(30)

                job_role = selected_role if selected_role else "Not specified"
                progress_bar.progress(50)

                # Analyze the resume with Google Gemini
                if use_custom_job_desc and custom_job_description:
                    analysis_result = analyzer.analyze_resume_with_gemini(
                        resume_text, job_role=job_role, job_description=custom_job_description)
                    st.session_state['used_custom_job_desc'] = True
                else:
                    analysis_result = analyzer.analyze_resume_with_gemini(
                        resume_text, job_role=job_role)
                    st.session_state['used_custom_job_desc'] = False

                progress_bar.progress(80)

                # Save the analysis to the database
                if analysis_result and "error" not in analysis_result:
                    save_ai_analysis_data(
                        None,  # No user_id needed
                        {
                            "model_used": selected_model,
                            "resume_score": analysis_result.get("resume_score", 0),
                            "job_role": job_role
                        }
                    )

                st.snow()
                progress_bar.progress(100)

                if not analysis_result or "error" in analysis_result:
                    error_msg = analysis_result.get('error', 'Unknown error') if analysis_result else 'Unknown error'
                    st.error(f"Analysis failed: {error_msg}")
                    return

                st.success("✅ Analysis complete!")

                full_response = analysis_result.get("analysis", "")
                resume_score = analysis_result.get("resume_score", 0)
                ats_score = analysis_result.get("ats_score", 0)
                model_used = analysis_result.get("model_used", selected_model)
                used_custom = st.session_state.get('used_custom_job_desc', False)

                # Stored for download
                st.session_state['full_analysis'] = full_response

                st.markdown("## Full Analysis Report")

                current_date = datetime.now().strftime("%B %d, %Y")
                overall_status = score_status(resume_score)
                custom_badge = (
                    '<p style="color: #4CAF50;"><strong>✓ Custom Job Description Used</strong></p>'
                    if used_custom else ''
                )

                # Report header (FIX: all <div> tags are now closed)
                st.markdown(f"""
                <div style="background-color: #262730; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
                    <h2 style="color: #ffffff; margin-bottom: 10px;">AI Resume Analysis Report</h2>
                    <div style="display: flex; flex-wrap: wrap; gap: 20px;">
                        <div style="flex: 1; min-width: 200px;">
                            <p style="color: #ffffff;"><strong>Job Role:</strong> {job_role}</p>
                            <p style="color: #ffffff;"><strong>Analysis Date:</strong> {current_date}</p>
                        </div>
                        <div style="flex: 1; min-width: 200px;">
                            <p style="color: #ffffff;"><strong>AI Model:</strong> {model_used}</p>
                            <p style="color: #ffffff;"><strong>Overall Score:</strong> {resume_score}/100 - {overall_status}</p>
                            {custom_badge}
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                # Gauge charts for scores
                col1, col2 = st.columns(2)

                with col1:
                    st.plotly_chart(gauge_figure(resume_score, "Resume Score"), width='stretch')
                    st.markdown(
                        f"<div style='text-align: center; font-weight: bold;'>{score_status(resume_score)}</div>",
                        unsafe_allow_html=True)

                with col2:
                    st.plotly_chart(gauge_figure(ats_score, "ATS Optimization Score"), width='stretch')
                    st.markdown(
                        f"<div style='text-align: center; font-weight: bold;'>{score_status(ats_score)}</div>",
                        unsafe_allow_html=True)

                # Job description match score (if a custom job description was used)
                if used_custom and custom_job_description:
                    job_match_score = analysis_result.get("job_match_score", 0)
                    if not job_match_score and "job_match" in analysis_result:
                        job_match_score = analysis_result["job_match"].get("score", 0)

                    if job_match_score:
                        st.markdown("""
                        <h3 style="background: linear-gradient(90deg, #4d7c0f, #84cc16); color: white; padding: 10px; border-radius: 5px; margin-top: 20px;">
                            <i class="fas fa-handshake"></i> Job Description Match Analysis
                        </h3>
                        """, unsafe_allow_html=True)

                        mcol1, mcol2 = st.columns(2)

                        with mcol1:
                            st.plotly_chart(gauge_figure(job_match_score, "Job Match Score"), width='stretch')
                            match_status = ("Excellent Match" if job_match_score >= 80
                                            else "Good Match" if job_match_score >= 60 else "Low Match")
                            st.markdown(
                                f"<div style='text-align: center; font-weight: bold;'>{match_status}</div>",
                                unsafe_allow_html=True)

                        with mcol2:
                            st.markdown("""
                            <div style="background-color: #262730; padding: 20px; border-radius: 10px; height: 100%;">
                                <h4 style="color: #ffffff; margin-bottom: 15px;">What This Means</h4>
                                <p style="color: #ffffff;">This score represents how well your resume matches the specific job description you provided.</p>
                                <ul style="color: #ffffff; padding-left: 20px;">
                                    <li><strong>80-100:</strong> Excellent match - your resume is highly aligned with this job</li>
                                    <li><strong>60-79:</strong> Good match - your resume matches many requirements</li>
                                    <li><strong>Below 60:</strong> Consider tailoring your resume more specifically to this job</li>
                                </ul>
                            </div>
                            """, unsafe_allow_html=True)

                # Styled full report
                formatted_analysis = format_analysis_report(full_response)

                st.markdown("""
                <style>
                    .report-section { margin-bottom: 25px; border: 1px solid #4B4B4B; border-radius: 8px; overflow: hidden; }
                    .section-content { padding: 15px; background-color: #262730; color: #ffffff; }
                    .report-section h3 { margin-top: 0; font-weight: 600; }
                    .report-section ul { padding-left: 20px; }
                    .report-section p { color: #ffffff; margin-bottom: 10px; }
                    .report-section li { color: #ffffff; margin-bottom: 5px; }
                </style>
                """, unsafe_allow_html=True)

                st.markdown(f"""
                <div style="background-color: #262730; padding: 20px; border-radius: 10px; border: 1px solid #4B4B4B; color: #ffffff;">

{formatted_analysis}

                </div>
                """, unsafe_allow_html=True)

                # Create a PDF report
                pdf_buffer = self.ai_analyzer.generate_pdf_report(
                    analysis_result={
                        "score": resume_score,
                        "ats_score": ats_score,
                        "model_used": model_used,
                        "full_response": full_response,
                        "strengths": analysis_result.get("strengths", []),
                        "weaknesses": analysis_result.get("weaknesses", []),
                        "used_custom_job_desc": used_custom,
                        "custom_job_description": custom_job_description if used_custom else ""
                    },
                    candidate_name=st.session_state.get('candidate_name', 'Candidate'),
                    job_role=selected_role
                )

                if pdf_buffer:
                    st.download_button(
                        label="📊 Download PDF Report",
                        data=pdf_buffer,
                        file_name=f"resume_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        width='stretch',
                        on_click=lambda: st.balloons()
                    )
                else:
                    st.error("PDF generation failed. Please try again later.")

            except Exception as ai_error:
                st.error(f"Error during AI analysis: {str(ai_error)}")
                st.code(traceback.format_exc())

    # ------------------------------------------------------------------
    # Other pages
    # ------------------------------------------------------------------
    def render_home(self):
        apply_modern_styles()

        hero_section(
            "Smart Resume AI",
            "Transform your career with AI-powered resume analysis and building. Get personalized insights and create professional resumes that stand out."
        )

        st.markdown('<div class="feature-grid">', unsafe_allow_html=True)

        feature_card(
            "fas fa-robot",
            "AI-Powered Analysis",
            "Get instant feedback on your resume with advanced AI analysis that identifies strengths and areas for improvement."
        )

        feature_card(
            "fas fa-magic",
            "Smart Resume Builder",
            "Create professional resumes with our intelligent builder that suggests optimal content and formatting."
        )

        feature_card(
            "fas fa-chart-line",
            "Career Insights",
            "Access detailed analytics and personalized recommendations to enhance your career prospects."
        )

        st.markdown('</div>', unsafe_allow_html=True)

        st.toast("Check out these repositories: [AI-Nexus(AI/ML)](https://github.com/Hunterdii/AI-Nexus)", icon="ℹ️")

        # Call-to-Action with Streamlit navigation
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            if st.button("Get Started", key="get_started_btn",
                         help="Click to start analyzing your resume",
                         type="primary",
                         width='stretch'):
                cleaned_name = "🔍 RESUME ANALYZER".lower().replace(" ", "_").replace("🔍", "").strip()
                st.session_state.page = cleaned_name
                st.rerun()

    def render_job_search(self):
        """Render the job search page"""
        render_job_search()

        st.toast("Check out these repositories: [GeeksforGeeks-POTD](https://github.com/Hunterdii/GeeksforGeeks-POTD)", icon="ℹ️")

    def render_feedback_page(self):
        """Render the feedback page"""
        apply_modern_styles()

        page_header(
            "Feedback & Suggestions",
            "Help us improve by sharing your thoughts"
        )

        feedback_manager = FeedbackManager()

        form_tab, stats_tab = st.tabs(["Submit Feedback", "Feedback Stats"])

        with form_tab:
            feedback_manager.render_feedback_form()

        with stats_tab:
            feedback_manager.render_feedback_stats()

        st.toast("Check out these repositories: [TryHackMe Free Rooms](https://github.com/Hunterdii/tryhackme-free-rooms)", icon="ℹ️")

    def show_repo_notification(self):
        message = """
<div style="background-color: #1e1e1e; border-radius: 10px; border: 1px solid #4b6cb7; padding: 10px; margin: 10px 0; color: white;">
    <div style="margin-bottom: 10px;">Check out these other repositories:</div>
    <div style="margin-bottom: 5px;"><b>Hacking Resources:</b></div>
    <ul style="margin-top: 0; padding-left: 20px;">
        <li><a href="https://github.com/Hunterdii/tryhackme-free-rooms" target="_blank" style="color: #4CAF50;">TryHackMe Free Rooms</a></li>
        <li><a href="https://github.com/Hunterdii/Awesome-Hacking" target="_blank" style="color: #4CAF50;">Awesome Hacking</a></li>
    </ul>
    <div style="margin-bottom: 5px;"><b>Programming Languages:</b></div>
    <ul style="margin-top: 0; padding-left: 20px;">
        <li><a href="https://github.com/Hunterdii/Awesome-Java" target="_blank" style="color: #4CAF50;">Awesome Java</a></li>
        <li><a href="https://github.com/Hunterdii/30-Days-Of-Rust" target="_blank" style="color: #4CAF50;">30 Days Of Rust</a></li>
    </ul>
    <div style="margin-bottom: 5px;"><b>Data Structures & Algorithms:</b></div>
    <ul style="margin-top: 0; padding-left: 20px;">
        <li><a href="https://github.com/Hunterdii/GeeksforGeeks-POTD" target="_blank" style="color: #4CAF50;">GeeksforGeeks POTD</a></li>
        <li><a href="https://github.com/Hunterdii/Leetcode-POTD" target="_blank" style="color: #4CAF50;">Leetcode POTD</a></li>
    </ul>
    <div style="margin-bottom: 5px;"><b>AI/ML Projects:</b></div>
    <ul style="margin-top: 0; padding-left: 20px;">
        <li><a href="https://github.com/Hunterdii/AI-Nexus" target="_blank" style="color: #4CAF50;">AI Nexus</a></li>
    </ul>
    <div style="margin-top: 10px;">If you find this project helpful, please consider ⭐ starring the repo!</div>
</div>
"""
        st.sidebar.markdown(message, unsafe_allow_html=True)

    def main(self):
        """Main application entry point"""
        self.apply_global_styles()

        # Admin login/logout in sidebar
        with st.sidebar:
            lottie_data = self.load_lottie_url(
                "https://assets5.lottiefiles.com/packages/lf20_xyadoh9h.json")
            if lottie_data:  # FIX: don't crash when offline
                st_lottie(lottie_data, height=200, key="sidebar_animation")
            st.title("Smart Resume AI")
            st.markdown("---")

            # Navigation buttons
            for page_name in self.pages.keys():
                if st.button(page_name, width='stretch'):
                    cleaned_name = (page_name.lower().replace(" ", "_").replace("🏠", "")
                                    .replace("🔍", "").replace("📝", "").replace("📊", "")
                                    .replace("🎯", "").replace("💬", "").replace("ℹ️", "").strip())
                    st.session_state.page = cleaned_name
                    st.rerun()

            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown("---")

            # Admin Login/Logout section at bottom
            if st.session_state.get('is_admin', False):
                st.success(f"Logged in as: {st.session_state.get('current_admin_email')}")
                if st.button("Logout", key="logout_button"):
                    try:
                        log_admin_action(st.session_state.get('current_admin_email'), "logout")
                        st.session_state.is_admin = False
                        st.session_state.current_admin_email = None
                        st.success("Logged out successfully!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error during logout: {str(e)}")
            else:
                with st.expander("👤 Admin Login"):
                    admin_email_input = st.text_input("Email", key="admin_email_input")
                    admin_password = st.text_input("Password", type="password", key="admin_password_input")
                    if st.button("Login", key="login_button"):
                        try:
                            if verify_admin(admin_email_input, admin_password):
                                st.session_state.is_admin = True
                                st.session_state.current_admin_email = admin_email_input
                                log_admin_action(admin_email_input, "login")
                                st.success("Logged in successfully!")
                                st.rerun()
                            else:
                                st.error("Invalid credentials")
                        except Exception as e:
                            st.error(f"Error during login: {str(e)}")

            self.show_repo_notification()

        # Force home page on first load
        if 'initial_load' not in st.session_state:
            st.session_state.initial_load = True
            st.session_state.page = 'home'
            st.rerun()

        # Get current page and render it
        current_page = st.session_state.get('page', 'home')

        # Mapping of cleaned page names to original names
        page_mapping = {
            (name.lower().replace(" ", "_").replace("🏠", "").replace("🔍", "")
             .replace("📝", "").replace("📊", "").replace("🎯", "")
             .replace("💬", "").replace("ℹ️", "").strip()): name
            for name in self.pages.keys()
        }

        if current_page in page_mapping:
            self.pages[page_mapping[current_page]]()
        else:
            # Default to home page if invalid page
            self.render_home()


if __name__ == "__main__":
    app = ResumeApp()
    app.main()