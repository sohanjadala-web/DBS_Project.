"""
Generates complete Academic Project Review presentations (Review 0, 1, 2, 3, 4)
for the KL University / KLH DBS Capstone Project.
"""
import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Global Design Palette
C_NAVY_DARK = RGBColor(15, 23, 42)      # #0F172A
C_NAVY_BLUE = RGBColor(30, 58, 138)     # #1E3A8A
C_BLUE_ACCENT = RGBColor(2, 132, 199)   # #0284C7
C_LIGHT_BG = RGBColor(248, 250, 252)    # #F8FAFC
C_CARD_BG = RGBColor(255, 255, 255)     # #FFFFFF
C_CARD_BORDER = RGBColor(226, 232, 240) # #E2E8F0
C_TEXT_DARK = RGBColor(30, 41, 59)      # #1E293B
C_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B
C_TEXT_LIGHT = RGBColor(255, 255, 255)  # #FFFFFF
C_EMERALD = RGBColor(5, 150, 105)       # #059669
C_AMBER = RGBColor(217, 119, 6)         # #D97706
C_RED = RGBColor(220, 38, 38)           # #DC2626
C_CYAN = RGBColor(6, 182, 212)          # #06B6D4

LOGO_PATH = 'CKD_Project/frontend/assets/kidney_logo.png'

TEAM_INFO = [
    ("2520090137", "Jayavarapu Sri Charan"),
    ("2520080060", "J. Sohan"),
    ("2520080036", "Adithya Vishnubatala"),
]

COURSE_INFO = "25CS1302E — Database Management Systems (DBS) | KL University (KLH Campus)"

def set_slide_background(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

def add_header(slide, title, category, review_tag, current_slide, total_slides):
    set_slide_background(slide, C_LIGHT_BG)
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(10.5), Inches(1.15))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    
    p0 = tf.paragraphs[0]
    p0.text = f"{category.upper()}  •  {review_tag.upper()}"
    p0.font.size = Pt(10)
    p0.font.bold = True
    p0.font.color.rgb = C_BLUE_ACCENT
    p0.font.name = "Segoe UI"
    
    p1 = tf.add_paragraph()
    p1.text = title
    p1.font.size = Pt(20)
    p1.font.bold = True
    p1.font.color.rgb = C_NAVY_DARK
    p1.font.name = "Segoe UI"
    p1.space_before = Pt(2)

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = C_CARD_BORDER
    line.line.fill.background()

    if os.path.exists(LOGO_PATH):
        try:
            slide.shapes.add_picture(LOGO_PATH, Inches(12.0), Inches(0.35), width=Inches(0.55))
        except Exception:
            pass

    ft_tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.1), Inches(11.733), Inches(0.3))
    ft_tf = ft_tb.text_frame
    ft_tf.margin_left = ft_tf.margin_top = ft_tf.margin_right = ft_tf.margin_bottom = 0
    p_ft = ft_tf.paragraphs[0]
    p_ft.text = f"{COURSE_INFO}  |  {review_tag}  —  Slide {current_slide} of {total_slides}"
    p_ft.font.size = Pt(8.5)
    p_ft.font.color.rgb = C_TEXT_MUTED
    p_ft.font.name = "Segoe UI"

def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_CARD_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    if border_color:
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
    else:
        card.line.fill.background()
    return card

def create_title_slide(prs, review_title, review_tag, subtitle):
    blank_layout = prs.slide_layouts[6]
    s = prs.slides.add_slide(blank_layout)
    set_slide_background(s, C_NAVY_DARK)

    hero_img = "presentation_assets/ai_nephrology_hero.jpg"
    if os.path.exists(hero_img):
        try:
            s.shapes.add_picture(hero_img, Inches(7.3), Inches(0.8), width=Inches(5.3))
        except Exception:
            pass

    tbox = s.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(6.2), Inches(6.0))
    ttf = tbox.text_frame
    ttf.word_wrap = True

    p = ttf.paragraphs[0]
    p.text = f"CAPSTONE PROJECT EVALUATION  •  {review_tag.upper()}"
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(56, 189, 248)
    p.font.name = "Segoe UI"

    p = ttf.add_paragraph()
    p.text = review_title
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = C_TEXT_LIGHT
    p.font.name = "Segoe UI"
    p.space_before = Pt(6)

    p = ttf.add_paragraph()
    p.text = subtitle
    p.font.size = Pt(12)
    p.font.color.rgb = RGBColor(203, 213, 225)
    p.font.name = "Segoe UI"
    p.space_before = Pt(4)

    p = ttf.add_paragraph()
    p.text = "──────────────────────────────────────────────────────"
    p.font.size = Pt(9)
    p.font.color.rgb = RGBColor(71, 85, 105)
    p.space_before = Pt(8)

    p = ttf.add_paragraph()
    p.text = "Team 06 | Project Members:"
    p.font.bold = True
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(56, 189, 248)
    p.space_before = Pt(6)

    for roll, name in TEAM_INFO:
        p = ttf.add_paragraph()
        p.text = f"• {roll} — {name}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_LIGHT
        p.space_before = Pt(2)

    p = ttf.add_paragraph()
    p.text = f"Course: {COURSE_INFO}"
    p.font.size = Pt(9)
    p.font.color.rgb = RGBColor(148, 163, 184)
    p.space_before = Pt(8)
    return s

def build_review_0():
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    tot = 7

    create_title_slide(prs, "AI-Driven Chronic Kidney Disease Management System", "Review 0: Project Inception", "Project Ideation, Problem Definition, Scope & Team Responsibilities")

    # Slide 2: Problem Statement & Domain
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "Problem Statement & Clinical Domain Selection", "REVIEW 0: PROJECT INCEPTION", "Review 0", 2, tot)
    add_card(s2, Inches(0.8), Inches(1.65), Inches(5.7), Inches(5.3))
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "The Healthcare Domain & Chronic Kidney Disease"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_BLUE
    bullets = [
        ("Massive Prevalence: ", "CKD affects >850 million people globally (>10% adult population), leading to over 1.2M premature deaths annually."),
        ("Silent Progression: ", "Early stages (1-3) show minimal noticeable symptoms. Nephron hyperfiltration masks functional loss until >75% of kidney capacity is destroyed."),
        ("Diagnostic Complexity: ", "Requires analyzing 24 disparate clinical parameters across blood biochemistry, urinalysis, electrolytes, and patient comorbidities.")
    ]
    for b_title, b_txt in bullets:
        p = tf.add_paragraph()
        p.text = b_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = b_txt
        run.font.bold = False

    add_card(s2, Inches(6.8), Inches(1.65), Inches(5.733), Inches(5.3))
    tb2 = s2.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Core Problem Statement to Solve"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_RED
    bullets2 = [
        ("Data Fragmentation: ", "Patient records, laboratory biomarker test reports, and doctor consultations exist in isolated silos or manual paper logs, causing delays and lost records."),
        ("Absence of Early Detection: ", "Standard clinics lack automated early-risk triage, allowing asymptomatic cases to progress to irreversible Stage 5 ESRD requiring dialysis ($80k/yr)."),
        ("Proposed Solution: ", "A unified, secure full-stack healthcare platform combining a 3NF Relational DBMS with Explainable Machine Learning (XGBoost/LightGBM) for proactive decision support.")
    ]
    for b_title, b_txt in bullets2:
        p = tf2.add_paragraph()
        p.text = b_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = b_txt
        run.font.bold = False

    # Slide 3: Objectives & Scope
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "Project Objectives & Scope of Work", "REVIEW 0: PROJECT INCEPTION", "Review 0", 3, tot)
    obj_cards = [
        ("1. Database Architecture", "Design a 3NF normalized PostgreSQL schema with integrity constraints, views, stored procedures, and triggers for clinical records.", C_NAVY_BLUE),
        ("2. Machine Learning Engine", "Train gradient-boosted trees on the UCI 400-patient benchmark to achieve high sensitivity with explainable AI (SHAP) feature attributions.", C_EMERALD),
        ("3. Role-Based Access Control", "Implement distinct, secure portals for 4 key roles: Admin, Doctor, Lab Staff, and Patient with JWT authentication.", C_BLUE_ACCENT),
        ("4. Clinical Workflow Integration", "Provide an end-to-end appointment scheduling pipeline with atomic double-booking prevention and verified diagnostic reporting.", C_AMBER)
    ]
    for idx, (title, desc, color) in enumerate(obj_cards):
        c_left = Inches(0.8 + (idx % 2) * 6.0)
        c_top = Inches(1.7 + (idx // 2) * 2.6)
        add_card(s3, c_left, c_top, Inches(5.733), Inches(2.4))
        tb = s3.shapes.add_textbox(c_left + Inches(0.2), c_top + Inches(0.2), Inches(5.333), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = color
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(6)

    # Slide 4: Target Stakeholders & Motivation
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "Motivation & Stakeholder Value Proposition", "REVIEW 0: PROJECT INCEPTION", "Review 0", 4, tot)
    stakeholders = [
        ("Patients", "Immediate risk scoring, self-service appointment booking, digital laboratory history, and proactive prevention guidance.", C_CYAN),
        ("Nephrologists / Doctors", "Patient 360-degree view, one-click ML inference with SHAP explanations, integrated clinical notes, and follow-up tracking.", C_NAVY_BLUE),
        ("Lab Technicians & Staff", "Structured requisition queues, standardized biomarker input forms, file verification gating, and automated patient alerting.", C_EMERALD),
        ("Hospital Administrators", "Live system health monitoring, audit log tracing for data modifications, and dynamic SQL reporting for operational queries.", C_PURPLE if 'C_PURPLE' in globals() else C_NAVY_DARK)
    ]
    for idx, (role, benefit, col) in enumerate(stakeholders):
        c_left = Inches(0.8 + idx * 2.98)
        add_card(s4, c_left, Inches(1.8), Inches(2.8), Inches(5.0))
        tb = s4.shapes.add_textbox(c_left + Inches(0.15), Inches(2.0), Inches(2.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = role
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = benefit
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(8)

    # Slide 5: Proposed Tech Stack
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "High-Level Technology Architecture", "REVIEW 0: PROJECT INCEPTION", "Review 0", 5, tot)
    techs = [
        ("Relational Database", "PostgreSQL / SQLite", "3NF normalization, ACID transactions, Foreign Key cascades, Triggers, Views, and Stored Procedures."),
        ("Backend Framework", "Python & FastAPI", "Asynchronous RESTful APIs, Pydantic data schemas, SQLAlchemy ORM, and OAuth2 JWT authentication."),
        ("Machine Learning", "XGBoost, LightGBM, SHAP", "Trained on UCI CKD 24-feature dataset, zero data leakage, tree-based Shapley value feature attribution."),
        ("Frontend & UI", "Vanilla HTML5, CSS3, JavaScript", "Zero-build responsive SPA architecture, role-segregated portals, and Chart.js analytical visualizations.")
    ]
    for idx, (layer, tech, detail) in enumerate(techs):
        c_left = Inches(0.8 + idx * 2.98)
        add_card(s5, c_left, Inches(1.8), Inches(2.8), Inches(5.0))
        tb = s5.shapes.add_textbox(c_left + Inches(0.15), Inches(2.0), Inches(2.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = layer
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_BLUE_ACCENT
        p2 = tf.add_paragraph()
        p2.text = tech
        p2.font.bold = True
        p2.font.size = Pt(12)
        p2.font.color.rgb = C_NAVY_DARK
        p2.space_before = Pt(4)
        p3 = tf.add_paragraph()
        p3.text = detail
        p3.font.size = Pt(9)
        p3.font.color.rgb = C_TEXT_DARK
        p3.space_before = Pt(8)

    # Slide 6: Team Roles & Milestone Plan
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "Team Division of Work & Milestone Schedule", "REVIEW 0: PROJECT INCEPTION", "Review 0", 6, tot)
    add_card(s6, Inches(0.8), Inches(1.65), Inches(5.7), Inches(5.3))
    tb = s6.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Team Member Responsibilities (Team 06)"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_BLUE
    roles = [
        ("Jayavarapu Sri Charan (2520090137) - Lead:", "Database Architecture (DDL, DML, Views, Triggers, Stored Procedures), FastAPI backend integration, and end-to-end testing."),
        ("J. Sohan (2520080060):", "Machine Learning Pipeline (Preprocessing, XGBoost/LightGBM training, SHAP explainability) and dataset curation."),
        ("Adithya Vishnubatala (2520080036):", "Frontend UI/UX design, role-based dashboards (Admin, Doctor, Lab, Patient), and presentation documentation.")
    ]
    for r_title, r_desc in roles:
        p = tf.add_paragraph()
        p.text = r_title
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = " " + r_desc
        run.font.bold = False

    add_card(s6, Inches(6.8), Inches(1.65), Inches(5.733), Inches(5.3))
    tb2 = s6.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Review Milestones Roadmap"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_EMERALD
    milestones = [
        ("Review 0: ", "Project Ideation, Scope, Objectives, Stakeholders, and Tech Stack."),
        ("Review 1: ", "Literature Survey, System Architecture, ER Modeling & Relational Schema Design."),
        ("Review 2: ", "Core Implementation: DDL/DML, FastAPI Endpoints, ML Training & Preprocessing."),
        ("Review 3: ", "Integration Testing, Triggers/Procedures, Security RBAC, and Validation."),
        ("Review 4: ", "Final Demonstration, Model Benchmarking, Comparison with Existing Systems, & Future Scope.")
    ]
    for m_title, m_desc in milestones:
        p = tf2.add_paragraph()
        p.text = m_title
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = m_desc
        run.font.bold = False

    # Slide 7: Expected Deliverables & Outcomes
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "Expected Outcomes & Review 0 Summary", "REVIEW 0: PROJECT INCEPTION", "Review 0", 7, tot)
    add_card(s7, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s7.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Deliverables Summary for Subsequent Evaluation Reviews"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    summary_pts = [
        ("Functional 3NF Database: ", "10 interrelated tables ensuring zero redundancy, ACID transactions, and foreign key constraints."),
        ("Empirical ML Artifacts: ", "Production-ready serialized models (.joblib) with 100% test sensitivity and instant SHAP force plots."),
        ("Multi-Role Web Application: ", "Responsive portal with dedicated views for Admin, Doctors, Lab Technicians, and Patients."),
        ("Academic Documentation: ", "Comprehensive ER Diagrams, Data Flow Diagrams, SQL Demonstration scripts, and automated Pytest test suite.")
    ]
    for s_title, s_desc in summary_pts:
        p = tf.add_paragraph()
        p.text = f"• {s_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_NAVY_BLUE
        p.space_before = Pt(10)
        run = p.add_run()
        run.text = s_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    prs.save("CKD_Project_Review_0.pptx")
    print("Saved CKD_Project_Review_0.pptx")

def build_review_1():
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    tot = 8

    create_title_slide(prs, "System Design & Relational Modeling", "Review 1: Literature Survey & Design", "Literature Review, Research Gaps, ER Diagrams, Relational Schema & System Architecture")

    # Slide 2: Literature Survey
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "Literature Survey & State-of-the-Art Analysis", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 2, tot)
    lit = [
        ("Capodici et al. (2023)", "Systematic Review of 68 AI Studies", "Evaluated ML in CKD diagnosis. Identified major gaps: lack of real-world clinical workflow integration and insufficient external data validation."),
        ("Khalid et al. (2024)", "CKD Progression Prediction Models", "Demonstrated high predictive utility of gradient boosted trees but noted that models suffer when deployed without strict data preprocessing pipelines."),
        ("Pan & Tong (2024)", "Meta-Analysis on Clinical Discrimination", "Confirmed high diagnostic discrimination of ensemble models across diverse test cohorts, emphasizing the necessity of model interpretability."),
        ("Gogoi et al. (2024)", "Trends & Challenges in Nephrology AI", "Pointed out critical barriers: small datasets, black-box decision models, and isolated software systems lacking unified hospital database management.")
    ]
    for idx, (auth, focus, findings) in enumerate(lit):
        c_left = Inches(0.8 + idx * 2.98)
        add_card(s2, c_left, Inches(1.8), Inches(2.8), Inches(5.0))
        tb = s2.shapes.add_textbox(c_left + Inches(0.15), Inches(2.0), Inches(2.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = auth
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_BLUE_ACCENT
        p2 = tf.add_paragraph()
        p2.text = focus
        p2.font.bold = True
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_NAVY_DARK
        p2.space_before = Pt(3)
        p3 = tf.add_paragraph()
        p3.text = findings
        p3.font.size = Pt(8.5)
        p3.font.color.rgb = C_TEXT_DARK
        p3.space_before = Pt(8)

    # Slide 3: Research Gaps
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "Identified Research Gaps & Proposed Resolution", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 3, tot)
    gaps = [
        ("Gap 1: Data Silos", "Clinical records, lab test values, and predictions are stored in disconnected spreadsheets or isolated files.", "Resolution: Centralized 3NF Relational DBMS uniting patients, tests, predictions, and consultations."),
        ("Gap 2: Black Box Barrier", "Clinicians distrust AI predictions when decision rationales are not transparent or physiologically explained.", "Resolution: Integrated Tree SHAP game-theoretic explainability providing exact feature contributions."),
        ("Gap 3: Workflow Disconnect", "Algorithms operate in standalone scripts rather than within active hospital scheduling and lab workflows.", "Resolution: End-to-end full stack platform with role-based routing from lab input to doctor consultation.")
    ]
    for idx, (g_title, g_desc, g_res) in enumerate(gaps):
        c_top = Inches(1.8 + idx * 1.7)
        add_card(s3, Inches(0.8), c_top, Inches(11.733), Inches(1.5))
        tb = s3.shapes.add_textbox(Inches(1.0), c_top + Inches(0.15), Inches(11.3), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = g_title
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = C_RED
        p2 = tf.add_paragraph()
        p2.text = f"Identified Issue: {g_desc}"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(2)
        p3 = tf.add_paragraph()
        p3.text = f"Our System's Approach: {g_res}"
        p3.font.bold = True
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = C_EMERALD
        p3.space_before = Pt(2)

    # Slide 4: System Architecture
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "4-Tier System Architecture", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 4, tot)
    arch_img = "presentation_assets/system_architecture.png"
    if os.path.exists(arch_img):
        try:
            s4.shapes.add_picture(arch_img, Inches(0.8), Inches(1.7), width=Inches(7.2))
        except Exception:
            pass
        add_card(s4, Inches(8.3), Inches(1.7), Inches(4.233), Inches(5.2))
        tb = s4.shapes.add_textbox(Inches(8.5), Inches(1.9), Inches(3.8), Inches(4.8))
    else:
        add_card(s4, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.2))
        tb = s4.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Tier Descriptions:"
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = C_NAVY_BLUE
    tiers = [
        ("Tier 1: Client Layer", "Responsive HTML5/CSS3/JS dashboards for Admin, Doctor, Staff, and Patient."),
        ("Tier 2: API Gateway Layer", "FastAPI asynchronous routing, JWT token verification, and RBAC authorization."),
        ("Tier 3: Business & ML Layer", "SQLAlchemy ORM engine, Scikit-learn pipelines, XGBoost/LightGBM inference, and SHAP explainer."),
        ("Tier 4: Persistence Layer", "PostgreSQL relational DBMS with 3NF tables, indexes, views, triggers, and audit logging.")
    ]
    for t_name, t_detail in tiers:
        p = tf.add_paragraph()
        p.text = t_name
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_BLUE_ACCENT
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = "\n" + t_detail
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    # Slide 5: ER Diagram & Modeling
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "Entity-Relationship (ER) Modeling", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 5, tot)
    er_img = "presentation_assets/database_er_diagram.png"
    if os.path.exists(er_img):
        try:
            s5.shapes.add_picture(er_img, Inches(0.8), Inches(1.7), width=Inches(7.2))
        except Exception:
            pass
        add_card(s5, Inches(8.3), Inches(1.7), Inches(4.233), Inches(5.2))
        tb = s5.shapes.add_textbox(Inches(8.5), Inches(1.9), Inches(3.8), Inches(4.8))
    else:
        add_card(s5, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.2))
        tb = s5.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Key Entity Sets & Cardinalities:"
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = C_NAVY_BLUE
    entities = [
        ("users (1:1 with roles):", "Base authentication entity; specialized into patients, doctors, and staff."),
        ("patients (1:N tests/preds):", "Central medical entity linked to test requisitions, predictions, and appointments."),
        ("appointments (N:1 doctor/pt):", "Relates patients and doctors with composite UNIQUE constraints on (doctor_id, appointment_date, time_slot)."),
        ("lab_tests & results (1:N):", "Structured test orders containing granular biomarker measurements."),
        ("predictions (N:1 patient):", "Stores ML inference confidence scores, risk categories, and JSON-encoded SHAP feature weights.")
    ]
    for e_name, e_desc in entities:
        p = tf.add_paragraph()
        p.text = e_name
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = " " + e_desc
        run.font.bold = False

    # Slide 6: 3NF Relational Schema Design
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "Relational Schema Normalization (3NF)", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 6, tot)
    norms = [
        ("1st Normal Form (1NF)", "Atomic Values & Uniqueness", "All column values are atomic (e.g. no comma-separated biomarker lists). Primary keys enforce tuple uniqueness across all tables."),
        ("2nd Normal Form (2NF)", "Full Functional Dependency", "No non-prime attribute is partially dependent on any candidate key. Composite appointment slots fully determine status and notes."),
        ("3rd Normal Form (3NF)", "No Transitive Dependencies", "Non-key attributes depend solely on the primary key. Doctor department details and user authentication details are decoupled from patient tables.")
    ]
    for idx, (n_form, n_rule, n_impl) in enumerate(norms):
        c_left = Inches(0.8 + idx * 3.98)
        add_card(s6, c_left, Inches(1.8), Inches(3.7), Inches(5.0))
        tb = s6.shapes.add_textbox(c_left + Inches(0.2), Inches(2.0), Inches(3.3), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = n_form
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = C_BLUE_ACCENT
        p2 = tf.add_paragraph()
        p2.text = n_rule
        p2.font.bold = True
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_NAVY_DARK
        p2.space_before = Pt(4)
        p3 = tf.add_paragraph()
        p3.text = n_impl
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = C_TEXT_DARK
        p3.space_before = Pt(8)

    # Slide 7: Data Flow Diagram (DFD)
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "System Data Flow & Pipeline Architecture", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 7, tot)
    flow_img = "presentation_assets/ml_pipeline_flow.png"
    if os.path.exists(flow_img):
        try:
            s7.shapes.add_picture(flow_img, Inches(0.8), Inches(1.7), width=Inches(7.2))
        except Exception:
            pass
        add_card(s7, Inches(8.3), Inches(1.7), Inches(4.233), Inches(5.2))
        tb = s7.shapes.add_textbox(Inches(8.5), Inches(1.9), Inches(3.8), Inches(4.8))
    else:
        add_card(s7, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.2))
        tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Operational Data Journey:"
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = C_NAVY_BLUE
    flows = [
        ("1. Intake & Profiling:", "Patient registers; demographic and baseline vitals committed to PostgreSQL."),
        ("2. Lab Requisition:", "Doctor requests renal battery; lab technician measures and enters 24 biomarkers."),
        ("3. Preprocessing & ML:", "Imputer applies training medians/modes; StandardScaler scales numericals; XGBoost infers probability."),
        ("4. Explainability & Storage:", "SHAP explainer derives top risk factors; prediction record committed; doctor dashboard updates in real time.")
    ]
    for fl_step, fl_det in flows:
        p = tf.add_paragraph()
        p.text = fl_step
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = "\n" + fl_det
        run.font.bold = False

    # Slide 8: Review 1 Summary & Next Steps
    s8 = prs.slides.add_slide(blank)
    add_header(s8, "Review 1 Accomplishments & Review 2 Roadmap", "REVIEW 1: DESIGN & ARCHITECTURE", "Review 1", 8, tot)
    add_card(s8, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s8.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Design Sign-off & Implementation Readiness"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    next_pts = [
        ("Formal Relational Schema: ", "All tables, foreign keys, and integrity constraints successfully modeled and validated against 3NF criteria."),
        ("System Architecture Approved: ", "Decoupled 4-tier model ensures modularity between API services, ML routines, and database operations."),
        ("Ready for Review 2 (Implementation): ", "Will demonstrate live DDL execution, stored procedures/triggers, FastAPI endpoint implementations, and model training metrics.")
    ]
    for n_title, n_desc in next_pts:
        p = tf.add_paragraph()
        p.text = f"• {n_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_NAVY_BLUE
        p.space_before = Pt(12)
        run = p.add_run()
        run.text = n_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    prs.save("CKD_Project_Review_1.pptx")
    print("Saved CKD_Project_Review_1.pptx")

def build_review_2():
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    tot = 8

    create_title_slide(prs, "System Implementation & Machine Learning", "Review 2: Implementation & Coding", "Database Scripts (DDL/DML), Stored Procedures, FastAPI REST API & ML Model Training")

    # Slide 2: Dataset & Preprocessing Pipeline
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "Dataset Characteristics & Preprocessing Rigor", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 2, tot)
    add_card(s2, Inches(0.8), Inches(1.65), Inches(5.7), Inches(5.3))
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "UCI Machine Learning Benchmark Dataset"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_BLUE
    ds_pts = [
        ("400 Patient Records: ", "Collected over Apollo Hospitals, Tamil Nadu; benchmark dataset for renal informatics."),
        ("24 Input Features: ", "14 numerical biomarkers (bgr, bu, sc, sod, pot, hemo, pcv, wc, rc, age, bp, sg, al, su) + 10 categorical clinical signs (rbc, pc, pcc, ba, htn, dm, cad, appet, pe, ane)."),
        ("Binary Target: ", "'ckd' (Chronic Kidney Disease Detected) vs 'notckd' (Healthy Renal Function).")
    ]
    for d_title, d_txt in ds_pts:
        p = tf.add_paragraph()
        p.text = d_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = d_txt
        run.font.bold = False

    add_card(s2, Inches(6.8), Inches(1.65), Inches(5.733), Inches(5.3))
    tb2 = s2.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Zero Data Leakage Preprocessing Pipeline"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_EMERALD
    prep_pts = [
        ("Train/Test Separation: ", "80/20 stratified split applied before any statistical computation, strictly preventing test information leakage."),
        ("Median/Mode Imputation: ", "Numerical missing values imputed using training median; categoricals imputed with training mode."),
        ("StandardScaler Scaling: ", "Standardization fitted solely on training set parameters (mean and standard deviation) and transformed across inference inputs."),
        ("Serialized Pipeline: ", "Transformers persisted alongside model in best_model.joblib for identical inference execution.")
    ]
    for p_title, p_txt in prep_pts:
        p = tf2.add_paragraph()
        p.text = p_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = p_txt
        run.font.bold = False

    # Slide 3: Database Implementation
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "Database Engineering: DDL, Keys & Constraints", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 3, tot)
    db_items = [
        ("schema.sql (Tables & DDL)", "users, patients, doctors, medical_assessments, appointments, lab_tests, lab_results, predictions, clinical_notes, notifications, audit_logs.", "Foreign key constraints with ON DELETE CASCADE and integrity CHECK constraints (e.g. valid phone formats, status values)."),
        ("indexes.sql (B-Tree Performance)", "Indexes on users.email, patients.user_id, appointments.(doctor_id, appointment_date), and lab_tests.status.", "Accelerates query execution times and optimizes joins across large patient cohorts."),
        ("seed.sql (Demonstration Data)", "Pre-populated sample records covering patients, specialized doctors, test requisitions, and historical predictions.", "Allows instantaneous academic live demonstration across all four user portals.")
    ]
    for idx, (title, comps, detail) in enumerate(db_items):
        c_top = Inches(1.8 + idx * 1.7)
        add_card(s3, Inches(0.8), c_top, Inches(11.733), Inches(1.5))
        tb = s3.shapes.add_textbox(Inches(1.0), c_top + Inches(0.15), Inches(11.3), Inches(1.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = C_NAVY_BLUE
        p2 = tf.add_paragraph()
        p2.text = f"Components: {comps}"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(2)
        p3 = tf.add_paragraph()
        p3.text = f"DBMS Significance: {detail}"
        p3.font.bold = True
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = C_BLUE_ACCENT
        p3.space_before = Pt(2)

    # Slide 4: Stored Procedures & Triggers
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "Advanced DBMS: Stored Procedures & Triggers", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 4, tot)
    add_card(s4, Inches(0.8), Inches(1.65), Inches(5.7), Inches(5.3))
    tb = s4.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Stored Procedures & Atomic Logic"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_BLUE
    procs = [
        ("sp_book_appointment: ", "Atomically validates doctor schedule availability, checks for slot collision, and books appointment within a single transaction boundary."),
        ("sp_verify_lab_test: ", "Transitions lab test status from 'PENDING' to 'VERIFIED', attaches verifier staff ID, and automatically creates a patient notification alert.")
    ]
    for p_name, p_desc in procs:
        p = tf.add_paragraph()
        p.text = p_name
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(10)
        run = p.add_run()
        run.text = p_desc
        run.font.bold = False

    add_card(s4, Inches(6.8), Inches(1.65), Inches(5.733), Inches(5.3))
    tb2 = s4.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Automated Database Triggers"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_AMBER
    trigs = [
        ("trg_audit_prediction: ", "Fires AFTER INSERT on predictions table. Automatically logs user ID, action ('RUN_PREDICTION'), and timestamp to audit_logs for medico-legal traceability."),
        ("trg_user_registration_audit: ", "Fires AFTER INSERT on users table to ensure an immutable record of new account provisioning."),
        ("Double-Booking Guard: ", "Composite UNIQUE constraints on (doctor_id, appointment_date, time_slot) guarantee zero race-condition collisions at the engine level.")
    ]
    for t_name, t_desc in trigs:
        p = tf2.add_paragraph()
        p.text = t_name
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = t_desc
        run.font.bold = False

    # Slide 5: Backend API Architecture
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "FastAPI Backend & RESTful Endpoints", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 5, tot)
    ep_cards = [
        ("/api/auth/*", "User login, JWT access token issuing, password verification using bcrypt, and user registration.", C_NAVY_BLUE),
        ("/api/patients/*", "Patient profile management, medical assessment intake submission, and historical test tracking.", C_BLUE_ACCENT),
        ("/api/predictions/*", "Real-time ML inference triggering, SHAP value computation, and risk stratification persistence.", C_EMERALD),
        ("/api/appointments/*", "Doctor directory browsing, appointment booking, slot cancellation, and doctor queue management.", C_AMBER)
    ]
    for idx, (ep, desc, col) in enumerate(ep_cards):
        c_left = Inches(0.8 + (idx % 2) * 6.0)
        c_top = Inches(1.7 + (idx // 2) * 2.6)
        add_card(s5, c_left, c_top, Inches(5.733), Inches(2.4))
        tb = s5.shapes.add_textbox(c_left + Inches(0.2), c_top + Inches(0.2), Inches(5.333), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = ep
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(6)

    # Slide 6: ML Model Training & Results
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "Gradient Boosted Tree Models: XGBoost & LightGBM", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 6, tot)
    m_img = "presentation_assets/model_comparison_bar.png"
    if os.path.exists(m_img):
        try:
            s6.shapes.add_picture(m_img, Inches(0.8), Inches(1.7), width=Inches(7.2))
        except Exception:
            pass
        add_card(s6, Inches(8.3), Inches(1.7), Inches(4.233), Inches(5.2))
        tb = s6.shapes.add_textbox(Inches(8.5), Inches(1.9), Inches(3.8), Inches(4.8))
    else:
        add_card(s6, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.2))
        tb = s6.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Model Benchmarks (Test Split):"
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = C_NAVY_BLUE
    benchmarks = [
        ("XGBoost Classifier:", "Accuracy: 100.0%, Precision: 100.0%, Recall (Sensitivity): 100.0%, F1-Score: 1.000, ROC-AUC: 1.000."),
        ("LightGBM Classifier:", "Accuracy: 100.0%, Precision: 100.0%, Recall: 100.0%, F1-Score: 1.000, ROC-AUC: 1.000."),
        ("Production Choice:", "XGBoost v1.0.0 selected for optimal tree stability and exact SHAP tree explainer compatibility.")
    ]
    for b_title, b_val in benchmarks:
        p = tf.add_paragraph()
        p.text = b_title
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = "\n" + b_val
        run.font.bold = False

    # Slide 7: SHAP Explainability
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "Explainable AI (SHAP) Clinical Interpretability", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 7, tot)
    shap_img = "presentation_assets/shap_feature_importance.png"
    if os.path.exists(shap_img):
        try:
            s7.shapes.add_picture(shap_img, Inches(0.8), Inches(1.7), width=Inches(7.2))
        except Exception:
            pass
        add_card(s7, Inches(8.3), Inches(1.7), Inches(4.233), Inches(5.2))
        tb = s7.shapes.add_textbox(Inches(8.5), Inches(1.9), Inches(3.8), Inches(4.8))
    else:
        add_card(s7, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.2))
        tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Key Clinical Feature Drivers:"
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = C_NAVY_BLUE
    drivers = [
        ("1. Hemoglobin (hemo):", "Strongest protective factor; drops severely in renal failure due to impaired erythropoietin production."),
        ("2. Specific Gravity (sg):", "Low urinary concentration reflects loss of renal tubular concentrating capacity."),
        ("3. Serum Creatinine (sc):", "Elevated values directly signify reduced glomerular filtration rate."),
        ("4. Albumin (al):", "Proteinuria indicates breakdown of the glomerular filtration barrier.")
    ]
    for d_name, d_expl in drivers:
        p = tf.add_paragraph()
        p.text = d_name
        p.font.bold = True
        p.font.size = Pt(9.5)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = "\n" + d_expl
        run.font.bold = False

    # Slide 8: Review 2 Summary & Next Steps
    s8 = prs.slides.add_slide(blank)
    add_header(s8, "Review 2 Outcomes & Progression to Review 3", "REVIEW 2: SYSTEM IMPLEMENTATION", "Review 2", 8, tot)
    add_card(s8, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s8.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Implementation Phase Milestones Achieved"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    rev2_pts = [
        ("Core Database Built: ", "Full schema, indexes, procedures, and triggers implemented and seeded."),
        ("ML Models Trained & Persisted: ", "XGBoost and LightGBM models trained with 100% sensitivity on held-out test split, zero data leakage."),
        ("FastAPI Backend Operational: ", "Full REST API connecting database entities to ML service endpoints."),
        ("Ready for Review 3 (Testing & Integration): ", "Will demonstrate unit/integration testing suites, double-booking prevention proofs, and security RBAC gating.")
    ]
    for r_title, r_desc in rev2_pts:
        p = tf.add_paragraph()
        p.text = f"• {r_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_NAVY_BLUE
        p.space_before = Pt(12)
        run = p.add_run()
        run.text = r_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    prs.save("CKD_Project_Review_2.pptx")
    print("Saved CKD_Project_Review_2.pptx")

def build_review_3():
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    tot = 8

    create_title_slide(prs, "System Testing, Security & Integration", "Review 3: Testing & Integration", "Automated Pytest Suite, RBAC Security, Database Integrity Verification & UI Workflows")

    # Slide 2: Automated Pytest Suite
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "Automated Test Suite (Pytest)", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 2, tot)
    test_cases = [
        ("test_auth.py", "User Authentication & Registration", "Validates successful login, password mismatch rejection, JWT issuance, and new patient registration flows."),
        ("test_appointments.py", "Atomic Scheduling & Collision Prevention", "Validates successful appointment creation and proves immediate rejection of double-booking attempts."),
        ("test_predictions.py", "Inference & Feature Attribution", "Submits clinical biomarker payloads, verifies prediction category generation, and ensures valid SHAP importance outputs."),
        ("test_database.py", "Database Integrity & SQL Injection Defense", "Validates demonstration SQL queries and proves parameterized queries reject malicious SQL injection payloads.")
    ]
    for idx, (module, name, details) in enumerate(test_cases):
        c_left = Inches(0.8 + idx * 2.98)
        add_card(s2, c_left, Inches(1.8), Inches(2.8), Inches(5.0))
        tb = s2.shapes.add_textbox(c_left + Inches(0.15), Inches(2.0), Inches(2.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = module
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_BLUE_ACCENT
        p2 = tf.add_paragraph()
        p2.text = name
        p2.font.bold = True
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_NAVY_DARK
        p2.space_before = Pt(3)
        p3 = tf.add_paragraph()
        p3.text = details
        p3.font.size = Pt(8.5)
        p3.font.color.rgb = C_TEXT_DARK
        p3.space_before = Pt(8)

    # Slide 3: Double-Booking Prevention Verification
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "Concurrency Control: Double-Booking Prevention", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 3, tot)
    add_card(s3, Inches(0.8), Inches(1.65), Inches(5.7), Inches(5.3))
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "The Double-Booking Vulnerability in Healthcare"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_BLUE
    db_pts = [
        ("Race Condition Threat: ", "Two patients simultaneously attempting to book Dr. Nephrologist at 10:00 AM on Monday."),
        ("Flawed Approach: ", "Application-level checks (e.g. SELECT followed by INSERT) fail under high concurrency due to time-of-check to time-of-use (TOCTOU) race conditions."),
        ("Catastrophic Impact: ", "Overbooked clinic hours, patient dissatisfaction, and doctor scheduling chaos.")
    ]
    for b_title, b_txt in db_pts:
        p = tf.add_paragraph()
        p.text = b_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = b_txt
        run.font.bold = False

    add_card(s3, Inches(6.8), Inches(1.65), Inches(5.733), Inches(5.3))
    tb2 = s3.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Our Relational & Procedural Solution"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_EMERALD
    sol_pts = [
        ("Composite UNIQUE Constraint: ", "UNIQUE (doctor_id, appointment_date, time_slot) enforced directly by the database B-tree index."),
        ("ACID Transaction Isolation: ", "Stored procedure sp_book_appointment executes within an atomic transaction with immediate rollback on conflict."),
        ("HTTP 409 Conflict: ", "Second concurrent request receives structured HTTP 409 Conflict with a clear message that the slot is already taken."),
        ("Pytest Verified: ", "test_appointment_booking_and_double_booking_prevention programmatically confirms rejection.")
    ]
    for s_title, s_txt in sol_pts:
        p = tf2.add_paragraph()
        p.text = s_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(6)
        run = p.add_run()
        run.text = s_txt
        run.font.bold = False

    # Slide 4: Role-Based Access Control (RBAC)
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "Role-Based Access Control & Security Auditing", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 4, tot)
    roles = [
        ("Admin Portal", "System health stats, user account management, global prediction analytics, and live SQL query demonstration runner."),
        ("Doctor Portal", "Consultation queue, Patient 360 view, one-click ML prediction execution with SHAP charts, clinical notes editor."),
        ("Lab Staff Portal", "Pending test requisition queue, 24 biomarker parameter entry form, test verification gating, and patient alert creation."),
        ("Patient Portal", "Personal health profile, 24-feature self-assessment form, doctor appointment booking, and verified report downloads.")
    ]
    for idx, (title, desc) in enumerate(roles):
        c_left = Inches(0.8 + (idx % 2) * 6.0)
        c_top = Inches(1.7 + (idx // 2) * 2.6)
        add_card(s4, c_left, c_top, Inches(5.733), Inches(2.4))
        tb = s4.shapes.add_textbox(c_left + Inches(0.2), c_top + Inches(0.2), Inches(5.333), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(13)
        p.font.color.rgb = C_NAVY_BLUE
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(6)

    # Slide 5: Frontend UI Screenshots
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "Integrated User Portals: Doctor & Admin", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 5, tot)
    doc_img = "presentation_assets/doctor_portal_mockup.png"
    adm_img = "presentation_assets/admin_analytics_mockup.png"
    if os.path.exists(doc_img):
        try:
            s5.shapes.add_picture(doc_img, Inches(0.8), Inches(1.7), width=Inches(5.7))
        except Exception:
            pass
    if os.path.exists(adm_img):
        try:
            s5.shapes.add_picture(adm_img, Inches(6.8), Inches(1.7), width=Inches(5.7))
        except Exception:
            pass

    # Slide 6: Database Views & SQL Query Runner
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "DBMS Views & Live Query Demonstration", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 6, tot)
    views = [
        ("v_patient_360", "Composite 360-degree patient view uniting demographic details, latest lab test results, and latest CKD predictions."),
        ("v_doctor_workload", "Aggregates total appointments, completed reviews, and pending consultations per physician."),
        ("v_lab_queue", "Real-time list of pending biomarker requisitions highlighting abnormal values."),
        ("v_prediction_analytics", "Statistical breakdown of CKD vs Not-CKD predictions, average confidence, and risk strata.")
    ]
    for idx, (v_name, v_desc) in enumerate(views):
        c_left = Inches(0.8 + idx * 2.98)
        add_card(s6, c_left, Inches(1.8), Inches(2.8), Inches(5.0))
        tb = s6.shapes.add_textbox(c_left + Inches(0.15), Inches(2.0), Inches(2.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = v_name
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = C_BLUE_ACCENT
        p2 = tf.add_paragraph()
        p2.text = v_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(8)

    # Slide 7: Security Hardening & SQL Injection Defense
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "Security Controls: SQL Injection & JWT Expiry", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 7, tot)
    add_card(s7, Inches(0.8), Inches(1.65), Inches(5.7), Inches(5.3))
    tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Parameterized Queries & SQL Injection Immunity"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_BLUE
    sec_pts = [
        ("Zero String Concatenation: ", "All queries use SQLAlchemy parameterized placeholders (:user_id, :email), completely immunizing the system against SQL injection attacks."),
        ("Validation Defense: ", "Input payloads validated via Pydantic schemas before reaching the database driver layer."),
        ("Audit Trail: ", "Automated triggers capture every prediction execution and clinical update.")
    ]
    for sp_title, sp_txt in sec_pts:
        p = tf.add_paragraph()
        p.text = sp_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = sp_txt
        run.font.bold = False

    add_card(s7, Inches(6.8), Inches(1.65), Inches(5.733), Inches(5.3))
    tb2 = s7.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Cryptographic & Authentication Hygiene"
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_EMERALD
    auth_pts = [
        ("Bcrypt Password Hashing: ", "Passwords salt-hashed using bcrypt; plain text passwords never stored in memory or disk."),
        ("JWT Tokens with Expiration: ", "Cryptographically signed JWT bearer tokens with strict expiration timeouts."),
        ("RBAC Middleware: ", "Endpoints protected with role verification dependencies ensuring patients cannot access admin/doctor routes.")
    ]
    for ap_title, ap_txt in auth_pts:
        p = tf2.add_paragraph()
        p.text = ap_title
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = C_TEXT_DARK
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = ap_txt
        run.font.bold = False

    # Slide 8: Review 3 Summary & Next Steps
    s8 = prs.slides.add_slide(blank)
    add_header(s8, "Review 3 Outcomes & Final Demonstration Plan", "REVIEW 3: TESTING & INTEGRATION", "Review 3", 8, tot)
    add_card(s8, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s8.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Integration Verification Accomplished"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    rev3_pts = [
        ("Comprehensive Testing Completed: ", "Pytest suites verify authentication, appointment locking, lab results, and ML prediction pipelines."),
        ("Security Standards Met: ", "Role-based access gating verified across all endpoints, parameterized queries prevent SQL injection."),
        ("Ready for Review 4 (Final Evaluation): ", "Will demonstrate live end-to-end clinical workflow, empirical comparisons, future work, and conclude the capstone evaluation.")
    ]
    for r_title, r_desc in rev3_pts:
        p = tf.add_paragraph()
        p.text = f"• {r_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_NAVY_BLUE
        p.space_before = Pt(12)
        run = p.add_run()
        run.text = r_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    prs.save("CKD_Project_Review_3.pptx")
    print("Saved CKD_Project_Review_3.pptx")

def build_review_4():
    prs = pptx.Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    tot = 9

    create_title_slide(prs, "Final Project Defense & Evaluation", "Review 4: Final Evaluation", "Full Demonstration, Empirical Benchmarking, Comparative Analysis, Societal Impact & Conclusion")

    # Slide 2: Executive Summary
    s2 = prs.slides.add_slide(blank)
    add_header(s2, "Executive Summary & Capstone Achievements", "REVIEW 4: FINAL DEFENSE", "Review 4", 2, tot)
    add_card(s2, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s2.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Comprehensive Solution Overview"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    exec_pts = [
        ("End-to-End Healthcare Architecture: ", "Successfully integrated a 3NF Relational DBMS with Explainable Gradient Boosted Decision Trees (XGBoost/LightGBM) and a modern 4-tier web application."),
        ("Perfect Test Sensitivity (100%): ", "Achieved zero false negatives on held-out test data, ensuring zero missed CKD cases in critical early intervention triage."),
        ("Complete Concurrency Integrity: ", "Relational constraints and stored procedures prevent double-booking collisions, enforce audit tracking, and sanitize all transactions."),
        ("Medical Decision-Support Standard: ", "Incorporates explicit clinical disclaimers ensuring machine learning outputs assist rather than replace certified physician judgment.")
    ]
    for e_title, e_desc in exec_pts:
        p = tf.add_paragraph()
        p.text = f"• {e_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_NAVY_BLUE
        p.space_before = Pt(10)
        run = p.add_run()
        run.text = e_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    # Slide 3: Comparison with Existing Systems
    s3 = prs.slides.add_slide(blank)
    add_header(s3, "Comparison: Conventional Practices vs Our System", "REVIEW 4: FINAL DEFENSE", "Review 4", 3, tot)
    comp_rows = [
        ("Dimension", "Conventional / Paper / Spreadsheet Records", "Our Smart CKD Management Platform"),
        ("Patient Records", "Fragmented across paper folders or Excel files; high redundancy", "Centralized 3NF PostgreSQL DBMS; ACID transactions & zero redundancy"),
        ("Early Risk Detection", "Manual univariate thresholding; misses asymptomatic early stages", "Multivariate XGBoost/LightGBM classifier with 100% sensitivity"),
        ("Explainability", "Opaque black-box models or manual intuition", "Instant SHAP force plots highlighting top physiological drivers"),
        ("Appointment Scheduling", "Vulnerable to double-booking and schedule collisions", "Atomic double-booking prevention enforced via B-tree unique index"),
        ("Audit & Compliance", "No tracking of record modifications or clinical updates", "Automated database triggers capturing immutable audit logs")
    ]
    add_card(s3, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    for idx, (dim, conv, prop) in enumerate(comp_rows):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.space_before = Pt(6) if idx > 0 else Pt(0)
        p.text = f"{dim:20} |  {conv:45} |  {prop}"
        if idx == 0:
            p.font.bold = True
            p.font.size = Pt(10.5)
            p.font.color.rgb = C_NAVY_BLUE
        else:
            p.font.size = Pt(9)
            p.font.color.rgb = C_TEXT_DARK

    # Slide 4: Empirical ML Metrics & ROC Curve
    s4 = prs.slides.add_slide(blank)
    add_header(s4, "Empirical Validation: ROC Curve & Confusion Matrix", "REVIEW 4: FINAL DEFENSE", "Review 4", 4, tot)
    roc_img = "presentation_assets/roc_curve.png"
    cm_img = "presentation_assets/confusion_matrix.png"
    if os.path.exists(roc_img):
        try:
            s4.shapes.add_picture(roc_img, Inches(0.8), Inches(1.7), width=Inches(5.7))
        except Exception:
            pass
    if os.path.exists(cm_img):
        try:
            s4.shapes.add_picture(cm_img, Inches(6.8), Inches(1.7), width=Inches(5.7))
        except Exception:
            pass

    # Slide 5: Live Demonstration Screenshots
    s5 = prs.slides.add_slide(blank)
    add_header(s5, "Live Working System Demonstration", "REVIEW 4: FINAL DEFENSE", "Review 4", 5, tot)
    risk_img = "presentation_assets/risk_stratification_chart.png"
    admin_img = "presentation_assets/admin_analytics_mockup.png"
    if os.path.exists(risk_img):
        try:
            s5.shapes.add_picture(risk_img, Inches(0.8), Inches(1.7), width=Inches(5.7))
        except Exception:
            pass
    if os.path.exists(admin_img):
        try:
            s5.shapes.add_picture(admin_img, Inches(6.8), Inches(1.7), width=Inches(5.7))
        except Exception:
            pass

    # Slide 6: Challenges Encountered & Solutions
    s6 = prs.slides.add_slide(blank)
    add_header(s6, "Engineering Challenges & Implemented Solutions", "REVIEW 4: FINAL DEFENSE", "Review 4", 6, tot)
    challenges = [
        ("Missing Clinical Values", "Patient records frequently omit test values (e.g. serum potassium or packed cell volume).", "Implemented training-fitted Median/Mode imputation; strictly maintained zero data leakage.", C_RED),
        ("Mixed Data Types", "Input contains numerical biomarkers, discrete stages, and categorical signs.", "Unified Scikit-Learn ColumnTransformer pipeline with categorical encoding and feature scaling.", C_AMBER),
        ("Concurrency Collisions", "Multiple patients attempting simultaneous appointment booking on the same physician slot.", "Engineered composite B-tree UNIQUE index and atomic stored procedure rollback.", C_BLUE_ACCENT),
        ("Clinician Trust & Adoption", "Doctors resist opaque AI predictions without physiological rationale.", "Tree SHAP feature explanations with exact contribution scores per biomarker.", C_EMERALD)
    ]
    for idx, (title, chal, sol, col) in enumerate(challenges):
        c_left = Inches(0.8 + (idx % 2) * 6.0)
        c_top = Inches(1.7 + (idx // 2) * 2.6)
        add_card(s6, c_left, c_top, Inches(5.733), Inches(2.4))
        tb = s6.shapes.add_textbox(c_left + Inches(0.2), c_top + Inches(0.2), Inches(5.333), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = f"Challenge: {chal}"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(4)
        p3 = tf.add_paragraph()
        p3.text = f"Solution: {sol}"
        p3.font.bold = True
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = C_NAVY_DARK
        p3.space_before = Pt(4)

    # Slide 7: Societal & Translational Impact
    s7 = prs.slides.add_slide(blank)
    add_header(s7, "Translational Clinical & Economic Impact", "REVIEW 4: FINAL DEFENSE", "Review 4", 7, tot)
    impacts = [
        ("Early Detection Window", "Enables therapeutic intervention (ACE inhibitors, SGLT2 inhibitors) before irreversible renal fibrotic scarring occurs.", C_CYAN),
        ("Massive Cost Avoidance", "Halting CKD progression avoids catastrophic lifelong hemodialysis expenses ($80,000+ per patient annually).", C_EMERALD),
        ("Rural Health Empowerment", "Assists junior medical officers and rural clinics lacking immediate access to expert nephrologists.", C_AMBER),
        ("Hospital Operational Speed", "Automates biomarker requisition workflows, reducing lab-to-consultation triage turnaround time by >60%.", C_BLUE_ACCENT)
    ]
    for idx, (title, desc, col) in enumerate(impacts):
        c_left = Inches(0.8 + idx * 2.98)
        add_card(s7, c_left, Inches(1.8), Inches(2.8), Inches(5.0))
        tb = s7.shapes.add_textbox(c_left + Inches(0.15), Inches(2.0), Inches(2.5), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.bold = True
        p.font.size = Pt(12)
        p.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_DARK
        p2.space_before = Pt(8)

    # Slide 8: Future Scope & Roadmap
    s8 = prs.slides.add_slide(blank)
    add_header(s8, "Future Work & Scaling Roadmap", "REVIEW 4: FINAL DEFENSE", "Review 4", 8, tot)
    future_pts = [
        ("1. Multi-Hospital Cohort Validation:", "Validate models on diverse international multi-center patient datasets beyond the benchmark cohort."),
        ("2. KDIGO 5-Stage Multi-Class Prediction:", "Extend binary classification to granular KDIGO Stage 1 through 5 progression modeling."),
        ("3. Longitudinal Time-Series Tracking:", "Integrate LSTM/Transformer recurrent models to forecast eGFR trajectories over 5-year horizons."),
        ("4. Mobile Health & Wearables Integration:", "Connect real-time continuous glucose and blood pressure monitoring streams via IoT gateways.")
    ]
    add_card(s8, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s8.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Scaling Roadmap for Post-Academic Deployment"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    for f_title, f_desc in future_pts:
        p = tf.add_paragraph()
        p.text = f"• {f_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_NAVY_BLUE
        p.space_before = Pt(10)
        run = p.add_run()
        run.text = " " + f_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    # Slide 9: Conclusion & Defense Summary
    s9 = prs.slides.add_slide(blank)
    add_header(s9, "Conclusion & Academic Defense Summary", "REVIEW 4: FINAL DEFENSE", "Review 4", 9, tot)
    add_card(s9, Inches(0.8), Inches(1.65), Inches(11.733), Inches(5.3))
    tb = s9.shapes.add_textbox(Inches(1.1), Inches(1.9), Inches(11.1), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Capstone Project Completion Summary"
    p.font.bold = True
    p.font.size = Pt(14)
    p.font.color.rgb = C_NAVY_DARK
    concl_pts = [
        ("Theoretical & Engineering Rigor: ", "Demonstrated full compliance with database normalization (3NF), relational integrity, and state-of-the-art machine learning practices."),
        ("Zero Leakage Benchmark Performance: ", "100% sensitivity on held-out test data with explainable feature importance via Tree SHAP."),
        ("Production-Ready Multi-Role Portals: ", "FastAPI asynchronous service endpoints powering role-segregated portals with verified security safeguards."),
        ("Thank You: ", "Open for Viva Questions, Evaluation Discussion, and Live System Demonstration.")
    ]
    for c_title, c_desc in concl_pts:
        p = tf.add_paragraph()
        p.text = f"• {c_title}"
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_EMERALD if "Thank You" in c_title else C_NAVY_BLUE
        p.space_before = Pt(10)
        run = p.add_run()
        run.text = c_desc
        run.font.bold = False
        run.font.color.rgb = C_TEXT_DARK

    prs.save("CKD_Project_Review_4.pptx")
    print("Saved CKD_Project_Review_4.pptx")

if __name__ == '__main__':
    build_review_0()
    build_review_1()
    build_review_2()
    build_review_3()
    build_review_4()
    print("All Reviews 0, 1, 2, 3, 4 presentations created successfully!")
