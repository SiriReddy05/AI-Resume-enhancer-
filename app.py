from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import PyPDF2
import docx
import re
import json
import os
from groq import Groq
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib import colors

# Add these new imports for DOCX generation
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
# FIX: Allow CORS for all domains so Vercel can access it!
CORS(app, resources={r"/*": {"origins": "*"}})

import os

# Change your Groq client line to this:
client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

@app.route('/')
def home():
    return "Backend is running!"

# 📄 Extract PDF
def extract_pdf(file_path):
    text = ""
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            if page.extract_text():
                text += page.extract_text() + "\n"
    return text

# 📄 Extract DOCX
def extract_docx(file_path):
    doc = docx.Document(file_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# 🧠 AI Enhancement (Strict JSON Structured Output)
def enhance_with_ai(raw_text):
    try:
        prompt = f"""
        You are an expert ATS-friendly resume writer. Extract and greatly enhance the following resume text.
        Improve the language, use strong action verbs for bullet points, fix grammar, and make it highly professional.
        
        You MUST return ONLY a valid JSON object. Do not include any explanations, markdown formatting, or markdown code blocks (no ```json).
        Use this EXACT JSON structure:
        {{
          "name": "Full Name",
          "contact": "Email | Phone | LinkedIn/Links",
          "summary": "A highly professional summary paragraph...",
          "skills": ["Skill 1", "Skill 2", "Skill 3"],
          "experience": [
            {{
              "title": "Job Title or Role",
              "subtitle": "Company Name",
              "date": "Duration (e.g., Jan 2021 - Present)",
              "details": ["High-impact bullet point 1", "High-impact bullet point 2"]
            }}
          ],
          "education": [
            {{
              "title": "Degree Name",
              "subtitle": "Institution Name",
              "date": "Graduation Year",
              "details": ["Optional academic detail or empty string"]
            }}
          ],
          "projects": [
            {{
              "title": "Project Name",
              "subtitle": "Technologies Used",
              "date": "Year or Duration",
              "details": ["Action-oriented bullet 1", "Action-oriented bullet 2"]
            }}
          ],
          "certifications": ["Certification Name 1", "Certification Name 2"]
        }}

        RAW RESUME TEXT:
        {raw_text}
        """

        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3 # Lower temperature for more consistent formatting
        )

        response_text = response.choices[0].message.content.strip()
        
        # Clean up any potential markdown formatting the AI might add by mistake
        if response_text.startswith("```json"):
            response_text = response_text[7:]
        if response_text.startswith("```"):
            response_text = response_text[3:]
        if response_text.endswith("```"):
            response_text = response_text[:-3]

        return json.loads(response_text.strip())

    except Exception as e:
        print("AI ERROR:", str(e))
        return {"error": "AI failed to process the resume properly."}

# 📄 Generate Clean PDF with Strict Formatting
# 📄 Generate Clean, 1-Page ATS PDF with Strict Formatting
def generate_pdf(data, filename="improved_resume.pdf"):
    # FIX: Drastically reduced margins to fit more on one page
    doc = SimpleDocTemplate(filename, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    # FIX: Tighter fonts and spacing
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], alignment=TA_CENTER, fontSize=16, spaceAfter=2)
    contact_style = ParagraphStyle(name='ContactStyle', parent=styles['Normal'], alignment=TA_CENTER, fontSize=9, textColor=colors.dimgrey, spaceAfter=8)
    section_heading = ParagraphStyle(name='SectionHeading', parent=styles['Heading2'], fontSize=11, spaceBefore=8, spaceAfter=2, textTransform='uppercase')
    item_title = ParagraphStyle(name='ItemTitle', parent=styles['Heading3'], fontSize=10, spaceBefore=4, spaceAfter=1)
    item_subtitle = ParagraphStyle(name='ItemSubtitle', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Oblique', textColor=colors.dimgrey, spaceAfter=2)
    normal_text = ParagraphStyle(name='NormalText', parent=styles['Normal'], fontSize=9, spaceAfter=4, leading=11)
    bullet_text = ParagraphStyle(name='BulletText', parent=styles['Normal'], fontSize=9, spaceAfter=1, leading=11)

    elements = []

    # 1. Header (Name & Contact)
    elements.append(Paragraph(f"<b>{data.get('name', 'Your Name')}</b>", title_style))
    elements.append(Paragraph(data.get('contact', 'Email | Phone'), contact_style))
    
    def add_section_header(title):
        elements.append(Paragraph(f"<b>{title}</b>", section_heading))
        elements.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceBefore=1, spaceAfter=4))

    # 2. Professional Summary
    if data.get('summary'):
        add_section_header("Professional Summary")
        elements.append(Paragraph(data['summary'], normal_text))

    # 3. Skills
    if data.get('skills') and len(data['skills']) > 0:
        add_section_header("Skills")
        skills_text = " • ".join(data['skills'])
        elements.append(Paragraph(skills_text, normal_text))

    # Helper for repetitive sections
    def add_list_section(title, key):
        items = data.get(key, [])
        if items and len(items) > 0:
            add_section_header(title)
            for item in items:
                title_str = f"<b>{item.get('title', '')}</b>"
                if item.get('date'):
                    title_str += f" <font color='grey'>| {item.get('date', '')}</font>"
                elements.append(Paragraph(title_str, item_title))
                
                if item.get('subtitle'):
                    elements.append(Paragraph(item.get('subtitle', ''), item_subtitle))
                
                # Bullet points
                if item.get('details') and isinstance(item['details'], list):
                    bullet_items = []
                    for detail in item['details']:
                        if str(detail).strip():
                            bullet_items.append(ListItem(Paragraph(str(detail), bullet_text), leftIndent=10, bulletOffsetY=-1))
                    if bullet_items:
                        elements.append(ListFlowable(bullet_items, bulletType='bullet', start='bulletchar', bulletFontName='Helvetica'))
                        elements.append(Spacer(1, 2))
    # 4. Experience, Education, Projects
    add_list_section("Experience", "experience")
    add_list_section("Projects", "projects") # Moved projects up for better flow
    add_list_section("Education", "education")
    # 5. Certifications
    if data.get('certifications') and len(data['certifications']) > 0:
        add_section_header("Certifications")
        cert_items = [ListItem(Paragraph(str(cert), bullet_text), leftIndent=10) for cert in data['certifications'] if str(cert).strip()]
        if cert_items:
             elements.append(ListFlowable(cert_items, bulletType='bullet', start='bulletchar'))

    # Build PDF
    doc.build(elements)
    return filename

# 📄 Generate Clean DOCX (NEW FUNCTION WITH SAFE FALLBACK)
def generate_docx(data, filename="improved_resume.docx"):
    doc = docx.Document()
    
    # 1. Header
    name_p = doc.add_heading(data.get('name', 'Your Name'), level=1)
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    contact_p = doc.add_paragraph(data.get('contact', 'Email | Phone'))
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def add_section_header(title):
        doc.add_heading(title, level=2)
        
    # 2. Professional Summary
    if data.get('summary'):
        add_section_header("Professional Summary")
        doc.add_paragraph(data['summary'])
        
    # 3. Skills
    if data.get('skills') and len(data['skills']) > 0:
        add_section_header("Skills")
        doc.add_paragraph(" • ".join(data['skills']))
        
    # Helper for repetitive sections (Experience, Education, Projects)
    def add_list_section(title, key):
        items = data.get(key, [])
        if items and len(items) > 0:
            add_section_header(title)
            for item in items:
                p = doc.add_paragraph()
                p.add_run(item.get('title', '')).bold = True
                if item.get('date'):
                    p.add_run(f"  |  {item.get('date', '')}")
                
                if item.get('subtitle'):
                    # FIX: SAFE FALLBACK manual italic
                    p_sub = doc.add_paragraph()
                    p_sub.add_run(item.get('subtitle', '')).italic = True
                    
                if item.get('details') and isinstance(item['details'], list):
                    for detail in item['details']:
                        if str(detail).strip():
                            # FIX: SAFE FALLBACK manual bullet
                            doc.add_paragraph(f"•  {str(detail)}")
                            
    # 4. Experience, Education, Projects
    add_list_section("Experience", "experience")
    add_list_section("Education", "education")
    add_list_section("Projects", "projects")
    
    # 5. Certifications
    if data.get('certifications') and len(data['certifications']) > 0:
        add_section_header("Certifications")
        for cert in data['certifications']:
            if str(cert).strip():
                doc.add_paragraph(f"•  {str(cert)}")
                
    doc.save(filename)
    return filename

# 📤 Upload & Process API
@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'resume' not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
            
        file = request.files['resume']
        file_path = "temp_" + file.filename
        file.save(file_path)

        # Extract Text
        text = ""
        if file.filename.lower().endswith(".pdf"):
            text = extract_pdf(file_path)
        elif file.filename.lower().endswith(".docx"):
            text = extract_docx(file_path)
        else:
            return jsonify({"error": "Unsupported format"}), 400

        # AI Enhance & Structure into JSON
        structured_data = enhance_with_ai(text)
        
        if "error" in structured_data:
             return jsonify(structured_data), 500

        # Generate BOTH formatted PDF and DOCX
        pdf_file = generate_pdf(structured_data)
        docx_file = generate_docx(structured_data)
        
        # Cleanup temp file
        if os.path.exists(file_path):
            os.remove(file_path)

        return jsonify({
            "resumeData": structured_data,
            "downloadUrl": f"http://127.0.0.1:5000/download/{pdf_file}"
        })

    except Exception as e:
        print("ERROR:", str(e))
        return jsonify({"error": str(e)}), 500

# 📥 Download API
@app.route('/download/<filename>')
def download_file(filename):
    if not os.path.exists(filename):
        return "File not found! Please click 'Upload & Enhance' to generate this document again.", 404
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)