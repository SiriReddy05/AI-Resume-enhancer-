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

from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

client = Groq(api_key=os.environ.get("GROQ_API_KEY"))

@app.route('/')
def home():
    return "Backend is running!"

def extract_pdf(file_path):
    text = ""
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            if page.extract_text():
                text += page.extract_text() + "\n"
    return text

def extract_docx(file_path):
    doc = docx.Document(file_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def enhance_with_ai(raw_text):
    try:
        system_prompt = """
        You are an expert ATS-friendly resume writer and a meticulous data extractor.
        
        CRITICAL RULES:
        1. YOU ARE FORBIDDEN FROM DELETING OR SUMMARIZING INFORMATION.
        2. You MUST retain EVERY single job experience, EVERY project, EVERY certification, EVERY achievement, and EVERY bullet point.
        3. EXPLICIT EXTRACTION: You must actively search for sections titled "Scholastic Achievements", "Position of Responsibility", "Leadership", or similar. Extract EVERY SINGLE bullet point under these headings and place them directly into the "achievements" array. Do not leave them behind.
        4. CATEGORIZED SKILLS: Group the candidate's skills into logical categories (e.g., Languages, Frameworks, Tools & Technologies, Soft Skills). Return them as an array of objects.
        5. Do not shorten the resume. Improve grammar and use strong action verbs.
        
        You MUST return ONLY a valid JSON object. Do not include any explanations, markdown formatting, or markdown code blocks (no ```json).
        Use this EXACT JSON structure:
        {
          "name": "Full Name",
          "contact": "Email | Phone | LinkedIn/Links",
          "summary": "A highly professional summary paragraph...",
          "skills": [
            {"category": "Languages", "items": "Python, JavaScript, SQL"},
            {"category": "Frameworks", "items": "React, Django, Flask"}
          ],
          "experience": [
            {
              "title": "Job Title or Role",
              "subtitle": "Company Name",
              "date": "Duration (e.g., Jan 2021 - Present)",
              "details": ["High-impact bullet point 1", "High-impact bullet point 2"]
            }
          ],
          "education": [
            {
              "title": "Degree Name",
              "subtitle": "Institution Name",
              "date": "Graduation Year",
              "details": ["Optional academic detail or empty string"]
            }
          ],
          "projects": [
            {
              "title": "Project Name",
              "subtitle": "Technologies Used",
              "date": "Year or Duration",
              "details": ["Action-oriented bullet 1", "Action-oriented bullet 2"]
            }
          ],
          "certifications": ["Certification 1", "Certification 2"],
          "achievements": ["Achievement/Leadership 1", "Achievement/Leadership 2"]
        }
        """

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"RAW RESUME TEXT:\n{raw_text}"}
            ],
            temperature=0.1,
            max_tokens=6000
        )

        response_text = response.choices[0].message.content.strip()
        
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
    
def generate_pdf(data, filename="improved_resume.pdf"):
    doc = SimpleDocTemplate(filename, pagesize=letter, rightMargin=25, leftMargin=25, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='TitleStyle', alignment=TA_CENTER, fontSize=15, leading=18, spaceAfter=2, fontName='Helvetica-Bold')
    contact_style = ParagraphStyle(name='ContactStyle', alignment=TA_CENTER, fontSize=9, spaceAfter=8)
    section_heading = ParagraphStyle(name='SectionHeading', fontSize=10, spaceBefore=6, spaceAfter=2, textTransform='uppercase', fontName='Helvetica-Bold')
    item_title = ParagraphStyle(name='ItemTitle', fontSize=9, spaceBefore=3, spaceAfter=0, fontName='Helvetica-Bold')
    normal_text = ParagraphStyle(name='NormalText', fontSize=9, leading=10, spaceAfter=2)
    bullet_text = ParagraphStyle(name='BulletText', fontSize=9, leading=10, leftIndent=10, firstLineIndent=0, spaceAfter=1)

    elements = []

    elements.append(Paragraph(data.get('name', '').upper(), title_style))
    elements.append(Paragraph(data.get('contact', ''), contact_style))
    
    def add_line():
        elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.black, spaceBefore=0, spaceAfter=2))

    if data.get('summary'):
        elements.append(Paragraph("<b>PROFESSIONAL SUMMARY</b>", section_heading))
        add_line()
        elements.append(Paragraph(data['summary'], normal_text))

    # UPDATED SKILLS DRAWING
    if data.get('skills'):
        elements.append(Paragraph("<b>SKILLS</b>", section_heading))
        add_line()
        for skill_group in data['skills']:
            cat = skill_group.get('category', '')
            items = skill_group.get('items', '')
            if cat and items:
                elements.append(Paragraph(f"<b>{cat}:</b> {items}", normal_text))
            elif items:
                elements.append(Paragraph(f"• {items}", normal_text))

    for section_title, key in [("EXPERIENCE", "experience"), ("PROJECTS", "projects")]:
        items = data.get(key, [])
        if items:
            elements.append(Paragraph(f"<b>{section_title}</b>", section_heading))
            add_line()
            for item in items:
                header = f"<b>{item.get('title', '')}</b>"
                if item.get('date'):
                    header += f" <font color='grey'>| {item.get('date', '')}</font>"
                elements.append(Paragraph(header, item_title))
                
                if item.get('subtitle'):
                    elements.append(Paragraph(f"<i>{item['subtitle']}</i>", normal_text))
                
                for detail in item.get('details', []):
                    if detail.strip():
                        elements.append(Paragraph(f"• {detail}", bullet_text))

    if data.get('education'):
        elements.append(Paragraph("<b>EDUCATION</b>", section_heading))
        add_line()
        for edu in data['education']:
            elements.append(Paragraph(f"<b>{edu.get('title', '')}</b> | {edu.get('date', '')}", item_title))
            elements.append(Paragraph(edu.get('subtitle', ''), normal_text))

    for section_title, key in [("CERTIFICATIONS", "certifications"), ("ACHIEVEMENTS & LEADERSHIP", "achievements")]:
        items = data.get(key, [])
        if items:
            elements.append(Paragraph(f"<b>{section_title}</b>", section_heading))
            add_line()
            for item in items:
                if str(item).strip():
                    elements.append(Paragraph(f"• {item}", bullet_text))

    doc.build(elements)
    return filename

def generate_docx(data, filename="improved_resume.docx"):
    doc = docx.Document()
    
    name_p = doc.add_heading(data.get('name', 'Your Name'), level=1)
    name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact_p = doc.add_paragraph(data.get('contact', 'Email | Phone'))
    contact_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def add_section_header(title):
        doc.add_heading(title, level=2)
        
    if data.get('summary'):
        add_section_header("Professional Summary")
        doc.add_paragraph(data['summary'])
        
    # UPDATED SKILLS DRAWING
    if data.get('skills') and len(data['skills']) > 0:
        add_section_header("Skills")
        for skill_group in data['skills']:
            p = doc.add_paragraph()
            cat = skill_group.get('category', '')
            items = skill_group.get('items', '')
            
            p.paragraph_format.left_indent = docx.shared.Inches(0.25)
            if cat:
                p.add_run(f"{cat}: ").bold = True
            p.add_run(items)
        
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
                    p_sub = doc.add_paragraph()
                    p_sub.add_run(item.get('subtitle', '')).italic = True
                    
                if item.get('details') and isinstance(item['details'], list):
                    for detail in item['details']:
                        if str(detail).strip():
                            doc.add_paragraph(f"•  {str(detail)}")
                            
    add_list_section("Experience", "experience")
    add_list_section("Education", "education")
    add_list_section("Projects", "projects")
    
    for section_title, key in [("Certifications", "certifications"), ("Achievements & Leadership", "achievements")]:
        items = data.get(key, [])
        if items and len(items) > 0:
            add_section_header(section_title)
            for item in items:
                if str(item).strip():
                    doc.add_paragraph(f"•  {str(item)}")
                
    doc.save(filename)
    return filename

@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'resume' not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
            
        file = request.files['resume']
        file_path = "temp_" + file.filename
        file.save(file_path)

        text = ""
        if file.filename.lower().endswith(".pdf"):
            text = extract_pdf(file_path)
        elif file.filename.lower().endswith(".docx"):
            text = extract_docx(file_path)
        else:
            return jsonify({"error": "Unsupported format"}), 400

        structured_data = enhance_with_ai(text)
        
        if "error" in structured_data:
             return jsonify(structured_data), 500

        pdf_file = generate_pdf(structured_data)
        docx_file = generate_docx(structured_data)
        
        if os.path.exists(file_path):
            os.remove(file_path)

        return jsonify({
            "resumeData": structured_data,
            "downloadUrl": f"http://127.0.0.1:5000/download/{pdf_file}"
        })

    except Exception as e:
        print("ERROR:", str(e))
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    if not os.path.exists(filename):
        return "File not found! Please click 'Upload & Enhance' to generate this document again.", 404
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)