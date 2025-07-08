"""
program_review_report_demo.py

This script demonstrates how to generate automated program review reports by department prefix
(e.g., ELC, SOC), including student enrollments and course sections by semester and modality.

The output is a PDF report formatted for academic use. This version uses synthetic structure only.

Tools: pandas, reportlab
"""

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph, Spacer
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ---- SETTINGS ----
input_file = "demo_course_data.xlsx"      # Replace with synthetic demo data
course_prefix = "ELC"                     # Department/course prefix to analyze
output_pdf = f"ProgramReview_{course_prefix}_Modality.pdf"

# ---- LOAD DATA ----
df = pd.read_excel(input_file)
df['TERM'] = df['TERM'].astype(str)

# Custom sort for term (e.g., Spring < Fall)
def sort_term(term):
    year = int(term[:4])
    semester = term[4:]
    semester_order = {'SP': 1, 'FA': 2}
    return year, semester_order.get(semester, 0)

df['Year'], df['SemesterOrder'] = zip(*df['TERM'].apply(sort_term))
df = df.sort_values(by=['Year', 'SemesterOrder'])

# ---- FILTER BY PREFIX ----
df = df[df['CRS'].str.startswith(course_prefix)]
courses = sorted(df['CRS'].unique())

# ---- PREPARE PDF ELEMENTS ----
doc = SimpleDocTemplate(output_pdf, pagesize=letter)
elements = []
styles = getSampleStyleSheet()

title_style = ParagraphStyle(
    'Title', parent=styles['Heading2'], alignment=1, fontSize=18, leading=34
)
footer_style = ParagraphStyle(
    'Footer', parent=styles['Heading3'], alignment=1, fontSize=14
)
note_style = ParagraphStyle(name='TinyNote', fontSize=9, leading=10)

# Cover Page
elements.append(Spacer(1, 100))
elements.append(Paragraph(
    f"<b>Fall and Spring Enrollments and Sections by<br/>"
    f"Modality for {course_prefix} Courses,<br/>"
    f"Spring 2020 - Spring 2024</b>", title_style))
elements.append(Spacer(1, 250))
elements.append(Paragraph("<b>Provided by<br/>Center for Institutional Effectiveness<br/>December 2024</b>", footer_style))
elements.append(PageBreak())

# ---- LOOP THROUGH COURSES ----
for course in courses:
    df_course = df[df['CRS'] == course]
    terms = sorted(df_course['TERM'].unique(), key=lambda t: sort_term(t))

    summary = []
    total_enrollment = 0
    total_sections = 0

    for term in terms:
        df_term = df_course[df_course['TERM'] == term]
        counts = {
            'Face-to-Face': (df_term['MODALITY'] == 'Face-to-Face').sum(),
            'Online': (df_term['MODALITY'] == 'Online').sum(),
            'Hybrid': (df_term['MODALITY'] == 'Hybrid').sum()
        }
        sections = {
            'Face-to-Face': df_term[df_term['MODALITY'] == 'Face-to-Face']['CRS_SECT'].nunique(),
            'Online': df_term[df_term['MODALITY'] == 'Online']['CRS_SECT'].nunique(),
            'Hybrid': df_term[df_term['MODALITY'] == 'Hybrid']['CRS_SECT'].nunique()
        }

        subtotal = sum(counts.values())
        subsect = sum(sections.values())
        total_enrollment += subtotal
        total_sections += subsect

        summary.extend([
            [term, 'Face-to-Face', counts['Face-to-Face'], sections['Face-to-Face']],
            ['', 'Online', counts['Online'], sections['Online']],
            ['', 'Hybrid', counts['Hybrid'], sections['Hybrid']],
            ['', 'Subtotal', subtotal, subsect]
        ])

    summary.append([
        '', f'{course} Grand Total:', total_enrollment, total_sections
    ])

    # Add table to PDF
    elements.append(Paragraph(f"<b>Course: {course}</b>", styles['Heading2']))
    data = [["Term", "Modality", "Enrollments", "Sections"]] + summary
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.purple),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.purple),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
    ]))
    elements.append(table)

    # Add data notes
    note_text = (
        "<br/><b>Source:</b> Institutional enrollment records (end-of-term)<br/><br/>"
        "<b>Note:</b> Enrollment counts reflect distinct seats per section (not unduplicated students).<br/>"
        "Modality is determined by section codes (e.g., WB = Online, HY = Hybrid).<br/>"
        "Time-of-day flags are based on section suffixes (below 599 = Day, above = Evening)."
    )
    elements.append(Paragraph(note_text, note_style))
    elements.append(PageBreak())

# ---- HEADER/FOOTER LOGIC ----
def no_header_footer(canvas, doc): pass

def header_footer(canvas, doc):
    canvas.setFont("Helvetica", 10)
    header = f"{course_prefix} Courses by Term and Modality (Spring 2020–Spring 2024)"
    canvas.drawCentredString(letter[0] / 2, letter[1] - 20, header)
    canvas.setFont("Helvetica", 8)
    canvas.drawString(30, 30, f"Data Packet: {course_prefix}")
    canvas.drawCentredString(letter[0] / 2, 30, f"Page {canvas.getPageNumber()}")
    canvas.drawString(letter[0] - 100, 30, "CIE")

# ---- BUILD PDF ----
doc.build(elements, onFirstPage=no_header_footer, onLaterPages=header_footer)
print(f" PDF generated: {output_pdf}")
