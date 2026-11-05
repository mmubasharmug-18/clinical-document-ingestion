import random
from datetime import datetime
from pathlib import Path

from docx import Document
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

OUT = Path(__file__).parent / "samples"
OUT.mkdir(exist_ok=True)
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"

NOTE_LINES = [
    "PROGRESS NOTE",
    "Patient: Sample Patient    MRN: 000123",
    "Date: 12 March 2026",
    "",
    "Reason for visit: follow up for type 2 diabetes.",
    "Blood pressure 128/82, pulse 76, weight 81 kg.",
    "HbA1c is 7.4 percent, down from 8.1 percent.",
    "Plan: continue metformin 500 mg twice daily.",
    "Repeat bloods in 3 months. Review diet with nurse.",
]


def note_image(width=1240, height=700):
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(FONT, 34)
    y = 60
    for line in NOTE_LINES:
        draw.text((70, y), line, fill=(25, 25, 25), font=font)
        y += 62
    img = img.rotate(0.7, expand=True, fillcolor="white")
    pixels = img.load()
    random.seed(4)
    for _ in range(9000):
        x, y = random.randrange(img.width), random.randrange(img.height)
        pixels[x, y] = (200, 200, 200)
    return img.filter(ImageFilter.GaussianBlur(0.6))


def make_txt():
    text = """DISCHARGE SUMMARY

Patient: Sample Patient
Admitted: 2 March 2026
Discharged: 6 March 2026

DIAGNOSIS:
Community acquired pneumonia, right lower lobe.

Hospital course:
Patient came in with fever and cough for four days. Chest x-ray showed
right lower lobe consolidation. Started on IV antibiotics and switched to
oral on day 3. Oxygen was not needed after day 2.

MEDICATIONS ON DISCHARGE:
- Amoxicillin 500 mg three times daily for 5 days
- Paracetamol 1 g as needed
- Salbutamol inhaler as needed

FOLLOW UP:
1. GP review in one week
2. Repeat chest x-ray in 6 weeks
"""
    (OUT / "discharge_summary.txt").write_text(text, encoding="utf-8")


def make_html():
    html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Lab Report - Sample Patient</title>
<meta name="author" content="City Lab">
<meta name="date" content="2026-03-10">
<style>body { font-family: Arial; }</style>
<script>console.log("ignore me");</script>
</head>
<body>
<h1>Laboratory Report</h1>
<p>Collected on 10 March 2026. Fasting sample.</p>
<h2>Blood tests</h2>
<table>
<tr><th>Test</th><th>Result</th><th>Range</th></tr>
<tr><td>Hemoglobin</td><td>13.6 g/dL</td><td>12.0 - 15.5</td></tr>
<tr><td>Fasting glucose</td><td>7.9 mmol/L</td><td>3.9 - 5.5</td></tr>
<tr><td>Creatinine</td><td>82 umol/L</td><td>45 - 90</td></tr>
</table>
<h2>Comments</h2>
<ul>
<li>Glucose is above range.</li>
<li>Other results are normal.</li>
</ul>
<p>Reported by the duty pathologist.</p>
</body>
</html>
"""
    (OUT / "lab_report.html").write_text(html, encoding="utf-8")


def make_docx():
    doc = Document()
    doc.core_properties.author = "Dr. A. Example"
    doc.core_properties.title = "Referral Letter"
    doc.core_properties.created = datetime(2026, 3, 9, 10, 30)
    doc.core_properties.modified = datetime(2026, 3, 9, 10, 45)
    doc.core_properties.subject = "Cardiology referral"
    doc.sections[0].header.paragraphs[0].text = "Example Family Clinic"
    doc.add_heading("Referral Letter", level=1)
    doc.add_paragraph("Dear Cardiology Team,")
    doc.add_paragraph(
        "Please see this 58 year old patient with chest tightness on exertion "
        "for the last three weeks. Resting ECG was normal."
    )
    doc.add_heading("Current medication", level=2)
    for item in ("Aspirin 75 mg daily", "Atorvastatin 20 mg at night", "Ramipril 5 mg daily"):
        doc.add_paragraph(item, style="List Bullet")
    doc.add_heading("Recent results", level=2)
    table = doc.add_table(rows=3, cols=3)
    rows = [("Test", "Result", "Date"), ("LDL", "3.1 mmol/L", "01/03/2026"), ("BP", "142/90", "08/03/2026")]
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            table.cell(r, c).text = value
    doc.add_paragraph("Kind regards, Dr. A. Example")
    doc.save(OUT / "referral_letter.docx")


def make_pdf():
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(
        str(OUT / "practice_guideline.pdf"),
        pagesize=A4,
        title="Hypertension Practice Guideline",
        author="Example Health Network",
    )
    body = styles["BodyText"]
    parts = [
        Paragraph("Hypertension Practice Guideline", styles["Title"]),
        Paragraph("1. Diagnosis", styles["Heading2"]),
        Paragraph(
            "Hypertension is diagnosed when clinic blood pressure stays at or above "
            "140/90 mmHg on repeat visits. Home or ambulatory readings should be used "
            "to confirm the diagnosis whenever possible.",
            body,
        ),
        Spacer(1, 8),
        Paragraph("2. First line treatment", styles["Heading2"]),
        Paragraph(
            "Start with lifestyle advice for every patient. Reduce salt, keep a healthy "
            "weight, be active and limit alcohol. If drugs are needed, an ACE inhibitor "
            "or a calcium channel blocker is the usual first choice.",
            body,
        ),
        Spacer(1, 8),
        Paragraph("3. Follow up", styles["Heading2"]),
        Paragraph(
            "Review the patient every four weeks until blood pressure is controlled, "
            "then every six months. Check kidney function and potassium after starting "
            "or changing an ACE inhibitor.",
            body,
        ),
    ]
    doc.build(parts)


def make_scanned_pdf():
    img = note_image()
    img.save(OUT / "_tmp_scan.png")
    page_w, page_h = A4
    c = canvas.Canvas(str(OUT / "scanned_progress_note.pdf"), pagesize=A4)
    c.setTitle("Scanned progress note")
    ratio = img.height / img.width
    c.drawImage(ImageReader(img), 30, page_h - 40 - (page_w - 60) * ratio, width=page_w - 60, height=(page_w - 60) * ratio)
    c.save()
    (OUT / "_tmp_scan.png").unlink()


def make_png():
    note_image(1100, 640).save(OUT / "clinic_note_photo.png")


if __name__ == "__main__":
    make_txt()
    make_html()
    make_docx()
    make_pdf()
    make_scanned_pdf()
    make_png()
    print(sorted(p.name for p in OUT.iterdir()))
