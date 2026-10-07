import io
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# 16:9 Aspect Ratio Landscape PDF (960x540 pt)
SLIDE_WIDTH = 960
SLIDE_HEIGHT = 540

def draw_cover_background(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor('#0F172A'))
    canvas.rect(0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, fill=True, stroke=False)
    canvas.restoreState()

def draw_slide_background(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(colors.HexColor('#2563EB'))
    canvas.rect(40, SLIDE_HEIGHT - 70, 10, 40, fill=True, stroke=False)
    canvas.restoreState()

def create_pdf(slide_data: dict) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=(SLIDE_WIDTH, SLIDE_HEIGHT),
        rightMargin=60,
        leftMargin=60,
        topMargin=50,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    cover_title_style = ParagraphStyle(
        'CoverTitle',
        parent=styles['Heading1'],
        fontSize=36,
        leading=44,
        textColor=colors.white,
        alignment=1
    )

    slide_title_style = ParagraphStyle(
        'SlideTitle',
        parent=styles['Heading2'],
        fontSize=24,
        leading=30,
        textColor=colors.HexColor('#0F172A'),
        spaceAfter=20
    )

    body_style = ParagraphStyle(
        'SlideBody',
        parent=styles['BodyText'],
        fontSize=16,
        leading=24,
        textColor=colors.HexColor('#334155'),
        spaceAfter=12
    )

    story = []

    # Cover Slide
    story.append(Spacer(1, 160))
    story.append(Paragraph(slide_data.get("title", "Prezentatsiya"), cover_title_style))
    story.append(PageBreak())

    # Content Slides
    slides = slide_data.get("slides", [])
    for idx, slide in enumerate(slides):
        story.append(Paragraph(slide.get("title", ""), slide_title_style))
        story.append(Spacer(1, 15))
        
        for point in slide.get("bullet_points", []):
            story.append(Paragraph(f"• {point}", body_style))

        if idx < len(slides) - 1:
            story.append(PageBreak())

    def on_page(canvas, doc):
        if doc.page == 1:
            draw_cover_background(canvas, doc)
        else:
            draw_slide_background(canvas, doc)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    buffer.seek(0)
    return buffer
