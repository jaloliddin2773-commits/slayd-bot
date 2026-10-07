"""Slayd yasash: Gemini'dan matn olish va .pptx faylga aylantirish."""
import json
import os
import re

import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")

DARK = RGBColor(0x1F, 0x2A, 0x44)      # to'q ko'k
ORANGE = RGBColor(0xF2, 0x7A, 0x1A)    # to'q sariq
LIGHT = RGBColor(0xF7, 0xF5, 0xF0)     # och fon
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

PROMPT = """Sen taqdimot (slayd) tayyorlovchi yordamchisan.
Foydalanuvchi so'rovi: "{request}"

Qoidalar:
- Til: foydalanuvchi qaysi tilda yozgan bo'lsa, shu tilda (odatda o'zbek tilida) yoz.
- Slaydlar soni: so'rovda raqam bo'lsa shuncha, bo'lmasa 7. Eng ko'pi 12.
- Birinchi slayd - sarlavha slaydi (bullets bo'sh bo'ladi, subtitle bo'ladi).
- Qolgan slaydlarda 3-5 ta qisqa, mazmunli band (har biri 15 so'zdan oshmasin).
- Oxirgi slayd: xulosa.
- Faqat JSON qaytar, boshqa hech narsa yozma.

JSON shakli:
{{"title": "Taqdimot nomi", "slides": [
  {{"title": "Slayd sarlavhasi", "subtitle": "faqat birinchi slaydda", "bullets": ["band 1", "band 2"]}}
]}}"""


def ask_gemini(request_text: str) -> dict:
    """Gemini'dan slaydlar tuzilmasini JSON ko'rinishida oladi."""
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent"
    )
    body = {
        "contents": [{"parts": [{"text": PROMPT.format(request=request_text)}]}],
        "generationConfig": {"responseMimeType": "application/json"},
    }
    resp = requests.post(
        url, params={"key": GEMINI_API_KEY}, json=body, timeout=90
    )
    resp.raise_for_status()
    text = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
    text = re.sub(r"^```(?:json)?|```$", "", text.strip()).strip()
    data = json.loads(text)
    if not data.get("slides"):
        raise ValueError("Gemini slaydlarni qaytarmadi")
    data["slides"] = data["slides"][:12]
    return data


def _fill_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_text(slide, x, y, w, h, text, size, color, bold=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    return box


def build_pptx(data: dict, path: str) -> str:
    """Tuzilmadan chiroyli .pptx fayl yasaydi."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    for i, s in enumerate(data["slides"]):
        slide = prs.slides.add_slide(blank)

        if i == 0:  # sarlavha slaydi
            _fill_bg(slide, DARK)
            bar = slide.shapes.add_shape(1, Inches(0.8), Inches(3.0), Inches(1.5), Inches(0.12))
            bar.fill.solid()
            bar.fill.fore_color.rgb = ORANGE
            bar.line.fill.background()
            _add_text(slide, 0.8, 1.3, 11.5, 1.7, data.get("title") or s["title"], 44, WHITE, True)
            sub = s.get("subtitle") or s.get("title", "")
            _add_text(slide, 0.8, 3.4, 11.5, 1.5, sub, 24, RGBColor(0xD0, 0xD5, 0xE0))
            continue

        _fill_bg(slide, LIGHT)
        side = slide.shapes.add_shape(1, 0, 0, Inches(0.35), prs.slide_height)
        side.fill.solid()
        side.fill.fore_color.rgb = ORANGE
        side.line.fill.background()

        _add_text(slide, 0.9, 0.5, 11.5, 1.2, s["title"], 34, DARK, True)

        box = slide.shapes.add_textbox(Inches(0.9), Inches(1.9), Inches(11.5), Inches(4.8))
        tf = box.text_frame
        tf.word_wrap = True
        for j, b in enumerate(s.get("bullets", [])):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.text = "•  " + b
            p.font.size = Pt(24)
            p.font.color.rgb = DARK
            p.space_after = Pt(14)

        _add_text(slide, 12.0, 6.9, 1.0, 0.4, str(i + 1), 14, ORANGE, True)

    prs.save(path)
    return path
