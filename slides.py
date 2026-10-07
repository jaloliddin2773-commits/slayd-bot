"""
AI Slayd Generator
Gemini -> JSON -> PowerPoint
"""

import json
import os
import re
import time

import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


# ============================================================
# GEMINI SETTINGS
# ============================================================

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()

# Render Environment'da GEMINI_MODEL bo'lsa, o'sha ishlaydi.
# Bo'lmasa 2.5 Flash ishlatiladi.
GEMINI_MODEL = os.environ.get(
    "GEMINI_MODEL",
    "gemini-2.5-flash"
).strip()


# ============================================================
# DESIGN COLORS
# ============================================================

DARK = RGBColor(0x1F, 0x2A, 0x44)
ORANGE = RGBColor(0xF2, 0x7A, 0x1A)
LIGHT = RGBColor(0xF7, 0xF5, 0xF0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0xD0, 0xD5, 0xE0)


# ============================================================
# AI PROMPT
# ============================================================

PROMPT = """
Sen professional PowerPoint prezentatsiya tayyorlovchi AI yordamchisan.

Foydalanuvchi so'rovi:
"{request}"

Vazifa:
Ushbu mavzu asosida professional prezentatsiya strukturasi yarat.

Qoidalar:

1. Foydalanuvchi qaysi tilda yozgan bo'lsa, shu tilda javob ber.
2. Agar foydalanuvchi slayd sonini ko'rsatgan bo'lsa, aynan shuncha slayd yarat.
3. Agar slayd soni ko'rsatilmagan bo'lsa, 7 ta slayd yarat.
4. Minimal 3 ta, maksimal 12 ta slayd bo'lsin.
5. Birinchi slayd — titul.
6. Birinchi slaydda bullets bo'lmasin.
7. Birinchi slaydda subtitle bo'lsin.
8. Oxirgi slayd — Xulosa.
9. Har bir oddiy slaydda 3-5 ta bullet bo'lsin.
10. Har bir bullet qisqa va tushunarli bo'lsin.
11. Bir bullet 15 so'zdan oshmasin.
12. Faktlarni imkon qadar aniq yoz.
13. Takroriy ma'lumotlardan foydalanma.
14. Maktab/o'quv prezentatsiyasi bo'lsa, tushunarli tilda yoz.
15. Professional va chiroyli prezentatsiya strukturasini yarat.

FAQAT JSON qaytar.

Hech qanday markdown yozma.
Hech qanday ```json ishlatma.
Hech qanday izoh yozma.

JSON format:

{
  "title": "Prezentatsiya nomi",
  "slides": [
    {
      "title": "Asosiy sarlavha",
      "subtitle": "Qisqa izoh",
      "bullets": []
    },
    {
      "title": "Slayd sarlavhasi",
      "subtitle": "",
      "bullets": [
        "Birinchi qisqa fikr",
        "Ikkinchi qisqa fikr",
        "Uchinchi qisqa fikr"
      ]
    }
  ]
}
"""


# ============================================================
# GEMINI API
# ============================================================

def ask_gemini(request_text: str) -> dict:

    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY topilmadi. "
            "Render Environment'da GEMINI_API_KEY qo'shing."
        )

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent"
    )

    prompt = PROMPT.format(
        request=request_text
    )

    body = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.7,
            "responseMimeType": "application/json"
        }
    }

    last_error = None

    # 4 marta urinib ko'ramiz
    for attempt in range(4):

        try:

            response = requests.post(
                url,
                params={
                    "key": GEMINI_API_KEY
                },
                json=body,
                timeout=180
            )

            # 503 / 429 bo'lsa qayta urinadi
            if response.status_code in (429, 500, 502, 503, 504):

                last_error = (
                    f"Gemini server xatosi: "
                    f"HTTP {response.status_code}"
                )

                if attempt < 3:
                    time.sleep(3 * (attempt + 1))
                    continue

                raise RuntimeError(last_error)

            response.raise_for_status()

            result = response.json()

            candidates = result.get("candidates", [])

            if not candidates:
                raise RuntimeError(
                    "Gemini javob qaytarmadi."
                )

            content = candidates[0].get("content", {})

            parts = content.get("parts", [])

            if not parts:
                raise RuntimeError(
                    "Gemini javobida matn topilmadi."
                )

            text = parts[0].get("text", "").strip()

            if not text:
                raise RuntimeError(
                    "Gemini bo'sh javob qaytardi."
                )

            # Markdown JSON bo'lsa tozalaymiz
            text = re.sub(
                r"^```json\s*",
                "",
                text,
                flags=re.IGNORECASE
            )

            text = re.sub(
                r"^```\s*",
                "",
                text
            )

            text = re.sub(
                r"\s*```$",
                "",
                text
            )

            text = text.strip()

            # JSON parse
            try:
                data = json.loads(text)

            except json.JSONDecodeError:

                # JSON ichidan { ... } qismini topishga urinamiz
                start = text.find("{")
                end = text.rfind("}")

                if start == -1 or end == -1:
                    raise RuntimeError(
                        "Gemini valid JSON qaytarmadi."
                    )

                data = json.loads(
                    text[start:end + 1]
                )

            # =================================================
            # VALIDATION
            # =================================================

            slides = data.get("slides")

            if not isinstance(slides, list):
                raise RuntimeError(
                    "Gemini 'slides' ro'yxatini qaytarmadi."
                )

            if len(slides) == 0:
                raise RuntimeError(
                    "Gemini slayd yaratmadi."
                )

            # Maximum 12
            data["slides"] = slides[:12]

            # Har bir slaydni tozalash
            for slide in data["slides"]:

                if not isinstance(slide, dict):
                    continue

                slide.setdefault(
                    "title",
                    "Slayd"
                )

                slide.setdefault(
                    "subtitle",
                    ""
                )

                slide.setdefault(
                    "bullets",
                    []
                )

                if not isinstance(
                    slide["bullets"],
                    list
                ):
                    slide["bullets"] = []

            return data

        except requests.RequestException as e:

            last_error = str(e)

            if attempt < 3:
                time.sleep(3 * (attempt + 1))
                continue

            raise RuntimeError(
                f"Gemini bilan ulanishda xatolik: {last_error}"
            )

        except Exception as e:

            # Biz yaratgan xatolarni qayta ko'taramiz
            raise e

    raise RuntimeError(
        last_error or "Gemini noma'lum xato berdi."
    )


# ============================================================
# POWERPOINT HELPERS
# ============================================================

def fill_background(slide, color):

    fill = slide.background.fill

    fill.solid()

    fill.fore_color.rgb = color


def add_text(
    slide,
    x,
    y,
    width,
    height,
    text,
    size,
    color,
    bold=False
):

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(width),
        Inches(height)
    )

    text_frame = box.text_frame

    text_frame.word_wrap = True

    paragraph = text_frame.paragraphs[0]

    paragraph.text = str(text)

    paragraph.font.size = Pt(size)

    paragraph.font.bold = bold

    paragraph.font.color.rgb = color

    return box


def add_number(slide, number):

    add_text(
        slide,
        12.0,
        6.85,
        0.7,
        0.3,
        str(number),
        14,
        ORANGE,
        True
    )


# ============================================================
# CREATE POWERPOINT
# ============================================================

def build_pptx(data: dict, path: str) -> str:

    presentation = Presentation()

    # 16:9
    presentation.slide_width = Inches(13.333)

    presentation.slide_height = Inches(7.5)

    blank_layout = presentation.slide_layouts[6]

    slides = data.get("slides", [])

    presentation_title = data.get(
        "title",
        "AI Presentation"
    )

    for index, slide_data in enumerate(slides):

        slide = presentation.slides.add_slide(
            blank_layout
        )

        title = slide_data.get(
            "title",
            "Slayd"
        )

        subtitle = slide_data.get(
            "subtitle",
            ""
        )

        bullets = slide_data.get(
            "bullets",
            []
        )

        # ====================================================
        # TITLE SLIDE
        # ====================================================

        if index == 0:

            fill_background(
                slide,
                DARK
            )

            # Orange line
            line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.8),
                Inches(3.0),
                Inches(1.6),
                Inches(0.12)
            )

            line.fill.solid()

            line.fill.fore_color.rgb = ORANGE

            line.line.fill.background()

            # Title
            add_text(
                slide,
                0.8,
                1.25,
                11.5,
                1.7,
                presentation_title or title,
                44,
                WHITE,
                True
            )

            # Subtitle
            add_text(
                slide,
                0.8,
                3.35,
                11.5,
                1.2,
                subtitle or title,
                23,
                GRAY,
                False
            )

            continue

        # ====================================================
        # NORMAL SLIDE
        # ====================================================

        fill_background(
            slide,
            LIGHT
        )

        # Orange side bar
        side_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0),
            Inches(0),
            Inches(0.32),
            presentation.slide_height
        )

        side_bar.fill.solid()

        side_bar.fill.fore_color.rgb = ORANGE

        side_bar.line.fill.background()

        # Title
        add_text(
            slide,
            0.85,
            0.45,
            11.2,
            1.0,
            title,
            34,
            DARK,
            True
        )

        # Bullet box
        box = slide.shapes.add_textbox(
            Inches(0.9),
            Inches(1.75),
            Inches(11.4),
            Inches(4.9)
        )

        text_frame = box.text_frame

        text_frame.word_wrap = True

        for bullet_index, bullet in enumerate(
            bullets
        ):

            paragraph = (
                text_frame.paragraphs[0]
                if bullet_index == 0
                else text_frame.add_paragraph()
            )

            paragraph.text = (
                "•  " + str(bullet)
            )

            paragraph.font.size = Pt(23)

            paragraph.font.color.rgb = DARK

            paragraph.space_after = Pt(15)

        # Slide number
        add_number(
            slide,
            index + 1
        )

    # ========================================================
    # SAVE
    # ========================================================

    presentation.save(path)

    return path
