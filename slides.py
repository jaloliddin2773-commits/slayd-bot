import json
import os
import re
from pathlib import Path

import requests

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    KeepTogether,
)


# =========================================================
# GEMINI CONFIG
# =========================================================

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

DEFAULT_MODELS = [
    "gemini-3.5-flash-lite",
    "gemini-3.6-flash",
    "gemini-3.7-flash",
    "gemini-3.8-flash",
]

GEMINI_MODELS = [
    x.strip()
    for x in os.getenv(
        "GEMINI_MODELS",
        ",".join(DEFAULT_MODELS),
    ).split(",")
    if x.strip()
]

GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "{model}:generateContent"
)


# =========================================================
# AI PROMPT
# =========================================================

PROMPT = r"""
You are a professional presentation strategist, educator and
PowerPoint content designer.

Create a complete presentation based on the user's request.

USER REQUEST:
__REQUEST__

IMPORTANT:

- Understand exactly what the user wants.
- If the user specifies the number of slides, follow it.
- If no number is specified, create 7 slides.
- If a language is specified, use that language.
- Otherwise use the language of the user's request.
- Make the presentation logical and professional.
- Avoid unnecessary text.
- Avoid repeating information.
- Do not invent statistics.
- Keep each bullet short.
- Maximum 5 bullets per slide.
- Make the content suitable for a real presentation.
- The first slide is created automatically by the application.
- Therefore "slides" contains only the content slides.

Use these layouts:

bullets
two_column
stats
quote

Return ONLY valid JSON.

Required JSON structure:

{
  "title": "Presentation title",
  "subtitle": "Short subtitle",
  "language": "uz",
  "slides": [
    {
      "title": "Slide title",
      "layout": "bullets",
      "bullets": [
        "Point one",
        "Point two",
        "Point three"
      ]
    }
  ]
}

For two_column:

{
  "title": "Slide title",
  "layout": "two_column",
  "left_title": "Left title",
  "left": [
    "Point",
    "Point"
  ],
  "right_title": "Right title",
  "right": [
    "Point",
    "Point"
  ]
}

For stats:

{
  "title": "Slide title",
  "layout": "stats",
  "stats": [
    {
      "value": "01",
      "label": "Label"
    },
    {
      "value": "02",
      "label": "Label"
    },
    {
      "value": "03",
      "label": "Label"
    }
  ]
}

For quote:

{
  "title": "Slide title",
  "layout": "quote",
  "quote": "Important statement",
  "source": "Author or source"
}

Return JSON only.
"""


# =========================================================
# JSON HELPERS
# =========================================================

def clean_json(text: str) -> str:
    text = text.strip()

    if text.startswith("```"):
        text = re.sub(
            r"^```(?:json)?\s*",
            "",
            text,
        )

        text = re.sub(
            r"\s*```$",
            "",
            text,
        )

    start = text.find("{")
    end = text.rfind("}")

    if start == -1 or end == -1:
        raise ValueError(
            "AI returned invalid JSON."
        )

    return text[start:end + 1]


def extract_response_text(data: dict) -> str:

    candidates = data.get(
        "candidates",
        [],
    )

    if not candidates:
        raise RuntimeError(
            "Gemini returned no candidates."
        )

    content = candidates[0].get(
        "content",
        {},
    )

    parts = content.get(
        "parts",
        [],
    )

    texts = []

    for part in parts:
        if (
            isinstance(part, dict)
            and "text" in part
        ):
            texts.append(
                part["text"]
            )

    result = "\n".join(texts).strip()

    if not result:
        raise RuntimeError(
            "Gemini returned an empty response."
        )

    return result


# =========================================================
# GEMINI
# =========================================================

def ask_gemini(request_text: str) -> dict:

    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY is missing."
        )

    prompt = PROMPT.replace(
        "__REQUEST__",
        request_text.strip(),
    )

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.7,
            "maxOutputTokens": 12000,
        },
    }

    errors = []

    for model in GEMINI_MODELS:

        url = GEMINI_URL.format(
            model=model
        )

        try:

            response = requests.post(
                url,
                headers={
                    "x-goog-api-key": GEMINI_API_KEY,
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=90,
            )

            # Temporary server/quota problems:
            # try another model.
            if response.status_code in (
                429,
                500,
                502,
                503,
                504,
            ):
                errors.append(
                    f"{model}: HTTP "
                    f"{response.status_code}"
                )
                continue

            response.raise_for_status()

            data = response.json()

            raw_text = extract_response_text(
                data
            )

            json_text = clean_json(
                raw_text
            )

            result = json.loads(
                json_text
            )

            if not isinstance(
                result,
                dict,
            ):
                raise ValueError(
                    "AI response is not an object."
                )

            if not isinstance(
                result.get("slides"),
                list,
            ):
                raise ValueError(
                    "AI did not return slides."
                )

            if not result["slides"]:
                raise ValueError(
                    "AI returned zero slides."
                )

            result["_model"] = model

            return result

        except requests.RequestException as exc:

            errors.append(
                f"{model}: {exc}"
            )

        except (
            ValueError,
            json.JSONDecodeError,
            RuntimeError,
        ) as exc:

            errors.append(
                f"{model}: {exc}"
            )

    raise RuntimeError(
        "AI xizmatlari vaqtincha ishlamadi.\n\n"
        + "\n".join(errors)
    )


# =========================================================
# POWERPOINT DESIGN
# =========================================================

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

BG = RGBColor(
    248,
    249,
    252,
)

DARK = RGBColor(
    22,
    27,
    34,
)

MUTED = RGBColor(
    92,
    99,
    112,
)

ACCENT = RGBColor(
    245,
    102,
    45,
)

WHITE = RGBColor(
    255,
    255,
    255,
)

LIGHT = RGBColor(
    232,
    235,
    240,
)


def add_background(slide):

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0,
        0,
        SLIDE_W,
        SLIDE_H,
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = BG
    shape.line.fill.background()


def add_top_accent(slide):

    shape = slide.shapes
