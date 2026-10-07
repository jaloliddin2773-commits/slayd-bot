import json
import logging
import google.generativeai as genai
from config import GEMINI_API_KEY, logger

genai.configure(api_key=GEMINI_API_KEY)

FALLBACK_MODELS = [
    "gemini-1.5-flash",
    "gemini-1.5-pro",
    "gemini-1.0-pro"
]

SYSTEM_PROMPT = """
Siz professional prezentatsiya generatorisiz. 
Foydalanuvchi bergan mavzu bo'yicha slaydlar tuzilmasini strictly JSON formatida qaytarishingiz shart.

JSON Formati:
{
  "title": "Prezentatsiya Sarlavhasi",
  "slides": [
    {
      "slide_number": 1,
      "title": "Slayd Sarlavhasi",
      "bullet_points": [
        "Birinchi muhim nuqta",
        "Ikkinchi muhim nuqta",
        "Uchinchi muhim nuqta"
      ],
      "speaker_notes": "Taqdimotchi uchun izoh"
    }
  ]
}

Qoidalar:
1. Matn foydalanuvchi yozgan tilda bo'lsin (O'zbek, Rus, Ingliz va h.k.).
2. Slaydlar soni: 5-8 ta.
3. FAQAT toza JSON qaytaring. ```json va boshqa ortiqcha belgilarni ishlatmang.
"""

async def generate_presentation_structure(prompt_text: str) -> dict:
    last_error = None
    for model_name in FALLBACK_MODELS:
        try:
            logger.info(f"Gemini so'rovi yuborilmoqda, model: {model_name}")
            model = genai.GenerativeModel(
                model_name=model_name,
                generation_config={"response_mime_type": "application/json"}
            )
            
            response = model.generate_content(f"{SYSTEM_PROMPT}\n\nMavzu: {prompt_text}")
            clean_text = response.text.strip()
            
            if clean_text.startswith("```json"):
                clean_text = clean_text[7:]
            if clean_text.startswith("```"):
                clean_text = clean_text[3:]
            if clean_text.endswith("```"):
                clean_text = clean_text[:-3]
                
            data = json.loads(clean_text.strip())
            return data
        except Exception as e:
            logger.warning(f"Model {model_name} ishlamadi: {e}. Keyingi fallback modelga o'tilmoqda...")
            last_error = e

    raise RuntimeError(f"Barcha Gemini modellarida xatolik yuz berdi! So'nggi xatolik: {last_error}")
