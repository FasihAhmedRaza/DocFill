"""
extractor.py — PDF → images (PyMuPDF) and Gemini Vision extraction.

Uses the modern Google GenAI SDK:
    from google import genai
    client = genai.Client(api_key=...)
    client.models.generate_content(model=..., contents=[prompt, Part.from_bytes(...)])

PyMuPDF (fitz) is pure-Python with zero system dependencies, so this runs
unchanged on Streamlit Community Cloud. No Tesseract / poppler / pdf2image.
"""

import io
import json
import re
import time

import fitz  # PyMuPDF
from PIL import Image

from validator import bank_arabic_from_english, canonical_emirate, normalize_iban

# Primary model + automatic fallbacks if one is overloaded / rate-limited.
MODELS_FALLBACK = [
    "gemini-2.5-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.0-flash",
]
GEMINI_MODEL = MODELS_FALLBACK[0]          # shown in the UI chip

EXTRACTION_PROMPT = """You are a UAE bank document data extraction specialist.
Examine ALL pages of this document (IBAN letter + Emirates ID card).
Return ONLY a valid JSON object — no markdown, no backticks, no extra text.

{
  "english_name":    "Full name in English exactly as printed",
  "arabic_name":     "Full name in Arabic (null if not visible)",
  "iban":            "IBAN — no spaces (e.g. AE190030013056648920001)",
  "account_number":  "Account number digits only",
  "swift":           "SWIFT/BIC code",
  "sort_code":       "Bank sort code if visible (else null)",
  "branch":          "Bank branch name if visible (else null)",
  "currency":        "AED / USD / EUR",
  "bank_name_en":    "Bank name in English",
  "bank_name_ar":    "Bank name in Arabic",
  "nationality":     "Nationality in English",
  "emirates_id":     "Emirates ID in format 784-XXXX-XXXXXXX-X",
  "date_of_birth":   "DD/MM/YYYY",
  "id_expiry":       "DD/MM/YYYY",
  "emirate":         "Emirate name if visible"
}

Rules:
- Use null for any field not clearly visible. Do NOT hallucinate.
- IBAN: remove all spaces.
- Return ONLY the JSON object, nothing else.
"""

FIELDS = [
    "english_name", "arabic_name", "iban", "account_number", "swift",
    "sort_code", "branch", "currency", "bank_name_en", "bank_name_ar",
    "nationality", "emirates_id", "date_of_birth", "id_expiry", "emirate",
]

_RETRYABLE = ("503", "429", "UNAVAILABLE", "RESOURCE_EXHAUSTED")


# ── PDF → images ──────────────────────────────────────────────────────────
def pdf_to_images(pdf_bytes: bytes, dpi: int = 200, max_pages: int = 8) -> list:
    """Render PDF pages to PIL Images using PyMuPDF."""
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for page_index in range(min(doc.page_count, max_pages)):
            pix = doc.load_page(page_index).get_pixmap(dpi=dpi)
            images.append(Image.open(io.BytesIO(pix.tobytes("png"))))
    finally:
        doc.close()
    return images


def _pil_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# ── JSON parsing ──────────────────────────────────────────────────────────
def _parse_json(raw: str) -> dict:
    """Strip code fences and parse the first JSON object Gemini returns."""
    raw = (raw or "").strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise


# ── Gemini call with model fallback + retry/backoff ───────────────────────
def extract_with_gemini(images: list, api_key: str) -> dict:
    """Send all page images to Gemini Vision → normalized field dict.

    Tries each model in MODELS_FALLBACK; retries transient 429/503 errors with
    exponential backoff before moving on. Raises if every model is exhausted.
    """
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)

    contents = [EXTRACTION_PROMPT]
    for img in images:
        contents.append(types.Part.from_bytes(data=_pil_to_png_bytes(img), mime_type="image/png"))

    last_exc = None
    for model_idx, model_name in enumerate(MODELS_FALLBACK):
        for attempt in range(3):
            try:
                response = client.models.generate_content(model=model_name, contents=contents)
                return _normalize(_parse_json(response.text))
            except Exception as exc:
                last_exc = exc
                err = str(exc)
                retryable = any(k in err for k in _RETRYABLE) or "overloaded" in err.lower()
                if not retryable:
                    raise
                if attempt < 2:
                    time.sleep((attempt + 1) * 3)   # 3s, 6s
                    continue
                break   # exhausted this model → try the next fallback
    raise RuntimeError(f"All Gemini models exhausted. Last error: {last_exc}")


# ── Normalization ─────────────────────────────────────────────────────────
def _normalize(data: dict) -> dict:
    """Coerce nulls to empty strings, clean IBAN, backfill Arabic bank name."""
    data = data or {}
    rec = {f: data.get(f) for f in FIELDS}
    rec = {k: ("" if v is None else str(v).strip()) for k, v in rec.items()}

    if rec.get("iban"):
        rec["iban"] = normalize_iban(rec["iban"])
    if not rec.get("currency"):
        rec["currency"] = "AED"
    if rec.get("emirate"):
        rec["emirate"] = canonical_emirate(rec["emirate"])
    if not rec.get("bank_name_ar") and rec.get("bank_name_en"):
        rec["bank_name_ar"] = bank_arabic_from_english(rec["bank_name_en"])
    return rec


def blank_record() -> dict:
    """An empty record (used when extraction fails — user fills manually)."""
    rec = {f: "" for f in FIELDS}
    rec["currency"] = "AED"
    return rec
