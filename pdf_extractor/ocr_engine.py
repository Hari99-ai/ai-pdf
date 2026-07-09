import base64
import os
import requests
from typing import Dict, Any, Optional
from pdf_extractor.logging_config import logger

try:
    import pytesseract
    pytesseract_available = True
except ImportError:
    pytesseract_available = False

def encode_image_to_base64(image_bytes: bytes) -> str:
    """Encode raw image bytes to a base64 string."""
    return base64.b64encode(image_bytes).decode("utf-8")

def extract_text_with_local_tesseract(image_bytes: bytes) -> Optional[str]:
    """Fallback OCR using local pytesseract (requires tesseract binary on path)."""
    if not pytesseract_available:
        return None
    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(image_bytes))
        logger.info("Performing local OCR via pytesseract...")
        text = pytesseract.image_to_string(img)
        return text.strip()
    except Exception as e:
        logger.warning(f"Local pytesseract OCR failed: {e}")
        return None

def perform_multimodal_ocr(
    image_bytes: bytes,
    api_key: str,
    model: str = "openai/gpt-4o-mini",
    site_url: str = "",
    app_name: str = ""
) -> str:
    """Perform high-fidelity OCR via OpenRouter Multimodal LLM (GPT-4o-mini)."""
    if not api_key:
        raise ValueError("OPENROUTER_API_KEY is missing. Cannot perform multimodal OCR.")

    base64_image = encode_image_to_base64(image_bytes)
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    if site_url:
        headers["HTTP-Referer"] = site_url
    if app_name:
        headers["X-OpenRouter-Title"] = app_name

    prompt = (
        "You are a professional document OCR agent. Extract all text and tables from this scanned document page. "
        "Preserve every number, decimal, currency symbol (e.g. $, €), date range, and percentage exactly as written. "
        "Do not round, modify, translate, or calculate anything. "
        "Format any tables you see as clear, structured text tables (using Markdown or CSV format)."
    )

    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}"
                        }
                    }
                ]
            }
        ],
        "temperature": 0.0,
        "max_tokens": 4000
    }

    logger.info(f"Invoking OpenRouter multimodal OCR using model {model}...")
    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=120
        )
        response.raise_for_status()
        data = response.json()
        ocr_text = data["choices"][0]["message"]["content"]
        logger.info("Multimodal OCR completed successfully.")
        return ocr_text.strip()
    except Exception as e:
        logger.error(f"Multimodal OCR request failed: {e}")
        raise RuntimeError(f"Multimodal OCR failed: {str(e)}")
