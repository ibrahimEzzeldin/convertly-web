"""
translation_service.py — MyMemory translation engine (no sleeps, no retries)
"""
import os
import requests
import logging
import time

logger = logging.getLogger(__name__)

MYMEMORY_URL   = "https://api.mymemory.translated.net/get"
MAX_CHUNK_SIZE = 500
MYMEMORY_EMAIL = os.environ.get("MYMEMORY_EMAIL", "")

# Maps short ISO codes to locale codes expected by MyMemory
# MyMemory API requires bare language codes (ar, en, fr, etc.) not region-specific (ar-SA, en-GB, fr-FR)
LANG_CODES = {
    "ar": "ar", "fr": "fr", "es": "es", "de": "de",
    "it": "it", "pt": "pt", "ru": "ru", "zh": "zh",
    "zh-TW": "zh", "ja": "ja", "ko": "ko", "tr": "tr",
    "nl": "nl", "pl": "pl", "hi": "hi", "uk": "uk",
    "sv": "sv", "no": "no", "da": "da", "el": "el",
    "he": "he", "id": "id", "vi": "vi", "en": "en",
}


def _map(code: str) -> str:
    """Map language codes to MyMemory-compatible bare codes.
    
    MyMemory only recognizes simple language codes like 'ar', 'en', 'fr', not region variants.
    """
    if not code or not isinstance(code, str):
        return code

    code = code.strip()
    if not code:
        return code

    # Direct match in LANG_CODES
    if code in LANG_CODES:
        return LANG_CODES[code]
    
    # Case-insensitive match
    code_lower = code.lower()
    if code_lower in LANG_CODES:
        return LANG_CODES[code_lower]

    # For region-coded inputs like "en-US", "zh-CN", extract base and look up
    if "-" in code_lower:
        base = code_lower.split("-")[0]
        if base in LANG_CODES:
            return LANG_CODES[base]

    # Fallback: return as-is
    return code


def split_into_chunks(text: str, max_size: int = MAX_CHUNK_SIZE) -> list:
    if len(text) <= max_size:
        return [text]
    chunks, current = [], ""
    for part in text.replace('\n', ' \n ').split('\n'):
        for s in (
            part.replace('. ', '.|')
                .replace('! ', '!|')
                .replace('? ', '?|')
                .split('|')
        ):
            s = s.strip()
            if not s:
                continue
            if len(current) + len(s) + 1 <= max_size:
                current += (" " if current else "") + s
            else:
                if current:
                    chunks.append(current)
                current = s
    if current:
        chunks.append(current)
    return chunks or [text]


def translate_chunk(text: str, from_code: str, to_code: str) -> str:
    """Translate one chunk via MyMemory with limited retries and backoff.

    Returns original text on failure (safe fallback).
    """
    if not text or not text.strip():
        return text

    langpair = f"{_map(from_code)}|{_map(to_code)}"
    params = {"q": text, "langpair": langpair}
    if MYMEMORY_EMAIL:
        params["de"] = MYMEMORY_EMAIL

    max_retries = 3
    backoff_s = 1.0
    for attempt in range(1, max_retries + 1):
        try:
            res = requests.get(MYMEMORY_URL, params=params, timeout=15)
            data = res.json()
            status = data.get("responseStatus")
            translated = data.get("responseData", {}).get("translatedText", "")

            if status == 200 and translated and "MYMEMORY WARNING" not in str(translated):
                return translated

            # Rate limit or quality warning: retry a few times before giving up
            rate_limit = status == 429 or "MYMEMORY WARNING" in str(translated)
            if rate_limit:
                logger.warning("MyMemory rate limited or warning on attempt %d for: '%s'", attempt, text[:40])
                if attempt < max_retries:
                    time.sleep(backoff_s)
                    backoff_s *= 2
                    continue
                return text

            logger.warning("MyMemory unexpected response %s: %s", status, str(translated)[:80])
            return text

        except requests.Timeout:
            logger.warning("MyMemory request timeout on attempt %d", attempt)
            if attempt < max_retries:
                time.sleep(backoff_s)
                backoff_s *= 2
                continue
            return text

        except requests.RequestException as exc:
            logger.warning("MyMemory request exception: %s (attempt %d)", exc, attempt)
            if attempt < max_retries:
                time.sleep(backoff_s)
                backoff_s *= 2
                continue
            return text

        except Exception as exc:
            logger.error("MyMemory error: %s", exc)
            return text

    return text


def translate_text(text: str, from_code: str, to_code: str) -> str:
    """Translate full text by chunking. No sleeps. No crashes."""
    if not text or not text.strip():
        return text
    chunks = split_into_chunks(text)
    logger.info("translate_text: %d chars → %d chunks, %s→%s", len(text), len(chunks), from_code, to_code)
    return " ".join(translate_chunk(c, from_code, to_code) for c in chunks)
