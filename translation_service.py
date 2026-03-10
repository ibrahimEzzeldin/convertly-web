"""
translation_service.py — MyMemory translation engine (no sleeps, no retries)
"""
import os
import requests
import logging

logger = logging.getLogger(__name__)

MYMEMORY_URL   = "https://api.mymemory.translated.net/get"
MAX_CHUNK_SIZE = 500
MYMEMORY_EMAIL = os.environ.get("MYMEMORY_EMAIL", "")

# Maps short ISO codes to locale codes expected by MyMemory
LANG_CODES = {
    "ar": "ar-SA", "fr": "fr-FR", "es": "es-ES", "de": "de-DE",
    "it": "it-IT", "pt": "pt-BR", "ru": "ru-RU", "zh": "zh-CN",
    "zh-TW": "zh-TW", "ja": "ja-JP", "ko": "ko-KR", "tr": "tr-TR",
    "nl": "nl-NL", "pl": "pl-PL", "hi": "hi-IN", "uk": "uk-UA",
    "sv": "sv-SE", "no": "no-NO", "da": "da-DK", "el": "el-GR",
    "he": "he-IL", "id": "id-ID", "vi": "vi-VN", "en": "en-GB",
}


def _map(code: str) -> str:
    return LANG_CODES.get(code, code)


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
    """Translate one chunk via MyMemory. No retries. No sleep. Returns original on failure."""
    if not text or not text.strip():
        return text
    params = {
        "q": text,
        "langpair": f"{_map(from_code)}|{_map(to_code)}",
    }
    if MYMEMORY_EMAIL:
        params["de"] = MYMEMORY_EMAIL
    try:
        res = requests.get(MYMEMORY_URL, params=params, timeout=8)
        data = res.json()
        status = data.get("responseStatus")
        translated = data.get("responseData", {}).get("translatedText", "")
        if status == 200 and translated and "MYMEMORY WARNING" not in str(translated):
            return translated
        if status == 429 or "MYMEMORY WARNING" in str(translated):
            logger.warning("MyMemory rate limited for: '%s'", text[:40])
            return text
        logger.warning("MyMemory unexpected response %s: %s", status, str(translated)[:80])
        return text
    except requests.Timeout:
        logger.warning("MyMemory request timed out")
        return text
    except Exception as exc:
        logger.error("MyMemory error: %s", exc)
        return text


def translate_text(text: str, from_code: str, to_code: str) -> str:
    """Translate full text by chunking. No sleeps. No crashes."""
    if not text or not text.strip():
        return text
    chunks = split_into_chunks(text)
    logger.info("translate_text: %d chars → %d chunks, %s→%s", len(text), len(chunks), from_code, to_code)
    return " ".join(translate_chunk(c, from_code, to_code) for c in chunks)
