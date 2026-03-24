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
LANG_CODES = {
    "ar": "ar-SA", "fr": "fr-FR", "es": "es-ES", "de": "de-DE",
    "it": "it-IT", "pt": "pt-BR", "ru": "ru-RU", "zh": "zh-CN",
    "zh-TW": "zh-TW", "ja": "ja-JP", "ko": "ko-KR", "tr": "tr-TR",
    "nl": "nl-NL", "pl": "pl-PL", "hi": "hi-IN", "uk": "uk-UA",
    "sv": "sv-SE", "no": "no-NO", "da": "da-DK", "el": "el-GR",
    "he": "he-IL", "id": "id-ID", "vi": "vi-VN", "en": "en-GB",
}


def _map(code: str) -> str:
    if not code or not isinstance(code, str):
        return code

    code = code.strip()
    if not code:
        return code

    # direct match (case-sensitive and case-insensitive)
    if code in LANG_CODES:
        return LANG_CODES[code]
    if code.lower() in LANG_CODES:
        return LANG_CODES[code.lower()]

    normalized = code.lower()
    aliases = {
        "en-us": "en-GB", "en-gb": "en-GB",
        "pt-br": "pt-BR", "pt-pt": "pt-BR",
        "zh-cn": "zh-CN", "zh-tw": "zh-TW",
        "ar-sa": "ar-SA", "ru-ru": "ru-RU",
        "ja-jp": "ja-JP", "ko-kr": "ko-KR",
        "tr-tr": "tr-TR", "nl-nl": "nl-NL",
        "pl-pl": "pl-PL", "hi-in": "hi-IN",
        "uk-ua": "uk-UA", "sv-se": "sv-SE",
        "no-no": "no-NO", "da-dk": "da-DK",
        "el-gr": "el-GR", "he-il": "he-IL",
        "id-id": "id-ID", "vi-vn": "vi-VN",
    }

    if normalized in aliases:
        return aliases[normalized]

    # Fall back to basic version for unknown region-coded languages
    if "-" in code:
        base, region = code.split("-", 1)
        if base.lower() in LANG_CODES:
            return LANG_CODES[base.lower()]

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
