"""
translation_service.py — DeepL (primary) + MyMemory (fallback) translation engine
"""
import os
import requests
import logging
import time
from typing import Optional

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


_deepl_absent_logged = False


def _translate_deepl(text: str, source_lang: str, target_lang: str) -> Optional[str]:
    """Translate using DeepL free API. Returns None if unavailable."""
    api_key = os.getenv("DEEPL_API_KEY", "")
    if not api_key:
        # Log once at startup-ish rather than per chunk. Without this it is
        # invisible that DeepL is configured in code but never actually used,
        # and every translation silently rides MyMemory's small free tier.
        global _deepl_absent_logged
        if not _deepl_absent_logged:
            _deepl_absent_logged = True
            logger.warning(
                "DEEPL_API_KEY is not set — all translation will use MyMemory "
                "(free tier: ~5k chars/day anonymous, 50k with MYMEMORY_EMAIL)."
            )
        return None

    # DeepL uses uppercase 2-letter codes
    src = source_lang.upper().split("-")[0] if source_lang.lower() != "auto" else None
    tgt = target_lang.upper().split("-")[0]

    payload = {"text": [text], "target_lang": tgt}
    if src:
        payload["source_lang"] = src

    try:
        resp = requests.post(
            "https://api-free.deepl.com/v2/translate",
            headers={"Authorization": f"DeepL-Auth-Key {api_key}", "Content-Type": "application/json"},
            json=payload,
            timeout=15,
        )
        if resp.status_code == 200:
            return resp.json()["translations"][0]["text"]
        logger.warning("DeepL returned %s: %s", resp.status_code, resp.text[:100])
        return None
    except Exception as e:
        logger.warning("DeepL error: %s", e)
        return None


def _translate_mymemory(text: str, from_code: str, to_code: str) -> Optional[str]:
    """Translate one chunk via MyMemory with limited retries and backoff.

    Returns None on failure so the caller can decide on the fallback.
    """
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
                return None

            logger.warning("MyMemory unexpected response %s: %s", status, str(translated)[:80])
            return None

        except requests.Timeout:
            logger.warning("MyMemory request timeout on attempt %d", attempt)
            if attempt < max_retries:
                time.sleep(backoff_s)
                backoff_s *= 2
                continue
            return None

        except requests.RequestException as exc:
            logger.warning("MyMemory request exception: %s (attempt %d)", exc, attempt)
            if attempt < max_retries:
                time.sleep(backoff_s)
                backoff_s *= 2
                continue
            return None

        except Exception as exc:
            logger.error("MyMemory error: %s", exc)
            return None

    return None


def _translate_chunk_detailed(text: str, from_code: str, to_code: str):
    """Translate one chunk. Returns (text, provider).

    provider is "deepl", "mymemory", or None when every provider failed — in
    which case the ORIGINAL text is returned so the caller can still assemble a
    document, but it now knows the chunk was not actually translated.
    """
    if not text or not text.strip():
        return text, "skip"

    # Try DeepL first (higher quality, more reliable)
    result = _translate_deepl(text, from_code, to_code)
    if result is not None:
        return result, "deepl"

    # Fall back to MyMemory
    result = _translate_mymemory(text, from_code, to_code)
    if result is not None:
        return result, "mymemory"

    # Both providers failed — return original text, flagged as a failure.
    logger.warning("TRANSLATE_SILENT_FALLBACK chunk not translated: '%s'", text[:40])
    return text, None


def translate_chunk(text: str, from_code: str, to_code: str) -> str:
    """Translate one chunk, trying DeepL first then falling back to MyMemory.

    Returns original text if all providers fail (safe fallback).
    """
    return _translate_chunk_detailed(text, from_code, to_code)[0]


def translate_text_detailed(text: str, from_code: str, to_code: str):
    """Translate full text by chunking. Returns (translated_text, stats).

    stats = {"total": int, "failed": int, "providers": [str]}

    Callers should inspect stats rather than trusting the text: when every
    provider fails this function still returns readable output, but it is the
    ORIGINAL untranslated text. Silently handing that back to a user — who has
    just spent a conversion credit — reads as a broken product, so routes turn
    a total failure into a real error.
    """
    if not text or not text.strip():
        return text, {"total": 0, "failed": 0, "providers": []}

    chunks = split_into_chunks(text)
    out, providers, failed = [], set(), 0
    for c in chunks:
        translated, provider = _translate_chunk_detailed(c, from_code, to_code)
        out.append(translated)
        if provider is None:
            failed += 1
        elif provider != "skip":
            providers.add(provider)

    stats = {"total": len(chunks), "failed": failed, "providers": sorted(providers)}
    logger.info(
        "translate_text: %d chars → %d chunks, %s→%s, failed=%d, via=%s",
        len(text), len(chunks), from_code, to_code, failed, ",".join(stats["providers"]) or "none",
    )
    if failed:
        logger.warning("TRANSLATE_PARTIAL_FAILURE %d/%d chunks untranslated (%s→%s)",
                       failed, len(chunks), from_code, to_code)
    return " ".join(out), stats


def translate_text(text: str, from_code: str, to_code: str) -> str:
    """Translate full text by chunking. No sleeps. No crashes.

    Backwards-compatible wrapper; prefer translate_text_detailed() so failures
    are visible.
    """
    return translate_text_detailed(text, from_code, to_code)[0]
