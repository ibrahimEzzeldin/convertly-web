"""
translation_service.py — LibreTranslate (primary) + MyMemory (fallback)
"""
import time
import logging
import requests

logger = logging.getLogger(__name__)

MAX_CHUNK_SIZE       = 400
DELAY_BETWEEN_CHUNKS = 0.5   # seconds between chunks
MAX_RETRIES          = 3

LIBRETRANSLATE_URL = "https://libretranslate.com/translate"
MYMEMORY_URL       = "https://api.mymemory.translated.net/get"

# Language name → ISO code used by both APIs
LANG_MAP = {
    "English":              "en",
    "Arabic":               "ar",
    "French":               "fr",
    "Spanish":              "es",
    "German":               "de",
    "Italian":              "it",
    "Portuguese":           "pt",
    "Russian":              "ru",
    "Chinese (Simplified)": "zh",
    "Chinese (Traditional)":"zh",
    "Japanese":             "ja",
    "Korean":               "ko",
    "Turkish":              "tr",
    "Dutch":                "nl",
    "Polish":               "pl",
    "Hindi":                "hi",
    "Ukrainian":            "uk",
    "Swedish":              "sv",
    "Norwegian":            "no",
    "Danish":               "da",
    "Greek":                "el",
    "Hebrew":               "he",
    "Indonesian":           "id",
    "Vietnamese":           "vi",
}

# MyMemory uses different codes for some locales
_MYMEMORY_CODE_OVERRIDE = {
    "Chinese (Simplified)":  "zh-CN",
    "Chinese (Traditional)": "zh-TW",
}


def split_into_chunks(text, max_size=MAX_CHUNK_SIZE):
    """Split text into sentence-boundary chunks each ≤ max_size chars."""
    sentences = []
    for line in text.replace('\n', ' \n ').split('\n'):
        for s in (
            line.replace('. ', '.|')
                .replace('! ', '!|')
                .replace('? ', '?|')
                .split('|')
        ):
            if s.strip():
                sentences.append(s.strip())

    chunks, current = [], ""
    for sentence in sentences:
        if len(current) + len(sentence) + 1 <= max_size:
            current += (" " if current else "") + sentence
        else:
            if current:
                chunks.append(current)
            if len(sentence) > max_size:
                for i in range(0, len(sentence), max_size):
                    chunks.append(sentence[i:i + max_size])
                current = ""
            else:
                current = sentence
    if current:
        chunks.append(current)
    return chunks or [text[:max_size]]


def translate_with_libretranslate(text, from_code, to_code):
    """Try LibreTranslate public API. Returns translated string or None on failure."""
    try:
        resp = requests.post(
            LIBRETRANSLATE_URL,
            json={"q": text, "source": from_code, "target": to_code, "format": "text"},
            timeout=20,
        )
        if resp.status_code == 200:
            data = resp.json()
            result = data.get("translatedText", "").strip()
            if result:
                return result
        logger.warning("LibreTranslate returned %d: %s", resp.status_code, resp.text[:200])
    except Exception as exc:
        logger.warning("LibreTranslate request failed: %s", exc)
    return None


def translate_with_mymemory(text, from_code, to_code):
    """Try MyMemory API with retry. Returns translated string or None on failure."""
    for attempt in range(MAX_RETRIES):
        try:
            resp = requests.get(
                MYMEMORY_URL,
                params={"q": text, "langpair": f"{from_code}|{to_code}"},
                timeout=15,
            )
            data = resp.json()
            translated = data.get("responseData", {}).get("translatedText", "")
            if resp.status_code == 429 or data.get("responseStatus") == 429 \
                    or "MYMEMORY WARNING" in str(translated):
                wait = 3 * (attempt + 1)
                logger.warning("MyMemory rate limit, retrying in %ds (attempt %d)", wait, attempt + 1)
                time.sleep(wait)
                continue
            if data.get("responseStatus") == 200 and translated:
                return translated.strip()
            logger.warning("MyMemory error %s: %s", data.get("responseStatus"), data.get("responseDetails", ""))
        except Exception as exc:
            logger.warning("MyMemory request failed (attempt %d): %s", attempt + 1, exc)
            time.sleep(2)
    return None


def translate_chunk(text, from_code, to_code):
    """Translate a single chunk: LibreTranslate first, MyMemory as fallback.
    Returns translated text, or original text if both engines fail."""
    if not text.strip():
        return text

    result = translate_with_libretranslate(text, from_code, to_code)
    if result:
        return result

    result = translate_with_mymemory(text, from_code, to_code)
    if result:
        return result

    logger.warning("Both translation engines failed — returning original text")
    return text


def translate_text(text, from_code, to_code):
    """Translate arbitrary-length text by chunking then joining results."""
    if not text.strip():
        return text
    if from_code == to_code:
        return text

    chunks = split_into_chunks(text)
    out = []
    for idx, chunk in enumerate(chunks):
        if idx > 0:
            time.sleep(DELAY_BETWEEN_CHUNKS)
        out.append(translate_chunk(chunk, from_code, to_code))
    return " ".join(out)
