# Pre-launch issues — found during closed testing

Collected 2026-07-22 from production logs and a code read, while the app is in
closed testing. Ordered by impact on user experience and store ratings.

---

## 1. Scanned PDFs charged a credit and returned a blank file — FIXED

**Evidence (production log, 2026-07-22 06:25):**

```
[WARNING] Words count: 0. It might be a scanned pdf, which is not supported yet.
[INFO] Terminated in 1.49s.
[INFO] Conversion counter incremented: fingerprint=362889e455df9df5, used=1, budget=3
```

A user uploaded a scanned Emirates ID. `pdf2docx` found no text, wrote a valid
but **empty** .docx, and `/convert` treated it as success because the only
post-check was `if not os.path.exists(out)` — the file exists, it is just empty.
The user lost 1 of 3 free conversions and got a blank document with no
explanation.

This is the worst possible first impression: it looks like the product simply
does not work, and it costs the user a credit to find out. `/translate-pdf`
already rejected scanned PDFs correctly; `/convert` did not.

**Fix:** pre-check PDFs for a text layer immediately after upload, before any
conversion work and before the quota increment. Returns the same message
`/translate-pdf` uses, and logs `SCANNED_PDF_REJECTED` so these can be counted.
The check is wrapped so a pre-check failure never blocks a legitimate
conversion.

**Verified:** scanned → HTTP 400 + message + **0 credits consumed** (3 uploads,
quota unchanged); text PDF → HTTP 200, converts and increments as before.

---

## 2. Translation failures are silent — NOT FIXED

`translation_service.py:192` returns the **original untranslated text** when
every provider fails:

```python
logger.warning("All translation providers failed for chunk: '%s'", text[:40])
return text
```

The user receives a "translated" PDF still in the source language, with no error
and a credit spent. Users will read this as a broken product.

**Suggested:** propagate a failure count to the route and return a real error (or
a partial-success warning) when a meaningful share of chunks fall through.

## 3. DeepL is configured in code but never used — NOT FIXED

`_translate_deepl()` returns `None` immediately unless `DEEPL_API_KEY` is set,
and that key is absent from both `.env.example` and `render.yaml`. So **all**
translation silently falls back to MyMemory, whose free tier is roughly 5k
chars/day anonymous (50k with `MYMEMORY_EMAIL` set). A single 25-page PDF can
exhaust it, after which everything hits issue #2 and returns untranslated text.

**Suggested:** either set `DEEPL_API_KEY` and add it to `render.yaml`, or drop
the DeepL path so the fallback behaviour is honest.

## 4. Language codes are not mapped on one of the two paths — NOT FIXED

`app.py:_translate_one_chunk` builds the pair raw:

```python
params={"q": text, "langpair": f"{src_code}|{target_code}"}
```

while `translation_service.py` maps codes through `LANG_CODES` first, with the
comment that MyMemory *"only recognizes simple language codes like 'ar', 'en',
'fr', not region variants."* `_MYMEMORY_LANG_CODE` emits `zh-CN` and `zh-TW`, so
Chinese requests on that path send exactly the region-coded values the code says
are unsupported.

There are also **two independent translation engines** in the codebase
(`translation_service.py` and `app.py::_mymemory_translate`), which is why the
two behave differently. Worth collapsing to one.

## 5. Translation throughput limits — BY DESIGN, worth surfacing to users

| Limit | Value | Location |
|---|---|---|
| Max pages | 25 | `app.py` translate-pdf |
| Chunk size | 400 (app.py) / 500 (service) | inconsistent between the two engines |
| Delay between chunks | 1.0s | `_CHUNK_DELAY` |
| Target languages | 24 | `TRANSLATE_TARGET_LANGS` |
| OCR | none | scanned PDFs unsupported everywhere |

At 1 second per ~400-char chunk, a large document takes minutes. The UI should
set that expectation rather than appearing to hang.

---

## Collecting errors from here

There is no error aggregation today; failures are `WARNING` lines that scroll
away. The cheapest useful step is a consistent, greppable prefix on each failure
class so Render logs can be counted:

- `SCANNED_PDF_REJECTED` (added in this change)
- `TRANSLATE_RATE_LIMITED`
- `TRANSLATE_SILENT_FALLBACK`
- `CONVERSION_FAILED`

Then `grep -c SCANNED_PDF_REJECTED` on the Render logs gives a real number per
failure mode, which is enough to prioritise before going live.
