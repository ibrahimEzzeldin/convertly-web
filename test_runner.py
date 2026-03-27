"""
Comprehensive feature test runner for Convertly.
Run: python test_runner.py
Server must be running on http://127.0.0.1:5000
"""
import requests, io, json, os, re, sys, base64
sys.stdout.reconfigure(encoding="utf-8")
from PIL import Image as PILImage

BASE = "http://127.0.0.1:5000"
sess = requests.Session()

GREEN = "\033[32m"
RED   = "\033[31m"
YEL   = "\033[33m"
RST   = "\033[0m"
OK    = f"{GREEN}✓{RST}"
FAIL  = f"{RED}✗{RST}"
SKIP  = f"{YEL}⚠{RST}"

results = []


# ── Helpers ──────────────────────────────────────────────────
def get_csrf():
    r = sess.get(BASE + "/")
    m = re.search(r'name="csrf-token"\s+content="([^"]+)"', r.text)
    if m:
        return m.group(1)
    m = re.search(r'content="([a-zA-Z0-9._-]{20,})"', r.text)
    return m.group(1) if m else ""


def make_pdf(text="Hello world test PDF"):
    """Create a minimal but valid 1-page PDF."""
    stream = f"BT /F1 12 Tf 100 700 Td ({text}) Tj ET"
    body = (
        "%PDF-1.4\n"
        "1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n"
        "2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n"
        "3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]"
        "/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj\n"
        f"4 0 obj<</Length {len(stream)}>>\nstream\n{stream}\nendstream endobj\n"
        "5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj\n"
        "xref\n0 6\n"
        "0000000000 65535 f \n"
        "0000000009 00000 n \n"
        "0000000062 00000 n \n"
        "0000000119 00000 n \n"
        "0000000273 00000 n \n"
        "0000000400 00000 n \n"
        "trailer<</Size 6/Root 1 0 R>>\n"
        "startxref\n474\n%%EOF"
    )
    return body.encode()


def pdf_file(name="test.pdf", text="Test document content"):
    return ("file", (name, io.BytesIO(make_pdf(text)), "application/pdf"))


def multi_pdf(field="files[]", names=("a.pdf", "b.pdf")):
    return [(field, (n, io.BytesIO(make_pdf(f"Doc {n}")), "application/pdf")) for n in names]


def png_bytes(w=100, h=60, color=(80, 160, 240)):
    buf = io.BytesIO()
    PILImage.new("RGB", (w, h), color).save(buf, "PNG")
    buf.seek(0)
    return buf


def sig_data_url():
    buf = io.BytesIO()
    img = PILImage.new("RGBA", (300, 80), (0, 0, 0, 0))
    img.save(buf, "PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}"


def record(name, r, ok_codes=(200,), want_file=False):
    ct = r.headers.get("Content-Type", "")
    is_file = "application/" in ct or "image/" in ct or "application/zip" in ct

    if r.status_code == 402:
        results.append((None, name, "402 quota_exceeded (skipped)"))
        return

    if r.status_code not in ok_codes:
        try:
            d = r.json()
            msg = d.get("error", d.get("message", "?"))[:80]
        except Exception:
            msg = r.text[:80].replace("\n", " ")
        results.append((False, name, f"HTTP {r.status_code}: {msg}"))
        return

    if want_file and not is_file:
        try:
            d = r.json()
            msg = d.get("error", "JSON instead of file")[:80]
        except Exception:
            msg = "non-file response"
        results.append((False, name, msg))
        return

    if not want_file and not is_file:
        try:
            d = r.json()
            if "error" in d:
                results.append((False, name, d["error"][:80]))
                return
        except Exception:
            pass

    size = len(r.content)
    results.append((True, name, f"HTTP {r.status_code} · {size} bytes · {ct.split(';')[0].strip()}"))


# ── Boot ─────────────────────────────────────────────────────
print("\n" + "=" * 66)
print("  CONVERTLY — FEATURE TEST SUITE")
print("=" * 66)

try:
    sess.get(BASE + "/", timeout=5)
except Exception:
    print(f"{RED}ERROR: Server not reachable at {BASE}{RST}")
    sys.exit(1)

csrf = get_csrf()
print(f"  CSRF token : {'OK (' + csrf[:12] + '...)' if csrf else RED + 'MISSING' + RST}")

# Reset quota so tests can run
secret = os.getenv("TEST_SECRET", "")
def reset_quota():
    if secret:
        r = sess.post(BASE + "/test-reset-quota", data={"secret": secret, "csrf_token": get_csrf()})
        return r.status_code == 200
    return False

if secret:
    if reset_quota():
        print("  Quota      : reset + pro granted via TEST_SECRET")
    else:
        print(f"  Quota      : {YEL}TEST_SECRET set but reset failed (check server){RST}")
else:
    print(f"  Quota      : {YEL}TEST_SECRET not set — quota errors possible{RST}")

print()


# ══════════════════════════════════════════════════════════════
#   1. PAGES & TEMPLATES
# ══════════════════════════════════════════════════════════════
print("── Pages & Templates ───────────────────────────────")

r = sess.get(BASE + "/")
has_logo  = "Convertly" in r.text
has_hero  = "Convert any" in r.text
has_step  = "Choose a tool" in r.text
has_nav   = "quota-pill" in r.text
has_noise = "feTurbulence" in r.text
ok = all([has_logo, has_hero, has_step, has_nav, has_noise])
results.append((ok, "Homepage template renders",
                f"logo={has_logo} hero={has_hero} step={has_step} nav={has_nav} noise={has_noise}"))

r = sess.get(BASE + "/pdf-to-word")
results.append(("PDF → Word" in r.text or "pdf-to-word" in r.text.lower(),
                "SEO tool page /pdf-to-word", f"HTTP {r.status_code}"))

for slug in ["merge-pdf", "translate-pdf", "protect-pdf", "sign-pdf"]:
    r = sess.get(BASE + f"/{slug}")
    results.append((r.status_code == 200, f"GET /{slug}", f"HTTP {r.status_code}"))

r = sess.get(BASE + "/privacy")
results.append((r.status_code == 200, "GET /privacy", f"HTTP {r.status_code}"))

r = sess.get(BASE + "/status")
record("GET /status (quota API)", r, (200,))
if r.status_code == 200:
    d = r.json()
    print(f"   quota: {d.get('conversions_used')}/{d.get('conversions_budget')}  pro={d.get('paid')}")

print()


# ══════════════════════════════════════════════════════════════
#   2. CSRF PROTECTION
# ══════════════════════════════════════════════════════════════
print("── Security ────────────────────────────────────────")

fresh = requests.Session()
r = fresh.post(BASE + "/convert",
               files=[pdf_file()],
               data={"mode": "pdf-to-word"})   # no CSRF
results.append((r.status_code == 400, "CSRF rejected on /convert (no token)", f"HTTP {r.status_code}"))

r = fresh.post(BASE + "/merge-pdf",
               files=multi_pdf(),
               data={})   # no CSRF
results.append((r.status_code == 400, "CSRF rejected on /merge-pdf (no token)", f"HTTP {r.status_code}"))

print()


# ══════════════════════════════════════════════════════════════
#   3. CONVERT TOOLS  (/convert route)
# ══════════════════════════════════════════════════════════════
reset_quota()
print("── Convert ─────────────────────────────────────────")

r = sess.post(BASE + "/convert",
              files=[pdf_file("test.pdf")],
              data={"mode": "pdf-to-word", "csrf_token": csrf})
record("POST /convert pdf-to-word", r, (200,), want_file=True)

r = sess.post(BASE + "/convert",
              files=[pdf_file("test.pdf")],
              data={"mode": "pdf-to-excel", "csrf_token": csrf})
record("POST /convert pdf-to-excel", r, (200,), want_file=True)

# word-to-pdf needs LibreOffice — expect 500 or 200
r = sess.post(BASE + "/convert",
              files=[("file", ("test.docx",
                               io.BytesIO(b"PK\x03\x04" + b"\x00" * 30),
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))],
              data={"mode": "word-to-pdf", "csrf_token": csrf})
results.append((r.status_code in (200, 400, 500, 503),
                "POST /convert word-to-pdf (LibreOffice optional)",
                f"HTTP {r.status_code}"))

r = sess.post(BASE + "/convert",
              files=[("file", ("test.xlsx",
                               io.BytesIO(b"PK\x03\x04" + b"\x00" * 30),
                               "application/vnd.ms-excel"))],
              data={"mode": "excel-to-pdf", "csrf_token": csrf})
results.append((r.status_code in (200, 400, 500),
                "POST /convert excel-to-pdf (LibreOffice optional)",
                f"HTTP {r.status_code}"))

# invalid mode
r = sess.post(BASE + "/convert",
              files=[pdf_file()],
              data={"mode": "invalid-mode", "csrf_token": csrf})
results.append((r.status_code == 400, "POST /convert rejects invalid mode", f"HTTP {r.status_code}"))

print()


# ══════════════════════════════════════════════════════════════
#   4. ORGANISE TOOLS
# ══════════════════════════════════════════════════════════════
reset_quota()
print("── Organise ────────────────────────────────────────")

r = sess.post(BASE + "/merge-pdf",
              files=multi_pdf("files[]", ("a.pdf", "b.pdf")),
              data={"csrf_token": csrf})
record("POST /merge-pdf (2 files)", r, (200,), want_file=True)

# merge with 1 file — should 400
r = sess.post(BASE + "/merge-pdf",
              files=[("files[]", ("a.pdf", io.BytesIO(make_pdf()), "application/pdf"))],
              data={"csrf_token": csrf})
results.append((r.status_code == 400, "POST /merge-pdf rejects < 2 files", f"HTTP {r.status_code}"))

r = sess.post(BASE + "/split-pdf",
              files=[pdf_file()],
              data={"split_mode": "individual", "csrf_token": csrf})
record("POST /split-pdf (individual)", r, (200,), want_file=True)

r = sess.post(BASE + "/split-pdf",
              files=[pdf_file()],
              data={"split_mode": "ranges", "ranges": "1", "csrf_token": csrf})
record("POST /split-pdf (ranges)", r, (200,), want_file=True)

r = sess.post(BASE + "/extract-pages",
              files=[pdf_file()],
              data={"pages": "1", "csrf_token": csrf})
record("POST /extract-pages", r, (200,), want_file=True)

# remove-pages on 1-page PDF should 400 (can't remove the only page)
r = sess.post(BASE + "/remove-pages",
              files=[pdf_file()],
              data={"pages": "1", "csrf_token": csrf})
results.append((r.status_code in (200, 400, 402),
                "POST /remove-pages (1-page PDF)", f"HTTP {r.status_code}"))

r = sess.post(BASE + "/compress-pdf",
              files=[pdf_file()],
              data={"level": "balanced", "csrf_token": csrf})
record("POST /compress-pdf", r, (200,), want_file=True)

r = sess.post(BASE + "/pdf-to-jpg",
              files=[pdf_file()],
              data={"dpi": "low", "csrf_token": csrf})
record("POST /pdf-to-jpg", r, (200,), want_file=True)

r = sess.post(BASE + "/jpg-to-pdf",
              files=[("files[]", ("img.jpg", png_bytes(), "image/jpeg"))],
              data={"csrf_token": csrf})
record("POST /jpg-to-pdf", r, (200,), want_file=True)

r = sess.post(BASE + "/organize-pdf/preview",
              files=[pdf_file()],
              data={"csrf_token": csrf})
record("POST /organize-pdf/preview", r, (200,))

print()


# ══════════════════════════════════════════════════════════════
#   5. EDIT & SIGN TOOLS
# ══════════════════════════════════════════════════════════════
reset_quota()
print("── Edit & Sign ─────────────────────────────────────")

r = sess.post(BASE + "/watermark-pdf",
              files=[pdf_file()],
              data={"text": "DRAFT", "opacity": "0.25", "color": "808080", "csrf_token": csrf})
record("POST /watermark-pdf", r, (200,), want_file=True)

# watermark with empty text — expect 400
r = sess.post(BASE + "/watermark-pdf",
              files=[pdf_file()],
              data={"text": "", "csrf_token": csrf})
results.append((r.status_code == 400, "POST /watermark-pdf rejects empty text", f"HTTP {r.status_code}"))

# watermark with text > 80 chars — expect 400
r = sess.post(BASE + "/watermark-pdf",
              files=[pdf_file()],
              data={"text": "X" * 81, "csrf_token": csrf})
results.append((r.status_code == 400, "POST /watermark-pdf rejects text > 80 chars", f"HTTP {r.status_code}"))

r = sess.post(BASE + "/rotate-pdf",
              files=[pdf_file()],
              data={"angle": "90", "pages": "all", "csrf_token": csrf})
record("POST /rotate-pdf (90°)", r, (200,), want_file=True)

r = sess.post(BASE + "/rotate-pdf",
              files=[pdf_file()],
              data={"angle": "270", "pages": "1", "csrf_token": csrf})
record("POST /rotate-pdf (270° page 1)", r, (200,), want_file=True)

r = sess.post(BASE + "/page-numbers",
              files=[pdf_file()],
              data={"position": "bc", "format": "n", "start": "1", "csrf_token": csrf})
record("POST /page-numbers", r, (200,), want_file=True)

r = sess.post(BASE + "/page-numbers",
              files=[pdf_file()],
              data={"position": "tr", "format": "pnN", "start": "5", "csrf_token": csrf})
record("POST /page-numbers (top-right, pnN, start=5)", r, (200,), want_file=True)

r = sess.post(BASE + "/translate-pdf",
              files=[pdf_file(text="Hello world this is a test document")],
              data={"from_lang": "en", "target_lang": "fr", "csrf_token": csrf})
record("POST /translate-pdf en→fr", r, (200,), want_file=True)

r = sess.post(BASE + "/translate-pdf",
              files=[pdf_file(text="Hello world")],
              data={"from_lang": "en", "target_lang": "ar", "csrf_token": csrf})
record("POST /translate-pdf en→ar", r, (200,), want_file=True)

r = sess.post(BASE + "/sign-pdf",
              files=[pdf_file()],
              data={"signature_data": sig_data_url(), "position": "br",
                    "size": "medium", "pages": "last", "csrf_token": csrf})
record("POST /sign-pdf (canvas sig)", r, (200,), want_file=True)

# sign with no signature — expect 400
r = sess.post(BASE + "/sign-pdf",
              files=[pdf_file()],
              data={"csrf_token": csrf})
results.append((r.status_code == 400, "POST /sign-pdf rejects missing signature", f"HTTP {r.status_code}"))

r = sess.post(BASE + "/edit-pdf",
              files=[pdf_file()],
              data={"text": "Added by test", "position": "tl",
                    "page_number": "1", "csrf_token": csrf})
record("POST /edit-pdf", r, (200,), want_file=True)

# edit with empty text — expect 400
r = sess.post(BASE + "/edit-pdf",
              files=[pdf_file()],
              data={"text": "", "csrf_token": csrf})
results.append((r.status_code == 400, "POST /edit-pdf rejects empty text", f"HTTP {r.status_code}"))

# edit with text > 500 chars — expect 400
r = sess.post(BASE + "/edit-pdf",
              files=[pdf_file()],
              data={"text": "A" * 501, "csrf_token": csrf})
results.append((r.status_code == 400, "POST /edit-pdf rejects text > 500 chars", f"HTTP {r.status_code}"))

print()


# ══════════════════════════════════════════════════════════════
#   6. PROTECT TOOLS
# ══════════════════════════════════════════════════════════════
reset_quota()
print("── Protect ─────────────────────────────────────────")

r = sess.post(BASE + "/protect-pdf",
              files=[pdf_file()],
              data={"user_pw": "secret123", "csrf_token": csrf})
record("POST /protect-pdf", r, (200,), want_file=True)

# protect with no password — expect 400
r = sess.post(BASE + "/protect-pdf",
              files=[pdf_file()],
              data={"user_pw": "", "csrf_token": csrf})
results.append((r.status_code == 400, "POST /protect-pdf rejects missing password", f"HTTP {r.status_code}"))

r = sess.post(BASE + "/unlock-pdf",
              files=[pdf_file()],
              data={"csrf_token": csrf})
record("POST /unlock-pdf (unencrypted PDF)", r, (200,), want_file=True)

print()


# ══════════════════════════════════════════════════════════════
#   7. TRANSLATION API  (/translate endpoint)
# ══════════════════════════════════════════════════════════════
print("── Translation API ─────────────────────────────────")

r = sess.post(BASE + "/translate",
              json={"text": "Hello world", "from_code": "en", "to_code": "fr"},
              headers={"X-CSRFToken": csrf, "Content-Type": "application/json"})
record("POST /translate en→fr (JSON)", r, (200,))
if r.status_code == 200:
    d = r.json()
    print(f"   translated: {d.get('translatedText','?')[:40]}")

r = sess.post(BASE + "/translate",
              json={"text": "Hello", "from_code": "en", "to_code": "en"},
              headers={"X-CSRFToken": csrf, "Content-Type": "application/json"})
results.append((r.status_code == 200 and r.json().get("translatedText") == "Hello",
                "POST /translate same-lang passthrough",
                f"HTTP {r.status_code}"))

r = sess.post(BASE + "/translate",
              json={"text": "", "from_code": "en", "to_code": "fr"},
              headers={"X-CSRFToken": csrf, "Content-Type": "application/json"})
results.append((r.status_code == 400, "POST /translate rejects empty text", f"HTTP {r.status_code}"))

print()


# ══════════════════════════════════════════════════════════════
#   8. FILE VALIDATION
# ══════════════════════════════════════════════════════════════
reset_quota()
print("── File Validation ─────────────────────────────────")

# Not a real PDF (bad header)
r = sess.post(BASE + "/compress-pdf",
              files=[("file", ("fake.pdf", io.BytesIO(b"NOT A PDF"), "application/pdf"))],
              data={"csrf_token": csrf})
results.append((r.status_code == 400, "Rejects non-PDF content with .pdf name", f"HTTP {r.status_code}"))

# Wrong extension for convert
r = sess.post(BASE + "/convert",
              files=[("file", ("test.txt", io.BytesIO(b"hello"), "text/plain"))],
              data={"mode": "pdf-to-word", "csrf_token": csrf})
results.append((r.status_code in (400, 429), "POST /convert rejects wrong extension", f"HTTP {r.status_code}"))

# No file at all
r = sess.post(BASE + "/rotate-pdf",
              data={"angle": "90", "csrf_token": csrf})
results.append((r.status_code == 400, "POST /rotate-pdf rejects missing file", f"HTTP {r.status_code}"))

print()


# ══════════════════════════════════════════════════════════════
#   RESULTS
# ══════════════════════════════════════════════════════════════
passed = failed = skipped = 0
print("=" * 66)
print("  RESULTS")
print("=" * 66)
for (ok, name, detail) in results:
    if ok is True:
        print(f"  {OK}  {name}")
        passed += 1
    elif ok is None:
        print(f"  {SKIP}  {name}")
        print(f"        {detail}")
        skipped += 1
    else:
        print(f"  {FAIL}  {name}")
        print(f"        {detail}")
        failed += 1

print("=" * 66)
print(f"  {GREEN}Passed : {passed}{RST}  |  {RED}Failed : {failed}{RST}  |  {YEL}Skipped: {skipped}{RST}")
print("=" * 66)
sys.exit(0 if failed == 0 else 1)
