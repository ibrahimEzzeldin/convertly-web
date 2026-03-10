"""
Comprehensive feature test for convertly-web.
Run: python test_features.py
"""

import requests, io, re, sys, os, time

BASE = "http://127.0.0.1:5000"

# ── helpers ──────────────────────────────────────────────────────────────────

def get_session():
    s = requests.Session()
    r = s.get(BASE + "/")
    assert r.status_code == 200, f"Home page returned {r.status_code}"
    m = re.search(r'<meta name="csrf-token" content="([^"]+)"', r.text)
    assert m, "CSRF token not found in page"
    token = m.group(1)
    return s, token

def make_minimal_pdf():
    """Return bytes of a tiny but valid single-page PDF with some text."""
    content = b"""%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/MediaBox[0 0 595 842]/Parent 2 0 R/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj
4 0 obj<</Length 44>>stream
BT /F1 12 Tf 72 720 Td (Hello World Test PDF) Tj ET
endstream
endobj
5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj
xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000266 00000 n
0000000360 00000 n
trailer<</Size 6/Root 1 0 R>>
startxref
441
%%EOF"""
    return content

def make_multi_page_pdf():
    """Return bytes of a 3-page PDF."""
    import fitz
    doc = fitz.open()
    for i in range(3):
        page = doc.new_page()
        page.insert_text((72, 720), f"Test page {i+1} — Hello World", fontsize=14)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()

RESULTS = []

def test(name, fn):
    try:
        fn()
        RESULTS.append(("PASS", name))
        print(f"  [PASS]  {name}")
    except Exception as e:
        RESULTS.append(("FAIL", name, str(e)))
        print(f"  [FAIL]  {name}: {e}")

# ── test functions ────────────────────────────────────────────────────────────

def t_home():
    r = requests.get(BASE + "/")
    assert r.status_code == 200
    assert "Convertly" in r.text or "convertly" in r.text.lower() or "PDF" in r.text

def t_status():
    r = requests.get(BASE + "/status")
    assert r.status_code == 200
    d = r.json()
    assert "conversions_remaining" in d

def t_pdf_to_word():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/convert", files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"mode": "pdf-to-word", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    assert len(r.content) > 100

def t_pdf_to_excel():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/convert", files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"mode": "pdf-to-excel", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    assert len(r.content) > 100

def t_merge_pdf():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/merge-pdf",
               files=[("files[]", ("a.pdf", pdf, "application/pdf")),
                      ("files[]", ("b.pdf", pdf, "application/pdf"))],
               data={"csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    assert len(r.content) > 100

def t_split_pdf():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/split-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"ranges": "1,2", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_remove_pages():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/remove-pages",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"pages": "1", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_extract_pages():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/extract-pages",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"pages": "1-2", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_compress_pdf():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/compress-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"quality": "lossless", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_repair_pdf_removed():
    """Confirm /repair-pdf endpoint was removed (404)."""
    r = requests.get(BASE + "/repair-pdf")
    assert r.status_code == 404, f"Expected 404 but got {r.status_code}"

def t_ocr_removed():
    """Confirm /ocr-pdf endpoint was removed (404)."""
    r = requests.get(BASE + "/ocr-pdf")
    assert r.status_code == 404, f"Expected 404 but got {r.status_code}"

def t_pdf_to_jpg():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/pdf-to-jpg",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"dpi": "72", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_watermark_pdf():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/watermark-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"text": "CONFIDENTIAL", "position": "center",
                     "opacity": "30", "font_size": "36", "pages": "all",
                     "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_rotate_pdf():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/rotate-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"angle": "90", "pages": "all", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_protect_pdf():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/protect-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"user_pw": "test123", "owner_pw": "owner123",
                     "allow_print": "on", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_unlock_pdf():
    """Test unlock with an unprotected PDF (should succeed or return graceful error)."""
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/unlock-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"password": "", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    # Either 200 (already unlocked) or 400 (needs password) — both are valid JSON
    assert r.status_code in (200, 400), f"status={r.status_code}"

def t_page_numbers():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/page-numbers",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"position": "bc", "start": "1", "font_size": "10",
                     "pages": "all", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_organize_preview():
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/organize-pdf/preview",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    d = r.json()
    assert "thumbnails" in d
    assert d["page_count"] == 3

def t_extract_pdf_paragraphs():
    """Test the fix for the fitz import bug that caused translation to fail."""
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/extract-pdf-paragraphs",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"csrf_token": tok},
               headers={"X-CSRFToken": tok})
    # Should return JSON (not HTML 500)
    assert r.status_code in (200, 400), f"status={r.status_code}, body starts: {r.text[:100]}"
    assert r.headers.get("Content-Type", "").startswith("application/json"), \
        f"Expected JSON, got: {r.headers.get('Content-Type')} body: {r.text[:100]}"

def t_translate_chunk():
    s, tok = get_session()
    r = s.post(BASE + "/translate-chunk",
               json={"text": "Hello, this is a test.", "target_lang": "French", "csrf_token": tok},
               headers={"Content-Type": "application/json",
                        "X-CSRFToken": tok, "X-CSRF-Token": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    d = r.json()
    assert "translated" in d, f"No 'translated' key: {d}"
    assert len(d["translated"]) > 0

def t_pdf_page_preview():
    """Test the new /pdf-page-preview endpoint for edit-pdf visual placement."""
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/pdf-page-preview",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"page": "0", "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    d = r.json()
    assert "image" in d, f"Missing 'image': {list(d.keys())}"
    assert d["total_pages"] == 3
    assert len(d["image"]) > 100  # base64 should have content

def t_edit_pdf_with_coords():
    """Test edit PDF with click-based x_pct/y_pct coordinates."""
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/edit-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"text": "Added by test", "x_pct": "20", "y_pct": "10",
                     "pages": "all", "font_size": "12", "color": "000000",
                     "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"
    assert len(r.content) > 100

def t_sign_pdf():
    """Test sign PDF with a signature image."""
    import fitz as _fitz
    # Generate a real PNG signature using fitz
    doc = _fitz.open()
    page = doc.new_page(width=200, height=80)
    page.draw_line(_fitz.Point(10, 40), _fitz.Point(190, 40), color=(0,0,0), width=2)
    pix = page.get_pixmap()
    sig_bytes = pix.tobytes("png")
    doc.close()

    s, tok = get_session()
    pdf = make_multi_page_pdf()
    r = s.post(BASE + "/sign-pdf",
               files={"file": ("test.pdf", pdf, "application/pdf"),
                      "signature_file": ("sig.png", sig_bytes, "image/png")},
               data={"position": "br", "pages": "last", "size": "medium",
                     "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code == 200, f"status={r.status_code} body={r.text[:200]}"

def t_extract_text_region():
    """Test the new /extract-text-region endpoint for highlight translation."""
    s, tok = get_session()
    pdf = make_multi_page_pdf()
    import json
    regions = json.dumps([{"page": 0, "x0_pct": 0, "y0_pct": 0, "x1_pct": 100, "y1_pct": 100}])
    r = s.post(BASE + "/extract-text-region",
               files={"file": ("test.pdf", pdf, "application/pdf")},
               data={"regions": regions, "csrf_token": tok},
               headers={"X-CSRFToken": tok})
    assert r.status_code in (200, 400), f"status={r.status_code} body={r.text[:200]}"
    assert r.headers.get("Content-Type", "").startswith("application/json"), \
        f"Expected JSON, got: {r.headers.get('Content-Type')}"

def t_argos_translate_endpoint():
    """Test the new /translate endpoint (Argos Translate with MyMemory fallback)."""
    s, tok = get_session()
    r = s.post(BASE + "/translate",
               json={"text": "Hello world", "from_code": "en", "to_code": "fr"},
               headers={"Content-Type": "application/json",
                        "X-CSRFToken": tok, "X-CSRF-Token": tok})
    # 200 = translated (Argos or fallback), 500 = both backends down (acceptable in test env)
    assert r.status_code in (200, 500), f"status={r.status_code} body={r.text[:200]}"
    if r.status_code == 200:
        d = r.json()
        assert "translatedText" in d, f"No 'translatedText' key: {d}"

def t_translate_modal_in_ui():
    """Confirm new translate modal with highlight tab is in the HTML."""
    r = requests.get(BASE + "/")
    assert "trModal" in r.text,            "trModal not in UI"
    assert "trTabHighlight" in r.text,     "trTabHighlight not in UI"
    assert "trHighlightCanvas" in r.text,  "trHighlightCanvas not in UI"
    assert "trDownloadBtn" in r.text,      "trDownloadBtn not in UI"
    assert "extract-text-region" in r.text, "extract-text-region not referenced in UI"

def t_no_repair_in_ui():
    """Confirm 'repair-pdf' mode button is gone from the HTML."""
    r = requests.get(BASE + "/")
    assert 'data-mode="repair-pdf"' not in r.text, "Repair PDF button still in UI"

def t_edit_preview_in_ui():
    """Confirm edit PDF preview modal exists in the HTML."""
    r = requests.get(BASE + "/")
    assert "editPreviewModal" in r.text, "editPreviewModal not in UI"
    assert "editOpenPreviewBtn" in r.text, "editOpenPreviewBtn not in UI"
    assert "pdf-page-preview" in r.text, "pdf-page-preview endpoint not referenced in UI"

# ── run all tests ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("\n" + "="*60)
    print("  CONVERTLY-WEB FEATURE TESTS")
    print("="*60)

    TESTS = [
        ("Home page loads",                  t_home),
        ("Status endpoint",                  t_status),
        ("PDF to Word",                      t_pdf_to_word),
        ("PDF to Excel",                     t_pdf_to_excel),
        ("Merge PDF",                        t_merge_pdf),
        ("Split PDF",                        t_split_pdf),
        ("Remove pages",                     t_remove_pages),
        ("Extract pages",                    t_extract_pages),
        ("Compress PDF",                     t_compress_pdf),
        ("Repair PDF removed (404)",         t_repair_pdf_removed),
        ("OCR PDF removed (404)",                  t_ocr_removed),
        ("PDF to JPG",                       t_pdf_to_jpg),
        ("Watermark PDF",                    t_watermark_pdf),
        ("Rotate PDF",                       t_rotate_pdf),
        ("Protect PDF",                      t_protect_pdf),
        ("Unlock PDF (unprotected)",         t_unlock_pdf),
        ("Page numbers",                     t_page_numbers),
        ("Organize PDF preview",             t_organize_preview),
        ("Extract PDF paragraphs (fitz fix)",t_extract_pdf_paragraphs),
        ("Translate chunk (MyMemory)",       t_translate_chunk),
        ("Extract text region (new)",        t_extract_text_region),
        ("Argos /translate endpoint",        t_argos_translate_endpoint),
        ("PDF page preview (new endpoint)",  t_pdf_page_preview),
        ("Edit PDF with coordinates",        t_edit_pdf_with_coords),
        ("Sign PDF",                         t_sign_pdf),
        ("UI: Repair PDF button removed",    t_no_repair_in_ui),
        ("UI: Edit PDF preview modal added", t_edit_preview_in_ui),
        ("UI: Translate modal updated",      t_translate_modal_in_ui),
    ]

    for name, fn in TESTS:
        test(name, fn)

    print("\n" + "="*60)
    passed = sum(1 for r in RESULTS if r[0] == "PASS")
    failed = sum(1 for r in RESULTS if r[0] == "FAIL")
    print(f"  Results: {passed} passed, {failed} failed out of {len(RESULTS)} tests")
    if failed:
        print("\n  FAILURES:")
        for r in RESULTS:
            if r[0] == "FAIL":
                print(f"    [FAIL] {r[1]}: {r[2]}")
    print("="*60 + "\n")
    sys.exit(0 if failed == 0 else 1)
