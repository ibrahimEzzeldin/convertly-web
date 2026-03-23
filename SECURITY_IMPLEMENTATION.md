"""
IMPLEMENTATION GUIDE: Quick Patches for app.py
This file shows EXACT changes needed to secure Convertly.

INSTRUCTIONS:
1. Add imports at the top of app.py
2. Add privacy policy route
3. Update endpoints with security decorators (see SECURITY_AUDIT.md for details)
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 1: ADD THESE IMPORTS TO TOP OF app.py
# ═════════════════════════════════════════════════════════════════════════

# Add after existing imports (around line 12):
"""
from security_manager import (
    get_client_fingerprint,
    ConversionCounter,
    TranslationQuota,
    ProToken,
    VoucherSecurity,
    CaptchaVerifier,
    require_captcha,
    check_conversion_quota,
    LogSanitizer,
)
import jwt  # For JWT token handling
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 2: ADD PRIVACY POLICY ROUTE
# Add this route after your other routes (around line 3600):
# ═════════════════════════════════════════════════════════════════════════

"""
@app.route("/privacy")
def privacy():
    '''Privacy Policy page.'''
    return render_template("privacy.html")
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 3: UPDATE /convert ENDPOINT (HIGH PRIORITY)
# ═════════════════════════════════════════════════════════════════════════

"""
CURRENT CODE (around line 382):

@app.route("/convert", methods=["POST"])
@limiter.limit(os.getenv("CONVERT_RATE_LIMIT", "10 per minute"))
def convert():
    # ── Quota check ────────────────────────────────────────────────────────
    conversions_used   = session.get("conversions_used", 0)
    conversions_budget = session.get("conversions_budget", FREE_CONVERSIONS_LIMIT)

    if conversions_used >= conversions_budget:
        return jsonify({...}), 402

CHANGE TO:

@app.route("/convert", methods=["POST"])
@limiter.limit("5 per minute; 50 per hour")  # Stricter rate limit
@csrf.protect  # Add CSRF protection
def convert():
    # ── SERVER-SIDE Quota check ────────────────────────────────────────────
    fingerprint = get_client_fingerprint(request)
    used, budget, remaining, is_pro = ConversionCounter.get_status(fingerprint)

    if used >= budget and not is_pro:
        return jsonify({
            "error": "quota_exceeded",
            "message": "You've used all your free conversions.",
            "used": used,
            "budget": budget,
        }), 402

    # ... [existing conversion code] ...

    # AFTER successful conversion, UPDATE quota:
    used, budget, remaining = ConversionCounter.increment(fingerprint, request)
    
    # Return updated quota info
    # Add to the response:
    response.headers["X-Conversions-Used"] = str(used)
    response.headers["X-Conversions-Remaining"] = str(remaining)
    return response
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 4: UPDATE /status ENDPOINT (QUICK FIX)
# ═════════════════════════════════════════════════════════════════════════

"""
CURRENT CODE (around line 345):

@app.route("/status")
def status():
    conversions_used   = session.get("conversions_used", 0)
    conversions_budget = session.get("conversions_budget", FREE_CONVERSIONS_LIMIT)
    paid               = session.get("paid", False)
    return jsonify({...})

CHANGE TO:

@app.route("/status")
def status():
    fingerprint = get_client_fingerprint(request)
    used, budget, remaining, is_pro = ConversionCounter.get_status(fingerprint)
    
    # Check JWT Pro token
    pro_token = request.cookies.get("pro_token")
    pro_valid = False
    if pro_token:
        valid, payload = ProToken.verify(pro_token)
        if valid and payload.get("fingerprint") == fingerprint:
            pro_valid = True
    
    is_pro = is_pro or pro_valid
    
    return jsonify({
        "conversions_used":      used,
        "conversions_budget":    budget,
        "conversions_remaining": remaining,
        "paid":                  is_pro,
        "pro_valid":             is_pro,
        "free_limit":            FREE_CONVERSIONS_LIMIT,
        "paid_amount":           PAID_CONVERSIONS_AMOUNT,
        "price_usd":             PAYPAL_PRICE_USD,
    })
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 5: UPDATE /payment-success (USE JWT TOKENS)
# ═════════════════════════════════════════════════════════════════════════

"""
CURRENT CODE (around line 1280):

@app.route("/payment-success")
def payment_success():
    # ... PayPal verification ...
    
    # Payment confirmed — extend the session budget
    current_budget = session.get("conversions_budget", FREE_CONVERSIONS_LIMIT)
    session["conversions_budget"] = current_budget + PAID_CONVERSIONS_AMOUNT
    session["paid"]               = True
    session["pro_unlocked"]       = True
    session.modified = True
    
    return redirect("/invoice")

CHANGE TO:

@app.route("/payment-success")
def payment_success():
    order_id = request.args.get("token", "")
    payer_id = request.args.get("PayerID", "")

    if not order_id or not payer_id:
        logger.warning("PayPal return missing token or PayerID")
        return redirect("/?error=payment_incomplete")

    try:
        access_token = _paypal_access_token()

        # Step 1: check current order status
        check_resp = _requests.get(
            f"{PAYPAL_API_BASE}/v2/checkout/orders/{order_id}",
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=15,
        )
        check_resp.raise_for_status()
        order_data = check_resp.json()
        status = order_data.get("status")

        # Step 2: if APPROVED, capture it
        if status == "APPROVED":
            cap_resp = _requests.post(
                f"{PAYPAL_API_BASE}/v2/checkout/orders/{order_id}/capture",
                headers={
                    "Content-Type":  "application/json",
                    "Authorization": f"Bearer {access_token}",
                },
                json={},
                timeout=15,
            )
            cap_resp.raise_for_status()
            order_data = cap_resp.json()
            status = order_data.get("status")

    except Exception as exc:
        logger.error("PayPal capture failed for order %s: %s", order_id, exc)
        return redirect("/?error=payment_error")

    if status != "COMPLETED":
        logger.warning("PayPal order %s status after capture: %s", order_id, status)
        return redirect("/?error=payment_incomplete")

    # ✅ PAYMENT CONFIRMED - Grant Pro access with JWT token
    fingerprint = get_client_fingerprint(request)
    
    # Create signed JWT
    pro_token = ProToken.create(fingerprint, int(os.getenv("PAID_CONVERSIONS_AMOUNT", 20)))
    
    # Grant server-side conversions
    ConversionCounter.grant_pro(fingerprint, int(os.getenv("PAID_CONVERSIONS_AMOUNT", 20)))
    
    # Store invoice details in session
    unit    = order_data.get("purchase_units", [{}])[0]
    capture = unit.get("payments", {}).get("captures", [{}])[0]
    session["last_invoice"] = {
        "order_id":  order_data.get("id", order_id),
        "item_name": f"{PAID_CONVERSIONS_AMOUNT} Additional File Conversions",
        "price":     capture.get("amount", {}).get("value", PAYPAL_PRICE_USD),
        "currency":  capture.get("amount", {}).get("currency_code", "USD"),
        "date":      capture.get("create_time", ""),
    }
    session.modified = True

    logger.info("PayPal payment COMPLETED for order %s", order_id)
    
    response = redirect("/invoice")
    # ✅ Set HttpOnly, Secure, SameSite=Strict Pro token
    response.set_cookie(
        "pro_token",
        pro_token,
        httponly=True,
        secure=_is_production,
        samesite="Strict",
        max_age=7*24*3600,  # 7 days
    )
    return response
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 6: UPDATE /redeem-voucher (RATE LIMITING)
# ═════════════════════════════════════════════════════════════════════════

"""
CURRENT CODE (around line 3520):

@app.route("/redeem-voucher", methods=["POST"])
@limiter.limit("10 per minute")
def redeem_voucher():
    data = request.get_json(silent=True) or {}
    code = str(data.get("code", "")).strip().upper()
    
    if not code:
        return jsonify({"error": "Please enter a voucher code."}), 400
    
    valid_codes = _load_voucher_codes()
    if not valid_codes:
        return jsonify({"error": "Voucher system is not enabled."}), 503
    
    if code not in valid_codes:
        return jsonify({"error": "Invalid code."}), 400
    
    # [rest of function]

CHANGE TO:

@app.route("/redeem-voucher", methods=["POST"])
@limiter.limit("5 per minute; 10 per hour")  # Stricter limits
@csrf.protect
def redeem_voucher():
    fingerprint = get_client_fingerprint(request)
    
    # ✅ Check attempt rate limit and lockout
    allowed, message = VoucherSecurity.check_attempt(fingerprint)
    if not allowed:
        return jsonify({"error": message, "locked": True}), 429  # Too Many Requests
    
    data = request.get_json(silent=True) or {}
    code = str(data.get("code", "")).strip().upper()
    
    if not code:
        return jsonify({"error": "Please enter a voucher code."}), 400
    
    valid_codes = _load_voucher_codes()
    if not valid_codes:
        return jsonify({"error": "Voucher system is not enabled."}), 503
    
    if code not in valid_codes:
        VoucherSecurity.record_attempt(fingerprint, False)  # ✅ Record failed attempt
        return jsonify({"error": "Invalid code."}), 400
    
    # Prevent double-redeem in session
    redeemed = session.get("redeemed_vouchers", [])
    if code in redeemed:
        VoucherSecurity.record_attempt(fingerprint, False)
        return jsonify({"error": "Already redeemed."}), 400
    
    # Grant conversions
    current_budget = session.get("conversions_budget", FREE_CONVERSIONS_LIMIT)
    current_used   = session.get("conversions_used", 0)
    new_budget     = current_budget + VOUCHER_GRANT
    
    session["conversions_budget"]  = new_budget
    session["conversions_used"]    = current_used
    session["redeemed_vouchers"]   = redeemed + [code]
    session.modified = True
    
    VoucherSecurity.record_attempt(fingerprint, True)  # ✅ Record success
    
    remaining = new_budget - current_used
    logger.info("Voucher redeemed: code=%s, granted=%d", LogSanitizer.sanitize(code), VOUCHER_GRANT)
    
    return jsonify({
        "success":   True,
        "granted":   VOUCHER_GRANT,
        "remaining": remaining,
        "budget":    new_budget,
    })
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 7: UPDATE /translate-chunk (SERVER-SIDE QUOTA)
# ═════════════════════════════════════════════════════════════════════════

"""
CURRENT CODE (around line 2750):

@app.route("/translate-chunk", methods=["POST"])
@limiter.limit("60 per minute")
def translate_chunk():
    try:
        data        = request.get_json(silent=True) or {}
        text        = str(data.get("text", "")).strip()
        target_lang = str(data.get("target_lang", "English")).strip()
        
        if not text:
            return jsonify({"error": "No text provided."}), 400
        
        target_lang = _LANG_CODE_TO_NAME.get(target_lang, target_lang)
        
        is_pro = session.get("pro_unlocked", False)  # ← CLIENT-SIDE!
        
        if not is_pro:
            today = _datetime.utcnow().strftime("%Y-%m-%d")
            usage = session.get("translation_usage", {"date": today, "count": 0})
            if usage.get("date") != today:
                usage = {"date": today, "count": 0}
            if usage["count"] >= FREE_DAILY_TRANSLATIONS:
                return jsonify({...}), 402

CHANGE TO:

@app.route("/translate-chunk", methods=["POST"])
@limiter.limit("3 per minute; 10 per hour")  # Stricter limits
@csrf.protect
def translate_chunk():
    try:
        fingerprint = get_client_fingerprint(request)
        
        data        = request.get_json(silent=True) or {}
        text        = str(data.get("text", "")).strip()
        target_lang = str(data.get("target_lang", "English")).strip()
        
        if not text:
            return jsonify({"error": "No text provided."}), 400
        
        target_lang = _LANG_CODE_TO_NAME.get(target_lang, target_lang)
        
        # ✅ Check JWT Pro token
        pro_token = request.cookies.get("pro_token")
        is_pro = False
        if pro_token:
            valid, payload = ProToken.verify(pro_token)
            if valid and payload.get("fingerprint") == fingerprint:
                is_pro = True
        
        if not is_pro:
            # ✅ Check server-side daily quota
            allowed, used, remaining = TranslationQuota.check_and_increment(
                fingerprint, 
                limit=FREE_DAILY_TRANSLATIONS
            )
            
            if not allowed:
                return jsonify({
                    "error":   "daily_limit_reached",
                    "used":    used,
                    "limit":   FREE_DAILY_TRANSLATIONS,
                    "message": f"You\\'ve used your {FREE_DAILY_TRANSLATIONS} free daily translations.",
                }), 402
        
        target_code = _MYMEMORY_LANG_CODE.get(target_lang, target_lang.lower())
        translated  = _ts_translate_text(text, "en", target_code)
        
        remaining = None
        if not is_pro:
            remaining = FREE_DAILY_TRANSLATIONS - used
        
        is_rtl = target_lang in _RTL_LANG_NAMES
        return jsonify({"translatedText": translated, "remaining": remaining, "isRtl": is_rtl})
    
    except Exception as exc:
        logger.error("translate-chunk error: %s", exc)
        return jsonify({"error": str(exc)}), 500
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 8: ADD CSRF PROTECTION TO ALL POST ENDPOINTS
# Example - do this for EVERY state-changing POST:
# ═════════════════════════════════════════════════════════════════════════

"""
// BEFORE:
@app.route("/merge-pdf", methods=["POST"])
@limiter.limit("10 per minute")
def merge_pdf_route():

// AFTER:
@app.route("/merge-pdf", methods=["POST"])
@limiter.limit("5 per minute")
@csrf.protect  # ✅ Add this
def merge_pdf_route():
"""

# ═════════════════════════════════════════════════════════════════════════
# PART 9: INJECT TURNSTILE SITE KEY INTO TEMPLATES
# ═════════════════════════════════════════════════════════════════════════

"""
Add this to app.before_request or as a context processor:

@app.context_processor
def inject_turnstile():
    return dict(turnstile_site_key=os.getenv("TURNSTILE_SITE_KEY", ""))

Then add to index.html before the translation form:

<script src="https://challenges.cloudflare.com/turnstile/v0/api.js" async defer></script>

<div class="cf-turnstile" data-sitekey="{{ turnstile_site_key }}" data-theme="light"></div>
"""

# ═════════════════════════════════════════════════════════════════════════
# TESTING COMMANDS
# ═════════════════════════════════════════════════════════════════════════

"""
After implementing the changes, test with:

# Test 1: Conversion quota (should fail after 3)
for i in {1..5}; do
  curl -X POST http://localhost:5000/convert \
    -F "file=@test.pdf" \
    -F "mode=pdf-to-word"
  echo "Attempt $i"
  sleep 1
done

# Test 2: Voucher rate limiting (should lock after 5 attempts)
for i in {1..10}; do
  curl -X POST http://localhost:5000/redeem-voucher \
    -H "Content-Type: application/json" \
    -d '{"code":"INVALID"}'
  echo "Attempt $i"
  sleep 1
done

# Test 3: Check status endpoint
curl http://localhost:5000/status

# Test 4: Check privacy page
curl http://localhost:5000/privacy
"""

# ═════════════════════════════════════════════════════════════════════════
# ENVIRONMENT VARIABLES TO ADD
# ═════════════════════════════════════════════════════════════════════════

"""
Add to .env:

# Security
SECRET_KEY=generate_new_strong_random_32_char_key_here
RATELIMIT_STORAGE_URI=memory://  # Use redis:// in production

# CAPTCHA (from https://dash.cloudflare.com/)
TURNSTILE_SITE_KEY=1x00000000000000000000AA
TURNSTILE_SECRET_KEY=1x0000000000000000000000000000000AA

# PayPal Webhook (from PayPal dashboard)
PAYPAL_WEBHOOK_ID=WH_xxxxxxxxxxxxxxxxxxxxxxx

# Existing
FLASK_ENV=production
FLASK_DEBUG=False
"""

# ═════════════════════════════════════════════════════════════════════════
# DEPLOYMENT CHECKLIST
# ═════════════════════════════════════════════════════════════════════════

"""
Before deploying:

☐ Update app.py with all security imports
☐ Add security_manager import: from security_manager import ...
☐ Update /convert endpoint with ConversionCounter
☐ Update /status endpoint with server-side quota
☐ Update /payment-success with JWT tokens
☐ Update /redeem-voucher with VoucherSecurity
☐ Update /translate-chunk with TranslationQuota
☐ Add @csrf.protect to all POST endpoints
☐ Add @limiter.limit() with stricter limits
☐ Create /privacy route
☐ Add Turnstile CAPTCHA script to templates
☐ Update .env with all new variables
☐ Run: pip install -r requirements.txt
☐ Run: python app.py (local test)
☐ Test all quota limits work
☐ Test voucher rate limiting
☐ Git commit and push
☐ Monitor logs on Render for any errors

After deployment:

☐ Check /status endpoint returns correct data
☐ Check /privacy page loads
☐ Test conversion quota resets properly
☐ Test translation daily limit
☐ Verify PayPal payments work
☐ Check logs contain no passwords
☐ Monitor for abuse attempts
"""

print("✅ Implementation guide created!")
print("📖 See SECURITY_AUDIT.md for detailed documentation")
print("🔒 All security fixes are ready to implement")
