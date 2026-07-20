# Google Play — developer page & store listing reference

Values and assets for the Play Console developer profile. Assets live in
`static/icons/`. Character counts were measured, not estimated.

---

## Developer profile (الملف الشخصي للمطوّر)

| Field | Value |
|---|---|
| `رمز مطوِّر` — Developer icon | `static/icons/play-store-icon-512.png` — 512×512, 24-bit, no alpha, 25 KB |
| `صورة العنوان` — Header image | `static/icons/play-developer-header-4096x2304.jpg` — 4096×2304, 24-bit, 213 KB |
| `الموقع الإلكتروني للمطوّر` | `https://convertly-web.onrender.com` |
| `تطبيق مميز` — Featured app | Blocked until the first build is uploaded |
| Developer name (store listing) | `Convertly` — set in Developer account → Account details, separate from the payments-profile legal name |

Both images satisfy Play's "24-bit, non-transparent, ≤1 MB" rule. The header keeps
all content in a centred safe area because Play crops it hard on narrow viewports.

## Promotional text (نص ترويجي) — max 140 chars

Default listing language is Arabic (`ar`).

| # | Text | Chars |
|---|---|---|
| **AR-1** *(recommended)* | حوِّل ملفاتك في ثوانٍ: PDF وWord وExcel والصور. بدون تسجيل، وبدون علامة مائية. | 78 |
| AR-2 | أدوات PDF مجانية وسريعة: تحويل ودمج وضغط وتوقيع وترجمة. بدون تسجيل، وملفاتك تُحذف بعد المعالجة. | 95 |
| AR-3 | حوِّل ودمج واضغط ملفات PDF من متصفحك مباشرة. مجانًا، وبدون حساب، وبخصوصية كاملة. | 80 |
| EN-1 | Convert your files in seconds: PDF, Word, Excel and images. No sign-up, no watermarks. | 86 |

AR-1 leads with the outcome and closes on the two objections the landing page
already answers ("no sign-up, no watermarks"). AR-2 is the option to use if
listing tool breadth matters more than speed.

---

## Data safety declaration — source material

Answer from `SECURITY_IMPLEMENTATION.md`; do not guess. The app **does** collect
user files, so declaring "no data collected" would be false.

- Files are uploaded for processing and deleted afterwards (`cleanup_old_files()`)
- Transfer is encrypted
- No account, no sign-up, no advertising ID
- Deletion requests aren't applicable because storage is ephemeral — say so
  explicitly rather than leaving it blank

Privacy and terms pages already exist and return 200: `/privacy`, `/terms`.

## Known gaps before production

- **Purchases are not durable.** `security_manager.py:31` defaults the quota DB to
  `/tmp/quota.db`; Render is on the Free plan (no persistent disk, sleeps when idle);
  `ProToken.EXPIRY_DAYS = 7`. Must be fixed before Play Billing is enabled.
- **Monetising publishes the developer's full legal address**, per Play's consumer
  protection policy. A business address requires a registered entity.
