# Data safety & content rating — prepared answers

Every answer below is traced to code. Confirm each one before submitting: these are
attestations, and a false declaration risks suspension.

---

## Data safety

### Does the app collect or share user data? **Yes**

Answering "no" would be false. The app receives user files, and the translate
feature transmits document text to third parties.

### Data types

| Type | Collected | Shared | Why |
|---|---|---|---|
| **Files and docs** | Yes | **Yes** | Uploaded for processing. Translate sends extracted text to DeepL and MyMemory (`translation_service.py:12`, `:99`) |
| Name, email, phone | No | No | No account, no sign-up |
| Location, contacts, photos | No | No | Never requested |
| Advertising ID | No | No | No ads, no analytics SDK |
| Payment info | No | No | PayPal handles it off-site; card details never touch the server |

### Required follow-ups for "Files and docs"

- **Purpose:** App functionality
- **Is collection required or optional?** Required — the tool cannot work without a file
- **Is data processed ephemerally?** Yes — deleted after processing (`cleanup_old_files()`)
- **Shared with third parties?** **Yes** — DeepL and MyMemory, for translation only
- **Encrypted in transit?** Yes — HTTPS
- **Can users request deletion?** No mechanism, because storage is ephemeral. Say so
  explicitly rather than leaving it blank.

### The one that is easy to get wrong

Files are deleted after processing, which makes "we don't keep anything" feel true —
but *sharing* is a separate question from *retention*. Translated text leaves the
server for DeepL and MyMemory. That must be declared.

---

## Content rating questionnaire

Category: **Utility / Productivity / Communication** (not a game)

| Question | Answer |
|---|---|
| Violence, blood, or scary content | No |
| Sexual or suggestive content | No |
| Profanity or crude humour | No |
| Drugs, alcohol, tobacco | No |
| Gambling or simulated gambling | No |
| User-generated content shared between users | **No** — files are private to the person who uploaded them; nothing is published or shared with other users |
| Users can interact or communicate | No |
| Shares user location | No |
| Allows purchase of digital goods | **No** for the current Android build — the TWA hides every payment path. **Change this to Yes** when Play Billing ships |
| Personal/sensitive info collected | Files only, as declared above |

Expected outcome: **Everyone / PEGI 3**.

Note: users may upload documents containing anything, but the rating covers content
*the app itself* presents, not what a user feeds into it.

---

## Also required before closed testing

- Privacy policy URL: `https://convertly-web.onrender.com/privacy` (live, 200)
- App category: Productivity
- Contact email: use a monitored address, not a personal Gmail — it is public
- Target audience: 18+ or 13+. Avoid selecting under-13, which triggers Families
  Policy obligations this app is not built for.
