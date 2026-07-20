# Android (Trusted Web Activity) build

The Android app is a TWA: a thin native shell that renders
`https://convertly-web.onrender.com` full-screen with no address bar. Deploying to
Render ships the Android app too — there is no separate frontend to maintain.

`twa-manifest.json` is hand-authored rather than produced by `bubblewrap init`,
which is interactive and can't run unattended.

## Load-bearing settings

| Setting | Why |
|---|---|
| `packageId: com.convertly.app` | **Permanent.** Cannot change after the first upload |
| `startUrl: "/?src=twa"` | Makes in-app detection deterministic. The `android-app://` referrer only survives the entry navigation, so relying on it alone is fragile |
| `features.playBilling.enabled` | Puts `com.android.vending.BILLING` in build #1. Play won't let you create the $2 in-app product until a build declaring that permission is uploaded — so including it now avoids a throwaway upload later |
| `iconUrl` / `maskableIconUrl` | Pulled from production, generated in `static/icons/` |

Play Billing is declared but **not wired up**: the paywall hides every payment path
inside the TWA (see `templates/base.html`). Android is free-tier only until the
durability work and the billing integration land.

## Prerequisites

```bash
npm install -g @bubblewrap/cli
bubblewrap doctor     # first run downloads JDK 17 + Android SDK (~hundreds of MB)
```

## Generating the signing key

**Run this yourself.** The password must not pass through a chat transcript or a
command history that gets shared.

```bash
cd android
"$HOME/.bubblewrap/jdk/jdk-17.0.11+9/bin/keytool" -genkeypair \
  -v -keystore android.keystore -alias convertly \
  -keyalg RSA -keysize 2048 -validity 10000
```

Then:

- **Back the keystore up somewhere you will still have in five years.** Losing it
  means you cannot ship updates.
- **Enrol in Play App Signing** at first upload. Google then holds the app signing
  key and this file becomes only the *upload* key — recoverable if lost. Without
  enrolment, losing this file is terminal for the app.
- `*.keystore` is gitignored. Never commit it.

## Building

```bash
cd android
bubblewrap build
```

Produces `app-release-bundle.aab` for Play, and `app-release-signed.apk` for
sideloading during testing.

## Digital Asset Links

`bubblewrap build` prints the SHA-256 fingerprint of the signing key. Put it in
`static/assetlinks.json` (see the `/.well-known/assetlinks.json` route in
`app.py`), in this shape:

```json
[{
  "relation": ["delegate_permission/common.handle_all_urls"],
  "target": {
    "namespace": "android_app",
    "package_name": "com.convertly.app",
    "sha256_cert_fingerprints": ["<fingerprint>"]
  }
}]
```

**Use the fingerprint Play Console shows after enrolling in Play App Signing**, not
the local upload key's — Google re-signs with its own key, so the local one will not
verify in production. Getting this wrong is why a TWA shows an address bar.

Verify after deploying:

```bash
curl -s https://convertly-web.onrender.com/.well-known/assetlinks.json
```
