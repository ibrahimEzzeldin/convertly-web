"""translation_service.py — Argos Translate integration for offline PDF translation.

Language packs (~100 MB each) are downloaded from the Argos package index on first use
and cached in the default Argos data directory.

Render.com free-tier note:
  Language packs are re-downloaded on every dyno restart because the free tier has no
  persistent disk.  Consider mounting a Render Disk (paid plan) and setting the
  environment variable ARGOS_PACKAGES_DIR to that mount path so packs survive restarts:
    export ARGOS_PACKAGES_DIR=/var/data/argos-packages
"""

from __future__ import annotations

import logging
import os
import threading
from typing import Optional

logger = logging.getLogger(__name__)

# Thread-safe set tracking which (from_code, to_code) pairs were installed this process
_installed: set[tuple[str, str]] = set()
_install_lock = threading.Lock()

# Optional: redirect Argos package storage to a persistent path
_packages_dir = os.getenv("ARGOS_PACKAGES_DIR")
if _packages_dir:
    os.makedirs(_packages_dir, exist_ok=True)
    os.environ["ARGOS_PACKAGES_DIR"] = _packages_dir


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _is_pair_installed(from_code: str, to_code: str) -> bool:
    """Return True if the language pair is already installed in Argos."""
    try:
        import argostranslate.translate as _at
        langs = _at.get_installed_languages()
        src = next((lang for lang in langs if lang.code == from_code), None)
        if src is None:
            return False
        return any(t.to_lang.code == to_code for t in src.translations_from)
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def install_language_pair(from_code: str, to_code: str) -> bool:
    """Install a language pair if not already installed.

    Downloads the appropriate Argos Translate package from the package index
    and installs it. Thread-safe — concurrent calls for the same pair are safe.

    Args:
        from_code: BCP-47 source language code (e.g. ``"en"``).
        to_code:   BCP-47 target language code (e.g. ``"ar"``).

    Returns:
        ``True`` on success, ``False`` if the pair is unavailable or install fails.
    """
    key = (from_code, to_code)

    if key in _installed:
        return True
    if _is_pair_installed(from_code, to_code):
        _installed.add(key)
        return True

    with _install_lock:
        # Double-check after acquiring the lock
        if key in _installed or _is_pair_installed(from_code, to_code):
            _installed.add(key)
            return True

        try:
            import argostranslate.package as _ap

            logger.info("Argos: updating package index…")
            _ap.update_package_index()
            available = _ap.get_available_packages()

            pkg = next(
                (p for p in available if p.from_code == from_code and p.to_code == to_code),
                None,
            )
            if pkg is None:
                logger.warning(
                    "Argos: no package available for %s→%s", from_code, to_code
                )
                return False

            logger.info("Argos: downloading %s→%s package (~100 MB)…", from_code, to_code)
            path = pkg.download()
            _ap.install_from_path(path)
            _installed.add(key)
            logger.info("Argos: installed %s→%s successfully", from_code, to_code)
            return True

        except Exception as exc:
            logger.warning(
                "Argos install failed (%s→%s): %s — will use online fallback",
                from_code, to_code, exc,
            )
            return False


def translate(text: str, from_code: str, to_code: str) -> str:
    """Translate *text* offline using Argos Translate.

    Automatically installs the required language pair on first use if it is not
    already present.

    Args:
        text:      Source text to translate (any length).
        from_code: BCP-47 source language code (e.g. ``"en"``).
        to_code:   BCP-47 target language code (e.g. ``"fr"``).

    Returns:
        Translated text string.

    Raises:
        RuntimeError: If the language pair is not available or translation fails.
    """
    if not text.strip():
        return text

    if not _is_pair_installed(from_code, to_code):
        ok = install_language_pair(from_code, to_code)
        if not ok:
            raise RuntimeError(
                f"Language pair {from_code}→{to_code} is not available in Argos Translate. "
                "Try a different language or use the online translation mode."
            )

    try:
        import argostranslate.translate as _at
        return _at.translate(text, from_code, to_code)
    except Exception as exc:
        raise RuntimeError(f"Argos Translate error: {exc}") from exc


def preinstall_pairs(pairs: list[tuple[str, str]]) -> None:
    """Pre-install language pairs in a background daemon thread (non-blocking).

    Call this once at application startup so that common pairs are ready before
    the first user request arrives.

    Args:
        pairs: List of ``(from_code, to_code)`` tuples to pre-install.

    Note:
        On Render.com free tier, packs are re-downloaded on every dyno restart
        (~100 MB per pair).  This causes a delay before the first translation
        request completes.  Use a Render Disk (paid plan) with ``ARGOS_PACKAGES_DIR``
        set to a persistent path to avoid this overhead.
    """
    def _run() -> None:
        for from_code, to_code in pairs:
            try:
                install_language_pair(from_code, to_code)
            except Exception as exc:
                logger.warning(
                    "Argos preinstall skipped (%s→%s): %s", from_code, to_code, exc
                )

    t = threading.Thread(target=_run, daemon=True, name="argos-preinstall")
    t.start()
    logger.info("Argos: background pre-installation started for %d pair(s)", len(pairs))
