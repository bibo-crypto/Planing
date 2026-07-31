"""
utils.py — Shared utilities, logging setup, and helper functions.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Writable data directory
# ---------------------------------------------------------------------------

def _get_app_data_dir() -> Path:
    """
    Return a directory the app is guaranteed to be able to write to.

    When frozen (built with PyInstaller) and installed under a protected
    location such as ``C:\\Program Files\\...``, a standard (non-admin) user
    account cannot create files there — writing the log file raises
    ``PermissionError: [Errno 13]``. To avoid that, frozen builds write to
    the current user's per-user AppData folder instead, which is always
    writable regardless of where the app itself is installed.

    When running from source (not frozen), logs/settings stay in a local
    "logs"/"settings" folder next to this file, for easy access while
    developing.  Using the file location instead of the current working
    directory keeps the app stable when launched from a shortcut or another
    directory.
    """
    if getattr(sys, "frozen", False):
        base = os.environ.get("LOCALAPPDATA") or str(Path.home())
        return Path(base) / "Delta Dyeing PO Converter"
    return Path(__file__).resolve().parent


APP_DATA_DIR = _get_app_data_dir()

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_DIR = APP_DATA_DIR / "logs"
LOG_FILE = LOG_DIR / "converter.log"


def setup_logging() -> logging.Logger:
    """Configure and return the application-wide logger."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger("converter")
    logger.setLevel(logging.DEBUG)

    if logger.handlers:
        # Already configured (e.g., called a second time in tests)
        return logger

    formatter = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # File handler — always DEBUG
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)

    # Console handler — INFO and above
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger


logger = setup_logging()


# ---------------------------------------------------------------------------
# Text / value helpers
# ---------------------------------------------------------------------------

def clean_text(value: Any) -> str:
    """Strip and normalise whitespace in a string-like value."""
    if value is None:
        return ""
    text = str(value).strip()
    # Collapse internal whitespace to single spaces
    return re.sub(r"\s+", " ", text)


def parse_number(value: Any) -> float | None:
    """
    Try to parse *value* as a float.
    Returns None when the value is not numeric rather than raising.
    """
    text = clean_text(value)
    if not text:
        return None

    # Accept both common formats: 1,234.56 and 1.234,56.  The last
    # separator is treated as the decimal separator when both are present;
    # a lone separator followed by exactly three digits is treated as a
    # thousands separator.
    text = re.sub(r"[^0-9+\-.,]", "", text)
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        tail = text.rsplit(",", 1)[1]
        text = text.replace(",", "") if len(tail) == 3 else text.replace(",", ".")
    elif "." in text:
        tail = text.rsplit(".", 1)[1]
        if len(tail) == 3 and text.count(".") == 1:
            text = text.replace(".", "")
    try:
        return float(text)
    except (ValueError, TypeError):
        return None


def divide_pos(raw_pos: Any) -> int | None:
    """
    Convert a raw POS value (e.g. 10, 20, 30 …) to the display value
    by dividing by 10.  Returns None if the value cannot be parsed.
    """
    num = parse_number(raw_pos)
    if num is None:
        return None
    return int(round(num / 10))


def find_pdfs(path: Path) -> list[Path]:
    """
    Return a sorted list of PDF files found at *path*.
    *path* may be a single PDF file or a directory.
    """
    if path.is_file() and path.suffix.lower() == ".pdf":
        return [path]
    if path.is_dir():
        return sorted(path.glob("*.pdf"))
    return []


def make_output_path(pdf_path: Path, output_dir: Path) -> Path:
    """Derive the Excel output path for a given PDF input path."""
    return output_dir / (pdf_path.stem + ".xlsx")


# ---------------------------------------------------------------------------
# Article description splitting
# ---------------------------------------------------------------------------

# Matches: "<yarn> Nm <nm> Ne <ne> <dye type ...> Lot : <lot>"
# e.g. "100% Cotton Pima Blend Nm 170/2 Ne 100/2 Vat Griege Lot : 52_LOT 1251B.3"
RE_ARTICLE_DESCRIPTION = re.compile(
    r"^(?P<yarn>.*?)\s*"
    r"Nm\s*(?P<nm>\S+)\s*"
    r"Ne\s*(?P<ne>\S+)\s*"
    r"(?P<dye>.+?)\s*"
    r"Lot\s*:?\s*(?P<lot>.+)$",
    re.IGNORECASE,
)

# Colour/state words that ride along next to the dye type (e.g. "Vat
# Griege") but aren't part of the dye type itself, so they're dropped.
ARTICLE_DYE_FILLER_WORDS = {"griege", "grey", "gray", "raw"}


def split_article_description(text: str) -> dict[str, str]:
    """
    Split a raw "Article Description" line into its component fields.

    Expected source pattern::

        <yarn composition> Nm <nm value> Ne <ne value> <dye type> [Griege] Lot : <lot>

    Example
    -------
    >>> split_article_description(
    ...     "100% Cotton Pima Blend Nm 170/2 Ne 100/2 Vat Griege Lot : 52_LOT 1251B.3"
    ... )
    {'yarn': '100% Cotton Pima Blend', 'nm': '170/2', 'ne': '100/2',
     'dye_type': 'Vat', 'lot': '52_LOT 1251B.3'}

    If the text doesn't match the expected pattern (unexpected PDF layout,
    OCR noise, etc.), all fields come back empty so the caller can fall
    back to showing the original, unsplit text — nothing is guessed.
    """
    empty = {"yarn": "", "nm": "", "ne": "", "dye_type": "", "lot": ""}

    text = clean_text(text)
    if not text:
        return empty

    match = RE_ARTICLE_DESCRIPTION.match(text)
    if not match:
        return empty

    dye_words = [
        w for w in match.group("dye").split()
        if w.lower() not in ARTICLE_DYE_FILLER_WORDS
    ]
    dye_type = " ".join(dye_words) if dye_words else match.group("dye").strip()

    return {
        "yarn": match.group("yarn").strip(),
        "nm": match.group("nm").strip(),
        "ne": match.group("ne").strip(),
        "dye_type": dye_type,
        "lot": match.group("lot").strip(),
    }


# ---------------------------------------------------------------------------
# Persistent settings (simple JSON file next to the application)
# ---------------------------------------------------------------------------

SETTINGS_FILE = APP_DATA_DIR / "settings" / "prefs.json"

def load_settings() -> dict[str, object]:
    """
    Load persisted settings from ``settings/prefs.json``.
    Returns an empty mapping when no settings file exists yet.

    Paths and checkbox values are optional, so callers already handle missing
    keys safely with ``dict.get``.  Keeping this function free of obsolete
    feature defaults prevents settings for removed UI controls from growing
    back on every launch.
    """
    settings: dict[str, object] = {}
    try:
        if SETTINGS_FILE.is_file():
            on_disk = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if isinstance(on_disk, dict):
                settings.update(on_disk)
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not load settings: %s", exc)
    return settings


def save_settings(settings: dict[str, object]) -> None:
    """
    Persist *settings* to ``settings/prefs.json``.
    Creates the directory if needed; silently skips on error.
    """
    try:
        SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(
            json.dumps(settings, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception as exc:  # noqa: BLE001
        logger.warning("Could not save settings: %s", exc)
