from __future__ import annotations

import math
import re
import unicodedata
from typing import Any

from .constants import EMPTY_SOURCE_MARKERS


def display_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\u200f", "").replace("\u200e", "").strip()


def normalized_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", display_text(value))
    text = text.translate(str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ى": "ي"}))
    return re.sub(r"\s+", " ", text).casefold()


def normalized_header(value: Any) -> str:
    """Compare headers flexibly: spaces, underscores, and punctuation are not significant."""
    header = normalized_text(value).replace("_", " ").replace("ة", "ه")
    header = re.sub(r"[^\w\u0600-\u06FF]+", " ", header)
    return re.sub(r"\s+", " ", header).strip()

def is_blank_source(value: Any) -> bool:
    return (
        value is None
        or (isinstance(value, float) and math.isnan(value))
        or display_text(value).casefold() in EMPTY_SOURCE_MARKERS
    )


def normalize_parcel_code(value: Any) -> str | None:
    if is_blank_source(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    result = display_text(value).translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
    if re.fullmatch(r"[+-]?\d+\.0+", result):
        result = result.split(".", 1)[0]
    return result or None


def values_equal(left: Any, right: Any) -> bool:
    if is_blank_source(left) and is_blank_source(right):
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return float(left) == float(right)
    return display_text(left) == display_text(right)
