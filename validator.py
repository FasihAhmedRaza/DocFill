"""
validator.py — Pure-Python validation for UAE client records.
No third-party packages: IBAN MOD-97, SWIFT regex, required-field checks.
"""

import re

# Bank English keyword → Arabic name (used to backfill bank_name_ar)
BANK_ARABIC_MAP = {
    "ADCB":                 "بنك أبوظبي التجاري",
    "ABU DHABI COMMERCIAL": "بنك أبوظبي التجاري",
    "RAKBANK":              "بنك رأس الخيمة الوطني",
    "RAS AL KHAIMAH":       "بنك رأس الخيمة الوطني",
    "EMIRATES NBD":         "بنك الإمارات دبي الوطني",
    "DUBAI ISLAMIC":        "بنك دبي الإسلامي",
    "FAB":                  "بنك أبوظبي الأول",
    "FIRST ABU DHABI":      "بنك أبوظبي الأول",
    "MASHREQ":              "بنك المشرق",
    "HSBC":                 "بنك إتش إس بي سي",
    "NOOR":                 "بنك نور",
    "CBD":                  "بنك دبي التجاري",
    "COMMERCIAL BANK OF DUBAI": "بنك دبي التجاري",
}

# Fields the Excel export depends on — used for the "ready" check.
REQUIRED_FIELDS = ["english_name", "arabic_name", "iban", "account_number", "emirate", "bank_name_ar"]

# Exact Emirate strings used by the template's dropdown (Sheet1!B2:B14).
# Writing one of these keeps column G inside the validated list.
EMIRATES = ["Abu Dhabi", "Ajman", "Dubai", "AL Fujairah",
            "Ras al-Khaimah", "Sharjah", "Umm al-Qaiwain"]

SWIFT_RE = re.compile(r"^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$")
EID_RE   = re.compile(r"^784-?\d{4}-?\d{7}-?\d$")


def validate_iban(iban: str) -> bool:
    """UAE/ISO IBAN MOD-97 checksum. Returns True only for a structurally
    valid, checksum-passing IBAN."""
    if not iban:
        return False
    try:
        s = re.sub(r"\s", "", str(iban)).upper()
        if not re.match(r"^[A-Z]{2}\d{2}[A-Z0-9]+$", s):
            return False
        reordered = s[4:] + s[:4]
        numeric = "".join(str(ord(c) - 55) if c.isalpha() else c for c in reordered)
        return int(numeric) % 97 == 1
    except Exception:
        return False


def validate_swift(swift: str) -> bool:
    """SWIFT/BIC must be 8 or 11 alphanumerics (AAAABBCC[XXX])."""
    if not swift:
        return False
    return bool(SWIFT_RE.match(str(swift).strip().upper()))


def validate_emirates_id(eid: str) -> bool:
    """Emirates ID in the 784-XXXX-XXXXXXX-X shape."""
    if not eid:
        return False
    return bool(EID_RE.match(str(eid).strip()))


def normalize_iban(iban: str) -> str:
    """Strip spaces and upper-case an IBAN."""
    return re.sub(r"\s", "", str(iban or "")).upper()


def canonical_emirate(value: str) -> str:
    """Snap a free-typed/extracted emirate to the template's exact dropdown
    string (e.g. 'abu dhabi', 'Abu-Dhabi', 'AbuDhabi' → 'Abu Dhabi';
    'fujairah' → 'AL Fujairah'; 'RAK' → 'Ras al-Khaimah'). Empty/unknown
    values pass through unchanged so we never invent an emirate."""
    if not value:
        return ""
    key = re.sub(r"[\s\-_]+", "", str(value)).lower()
    aliases = {
        "abudhabi": "Abu Dhabi", "ajman": "Ajman", "dubai": "Dubai",
        "fujairah": "AL Fujairah", "alfujairah": "AL Fujairah",
        "rasalkhaimah": "Ras al-Khaimah", "rak": "Ras al-Khaimah",
        "sharjah": "Sharjah",
        "ummalquwain": "Umm al-Qaiwain", "ummalqaiwain": "Umm al-Qaiwain",
        "ummalqaywayn": "Umm al-Qaiwain",
    }
    return aliases.get(key, value)


def bank_arabic_from_english(bank_en: str) -> str:
    """Best-effort lookup of an Arabic bank name from an English one."""
    up = (bank_en or "").upper()
    for kw, arabic in BANK_ARABIC_MAP.items():
        if kw in up:
            return arabic
    return ""


def field_status(record: dict) -> dict:
    """Per-field status for the UI checklist.

    Returns {field: "ok" | "warn"} — "ok" when present (and valid for the
    fields we can validate), "warn" when missing or failing validation.
    """
    status = {}
    for key in [
        "english_name", "arabic_name", "iban", "account_number", "swift",
        "bank_name_ar", "nationality", "emirates_id", "date_of_birth",
        "id_expiry", "emirate", "currency",
    ]:
        val = record.get(key)
        if key == "iban":
            status[key] = "ok" if validate_iban(val) else "warn"
        elif key == "swift":
            status[key] = "ok" if validate_swift(val) else "warn"
        elif key == "emirates_id":
            status[key] = "ok" if validate_emirates_id(val) else "warn"
        else:
            status[key] = "ok" if (val and str(val).strip()) else "warn"
    return status


def count_ok(record: dict) -> int:
    """Number of fields that passed (used for the card badge)."""
    return sum(1 for s in field_status(record).values() if s == "ok")


def is_export_ready(record: dict) -> bool:
    """All Excel-required fields present and IBAN valid."""
    if not validate_iban(record.get("iban", "")):
        return False
    return all(str(record.get(f, "")).strip() for f in REQUIRED_FIELDS)
