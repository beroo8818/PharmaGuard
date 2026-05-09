import re


def clean_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text)


def slug(text):
    text = clean_text(text).lower()
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = re.sub(r"[^a-z0-9\u0600-\u06ff]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def contains_arabic(text_value):
    s = str(text_value)
    return any(
        "\u0600" <= ch <= "\u06FF"
        or "\u0750" <= ch <= "\u077F"
        or "\u08A0" <= ch <= "\u08FF"
        for ch in s
    )


def normalize_search_text(text_value):
    s = "" if text_value is None else str(text_value)
    s = s.strip().lower()
    if not s:
        return ""

    s = re.sub(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED\u0640]", "", s)

    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ى": "ي",
        "ئ": "ي",
        "ؤ": "و",
        "ة": "ه",
    }

    for old, new in replacements.items():
        s = s.replace(old, new)

    s = re.sub(r"[^\w\s\u0600-\u06FF]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s
