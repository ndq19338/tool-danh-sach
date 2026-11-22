from __future__ import annotations

import random
import re
import unicodedata
from datetime import date
from typing import Callable, Optional

# Standard 3-digit Vietnamese mobile network prefixes (10-digit numbers)
VN_MOBILE_PREFIXES: tuple[str, ...] = (
    # Viettel
    "032", "033", "034", "035", "036", "037", "038", "039",
    "086", "096", "097", "098",
    # Vinaphone
    "081", "082", "083", "084", "085", "088", "091", "094",
    # MobiFone
    "070", "076", "077", "078", "079", "089", "090", "093",
    # Vietnamobile / Gmobile
    "056", "058", "092",
)

_MISSING_PHONE_TOKENS = {
    "",
    "-",
    "--",
    "...",
    "…",
    "không",
    "khong",
    "không có",
    "khong co",
    "chưa có",
    "chua co",
    "none",
    "n/a",
}

_DATE_DMY_RE = re.compile(r"^(\d{1,2})\s*[/\-\.]\s*(\d{1,2})\s*[/\-\.]\s*(\d{4})$")
_DATE_YMD_RE = re.compile(r"^(\d{4})\s*[/\-\.]\s*(\d{1,2})\s*[/\-\.]\s*(\d{1,2})$")

_GENDER_MALE = {"nam", "nám", "nãm", "male", "m"}
_GENDER_FEMALE = {"nữ", "nu", "nử", "nư", "nứ", "female", "f"}
_ETHNICITY_KINH = {"kinh", "kịn", "kính", "kình", "kin", "kỉnh"}

_RELIGION_NONE = {
    "không",
    "khong",
    "không có",
    "khong co",
    "0",
    "-",
    "ko",
    "k",
    "none",
    "không theo",
    "khong theo",
    "chưa",
}

_RELIGION_CANONICAL = {
    "công giáo": "Công giáo",
    "cong giao": "Công giáo",
    "phật giáo": "Phật giáo",
    "phat giao": "Phật giáo",
    "tin lành": "Tin lành",
    "tin lanh": "Tin lành",
    "cao đài": "Cao Đài",
    "cao dai": "Cao Đài",
    "hòa hảo": "Hòa Hảo",
    "hoà hảo": "Hòa Hảo",
    "hoa hao": "Hòa Hảo",
    "hồi giáo": "Hồi giáo",
    "hoi giao": "Hồi giáo",
    "thiên chúa": "Thiên Chúa giáo",
    "thiên chúa giáo": "Thiên Chúa giáo",
}

_TRUTHY_FLAGS = {
    "x",
    "có",
    "co",
    "c",
    "1",
    "yes",
    "true",
    "v",
    "✓",
    "✔",
    "đã có",
    "da co",
}

_FALSY_FLAGS = {
    "",
    "không",
    "khong",
    "0",
    "no",
    "false",
    "-",
    "chưa",
    "chưa có",
    "chua",
    "chua co",
    "k",
    "ko",
}

_POSITION_CANONICAL = {
    "bí thư": "Bí thư",
    "bi thu": "Bí thư",
    "phó bí thư": "Phó Bí thư",
    "pho bi thu": "Phó Bí thư",
    "ủy viên": "Ủy viên",
    "uỷ viên": "Ủy viên",
    "uy vien": "Ủy viên",
    "uv": "Ủy viên",
    "ủy viên bch": "Ủy viên BCH",
    "uỷ viên bch": "Ủy viên BCH",
}


def clean_text(text: str) -> str:
    """Normalize Unicode NFC, strip non-breaking spaces/newlines, collapse whitespace."""
    if not text:
        return ""
    s = unicodedata.normalize("NFC", text)
    s = s.replace("\xa0", " ").replace("\u200b", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _capitalize_word(word: str) -> str:
    if not word:
        return ""
    if "-" in word and word != "-":
        return "-".join(_capitalize_word(part) for part in word.split("-"))
    return word[0].upper() + word[1:].lower()


def vn_title_case(text: str, *, is_place: bool = False) -> str:
    """Vietnamese Title Case for names and birthplaces."""
    cleaned = clean_text(text)
    if not cleaned:
        return ""

    if is_place:
        low = cleaned.lower()
        if low in {"tphcm", "tp.hcm", "tp hcm", "tp. hồ chí minh", "tp hồ chí minh"}:
            return "TP. Hồ Chí Minh"
        if low in {"thành phố hồ chí minh", "thanh pho ho chi minh"}:
            return "Thành phố Hồ Chí Minh"

    tokens = cleaned.split(" ")
    out: list[str] = []
    for idx, tok in enumerate(tokens):
        low_tok = tok.lower()
        if is_place:
            if low_tok in {"tp", "tp."}:
                out.append("TP.")
                continue
            if idx == 1 and low_tok == "phố" and out and out[0] == "Thành":
                out.append("phố")
                continue
            if tok in {"-", "–"}:
                out.append("-")
                continue
        out.append(_capitalize_word(tok))
    return " ".join(out)


def normalize_gender(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""
    low = cleaned.lower()
    if low in _GENDER_MALE:
        return "Nam"
    if low in _GENDER_FEMALE:
        return "Nữ"
    return vn_title_case(cleaned)


def normalize_ethnicity(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""
    low = cleaned.lower()
    if low in _ETHNICITY_KINH:
        return "Kinh"
    return vn_title_case(cleaned)


def normalize_religion(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""
    low = cleaned.lower()
    if low in _RELIGION_NONE:
        return "Không"
    if low in _RELIGION_CANONICAL:
        return _RELIGION_CANONICAL[low]
    return cleaned[0].upper() + cleaned[1:].lower()


def normalize_position(text: str) -> str:
    cleaned = clean_text(text)
    if not cleaned:
        return ""
    low = cleaned.lower()
    if low in _POSITION_CANONICAL:
        return _POSITION_CANONICAL[low]
    return cleaned[0].upper() + cleaned[1:]


def parse_truthy_flag(text: str) -> bool:
    cleaned = clean_text(text)
    if not cleaned:
        return False
    low = cleaned.lower()
    if low in _FALSY_FLAGS:
        return False
    if low in _TRUTHY_FLAGS:
        return True
    return True


def format_stat_number(n: int) -> str:
    """Same rule as main project: leading '0' for 1..9, '0' stays '0', >=10 unchanged."""
    return f"0{n}" if 1 <= n <= 9 else str(n)


def format_vn_date(dt: date) -> str:
    """
    Vietnamese administrative date format (same as main project):
    - Day 1..9 has leading '0' (01..09, 10..31)
    - Month 1 & 2 have leading '0' (01, 02); months 3..12 have no leading '0' (3..12)
    """
    day_str = f"{dt.day:02d}"
    month_str = f"{dt.month:02d}" if dt.month in (1, 2) else str(dt.month)
    return f"{day_str}/{month_str}/{dt.year:04d}"


def generate_vn_phone(rng: Optional[random.Random] = None) -> str:
    gen = rng if rng is not None else random
    prefix = gen.choice(VN_MOBILE_PREFIXES)
    suffix = f"{gen.randint(0, 9_999_999):07d}"
    return f"{prefix}{suffix}"


def clean_and_validate_date(
    text: str,
    *,
    is_dob: bool = True,
    min_year: int = 1990,
    max_year: int = 2010,
) -> tuple[str, Optional[str]]:
    cleaned = clean_text(text)
    if not cleaned:
        return "", None

    day = month = year = None
    m_dmy = _DATE_DMY_RE.match(cleaned)
    if m_dmy:
        day, month, year = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
    else:
        m_ymd = _DATE_YMD_RE.match(cleaned)
        if m_ymd:
            year, month, day = int(m_ymd.group(1)), int(m_ymd.group(2)), int(m_ymd.group(3))
        else:
            label = "Ngày sinh" if is_dob else "Ngày vào Đảng/Đoàn"
            return cleaned, f"{label} sai định dạng (cần ngày/tháng/năm)"

    if month < 1 or month > 12:
        return cleaned, f"Tháng không hợp lệ ({month} > 12)" if month > 12 else f"Tháng không hợp lệ ({month})"

    if is_dob:
        if year < min_year or year > max_year:
            return cleaned, f"Năm sinh vô lý ({year}, ngoài khoảng {min_year}–{max_year})"
    else:
        if year < 1995 or year > 2035:
            return cleaned, f"Năm vào Đảng/Đoàn bất thường ({year})"

    try:
        valid_dt = date(year, month, day)
    except ValueError:
        return cleaned, f"Ngày không tồn tại trong lịch ({day:02d}/{month:02d}/{year:04d})"

    return format_vn_date(valid_dt), None


def clean_and_validate_phone(
    text: str,
    *,
    fill_missing: bool = True,
    phone_generator: Optional[Callable[[], str]] = None,
) -> tuple[str, Optional[str]]:
    cleaned = clean_text(text)
    if cleaned.lower() in _MISSING_PHONE_TOKENS:
        if fill_missing:
            gen_fn = phone_generator if phone_generator is not None else generate_vn_phone
            return gen_fn(), None
        return "", "Thiếu số điện thoại (phải đủ 10 số bắt đầu bằng 0)"

    compact = re.sub(r"[\s\.\-\(\)]+", "", cleaned)
    if compact.startswith("+84") and len(compact) == 12 and compact[3:].isdigit():
        compact = "0" + compact[3:]

    if re.fullmatch(r"0\d{9}", compact):
        return compact, None

    if not compact.isdigit():
        return cleaned, "SĐT chứa ký tự không phải chữ số"
    if not compact.startswith("0"):
        return cleaned, f"SĐT không bắt đầu bằng số 0 ({len(compact)} chữ số)"
    return cleaned, f"SĐT có {len(compact)} chữ số (phải đủ 10 số bắt đầu bằng 0)"
