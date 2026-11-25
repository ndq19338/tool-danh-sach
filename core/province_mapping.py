"""
Bảng ánh xạ quy đổi tên tỉnh/thành phố cũ sang 34 tỉnh/thành phố trực thuộc Trung ương
của Việt Nam theo Nghị quyết số 202/2025/QH15 (ngày 12/6/2025) và các đợt sáp nhập trước đây
(như Hà Tây -> Hà Nội).

File được tách riêng để dễ dàng kiểm tra, đối chiếu và chỉnh sửa khi cần.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Optional

from .cleaner import clean_text

# ==============================================================================
# 1. DANH SÁCH CHÍNH THỨC 34 TỈNH/THÀNH PHỐ SAU SÁP NHẬP (NGHỊ QUYẾT 202/2025/QH15)
#    (11 địa phương giữ nguyên + 23 địa phương hình thành sau sáp nhập từ 52 tỉnh/TP cũ)
# ==============================================================================
NEW_PROVINCES_34: tuple[str, ...] = (
    # --- 11 tỉnh/thành phố giữ nguyên ---
    "Hà Nội",
    "Huế",
    "Lai Châu",
    "Điện Biên",
    "Sơn La",
    "Lạng Sơn",
    "Quảng Ninh",
    "Thanh Hóa",
    "Nghệ An",
    "Hà Tĩnh",
    "Cao Bằng",
    # --- 23 tỉnh/thành phố hình thành sau sáp nhập ---
    "Tuyên Quang",
    "Lào Cai",
    "Thái Nguyên",
    "Phú Thọ",
    "Bắc Ninh",
    "Hải Phòng",
    "Hưng Yên",
    "Ninh Bình",
    "Quảng Trị",
    "Đà Nẵng",
    "Quảng Ngãi",
    "Gia Lai",
    "Khánh Hòa",
    "Lâm Đồng",
    "Đắk Lắk",
    "TP. Hồ Chí Minh",
    "Đồng Nai",
    "Tây Ninh",
    "Cần Thơ",
    "Vĩnh Long",
    "Đồng Tháp",
    "An Giang",
    "Cà Mau",
)

# ==============================================================================
# 2. BẢNG ÁNH XẠ ĐẦY ĐỦ: TỈNH/THÀNH CŨ -> TỈNH/THÀNH MỚI (34 TỈNH/TP)
#    Bao gồm toàn bộ 63 tỉnh/thành trước 2025 + các tỉnh sáp nhập lịch sử không gây mơ hồ.
# ==============================================================================
PROVINCE_MAPPING: dict[str, str] = {
    # --- 11 địa phương giữ nguyên ---
    "Hà Nội": "Hà Nội",
    "Huế": "Huế",
    "Thừa Thiên Huế": "Huế",
    "Lai Châu": "Lai Châu",
    "Điện Biên": "Điện Biên",
    "Sơn La": "Sơn La",
    "Lạng Sơn": "Lạng Sơn",
    "Quảng Ninh": "Quảng Ninh",
    "Thanh Hóa": "Thanh Hóa",
    "Nghệ An": "Nghệ An",
    "Hà Tĩnh": "Hà Tĩnh",
    "Cao Bằng": "Cao Bằng",

    # --- 23 địa phương sau sáp nhập (52 tỉnh/TP cũ -> 23 tỉnh/TP mới) ---
    # 1. Tuyên Quang = Tuyên Quang + Hà Giang
    "Tuyên Quang": "Tuyên Quang",
    "Hà Giang": "Tuyên Quang",

    # 2. Lào Cai = Lào Cai + Yên Bái
    "Lào Cai": "Lào Cai",
    "Yên Bái": "Lào Cai",

    # 3. Thái Nguyên = Thái Nguyên + Bắc Kạn
    "Thái Nguyên": "Thái Nguyên",
    "Bắc Kạn": "Thái Nguyên",

    # 4. Phú Thọ = Phú Thọ + Vĩnh Phúc + Hòa Bình
    "Phú Thọ": "Phú Thọ",
    "Vĩnh Phúc": "Phú Thọ",
    "Hòa Bình": "Phú Thọ",

    # 5. Bắc Ninh = Bắc Ninh + Bắc Giang
    "Bắc Ninh": "Bắc Ninh",
    "Bắc Giang": "Bắc Ninh",

    # 6. Hải Phòng = Hải Phòng + Hải Dương
    "Hải Phòng": "Hải Phòng",
    "Hải Dương": "Hải Phòng",

    # 7. Hưng Yên = Hưng Yên + Thái Bình
    "Hưng Yên": "Hưng Yên",
    "Thái Bình": "Hưng Yên",

    # 8. Ninh Bình = Ninh Bình + Hà Nam + Nam Định
    "Ninh Bình": "Ninh Bình",
    "Hà Nam": "Ninh Bình",
    "Nam Định": "Ninh Bình",

    # 9. Quảng Trị = Quảng Trị + Quảng Bình
    "Quảng Trị": "Quảng Trị",
    "Quảng Bình": "Quảng Trị",

    # 10. Đà Nẵng = Đà Nẵng + Quảng Nam
    "Đà Nẵng": "Đà Nẵng",
    "Quảng Nam": "Đà Nẵng",

    # 11. Quảng Ngãi = Quảng Ngãi + Kon Tum
    "Quảng Ngãi": "Quảng Ngãi",
    "Kon Tum": "Quảng Ngãi",

    # 12. Gia Lai = Gia Lai + Bình Định
    "Gia Lai": "Gia Lai",
    "Bình Định": "Gia Lai",

    # 13. Khánh Hòa = Khánh Hòa + Ninh Thuận
    "Khánh Hòa": "Khánh Hòa",
    "Ninh Thuận": "Khánh Hòa",

    # 14. Lâm Đồng = Lâm Đồng + Đắk Nông + Bình Thuận
    "Lâm Đồng": "Lâm Đồng",
    "Đắk Nông": "Lâm Đồng",
    "Bình Thuận": "Lâm Đồng",

    # 15. Đắk Lắk = Đắk Lắk + Phú Yên
    "Đắk Lắk": "Đắk Lắk",
    "Phú Yên": "Đắk Lắk",

    # 16. TP. Hồ Chí Minh = TP. Hồ Chí Minh + Bình Dương + Bà Rịa - Vũng Tàu
    "TP. Hồ Chí Minh": "TP. Hồ Chí Minh",
    "Bình Dương": "TP. Hồ Chí Minh",
    "Bà Rịa - Vũng Tàu": "TP. Hồ Chí Minh",

    # 17. Đồng Nai = Đồng Nai + Bình Phước
    "Đồng Nai": "Đồng Nai",
    "Bình Phước": "Đồng Nai",

    # 18. Tây Ninh = Tây Ninh + Long An
    "Tây Ninh": "Tây Ninh",
    "Long An": "Tây Ninh",

    # 19. Cần Thơ = Cần Thơ + Sóc Trăng + Hậu Giang
    "Cần Thơ": "Cần Thơ",
    "Sóc Trăng": "Cần Thơ",
    "Hậu Giang": "Cần Thơ",

    # 20. Vĩnh Long = Vĩnh Long + Bến Tre + Trà Vinh
    "Vĩnh Long": "Vĩnh Long",
    "Bến Tre": "Vĩnh Long",
    "Trà Vinh": "Vĩnh Long",

    # 21. Đồng Tháp = Đồng Tháp + Tiền Giang
    "Đồng Tháp": "Đồng Tháp",
    "Tiền Giang": "Đồng Tháp",

    # 22. An Giang = An Giang + Kiên Giang
    "An Giang": "An Giang",
    "Kiên Giang": "An Giang",

    # 23. Cà Mau = Cà Mau + Bạc Liêu
    "Cà Mau": "Cà Mau",
    "Bạc Liêu": "Cà Mau",

    # --- Các tỉnh lịch sử đã sáp nhập trước 2025 (không mơ hồ) ---
    "Hà Tây": "Hà Nội",
    "Hà Bắc": "Bắc Ninh",
    "Hà Nam Ninh": "Ninh Bình",
    "Vĩnh Phú": "Phú Thọ",
    "Hà Tuyên": "Tuyên Quang",
    "Hoàng Liên Sơn": "Lào Cai",
    "Bắc Thái": "Thái Nguyên",
    "Cửu Long": "Vĩnh Long",
    "Minh Hải": "Cà Mau",
}

# ==============================================================================
# 3. CÁC BIẾN THỂ VIẾT TẮT / LỖI CHÍNH TẢ PHỔ BIẾN -> TÊN CHUẨN TRƯỚC QUY ĐỔI
#    (Keys được viết dạng chữ thường không dấu, đã bỏ tiền tố Tỉnh/TP.)
# ==============================================================================
PROVINCE_ALIASES: dict[str, str] = {
    "ho chi minh": "TP. Hồ Chí Minh",
    "hcm": "TP. Hồ Chí Minh",
    "tphcm": "TP. Hồ Chí Minh",
    "tp hcm": "TP. Hồ Chí Minh",
    "sai gon": "TP. Hồ Chí Minh",
    "hn": "Hà Nội",
    "hp": "Hải Phòng",
    "thua thien hue": "Thừa Thiên Huế",
    "tt hue": "Thừa Thiên Huế",
    "ba ria vung tau": "Bà Rịa - Vũng Tàu",
    "br vt": "Bà Rịa - Vũng Tàu",
    "brvt": "Bà Rịa - Vũng Tàu",
    "vung tau": "Bà Rịa - Vũng Tàu",
    "dac lac": "Đắk Lắk",
    "dac lak": "Đắk Lắk",
    "dak lac": "Đắk Lắk",
    "daklak": "Đắk Lắk",
    "dac nong": "Đắk Nông",
    "daknong": "Đắk Nông",
    "bac can": "Bắc Kạn",
    "kontum": "Kon Tum",
    "khanh hoaa": "Khánh Hòa",
}

# Tiền tố hành chính thường bị thừa khi sinh viên nhập Nơi sinh
_ADMIN_PREFIX_RE = re.compile(
    r"^(?:thành\s+phố|thanh\s+pho|tp\.?|tỉnh|tinh|t\.)\s+",
    re.IGNORECASE,
)

# Chuẩn hóa vị trí dấu kiểu cũ -> kiểu mới (vd: Hoá -> Hóa, Hoà -> Hòa, Thuỷ -> Thủy)
_OLD_TONE_FIXES: tuple[tuple[str, str], ...] = (
    ("oà", "òa"), ("oá", "óa"), ("oả", "ỏa"), ("oã", "õa"), ("oạ", "ọa"),
    ("uỳ", "ùy"), ("uý", "úy"), ("uỷ", "ủy"), ("uỹ", "ũy"), ("uỵ", "ụy"),
    ("oè", "òe"), ("oé", "óe"), ("oẻ", "ỏe"), ("oẽ", "õe"), ("oẹ", "ọe"),
)


def _fold_vn_key(text: str) -> str:
    """
    Chuyển chuỗi tiếng Việt sang khóa so khớp chữ thường, bỏ dấu thanh & dấu phụ,
    bỏ tiền tố hành chính (Tỉnh / TP. / Thành phố) và chuẩn hóa dấu gạch nối.
    Ví dụ:
      - 'Tỉnh Hải Dương'         -> 'hai duong'
      - 'Thanh Hoá'              -> 'thanh hoa'
      - 'Haỉ Phòng'              -> 'hai phong'
      - 'Thành phố Hồ Chí Minh'  -> 'ho chi minh'
      - 'Bà Rịa – Vũng Tàu'      -> 'ba ria vung tau'
    """
    s = clean_text(text)
    if not s:
        return ""

    # Lặp tối đa 2 lần phòng trường hợp viết "Tỉnh TP. ..." hoặc dư dấu chấm
    for _ in range(2):
        s_new = _ADMIN_PREFIX_RE.sub("", s).strip()
        if s_new == s:
            break
        s = s_new

    s = s.lower()
    for old_seq, new_seq in _OLD_TONE_FIXES:
        s = s.replace(old_seq, new_seq)

    # Thay đ -> d trước khi tách NFD
    s = s.replace("đ", "d")
    nfd = unicodedata.normalize("NFD", s)
    no_marks = "".join(ch for ch in nfd if unicodedata.category(ch) != "Mn")
    # Thay mọi ký tự phân tách (-, –, ., _, /) bằng khoảng trắng
    no_sep = re.sub(r"[^a-z0-9]+", " ", no_marks)
    return re.sub(r"\s+", " ", no_sep).strip()


def _build_lookup_index() -> dict[str, str]:
    """
    Xây dựng bảng tra cứu nhanh: folded_key -> tên chuẩn trước sáp nhập (key trong PROVINCE_MAPPING).
    Đảm bảo toàn bộ 63 tỉnh/TP + các tỉnh lịch sử + alias đều có khóa duy nhất.
    """
    index: dict[str, str] = {}
    for canonical_old in PROVINCE_MAPPING:
        key = _fold_vn_key(canonical_old)
        index[key] = canonical_old
    for alias_key, canonical_old in PROVINCE_ALIASES.items():
        index[_fold_vn_key(alias_key)] = canonical_old
    return index


_FOLDED_LOOKUP: dict[str, str] = _build_lookup_index()


def normalize_birthplace(
    raw_text: str,
    *,
    apply_merger: bool = True,
) -> tuple[str, Optional[str]]:
    """
    Chuẩn hóa và (tùy chọn) quy đổi cột Nơi sinh sang tên tỉnh/thành phố mới sau sáp nhập.

    Args:
        raw_text: Giá trị gốc ở cột Nơi sinh.
        apply_merger: Nếu True, quy đổi tỉnh/thành cũ sang tên mới (34 tỉnh/TP + Hà Tây -> Hà Nội).
                      Nếu False, chỉ chuẩn hóa chính tả/dấu thanh/tiền tố về tên tỉnh/TP chuẩn.

    Returns:
        (output_value, warning_reason)
        - Nếu nhận diện được tỉnh/thành hợp lệ: (tên chuẩn mới/cũ, None).
        - Nếu ô trống: ("", None).
        - Nếu không khớp tỉnh/thành nào (sai chính tả nặng, ghi địa chỉ xã/huyện, nước ngoài...):
          giữ nguyên text gốc (đã trim) và trả về lý do cảnh báo để ghi vào Canh_bao_loi.
    """
    cleaned = clean_text(raw_text)
    if not cleaned:
        return "", None

    folded = _fold_vn_key(cleaned)
    canonical_old = _FOLDED_LOOKUP.get(folded)

    if canonical_old is None:
        return (
            cleaned,
            "Không nhận diện được tỉnh/thành phố hợp lệ (sai chính tả hoặc ghi chi tiết cấp xã/huyện)",
        )

    if apply_merger:
        return PROVINCE_MAPPING[canonical_old], None

    # Khi tắt chế độ quy đổi sáp nhập: giữ tên chuẩn trước sáp nhập
    return canonical_old, None
