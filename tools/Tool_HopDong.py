# -*- coding: utf-8 -*-
"""
Tool tạo Hợp đồng tín dụng GQVL (Giải quyết việc làm) từ file mẫu Word.
- Nhập thông tin khách hàng -> tự tính: bằng chữ, hạn trả nợ cuối cùng, phân kỳ, lãi quá hạn
- Điền vào file mẫu templates/hopdongmau.docx (giữ nguyên định dạng của mẫu)
- Tạo hàng loạt: nhập từ Excel hoặc thêm từng khách vào danh sách rồi xuất tất cả cùng lúc
- Có thể lấy thông tin khách hàng từ SAO KÊ dùng chung (nạp ở launcher)

Thư viện: pandas, openpyxl, lxml (đã có khi cài python-docx)
"""
import json
import logging
import os
import re
import sys
import zipfile
from calendar import monthrange
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import pandas as pd
from lxml import etree

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from . import _ui as ui
    from . import _paths as paths
    from ._shared import store, get_field, cell_to_str, clean_header, row_empty, remove_accents, \
        K_TEN, K_CCCD, K_NGAYSINH, K_NGAYCAP, K_NOICAP, K_SDT, K_XA, K_DIACHI
except ImportError:          # chạy trực tiếp file này, không qua launcher
    import _ui as ui
    import _paths as paths
    from _shared import store, get_field, cell_to_str, clean_header, row_empty, remove_accents, \
        K_TEN, K_CCCD, K_NGAYSINH, K_NGAYCAP, K_NOICAP, K_SDT, K_XA, K_DIACHI

BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
CONFIG_SECTION = "hopdong"       # mục cấu hình của tool này trong config.json
# File mẫu: ưu tiên bản đặt cạnh file exe (sửa mẫu không cần đóng gói lại), sau đó đến bản đóng gói sẵn bên trong exe
if getattr(sys, "frozen", False):
    RESOURCE_DIR = os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), "tools")
else:
    RESOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_CANDIDATES = [os.path.join(BASE_DIR, "templates", "hopdongmau.docx"),
                       os.path.join(RESOURCE_DIR, "templates", "hopdongmau.docx")]

# ---------- Giá trị mặc định khi để trống ----------
DOT_TEN_TK = "." * 87
DOT_STK = "." * 31
DOT_NGAN_HANG = "." * 31
DOT_CHI_NHANH = "." * 30
DOT_UY_QUYEN = "." * 80
DOT_SHORT = "." * 15
DOT_DATE = "...../...../........."

NOI_CAP = ["Cục CSQLHC về TTXH", "Bộ Công an"]
LAI_SUAT = ["7,9248", "6,24"]
CHUC_VU = ["Giám đốc", "Phó Giám đốc"]
COT_THON = 7                                        # cột H của sao kê: Tên thôn
LOAI_GT_MAC_DINH = "Căn cước công dân số:"          # ô [loaigt] trong mẫu (không còn nhập tay)
NGAY_DOI_NOI_CAP = date(2024, 7, 1)                 # cấp từ ngày này về trước: Cục CSQLHC; sau đó: Bộ Công an

# 34 tỉnh, thành phố sau sáp nhập (Nghị quyết 202/2025/QH15)
TINH_LIST = [
    "An Giang", "Bắc Ninh", "Cà Mau", "Cao Bằng", "Cần Thơ", "Đà Nẵng", "Đắk Lắk", "Điện Biên",
    "Đồng Nai", "Đồng Tháp", "Gia Lai", "Hà Nội", "Hà Tĩnh", "Hải Phòng", "Hưng Yên", "Huế",
    "Khánh Hòa", "Lai Châu", "Lâm Đồng", "Lạng Sơn", "Lào Cai", "Nghệ An", "Ninh Bình", "Phú Thọ",
    "Quảng Ngãi", "Quảng Ninh", "Quảng Trị", "Sơn La", "Tây Ninh", "Thái Nguyên", "Thanh Hóa",
    "Thành phố Hồ Chí Minh", "Tuyên Quang", "Vĩnh Long",
]
assert len(TINH_LIST) == 34
HE_SO_QUA_HAN = Decimal("1.3")        # lãi suất quá hạn = 130% lãi suất cho vay


# ================== TÍNH TOÁN (không phụ thuộc giao diện) ==================
def parse_date(text):
    """Nhận 'd/m/yyyy', 'dd-mm-yyyy', 'ddmmyyyy'... -> date. Lỗi -> ValueError."""
    s = (text or "").strip()
    iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$", s)       # 1992-07-09 (có thể kèm giờ)
    if iso:
        return date(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
    s = s.replace("-", "/").replace(".", "/").replace(" ", "")
    if re.fullmatch(r"\d{8}", s):
        s = f"{s[:2]}/{s[2:4]}/{s[4:]}"
    return datetime.strptime(s, "%d/%m/%Y").date()


def fmt_date(d):
    return d.strftime("%d/%m/%Y")


def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y, m = d.year + y, m + 1
    return date(y, m, min(d.day, monthrange(y, m)[1]))


# ---------- Định dạng khi nhập ----------
def fmt_name(s):
    """Viết hoa chữ cái đầu mỗi từ, còn lại viết thường (kể cả dữ liệu sao kê viết hoa hết)."""
    s = re.sub(r"\s{2,}", " ", (s or "").lstrip())
    return re.sub(r"\S+", lambda m: m.group(0)[:1].upper() + m.group(0)[1:].lower(), s)


def fmt_date_input(s):
    """Gõ liền 19112026 hoặc 19/11/2026 -> tự chèn dấu '/' theo dd/mm/yyyy; '5/1/1990' -> '05/01/1990'."""
    s = (s or "").strip()
    iso = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?", s)
    if iso:
        return f"{int(iso.group(3)):02d}/{int(iso.group(2)):02d}/{iso.group(1)}"
    if re.search(r"\D", s):
        parts = re.split(r"\D+", s)
        toks = []
        for i, p in enumerate(parts[:3]):
            if p == "":
                continue
            if i < 2 and len(p) == 1 and i < len(parts) - 1:
                p = "0" + p                          # "5/" -> "05/"
            toks.append(p[:4] if i == 2 else p[:2])
        out = "/".join(toks)
        if parts[-1] == "" and 0 < len(toks) < 3:
            out += "/"
        return out
    d = s[:8]
    out = d[:2]
    if len(d) > 2:
        out += "/" + d[2:4]
    if len(d) > 4:
        out += "/" + d[4:]
    return out


def fmt_digits(max_len):
    """Chỉ cho nhập số, tối đa max_len chữ số."""
    return lambda s: re.sub(r"\D", "", s or "")[:max_len]


def fmt_money_input(s):
    n = re.sub(r"\D", "", s or "")
    return fmt_money(int(n)) if n else ""


def suggest_noi_cap(ngay_cap_text):
    """Gợi ý nơi cấp theo ngày cấp: từ 01/07/2024 về trước -> Cục CSQLHC về TTXH, sau đó -> Bộ Công an."""
    try:
        d = parse_date(ngay_cap_text)
    except ValueError:
        return ""
    return NOI_CAP[0] if d <= NGAY_DOI_NOI_CAP else NOI_CAP[1]


def _tinh_key(text):
    t = re.sub(r"(?i)^\s*(tỉnh|thành phố|tp\.?)\s+", "", text or "")
    return remove_accents(t)


def normalize_tinh(text):
    """'tuyên quang', 'Tỉnh Tuyên Quang', 'TP HCM'... -> tên chuẩn trong TINH_LIST; không khớp thì trả nguyên văn."""
    text = (text or "").strip()
    key = _tinh_key(text)
    if key in ("tphcm", "hcm", "saigon", "hochiminh"):
        return "Thành phố Hồ Chí Minh"
    for t in TINH_LIST:
        if _tinh_key(t) == key:
            return t
    return text


def clean_pgd(text):
    """Bỏ chữ 'PGD' / 'Phòng giao dịch' trong tên đơn vị lấy từ sao kê, viết hoa chữ cái đầu."""
    s = re.sub(r"(?i)\bphòng giao dịch\b", " ", text or "")
    s = re.sub(r"(?i)\bp\.?\s?gd\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" -:–.,")
    return fmt_name(s)


def pgd_from_sao_ke():
    """Tên PGD lấy từ cột D (số 3) của sao kê dùng chung, đã bỏ chữ PGD."""
    for i in range(min(50, len(store.records))):
        val = store.rows[i][3] if store.has_positions and len(store.rows[i]) > 3 else ""
        if not val:
            val = get_field(store.records[i], "Tên PGD", "PGD")
        val = clean_pgd(val)
        if val:
            return val
    return ""


def fmt_ky_line(ngay, tien):
    return f"Ngày: {fmt_date(ngay)} Số tiền: {fmt_money(tien)} đồng"


def parse_money(text):
    digits = re.sub(r"\D", "", text or "")
    return int(digits) if digits else 0


def fmt_money(n):
    return f"{int(n):,}".replace(",", ".")


def so_thanh_chu(so):
    so = int(so)
    if so == 0:
        return "Không đồng chẵn."
    chu = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    don_vi = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]
    groups, n = [], so
    while n > 0:
        groups.append(n % 1000)
        n //= 1000
    if len(groups) > len(don_vi):
        raise ValueError("Số tiền quá lớn để đọc thành chữ.")

    def doc3(n, day_du):
        tram, chuc, dv = n // 100, (n % 100) // 10, n % 10
        out = []
        if tram > 0 or day_du:
            out.append(chu[tram] + " trăm")
        if chuc > 1:
            out.append(chu[chuc] + " mươi")
            if dv == 1:
                out.append("mốt")
            elif dv == 5:
                out.append("lăm")
            elif dv > 0:
                out.append(chu[dv])
        elif chuc == 1:
            out.append("mười")
            if dv == 5:
                out.append("lăm")
            elif dv > 0:
                out.append(chu[dv])
        elif dv > 0:
            if tram > 0 or day_du:
                out.append("lẻ")
            out.append(chu[dv])
        return " ".join(out)

    parts = []
    for i in range(len(groups) - 1, -1, -1):
        if groups[i] == 0:
            continue
        parts.append((doc3(groups[i], i < len(groups) - 1) + " " + don_vi[i]).strip())
    text = " ".join(parts)
    return text[:1].upper() + text[1:] + " đồng chẵn."


def ky_han_months(han_vay):
    """Phân kỳ tự động: <=12 tháng -> 12; <=60 tháng -> 24; <=120 tháng -> 36."""
    if han_vay <= 12:
        return 12
    if han_vay <= 60:
        return 24
    return 36


def build_schedule(ngay_vay, han_vay, tong):
    """Danh sách (ngày trả, số tiền) các kỳ trả nợ gốc. Kỳ cuối rơi đúng hạn cuối cùng.
    Số tiền chia đều các kỳ, làm tròn nghìn đồng, phần lẻ dồn vào kỳ cuối."""
    step = ky_han_months(han_vay)
    months = list(range(step, han_vay, step)) + [han_vay]
    n = len(months)
    base = (tong // n // 1000) * 1000
    amounts = [base] * (n - 1) + [tong - base * (n - 1)]
    return [(add_months(ngay_vay, m), a) for m, a in zip(months, amounts)]


def lai_suat_qua_han(ls_text):
    try:
        v = Decimal(ls_text.replace(",", ".")) * HE_SO_QUA_HAN
    except (InvalidOperation, AttributeError):
        return ""
    return format(v.normalize(), "f").replace(".", ",")


def default_text(value, default):
    value = (value or "").strip()
    return value if value else default


# ================== ĐIỀN FILE WORD ==================
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


RED = {"FF0000", "EE0000", "C00000", "FF0000"}


def _black(t):
    """Mẫu tô đỏ các ô [..]; sau khi điền thì đổi về chữ đen."""
    run = t.getparent()
    rpr = run.find(f"{{{W}}}rPr") if run is not None else None
    color = rpr.find(f"{{{W}}}color") if rpr is not None else None
    if color is not None and (color.get(f"{{{W}}}val") or "").upper() in RED:
        color.set(f"{{{W}}}val", "000000")


def _replace_in_paragraph(p, pattern, mapping):
    """Thay các [placeholder] kể cả khi Word tách chúng ra nhiều run / nằm trong content control."""
    ts = list(p.iter(f"{{{W}}}t"))
    if not ts:
        return
    texts = [t.text or "" for t in ts]
    full = "".join(texts)
    if "[" not in full:
        return
    matches = list(pattern.finditer(full))
    if not matches:
        return

    owner = []
    for i, tx in enumerate(texts):
        owner += [i] * len(tx)
    new = [""] * len(ts)
    pos = 0
    for m in matches:
        for k in range(pos, m.start()):
            new[owner[k]] += full[k]
        new[owner[m.start()]] += mapping[m.group(0)]
        pos = m.end()
    for k in range(pos, len(full)):
        new[owner[k]] += full[k]

    for t, old, txt in zip(ts, texts, new):
        if txt == old:
            continue
        _black(t)
        parts = txt.split("\n")
        t.text = parts[0]
        t.set(XML_SPACE, "preserve")
        prev = t
        for part in parts[1:]:                     # xuống dòng trong cùng đoạn
            br = etree.Element(f"{{{W}}}br")
            prev.addnext(br)
            t2 = etree.Element(f"{{{W}}}t")
            t2.text = part
            t2.set(XML_SPACE, "preserve")
            br.addnext(t2)
            prev = t2


def fill_template(template_path, out_path, mapping):
    """Điền mapping vào mẫu, ghi ra out_path. Trả về list placeholder còn sót (nếu có)."""
    keys = sorted(mapping, key=len, reverse=True)
    pattern = re.compile("|".join(re.escape(k) for k in keys))
    with zipfile.ZipFile(template_path) as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        for p in root.iter(f"{{{W}}}p"):
            _replace_in_paragraph(p, pattern, mapping)
        new_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        leftover = sorted(set(re.findall(r"\[[A-Za-z0-9_/]+\]", "".join(
            t.text or "" for t in root.iter(f"{{{W}}}t")))))
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = new_xml if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)
    return leftover


def build_mapping(d):
    """d: dict đã kiểm tra hợp lệ -> dict {placeholder: nội dung}."""
    lap, vay = d["ngay_lap"], d["ngay_vay"]
    han_cuoi = add_months(vay, d["han_vay"])
    schedule = build_schedule(vay, d["han_vay"], d["tong"])
    m = {
        "[sohopdong]": d["so_hd"],
        "[namlaphd]": str(lap.year),
        "[ngaylaphdchu]": f"ngày {lap.day:02d} tháng {lap.month:02d} năm {lap.year}",
        "[ngaylaphd]": fmt_date(lap),
        "[pgd]": default_text(d["pgd"], DOT_SHORT),
        "[diachipgd]": default_text(d["dia_chi_pgd"], DOT_SHORT),
        "[sdtpgd]": default_text(d["sdt_pgd"], DOT_SHORT),
        "[lanhdao]": d["lanh_dao"],
        "[chucvu]": d["chuc_vu"],
        "[giayuyquyen]": default_text(d["uy_quyen"], DOT_UY_QUYEN),
        "[hotenhv]": d["ho_ten"],
        "[namsinhhv]": fmt_date(d["ngay_sinh"]),
        "[loaigt]": d["loai_gt"],
        "[cccd/cmt]": d["cccd"],
        "[ngaycap]": fmt_date(d["ngay_cap"]) if d["ngay_cap"] else DOT_DATE,
        "[noicap]": default_text(d["noi_cap"], DOT_SHORT),
        "[ngayhethancccd]": default_text(d["het_han"], DOT_SHORT),
        # Mẫu đã có sẵn chữ "Thôn (tổ)" / "xã (phường)" nên bỏ chữ Thôn / Xã / Phường đứng đầu giá trị
        "[thon]": default_text(re.sub(r"(?i)^\s*thôn\s+", "", d["thon"]), DOT_SHORT),
        "[xa]": default_text(re.sub(r"(?i)^\s*(xã|phường)\s+", "", d["xa"]), DOT_SHORT),
        "[tinh]": default_text(d["tinh"], DOT_SHORT),
        "[dienthoaihv]": default_text(d["dien_thoai"], DOT_SHORT),
        # Mẫu viết "[tongvaycatso].000.000": thay cả cụm bằng số tiền đầy đủ
        "[tongvaycatso].000.000": fmt_money(d["tong"]),
        "[tongvaycatso]": fmt_money(d["tong"]),
        "[tongtienvaychu]": so_thanh_chu(d["tong"]),
        "[sotaovieclam]": str(d["so_ld"]),
        "[hanvay]": str(d["han_vay"]),
        "[hancuoi]": fmt_date(han_cuoi),
        "[ngayvay]": fmt_date(vay),
        "[laisuat]": d["lai_suat"],
        "[laisuatqh]": d["lai_qh"],
        "[tennguoinhantien]": default_text(d["ten_tk"], DOT_TEN_TK),
        "[sotaikhoannhantien]": default_text(d["stk"], DOT_STK),
        "[nganhangnhantien]": default_text(d["ngan_hang"], DOT_NGAN_HANG),
        "[cnnhnhantien]": default_text(d["chi_nhanh"], DOT_CHI_NHANH),
        "[mucdichsudung]": d["muc_dich"],
    }
    # 10 ô kỳ hạn trả nợ gốc: mỗi kỳ một dòng, ô thừa để trống
    for i in range(1, 11):
        key = f"[inpk{i}tien{i}]"
        if i <= len(schedule):
            ngay, tien = schedule[i - 1]
            line = fmt_ky_line(ngay, tien)
            m[key] = line if i == 1 else "\n" + line
        else:
            m[key] = ""
    return m


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]+', "", s).strip().replace(" ", "_")


# ================== KIỂM TRA DỮ LIỆU MỘT HỢP ĐỒNG ==================
UNIT_KEYS = ("pgd", "dia_chi_pgd", "sdt_pgd", "lanh_dao", "chuc_vu", "uy_quyen")   # đơn vị cho vay: lưu & dùng cho mọi hợp đồng
RECORD_KEYS = ("ho_ten", "ngay_sinh", "cccd", "ngay_cap", "noi_cap", "het_han",
               "dien_thoai", "thon", "xa", "tinh", "so_hd", "ngay_lap", "ngay_vay", "han_vay",
               "lai_suat", "lai_qh", "tong_tien", "so_ld", "muc_dich", "ten_tk", "stk",
               "ngan_hang", "chi_nhanh")                                  # riêng từng khách hàng
IMPORT_DEFAULT_KEYS = ("ngay_lap", "lai_suat", "tinh", "xa")              # ô Excel trống -> lấy theo form
SAVE_KEYS = UNIT_KEYS + ("tinh", "xa")                                    # các ô được nhớ giữa các lần mở tool


def parse_record(vals):
    """vals: dict chuỗi (khóa như trong form) -> (dict dữ liệu đã chuẩn hóa, list lỗi)."""
    g = lambda k: str(vals.get(k) or "").strip()
    errors, d = [], {}

    for key, label in (("lanh_dao", "Lãnh đạo"), ("chuc_vu", "Chức vụ"), ("ho_ten", "Họ và tên KH"),
                       ("cccd", "CCCD"), ("so_hd", "Số hợp đồng"), ("lai_suat", "Lãi suất"),
                       ("muc_dich", "Mục đích sử dụng vốn")):
        d[key] = g(key)
        if not d[key]:
            errors.append(f"Chưa nhập: {label}")

    d["ho_ten"] = fmt_name(d["ho_ten"])
    d["lanh_dao"] = fmt_name(d["lanh_dao"])
    d["lai_suat"] = d["lai_suat"].replace(".", ",")
    if d["lai_suat"] and d["lai_suat"] not in LAI_SUAT:
        errors.append("Lãi suất phải là " + " hoặc ".join(LAI_SUAT))

    for key, label, required in (("ngay_sinh", "Ngày sinh", True), ("ngay_cap", "Ngày cấp", False),
                                 ("ngay_lap", "Ngày lập hợp đồng", True),
                                 ("ngay_vay", "Ngày cho vay", True)):
        text = g(key)
        if not text:
            d[key] = None
            if required:
                errors.append(f"Chưa nhập: {label}")
            continue
        try:
            d[key] = parse_date(text)
        except ValueError:
            d[key] = None
            errors.append(f"{label} không hợp lệ (nhập dạng dd/mm/yyyy)")

    han = g("han_vay")
    if han.isdigit() and 1 <= int(han) <= 120:
        d["han_vay"] = int(han)
    else:
        d["han_vay"] = 0
        errors.append("Thời hạn vay phải là số tháng từ 1 đến 120")

    d["tong"] = parse_money(g("tong_tien"))
    if d["tong"] <= 0:
        errors.append("Chưa nhập Tổng tiền vay")

    ld = g("so_ld")
    if ld.isdigit():
        d["so_ld"] = int(ld)
    else:
        d["so_ld"] = 0
        errors.append("Số lao động được tạo việc làm phải là số")

    for key in ("pgd", "dia_chi_pgd", "sdt_pgd", "uy_quyen", "het_han", "dien_thoai",
                "ten_tk", "stk", "ngan_hang", "chi_nhanh"):
        d[key] = g(key)
    d["thon"], d["xa"] = fmt_name(g("thon")), fmt_name(g("xa"))
    d["loai_gt"] = LOAI_GT_MAC_DINH
    d["noi_cap"] = g("noi_cap") or suggest_noi_cap(g("ngay_cap"))
    tinh = g("tinh")
    d["tinh"] = normalize_tinh(tinh) if tinh else ""
    if tinh and d["tinh"] not in TINH_LIST:
        errors.append("Tỉnh không thuộc 34 tỉnh/thành sau sáp nhập")
    d["lai_qh"] = g("lai_qh") or lai_suat_qua_han(d["lai_suat"])
    return d, errors


def sao_ke_fields(row):
    """Một dòng sao kê -> các ô của form hợp đồng (chỉ ô có dữ liệu)."""
    def date_text(keys):
        text = get_field(row, *keys)
        try:
            return fmt_date(parse_date(text)) if text else ""
        except ValueError:
            return text

    ngay_cap = date_text(K_NGAYCAP)
    noi_cap = suggest_noi_cap(ngay_cap)
    if not noi_cap:
        raw = get_field(row, *K_NOICAP)
        noi_cap = raw if raw in NOI_CAP else ""
    thon = store.value_at(row, COT_THON) or get_field(row, "Tên thôn", "Thôn")        # cột H
    out = {
        "ho_ten": fmt_name(get_field(row, *K_TEN)),
        "ngay_sinh": date_text(K_NGAYSINH),
        "cccd": get_field(row, *K_CCCD),
        "ngay_cap": ngay_cap,
        "noi_cap": noi_cap,
        "dien_thoai": get_field(row, *K_SDT),
        "xa": fmt_name(get_field(row, *K_XA)),
        "thon": fmt_name(thon.strip()),
    }
    return {k: v for k, v in out.items() if v}


# ================== HÀNG LOẠT: EXCEL & TẠO FILE ==================
EXCEL_COLUMNS = [        # (khóa, tiêu đề cột, ví dụ ở file mẫu)
    ("ho_ten", "Họ và tên KH", "Nguyễn Văn A"),
    ("ngay_sinh", "Ngày sinh", "01/05/1990"),
    ("cccd", "CCCD", "010190001234"),
    ("ngay_cap", "Ngày cấp", "03/03/2022"),
    ("noi_cap", "Nơi cấp", ""),
    ("het_han", "Có thời hạn đến", ""),
    ("dien_thoai", "Điện thoại", "0912345678"),
    ("thon", "Thôn", "Thôn Pà Vầy"),
    ("xa", "Xã", ""),
    ("tinh", "Tỉnh", ""),
    ("so_hd", "Số hợp đồng", "01"),
    ("ngay_lap", "Ngày lập hợp đồng", ""),
    ("ngay_vay", "Ngày cho vay", "19/11/2026"),
    ("han_vay", "Thời hạn vay (tháng)", "60"),
    ("lai_suat", "Lãi suất", ""),
    ("tong_tien", "Tổng tiền vay", "100000000"),
    ("so_ld", "Số lao động", "2"),
    ("muc_dich", "Mục đích sử dụng vốn", "Chăn nuôi bò sinh sản"),
    ("ten_tk", "Tên tài khoản nhận tiền", ""),
    ("stk", "Số tài khoản", ""),
    ("ngan_hang", "Ngân hàng", ""),
    ("chi_nhanh", "Chi nhánh ngân hàng", ""),
]
EXCEL_ALIASES = {
    "Họ tên": "ho_ten", "Họ và tên": "ho_ten", "Tên khách hàng": "ho_ten", "Tên KH": "ho_ten",
    "Ngày tháng năm sinh": "ngay_sinh", "Số CCCD": "cccd", "CMND": "cccd", "Số điện thoại": "dien_thoai",
    "Ngày lập HĐ": "ngay_lap", "Ngày vay": "ngay_vay", "Thời hạn vay": "han_vay", "Thời hạn": "han_vay",
    "Số tiền vay": "tong_tien", "Số LĐ": "so_ld", "Số lao động được tạo việc làm": "so_ld",
    "Mục đích": "muc_dich", "Số HĐ": "so_hd",
}
EXCEL_NOTES = [
    "Mỗi dòng là một khách hàng; xóa dòng ví dụ trước khi dùng.",
    "Bắt buộc: Họ và tên KH, Ngày sinh, CCCD, Số hợp đồng, Ngày cho vay, Thời hạn vay (tháng), "
    "Tổng tiền vay, Số lao động, Mục đích sử dụng vốn.",
    "Họ tên, thôn, xã tự được viết hoa chữ cái đầu mỗi từ. Ngày nhập dạng dd/mm/yyyy.",
    "Nơi cấp để trống -> tự gợi ý theo Ngày cấp: từ 01/07/2024 về trước = Cục CSQLHC về TTXH, sau đó = Bộ Công an.",
    "Ô để trống ở: Ngày lập hợp đồng, Lãi suất, Xã, Tỉnh -> lấy theo giá trị đang chọn trong form.",
    "Tỉnh phải thuộc 34 tỉnh/thành sau sáp nhập (vd: Tuyên Quang, Lào Cai, Hà Nội...).",
    "Nếu có CCCD và đã nạp sao kê: các ô khách hàng còn trống (tên, ngày sinh, ngày cấp, SĐT, xã, thôn "
    "(cột H sao kê)) tự lấy từ sao kê.",
    "Lãi suất chỉ nhận 7,9248 hoặc 6,24.",
    "Thông tin đơn vị (PGD, Lãnh đạo, Chức vụ, Giấy ủy quyền) lấy từ form, không cần cột trong Excel.",
]


def read_batch_excel(path):
    """Đọc file Excel danh sách khách hàng -> list[dict khóa form -> chuỗi]."""
    raw = pd.read_excel(path, header=0, dtype=object)
    lookup = {remove_accents(label): key for key, label, _ in EXCEL_COLUMNS}
    lookup.update({remove_accents(a): k for a, k in EXCEL_ALIASES.items()})
    col_keys = [lookup.get(remove_accents(clean_header(h))) for h in raw.columns]
    if not any(col_keys):
        raise ValueError("Không nhận ra tiêu đề cột nào. Hãy dùng file Excel mẫu (nút 'Tải file Excel mẫu').")
    records = []
    for row in raw.values.tolist():
        vals = [cell_to_str(c) for c in row]
        if row_empty(vals):
            continue
        rec = {k: v.strip() for k, v in zip(col_keys, vals) if k and v.strip()}
        if rec:
            records.append(rec)
    return records


def normalize_record(rec):
    """Chuẩn hóa dòng nhập từ Excel: viết hoa tên, ngày dd/mm/yyyy, sửa số 0 đầu, tỉnh chuẩn, gợi ý nơi cấp."""
    rec = dict(rec)
    for key in ("ho_ten", "thon", "xa"):
        if rec.get(key):
            rec[key] = fmt_name(rec[key])
    for key in ("ngay_sinh", "ngay_cap", "ngay_lap", "ngay_vay"):
        if rec.get(key):
            try:
                rec[key] = fmt_date(parse_date(rec[key]))
            except ValueError:
                pass                                   # giữ nguyên, kiểm tra sau sẽ báo lỗi
    if rec.get("lai_suat"):
        rec["lai_suat"] = rec["lai_suat"].replace(".", ",")
    if rec.get("tinh"):
        rec["tinh"] = normalize_tinh(rec["tinh"])
    cccd = rec.get("cccd", "")
    if cccd.isdigit() and len(cccd) == 11:
        rec["cccd"] = cccd.zfill(12)
    sdt = rec.get("dien_thoai", "")
    if sdt.isdigit() and len(sdt) == 9:
        rec["dien_thoai"] = "0" + sdt
    if not rec.get("noi_cap") and rec.get("ngay_cap"):
        rec["noi_cap"] = suggest_noi_cap(rec["ngay_cap"])
    return rec


def fill_blank_from_sao_ke(rec):
    """Điền các ô khách hàng còn trống từ sao kê (tìm theo CCCD). Trả về (rec, có_điền)."""
    rec = dict(rec)
    if not rec.get("cccd") or not store.records:
        return rec, False
    found = store.find_customers(rec["cccd"])
    if not found:
        return rec, False
    filled = False
    for key, val in sao_ke_fields(found[0]).items():
        if not rec.get(key):
            rec[key] = val
            filled = True
    return rec, filled


def write_batch_template(path):
    """Ghi file Excel mẫu (sheet DanhSach + sheet HuongDan)."""
    from openpyxl.utils import get_column_letter
    cols = [label for _, label, _ in EXCEL_COLUMNS]
    example = [ex for _, _, ex in EXCEL_COLUMNS]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame([example], columns=cols).to_excel(xw, index=False, sheet_name="DanhSach")
        pd.DataFrame({"Hướng dẫn": EXCEL_NOTES}).to_excel(xw, index=False, sheet_name="HuongDan")
        for i, c in enumerate(cols, 1):
            xw.sheets["DanhSach"].column_dimensions[get_column_letter(i)].width = max(14, len(c) + 4)
        xw.sheets["HuongDan"].column_dimensions["A"].width = 120


def unique_path(folder, name, used):
    base, ext = os.path.splitext(name)
    path, n = os.path.join(folder, name), 2
    while path.lower() in used:                      # trùng tên trong cùng đợt -> thêm _2, _3...
        path = os.path.join(folder, f"{base}_{n}{ext}")
        n += 1
    used.add(path.lower())
    return path


def generate_batch(template, folder, records, unit, on_progress=None):
    """Tạo hợp đồng cho từng record. Trả về (list file đã tạo, list (số thứ tự, tên, lỗi))."""
    done, errors, used = [], [], set()
    n = len(records)
    for i, rec in enumerate(records, 1):
        name = rec.get("ho_ten") or f"dòng {i}"
        if on_progress:
            on_progress(i - 1, n, name)
        d, errs = parse_record({**unit, **rec})
        if errs:
            errors.append((i, name, "; ".join(errs)))
            continue
        out = unique_path(folder, safe_name(f"HopDong_{d['ho_ten']}_{d['so_hd']}") + ".docx", used)
        try:
            fill_template(template, out, build_mapping(d))
            done.append(out)
        except PermissionError:
            errors.append((i, name, "File đang mở (ví dụ trong Word), không ghi được"))
        except Exception as e:
            logging.exception("generate_batch")
            errors.append((i, name, str(e)))
    if on_progress:
        on_progress(n, n, "")
    return done, errors


# ================== GIAO DIỆN ==================
class HopDongApp:
    def __init__(self, root):
        self.root = root
        if isinstance(root, tk.Tk):
            ui.apply_theme(root)
            root.title("Tạo hợp đồng tín dụng GQVL")
            root.geometry("1100x720")

        self.main = main = tk.Frame(root, bg=ui.BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)
        self.v = {}                       # key -> StringVar (form một hợp đồng)
        self.cfg = self._load_config()
        self.batch = []                   # danh sách khách hàng để tạo hàng loạt
        self._editing = None              # chỉ số dòng đang sửa trong danh sách

        head = tk.Frame(main, bg=ui.BG)
        head.pack(fill="x", pady=(0, 8))
        left = tk.Frame(head, bg=ui.BG)
        left.pack(side="left")
        ui.label(left, "Hợp đồng tín dụng GQVL", bold=True, size=15).pack(anchor="w")
        ui.label(left, "Tạo từng hợp đồng hoặc tạo hàng loạt nhiều khách hàng theo mẫu Word.",
                 muted=True).pack(anchor="w")
        tabs = tk.Frame(head, bg=ui.BG)
        tabs.pack(side="right")
        self.tab_single = ui.RoundButton(tabs, text="Từng hợp đồng", width=150, pady=7,
                                         command=lambda: self._show_tab("single"))
        self.tab_single.pack(side="left", padx=3)
        self.tab_batch = ui.RoundButton(tabs, text="Hàng loạt (0)", width=150, pady=7,
                                        command=lambda: self._show_tab("batch"))
        self.tab_batch.pack(side="left", padx=3)

        self.pages = tk.Frame(main, bg=ui.BG)
        self.pages.pack(fill="both", expand=True)
        self.single = tk.Frame(self.pages, bg=ui.BG)
        self.batch_page = tk.Frame(self.pages, bg=ui.BG)
        self._save_job = None
        self._build_single(self.single)
        self._build_batch(self.batch_page)

        self._apply_defaults()
        store.subscribe(self._sk_changed)
        self._sk_changed()                      # tự điền Tên PGD từ sao kê nếu đang trống
        for key in ("han_vay", "ngay_vay", "tong_tien", "lai_suat"):
            self.v[key].trace_add("write", self._refresh)
        self.v["ngay_cap"].trace_add("write", self._on_ngay_cap)      # gợi ý nơi cấp
        for key in SAVE_KEYS:                   # đơn vị cho vay, tỉnh, xã: lưu ngay khi thay đổi
            self.v[key].trace_add("write", self._schedule_save)
        self._refresh()
        self.refresh_batch()
        self._show_tab("single")

    # ---------- chuyển tab ----------
    def _show_tab(self, name):
        self.single.pack_forget()
        self.batch_page.pack_forget()
        page = self.single if name == "single" else self.batch_page
        page.pack(fill="both", expand=True)
        for btn, key in ((self.tab_single, "single"), (self.tab_batch, "batch")):
            if key == name:
                btn.config(bg=ui.PRIMARY, hover=ui.PRIMARY_H, fg="white")
            else:
                btn.config(bg=ui.SECOND, hover=ui.SECOND_H, fg=ui.TEXT)
        if name == "batch":
            self._update_unit_info()

    # ================== TAB 1: TỪNG HỢP ĐỒNG ==================
    def _build_single(self, parent):
        bar = tk.Frame(parent, bg=ui.BG)
        bar.pack(side="bottom", fill="x", pady=(10, 0))
        self.btn_run = ui.RoundButton(bar, text="TẠO HỢP ĐỒNG", command=self.generate, padx=26, pady=10)
        self.btn_run.pack(side="left")
        self.btn_add = ui.secondary_button(bar, "Thêm vào danh sách", self.add_to_list, padx=18, pady=10)
        self.btn_add.pack(side="left", padx=10)
        ui.secondary_button(bar, "Nhập khách hàng mới", self.reset_customer, padx=18,
                            pady=10).pack(side="left")
        self.status = ui.label(bar, "", muted=True)
        self.status.pack(side="left", padx=12)

        scroll = ui.ScrollFrame(parent)
        scroll.pack(fill="both", expand=True)
        inner = scroll.inner
        inner.columnconfigure(0, weight=1, uniform="col")
        inner.columnconfigure(1, weight=1, uniform="col")
        left = tk.Frame(inner, bg=ui.BG)
        left.grid(row=0, column=0, sticky="new", padx=(0, 6))
        right = tk.Frame(inner, bg=ui.BG)
        right.grid(row=0, column=1, sticky="new", padx=(6, 0))

        self._card_customer(left)
        self._card_account(left)
        self._card_unit(right)
        self._card_loan(right)

    def _field(self, card, row, key, label, kind="entry", values=None, readonly=False, fmt=None):
        """fmt: hàm định dạng áp dụng ngay khi gõ (viết hoa tên, dd/mm/yyyy, chỉ số...)."""
        parent = card.content
        parent.columnconfigure(0, minsize=150)
        parent.columnconfigure(1, weight=1)
        ui.label(parent, label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        var = self.v.setdefault(key, tk.StringVar(master=self.main))
        if kind == "combo":
            w = ttk.Combobox(parent, textvariable=var, values=values, state="readonly")
        else:
            w = ttk.Entry(parent, textvariable=var, state="readonly" if readonly else "normal")
        w.grid(row=row, column=1, sticky="ew", pady=4)
        if fmt:
            self._bind_format(w, var, fmt)
        return w

    def _bind_format(self, entry, var, fn):
        """Áp dụng fn lên nội dung ô sau mỗi lần gõ / dán / rời ô, giữ nguyên vị trí con trỏ."""
        def apply(_e=None):
            s = var.get()
            new = fn(s)
            if new == s:
                return
            try:
                pos = entry.index("insert")
            except tk.TclError:
                pos = len(s)
            at_end = pos >= len(s)
            keep = sum(ch.isalnum() for ch in s[:pos])        # số ký tự chữ/số đứng trước con trỏ
            var.set(new)
            if at_end:
                entry.icursor("end")
                return
            newpos, count = 0, 0
            if keep:
                newpos = len(new)
                for i, ch in enumerate(new):
                    if ch.isalnum():
                        count += 1
                        if count >= keep:
                            newpos = i + 1
                            break
            entry.icursor(newpos)

        entry.bind("<KeyRelease>", apply, add="+")
        entry.bind("<FocusOut>", apply, add="+")
        entry.bind("<<Paste>>", lambda _e: entry.after_idle(apply), add="+")

    def _card_unit(self, parent):
        c = ui.Card(parent, title="Đơn vị cho vay (tự lưu, dùng cho mọi hợp đồng)")
        c.pack(fill="x", pady=(0, 10))
        self._field(c, 0, "pgd", "Tên PGD:")
        self._field(c, 1, "dia_chi_pgd", "Địa chỉ PGD:")
        self._field(c, 2, "sdt_pgd", "Điện thoại PGD:")
        self._field(c, 3, "lanh_dao", "Lãnh đạo (Ông):", fmt=fmt_name)
        self._field(c, 4, "chuc_vu", "Chức vụ:", "combo", CHUC_VU)
        self._field(c, 5, "uy_quyen", "Giấy ủy quyền:")

    def _card_customer(self, parent):
        c = ui.Card(parent, title="Khách hàng (Bên vay)")
        c.pack(fill="x", pady=(0, 10))
        self._lookup_bar(c)
        self._field(c, 0, "ho_ten", "Họ và tên KH:", fmt=fmt_name)
        self._field(c, 1, "ngay_sinh", "Ngày tháng năm sinh:", fmt=fmt_date_input)
        self._field(c, 2, "cccd", "CCCD:", fmt=fmt_digits(12))
        self._field(c, 3, "ngay_cap", "Ngày cấp:", fmt=fmt_date_input)
        self._field(c, 4, "noi_cap", "Nơi cấp (tự gợi ý):", "combo", NOI_CAP)
        self._field(c, 5, "het_han", "Có thời hạn đến:")
        self._field(c, 6, "dien_thoai", "Điện thoại:", fmt=fmt_digits(11))
        self._field(c, 7, "thon", "Thôn:", fmt=fmt_name)
        self._field(c, 8, "xa", "Xã:", fmt=fmt_name)
        self._field(c, 9, "tinh", "Tỉnh:", "combo", TINH_LIST)

    # ---------- lấy thông tin từ sao kê dùng chung ----------
    def _lookup_bar(self, card):
        bar = tk.Frame(card.body, bg=ui.CARD)
        bar.pack(before=card.content, fill="x", pady=(0, 2))
        ui.label(bar, "Tìm trong sao kê:").pack(side="left")
        self.v["tim_sk"] = tk.StringVar(master=self.main)
        ent = ttk.Entry(bar, textvariable=self.v["tim_sk"])
        ent.pack(side="left", fill="x", expand=True, padx=8)
        ent.bind("<Return>", lambda _e: self.lookup())
        ui.RoundButton(bar, text="Lấy thông tin", command=self.lookup, padx=14, pady=5,
                       size=9).pack(side="right")
        self.sk_hint = ui.label(card.body, "", muted=True, size=9)
        self.sk_hint.pack(before=card.content, anchor="w", pady=(0, 8))

    def _sk_changed(self):
        if not self.v["pgd"].get().strip():                  # Tên PGD lấy từ cột D của sao kê (bỏ chữ PGD)
            pgd = pgd_from_sao_ke()
            if pgd:
                self.v["pgd"].set(pgd)
        self.sk_hint.config(text=("Nhập CCCD, mã KH, số khế ước hoặc tên. Sao kê: " + store.summary())
                            if store.records else
                            "Chưa có sao kê - nạp ở khung \"Dữ liệu dùng chung\" bên trái để tự điền thông tin.")

    def lookup(self):
        q = self.v["tim_sk"].get().strip()
        if not store.records:
            messagebox.showinfo("Chưa có sao kê", "Hãy nạp file sao kê ở khung \"Dữ liệu dùng chung\" "
                                "bên trái màn hình.")
            return
        if not q:
            return
        found = store.find_customers(q)
        if not found:
            messagebox.showinfo("Không tìm thấy", f"Không có khách hàng nào khớp với \"{q}\" trong sao kê.")
        elif len(found) == 1:
            self._fill_from(found[0])
        else:
            self._choose(found)

    def _choose(self, found):
        top = tk.Toplevel(self.main)
        top.title("Chọn khách hàng")
        top.configure(bg=ui.BG)
        top.transient(self.main.winfo_toplevel())
        ui.label(top, f"Có {len(found)} kết quả, chọn một khách hàng:", bold=True,
                 bg=ui.BG).pack(anchor="w", padx=14, pady=(12, 6))
        lb = tk.Listbox(top, width=80, height=min(12, len(found)), font=ui.font(10),
                        activestyle="none", selectbackground=ui.PRIMARY, relief="flat",
                        highlightthickness=1, highlightbackground=ui.BORDER)
        for r in found:
            lb.insert("end", "  " + " | ".join(x for x in (
                get_field(r, *K_TEN), get_field(r, *K_CCCD), get_field(r, *K_XA),
                get_field(r, *K_DIACHI)) if x))
        lb.selection_set(0)
        lb.pack(padx=14)

        def pick(_e=None):
            sel = lb.curselection()
            if sel:
                top.destroy()
                self._fill_from(found[sel[0]])

        lb.bind("<Double-Button-1>", pick)
        lb.bind("<Return>", pick)
        ui.RoundButton(top, text="Chọn", command=pick, padx=24, pady=7).pack(pady=12)
        lb.focus_set()
        top.grab_set()

    def _fill_from(self, row):
        for key, val in sao_ke_fields(row).items():
            self.v[key].set(val)
        self.status.config(text="Đã lấy thông tin từ sao kê - hãy kiểm tra lại các ô trước khi tạo.",
                           fg=ui.OK)

    def _card_loan(self, parent):
        c = ui.Card(parent, title="Khoản vay")
        c.pack(fill="x", pady=(0, 10))
        self._field(c, 0, "so_hd", "Số hợp đồng:", fmt=fmt_digits(6))
        self._field(c, 1, "ngay_lap", "Ngày lập hợp đồng:", fmt=fmt_date_input)
        self._field(c, 2, "ngay_vay", "Ngày cho vay:", fmt=fmt_date_input)
        self._field(c, 3, "han_vay", "Thời hạn vay (tháng):", fmt=fmt_digits(3))
        self._field(c, 4, "phan_ky", "Phân kỳ (tự động):", readonly=True)
        self._field(c, 5, "lai_suat", "Lãi suất (%/năm):", "combo", LAI_SUAT)
        self._field(c, 6, "lai_qh", "Lãi quá hạn (130%):")
        self._field(c, 7, "tong_tien", "Tổng tiền vay (đồng):", fmt=fmt_money_input)
        self._field(c, 8, "bang_chu", "Bằng chữ (tự động):", readonly=True)
        self._field(c, 9, "so_ld", "Số LĐ được tạo việc làm:", fmt=fmt_digits(4))
        self._field(c, 10, "muc_dich", "Mục đích sử dụng vốn:")
        ui.label(c.content, "Kỳ hạn trả nợ gốc (xem trước):", muted=True).grid(
            row=11, column=0, sticky="nw", pady=(8, 0))
        self.preview = ui.label(c.content, "", muted=True, justify="left", anchor="w")
        self.preview.grid(row=11, column=1, sticky="w", pady=(8, 0))

    def _card_account(self, parent):
        c = ui.Card(parent, title="Tài khoản nhận tiền (để trống = in dấu chấm)")
        c.pack(fill="x", pady=(0, 10))
        self._field(c, 0, "ten_tk", "Tên tài khoản nhận tiền:")
        self._field(c, 1, "stk", "Số tài khoản:", fmt=fmt_digits(25))
        self._field(c, 2, "ngan_hang", "Ngân hàng nhận:")
        self._field(c, 3, "chi_nhanh", "Chi nhánh ngân hàng:")

    # ---------- xử lý ô nhập ----------
    def _apply_defaults(self):
        c = self.cfg
        for key in ("pgd", "dia_chi_pgd", "sdt_pgd", "lanh_dao", "uy_quyen", "xa"):
            self.v[key].set(c.get(key, ""))
        tinh = normalize_tinh(c.get("tinh", ""))
        self.v["tinh"].set(tinh if tinh in TINH_LIST else "")
        self.v["chuc_vu"].set(c.get("chuc_vu", CHUC_VU[0]))
        self.v["ngay_lap"].set(fmt_date(date.today()))

    def _on_ngay_cap(self, *_):
        """Gõ xong Ngày cấp -> gợi ý Nơi cấp (vẫn chọn lại được)."""
        text = self.v["ngay_cap"].get().strip()
        if len(text) == 10:
            goi_y = suggest_noi_cap(text)
            if goi_y:
                self.v["noi_cap"].set(goi_y)

    def _refresh(self, *_):
        try:
            self.v["lai_qh"].set(lai_suat_qua_han(self.v["lai_suat"].get()))
            tong = parse_money(self.v["tong_tien"].get())
            self.v["bang_chu"].set(so_thanh_chu(tong) if tong > 0 else "")

            han = self.v["han_vay"].get().strip()
            if han.isdigit() and 1 <= int(han) <= 120:
                han = int(han)
                step = ky_han_months(han)
                n = len(list(range(step, han, step))) + 1
                self.v["phan_ky"].set(f"Mỗi kỳ {step} tháng ({n} kỳ)")
                preview = ""
                try:
                    vay = parse_date(self.v["ngay_vay"].get())
                    if tong > 0:
                        preview = "\n".join(fmt_ky_line(d, a) for d, a in build_schedule(vay, han, tong))
                except ValueError:
                    pass
                self.preview.config(text=preview)
            else:
                self.v["phan_ky"].set("")
                self.preview.config(text="")
        except Exception:
            logging.exception("refresh")

    def reset_customer(self):
        keep = set(UNIT_KEYS) | {"tinh", "xa", "ngay_lap", "lai_suat"}
        for key, var in self.v.items():
            if key not in keep and key not in ("phan_ky", "lai_qh", "bang_chu"):
                var.set("")
        self._editing = None
        self.btn_add.config(text="Thêm vào danh sách")
        self.status.config(text="", fg=ui.MUTED)

    # ---------- cấu hình ----------
    def _load_config(self):
        return paths.load_section(CONFIG_SECTION)

    def _schedule_save(self, *_):
        """Lưu thông tin đơn vị cho vay sau khi ngừng gõ ~0,6 giây."""
        if self._save_job is not None:
            self.main.after_cancel(self._save_job)
        self._save_job = self.main.after(600, self._save_config)

    def _save_config(self, **extra):
        self._save_job = None
        self.cfg.update({k: self.v[k].get().strip() for k in SAVE_KEYS})
        self.cfg.update(extra)
        paths.save_section(CONFIG_SECTION, self.cfg)

    @staticmethod
    def _template_cfg(template):
        """Chỉ nhớ đường dẫn mẫu do người dùng tự chọn (mẫu mặc định nằm trong thư mục tạm của exe)."""
        return {} if template in TEMPLATE_CANDIDATES else {"template": template}

    def _find_template(self):
        for path in [self.cfg.get("template")] + TEMPLATE_CANDIDATES:
            if path and os.path.exists(path):
                return path
        messagebox.showinfo("Chọn file mẫu", "Không thấy file mẫu hợp đồng. Hãy chọn file hopdongmau.docx.")
        path = filedialog.askopenfilename(title="Chọn file mẫu hợp đồng",
                                          filetypes=[("Word files", "*.docx")])
        return path or None

    def _unit(self):
        return {k: self.v[k].get().strip() for k in UNIT_KEYS}

    def _record_from_form(self):
        return {k: self.v[k].get().strip() for k in RECORD_KEYS}

    # ---------- tạo một hợp đồng ----------
    def _collect(self):
        return parse_record({k: var.get() for k, var in self.v.items()})

    def generate(self):
        d, errors = self._collect()
        if errors:
            messagebox.showwarning("Thiếu / sai thông tin", "\n".join(f"• {e}" for e in errors))
            return
        template = self._find_template()
        if not template:
            return

        default_name = safe_name(f"HopDong_{d['ho_ten']}_{d['so_hd']}") + ".docx"
        out = filedialog.asksaveasfilename(
            title="Lưu hợp đồng", defaultextension=".docx", initialfile=default_name,
            initialdir=paths.output_dir("hopdong"), filetypes=[("Word files", "*.docx")])
        if not out:
            return
        try:
            leftover = fill_template(template, out, build_mapping(d))
        except PermissionError:
            messagebox.showerror("Không ghi được file", "File đang được mở (ví dụ trong Word). "
                                 "Hãy đóng file hoặc chọn tên khác.")
            return
        except Exception as e:
            logging.exception("generate hop dong")
            messagebox.showerror("Lỗi", f"Không tạo được hợp đồng:\n{e}")
            return

        self._save_config(**self._template_cfg(template))
        self.status.config(text=f"Đã tạo: {os.path.basename(out)}", fg=ui.OK)
        msg = "Đã tạo hợp đồng.\nMở file ngay?"
        if leftover:
            msg = "Lưu ý: mẫu còn ô chưa được điền: " + ", ".join(leftover) + "\n\n" + msg
        if messagebox.askyesno("Hoàn tất", msg) and hasattr(os, "startfile"):
            try:
                os.startfile(out)
            except OSError:
                pass

    # ================== TAB 2: HÀNG LOẠT ==================
    def _build_batch(self, parent):
        # thanh tạo tất cả (pack trước để luôn hiện dưới cùng)
        bar = tk.Frame(parent, bg=ui.BG)
        bar.pack(side="bottom", fill="x", pady=(10, 0))
        self.btn_all = ui.RoundButton(bar, text="TẠO TẤT CẢ HỢP ĐỒNG", command=self.generate_all,
                                      padx=26, pady=10)
        self.btn_all.pack(side="left")
        self.progress = ttk.Progressbar(bar, length=220, mode="determinate")
        self.progress.pack(side="left", padx=14)
        self.b_status = ui.label(bar, "", muted=True)
        self.b_status.pack(side="left")

        self.unit_info = ui.label(parent, "", muted=True)
        self.unit_info.pack(anchor="w", pady=(0, 6))

        tools_bar = tk.Frame(parent, bg=ui.BG)
        tools_bar.pack(fill="x", pady=(0, 8))
        ui.RoundButton(tools_bar, text="Nhập từ Excel", command=self.import_excel, padx=16,
                       pady=7).pack(side="left")
        ui.secondary_button(tools_bar, "Tải file Excel mẫu", self.download_template, padx=16,
                            pady=7).pack(side="left", padx=8)
        ui.secondary_button(tools_bar, "Sửa dòng chọn", self.edit_selected, padx=14,
                            pady=7).pack(side="left")
        ui.secondary_button(tools_bar, "Xóa dòng chọn", self.delete_selected, padx=14,
                            pady=7).pack(side="left", padx=8)
        ui.RoundButton(tools_bar, text="Xóa hết", command=self.clear_batch, bg=ui.DANGER,
                       hover=ui.DANGER_H, padx=14, pady=7).pack(side="right")

        card = ui.Card(parent, title="Danh sách khách hàng", fill=True)
        card.pack(fill="both", expand=True)
        holder = tk.Frame(card.content, bg=ui.CARD)
        holder.pack(fill="both", expand=True)
        cols = ("stt", "ten", "cccd", "so_hd", "tien", "han", "note")
        self.btree = ttk.Treeview(holder, columns=cols, show="headings", selectmode="extended")
        for col, text, w, anchor in (("stt", "STT", 45, "center"), ("ten", "Họ và tên", 170, "w"),
                                     ("cccd", "CCCD", 120, "w"), ("so_hd", "Số HĐ", 70, "w"),
                                     ("tien", "Tổng tiền vay", 120, "e"), ("han", "Tháng", 60, "center"),
                                     ("note", "Trạng thái", 300, "w")):
            self.btree.heading(col, text=text)
            self.btree.column(col, width=w, anchor=anchor)
        vsb = ttk.Scrollbar(holder, orient="vertical", command=self.btree.yview)
        self.btree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.btree.pack(side="left", fill="both", expand=True)
        self.btree.tag_configure("bad", foreground=ui.ERR)
        self.btree.bind("<Double-Button-1>", lambda _e: self.edit_selected())

    def _update_unit_info(self):
        u = self._unit()
        who = f"{u['chuc_vu']} {u['lanh_dao']}".strip() or "(chưa nhập lãnh đạo)"
        self.unit_info.config(text=f"Đơn vị áp dụng cho mọi hợp đồng: PGD {u['pgd'] or '(chưa nhập)'} - "
                                   f"{who}. Sửa ở tab \"Từng hợp đồng\".")

    def refresh_batch(self):
        """Vẽ lại danh sách + kiểm tra từng dòng."""
        self.btree.delete(*self.btree.get_children())
        check_unit = {**self._unit(), "lanh_dao": "x", "chuc_vu": "x"}    # lỗi đơn vị báo riêng khi tạo
        bad = 0
        for i, rec in enumerate(self.batch):
            _, errs = parse_record({**check_unit, **rec})
            if errs:
                bad += 1
            note = "OK" if not errs else "Thiếu/sai: " + "; ".join(
                e.replace("Chưa nhập: ", "") for e in errs[:3]) + (" ..." if len(errs) > 3 else "")
            tien = parse_money(rec.get("tong_tien", ""))
            self.btree.insert("", "end", iid=str(i), tags=("bad",) if errs else (), values=(
                i + 1, rec.get("ho_ten", ""), rec.get("cccd", ""), rec.get("so_hd", ""),
                fmt_money(tien) if tien else "", rec.get("han_vay", ""), note))
        self.tab_batch.config(text=f"Hàng loạt ({len(self.batch)})")
        return bad

    def add_to_list(self):
        rec = self._record_from_form()
        _, errs = parse_record({**self._unit(), **rec})
        errs = [e for e in errs if "Lãnh đạo" not in e and "Chức vụ" not in e]
        if errs:
            messagebox.showwarning("Thiếu / sai thông tin", "\n".join(f"• {e}" for e in errs))
            return
        if self._editing is not None and self._editing < len(self.batch):
            self.batch[self._editing] = rec
            msg = f"Đã cập nhật {rec['ho_ten']} trong danh sách."
        else:
            self.batch.append(rec)
            msg = f"Đã thêm {rec['ho_ten']} vào danh sách ({len(self.batch)} khách hàng)."
        self.refresh_batch()
        self.reset_customer()                      # sẵn sàng nhập khách tiếp theo
        self.status.config(text=msg, fg=ui.OK)

    def import_excel(self):
        path = filedialog.askopenfilename(title="Chọn file Excel danh sách khách hàng",
                                          filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            rows = read_batch_excel(path)
        except Exception as e:
            logging.exception("read_batch_excel")
            messagebox.showerror("Lỗi", f"Không đọc được file Excel:\n{e}")
            return
        defaults = {k: self.v[k].get().strip() for k in IMPORT_DEFAULT_KEYS}
        from_sk = 0
        for rec in rows:
            rec = normalize_record(rec)
            rec, filled = fill_blank_from_sao_ke(rec)
            from_sk += filled
            for key, val in defaults.items():
                if val and not rec.get(key):
                    rec[key] = val
            self.batch.append(rec)
        bad = self.refresh_batch()
        msg = f"Đã nhập {len(rows)} khách hàng."
        if from_sk:
            msg += f"\n{from_sk} dòng được bổ sung thông tin từ sao kê."
        if bad:
            msg += f"\n{bad} dòng còn thiếu / sai thông tin (hiện màu đỏ). Bấm đúp vào dòng để sửa."
        messagebox.showinfo("Nhập từ Excel", msg)
        self.b_status.config(text=f"{len(self.batch)} khách hàng trong danh sách", fg=ui.MUTED)

    def download_template(self):
        path = filedialog.asksaveasfilename(title="Lưu file Excel mẫu", defaultextension=".xlsx",
                                            initialfile="MauDanhSachHopDong.xlsx",
                                            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            write_batch_template(path)
        except PermissionError:
            messagebox.showerror("Không ghi được file", "File đang được mở. Hãy đóng file hoặc chọn tên khác.")
            return
        except Exception as e:
            logging.exception("write_batch_template")
            messagebox.showerror("Lỗi", f"Không tạo được file mẫu:\n{e}")
            return
        self.b_status.config(text="Đã lưu file Excel mẫu.", fg=ui.OK)
        if hasattr(os, "startfile"):
            try:
                os.startfile(path)
            except OSError:
                pass

    def _selected_indexes(self):
        return sorted(int(i) for i in self.btree.selection())

    def edit_selected(self):
        idx = self._selected_indexes()
        if not idx:
            messagebox.showinfo("Chọn dòng", "Hãy chọn một dòng trong danh sách để sửa.")
            return
        i = idx[0]
        rec = self.batch[i]
        for key in RECORD_KEYS:
            self.v[key].set(rec.get(key, ""))
        if not rec.get("lai_qh"):
            self.v["lai_qh"].set(lai_suat_qua_han(rec.get("lai_suat", "")))
        self._editing = i
        self.btn_add.config(text="Cập nhật danh sách")
        self.status.config(text=f"Đang sửa dòng {i + 1}: {rec.get('ho_ten', '')}", fg=ui.INFO)
        self._show_tab("single")

    def delete_selected(self):
        idx = self._selected_indexes()
        if not idx:
            return
        for i in reversed(idx):
            del self.batch[i]
        self._editing = None
        self.btn_add.config(text="Thêm vào danh sách")
        self.refresh_batch()

    def clear_batch(self):
        if self.batch and messagebox.askyesno("Xác nhận", f"Xóa toàn bộ {len(self.batch)} khách hàng "
                                              "trong danh sách?"):
            self.batch.clear()
            self._editing = None
            self.btn_add.config(text="Thêm vào danh sách")
            self.refresh_batch()

    def generate_all(self):
        if not self.batch:
            messagebox.showinfo("Danh sách trống", "Hãy nhập từ Excel hoặc thêm khách hàng vào danh sách trước.")
            return
        unit = self._unit()
        if not unit["lanh_dao"] or not unit["chuc_vu"]:
            messagebox.showwarning("Thiếu thông tin đơn vị",
                                   "Chưa nhập Lãnh đạo / Chức vụ. Hãy điền ở tab \"Từng hợp đồng\" (khung Đơn vị cho vay).")
            return
        template = self._find_template()
        if not template:
            return
        # Mỗi đợt hàng loạt lưu vào một thư mục con riêng trong "Hop dong GQVL", tránh lẫn / ghi đè đợt trước
        folder = os.path.join(paths.output_dir("hopdong"), datetime.now().strftime("Hang loat %d-%m-%Y %Hh%M"))
        os.makedirs(folder, exist_ok=True)

        self.btn_all.config(state="disabled")
        self.progress.config(maximum=len(self.batch), value=0)

        def on_progress(i, n, name):
            self.progress.config(value=i)
            self.b_status.config(text=f"Đang tạo {i + 1}/{n}: {name}" if i < n else "Hoàn tất", fg=ui.MUTED)
            self.main.update_idletasks()

        try:
            done, errors = generate_batch(template, folder, self.batch, unit, on_progress)
        finally:
            self.btn_all.config(state="normal")

        self._save_config(**self._template_cfg(template))
        self.b_status.config(text=f"Đã tạo {len(done)}/{len(self.batch)} hợp đồng",
                             fg=ui.OK if not errors else ui.ERR)
        msg = f"Đã tạo {len(done)} hợp đồng trong:\n{folder}"
        if errors:
            lines = [f"• Dòng {i} ({name}): {err}" for i, name, err in errors[:12]]
            if len(errors) > 12:
                lines.append(f"... và {len(errors) - 12} dòng lỗi khác")
            msg += f"\n\n{len(errors)} dòng chưa tạo được:\n" + "\n".join(lines)
        if done:
            msg += "\n\nMở thư mục ngay?"
            if messagebox.askyesno("Hoàn tất", msg):
                paths.open_folder(folder)
        else:
            messagebox.showwarning("Không tạo được hợp đồng nào", msg)


# ===== Khai báo cho launcher =====
TOOL_NAME = "Hợp đồng GQVL"
TOOL_CLASS = HopDongApp
TOOL_ORDER = 30


if __name__ == "__main__":
    root = tk.Tk()
    HopDongApp(root)
    root.mainloop()
