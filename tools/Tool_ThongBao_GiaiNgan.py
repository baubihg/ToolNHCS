# -*- coding: utf-8 -*-
"""
Tool tạo Thông báo phê duyệt (Mẫu 04) & Mẫu 06/TD  -  V18
- Đọc "Danh sách khoản vay được phê duyệt giải ngân" (Mẫu số 05/GDX, bố cục mới)
- Ưu tiên đọc theo BẢNG (extract_tables), dự phòng đọc theo DÒNG VĂN BẢN
- Đối chiếu tổng tiền từng tổ / tổng cộng với PDF, xem trước trước khi xuất file

Thư viện: pandas, openpyxl (xlsx), xlrd (nếu dùng .xls), python-docx, pdfplumber
"""
import json
import logging
import numbers
import os
import queue
import re
import sys
import threading
import unicodedata
from collections import Counter
from copy import deepcopy
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import pdfplumber
from docx import Document

try:
    from . import _ui as ui
    from . import _paths as paths
except ImportError:          # chạy trực tiếp file này, không qua launcher
    import _ui as ui
    import _paths as paths

paths.setup_logging()        # nhật ký chung của chương trình (chỉ tạo file khi có ghi)
CONFIG_SECTION = "thongbao"  # mục cấu hình của tool này trong config.json

# Tên cột Excel (đã chuẩn hóa bằng _norm: chữ thường, gộp khoảng trắng)
COL_DIEM = "điểm giao dịch"
COL_MA_DIEM = ("mã điểm giao dịch", "mã điểm")


# ================== CHUẨN HÓA CHUỖI ==================
def _norm(s):
    """NFC + bỏ ký tự ẩn + gộp whitespace + lowercase."""
    s = unicodedata.normalize("NFC", str(s) if s is not None else "")
    s = s.replace("\xa0", " ")
    s = re.sub(r"[\u200b-\u200f\ufeff]", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def _strip_accents(s):
    """Bỏ dấu tiếng Việt (xử lý cả đ/Đ vì NFD không tách được)."""
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return unicodedata.normalize("NFC", s)


def _title(s):
    return " ".join(w.capitalize() for w in s.split())


def _safe_name(s):
    return re.sub(r'[\\/*?:"<>|]', "", s).strip()


def _val(row, col, default=""):
    """Lấy giá trị ô Excel; ô trống/NaN trả về default (không bao giờ ra chuỗi 'nan')."""
    v = row.get(col)
    if v is None or pd.isna(v):
        return default
    s = str(v).strip()
    return default if s == "" or s.lower() == "nan" else s


def fmt_money(x):
    return "{:,.0f}".format(x).replace(",", ".")


def fmt_trieu(v):
    """Đổi sang triệu; nếu không tròn triệu thì giữ phần lẻ (25,5) thay vì cắt bỏ."""
    return ("%g" % (v / 1_000_000)).replace(".", ",")


# ================== ĐỌC SỐ THÀNH CHỮ ==================
def doc_so_thanh_chu(so):
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
        if tram > 0 or day_du:               # nhóm không phải nhóm đầu: "không trăm"
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
        day_du = i < len(groups) - 1
        parts.append((doc3(groups[i], day_du) + " " + don_vi[i]).strip())
    text = " ".join(parts)
    return text[:1].upper() + text[1:] + " đồng chẵn."


# ================== NGÀY THÁNG ==================
def format_ngay(value, ten_cot):
    """Lấy NGÀY từ ô Excel (số 13, chuỗi '13', ô ngày tháng). Tháng/năm theo ngày chạy máy.
    Ô trống/không đọc được -> báo lỗi (không âm thầm lấy ngày hôm nay)."""
    ngay = None
    if value is not None and not pd.isna(value):
        try:
            if isinstance(value, (datetime, pd.Timestamp)):
                ngay = value.day
            elif isinstance(value, numbers.Real) and not isinstance(value, bool):
                f = float(value)
                if f.is_integer() and 1 <= int(f) <= 31:
                    ngay = int(f)
                else:
                    p = pd.to_datetime(f, errors="coerce", unit="D", origin="1899-12-30")
                    ngay = None if pd.isna(p) else p.day
            elif isinstance(value, str):
                s = value.strip()
                if s.isdigit() and 1 <= int(s) <= 31:
                    ngay = int(s)
                else:
                    p = pd.to_datetime(s, errors="coerce", dayfirst=True)
                    ngay = None if pd.isna(p) else p.day
        except (ValueError, TypeError, OverflowError):
            ngay = None
    if ngay is None:
        raise ValueError(f"Cột '{ten_cot}' trong Excel trống hoặc không đọc được (giá trị: {value!r}).")
    now = datetime.now()
    thang = f"0{now.month}" if now.month in (1, 2) else str(now.month)
    return f"ngày {ngay:02d} tháng {thang} năm {now.year}"


# ================== ĐỊA CHỈ / TÊN KHÁCH HÀNG ==================
def addr_parts(addr):
    """'Nà Sang~Tát Ngà~-~Tuyên Quang' -> ['Nà Sang', 'Tát Ngà', 'Tuyên Quang']"""
    addr = addr or ""
    if "~" in addr:
        parts = addr.split("~")
    else:
        parts = re.split(r"\s*,\s*|\s+-\s+", addr)
    return [p.strip() for p in parts if p.strip() and p.strip() not in ("-", "–")]


def thon_from_address(addr, xa, tinh):
    parts = addr_parts(addr)
    if len(parts) > 1:
        return parts[0]
    t = parts[0] if parts else ""
    for suffix in (tinh, xa):
        i = t.lower().rfind(suffix.lower()) if suffix else -1
        if i > 0:
            t = t[:i].strip()
    return t


def split_name_address(text):
    """Chế độ đọc theo dòng văn bản: tên KH là cụm từ IN HOA đứng đầu, phần còn lại là địa chỉ."""
    words = text.split()
    i = 0
    while i < len(words) and words[i] == words[i].upper() and any(c.isalpha() for c in words[i]):
        i += 1
    return " ".join(words[:i]), " ".join(words[i:])


# ================== ĐỌC PDF ==================
MONEY = r"\d{1,3}(?:[.,]\d{3})+"
RE_CONG = re.compile(r"Cộng\s*\(theo tổ\)\s*-?\s*(.*?)\s*:?\s*(" + MONEY + r")?\s*$", re.I)
RE_ROW = re.compile(r"^(\d+)\s+(?:(\d{6,9})\s+)?(\d{10})\s+(.*?)\s+(\d{16})\s+([A-Z][A-Z0-9]*)\b(.*)$")

COL_KEYS = {"ma_to": "mã tổ", "ma_kh": "mã khách hàng", "ten": "tên khách hàng",
            "dia_chi": "địa chỉ", "ma_kv": "mã khoản vay", "sp": "tên sản phẩm vay",
            "tien": "số tiền giải ngân"}
DEFAULT_COLMAP = {"ma_to": 1, "ma_kh": 2, "ten": 3, "dia_chi": 4, "ma_kv": 5, "sp": 6, "tien": 8}


def parse_money(s, need_sep=False):
    """'100.000.000' hoặc '100,000,000' -> 100000000. None nếu không phải số tiền."""
    s = (s or "").strip()
    pattern = MONEY if need_sep else r"\d{1,3}(?:[.,]\d{3})*"
    if re.fullmatch(pattern, s):
        return int(re.sub(r"\D", "", s))
    return None


class RowCollector:
    """Gom khách hàng theo tổ; tổ được chốt khi gặp dòng 'Cộng (theo tổ) - <tên tổ>'."""

    def __init__(self):
        self.rows, self.pending = [], []
        self.ma_to = ""
        self.to_totals = []          # (mã tổ, tên tổ, tổng tính được, tổng trong PDF)
        self.grand_total = None
        self.warnings = []
        self.missing_amount = 0

    def add_customer(self, ma_to, ma_kh, ten, dia_chi, ma_kv, sp, tien):
        if ma_to:
            self.ma_to = ma_to
        if tien is None:
            self.missing_amount += 1
        self.pending.append({"Mã KH": ma_kh, "Tên KH": ten.upper(), "Địa chỉ": dia_chi,
                             "Mã KV": ma_kv, "Chương trình": sp, "Tiền": tien or 0})

    def close_group(self, to_name, total):
        if not self.pending:
            self.warnings.append(f"Dòng 'Cộng (theo tổ) - {to_name}' không có khách hàng nào phía trên.")
            return
        s = sum(r["Tiền"] for r in self.pending)
        if total is not None and s != total:
            self.warnings.append(f"Tổ '{to_name}': tổng các khoản vay {fmt_money(s)} "
                                 f"khác dòng Cộng trong PDF {fmt_money(total)}.")
        if not self.ma_to:
            self.warnings.append(f"Tổ '{to_name}': không đọc được Mã tổ.")
        for r in self.pending:
            r["Tổ"], r["Mã Tổ"] = to_name, self.ma_to
            self.rows.append(r)
        self.to_totals.append((self.ma_to, to_name, s, total))
        self.pending, self.ma_to = [], ""

    def finalize(self):
        if self.pending:
            raise ValueError(f"Còn {len(self.pending)} khoản vay chưa có dòng 'Cộng (theo tổ)' đi kèm.")
        if not self.rows:
            raise ValueError("Không đọc được khoản vay nào.")
        if self.missing_amount:
            raise ValueError(f"Không đọc được số tiền của {self.missing_amount} khoản vay.")
        s = sum(r["Tiền"] for r in self.rows)
        if self.grand_total is None:
            self.warnings.append("Không tìm thấy dòng 'Tổng cộng' trong PDF để đối chiếu.")
        elif s != self.grand_total:
            self.warnings.append(f"Tổng các khoản vay {fmt_money(s)} khác 'Tổng cộng' trong PDF "
                                 f"{fmt_money(self.grand_total)}.")


def _cell(cells, idx):
    return cells[idx] if idx is not None and 0 <= idx < len(cells) else ""


def _map_columns(cells):
    normed = [_norm(c) for c in cells]
    if "tên khách hàng" not in normed or "số tiền giải ngân" not in normed:
        return None
    return {k: normed.index(v) for k, v in COL_KEYS.items() if v in normed}


def _consume_cells(col, cells, cm):
    cong = next((c for c in cells if _norm(c).startswith("cộng (theo tổ)")), None)
    tong = any(_norm(c).startswith("tổng cộng") for c in cells)
    if cong or tong:
        amount = parse_money(_cell(cells, cm["tien"]))
        if amount is None:
            amount = next((parse_money(c, True) for c in cells if parse_money(c, True) is not None), None)
        if cong:
            m = RE_CONG.match(cong)
            col.close_group(m.group(1).strip() if m else cong, amount)
        else:
            col.grand_total = amount
        return
    ma_kv = next((c for c in cells if re.fullmatch(r"\d{16}", c)), "")
    if not ma_kv:
        return
    col.add_customer(
        ma_to=re.sub(r"\D", "", _cell(cells, cm["ma_to"])),
        ma_kh=re.sub(r"\D", "", _cell(cells, cm["ma_kh"])),
        ten=_cell(cells, cm["ten"]), dia_chi=_cell(cells, cm["dia_chi"]),
        ma_kv=ma_kv, sp=_cell(cells, cm["sp"]), tien=parse_money(_cell(cells, cm["tien"])))


def _parse_by_tables(pages):
    col, colmap = RowCollector(), dict(DEFAULT_COLMAP)
    for page in pages:
        for table in page.extract_tables():
            for raw in table:
                if not raw:
                    continue
                cells = [re.sub(r"\s+", " ", unicodedata.normalize("NFC", c or "")).strip() for c in raw]
                cm = _map_columns(cells)
                if cm:                       # dòng tiêu đề (trang sau có thể không lặp lại)
                    colmap = {**DEFAULT_COLMAP, **cm}
                    continue
                _consume_cells(col, cells, colmap)
    return col


def _parse_by_text(texts):
    col = RowCollector()
    for text in texts:
        for line in text.split("\n"):
            line = line.strip()
            n = _norm(line)
            if n.startswith("cộng (theo tổ)"):
                m = RE_CONG.match(line)
                col.close_group(m.group(1).strip() if m else line, parse_money(m.group(2)) if m else None)
            elif n.startswith("tổng cộng"):
                m = re.search(MONEY, line)
                col.grand_total = parse_money(m.group(0)) if m else None
            else:
                m = RE_ROW.match(line)
                if not m:
                    continue
                _stt, ma_to, ma_kh, mid, ma_kv, sp, rest = m.groups()
                ten, dia_chi = split_name_address(mid)
                money = re.findall(MONEY, rest)
                col.add_customer(ma_to or "", ma_kh, ten, dia_chi, ma_kv, sp,
                                 parse_money(money[0]) if money else None)
    return col


def parse_header(texts):
    diems, ma_diem, pgd, chi_nhanh = [], "", "", ""
    for text in texts:
        for line in text.split("\n"):
            m = re.search(r"Tên điểm giao dịch\s*:\s*(.*)", line, re.I)
            if m:
                name = re.split(r"\s*Mã điểm giao dịch|\s*Đơn vị", m.group(1), flags=re.I)[0].strip()
                if name:
                    diems.append(name)
                m2 = re.search(r"Mã điểm giao dịch\s*:\s*(\S+)", line, re.I)
                if m2:
                    ma_diem = m2.group(1).strip()
                continue
            m = re.match(r"\s*Điểm\s*:\s*(.*)", line, re.I)      # mẫu cũ, neo đầu dòng
            if m:
                name = re.split(r"Mã điểm|Đơn vị", m.group(1), flags=re.I)[0].strip()
                if name:
                    diems.append(name)
                continue
            m = re.search(r"PGD\s+(.+?)\s*-\s*Chi nhánh\s+(.+)$", line)
            if m and not pgd:
                pgd = m.group(1).strip()
                chi_nhanh = re.sub(r"NGÂN HÀNG.*$", "", m.group(2)).strip()
    uniq = list(dict.fromkeys(diems))
    if not uniq:
        raise ValueError("Không đọc được 'Tên điểm giao dịch' từ PDF.")
    if len(uniq) > 1:
        raise ValueError("PDF chứa nhiều điểm giao dịch (" + ", ".join(uniq) +
                         "). Hãy tách thành từng file riêng để tránh gán nhầm ngày.")
    return {"diem": uniq[0], "ma_diem": ma_diem, "pgd": pgd, "chi_nhanh": chi_nhanh}


def parse_pdf(path):
    errors = []
    with pdfplumber.open(path) as pdf:
        pages = list(pdf.pages)
        texts = [unicodedata.normalize("NFC", p.extract_text() or "") for p in pages]
        header = parse_header(texts)
        for mode, fn in (("bảng", lambda: _parse_by_tables(pages)),
                         ("văn bản", lambda: _parse_by_text(texts))):
            try:
                col = fn()
                col.finalize()
                logging.info("Đọc PDF theo %s: %d khoản vay, %d tổ", mode, len(col.rows), len(col.to_totals))
                return {**header, "rows": col.rows, "to_totals": col.to_totals,
                        "grand_total": col.grand_total, "warnings": col.warnings, "mode": mode}
            except ValueError as e:
                errors.append(f"[đọc theo {mode}] {e}")
    raise ValueError("Không bóc tách được dữ liệu PDF.\n" + "\n".join(errors) +
                     "\n\nBấm 'Chẩn đoán PDF' để xuất nội dung PDF đọc được ra file kiểm tra.")


def dump_pdf_debug(pdf_path):
    out = paths.ensure_parent(os.path.join(os.path.dirname(paths.log_path()), "chan_doan_pdf.txt"))
    with pdfplumber.open(pdf_path) as pdf, open(out, "w", encoding="utf-8") as f:
        for i, page in enumerate(pdf.pages, 1):
            f.write(f"===== TRANG {i} - TEXT =====\n{page.extract_text()}\n")
            f.write(f"===== TRANG {i} - TABLES =====\n")
            for t in page.extract_tables():
                for r in t:
                    f.write(repr(r) + "\n")
                f.write("-----\n")
    return out


# ================== ĐỌC EXCEL & KHỚP ĐIỂM GIAO DỊCH ==================
def load_excel(path):
    df = pd.read_excel(path)
    df.columns = [_norm(c) for c in df.columns]
    return df


def _contains(longer, shorter):
    return re.search(r"(?<!\w)" + re.escape(shorter) + r"(?!\w)", longer) is not None


def find_info_row(df, diem, ma_diem=""):
    """Khớp: mã điểm (nếu Excel có cột) -> tên chính xác -> không dấu -> cụm từ -> cụm từ không dấu.
    Không bao giờ tự lấy dòng đầu; nhập nhằng thì báo lỗi."""
    if df.empty:
        raise ValueError("File Excel không có dữ liệu.")
    if COL_DIEM not in df.columns:
        raise ValueError("Không tìm thấy cột 'Điểm giao dịch' trong Excel. Các cột hiện có: "
                         + ", ".join(map(str, df.columns)))
    d = _norm(diem)
    if not d:
        raise ValueError("Không đọc được tên điểm giao dịch từ PDF.")

    for mc in COL_MA_DIEM:
        if ma_diem and mc in df.columns:
            m = df.loc[df[mc].astype(str).str.strip().str.upper() == ma_diem.strip().upper()]
            if len(m) == 1:
                logging.info("[MATCH] MÃ ĐIỂM %s", ma_diem)
                return m.iloc[0]

    names = df[COL_DIEM].apply(lambda v: "" if pd.isna(v) else _norm(v))

    for tf, label in ((lambda s: s, "EXACT"), (_strip_accents, "EXACT-NO-ACCENT")):
        m = df.loc[names.apply(tf) == tf(d)]
        if len(m) == 1:
            logging.info("[MATCH] %s -> %s", label, m.iloc[0][COL_DIEM])
            return m.iloc[0]
        if len(m) > 1:
            raise ValueError(f"Excel có nhiều dòng trùng tên điểm giao dịch '{diem}'.")

    for tf, label in ((lambda s: s, "PHRASE"), (_strip_accents, "PHRASE-NO-ACCENT")):
        key, series, cands = tf(d), names.apply(tf), []
        for idx, en in series.items():
            if not en:
                continue
            if _contains(key, en):
                cands.append((len(en), idx))
            elif _contains(en, key):
                cands.append((len(key), idx))
        if cands:
            top = max(c[0] for c in cands)
            best = [c for c in cands if c[0] == top]
            if len(best) == 1:
                logging.info("[MATCH] %s -> %s", label, df.loc[best[0][1]][COL_DIEM])
                return df.loc[best[0][1]]

    ds = ", ".join(str(x) for x in df[COL_DIEM].dropna().tolist())
    raise ValueError(f"Không xác định được duy nhất điểm giao dịch '{diem}' trong Excel "
                     f"(không tự lấy dòng đầu để tránh sai ngày).\nDanh sách trong Excel: {ds}")


# ================== DỰNG NGỮ CẢNH (ĐỌC + KIỂM TRA) ==================
def build_context(file_pdf, file_excel):
    pdf = parse_pdf(file_pdf)
    df = load_excel(file_excel)
    info = find_info_row(df, pdf["diem"], pdf["ma_diem"])
    warnings = list(pdf["warnings"])
    rows = pdf["rows"]

    # xã / tỉnh suy ra từ địa chỉ PDF làm dự phòng khi Excel trống
    cx, ct = Counter(), Counter()
    for r in rows:
        p = addr_parts(r["Địa chỉ"])
        if len(p) >= 3:
            cx[p[1]] += 1
            ct[p[-1]] += 1
    ten_xa = _val(info, "xã") or (cx.most_common(1)[0][0] if cx else "")
    ten_tinh = _val(info, "tỉnh") or (ct.most_common(1)[0][0] if ct else "") or pdf["chi_nhanh"]
    ten_pgd = _val(info, "pgd") or pdf["pgd"]
    thieu = [n for n, v in (("Xã", ten_xa), ("Tỉnh", ten_tinh), ("PGD", ten_pgd)) if not v]
    if thieu:
        raise ValueError("Không xác định được: " + ", ".join(thieu) + " (Excel trống và PDF không có).")

    ngay_tb = format_ngay(info.get("ngày thông báo"), "Ngày thông báo")
    ngay_pt = format_ngay(info.get("ngày phát tiền"), "Ngày phát tiền")

    groups = {}
    for r in rows:
        groups.setdefault((r["Mã Tổ"], r["Tổ"]), []).append(r)
    to_list = []
    for (ma_to, to), rs in groups.items():
        thons = Counter(thon_from_address(r["Địa chỉ"], ten_xa, ten_tinh) for r in rs)
        thons.pop("", None)
        if len(thons) > 1:
            warnings.append(f"Tổ '{to}' có khách ở nhiều thôn ({', '.join(thons)}); "
                            f"Mẫu 06 lấy thôn phổ biến nhất: {thons.most_common(1)[0][0]}.")
        if any(not r["Tên KH"] or not r["Địa chỉ"] for r in rs):
            warnings.append(f"Tổ '{to}': có khách thiếu tên hoặc địa chỉ.")
        if any(r["Tiền"] % 1_000_000 for r in rs):
            warnings.append(f"Tổ '{to}': có khoản không tròn triệu (Mẫu 06 sẽ ghi số lẻ).")
        to_list.append({"ma_to": ma_to, "to": to, "rows": rs, "tong": sum(r["Tiền"] for r in rs),
                        "thon": thons.most_common(1)[0][0] if thons else pdf["diem"]})

    return {"pdf": pdf, "diem": pdf["diem"], "ma_diem": pdf["ma_diem"],
            "excel_diem": _val(info, COL_DIEM), "xa": ten_xa, "tinh": ten_tinh, "pgd": ten_pgd,
            "ngay_tb": ngay_tb, "ngay_pt": ngay_pt, "to_list": to_list, "rows": rows,
            "tong": sum(r["Tiền"] for r in rows), "warnings": warnings}


# ================== GHI WORD ==================
def replace_in_paragraph(p, mapping):
    """Thay placeholder, kể cả khi Word tách {{KEY}} thành nhiều run. Giữ định dạng run đầu tiên."""
    for key, value in mapping.items():
        value, start = str(value), 0
        while True:
            runs = p.runs
            lens = [len(r.text) for r in runs]
            idx = "".join(r.text for r in runs).find(key, start)
            if idx < 0:
                break
            end, pos, first = idx + len(key), 0, True
            for r, ln in zip(runs, lens):
                r_start, r_end = pos, pos + ln
                pos = r_end
                if r_end <= idx or r_start >= end:
                    continue
                s, e = max(idx, r_start) - r_start, min(end, r_end) - r_start
                if first:
                    r.text = r.text[:s] + value + r.text[e:]
                    first = False
                else:
                    r.text = r.text[:s] + r.text[e:]
            start = idx + len(value)


def _iter_table_paragraphs(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for t in cell.tables:
                yield from _iter_table_paragraphs(t)


def iter_paragraphs(doc):
    """Thân văn bản + bảng (lồng nhau) + header/footer. (Text box chưa được hỗ trợ.)"""
    yield from doc.paragraphs
    for t in doc.tables:
        yield from _iter_table_paragraphs(t)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            if part.is_linked_to_previous:     # tránh vô tình tạo header mới
                continue
            yield from part.paragraphs
            for t in part.tables:
                yield from _iter_table_paragraphs(t)


def replace_everywhere(doc, mapping):
    for p in iter_paragraphs(doc):
        replace_in_paragraph(p, mapping)


def save_docx(doc, path, used):
    base, ext = os.path.splitext(path)
    n = 2
    while path in used:                        # trùng tên TRONG LẦN CHẠY NÀY -> thêm số
        path = f"{base} ({n}){ext}"
        n += 1
    used.add(path)
    try:
        doc.save(path)
    except PermissionError:
        raise ValueError(f"Không ghi được file (đang mở trong Word?):\n{path}")


def generate_files(ctx, form, base_dir):
    date_folder = datetime.now().strftime("%d-%m-%Y")
    out_04 = os.path.join(base_dir, date_folder, "Mau_04")
    out_06 = os.path.join(base_dir, date_folder, "Mau_06")
    used, c04, c06, notes = set(), 0, 0, []

    if form["chuc_vu"] == "Phó Giám đốc":
        chu_ky, title_lap = "KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC", "\nLẬP BẢNG"
    else:
        chu_ky, title_lap = "GIÁM ĐỐC", "LẬP BẢNG"

    # ---------- Mẫu 04: mỗi (Tổ, Chương trình) một file ----------
    if form["file_04"]:
        os.makedirs(out_04, exist_ok=True)
        for t in ctx["to_list"]:
            by_ct = {}
            for r in t["rows"]:
                by_ct.setdefault(r["Chương trình"], []).append(r)
            for ct, rs in by_ct.items():
                tong = sum(r["Tiền"] for r in rs)
                mapping = {
                    "{{NGAY_THONG_BAO}}": ctx["ngay_tb"], "{{TEN_XA}}": ctx["xa"],
                    "{{TEN_PGD_UPPER}}": ctx["pgd"].upper(), "{{TEN_TINH_UPPER}}": ctx["tinh"].upper(),
                    "{{TEN_PGD}}": ctx["pgd"], "{{TEN_TINH}}": ctx["tinh"],
                    "{{SO_HO}}": f"{len(rs):02d}", "{{SO_TIEN}}": fmt_money(tong),
                    "{{SO_TIEN_CHU}}": doc_so_thanh_chu(tong), "{{NGAY_PHAT_TIEN}}": ctx["ngay_pt"],
                    "{{DIEM_GIAO_DICH}}": ctx["diem"], "{{NGUOI_LAP_BANG}}": form["nguoi_lap"],
                    "{{TITLE_LAP}}": title_lap, "{{CHUC_VU}}": chu_ky,
                    "{{TEN_LANH_DAO}}": form["ten_lanh_dao"],
                }
                doc = Document(form["file_04"])
                replace_everywhere(doc, mapping)
                name = f"{_safe_name(t['to'])} - {_safe_name(ctx['diem'])} - {_safe_name(ct)}.docx"
                save_docx(doc, os.path.join(out_04, name), used)
                c04 += 1

    # ---------- Mẫu 06: mỗi tổ một file ----------
    if form["file_06"]:
        os.makedirs(out_06, exist_ok=True)
        for t in ctx["to_list"]:
            rs, tong_str = t["rows"], fmt_trieu(t["tong"])
            mapping = {"{{DIA_BAN}}": f"{t['thon']} - {ctx['xa']} - {ctx['tinh']}",
                       "{{TO_TKVV}}": f"{t['to']} ({t['ma_to']})", "{{TONG_TIEN}}": tong_str}
            doc = Document(form["file_06"])
            found = False
            for table in doc.tables:
                idx = next((i for i, r in enumerate(table.rows)
                            if any("{{TEN_KH}}" in c.text for c in r.cells)), -1)
                if idx < 0:
                    continue
                found = True
                tpl = table.rows[idx]
                for _ in range(len(rs) - 1):
                    tpl._tr.addnext(deepcopy(tpl._tr))
                data_rows = list(table.rows)[idx: idx + len(rs)]
                for i, (row_tb, kh) in enumerate(zip(data_rows, rs), 1):
                    data = {"{{STT}}": i, "{{TEN_KH}}": kh["Tên KH"],
                            "{{MA_KV}}": f"66-{kh['Mã KV'][-4:]}", "{{MA_CT}}": kh["Chương trình"],
                            "{{TIEN}}": fmt_trieu(kh["Tiền"])}
                    for cell in row_tb.cells:
                        for p in cell.paragraphs:
                            replace_in_paragraph(p, data)
                # điền ô tổng ở dòng "Cộng" nếu còn trống
                for row_tb in table.rows:
                    cells = row_tb.cells
                    first = " ".join(c.text.strip().lower() for c in cells[:2])
                    if "cộng" in first and len(cells) >= 5:
                        for ci in (4, 5):
                            if ci < len(cells) and not cells[ci].text.strip():
                                cells[ci].paragraphs[0].text = tong_str
            if not found:
                notes.append(f"Mẫu 06 không có dòng chứa {{{{TEN_KH}}}} - tổ '{t['to']}' chưa có bảng khách hàng.")
            replace_everywhere(doc, mapping)
            name = f"Mau06_{_safe_name(t['to'])}_{t['ma_to']}.docx"
            save_docx(doc, os.path.join(out_06, name), used)
            c06 += 1
    return c04, c06, notes, os.path.join(base_dir, date_folder)


# ================== GIAO DIỆN ==================
class ThongBaoGiaiNganApp:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        if isinstance(root, tk.Tk):                 # chỉ đặt title/size khi chạy độc lập
            ui.apply_theme(root)
            root.title("Tool Tạo Thông Báo & Mẫu 06/TD (V18 - Mẫu 05/GDX mới)")
            root.geometry("820x680")

        main = tk.Frame(root, bg=ui.BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)
        ui.label(main, "Thông báo giải ngân & Mẫu 06/TD", bold=True, size=15).pack(anchor="w")
        ui.label(main, "V18 - Mẫu 05/GDX mới", muted=True).pack(anchor="w", pady=(0, 10))

        # ---- Khung 1: người lập / lãnh đạo ----
        card1 = ui.Card(main, title="Thông tin người lập")
        card1.pack(fill="x")
        f = card1.content
        f.columnconfigure(1, weight=1)

        ui.label(f, "Người lập bảng:").grid(row=0, column=0, sticky="w", pady=5, padx=(0, 12))
        self.entry_nguoi_lap = ttk.Entry(f)
        self.entry_nguoi_lap.grid(row=0, column=1, sticky="ew", pady=5)
        self.entry_nguoi_lap.bind("<FocusOut>", self.auto_capitalize)

        ui.label(f, "Chức vụ Lãnh đạo:").grid(row=1, column=0, sticky="w", pady=5, padx=(0, 12))
        self.chuc_vu_var = tk.StringVar(master=main, value="Giám đốc")
        ttk.Combobox(f, textvariable=self.chuc_vu_var, values=["Giám đốc", "Phó Giám đốc"],
                     state="readonly").grid(row=1, column=1, sticky="ew", pady=5)

        ui.label(f, "Tên lãnh đạo:").grid(row=2, column=0, sticky="w", pady=5, padx=(0, 12))
        self.entry_ten_lanh_dao = ttk.Entry(f)
        self.entry_ten_lanh_dao.grid(row=2, column=1, sticky="ew", pady=5)
        self.entry_ten_lanh_dao.bind("<FocusOut>", self.auto_capitalize)

        # ---- Khung 2: nguồn file ----
        card2 = ui.Card(main, title="Nguồn file")
        card2.pack(fill="x", pady=(10, 0))
        g = card2.content
        g.columnconfigure(1, weight=1)

        self.excel_path_var = tk.StringVar(master=main)
        self.pdf_path_var = tk.StringVar(master=main)
        self.template_04_var = tk.StringVar(master=main)
        self.template_06_var = tk.StringVar(master=main)
        rows = [("Dữ liệu điểm giao dịch - excel:", self.excel_path_var, "Excel"),
                ("Danh sách phê duyệt giải ngân - pdf:", self.pdf_path_var, "PDF"),
                ("Mẫu thông báo phê duyệt - word:", self.template_04_var, "Word"),
                ("Mẫu kiểm tra sau cho vay - word:", self.template_06_var, "Word")]
        for i, (label, var, ft) in enumerate(rows):
            ui.label(g, label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 12))
            ttk.Entry(g, textvariable=var, state="readonly").grid(row=i, column=1, sticky="ew", pady=5)
            ui.RoundButton(g, text="Chọn", command=lambda v=var, t=ft: self.browse_file(v, t),
                           padx=14, pady=5, size=9).grid(row=i, column=2, padx=(8, 0))

        # Thư mục xuất file (mặc định: "Thong bao phe duyet" cạnh chương trình)
        r = len(rows)
        self.out_dir_var = tk.StringVar(master=main, value=paths.output_dir("thongbao"))
        ui.label(g, "Thư mục xuất file:").grid(row=r, column=0, sticky="w", pady=5, padx=(0, 12))
        ttk.Entry(g, textvariable=self.out_dir_var, state="readonly").grid(row=r, column=1, sticky="ew", pady=5)
        out_btns = tk.Frame(g, bg=ui.CARD)
        out_btns.grid(row=r, column=2, padx=(8, 0))
        ui.RoundButton(out_btns, text="Chọn", command=self.browse_out_dir, padx=14, pady=5,
                       size=9).pack(side="left")
        ui.secondary_button(out_btns, "Mở", lambda: paths.open_folder(self.out_dir_var.get()),
                            padx=12, pady=5, size=9).pack(side="left", padx=(6, 0))

        # ---- Nút chạy ----
        bar = tk.Frame(main, bg=ui.BG)
        bar.pack(pady=14)
        self.btn_run = ui.RoundButton(bar, text="TẠO BÁO CÁO (04 & 06)", command=self.start,
                                      padx=26, pady=10, size=11)
        self.btn_run.pack(side="left", padx=8)
        ui.secondary_button(bar, "Chẩn đoán PDF", self.diagnose, padx=20, pady=10,
                            size=11).pack(side="left", padx=8)
        self.status = tk.StringVar(master=main)
        tk.Label(main, textvariable=self.status, bg=ui.BG, fg=ui.MUTED, font=ui.font(10)).pack()

        self.load_config()

    # ---- tiện ích giao diện ----
    def auto_capitalize(self, event):
        w = event.widget
        val = w.get().strip()
        if val:
            w.delete(0, tk.END)
            w.insert(0, _title(val))

    def load_config(self):
        d = paths.load_section(CONFIG_SECTION)
        if not d:
            return
        self.entry_nguoi_lap.insert(0, d.get("nguoi_lap", ""))
        self.chuc_vu_var.set(d.get("chuc_vu", "Giám đốc"))
        self.entry_ten_lanh_dao.insert(0, d.get("ten_lanh_dao", ""))
        self.excel_path_var.set(d.get("excel_path", ""))
        self.template_04_var.set(d.get("template_04", ""))
        self.template_06_var.set(d.get("template_06", ""))
        if d.get("out_dir"):
            self.out_dir_var.set(d["out_dir"])

    def save_config(self):
        d = {"nguoi_lap": self.entry_nguoi_lap.get().strip(), "chuc_vu": self.chuc_vu_var.get(),
             "ten_lanh_dao": self.entry_ten_lanh_dao.get().strip(),
             "excel_path": self.excel_path_var.get(), "template_04": self.template_04_var.get(),
             "template_06": self.template_06_var.get(), "out_dir": self.out_dir_var.get()}
        paths.save_section(CONFIG_SECTION, d)

    def browse_out_dir(self):
        cur = self.out_dir_var.get()
        folder = filedialog.askdirectory(title="Chọn thư mục xuất file",
                                         initialdir=cur if os.path.isdir(cur) else paths.output_dir("thongbao"))
        if folder:
            self.out_dir_var.set(os.path.normpath(folder))
            self.save_config()

    def browse_file(self, var, ftype):
        types = {"Excel": [("Excel files", "*.xlsx *.xls")], "PDF": [("PDF files", "*.pdf")],
                 "Word": [("Word files", "*.docx")]}[ftype]
        path = filedialog.askopenfilename(title=f"Chọn file {ftype}", filetypes=types)
        if path:
            var.set(path)

    def diagnose(self):
        if not self.pdf_path_var.get():
            messagebox.showwarning("Thiếu file", "Hãy chọn file PDF trước.")
            return
        try:
            out = dump_pdf_debug(self.pdf_path_var.get())
            if messagebox.askyesno("Chẩn đoán PDF", f"Đã ghi nội dung PDF đọc được ra:\n{out}\n\nMở file ngay?") \
                    and hasattr(os, "startfile"):
                os.startfile(out)
        except Exception as e:
            logging.exception("diagnose")
            messagebox.showerror("Lỗi", str(e))

    # ---- luồng xử lý ----
    def start(self):
        for ent in (self.entry_nguoi_lap, self.entry_ten_lanh_dao):
            v = ent.get().strip()
            ent.delete(0, tk.END)
            ent.insert(0, _title(v))
        form = {"nguoi_lap": self.entry_nguoi_lap.get().strip(), "chuc_vu": self.chuc_vu_var.get(),
                "ten_lanh_dao": self.entry_ten_lanh_dao.get().strip(),
                "excel": self.excel_path_var.get(), "pdf": self.pdf_path_var.get(),
                "file_04": self.template_04_var.get(), "file_06": self.template_06_var.get(),
                "out_dir": self.out_dir_var.get().strip() or paths.output_dir("thongbao")}
        if not form["excel"] or not form["pdf"]:
            messagebox.showwarning("Lỗi", "Vui lòng chọn File Excel Dữ liệu và Danh sách phê duyệt giải ngân!")
            return
        if not form["file_04"] and not form["file_06"]:
            messagebox.showwarning("Lỗi", "Hãy chọn ít nhất một mẫu Word (04 hoặc 06).")
            return
        self.save_config()
        self.btn_run.config(state="disabled")
        self.status.set("Đang đọc PDF và Excel...")
        threading.Thread(target=self._worker, args=(form,), daemon=True).start()
        self.root.after(100, self._poll)

    def _worker(self, form):
        try:
            self.q.put(("ok", build_context(form["pdf"], form["excel"]), form))
        except Exception as e:
            logging.exception("build_context")
            self.q.put(("err", str(e)))

    def _poll(self):
        try:
            item = self.q.get_nowait()
        except queue.Empty:
            self.root.after(100, self._poll)
            return
        self.btn_run.config(state="normal")
        if item[0] == "err":
            self.status.set("")
            messagebox.showerror("Lỗi thực thi", item[1])
            return
        self._confirm_and_generate(item[1], item[2])

    def _confirm_and_generate(self, ctx, form):
        lines = [f"Điểm giao dịch (PDF): {ctx['diem']} {('(' + ctx['ma_diem'] + ')') if ctx['ma_diem'] else ''}",
                 f"Khớp dòng Excel: {ctx['excel_diem']}  |  Xã: {ctx['xa']}  |  Tỉnh: {ctx['tinh']}  |  PGD: {ctx['pgd']}",
                 f"Ngày thông báo: {ctx['ngay_tb']}", f"Ngày phát tiền: {ctx['ngay_pt']}", "",
                 f"{len(ctx['to_list'])} tổ - {len(ctx['rows'])} khách - Tổng {fmt_money(ctx['tong'])} đồng", ""]
        for t in ctx["to_list"]:
            lines.append(f"• {t['to']} ({t['ma_to'] or '?'}): {len(t['rows'])} hộ - {fmt_money(t['tong'])}")
        if ctx["warnings"]:
            lines += ["", "CẢNH BÁO:"] + [f"! {w}" for w in ctx["warnings"]]
        lines += ["", "Tiếp tục tạo file?"]
        if not messagebox.askyesno("Kiểm tra trước khi xuất", "\n".join(lines)):
            self.status.set("Đã hủy.")
            return
        self.status.set("Đang tạo file Word...")
        self.root.update_idletasks()
        try:
            c04, c06, notes, out_root = generate_files(ctx, form, form["out_dir"])
            self.status.set("")
            msg = f"Thành công!\n- Tạo {c04} file Mẫu 04\n- Tạo {c06} file Mẫu 06\n\nĐã lưu trong:\n{out_root}"
            if notes:
                msg += "\n\nLưu ý:\n" + "\n".join(notes)
            if messagebox.askyesno("Hoàn tất", msg + "\n\nMở thư mục chứa file ngay?"):
                paths.open_folder(out_root)
        except Exception as e:
            logging.exception("generate_files")
            self.status.set("")
            messagebox.showerror("Lỗi thực thi", str(e))


# ===== Khai báo cho launcher =====
TOOL_NAME = "Thông báo giải ngân"
TOOL_CLASS = ThongBaoGiaiNganApp
TOOL_ORDER = 10


if __name__ == "__main__":
    root = tk.Tk()
    ThongBaoGiaiNganApp(root)
    root.mainloop()