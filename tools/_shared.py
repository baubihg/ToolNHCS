# -*- coding: utf-8 -*-
"""
Kho dữ liệu SAO KÊ và DANH SÁCH HỘ NGHÈO dùng chung cho launcher và mọi tool.

Cách dùng trong một tool:
    from tools._shared import store          # (trong tool: try ". import _shared" như _ui)
    store.records        # list[dict] các dòng sao kê, khóa = tên cột (rỗng nếu chưa nạp)
    store.rows           # list[list[str]] cùng các dòng đó nhưng theo VỊ TRÍ cột (A=0, B=1, ...)
    store.ngay           # "Ngày số liệu" của sao kê
    store.subscribe(fn)  # fn() được gọi mỗi khi sao kê được nạp mới / xóa
    store.find_customers("010190001234")   # tìm theo CCCD / mã KH / số khế ước / tên
    store.phan_loai_ho("010190001234")     # "Hộ nghèo" / "Hộ cận nghèo" / "" (theo DS hộ nghèo đã nạp)

Dữ liệu lưu trong thư mục dữ liệu chung của chương trình (xem _paths.py), nên lần sau mở launcher có dữ liệu ngay.
File bắt đầu bằng "_" nên launcher không coi đây là một tool.
"""
import logging
import numbers
import os
import pickle
import re
import unicodedata
from datetime import date, datetime

import pandas as pd

try:
    from . import _paths as paths
except ImportError:          # chạy trực tiếp, không qua launcher
    import _paths as paths

SAOKE_FILE = paths.data_path("saoke.pkl")
HONGHEO_FILE = paths.data_path("hongheo.pkl")
LEGACY_FILE = paths.data_path("tracuu_legacy.pkl")           # bản cũ lưu sao kê + hộ nghèo chung 1 file (đọc xong thì xóa)

# Tên cột có thể gặp trong file sao kê
K_TEN = ("Tên KH", "Tên khách hàng", "Họ tên")
K_CCCD = ("Số CMND", "Số CCCD", "CMND", "CCCD")
K_MAKH = ("Mã KH", "Mã khách hàng")
K_KHEUOC = ("Số khế ước", "Mã món vay")
K_NGAYSINH = ("Ngày sinh",)
K_NGAYCAP = ("Ngày cấp CMND", "Ngày cấp")
K_NOICAP = ("Nơi cấp CMND", "Nơi cấp")
K_SDT = ("Số điện thoại", "Điện thoại", "SĐT")
K_XA = ("Tên xã", "Xã")
K_DIACHI = ("Địa chỉ",)


# ================== TIỆN ÍCH ==================
def remove_accents(s):
    """Bỏ dấu tiếng Việt + ký tự đặc biệt, chữ thường."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("đ", "d").replace("Đ", "D")
    return re.sub(r"[^a-z0-9]", "", s.lower())


def cell_to_str(v):
    """Đổi giá trị ô Excel sang chuỗi (số nguyên không còn '.0', ngày -> dd/mm/yyyy)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, numbers.Integral):
        return str(int(v))
    if isinstance(v, numbers.Real):
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    return unicodedata.normalize("NFC", str(v))


def clean_header(h):
    return re.sub(r"[\n\r\t]", "", cell_to_str(h)).strip()


def row_empty(values):
    return all(v == "" for v in values)


def get_field(row, *keys):
    """Giá trị không rỗng đầu tiên trong các tên cột đưa vào."""
    for k in keys:
        v = row.get(k)
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return ""


# ================== ĐỌC FILE SAO KÊ ==================
def read_sao_ke(path):
    """Đọc file sao kê -> (list[dict theo tên cột], ngày số liệu, list[list theo vị trí cột])."""
    raw = pd.read_excel(path, header=None, dtype=object)
    raw_rows = raw.values.tolist()

    header_idx = -1
    for i in range(min(20, len(raw_rows))):
        s = "".join(cell_to_str(c) for c in raw_rows[i]).lower()
        if "tên kh" in s and "số khế ước" in s:
            header_idx = i
            break
    if header_idx == -1:
        raise ValueError("Không tìm thấy dòng tiêu đề (cần có cột 'Tên KH' và 'Số khế ước').")

    headers = [clean_header(h) for h in raw_rows[header_idx]]
    data, rows = [], []
    for r in raw_rows[header_idx + 1:]:
        vals = [cell_to_str(c) for c in r]
        if row_empty(vals):
            continue
        rows.append(vals)           # giữ nguyên vị trí cột (cột A = 0, B = 1, ...)
        data.append({h: vals[i] for i, h in enumerate(headers) if h and i < len(vals)})

    ngay = str(data[0].get("Ngày số liệu", "")).strip() if data else ""
    return data, ngay, rows


def read_ho_ngheo(path):
    """Đọc file danh sách hộ nghèo (tiêu đề ở dòng đầu) -> list[dict]."""
    raw = pd.read_excel(path, header=0, dtype=object)
    headers = [clean_header(h) for h in raw.columns]
    data = []
    for r in raw.values.tolist():
        vals = [cell_to_str(c) for c in r]
        if row_empty(vals):
            continue
        data.append({h: vals[i] for i, h in enumerate(headers) if h})
    return data


def build_map_ho_ngheo(rows):
    """{CCCD bỏ số 0 đầu: phân loại hộ}"""
    m = {}
    for row in rows:
        cccd = get_field(row, "Số CMND/Giấy KS", "Số CMND", "CCCD")
        phan_loai = get_field(row, "Phân loại hộ", "Phân Loại Hộ")
        if cccd:
            m[cccd.lstrip("0")] = phan_loai
    return m


# ================== KHO DỮ LIỆU ==================
class SaoKeStore:
    def __init__(self):
        self.records = []
        self.rows = []
        self.ngay = ""
        self.ho_ngheo = []
        self._ho_map = {}
        self._listeners = []
        self._load_saved()

    # ---- đăng ký nhận thông báo ----
    def subscribe(self, fn):
        if fn not in self._listeners:
            self._listeners.append(fn)
        return fn

    def unsubscribe(self, fn):
        if fn in self._listeners:
            self._listeners.remove(fn)

    def _notify(self):
        for fn in list(self._listeners):
            try:
                fn()
            except Exception:
                logging.exception("Lỗi khi cập nhật giao diện sau khi đổi sao kê")

    # ---- lưu / đọc ----
    def _read_pickle(self, path):
        with open(path, "rb") as fh:
            return pickle.load(fh)

    def _write_pickle(self, path, obj):
        try:
            with open(paths.ensure_parent(path), "wb") as fh:
                pickle.dump(obj, fh, protocol=pickle.HIGHEST_PROTOCOL)
        except OSError as e:
            logging.warning("Không lưu được %s: %s", os.path.basename(path), e)

    def _load_saved(self):
        try:
            if os.path.exists(SAOKE_FILE):
                d = self._read_pickle(SAOKE_FILE)
                self.records, self.ngay = d.get("sao_ke", []), d.get("ngay", "")
                self.rows = d.get("rows", [])
            if os.path.exists(HONGHEO_FILE):
                self.ho_ngheo = self._read_pickle(HONGHEO_FILE).get("ho_ngheo", [])
            if os.path.exists(LEGACY_FILE):                  # chuyển dữ liệu từ bản cũ (nếu còn thiếu)
                d = self._read_pickle(LEGACY_FILE)
                if not self.records and d.get("sao_ke"):
                    self.records, self.ngay = d["sao_ke"], d.get("ngay", "")
                    self._save()
                if not self.ho_ngheo and d.get("ho_ngheo"):
                    self.ho_ngheo = d["ho_ngheo"]
                    self._save_ho_ngheo()
                os.remove(LEGACY_FILE)                       # đã chuyển sang file mới, bỏ file cũ cho gọn
        except Exception as e:
            logging.warning("Không đọc được dữ liệu dùng chung đã lưu: %s", e)
        self._ho_map = build_map_ho_ngheo(self.ho_ngheo)

    def _save(self):
        self._write_pickle(SAOKE_FILE, {"sao_ke": self.records, "ngay": self.ngay, "rows": self.rows})

    def _save_ho_ngheo(self):
        self._write_pickle(HONGHEO_FILE, {"ho_ngheo": self.ho_ngheo})

    # ---- thao tác ----
    def load_file(self, path):
        """Nạp file sao kê mới (lỗi thì giữ nguyên dữ liệu cũ và ném exception)."""
        data, ngay, rows = read_sao_ke(path)
        self.records, self.ngay, self.rows = data, ngay, rows
        self._save()
        self._notify()

    def clear(self):
        self.records, self.ngay, self.rows = [], "", []
        try:
            if os.path.exists(SAOKE_FILE):
                os.remove(SAOKE_FILE)
        except OSError as e:
            logging.warning("Không xóa được file sao kê: %s", e)
        self._notify()

    def load_ho_ngheo_file(self, path):
        """Nạp file danh sách hộ nghèo (lỗi thì giữ nguyên dữ liệu cũ và ném exception)."""
        data = read_ho_ngheo(path)
        self.ho_ngheo = data
        self._ho_map = build_map_ho_ngheo(data)
        self._save_ho_ngheo()
        self._notify()

    def clear_ho_ngheo(self):
        self.ho_ngheo, self._ho_map = [], {}
        try:
            if os.path.exists(HONGHEO_FILE):
                os.remove(HONGHEO_FILE)
        except OSError as e:
            logging.warning("Không xóa được file hộ nghèo: %s", e)
        self._notify()

    def phan_loai_ho(self, cccd):
        """Phân loại hộ (vd 'Hộ nghèo') của một CCCD, '' nếu không có trong danh sách."""
        return self._ho_map.get(str(cccd or "").strip().lstrip("0"), "")

    def value_at(self, record, col):
        """Giá trị ở cột số `col` (A=0, B=1, ...) của một dòng trong store.records. '' nếu không có."""
        if not self.has_positions:
            return ""
        for i, r in enumerate(self.records):
            if r is record:
                row = self.rows[i]
                return row[col] if col < len(row) else ""
        return ""

    @property
    def has_positions(self):
        """True nếu có dữ liệu theo vị trí cột (sao kê nạp bằng phiên bản mới)."""
        return bool(self.records) and len(self.rows) == len(self.records)

    def summary(self):
        if not self.records:
            return "Chưa nạp sao kê"
        txt = f"{len(self.records):,} bản ghi".replace(",", ".")
        return txt + (f" · ngày {self.ngay}" if self.ngay else "")

    def ho_ngheo_summary(self):
        if not self.ho_ngheo:
            return "Chưa nạp hộ nghèo"
        return f"{len(self.ho_ngheo):,} nhân khẩu".replace(",", ".")

    # ---- tra cứu nhanh cho các tool ----
    def find_customers(self, query, limit=50):
        """Tìm theo CCCD / mã KH / số khế ước (khớp đúng) hoặc tên (khớp một phần).
        Mỗi khách hàng chỉ trả về một dòng dù có nhiều khoản vay."""
        q = (query or "").strip()
        nq = remove_accents(q)
        if not nq:
            return []
        digits = re.sub(r"\D", "", q)
        by_name = not q.replace(" ", "").isdigit()
        out, seen = [], set()
        for r in self.records:
            cccd = get_field(r, *K_CCCD)
            ma_kh, khe_uoc = get_field(r, *K_MAKH), get_field(r, *K_KHEUOC)
            ten = get_field(r, *K_TEN)
            hit = (nq in (remove_accents(ma_kh), remove_accents(khe_uoc))
                   or (digits and cccd and cccd.lstrip("0") == digits.lstrip("0"))
                   or (by_name and nq in remove_accents(ten)))
            if not hit:
                continue
            key = cccd.lstrip("0") or ma_kh or (ten + get_field(r, *K_DIACHI))
            if key in seen:
                continue
            seen.add(key)
            out.append(r)
            if len(out) >= limit:
                break
        return out


store = SaoKeStore()
