"""
Kiểm tra dữ liệu tín dụng trên SAO KÊ dùng chung (nạp ở khung "Dữ liệu dùng chung" của launcher).
Các cột được đọc theo VỊ TRÍ cột Excel (xem COL bên dưới: F=5, H=7, ...).
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from datetime import datetime
import os

try:
    from . import _ui as ui
    from ._shared import store
except ImportError:
    import _ui as ui
    from _shared import store

# ===================== ÁNH XẠ CỘT (0-based) =====================
COL = {
    "ten_xa":          5,   # F
    "ten_thon":        7,   # H
    "ma_kh":           8,   # I
    "ten_kh":          9,   # J
    "ten_dvut":        27,  # AB - Tên ĐVUT
    "so_khe_uoc":      28,  # AC
    "ngay_sinh":       10,  # K
    "gioi_tinh":       13,  # N
    "so_cmnd":         16,  # Q
    "noi_cap":         17,  # R
    "dien_thoai":      21,  # V
    "ngay_vay":        29,  # AD
    "ngay_dh_hop_dong":30,  # AE
    "thoi_han_vay":    33,  # AH
    "tinh_trang":      40,  # AO
    "du_no_qua_han":   58,  # BG
    "lai_ton_th":      66,  # BO
    "nguon_von":       48,  # AW
    "ma_ndt":         135,  # EF
    "ma_to":           22,  # W
    "ten_to_truong":   25,  # Z
    "du_no":           60,  # BI
    "ton_rpa":        104,  # DA
}

HINT_DETAIL = "Bấm vào mục lỗi bên trái để xem chi tiết  |  Bấm đúp vào dòng để xem đầy đủ"

DISPLAY_COLS = ["ten_xa", "ten_thon", "ten_kh", "ma_kh", "ten_dvut", "so_khe_uoc"]

LABELS = {
    "ten_xa":     "Tên xã", "ten_thon":   "Tên thôn", "ten_kh":     "Tên KH",
    "ma_kh":      "Mã KH", "ten_dvut":   "Tên hội đoàn thể", "so_khe_uoc": "Số khế ước",
    "Lỗi":        "Lỗi", "tinh_trang": "Tình trạng (AO)", "ngay_vay":   "Ngày vay",
    "ngay_dh_hop_dong": "Ngày ĐH HĐ", "thoi_han_vay": "Thời hạn vay", "so_thang":   "Số tháng",
    "lai_ton_th": "Lãi tồn TH", "du_no_qua_han": "Dư nợ QH", "noi_cap":    "Nơi cấp",
    "dien_thoai": "Số điện thoại", "gioi_tinh":  "Giới tính", "so_cmnd":    "Số CMND",
    "ngay_sinh":  "Ngày sinh", "nguon_von":  "Nguồn vốn (AW)", "ma_ndt":     "Mã NĐT (EF)",
    "ma_to":      "Mã tổ", "ten_to_truong": "Tên tổ trưởng", "du_no":      "Dư nợ (BI)",
    "ton_rpa":    "Tồn RPA (DA)", "so_thanh_vien": "Số thành viên",
}

def build_dataframe(rows):
    """rows: list[list[str]] theo vị trí cột (store.rows) -> DataFrame với các cột đặt tên theo COL."""
    data = pd.DataFrame({name: [r[idx] if idx < len(r) else "" for r in rows]
                         for name, idx in COL.items()})
    mask = data["so_khe_uoc"].apply(safe_str) != "Số khế ước"      # bỏ dòng tiêu đề lọt vào (nếu có)
    return data[mask].reset_index(drop=True)


# ===================== TIỆN ÍCH =====================
def safe_str(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    if pd.isna(v): return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit(): s = s[:-2]
    return s

def to_date(v):
    if pd.isna(v): return pd.NaT
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.to_pydatetime() if isinstance(v, pd.Timestamp) else v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try: return datetime.strptime(s, fmt)
        except ValueError: continue
    return pd.NaT

def months_between(d1, d2):
    if pd.isna(d1) or pd.isna(d2): return np.nan
    m = (d2.year - d1.year) * 12 + (d2.month - d1.month)
    if d2.day < d1.day: m -= 1
    return m

def is_open(v):
    return safe_str(v).upper() == "OPEN"

def to_num(v):
    if pd.isna(v): return 0.0
    try: return float(str(v).replace(",", "").replace(" ", "").strip() or 0)
    except ValueError: return 0.0

def fmt_num(v):
    try: return f"{float(v):,.0f}"
    except (ValueError, TypeError): return str(v) if v is not None else ""

def _ma_gioi_tinh_cccd(nam_sinh, gioi_tinh):
    if pd.isna(nam_sinh): return None
    y = nam_sinh.year
    if 1900 <= y <= 1999: base = 0
    elif 2000 <= y <= 2099: base = 2
    elif 2100 <= y <= 2199: base = 4
    elif 2200 <= y <= 2299: base = 6
    elif 2300 <= y <= 2399: base = 8
    else: return None
    
    if gioi_tinh == "01": return str(base)
    elif gioi_tinh == "02": return str(base + 1)
    return None

def _phan_loai_theo_thang(thang):
    if pd.isna(thang): return None
    if thang <= 12: return "Ngắn hạn"
    if thang <= 60: return "Trung hạn"
    return "Dài hạn"

# ===================== CÁC HÀM KIỂM TRA =====================
def _base_cols(df, mask):
    return df.loc[mask, DISPLAY_COLS].copy()

def loi_1_lai_ton(df):
    m = (df["lai_ton_th"].apply(to_num) > 1_000_000) & df["tinh_trang"].apply(is_open) & (df["du_no_qua_han"].apply(to_num) == 0)
    out = _base_cols(df, m)
    if len(out) == 0: return out
    out["Lãi tồn TH"] = df.loc[m, "lai_ton_th"].apply(fmt_num).values
    out["Dư nợ QH"] = df.loc[m, "du_no_qua_han"].apply(fmt_num).values
    out["Lỗi"] = "Lãi tồn > 1 triệu nhưng dư nợ QH = 0"
    return out

def loi_2_ngay_den_han(df):
    ad = df["ngay_vay"].apply(safe_str)
    ae = df["ngay_dh_hop_dong"].apply(safe_str)
    m = (ad != "") & (ae != "") & (ad.str[:2] != ae.str[:2]) & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, m)
    if len(out) == 0: return out
    out["Ngày vay"] = ad[m].values
    out["Ngày ĐH HĐ"] = ae[m].values
    out["Lỗi"] = "Ngày đến hạn khác ngày vay (2 ký tự đầu)"
    return out

def loi_3_noi_cap(df):
    allowed = {"BỘ CÔNG AN", "CỤC CSQLHC VỀ TTXH"}
    r = df["noi_cap"].apply(safe_str).str.upper()
    ac = df["so_khe_uoc"].apply(safe_str)
    m = (~r.isin(allowed)) & (ac != "") & (ac != "Số khế ước") & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, m)
    if len(out) == 0: return out
    out["Nơi cấp"] = r[m].values
    out["Lỗi"] = "Nơi cấp không phải BỘ CÔNG AN / CỤC CSQLHC VỀ TTXH"
    return out

def loi_4_dien_thoai(df):
    v = df["dien_thoai"].apply(safe_str)
    m = (v.str.len() > 10) & (v != "") & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, m)
    if len(out) == 0: return out
    out["Số điện thoại"] = v[m].values
    out["Lỗi"] = "Số điện thoại > 10 ký tự"
    return out

def loi_5_gioi_tinh(df):
    gt = df["gioi_tinh"].apply(safe_str)
    cmnd = df["so_cmnd"].apply(safe_str)
    ns = df["ngay_sinh"].apply(to_date)

    def bad(i):
        g = gt.iloc[i]
        c = cmnd.iloc[i]
        d = ns.iloc[i]
        if len(c) != 12 or pd.isna(d): return False
        ma_ky_vong = _ma_gioi_tinh_cccd(d, g)
        if ma_ky_vong is None: return False
        chu_so_4 = c[3]
        if not chu_so_4.isdigit(): return False
        return chu_so_4 != ma_ky_vong

    mask = pd.Series([bad(i) for i in range(len(df))], index=df.index) & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, mask)
    if len(out) == 0: return out

    out["Giới tính"] = gt[mask].values
    out["Số CMND"] = cmnd[mask].values
    out["Ngày sinh"] = ns[mask].dt.strftime("%d/%m/%Y").fillna("").values

    def mo_ta(i):
        g = gt.iloc[i]
        c = cmnd.iloc[i]
        d = ns.iloc[i]
        ten_gt = "Nam" if g == "01" else "Nữ" if g == "02" else "?"
        ma_dung = _ma_gioi_tinh_cccd(d, g)
        return f"CCCD số thứ 4 là '{c[3]}', nhưng {ten_gt} sinh năm {d.year if not pd.isna(d) else '?'} phải là '{ma_dung}'"

    out["Lỗi"] = [mo_ta(i) for i in out.index]
    return out

def loi_6_nam_sinh(df):
    q = df["so_cmnd"].apply(safe_str)
    ns = df["ngay_sinh"].apply(to_date)

    def bad(i):
        c = q.iloc[i]
        d = ns.iloc[i]
        if len(c) != 12 or pd.isna(d): return False
        try: yy = d.strftime("%Y")[-2:]
        except Exception: return False
        return c[4:6] != yy

    mask = pd.Series([bad(i) for i in range(len(df))], index=df.index) & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, mask)
    if len(out) == 0: return out
    out["Số CMND"] = q[mask].values
    out["Ngày sinh"] = ns[mask].dt.strftime("%d/%m/%Y").fillna("").values
    out["Lỗi"] = "Năm sinh trong CMND (vị trí 5-6) khác năm sinh khách hàng"
    return out

def loi_7_thoi_han(df):
    ad = df["ngay_vay"].apply(to_date)
    ae = df["ngay_dh_hop_dong"].apply(to_date)
    loai = df["thoi_han_vay"].apply(safe_str)

    def bad(i):
        d1, d2 = ad.iloc[i], ae.iloc[i]
        if pd.isna(d1) or pd.isna(d2): return False
        thang = months_between(d1, d2)
        if pd.isna(thang): return False
        pl = loai.iloc[i]
        if pl == "Ngắn hạn": return thang > 12
        if pl == "Trung hạn": return thang <= 12 or thang > 60
        if pl == "Dài hạn": return thang <= 60
        return False

    mask = pd.Series([bad(i) for i in range(len(df))], index=df.index) & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, mask)
    if len(out) == 0: return out
    out["Ngày vay"] = ad[mask].dt.strftime("%d/%m/%Y").fillna("").values
    out["Ngày ĐH HĐ"] = ae[mask].dt.strftime("%d/%m/%Y").fillna("").values
    out["Thời hạn vay"] = loai[mask].values
    out["Số tháng"] = [months_between(ad.iloc[i], ae.iloc[i]) for i in out.index]
    out["Lỗi"] = [f"Phân loại '{loai.iloc[i]}' không khớp số tháng ({months_between(ad.iloc[i], ae.iloc[i])})" for i in out.index]
    return out

def loi_8_chua_co_dien_thoai(df):
    v = df["dien_thoai"].apply(safe_str)
    ac = df["so_khe_uoc"].apply(safe_str)
    m = (v == "") & (ac != "") & (ac != "Số khế ước") & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, m)
    if len(out) == 0: return out
    out["Số điện thoại"] = "(trống)"
    out["Lỗi"] = "Chưa có số điện thoại"
    return out

def loi_9_sai_ky_han(df):
    ad = df["ngay_vay"].apply(to_date)
    ae = df["ngay_dh_hop_dong"].apply(to_date)
    ah = df["thoi_han_vay"].apply(safe_str)

    def bad(i):
        d1, d2 = ad.iloc[i], ae.iloc[i]
        if pd.isna(d1) or pd.isna(d2): return False
        thang = months_between(d1, d2)
        if pd.isna(thang): return False
        phan_loai_dung = _phan_loai_theo_thang(thang)
        if phan_loai_dung is None: return False
        return ah.iloc[i] != phan_loai_dung

    mask = pd.Series([bad(i) for i in range(len(df))], index=df.index) & df["tinh_trang"].apply(is_open)
    out = _base_cols(df, mask)
    if len(out) == 0: return out
    out["Ngày vay"] = ad[mask].dt.strftime("%d/%m/%Y").fillna("").values
    out["Ngày ĐH HĐ"] = ae[mask].dt.strftime("%d/%m/%Y").fillna("").values
    out["Số tháng"] = [months_between(ad.iloc[i], ae.iloc[i]) for i in out.index]
    out["Thời hạn vay"] = ah[mask].values

    def mo_ta(i):
        thang = months_between(ad.iloc[i], ae.iloc[i])
        dung = _phan_loai_theo_thang(thang)
        return f"Cột AH ghi '{ah.iloc[i]}', nhưng {thang} tháng phải là '{dung}'"

    out["Lỗi"] = [mo_ta(i) for i in out.index]
    return out

def loi_10_sai_ma_ndt(df):
    ma = df["ma_ndt"].apply(safe_str)
    nv = df["nguon_von"].apply(safe_str)

    def chuan_hoa_aw(x):
        x = x.strip()
        if x.endswith(".0"): x = x[:-2]
        return x
    nv_chuan = nv.apply(chuan_hoa_aw)

    BANG = {"INV0207250086489": "2", "INV0107190050369": "1", "": "1"}

    def bad(i):
        m = ma.iloc[i]
        a = nv_chuan.iloc[i]
        if m in BANG: return a != BANG[m]
        return True

    ac = df["so_khe_uoc"].apply(safe_str)
    mask = pd.Series([bad(i) for i in range(len(df))], index=df.index) & df["tinh_trang"].apply(is_open) & (ac != "") & (ac != "Số khế ước")
    out = _base_cols(df, mask)
    if len(out) == 0: return out

    out["Mã nhà đầu tư (EF)"] = ma[mask].replace("", "(trống)").values
    out["Nguồn vốn (AW)"] = nv_chuan[mask].values

    def mo_ta(i):
        m = ma.iloc[i]
        a = nv_chuan.iloc[i]
        if m in BANG:
            return f"Mã NĐT '{m if m else '(trống)'}' phải có nguồn vốn = {BANG[m]}, nhưng đang là {a}"
        return f"Mã NĐT '{m}' không có trong danh mục quy định"

    out["Lỗi"] = [mo_ta(i) for i in out.index]
    return out

def loi_11_to_khong_dat(df):
    ma_to  = df["ma_to"].apply(safe_str)
    ma_kh  = df["ma_kh"].apply(safe_str)
    ten_tt = df["ten_to_truong"].apply(safe_str)
    tt_open = df["tinh_trang"].apply(is_open)

    df_open = pd.DataFrame({"ma_to": ma_to, "ma_kh": ma_kh, "ten_tt": ten_tt, "open": tt_open})
    df_open = df_open[(df_open["open"]) & (df_open["ma_to"] != "") & (df_open["ma_kh"] != "")]

    if len(df_open) == 0:
        return pd.DataFrame(columns=["Mã tổ", "Tên tổ trưởng", "Số thành viên", "Lỗi"])

    dem = df_open.groupby("ma_to", sort=True).agg(
        ten_to_truong=("ten_tt", "first"),
        so_thanh_vien=("ma_kh", "nunique"),
    ).reset_index()

    dem = dem.rename(columns={"ma_to": "Mã tổ", "ten_to_truong": "Tên tổ trưởng", "so_thanh_vien": "Số thành viên"})
    ket_qua = dem[dem["Số thành viên"] < 20].copy()
    if len(ket_qua) == 0:
        return pd.DataFrame(columns=["Mã tổ", "Tên tổ trưởng", "Số thành viên", "Lỗi"])

    ket_qua["Lỗi"] = ket_qua["Số thành viên"].apply(lambda n: f"Tổ chỉ có {n} khách hàng (< 20)")
    return ket_qua.reset_index(drop=True)

def loi_12_het_du_no_can_ra_to(df):
    ma_to = df["ma_to"].apply(safe_str)
    du_no = df["du_no"].apply(to_num)
    ma_kh = df["ma_kh"].apply(safe_str)
    ten_kh = df["ten_kh"].apply(safe_str)
    ac = df["so_khe_uoc"].apply(safe_str)

    m = (ma_to != "") & (du_no == 0) & (ma_kh != "") & (ac != "") & (ac != "Số khế ước") & df["tinh_trang"].apply(is_open)
    out = df.loc[m, ["ma_kh", "ten_kh", "ma_to"]].copy()
    
    if len(out) == 0:
        return pd.DataFrame(columns=["Mã khách hàng", "Tên khách hàng", "Mã tổ", "Dư nợ (BI)", "Lỗi"])
        
    out = out.rename(columns={"ma_kh": "Mã khách hàng", "ten_kh": "Tên khách hàng", "ma_to": "Mã tổ"})
    out["Dư nợ (BI)"] = df.loc[m, "du_no"].apply(fmt_num).values
    out["Lỗi"] = "Dư nợ = 0 nhưng vẫn còn mã tổ (cần ra khỏi tổ)"
    return out.reset_index(drop=True)

CHECKS = [
    ("1. Lãi tồn trên 1 triệu",    loi_1_lai_ton),
    ("2. Sai ngày đến hạn",        loi_2_ngay_den_han),
    ("3. Sai nơi cấp",              loi_3_noi_cap),
    ("4. Sai số điện thoại",        loi_4_dien_thoai),
    ("5. Sai giới tính",            loi_5_gioi_tinh),
    ("6. Sai năm sinh",             loi_6_nam_sinh),
    ("7. Sai thời hạn (phân loại)", loi_7_thoi_han),
    ("8. Chưa có số điện thoại",    loi_8_chua_co_dien_thoai),
    ("9. Sai kỳ hạn vay",           loi_9_sai_ky_han),
    ("10. Sai mã nhà đầu tư",       loi_10_sai_ma_ndt),
    ("11. Tổ không đạt yêu cầu",    loi_11_to_khong_dat),
    ("12. Hết dư nợ cần ra tổ",     loi_12_het_du_no_can_ra_to),
]


class KiemTraDuLieuTool:
    def __init__(self, root):
        self.root = root
        if isinstance(root, tk.Tk):
            ui.apply_theme(root)
            root.title("Kiểm tra dữ liệu tín dụng")
            root.geometry("1200x800")

        self.results = {}
        self.buttons = {}
        self.current_df = None
        self._src_key = None
        
        self._build_ui()
        
        store.subscribe(self.on_data_changed)       # tự chạy lại khi nạp / xóa sao kê ở launcher
        self.on_data_changed()                      # chạy ngay nếu kho đã có dữ liệu

    def _style_button(self, btn, n):
        """n = None: chưa có dữ liệu; n > 0: có lỗi (đỏ); n == 0: đạt (xanh)."""
        if n is None:
            btn.config(bg=ui.SECOND, hover=ui.SECOND_H, fg=ui.MUTED)
        elif n > 0:
            btn.config(bg="#fee2e2", hover="#fecaca", fg="#b91c1c")
        else:
            btn.config(bg="#dcfce7", hover="#bbf7d0", fg="#15803d")

    def _reset_buttons(self):
        for title, btn in self.buttons.items():
            btn.config(text=f"{title} (0)")
            self._style_button(btn, None)

    def _build_ui(self):
        main = tk.Frame(self.root, bg=ui.BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)

        # --- TIÊU ĐỀ ---
        head = tk.Frame(main, bg=ui.BG)
        head.pack(fill="x", pady=(0, 10))
        left = tk.Frame(head, bg=ui.BG)
        left.pack(side="left")
        ui.label(left, "Kiểm tra dữ liệu tín dụng", bold=True, size=15).pack(anchor="w")
        self.lbl_file = ui.label(left, "Chưa có dữ liệu (hãy nạp sao kê ở khung \"Dữ liệu dùng chung\" bên trái)",
                                 muted=True)
        self.lbl_file.pack(anchor="w")
        self.lbl_status = tk.Label(head, text="Sẵn sàng", bg=ui.BG, fg=ui.MUTED, font=ui.font(11, True))
        self.lbl_status.pack(side="right")

        # --- NÚT CHỨC NĂNG ---
        tool = tk.Frame(main, bg=ui.BG)
        tool.pack(fill="x", pady=(0, 10))
        self.btn_summary = ui.RoundButton(tool, text="Tổng hợp tất cả", command=self.show_summary,
                                          padx=18, pady=8)
        self.btn_summary.pack(side="left")
        self.btn_export = ui.secondary_button(tool, "Xuất Excel", self.export_excel, padx=18, pady=8)
        self.btn_export.pack(side="left", padx=10)

        split = tk.Frame(main, bg=ui.BG)
        split.pack(fill="both", expand=True)

        # --- CỘT TRÁI: DANH MỤC LỖI (1 cột, đủ rộng, có cuộn nếu cửa sổ thấp) ---
        sidebar = ui.Card(split, title="Danh mục lỗi", fill=True)
        sidebar.canvas.configure(width=340)
        sidebar.pack(side="left", fill="y", padx=(0, 10))
        sf = ui.ScrollFrame(sidebar.content, bg=ui.CARD)
        sf.pack(fill="both", expand=True)
        for title, _ in CHECKS:
            btn = ui.RoundButton(sf.inner, text=f"{title} (0)", anchor="w", padx=14, pady=9,
                                 bold=False, parent_bg=ui.CARD,
                                 command=lambda t=title: self.show_error(t))
            btn.pack(fill="x", pady=3, padx=(0, 6))
            self.buttons[title] = btn
        self._reset_buttons()

        # --- CỘT PHẢI: BẢNG KẾT QUẢ (giãn theo chiều cao cửa sổ) ---
        table_card = ui.Card(split, fill=True)
        table_card.pack(side="left", fill="both", expand=True)

        self.lbl_detail = tk.Label(table_card.content, text=HINT_DETAIL, bg=ui.CARD, fg=ui.TEXT,
                                   font=ui.font(11, True), anchor="w", pady=4)
        self.lbl_detail.pack(fill="x")

        tree_wrap = tk.Frame(table_card.content, bg=ui.CARD)
        tree_wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_wrap, show="headings")
        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        self.tree.tag_configure("odd",  background="#f8fafc")
        self.tree.tag_configure("even", background="#ffffff")

        self.tree.bind("<Double-1>", self._on_double_click)

        # Tooltip
        self.tooltip = tk.Toplevel(self.root)
        self.tooltip.withdraw()
        self.tooltip.overrideredirect(True)
        self.tooltip_label = tk.Label(self.tooltip, bg="#ffffe0", fg="#1e293b",
                                      font=ui.font(10), justify="left", relief="solid", bd=1,
                                      padx=5, pady=5)
        self.tooltip_label.pack()

        self.tree.bind("<Motion>", self._on_motion)
        self.tree.bind("<Leave>", lambda e: self.tooltip.withdraw())

    def on_data_changed(self):
        """Được gọi khi sao kê dùng chung thay đổi (nạp mới / xóa) hoặc khi mở tool."""
        records = store.records
        if not records:
            self._src_key = None
            self.current_df = None
            self.results.clear()
            self.lbl_file.config(text="Chưa có dữ liệu (hãy nạp sao kê ở khung \"Dữ liệu dùng chung\" bên trái)")
            self.lbl_status.config(text="Trống", fg="#f59e0b")
            self.clear_tree()
            self._reset_buttons()
            return

        key = (id(records), len(records))
        if key == self._src_key and self.current_df is not None:
            return                                   # chỉ đổi hộ nghèo, sao kê vẫn như cũ

        if not store.has_positions:                  # dữ liệu lưu bằng bản cũ: thiếu vị trí cột
            self._src_key = None
            self.current_df = None
            self.clear_tree()
            self._reset_buttons()
            self.lbl_file.config(text="Sao kê đang lưu bằng phiên bản cũ nên chưa biết vị trí cột. "
                                      "Hãy bấm \"Nạp sao kê\" nạp lại file một lần.")
            self.lbl_status.config(text="Cần nạp lại sao kê", fg=ui.ERR)
            return

        self.lbl_status.config(text="Đang xử lý...", fg="#f59e0b")
        self.root.update_idletasks()
        try:
            self.current_df = build_dataframe(store.rows)
            text = "Dữ liệu dùng chung"
            if store.ngay:
                text += f"  (Ngày sao kê: {store.ngay})"
            self.lbl_file.config(text=text)
            self.run_checks()
            self._src_key = key
            self.lbl_status.config(text=f"Đã nạp {len(self.current_df)} dòng", fg=ui.OK)
        except Exception as e:
            self.lbl_status.config(text="Lỗi xử lý dữ liệu", fg=ui.ERR)
            messagebox.showerror("Lỗi", f"Không thể phân tích dữ liệu sao kê:\n{e}")

    def run_checks(self):
        if self.current_df is None or len(self.current_df) == 0:
            return
        self.results.clear()
        for title, fn in CHECKS:
            try:
                res = fn(self.current_df)
            except Exception as e:
                res = pd.DataFrame({"Lỗi": [f"Lỗi kiểm tra: {e}"]})
            self.results[title] = res
            btn = self.buttons[title]
            btn.config(text=f"{title} ({len(res)})")
            self._style_button(btn, len(res))

        self.clear_tree()
        self.lbl_detail.config(text=HINT_DETAIL, fg=ui.TEXT)

    def show_error(self, title):
        res = self.results.get(title)
        if res is None: return
        self.fill_tree(res)
        self.lbl_detail.config(text=f"{title}  —  {len(res)} lỗi", fg="#dc2626" if len(res) else "#16a34a")

    def show_summary(self):
        if not self.results: return
        rows = [(t, len(r)) for t, r in self.results.items()]
        summary = pd.DataFrame(rows, columns=["Loại lỗi", "Số lượng"])
        self.fill_tree(summary)
        total = sum(len(r) for r in self.results.values())
        self.lbl_detail.config(text=f"Tổng hợp  —  {total} lỗi", fg="#7c3aed")

    def export_excel(self):
        if not self.results:
            messagebox.showinfo("Thông báo", "Chưa có dữ liệu để xuất.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile="ket_qua_kiem_tra.xlsx")
        if not path: return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                for title, r in self.results.items():
                    r.to_excel(w, sheet_name=title[:31], index=False)
            messagebox.showinfo("Thành công", f"Đã xuất:\n{path}")
        except Exception as e:
            messagebox.showerror("Lỗi xuất file", str(e))

    def clear_tree(self):
        for c in self.tree["columns"]:
            self.tree.heading(c, text="")
        self.tree["columns"] = ()
        for i in self.tree.get_children():
            self.tree.delete(i)

    def fill_tree(self, df):
        self.clear_tree()
        cols = list(df.columns)
        self.tree["columns"] = cols
        
        for c in cols:
            label = LABELS.get(c, c)
            self.tree.heading(c, text=label)
            max_w = len(label) * 10 + 20
            if not df.empty:
                sample_vals = df[c].head(100).astype(str).tolist()
                for v in sample_vals:
                    if v and v != "nan":
                        w = len(v) * 8 + 20
                        if w > max_w: max_w = w
            if max_w > 450: max_w = 450
            elif max_w < 100: max_w = 100
                
            self.tree.column(c, width=max_w, anchor="w", stretch=False)

        for idx, (_, row) in enumerate(df.head(5000).iterrows()):
            vals = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            tag = "odd" if idx % 2 else "even"
            self.tree.insert("", "end", values=vals, tags=(tag,))

        if len(df) > 5000:
            self.lbl_detail.config(text=f"(Hiển thị 5000/{len(df)} dòng — xuất Excel để xem đầy đủ)")

    def _on_motion(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if row_id and col_id:
            col_idx = int(col_id.replace("#", "")) - 1
            vals = self.tree.item(row_id, "values")
            if col_idx < len(vals):
                text = str(vals[col_idx])
                if text and len(text) > 20:
                    self.tooltip_label.config(text=text)
                    self.tooltip.geometry(f"+{event.x_root + 15}+{event.y_root + 10}")
                    self.tooltip.deiconify()
                    return
        self.tooltip.withdraw()

    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item: return
        vals = self.tree.item(item, "values")
        cols = self.tree["columns"]
        self._show_row_popup(cols, vals)

    def _show_row_popup(self, cols, vals):
        win = tk.Toplevel(self.root)
        win.title("Chi tiết dòng")
        win.geometry("700x500")
        win.transient(self.root.winfo_toplevel())

        wrap = tk.Frame(win, bg="#ffffff", padx=16, pady=16)
        wrap.pack(fill="both", expand=True, padx=12, pady=12)

        canvas = tk.Canvas(wrap, bg="#ffffff", highlightthickness=0)
        scroll = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg="#ffffff")

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)

        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        for c, v in zip(cols, vals):
            label = LABELS.get(c, c)
            row = tk.Frame(inner, bg="#ffffff")
            row.pack(fill="x", pady=4)

            tk.Label(row, text=label + ":", bg="#ffffff", fg="#64748b", font=("Segoe UI", 10, "bold"), width=18, anchor="w").pack(side="left", anchor="n")
            tk.Label(row, text=v if v else "(trống)", bg="#ffffff", fg="#1e293b", font=("Segoe UI", 11), anchor="w", justify="left", wraplength=480).pack(side="left", fill="x", expand=True)

        ui.RoundButton(win, text="Đóng", command=win.destroy).pack(pady=10)

# ===================== KHAI BÁO TOOL =====================
TOOL_NAME = "Kiểm tra dữ liệu tín dụng"
TOOL_CLASS = KiemTraDuLieuTool
TOOL_ORDER = 40

if __name__ == "__main__":
    root = tk.Tk()
    KiemTraDuLieuTool(root)
    root.mainloop()