# -*- coding: utf-8 -*-
"""
Tool Tra cứu thông tin tín dụng.
- Dùng dữ liệu SAO KÊ và DS HỘ NGHÈO dùng chung (nạp ở khung "Dữ liệu dùng chung" của launcher)
- Tìm theo xã, tên KH, CCCD, mã KH, tên vợ/chồng, SĐT, mã món vay (kết hợp tất cả ô đã nhập)
"""
import math
import re
import unicodedata

import tkinter as tk
from tkinter import ttk, messagebox

try:
    from . import _ui as ui
    from ._shared import store, remove_accents
except ImportError:          # chạy trực tiếp file này, không qua launcher
    import _ui as ui
    from _shared import store, remove_accents

ALL_XA = "Tất Cả"
HINT = "Vui lòng chọn một dòng ở danh sách bên trái để xem chi tiết."


# ================== TIỆN ÍCH ==================
def match_number(cell, search):
    """So khớp số, bỏ qua số 0 đầu (CCCD/SĐT bị Excel làm mất số 0)."""
    if not search:
        return True
    c = str(cell if cell is not None else "").strip()
    s = str(search).strip()
    if c == s:
        return True
    if c.lstrip("0") == s.lstrip("0"):
        return True
    return s in c


def match_text(cell, norm_search):
    if not norm_search:
        return True
    return norm_search in remove_accents(cell)


def resolve_key(headers, names):
    """Tìm tên cột thật trong file ứng với danh sách tên có thể có."""
    for n in names:
        if n in headers:
            return n
    norm_names = [remove_accents(n) for n in names]
    for h in headers:
        ch = remove_accents(h)
        for nn in norm_names:
            if ch == nn or nn in ch:
                return h
    return None


def fmt_money(s):
    if not s:
        return ""
    try:
        v = float(s)
    except ValueError:
        return ""
    if not math.isfinite(v):
        return ""
    if v.is_integer():
        txt = f"{int(v):,}".replace(",", ".")
    else:
        txt = f"{v:,.3f}".rstrip("0").rstrip(".")
        txt = txt.replace(",", "§").replace(".", ",").replace("§", ".")
    return txt + " đ"


# ================== GIAO DIỆN ==================
class TraCuuApp:
    def __init__(self, root):
        self.root = root
        if isinstance(root, tk.Tk):                 # chỉ đặt title/size khi chạy độc lập
            ui.apply_theme(root)
            root.title("Tra cứu thông tin tín dụng")
            root.geometry("1100x720")

        self._shown = []

        self.main = main = tk.Frame(root, bg=ui.BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)
        self._build_header(main)
        self._build_search(main)
        self._build_results(main)

        store.subscribe(self._on_store_changed)      # dữ liệu dùng chung: nạp ở launcher là cập nhật ở đây
        self._on_store_changed()

    # Sao kê nằm trong kho dùng chung (tools/_shared.py)
    @property
    def excel_data(self):
        return store.records

    @property
    def ngay_sao_ke(self):
        return store.ngay

    # ---------- dựng giao diện ----------
    def _build_header(self, parent):
        head = tk.Frame(parent, bg=ui.BG)
        head.pack(fill="x", pady=(0, 8))
        left = tk.Frame(head, bg=ui.BG)
        left.pack(side="left")
        ui.label(left, "Tra cứu thông tin tín dụng", bold=True, size=15).pack(anchor="w")
        self.info_lbl = ui.label(left, "", muted=True)
        self.info_lbl.pack(anchor="w")
        self.ngay_lbl = tk.Label(head, text="", bg="#ff9800", fg="white", font=ui.font(10, True),
                                 padx=12, pady=3)

    def _build_search(self, parent):
        card = ui.Card(parent, title="Điều kiện truy vấn")
        card.pack(fill="x", pady=(10, 0))
        box = card.content

        top = tk.Frame(box, bg=ui.CARD)
        top.pack(fill="x", pady=(0, 6))
        ui.label(top, "Đơn vị hành chính:", bold=True).pack(side="left")
        self.xa_var = tk.StringVar(master=parent, value=ALL_XA)
        self.xa_cb = ttk.Combobox(top, textvariable=self.xa_var, values=[ALL_XA],
                                  state="readonly", width=32)
        self.xa_cb.pack(side="left", padx=10)
        self.xa_cb.bind("<<ComboboxSelected>>", lambda e: self.perform_search())
        ui.RoundButton(top, text="Tìm kiếm", command=self.perform_search, padx=18,
                       pady=6).pack(side="right")
        ui.secondary_button(top, "Xóa điều kiện", self.clear_filters, padx=14,
                            pady=6).pack(side="right", padx=8)

        grid = tk.Frame(box, bg=ui.CARD)
        grid.pack(fill="x")
        self.type_var = tk.StringVar(master=parent, value="ten_kh")
        fields = [("ten_kh", "Tên khách hàng:"), ("cccd", "CC khách hàng:"), ("ma_kh", "Mã khách hàng:"),
                  ("ten_vc", "Tên vợ (chồng):"), ("sdt", "Số điện thoại:"), ("ma_mv", "Mã món vay:")]
        self.entries = {}
        for i, (key, label) in enumerate(fields):
            r, c = divmod(i, 3)
            cell = tk.Frame(grid, bg=ui.CARD)
            cell.grid(row=r, column=c, sticky="ew", padx=6, pady=4)
            grid.columnconfigure(c, weight=1)
            tk.Radiobutton(cell, variable=self.type_var, value=key, bg=ui.CARD,
                           activebackground=ui.CARD, highlightthickness=0).pack(side="left")
            ui.label(cell, label, width=14, anchor="w").pack(side="left")
            ent = ttk.Entry(cell)
            ent.pack(side="left", fill="x", expand=True)
            ent.bind("<FocusIn>", lambda e, k=key: self.type_var.set(k))
            ent.bind("<Button-1>", lambda e, k=key: self.type_var.set(k))
            ent.bind("<Return>", lambda e: self.perform_search())
            self.entries[key] = ent

    def _build_results(self, parent):
        area = tk.Frame(parent, bg=ui.BG)
        area.pack(fill="both", expand=True, pady=(10, 0))
        area.rowconfigure(0, weight=1)
        area.columnconfigure(0, weight=3)
        area.columnconfigure(1, weight=2)

        # --- trái: danh sách ---
        left = ui.Card(area, title="Danh sách kết quả", fill=True)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        cols = ("xa", "dia_chi", "to", "ten")
        holder = tk.Frame(left.content, bg=ui.CARD)
        holder.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(holder, columns=cols, show="headings", selectmode="browse")
        for col, text, w in (("xa", "Tên xã", 110), ("dia_chi", "Địa chỉ", 180),
                             ("to", "Tên tổ trưởng", 120), ("ten", "Tên khách hàng", 150)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=w, anchor="w")
        vsb = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(holder, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # --- phải: chi tiết ---
        right = ui.Card(area, title="Thông tin chi tiết", fill=True)
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        holder2 = tk.Frame(right.content, bg=ui.CARD)
        holder2.pack(fill="both", expand=True)
        self.detail = tk.Text(holder2, wrap="word", font=ui.font(10), bg=ui.CARD, fg=ui.TEXT,
                              relief="flat", highlightthickness=0, padx=4, pady=4, cursor="arrow")
        dvsb = ttk.Scrollbar(holder2, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dvsb.set)
        dvsb.pack(side="right", fill="y")
        self.detail.pack(side="left", fill="both", expand=True)
        self.detail.tag_configure("b", font=ui.font(10, True))
        self.detail.tag_configure("blue", foreground=ui.INFO, font=ui.font(10, True))
        self.detail.tag_configure("hl", foreground="#d32f2f", font=ui.font(10, True), spacing1=10)
        self.detail.tag_configure("hint", foreground=ui.MUTED, justify="center", spacing1=40)
        self.detail.tag_configure("err", foreground=ui.ERR, font=ui.font(10, True))
        self.detail.tag_configure("hr", foreground="#bbbbbb")
        self.show_hint()

    def clear_filters(self):
        for ent in self.entries.values():
            ent.delete(0, "end")
        self.xa_var.set(ALL_XA)
        if self.excel_data:
            self.render_table(self.excel_data)

    # ---------- dữ liệu dùng chung ----------
    def _on_store_changed(self):
        """Được gọi khi sao kê / hộ nghèo dùng chung thay đổi (nạp mới / xóa) ở bất kỳ nơi nào."""
        self.info_lbl.config(
            text=f"Sao kê: {store.summary()}   |   Hộ nghèo: {store.ho_ngheo_summary()}"
                 "   (nạp / xóa dữ liệu ở khung \"Dữ liệu dùng chung\" bên trái)")
        if store.records:
            self.ngay_lbl.config(text=f"Ngày sao kê: {store.ngay or 'Không xác định'}")
            self.ngay_lbl.pack(side="right")
        else:
            self.ngay_lbl.pack_forget()
        self.populate_xa()
        if store.records:
            self.render_table(store.records)
        else:
            self.tree.delete(*self.tree.get_children())
            self._shown = []
            self.show_hint()

    # ---------- tìm kiếm ----------
    def populate_xa(self):
        xs = {str(r.get("Tên xã", "")).strip() for r in self.excel_data}
        xs.discard("")
        key = lambda s: unicodedata.normalize("NFKD", s.replace("đ", "d").replace("Đ", "D")).lower()
        self.xa_cb.config(values=[ALL_XA] + sorted(xs, key=key))
        self.xa_var.set(ALL_XA)

    def perform_search(self):
        if not self.excel_data:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng nạp file sao kê trước khi tìm kiếm!")
            return

        xa = self.xa_var.get()
        xa = "" if xa == ALL_XA else xa
        g = lambda k: self.entries[k].get()
        crit = {
            "xa": remove_accents(xa), "ten_kh": remove_accents(g("ten_kh")),
            "cccd": g("cccd").strip(), "ma_kh": remove_accents(g("ma_kh")),
            "ten_vc": remove_accents(g("ten_vc")), "sdt": g("sdt").strip(),
            "ma_mv": remove_accents(g("ma_mv")),
        }
        headers = list(self.excel_data[0].keys())
        names = {
            "xa": ["Tên xã"], "ten_kh": ["Tên KH", "Tên khách hàng", "Họ tên"],
            "cccd": ["Số CMND", "Số CCCD", "CMND", "CCCD"], "ma_kh": ["Mã KH", "Mã khách hàng"],
            "ten_vc": ["Tên vợ/chồng", "Tên vợ (chồng)", "Vợ/Chồng"],
            "sdt": ["Số điện thoại", "Điện thoại", "SĐT"], "ma_mv": ["Số khế ước", "Mã món vay"],
        }
        is_number = {"cccd", "sdt"}
        # Chỉ xét các tiêu chí có nhập; tìm sẵn tên cột thật một lần cho nhanh
        active = [(k, resolve_key(headers, names[k]), v) for k, v in crit.items() if v]

        result = []
        for row in self.excel_data:
            ok = True
            for k, col, v in active:
                cell = row.get(col, "") if col else ""
                if not (match_number(cell, v) if k in is_number else match_text(cell, v)):
                    ok = False
                    break
            if ok:
                result.append(row)
        self.render_table(result)

    # ---------- danh sách ----------
    def render_table(self, data):
        self.tree.delete(*self.tree.get_children())
        self._shown = data
        self.show_hint()
        if not data:
            self.tree.insert("", "end", iid="none",
                             values=("", "Không tìm thấy dữ liệu phù hợp", "", ""))
            return
        for i, r in enumerate(data):
            self.tree.insert("", "end", iid=str(i), values=(
                r.get("Tên xã") or r.get("Xã") or "",
                r.get("Địa chỉ") or "",
                r.get("Tên tổ") or r.get("Tổ trưởng") or "",
                r.get("Tên KH") or r.get("Tên khách hàng") or "",
            ))

    def on_select(self, _event=None):
        sel = self.tree.selection()
        if sel and sel[0].isdigit():
            self.show_detail(self._shown[int(sel[0])])

    # ---------- chi tiết ----------
    def show_hint(self):
        self.detail.config(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("end", HINT, "hint")
        self.detail.config(state="disabled")

    def _add(self, *parts):
        for p in parts:
            txt, tag = p if isinstance(p, tuple) else (p, ())
            self.detail.insert("end", txt, tag)
        self.detail.insert("end", "\n")

    def show_detail(self, row):
        def get(*keys):
            for k in keys:
                v = row.get(k)
                if v is not None and str(v).strip() != "":
                    return str(v).strip()
            return ""

        self.detail.config(state="normal")
        self.detail.delete("1.0", "end")
        n0 = 0

        ma_mv, ma_kh = get("Số khế ước", "Mã món vay"), get("Mã KH", "Mã khách hàng")
        if ma_mv or ma_kh:
            self._add("Mã món vay: ", (ma_mv, "b"), f" (Mã KH: {ma_kh})" if ma_kh else ""); n0 += 1

        chuong_trinh = get("Tên chương trình")
        nguon_raw = get("Nguồn vốn")
        ma_ndt = get("Mã nhà đầu tư", "Mã NĐT", "Mã nhà ĐT")
        if "giaiquyetvieclam" in remove_accents(chuong_trinh):
            if ma_ndt == "INV0107190050369":
                nguon = "Trung ương Quỹ quốc gia"
            elif ma_ndt == "INV0207250086489":
                nguon = "Địa phương"
            elif nguon_raw == "1":
                nguon = "Trung ương huy động"
            elif nguon_raw == "2":
                nguon = "Địa phương"
            else:
                nguon = nguon_raw
        else:
            nguon = {"1": "Trung ương", "2": "Địa phương"}.get(nguon_raw, nguon_raw)
        if chuong_trinh:
            self._add("Chương trình: ", (chuong_trinh, "b"), f" (Nguồn vốn: {nguon})" if nguon else "")
            n0 += 1

        muc_dich = get("Tên PNKT51", "Mục đích")
        if muc_dich:
            self._add(f"Mục đích: {muc_dich}"); n0 += 1

        du_no = fmt_money(get("Tổng dư nợ"))
        if du_no:
            self._add("Dư nợ: ", (du_no, "b"), " (TH)"); n0 += 1
        giai_ngan = fmt_money(get("Tổng giải ngân"))
        if giai_ngan:
            self._add("Giải ngân: ", (giai_ngan, "b")); n0 += 1
        du_co = fmt_money(get("Dư nợ trong hạn"))
        if du_co:
            self._add(f"Dư cơ: {du_co}"); n0 += 1

        lai_thu, lai_ton = fmt_money(get("Tổng thu lãi TH")), fmt_money(get("Lãi tồn TH"))
        if lai_thu or lai_ton:
            self._add(f"Lãi đã thu: {lai_thu}" + (f"; Lãi tồn: {lai_ton}" if lai_ton else "")); n0 += 1

        ngay_vay, han_vay = get("Ngày vay"), get("Ngày ĐH theo hợp đồng")
        if ngay_vay or han_vay:
            self._add(f"Ngày vay: {ngay_vay}" + (f"; Hạn vay: {han_vay}" if han_vay else "")); n0 += 1

        ngay_gd = get("Ngày giao dịch gần nhất")
        if ngay_gd:
            self._add(f"Ngày giao dịch gần nhất: {ngay_gd}"); n0 += 1

        tk105 = fmt_money(get("Số dư tiền gửi 105", "Số dư tiết kiệm 105"))
        if tk105:
            self._add(("Số dư tiền gửi 105: " + tk105, ui.INFO)); n0 += 1

        self._add(("─" * 38, "hr"))

        n1 = 0
        ho_ten, sdt = get("Tên KH", "Tên khách hàng"), get("Số điện thoại", "SĐT")
        if ho_ten or sdt:
            self._add("Họ và tên: ", (ho_ten, "b"), f" (SĐT: {sdt})" if sdt else ""); n1 += 1

        sinh, dan_toc, gt_raw = get("Ngày sinh"), get("Tên DT", "Dân tộc"), get("Giới tính")
        gt = {"01": "Nam", "1": "Nam", "02": "Nữ", "2": "Nữ"}.get(gt_raw, gt_raw)
        parts = []
        if sinh:
            parts.append(f"Sinh ngày: {sinh}")
        if gt:
            parts.append(f"Giới tính: {gt}")
        if dan_toc:
            parts.append(f"Dân tộc: {dan_toc}")
        if parts:
            self._add("; ".join(parts)); n1 += 1

        cccd = get("Số CMND", "CCCD")
        ngay_cap, noi_cap = get("Ngày cấp CMND"), get("Nơi cấp CMND")
        if cccd:
            self._add(f"Căn cước: {cccd}" + (f" (Cấp: {ngay_cap} - {noi_cap})" if ngay_cap or noi_cap else ""))
            n1 += 1

        thua_ke = get("Tên vợ/chồng", "Tên vợ (chồng)")
        if thua_ke:
            self._add(f"Tên người thừa kế: {thua_ke}"); n1 += 1

        dia_chi = get("Địa chỉ")
        if dia_chi:
            self._add(f"Địa chỉ: {dia_chi}"); n1 += 1

        # Liên kết dữ liệu hộ nghèo (so khớp CCCD, bỏ số 0 đầu)
        if cccd:
            phan_loai = store.phan_loai_ho(cccd)
            if phan_loai:
                self._add(("Điều tra biểu dân số mới nhất: thuộc hộ " + phan_loai.lower(), "hl"))
                n1 += 1

        if n0 + n1 == 0:
            self.detail.delete("1.0", "end")
            self.detail.insert("end", "Không có thông tin chi tiết cho bản ghi này.", "err")

        self.detail.config(state="disabled")


# ===== Khai báo cho launcher =====
TOOL_NAME = "Tra cứu thông tin"
TOOL_CLASS = TraCuuApp
TOOL_ORDER = 20


if __name__ == "__main__":
    root = tk.Tk()
    TraCuuApp(root)
    root.mainloop()
