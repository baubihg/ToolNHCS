# -*- coding: utf-8 -*-
"""
Giao diện dùng chung cho launcher và các tool (bảng màu, font, nút bo góc, khung bo góc).
File bắt đầu bằng dấu "_" nên launcher không coi đây là một tool.
"""
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk

# ---------- Bảng màu ----------
BG = "#eef1f5"          # nền chung
CARD = "#ffffff"        # nền khung
BORDER = "#d8dee6"
TEXT = "#1f2933"
MUTED = "#6b7785"
PRIMARY = "#16a085"
PRIMARY_H = "#13896f"
SECOND = "#e3e8ee"      # nút phụ
SECOND_H = "#d3dae3"
DANGER = "#e74c3c"
DANGER_H = "#cf3e2f"
OK = "#1e8449"
ERR = "#c0392b"
INFO = "#2563eb"
SIDEBAR = "#22313f"
SIDEBAR_BTN = "#2c3e50"
DISABLED = "#b8c0c8"

FAMILY = "Segoe UI"     # được chọn lại trong apply_theme theo font máy có


def font(size=10, bold=False):
    return (FAMILY, size, "bold" if bold else "normal")


def _parent_bg(widget):
    try:
        return widget.cget("bg")
    except tk.TclError:
        return BG


def apply_theme(root):
    """Gọi 1 lần ở cửa sổ gốc: đặt font mặc định + kiểu ttk."""
    global FAMILY
    families = set(tkfont.families(root))
    for f in ("Segoe UI", "Roboto", "Helvetica Neue", "Arial"):
        if f in families:
            FAMILY = f
            break

    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(name).configure(family=FAMILY, size=10)
        except tk.TclError:
            pass
    root.configure(bg=BG)
    root.option_add("*TCombobox*Listbox.font", (FAMILY, 10))
    root.option_add("*TCombobox*Listbox.selectBackground", PRIMARY)

    st = ttk.Style(root)
    st.theme_use("clam")
    st.configure(".", font=(FAMILY, 10), background=BG, foreground=TEXT)
    st.configure("TFrame", background=BG)
    st.configure("TLabel", background=BG, foreground=TEXT)

    field = dict(fieldbackground="white", bordercolor=BORDER, lightcolor=BORDER,
                 darkcolor=BORDER, padding=(8, 6), relief="flat")
    st.configure("TEntry", **field)
    st.configure("TCombobox", arrowcolor=MUTED, **field)
    for w in ("TEntry", "TCombobox"):
        st.map(w, bordercolor=[("focus", PRIMARY)], lightcolor=[("focus", PRIMARY)],
               darkcolor=[("focus", PRIMARY)],
               fieldbackground=[("readonly", "#f6f8fa")])
    st.map("TCombobox", fieldbackground=[("readonly", "white")],
           selectbackground=[("readonly", "white")], selectforeground=[("readonly", TEXT)])

    st.configure("Treeview", rowheight=30, fieldbackground="white", background="white",
                 borderwidth=0, font=(FAMILY, 10))
    st.configure("Treeview.Heading", background=PRIMARY, foreground="white", relief="flat",
                 font=(FAMILY, 10, "bold"), padding=(8, 7))
    st.map("Treeview.Heading", background=[("active", PRIMARY_H)])
    st.map("Treeview", background=[("selected", "#cdeee6")], foreground=[("selected", TEXT)])

    st.configure("Horizontal.TProgressbar", troughcolor=BORDER, background=PRIMARY,
                 bordercolor=BORDER, lightcolor=PRIMARY, darkcolor=PRIMARY, thickness=10)
    st.configure("TScrollbar", troughcolor=BG, background="#c3ccd6", bordercolor=BG,
                 arrowcolor=MUTED, relief="flat")
    st.map("TScrollbar", background=[("active", "#a9b4c0")])


# ---------- Hình bo góc ----------
def round_rect(canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


class RoundButton(tk.Canvas):
    """Nút bo góc. Hỗ trợ config(text=..., state='normal'/'disabled', command=..., bg=..., hover=...)."""

    def __init__(self, parent, text="", command=None, bg=PRIMARY, hover=PRIMARY_H, fg="white",
                 radius=10, padx=18, pady=8, anchor="center", width=None, size=10, bold=True,
                 parent_bg=None):
        self._text, self._command = text, command
        self._bg, self._hover, self._fg = bg, hover, fg
        self._r, self._anchor, self._padx = radius, anchor, padx
        self._state = "normal"
        self._cur = bg
        self._font = tkfont.Font(family=FAMILY, size=size, weight="bold" if bold else "normal")
        w = width or self._font.measure(text) + 2 * padx
        h = self._font.metrics("linespace") + 2 * pady
        super().__init__(parent, width=w, height=h, highlightthickness=0, bd=0,
                         bg=parent_bg or _parent_bg(parent), cursor="hand2")
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", lambda e: self._paint(self._hover))
        self.bind("<Leave>", lambda e: self._paint(self._bg))
        self.bind("<ButtonPress-1>", lambda e: self._paint(self._hover))
        self.bind("<ButtonRelease-1>", self._release)
        self._draw()

    def _paint(self, color):
        if self._state == "normal":
            self._cur = color
            self._draw()

    def _draw(self):
        self.delete("all")
        w = max(self.winfo_width(), int(self.cget("width")))
        h = max(self.winfo_height(), int(self.cget("height")))
        color = self._cur if self._state == "normal" else DISABLED
        round_rect(self, 1, 1, w - 1, h - 1, self._r, fill=color, outline=color)
        if self._anchor == "w":
            self.create_text(self._padx, h // 2, text=self._text, fill=self._fg,
                             font=self._font, anchor="w")
        else:
            self.create_text(w // 2, h // 2, text=self._text, fill=self._fg, font=self._font)

    def _release(self, e):
        if self._state == "normal" and 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height():
            if self._command:
                self._command()

    def configure(self, cnf=None, **kw):
        if cnf:
            kw.update(cnf)
        if not kw:
            return super().configure()
        redraw = False
        for key, attr in (("text", "_text"), ("command", "_command"), ("hover", "_hover"),
                          ("fg", "_fg")):
            if key in kw:
                setattr(self, attr, kw.pop(key))
                redraw = True
        if "bg" in kw:
            self._bg = self._cur = kw.pop("bg")
            redraw = True
        if "state" in kw:
            self._state = kw.pop("state")
            self.config(cursor="hand2" if self._state == "normal" else "arrow")
            if self._state == "normal":
                self._cur = self._bg
            redraw = True
        if kw:
            super().configure(**kw)
        if redraw:
            self._draw()

    config = configure


class Card(tk.Frame):
    """Khung trắng bo góc. Đặt widget con vào card.content.
    fill=True: phần nội dung giãn theo chiều cao khung (dùng cho bảng/danh sách)."""

    def __init__(self, parent, title=None, pad=14, radius=12, fill=False, parent_bg=None):
        bg = parent_bg or _parent_bg(parent)
        super().__init__(parent, bg=bg)
        self._pad, self._r, self._fill = pad, radius, fill
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0, height=40)
        self.canvas.pack(fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=CARD)
        self._win = self.canvas.create_window(pad, pad, window=self.body, anchor="nw")
        if title:
            tk.Label(self.body, text=title, bg=CARD, fg=TEXT,
                     font=font(11, True)).pack(anchor="w", pady=(0, 8))
        self.content = tk.Frame(self.body, bg=CARD)
        self.content.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", self._on_canvas)
        self.body.bind("<Configure>", self._on_body)

    def _on_canvas(self, e):
        c = self.canvas
        c.delete("bg")
        round_rect(c, 1, 1, max(e.width - 1, 2), max(e.height - 1, 2), self._r,
                   fill=CARD, outline=BORDER, tags="bg")
        c.tag_lower("bg")
        opts = {"width": max(e.width - 2 * self._pad, 1)}
        if self._fill:
            opts["height"] = max(e.height - 2 * self._pad, 1)
        c.itemconfigure(self._win, **opts)

    def _on_body(self, e):
        if not self._fill:
            self.canvas.configure(height=self.body.winfo_reqheight() + 2 * self._pad)


class ScrollFrame(tk.Frame):
    """Vùng cuộn dọc. Đặt widget con vào scrollframe.inner."""

    def __init__(self, parent, bg=BG):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0, width=1, height=1)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window(0, 0, window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._on_wheel, add="+")

    def _on_wheel(self, e):
        try:
            w = self.winfo_containing(e.x_root, e.y_root)
        except (KeyError, tk.TclError):
            return
        while w is not None and w is not self:
            w = w.master
        if w is None:
            return                      # con trỏ không nằm trong vùng cuộn này
        if self.inner.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-e.delta / 120), "units")


# ---------- Hàm tạo widget nhanh ----------
def label(parent, text="", bold=False, muted=False, size=10, **kw):
    bg = kw.pop("bg", None) or _parent_bg(parent)
    fg = kw.pop("fg", None) or (MUTED if muted else TEXT)
    return tk.Label(parent, text=text, bg=bg, fg=fg, font=font(size, bold), **kw)


def secondary_button(parent, text, command, **kw):
    return RoundButton(parent, text=text, command=command, bg=SECOND, hover=SECOND_H,
                       fg=TEXT, **kw)
