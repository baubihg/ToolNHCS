# launcher.py
import importlib
import logging
import os
import pkgutil
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import tools                      # thư mục tools/ (có __init__.py)
import updater                    # cập nhật bắt buộc qua GitHub Releases
from tools import _paths as paths # nơi lưu dữ liệu / cấu hình / nhật ký
from tools import _ui as ui       # giao diện dùng chung
from tools._shared import store   # dữ liệu sao kê dùng chung cho mọi tool
from version import APP_NAME, APP_VERSION


def discover_tools():
    """Quét tools/, trả về list (order, name, class) và list lỗi import."""
    found, errors = [], []
    for m in pkgutil.iter_modules(tools.__path__):
        if m.name.startswith("_"):
            continue
        try:
            mod = importlib.import_module(f"tools.{m.name}")
        except Exception:
            errors.append((m.name, traceback.format_exc()))
            continue
        name = getattr(mod, "TOOL_NAME", None)
        cls = getattr(mod, "TOOL_CLASS", None)
        if name and cls:
            found.append((getattr(mod, "TOOL_ORDER", 100), name, cls))
    found.sort(key=lambda t: (t[0], t[1]))
    return found, errors


class Launcher(tk.Tk):
    def report_callback_exception(self, exc, val, tb):
        """Lỗi bất ngờ trong giao diện: ghi nhật ký + báo gọn (bản exe không có cửa sổ lệnh để xem lỗi)."""
        logging.error("Lỗi chưa xử lý", exc_info=(exc, val, tb))
        messagebox.showerror("Có lỗi xảy ra", f"{exc.__name__}: {val}\n\nChi tiết đã ghi vào:\n{paths.log_path()}")

    def __init__(self):
        super().__init__()
        paths.setup_logging()
        ui.apply_theme(self)
        self.title(APP_NAME)
        try:                                           # icon cửa sổ + thanh tác vụ (áp dụng cho cả hộp thoại)
            self.iconbitmap(default=paths.resource_path(os.path.join("assets", "app.ico")))
        except tk.TclError:
            pass
        self.geometry("1180x760")
        self.minsize(960, 620)

        # Bản đóng gói (.exe): bắt buộc kiểm tra / cập nhật trước, qua được mới dựng giao diện.
        # Chạy bằng "python launcher.py": bỏ qua kiểm tra.
        if updater.is_frozen():
            updater.cleanup_leftovers()
            updater.UpdateGate(self, on_pass=self._build_ui)
        else:
            self._build_ui()

    def _build_ui(self):
        paths.ensure_output_dirs()          # lần đầu chạy: tự tạo "Thong bao phe duyet" và "Hop dong GQVL"

        # ---------- Sidebar trái ----------
        self.sidebar = tk.Frame(self, width=230, bg=ui.SIDEBAR)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        tk.Label(self.sidebar, text=APP_NAME, bg=ui.SIDEBAR, fg="white",
                 font=ui.font(13, True), anchor="w").pack(fill="x", padx=20, pady=(24, 2))
        tk.Label(self.sidebar, text="Chọn một công cụ bên dưới", bg=ui.SIDEBAR, fg="#8fa1b3",
                 font=ui.font(9), anchor="w").pack(fill="x", padx=20, pady=(0, 18))

        tk.Label(self.sidebar, text=f"Phiên bản {APP_VERSION}", bg=ui.SIDEBAR, fg="#6f8397",
                 font=ui.font(8), anchor="w").pack(side="bottom", fill="x", padx=20, pady=(0, 10))

        # ---------- Dữ liệu dùng chung (cuối sidebar) ----------
        panel = tk.Frame(self.sidebar, bg=ui.SIDEBAR)
        panel.pack(side="bottom", fill="x", padx=14, pady=(16, 4))
        tk.Label(panel, text="DỮ LIỆU DÙNG CHUNG", bg=ui.SIDEBAR, fg="#8fa1b3",
                 font=ui.font(8, True), anchor="w").pack(fill="x")
        self.sk_label = self._data_block(panel, self.load_sao_ke, "Nạp sao kê", self.clear_sao_ke)
        self.hn_label = self._data_block(panel, self.load_ho_ngheo, "Nạp hộ nghèo", self.clear_ho_ngheo)
        store.subscribe(self._refresh_shared)
        self._refresh_shared()

        # ---------- Vùng nội dung phải ----------
        self.content = tk.Frame(self, bg=ui.BG)
        self.content.pack(side="right", fill="both", expand=True)

        # ---------- Registry ----------
        self._classes, self._instances, self._buttons = {}, {}, {}
        self._current = None

        found, errors = discover_tools()
        for _, name, cls in found:
            self.register(name, cls)
        for mod_name, tb in errors:        # tool lỗi import: không làm sập launcher
            self.register_error(mod_name, tb)

        if self._classes:
            self.show(next(iter(self._classes)))

    # ---------- dữ liệu dùng chung ----------
    def _data_block(self, parent, load_cmd, load_text, clear_cmd):
        lbl = tk.Label(parent, text="", bg=ui.SIDEBAR, fg="white", font=ui.font(9),
                       anchor="w", justify="left", wraplength=190)
        lbl.pack(fill="x", pady=(10, 5))
        row = tk.Frame(parent, bg=ui.SIDEBAR)
        row.pack(fill="x")
        ui.RoundButton(row, text=load_text, command=load_cmd, padx=14, pady=7, size=9,
                       parent_bg=ui.SIDEBAR).pack(side="left")
        ui.RoundButton(row, text="Xóa", command=clear_cmd, bg=ui.SIDEBAR_BTN, hover=ui.DANGER,
                       padx=12, pady=7, size=9, parent_bg=ui.SIDEBAR).pack(side="left", padx=6)
        return lbl

    def _refresh_shared(self):
        self.sk_label.config(text="Sao kê: " + store.summary(),
                             fg="#7ee0b8" if store.records else "#f5b7b1")
        self.hn_label.config(text="Hộ nghèo: " + store.ho_ngheo_summary(),
                             fg="#7ee0b8" if store.ho_ngheo else "#f5b7b1")

    def _load(self, title, loader, label):
        path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        label.config(text="Đang xử lý file...", fg="white")
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            loader(path)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi đọc file:\n{e}")
            self._refresh_shared()
        finally:
            self.config(cursor="")

    def load_sao_ke(self):
        self._load("Chọn file Sao kê", store.load_file, self.sk_label)

    def load_ho_ngheo(self):
        self._load("Chọn file Danh sách hộ nghèo", store.load_ho_ngheo_file, self.hn_label)

    def clear_sao_ke(self):
        if store.records and messagebox.askyesno(
                "Xác nhận", "Xóa dữ liệu sao kê? Tất cả công cụ sẽ không còn dữ liệu này."):
            store.clear()

    def clear_ho_ngheo(self):
        if store.ho_ngheo and messagebox.askyesno(
                "Xác nhận", "Xóa danh sách hộ nghèo? Tất cả công cụ sẽ không còn dữ liệu này."):
            store.clear_ho_ngheo()

    # ---------- đăng ký ----------
    def _add_button(self, text, command, fg="white"):
        btn = ui.RoundButton(self.sidebar, text=text, command=command, bg=ui.SIDEBAR,
                             hover=ui.SIDEBAR_BTN, fg=fg, anchor="w", radius=10,
                             padx=16, pady=11, size=10, bold=False, parent_bg=ui.SIDEBAR)
        btn.pack(fill="x", padx=12, pady=2)
        return btn

    def register(self, name, app_class):
        """Đăng ký một tool; số thứ tự tự tăng theo thứ tự (TOOL_ORDER) nên tool thêm sau tự có số tiếp theo."""
        self._classes[name] = app_class
        number = len(self._classes)
        self._buttons[name] = self._add_button(f"{number}. {name}", lambda n=name: self.show(n))

    def register_error(self, mod_name, tb):
        self._add_button(f"⚠ {mod_name}", lambda: messagebox.showerror(
            f"Không nạp được tool: {mod_name}", tb), fg="#f5b7b1")

    # ---------- hiển thị ----------
    def show(self, name):
        if self._current:
            self._current.pack_forget()

        if name not in self._instances:
            wrapper = tk.Frame(self.content, bg=ui.BG)
            try:
                self._classes[name](wrapper)
            except Exception:
                wrapper.destroy()
                messagebox.showerror(f"Lỗi khi mở: {name}", traceback.format_exc())
                return
            self._instances[name] = wrapper

        self._current = self._instances[name]
        self._current.pack(fill="both", expand=True)
        self.update_idletasks()

        for n, b in self._buttons.items():
            if n == name:
                b.config(bg=ui.PRIMARY, hover=ui.PRIMARY_H)
            else:
                b.config(bg=ui.SIDEBAR, hover=ui.SIDEBAR_BTN)
        self.title(f"{APP_NAME} - {name}")


if __name__ == "__main__":
    Launcher().mainloop()
