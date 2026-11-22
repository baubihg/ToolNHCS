# Tool_BaoCao.py
import tkinter as tk
from tkinter import ttk

class BaoCaoApp:
    def __init__(self, root):
        self.root = root
        if isinstance(root, (tk.Tk, tk.Toplevel)):
            root.title("Báo cáo")
            root.geometry("800x600")

        # ---- UI của app (pack vào root) ----
        tk.Label(root, text="Báo cáo tháng", font=("Arial", 14, "bold")).pack(pady=20)
        ttk.Button(root, text="Xuất Excel").pack()
        # ... phần logic của app