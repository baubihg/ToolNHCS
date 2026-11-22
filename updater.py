# -*- coding: utf-8 -*-
"""
Cập nhật BẮT BUỘC cho launcher (bản đóng gói .exe) qua GitHub Releases.

Cách hoạt động
  1. Mỗi lần mở: hỏi GitHub bản Release mới nhất (không dùng API nên không bị giới hạn lượt gọi).
  2. Có bản mới hơn APP_VERSION -> màn hình chặn: chỉ có "Cập nhật ngay" hoặc "Thoát".
     Không cập nhật = chương trình đóng. Cập nhật = tải về, thay file exe, tự mở lại bản mới.
  3. Không có Internet: cho dùng tạm tối đa OFFLINE_GRACE_DAYS ngày kể từ lần kiểm tra thành công gần nhất,
     quá hạn thì cũng bị chặn.
  4. Chạy bằng "python launcher.py" (không phải exe) thì bỏ qua kiểm tra.

Phát hành bản mới: sửa version.py -> chạy build.bat -> tạo Release trên GitHub (tag v<phiên bản>)
và đính kèm ToolNHCS.exe + ToolNHCS.exe.sha256 (xem README.md).
"""
import hashlib
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import tkinter as tk
from tkinter import ttk

from tools import _paths as paths
from tools import _ui as ui
from version import APP_NAME, APP_VERSION

# ---------- Cấu hình ----------
REPO = "baubihg/ToolNHCS"
ASSET_NAME = "ToolNHCS.exe"          # tên file đính kèm trong Release
WEB = "https://github.com"
API = "https://api.github.com"
OFFLINE_GRACE_DAYS = 7               # số ngày được dùng tạm khi không kiểm tra được cập nhật
TIMEOUT = 8                          # giây, cho thao tác kiểm tra
USER_AGENT = "ToolNHCS-Updater"


class UpdateError(Exception):
    pass


class NetworkError(UpdateError):
    pass


# ================== TIỆN ÍCH ==================
def is_frozen():
    return bool(getattr(sys, "frozen", False))


def exe_path():
    return os.path.abspath(sys.executable)


def parse_version(text):
    nums = re.findall(r"\d+", text or "")
    return tuple(int(n) for n in nums) or (0,)


def is_newer(remote, local):
    a, b = list(parse_version(remote)), list(parse_version(local))
    n = max(len(a), len(b))
    a += [0] * (n - len(a))
    b += [0] * (n - len(b))
    return a > b


def mark_checked_ok():
    paths.save_section("update", {"last_ok": time.time(), "version": APP_VERSION})


def within_grace(now=None):
    """True nếu lần kiểm tra thành công gần nhất cách đây chưa quá OFFLINE_GRACE_DAYS ngày."""
    try:
        last = float(paths.load_section("update").get("last_ok", 0))
    except (ValueError, TypeError):
        return False
    return 0 <= (now or time.time()) - last <= OFFLINE_GRACE_DAYS * 86400


def cleanup_leftovers():
    """Xóa file rác của lần cập nhật trước (bản cũ, bản tải dở, script hỗ trợ)."""
    exe = exe_path()
    for path in (exe + ".old", exe + ".new", exe + ".new.part",
                 os.path.join(os.path.dirname(exe), "update_helper.bat")):
        try:
            os.remove(path)
        except OSError:
            pass


# ================== MẠNG ==================
def _open(url, timeout=TIMEOUT):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "*/*"})
    return urllib.request.urlopen(req, timeout=timeout)


def fetch_latest():
    """Thông tin Release mới nhất: dict(tag, version, url, sha_url, notes) hoặc None nếu chưa có Release.
    Ném NetworkError khi không kết nối được."""
    try:
        with _open(f"{WEB}/{REPO}/releases/latest") as resp:    # GitHub chuyển hướng sang .../releases/tag/<tag>
            final = resp.geturl()
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise NetworkError(f"GitHub trả về lỗi {e.code}") from e
    except (urllib.error.URLError, OSError) as e:
        raise NetworkError(str(getattr(e, "reason", e))) from e

    m = re.search(r"/releases/tag/([^/?#]+)", final)
    if not m:
        return None
    tag = urllib.parse.unquote(m.group(1))
    base = f"{WEB}/{REPO}/releases/download/{urllib.parse.quote(tag)}"
    info = {"tag": tag, "version": tag.lstrip("vV"), "url": f"{base}/{ASSET_NAME}",
            "sha_url": f"{base}/{ASSET_NAME}.sha256", "notes": ""}
    try:                                                      # ghi chú phát hành: có thì hiện, không có thì thôi
        with _open(f"{API}/repos/{REPO}/releases/tags/{urllib.parse.quote(tag)}") as resp:
            info["notes"] = (json.load(resp).get("body") or "").strip()
    except Exception:
        pass
    return info


def download(info, dest, progress=None):
    """Tải file exe mới về dest, kiểm tra đủ dung lượng + SHA256 (nếu Release có file .sha256)."""
    expected = None
    try:
        with _open(info["sha_url"]) as resp:
            m = re.search(r"[0-9a-fA-F]{64}", resp.read(4096).decode("utf-8", "ignore"))
            expected = m.group(0).lower() if m else None
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise NetworkError(f"Không tải được mã kiểm tra (lỗi {e.code})") from e
    except (urllib.error.URLError, OSError) as e:
        raise NetworkError(str(getattr(e, "reason", e))) from e

    tmp = dest + ".part"
    sha = hashlib.sha256()
    got = total = 0
    try:
        with _open(info["url"], timeout=30) as resp, open(tmp, "wb") as fh:
            total = int(resp.headers.get("Content-Length") or 0)
            while True:
                chunk = resp.read(256 * 1024)
                if not chunk:
                    break
                fh.write(chunk)
                sha.update(chunk)
                got += len(chunk)
                if progress:
                    progress(got, total)
    except urllib.error.HTTPError as e:
        _silent_remove(tmp)
        raise UpdateError(f"Không tải được bản cập nhật (lỗi {e.code}). "
                          f"Release cần có file đính kèm tên {ASSET_NAME}.") from e
    except (urllib.error.URLError, OSError) as e:
        _silent_remove(tmp)
        raise NetworkError(str(getattr(e, "reason", e))) from e

    if total and got != total:
        _silent_remove(tmp)
        raise UpdateError("Tải về bị thiếu dữ liệu, hãy thử lại.")
    if expected and sha.hexdigest() != expected:
        _silent_remove(tmp)
        raise UpdateError("Sai mã kiểm tra SHA256, file tải về không đáng tin.")
    with open(tmp, "rb") as fh:
        if fh.read(2) != b"MZ":                               # file chương trình Windows luôn bắt đầu bằng "MZ"
            _silent_remove(tmp)
            raise UpdateError("File tải về không phải chương trình hợp lệ.")
    os.replace(tmp, dest)


def _silent_remove(path):
    try:
        os.remove(path)
    except OSError:
        pass


# ================== THAY FILE EXE & KHỞI ĐỘNG LẠI ==================
def replace_executable(exe, new):
    """Windows cho phép ĐỔI TÊN file exe đang chạy (không cho ghi đè/xóa): đổi exe -> .old rồi đặt bản mới vào."""
    old = exe + ".old"
    _silent_remove(old)
    os.replace(exe, old)
    try:
        os.replace(new, exe)
    except OSError:
        os.replace(old, exe)                                  # lỗi giữa chừng: trả lại bản cũ
        raise


def _fresh_env():
    """PyInstaller (onefile): bắt tiến trình con tự giải nén riêng, không dùng chung thư mục tạm với tiến trình này."""
    env = os.environ.copy()
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
    for key in [k for k in env if k.startswith("_PYI") or k == "_MEIPASS2"]:
        env.pop(key, None)
    return env


def relaunch(exe):
    flags = 0x00000008 | 0x00000200 if os.name == "nt" else 0      # DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP
    subprocess.Popen([exe], env=_fresh_env(), cwd=os.path.dirname(exe), close_fds=True, creationflags=flags)


def run_helper_script(exe, new):
    """Phương án dự phòng khi không đổi tên được exe: script chờ chương trình tắt rồi thay file và mở lại."""
    bat = os.path.join(os.path.dirname(exe), "update_helper.bat")
    with open(bat, "w", encoding="mbcs", errors="replace") as fh:
        fh.write(f'''@echo off
set /a n=0
:retry
ping 127.0.0.1 -n 2 >nul
move /y "{new}" "{exe}" >nul 2>&1
if errorlevel 1 (
  set /a n+=1
  if %n% lss 60 goto retry
  exit /b 1
)
start "" "{exe}"
del "%~f0"
''')
    flags = 0x00000008 | 0x00000200 if os.name == "nt" else 0
    subprocess.Popen(["cmd", "/c", bat], env=_fresh_env(), cwd=os.path.dirname(exe), close_fds=True,
                     creationflags=flags)


# ================== MÀN HÌNH CHẶN ==================
class UpdateGate(tk.Frame):
    """Phủ kín cửa sổ launcher cho tới khi kiểm tra xong. Chưa cập nhật thì không vào được."""

    def __init__(self, root, on_pass):
        super().__init__(root, bg=ui.BG)
        self.root, self.on_pass = root, on_pass
        self.q = queue.Queue()
        self.info = None
        self._alive = True
        self.place(x=0, y=0, relwidth=1, relheight=1)
        root.protocol("WM_DELETE_WINDOW", self.quit_app)        # đóng cửa sổ = thoát chương trình

        tk.Label(self, text=APP_NAME, bg=ui.BG, fg=ui.TEXT,
                 font=ui.font(22, True)).place(relx=0.5, rely=0.2, anchor="center")
        tk.Label(self, text=f"Phiên bản {APP_VERSION}", bg=ui.BG, fg=ui.MUTED,
                 font=ui.font(10)).place(relx=0.5, rely=0.26, anchor="center")

        self.card = ui.Card(self, pad=22)
        self.card.place(relx=0.5, rely=0.55, anchor="center", width=580)
        c = self.card.content
        self.title_lbl = tk.Label(c, text="", bg=ui.CARD, fg=ui.TEXT, font=ui.font(14, True), anchor="w")
        self.title_lbl.pack(fill="x")
        self.msg_lbl = tk.Label(c, text="", bg=ui.CARD, fg=ui.MUTED, font=ui.font(10), anchor="w",
                                justify="left", wraplength=520)
        self.msg_lbl.pack(fill="x", pady=(6, 8))
        self.notes = tk.Text(c, height=6, wrap="word", font=ui.font(9), bg="#f6f8fa", relief="flat",
                             highlightthickness=1, highlightbackground=ui.BORDER, padx=8, pady=6)
        self.bar = ttk.Progressbar(c, mode="indeterminate", length=520)
        self.bar_lbl = tk.Label(c, text="", bg=ui.CARD, fg=ui.MUTED, font=ui.font(9), anchor="w")
        self.btns = tk.Frame(c, bg=ui.CARD)

        self._show_checking()
        self.after(100, self._poll)
        threading.Thread(target=self._check_worker, daemon=True).start()

    # ---------- khung hiển thị ----------
    def _layout(self, title, msg, notes="", bar=None, bar_text="", buttons=()):
        """bar: None = ẩn, 'wait' = chạy qua lại, 'det' = có phần trăm."""
        self.title_lbl.config(text=title)
        self.msg_lbl.config(text=msg)
        for w in (self.notes, self.bar, self.bar_lbl, self.btns):
            w.pack_forget()
        if notes:
            self.notes.config(state="normal")
            self.notes.delete("1.0", "end")
            self.notes.insert("1.0", notes[:1500])
            self.notes.config(state="disabled")
            self.notes.pack(fill="x", pady=(0, 8))
        self.bar.stop()
        if bar:
            self.bar.config(mode="indeterminate" if bar == "wait" else "determinate", value=0, maximum=100)
            self.bar.pack(fill="x", pady=(2, 4))
            if bar == "wait":
                self.bar.start(12)
            self.bar_lbl.config(text=bar_text)
            self.bar_lbl.pack(fill="x")
        for w in self.btns.winfo_children():
            w.destroy()
        for text, cmd, primary in buttons:
            if primary:
                b = ui.RoundButton(self.btns, text=text, command=cmd, padx=22, pady=9)
            else:
                b = ui.secondary_button(self.btns, text, cmd, padx=22, pady=9)
            b.pack(side="left", padx=(0, 10))
        if buttons:
            self.btns.pack(fill="x", pady=(14, 0))

    def _show_checking(self):
        self._layout("Đang kiểm tra cập nhật...", "Vui lòng đợi trong giây lát.", bar="wait",
                     buttons=[("Thoát", self.quit_app, False)])

    def _show_required(self, info):
        self._layout(
            f"Có phiên bản mới: {info['version']}",
            f"Phiên bản hiện tại: {APP_VERSION}.\nBạn cần cập nhật để tiếp tục sử dụng. "
            "Nếu không cập nhật, chương trình sẽ đóng.",
            notes=info.get("notes", ""),
            buttons=[("Cập nhật ngay", self._start_download, True), ("Thoát", self.quit_app, False)])

    def _show_downloading(self):
        self._layout("Đang tải bản cập nhật...", "Không tắt chương trình trong lúc tải.", bar="det",
                     bar_text="Đang kết nối...", buttons=[("Thoát", self.quit_app, False)])

    def _show_installing(self):
        self._layout("Đang cài đặt...", "Chương trình sẽ tự khởi động lại bằng phiên bản mới.", bar="wait")

    def _show_error(self, title, msg, retry):
        self._layout(title, msg, buttons=[("Thử lại", retry, True), ("Thoát", self.quit_app, False)])

    # ---------- kiểm tra ----------
    def _check_worker(self):
        try:
            self.q.put(("checked", fetch_latest(), None))
        except Exception as e:                                  # mạng lỗi hoặc GitHub không phản hồi
            self.q.put(("checked", None, str(e) or e.__class__.__name__))

    def _retry(self):
        self._show_checking()
        threading.Thread(target=self._check_worker, daemon=True).start()

    def _on_checked(self, info, err):
        if err:
            if within_grace():                                  # tạm cho dùng khi offline
                self._pass()
            else:
                self._show_error(
                    "Không kiểm tra được cập nhật",
                    "Chương trình cần kết nối Internet để kiểm tra phiên bản mới "
                    f"(được dùng tạm {OFFLINE_GRACE_DAYS} ngày kể từ lần kiểm tra gần nhất).\n"
                    f"Chi tiết: {err}", self._retry)
            return
        mark_checked_ok()
        if info and is_newer(info["version"], APP_VERSION):
            self.info = info
            self._show_required(info)
        else:
            self._pass()

    # ---------- tải & cài ----------
    def _start_download(self):
        self._show_downloading()
        dest = exe_path() + ".new"

        def work():
            try:
                download(self.info, dest, progress=lambda g, t: self.q.put(("progress", g, t)))
                self.q.put(("downloaded", dest))
            except Exception as e:
                self.q.put(("dl_error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _install(self, new_path):
        self._show_installing()
        self.update_idletasks()
        exe = exe_path()
        try:
            replace_executable(exe, new_path)
            relaunch(exe)
        except OSError:
            try:
                run_helper_script(exe, new_path)               # dự phòng
            except Exception as e:
                self._show_error("Không cài được bản cập nhật",
                                 f"{e}\nHãy tải bản mới từ GitHub và chép đè lên file chương trình.",
                                 self._start_download)
                return
        self.root.after(300, lambda: os._exit(0))

    # ---------- vòng nhận kết quả từ luồng nền ----------
    def _poll(self):
        if not self._alive:
            return
        try:
            while True:
                self._handle(self.q.get_nowait())
        except queue.Empty:
            pass
        if self._alive:                                         # _handle có thể đã đóng màn hình chặn
            self.after(100, self._poll)

    def _handle(self, msg):
        kind = msg[0]
        if kind == "checked":
            self._on_checked(msg[1], msg[2])
        elif kind == "progress":
            got, total = msg[1], msg[2]
            if total:
                self.bar.config(value=got * 100 / total)
                self.bar_lbl.config(text=f"{got / 1048576:.1f} / {total / 1048576:.1f} MB")
            else:
                self.bar_lbl.config(text=f"{got / 1048576:.1f} MB")
        elif kind == "downloaded":
            self._install(msg[1])
        elif kind == "dl_error":
            self._show_error("Tải bản cập nhật thất bại", msg[1], self._start_download)

    # ---------- kết thúc ----------
    def _pass(self):
        self._alive = False
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        self.destroy()
        self.on_pass()

    def quit_app(self):
        self._alive = False
        self.root.destroy()
