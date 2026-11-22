# -*- coding: utf-8 -*-
"""
Nơi lưu mọi file phát sinh khi chạy chương trình - gom vào MỘT thư mục, không rải file cạnh file .exe.

  Thư mục XUẤT FILE (tự tạo ở lần chạy đầu, cạnh file exe):
        Thong bao phe duyet\\        nơi xuất file của tool Thông báo giải ngân
        Hop dong GQVL\\              nơi xuất file của tool Hợp đồng GQVL
    Nếu thư mục cạnh exe không ghi được (vd: Program Files) thì dùng Documents\\ToolNHCS.

    %LOCALAPPDATA%\\ToolNHCS\\          (mở nhanh: nhấn Win+R, gõ  %LOCALAPPDATA%\\ToolNHCS )
        config.json                  cấu hình của các tool + trạng thái cập nhật (1 file duy nhất)
        data\\saoke.pkl               sao kê dùng chung
        data\\hongheo.pkl             danh sách hộ nghèo
        logs\\app.log                 nhật ký lỗi (tự xoay vòng, chỉ tạo khi có gì để ghi)

Chạy bằng "python launcher.py" hay bằng file .exe đều dùng chung thư mục này.
File cũ (saoke_data.pkl, config_*.json, thongbao.log...) cạnh exe / trong tools\\ được tự chuyển sang
thư mục mới ở lần chạy đầu rồi xóa khỏi chỗ cũ.

Đặt biến môi trường NHCS_DATA_DIR để dùng thư mục khác (ví dụ khi thử nghiệm).
File bắt đầu bằng "_" nên launcher không coi đây là một tool.
"""
import json
import logging
import logging.handlers
import os
import shutil
import subprocess
import sys
import tempfile

APP_DIR_NAME = "ToolNHCS"
_root = None
_migrated = False
_logging_ready = False


# ================== ĐƯỜNG DẪN ==================
def data_root():
    """Thư mục gốc chứa dữ liệu (tự tạo, tự chuyển file cũ sang ở lần gọi đầu tiên)."""
    global _root
    if _root:
        return _root
    base = os.environ.get("NHCS_DATA_DIR")
    if not base:
        local = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".local", "share")
        base = os.path.join(local, APP_DIR_NAME)
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:                                   # không ghi được (hiếm): dùng thư mục tạm
        base = os.path.join(tempfile.gettempdir(), APP_DIR_NAME)
        os.makedirs(base, exist_ok=True)
    _root = base
    migrate_legacy()
    return _root


def config_path():
    return os.path.join(data_root(), "config.json")


def data_path(filename):
    """Đường dẫn file dữ liệu trong thư mục data\\ (thư mục chỉ được tạo khi ghi: dùng ensure_parent)."""
    return os.path.join(data_root(), "data", filename)


def log_path():
    return os.path.join(data_root(), "logs", "app.log")


def ensure_parent(path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    return path


# ================== THƯ MỤC XUẤT FILE ==================
OUTPUT_DIRS = {"thongbao": "Thong bao phe duyet", "hopdong": "Hop dong GQVL"}
_out_base = None


def _app_dir():
    """Thư mục chứa file exe (hoặc chứa launcher.py khi chạy bằng Python)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def resource_path(rel):
    """Đường dẫn file đi kèm chương trình (icon...). Trong file exe nằm ở thư mục tạm của PyInstaller,
    khi chạy bằng Python thì nằm cạnh launcher.py."""
    base = getattr(sys, "_MEIPASS", None) if getattr(sys, "frozen", False) else None
    return os.path.join(base or _app_dir(), rel)


def output_base():
    """Thư mục chứa các thư mục xuất file. Chọn chỗ đầu tiên tạo được:
    (NHCS_OUTPUT_DIR nếu có) -> cạnh file exe -> Documents\\ToolNHCS."""
    global _out_base
    if _out_base:
        return _out_base
    candidates = [os.environ.get("NHCS_OUTPUT_DIR"), _app_dir(),
                  os.path.join(os.path.expanduser("~"), "Documents", APP_DIR_NAME)]
    for base in filter(None, candidates):
        try:
            for name in OUTPUT_DIRS.values():
                os.makedirs(os.path.join(base, name), exist_ok=True)
            _out_base = base
            return base
        except OSError:
            continue
    _out_base = os.path.join(tempfile.gettempdir(), APP_DIR_NAME)
    return _out_base


def output_dir(key):
    """Thư mục xuất file của một tool ('thongbao' | 'hopdong'); tự tạo lại nếu bị xóa."""
    path = os.path.join(output_base(), OUTPUT_DIRS[key])
    try:
        os.makedirs(path, exist_ok=True)
    except OSError as e:
        logging.warning("Không tạo được thư mục %s: %s", path, e)
    return path


def ensure_output_dirs():
    for key in OUTPUT_DIRS:
        output_dir(key)


def open_folder(path):
    """Mở thư mục trong trình quản lý file (Windows Explorer)."""
    try:
        os.makedirs(path, exist_ok=True)
        if hasattr(os, "startfile"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as e:
        logging.warning("Không mở được thư mục %s: %s", path, e)


# ================== CẤU HÌNH (một file config.json, mỗi tool một mục) ==================
def _read_json(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            d = json.load(fh)
        return d if isinstance(d, dict) else {}
    except (OSError, ValueError):
        return {}


def _write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)                             # ghi nguyên tử: mất điện giữa chừng không hỏng file


def load_section(name):
    d = _read_json(config_path()).get(name)
    return dict(d) if isinstance(d, dict) else {}


def save_section(name, data):
    """Lưu cấu hình của một tool, giữ nguyên mục của các tool khác. Trả về True nếu ghi được."""
    cfg = _read_json(config_path())
    cfg[name] = data
    try:
        _write_json(config_path(), cfg)
        return True
    except OSError as e:
        logging.warning("Không lưu được cấu hình '%s': %s", name, e)
        return False


# ================== NHẬT KÝ ==================
class _LazyRotatingHandler(logging.handlers.RotatingFileHandler):
    """Chỉ tạo thư mục logs\\ và file app.log khi thật sự có dòng nhật ký đầu tiên."""

    def _open(self):
        os.makedirs(os.path.dirname(self.baseFilename), exist_ok=True)
        return super()._open()


def setup_logging():
    """Ghi nhật ký vào logs\\app.log (tối đa ~1 MB). Gọi nhiều lần cũng chỉ cài một lần."""
    global _logging_ready
    if _logging_ready:
        return
    _logging_ready = True
    root = logging.getLogger()
    if any(getattr(h, "_nhcs", False) for h in root.handlers):
        return
    try:
        handler = _LazyRotatingHandler(log_path(), maxBytes=512 * 1024, backupCount=1,
                                       encoding="utf-8", delay=True)
    except OSError:
        return
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    handler._nhcs = True
    root.addHandler(handler)
    root.setLevel(logging.INFO)


# ================== CHUYỂN FILE CŨ ==================
_LEGACY_FILES = {                                     # tên cũ -> (thư mục con, tên mới)
    "saoke_data.pkl": ("data", "saoke.pkl"),
    "hongheo_data.pkl": ("data", "hongheo.pkl"),
    "tracuu_data.pkl": ("data", "tracuu_legacy.pkl"),
}
_LEGACY_CONFIGS = {                                   # file json cũ -> mục trong config.json
    "config_thongbao.json": "thongbao",
    "config_hopdong.json": "hopdong",
    "update_state.json": "update",
}
_LEGACY_JUNK = ("thongbao.log",)                      # file nhật ký cũ: xóa


def _legacy_dirs():
    if getattr(sys, "frozen", False):
        return [os.path.dirname(os.path.abspath(sys.executable))]
    here = os.path.dirname(os.path.abspath(__file__))          # thư mục tools\
    return [here, os.path.dirname(here)]


def migrate_legacy():
    """Gom file rải rác ở chỗ cũ về thư mục dữ liệu mới. An toàn khi gọi lại nhiều lần."""
    global _migrated
    if _migrated or not _root:
        return
    _migrated = True
    for d in _legacy_dirs():
        if os.path.abspath(d) == os.path.abspath(_root):
            continue
        try:
            for old, (sub, new) in _LEGACY_FILES.items():
                src, dst = os.path.join(d, old), os.path.join(_root, sub, new)
                if os.path.isfile(src) and not os.path.exists(dst):
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.move(src, dst)
            for old, section in _LEGACY_CONFIGS.items():
                src = os.path.join(d, old)
                if not os.path.isfile(src):
                    continue
                cfg_file = os.path.join(_root, "config.json")
                cfg, old_data = _read_json(cfg_file), _read_json(src)
                if old_data and section not in cfg:
                    cfg[section] = old_data
                    _write_json(cfg_file, cfg)
                os.remove(src)
            for old in _LEGACY_JUNK:
                src = os.path.join(d, old)
                if os.path.isfile(src):
                    os.remove(src)
        except OSError as e:                          # file đang bị khóa...: bỏ qua, lần sau thử lại
            logging.warning("Chưa chuyển hết file cũ: %s", e)
