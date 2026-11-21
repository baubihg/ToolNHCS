# ToolNHCS

Bộ công cụ nghiệp vụ Ngân hàng Chính sách xã hội (launcher + các tool trong thư mục `tools/`).

## Cấu trúc

```
launcher.py        Cửa sổ chính, tự quét thư mục tools/
updater.py         Cập nhật bắt buộc qua GitHub Releases
version.py         Số phiên bản (APP_VERSION)
build.bat          Đóng gói .exe + tạo mã SHA256
tools/             Các công cụ (_ui.py, _shared.py là file dùng chung)
tools/templates/   File mẫu Word (hopdongmau.docx)
```

## Cách cập nhật hoạt động

- Mỗi lần mở file `.exe`, chương trình hỏi GitHub bản **Release mới nhất**.
- Nếu số phiên bản của Release **lớn hơn** `APP_VERSION` trong bản đang chạy: hiện màn hình bắt buộc cập nhật.
  Chọn **Cập nhật ngay** thì tự tải, thay file và mở lại; chọn **Thoát** (hoặc đóng cửa sổ) thì chương trình đóng.
- Mất mạng: được dùng tạm tối đa 7 ngày kể từ lần kiểm tra thành công gần nhất (`OFFLINE_GRACE_DAYS` trong `updater.py`).
- Chạy bằng `python launcher.py` thì bỏ qua kiểm tra.

## Phát hành phiên bản mới

1. Sửa `version.py`, ví dụ `APP_VERSION = "1.0.1"`.
2. Chạy `build.bat` (tạo `dist\ToolNHCS.exe` và `dist\ToolNHCS.exe.sha256`).
3. Đẩy mã nguồn lên GitHub:
   ```
   git add .
   git commit -m "Phiên bản 1.0.1"
   git push
   ```
4. Tạo **Release** trên GitHub (Releases → Draft a new release):
   - Tag: `v1.0.1` (đúng bằng số trong `version.py`, thêm chữ `v`)
   - Đính kèm đúng 2 file: `ToolNHCS.exe` và `ToolNHCS.exe.sha256` (tên file không được đổi)
   - Ô mô tả: ghi nội dung thay đổi, người dùng sẽ thấy ngay trong màn hình cập nhật
   - Bấm **Publish release** (không chọn "pre-release", vì bản pre-release không được coi là "latest")

Hoặc dùng GitHub CLI: `gh release create v1.0.1 dist\ToolNHCS.exe dist\ToolNHCS.exe.sha256 --title "v1.0.1" --notes "..."`

## Lần đầu đưa lên GitHub

```
git init
git branch -M main
git add .
git commit -m "Phiên bản 1.0.0"
git remote add origin https://github.com/baubihg/ToolNHCS.git
git push -u origin main
```
Sau đó tạo Release `v1.0.0` như trên và gửi file `ToolNHCS.exe` này cho người dùng **thủ công một lần**
(các bản cũ chưa có cơ chế cập nhật nên không tự nâng cấp được). Từ bản này trở đi sẽ tự cập nhật.

## Lưu ý

- Cần có số phiên bản **tăng dần**: tag `v1.0.2` mới hơn `v1.0.1`. Tăng lên mới bắt buộc người dùng cập nhật.
- Các file dữ liệu (`saoke_data.pkl`, `hongheo_data.pkl`, `config_*.json`) nằm cạnh file exe trên máy người dùng và
  đã được `.gitignore` chặn, **không bao giờ đưa lên GitHub** vì chứa thông tin khách hàng.
- Bật xác thực 2 lớp cho tài khoản GitHub: ai có quyền tạo Release là có quyền đưa chương trình mới đến mọi máy.
- Nếu mạng cơ quan chặn `github.com`, chương trình sẽ không kiểm tra được và bị chặn sau thời gian dùng tạm.
