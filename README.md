# ToolNHCS by Mr.CK

Bộ công cụ nghiệp vụ Ngân hàng Chính sách xã hội (launcher + các tool trong thư mục `tools/`).

## Cấu trúc

```
launcher.py        Cửa sổ chính, tự quét thư mục tools/
assets/app.ico     Icon chương trình
updater.py         Cập nhật bắt buộc qua GitHub Releases
version.py         Số phiên bản (APP_VERSION)
build.bat          Đóng gói .exe + tạo mã SHA256
tools/             Các công cụ (_ui.py, _shared.py là file dùng chung)
tools/templates/   File mẫu Word (hopdongmau.docx)
```

## Icon

Icon nằm ở `assets\app.ico` và được dùng ở 2 nơi: **file exe** (`--icon` trong `build.bat`, hiện ở Explorer / màn hình nền / thanh tác vụ)
và **cửa sổ chương trình** (`launcher.py` nạp icon từ gói exe). Muốn đổi icon:
1. Chuẩn bị ảnh logo vuông, nền trong suốt, nên từ 256x256 trở lên (PNG).
2. Đổi sang `.ico` nhiều kích thước và ghi đè `assets\app.ico`:
   ```
   python -c "from PIL import Image; Image.open('logo.png').convert('RGBA').save('assets/app.ico', sizes=[(16,16),(24,24),(32,32),(48,48),(64,64),(128,128),(256,256)])"
   ```
3. Tăng số trong `version.py`, chạy `build.bat`, phát hành như bình thường.
Nếu Explorer vẫn hiện icon cũ sau khi build: đổi tên file exe tạm thời hoặc chép sang thư mục khác (Windows lưu đệm icon).

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
- **Thư mục xuất file**: lần chạy đầu, chương trình tự tạo 2 thư mục **cạnh file exe**:
  `Thong bao phe duyet\` (tool Thông báo giải ngân; trong đó tự chia theo ngày/Mau_04/Mau_06) và
  `Hop dong GQVL\` (tool Hợp đồng GQVL; mỗi đợt tạo hàng loạt nằm trong một thư mục con `Hang loat <ngày giờ>`).
  Tool Thông báo giải ngân có nút **Chọn** để đổi thư mục xuất, sau khi tạo xong sẽ hỏi có mở thư mục không.
  Nếu thư mục cạnh exe không ghi được (ví dụ `Program Files`) thì dùng `Documents\ToolNHCS`.
- **Số thứ tự tool**: tự đánh theo thứ tự `TOOL_ORDER` (1, 2, 3...). Tool mới thêm vào `tools\` tự có số tiếp theo.
- Ngoài 2 thư mục xuất file, cạnh file exe không có gì khác. Dữ liệu và cấu hình nằm gọn trong **một thư mục riêng** trên máy người dùng:
  `%LOCALAPPDATA%\ToolNHCS\` (mở nhanh: nhấn Win+R, gõ `%LOCALAPPDATA%\ToolNHCS`):
  ```
  config.json        cấu hình các tool + trạng thái cập nhật
  data\saoke.pkl     sao kê dùng chung
  data\hongheo.pkl   danh sách hộ nghèo
  logs\app.log       nhật ký lỗi (chỉ có khi có lỗi/ghi chú)
  ```
  Chạy bằng `python launcher.py` hay bằng exe đều dùng chung thư mục này. File cũ (cạnh exe hoặc trong `tools\`)
  tự được chuyển sang đây ở lần chạy đầu. Muốn "đặt lại như mới": đóng chương trình rồi xóa thư mục này.
  Các file này chứa thông tin khách hàng nên **không bao giờ đưa lên GitHub**.
- Bật xác thực 2 lớp cho tài khoản GitHub: ai có quyền tạo Release là có quyền đưa chương trình mới đến mọi máy.
- Nếu mạng cơ quan chặn `github.com`, chương trình sẽ không kiểm tra được và bị chặn sau thời gian dùng tạm.
