# Excel API Updater

App desktop (Electron) cập nhật **Profile Name**, **Note**, **Proxy** của Hidemium từ file Excel.

## Chạy app

- Bản đóng gói: `dist\Updater_Profile_Hidemium-win32-x64\Updater_Profile_Hidemium.exe`
- Chạy từ source: `npm install` rồi `npm start`
- Đóng gói lại sau khi sửa code: `npm run package`
- Đóng gói + nén zip để mang sang máy khác: `npm run zip`
  → `dist\Updater_Profile_Hidemium-win32-x64.zip`. Máy đích chỉ cần Windows 64-bit và Hidemium đang chạy,
  giải nén rồi mở `Updater_Profile_Hidemium.exe` (không cần cài Node.js).

## Cấu trúc file Excel

| Cột | Nội dung | API gọi khi có giá trị |
| --- | --- | --- |
| A | UID profile | – |
| B | Profile Name | `PUT /v2/browser/update-once` `{ column: "name", profile_uuid, data }` |
| C | Note | `PUT /v2/browser/update-note` `{ note, profile_uuid }` |
| D | Proxy | `POST /v2/browser/proxy/update` `{ browser_update: [{ uuid, proxy }] }` |
| E | Status | App ghi kết quả vào cột này khi bấm **Xuất kết quả** (lưu ra file mới) |

- Dòng 1 là tiêu đề (bỏ qua nếu cột A không phải UUID). Dòng có cột A trống bị bỏ qua.
- Ô **trống** hoặc chữ **`null`** → không gọi API cho cột đó.
- Chế độ **Local** → `?is_local=true`, **Cloud** → `?is_local=false`.
  App cảnh báo nếu UID dạng `local-…` mà đang chọn Cloud (và ngược lại).
- Có thể chọn sheet, tick/bỏ tick từng dòng, Dừng giữa chừng, rồi bấm
  **Chọn dòng lỗi / chưa chạy** để chạy tiếp phần còn lại.

## Định dạng proxy (cột D)

Tự chuyển sang định dạng Hidemium `TYPE|host|port|user|pass`:

| Trong Excel | Gửi lên API |
| --- | --- |
| `127.0.0.1:60007` | `HTTP\|127.0.0.1\|60007` |
| `103.182.23.51:60009:user:pass` | `HTTP\|103.182.23.51\|60009\|user\|pass` |
| `http://user:pass@103.182.23.51:60009` | `HTTP\|103.182.23.51\|60009\|user\|pass` |
| `socks5://user:pass@1.2.3.4:1080` | `SOCKS5\|1.2.3.4\|1080\|user\|pass` |
| `SOCKS5\|1.2.3.4\|1080\|user\|pass` | giữ nguyên |

## Mã nguồn

- `main.js` – cửa sổ Electron, hộp thoại chọn/lưu file, IPC.
- `preload.js` – cầu nối `window.api` cho giao diện.
- `src/core.js` – đọc Excel, chuyển proxy, gọi API, ghi kết quả (Node thuần, test được độc lập).
- `renderer/` – giao diện (HTML/CSS/JS).
- `images/` – logo (`logo_*.png` dùng ở header) và `app.ico` (icon cửa sổ và file .exe, ghép từ 4 file PNG).
