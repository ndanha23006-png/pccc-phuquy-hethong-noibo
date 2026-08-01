# Hướng dẫn triển khai Web App — Hệ thống biểu mẫu PCCC Phú Quý

Bộ file này biến trang HTML thành **Web App Google Apps Script thật**: mỗi lần gửi biểu mẫu "Xử lý phản hồi tổ đội" sẽ **tạo 1 file Word (.docx) vào thư mục Drive** và (tuỳ chọn) gửi email nội bộ kèm file.

Kiến trúc **nhiều trang riêng** để dễ quản trị: mỗi tính năng là 1 file `.html`. Trang chủ (`index`) chỉ chứa các nút dẫn sang từng trang.

```
Code.gs         → backend: router + phiếu Word + báo cáo tuần (Excel)
styles.html     → CSS dùng chung (mọi trang nhúng file này)
index.html      → Trang chủ (các nút công cụ)
phanhoi.html    → Biểu mẫu "Xử lý phản hồi tổ đội"
baocaohub.html  → Trang danh sách tuần + thêm tuần mới
baocao.html     → Form nhập báo cáo tuần (ghi vào Excel)
appsscript.json → cấu hình quyền (manifest)
```

> Khi tạo file HTML trên Apps Script, gõ tên **không kèm `.html`**: `styles`, `index`, `phanhoi`, `baocaohub`, `baocao`.

---

## BƯỚC 1 — Tạo thư mục Drive lưu phiếu Word

1. Vào https://drive.google.com → **Mới → Thư mục mới** → đặt tên ví dụ: `Phiếu xử lý phản hồi tổ đội`.
2. Mở thư mục đó. Nhìn lên thanh địa chỉ, sao chép đoạn **ID** (phần sau `/folders/`):
   ```
   https://drive.google.com/drive/folders/1AbCdEfGhIJKlmNOPqrs_TUvWxyz  ← ID là đoạn bôi đậm
   ```
3. Giữ lại ID này cho Bước 3.

---

## BƯỚC 2 — Tạo dự án Apps Script và dán các file

1. Vào https://script.google.com → **Dự án mới (New project)**.
2. Đổi tên dự án (góc trên trái) thành `PCCC Phú Quý - Hệ thống nội bộ`.
3. **Tạo đủ 4 file** trong dự án (nút **+** cạnh "Files"):
   - File `Code.gs` đã có sẵn → **xoá hết** nội dung mặc định, dán toàn bộ nội dung file **`Code.gs`** của tôi vào.
   - Bấm **+ → HTML**, đặt tên `styles` → dán nội dung **`styles.html`**.
   - Bấm **+ → HTML**, đặt tên `index` → dán nội dung **`index.html`**.
   - Bấm **+ → HTML**, đặt tên `phanhoi` → dán nội dung **`phanhoi.html`**.

   > Lưu ý: khi tạo file HTML, Apps Script tự thêm đuôi `.html`. Bạn chỉ gõ tên `styles`, `index`, `phanhoi` (không gõ `.html`).

4. (Tuỳ chọn nhưng nên làm) Hiện file manifest: **Cài đặt dự án (bánh răng) → bật "Hiển thị tệp kê khai appsscript.json"**. Quay lại, mở `appsscript.json`, dán nội dung file manifest của tôi vào.

---

## BƯỚC 3 — Điền cấu hình trong Code.gs

Ở đầu file `Code.gs`, sửa 3 dòng:

```javascript
var FOLDER_ID = 'DAN_ID_THU_MUC_DRIVE_VAO_DAY';        // ← dán ID thư mục ở Bước 1
var EMAIL_NHAN_MAC_DINH = 'ndanha23006@cusc.ctu.edu.vn'; // ← email nhận tin nội bộ
var GUI_EMAIL = true;                                   // true = gửi mail kèm file | false = chỉ tạo file Word
var BAOCAO_FOLDER_ID = '1W-F_GHTIf--zuY0v79D0YfkB6KNE9EXC'; // ← thư mục chứa file Excel báo cáo tuần
```

> `BAOCAO_FOLDER_ID` đã điền sẵn thư mục "Báo cáo công việc tuần - App hệ thống nội bộ" của anh. Nếu đổi thư mục khác thì thay ID tại đây.

Bấm **Lưu** (biểu tượng đĩa mềm hoặc Ctrl+S).

---

## BƯỚC 4 — Cấp quyền lần đầu

1. Trên thanh công cụ, chọn hàm **`kiemTraCauHinh`** trong ô dropdown → bấm **Run (Chạy)**.
2. Google hiện bảng xin quyền → **Review permissions → chọn tài khoản → Advanced → Go to ... (unsafe) → Allow**.
   (Đây là app của chính bạn nên an toàn; cảnh báo "unsafe" là mặc định của Google với app chưa xác minh.)
3. Vào menu **View → Logs** (hoặc Execution log). Nếu thấy dòng `OK - Thư mục lưu trữ: ...` là cấu hình đúng.

---

## BƯỚC 5 — Triển khai Web App

1. Góc trên phải: **Deploy (Triển khai) → New deployment**.
2. Bấm bánh răng cạnh "Select type" → chọn **Web app**.
3. Điền:
   - **Description**: `Phiên bản 1`
   - **Execute as**: **Me (chính bạn)** — để hệ thống dùng Drive của bạn tạo file.
   - **Who has access**:
     - `Anyone` (bất kỳ ai có link) — tiện cho nhân viên dùng nhanh, hoặc
     - `Anyone with Google account` — an toàn hơn, cần đăng nhập Google.
4. Bấm **Deploy** → sao chép **Web app URL**. Đây là **link hệ thống** để gửi cho nhân viên (mở được trên điện thoại).

> Mỗi lần sửa code sau này: **Deploy → Manage deployments → (bút chì) Edit → Version: New version → Deploy** để cập nhật cùng một link.

---

## BƯỚC 6 — Kiểm tra

1. Mở Web app URL → thấy **trang chủ** với các nút.
2. Bấm **"Xử lý phản hồi tổ đội"** → điền form → **Tạo phiếu & lưu file Word**.
3. Kiểm tra: thư mục Drive xuất hiện file `PH-....docx`; nếu bật `GUI_EMAIL` thì email nội bộ nhận được thư kèm file.
4. Thử bỏ trống 1 ô bắt buộc → hệ thống chặn và báo trường thiếu (không tạo file).

---

## THÊM TÍNH NĂNG MỚI SAU NÀY (trang riêng)

Ví dụ làm "Báo cáo công việc tuần":

1. Tạo file HTML mới, ví dụ `baocao` (dán markup trang mới, nhớ `<?!= include('styles') ?>` để dùng chung CSS).
2. Trong `Code.gs`, thêm tên trang vào danh sách:
   ```javascript
   var TRANG_HOP_LE = ['index', 'phanhoi', 'baocao'];
   ```
3. Ở `index.html`, đổi nút "Sắp ra mắt" tương ứng thành link:
   ```html
   <a class="app live" href="<?= baseUrl ?>?page=baocao" style="text-decoration:none;color:inherit"> ... </a>
   ```
4. Viết hàm xử lý riêng trong `Code.gs` (ví dụ `luuBaoCao(data)`), gọi bằng `google.script.run` như trang `phanhoi`.
5. **Deploy → New version**.

Mỗi trang tách biệt nên sửa trang này không ảnh hưởng trang kia — dễ bảo trì.

---

## KHI CÓ API ZALO (nâng cấp sau)

Trong `Code.gs`, thêm 1 hàm `guiZalo_(msg)` dùng `UrlFetchApp.fetch(...)` gọi endpoint Zalo OA, rồi chèn lời gọi ngay cạnh dòng `MailApp.sendEmail(...)`. Phần tạo file Word giữ nguyên.
