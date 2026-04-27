# VIE Charset Converter (Word Web Add-in)
Add-in cá nhân cho Word để chuyển tiếng Việt từ bảng mã non-Unicode (`TCVN3`, `VNI`) sang Unicode.

### Mô hình manifest (không xung đột local/prod)

- `manifest.xml`: bản production, trỏ GitHub Pages.
- `manifest.local.xml`: bản local dev, trỏ `https://localhost:3000`.
- Hai manifest dùng **2 ID khác nhau**, nên có thể cài song song và không ghi đè nhau.

## Chạy local (dev)

Yêu cầu:
- Node.js 18+
- Word Desktop (Microsoft 365)

Lần đầu trên mỗi máy:

```bash
npm install
npm run trust-cert
```

Chạy local:

```bash
npm start
```

`npm start` sẽ:
- chạy HTTPS server `https://localhost:3000`
- sideload `manifest.local.xml`
- mở Word Desktop với add-in local

Lệnh hữu ích:

```bash
npm run start:server
npm run stop:desktop
npm run validate
npm run validate:prod
```

## Thêm add-in trong Word

### Cách 1: Upload manifest (đơn giản nhất)

Word Online:
1. Mở Word Online.
2. `Insert` -> `Add-ins` -> `Manage My Add-ins`.
3. `Upload My Add-in` và chọn:
   - `manifest.local.xml` nếu test local
   - `manifest.xml` nếu dùng bản production

Word Desktop:
1. `Insert` -> `My Add-ins` -> `Manage My Add-ins`.
2. Trình duyệt mở trang quản lý add-in.
3. Upload file manifest tương ứng.

### Cách 2: Shared Folder Catalog (khi không upload được trên desktop)

1. Tạo thư mục, ví dụ `C:\OfficeAddins`.
2. Copy manifest vào đó (`manifest.xml` hoặc `manifest.local.xml`).
3. Chia sẻ thư mục (Properties -> Sharing -> Advanced Sharing).
4. Dùng UNC path: `\\localhost\OfficeAddins`.
5. Trong Word Desktop: `File` -> `Options` -> `Trust Center` -> `Trust Center Settings` -> `Trusted Add-in Catalogs`.
6. Ở `Catalog Url`, nhập `\\localhost\OfficeAddins`.
7. Tick `Show in Menu`, bấm `Add Catalog`, `OK`.
8. Đóng hẳn Word và mở lại.
9. Vào `Insert` -> `My Add-ins` -> `Shared Folder` để thêm add-in.

Lưu ý: ô `Catalog URL` không nhận `C:\...`, chỉ nhận `https://...` hoặc UNC `\\server\share`.

## Public lên GitHub Pages

Repo đã có workflow: `.github/workflows/deploy-pages.yml`.

Các bước:
1. Push code lên nhánh `main`.
2. Vào GitHub repo -> `Settings` -> `Pages` -> chọn `Source: GitHub Actions`.
3. Chờ workflow `Deploy GitHub Pages` chạy xong trong tab `Actions`.
4. Kiểm tra URL:
   - `https://nnhieu97.github.io/vie-converter/taskpane.html`
5. Upload `manifest.xml` vào Word để dùng bản online.

## Cấu trúc chính

- `src/lib/*`: engine detect/convert charset
- `src/taskpane/*`: UI + logic Preview/Apply
- `scripts/build.js`: build vào `dist/`
- `scripts/dev-server.js`: HTTPS local + auto rebuild
- `scripts/start-desktop.js`: sideload local manifest + mở Word
- `.github/workflows/deploy-pages.yml`: deploy GitHub Pages

## Tài liệu kiến trúc

- `architecture_flow.md`: sơ đồ khối tổng thể, sơ đồ runtime, và sequence diagram chi tiết của add-in
