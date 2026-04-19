# VIE Charset Converter (Word Web Add-in)

Word Web Add-in cá nhân để chuyển tiếng Việt từ bảng mã non-Unicode (`TCVN3`, `VNI`) sang Unicode.

## MVP hiện tại

- Chế độ an toàn: chỉ xử lý **vùng đang chọn**.
- Luồng làm việc: `Preview` trước, `Apply` sau.
- Nguồn bảng mã: `Auto detect` hoặc chọn tay `TCVN3` / `VNI`.
- Engine chuyển mã: [`vietnamese-conversion`](https://github.com/duydev/vietnamese-conversion).

## Yêu cầu

- Node.js 18+
- Microsoft Word Desktop (Microsoft 365) hoặc Word on the web.

## Chạy local (Desktop tự mở Word)

1. Cài package:

```bash
npm install
```

2. Cài certificate dev (mỗi máy làm 1 lần):

```bash
npm run trust-cert
```

3. Chạy lệnh start:

```bash
npm start
```

Lệnh này sẽ:
- chạy HTTPS server tại `https://localhost:3000`,
- sideload `manifest.xml`,
- mở Word Desktop với add-in.

## Chạy chỉ server (không mở Word)

```bash
npm run start:server
```

## Dừng sideload desktop

```bash
npm run stop:desktop
```

## Word on the web

Nếu dùng Word Online, vẫn có thể upload thủ công manifest:
1. `Insert` -> `Add-ins` -> `My Add-ins` -> `Upload My Add-in`.
2. Chọn `manifest.xml`.

## Dùng trên máy khác của bạn

1. Copy toàn bộ thư mục `vie-converter` sang máy mới.
2. Trên máy mới chạy:

```bash
npm install
npm run trust-cert
npm start
```

## Cấu trúc

- `manifest.xml`: khai báo add-in cho Word.
- `src/taskpane/*`: UI + logic convert.
- `scripts/build.js`: build JS bundle và copy static files vào `dist/`.
- `scripts/dev-server.js`: HTTPS server local + auto rebuild khi sửa file.
- `scripts/start-desktop.js`: start server và tự sideload/mở Word Desktop.

## Lệnh tiện ích

```bash
npm run build
npm run validate
```

## Giới hạn MVP

- Chưa bật chế độ `Main body` (đã để hook trong UI).
- Nếu đoạn có format rất phức tạp, replace text có thể ảnh hưởng format cục bộ. Nên chạy từng đoạn nhỏ.
