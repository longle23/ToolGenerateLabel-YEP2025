## Label Generator YEP 2025

Tool Node.js dùng `canvas` và `csv-parser` để generate label nhân viên (PNG) từ file CSV.

### Cấu trúc

```text
label-generator/
 ├── index.js
 ├── config.json
 ├── input.csv          # copy từ RegistrationYEP2025.csv
 ├── package.json
 ├── fonts/
 │    └── Roboto-Bold.ttf (tự copy vào)
 └── output/            # ảnh PNG sẽ được sinh ở đây
```

### Cách chuẩn bị

1. Copy file `RegistrationYEP2025.csv` vào thư mục `label-generator` và đổi tên thành `input.csv`.
2. Tạo thư mục `fonts/` và copy file font (vd: `Roboto-Bold.ttf`) vào đó.
3. Tạo thư mục trống `output/`.

### Cài đặt

```bash
cd label-generator
npm install
```

> Nếu cài `canvas` trên Windows bị lỗi, hãy cài thêm Visual Studio Build Tools theo hướng dẫn trên trang npm của `canvas`.

### Chạy

```bash
cd label-generator
node index.js
# hoặc
npm start
```

Mỗi dòng trong `input.csv` (có header `STT,Code,Name`) sẽ sinh một file PNG trong thư mục `output/` với tên dạng `label_001.png`, `label_002.png`, ...


