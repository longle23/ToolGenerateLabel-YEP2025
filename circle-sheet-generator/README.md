# Circle Sheet Generator

Tool tạo sheet A4 với vòng tròn và số (3 cột x 5 dòng = 15 vòng tròn/sheet).

## Cách sử dụng

### Chạy với mặc định (số bắt đầu từ 001, tạo 1 sheet):
```bash
node generate-circle-sheet.cjs
```

### Chạy với số bắt đầu tùy chỉnh:
```bash
node generate-circle-sheet.cjs 50
```
Tạo 1 sheet bắt đầu từ số 050.

### Chạy với số bắt đầu và số lượng sheet:
```bash
node generate-circle-sheet.cjs 1 3
```
Tạo 3 sheet:
- Sheet 1: số 001-015
- Sheet 2: số 016-030
- Sheet 3: số 031-045

## Output

Files được lưu trong thư mục `circle-sheet-generator/output/`:
- `sheet_circle_001.png` - File PNG
- `sheet_circle_001.pdf` - File PDF

## Cấu hình

Tool sử dụng file `config.json` ở thư mục gốc để lấy cấu hình:
- Vòng tròn: `config.circle` (màu viền, độ dày, tỷ lệ bán kính)
- Số: `config.stt` (font, kích thước, màu sắc)

