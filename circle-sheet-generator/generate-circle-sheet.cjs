// generate-circle-sheet.cjs
// Tool tạo sheet A4 với vòng tròn và số (3 cột x 5 dòng)

const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const PDFDocument = require('pdfkit');
const Papa = require('papaparse');

// ===== Helper: convert mm -> px theo DPI =====
function mmToPx(mm, dpi) {
  // 1 inch = 25.4 mm
  return Math.round((mm / 25.4) * dpi);
}

// ===== Escape text cho SVG =====
function escapeXml(unsafe) {
  if (!unsafe) return '';
  return String(unsafe)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ===== Load config =====
const configPath = path.join(__dirname, '..', 'config.json');
if (!fs.existsSync(configPath)) {
  console.error('Không tìm thấy config.json');
  process.exit(1);
}
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const { page } = config;
const DPI = page.dpi || 300;

// ===== Tạo sheet A4 với vòng tròn và số (3 cột x 5 dòng) =====
async function createCircleNumberSheet(numbers, sheetIndex = 0) {
  // Kích thước A4 DỌC: 210mm x 297mm (portrait)
  const a4WidthMm = 210;
  const a4HeightMm = 297;
  const a4WidthPx = mmToPx(a4WidthMm, DPI);
  const a4HeightPx = mmToPx(a4HeightMm, DPI);
  
  // Layout: 5 hàng x 3 cột = 15 vòng tròn
  const cols = 3;
  const rows = 5;
  const totalCircles = cols * rows;
  
  // Kích thước mỗi cell trong sheet
  const cellWidthMm = a4WidthMm / cols;   // 210/3 = 70mm
  const cellHeightMm = a4HeightMm / rows; // 297/5 = 59.4mm
  const cellWidthPx = mmToPx(cellWidthMm, DPI);
  const cellHeightPx = mmToPx(cellHeightMm, DPI);
  
  // Lấy config cho vòng tròn và số
  const circleCfg = config.circle || {};
  const sttCfg = config.stt || {};
  const strokeColor = circleCfg.strokeColor || '#000000';
  const lineWidthPx = circleCfg.lineWidth || 4;
  const radiusRatio = circleCfg.radiusRatio || 0.35;
  
  // Tính bán kính vòng tròn dựa trên cell size
  const minCellSizePx = Math.min(cellWidthPx, cellHeightPx);
  const radiusPx = minCellSizePx * radiusRatio;
  
  // Font size cho số (tương tự như STT trong config)
  const fontSizePx =  250;
  const fontFamily = sttCfg.fontFamily || 'Arial';
  const fontWeight = sttCfg.fontWeight || 'bold';
  const textColor = sttCfg.color || '#000000';
  
  // Tạo SVG cho sheet
  let circlesSvg = '';
  let numbersSvg = '';
  let cutLinesSvg = '';
  
  // Tạo đường kẻ nét đứt (cut lines) để dễ cắt
  const dashArray = '10,10';
  const cutLineWidth = 2;
  const cutLineColor = '#CCCCCC';
  
  // Đường kẻ ngang giữa các hàng (4 đường cho 5 hàng)
  for (let i = 1; i < rows; i++) {
    const y = i * cellHeightPx;
    cutLinesSvg += `
  <line x1="0" y1="${y}" x2="${a4WidthPx}" y2="${y}" 
        stroke="${cutLineColor}" stroke-width="${cutLineWidth}" 
        stroke-dasharray="${dashArray}" />`;
  }
  
  // Đường kẻ dọc giữa các cột (2 đường cho 3 cột)
  for (let i = 1; i < cols; i++) {
    const x = i * cellWidthPx;
    cutLinesSvg += `
  <line x1="${x}" y1="0" x2="${x}" y2="${a4HeightPx}" 
        stroke="${cutLineColor}" stroke-width="${cutLineWidth}" 
        stroke-dasharray="${dashArray}" />`;
  }
  
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      const index = row * cols + col;
      if (index < totalCircles && index < numbers.length) {
        const number = numbers[index];
        const numberText = String(number).padStart(3, '0');
        
        // Vị trí center của cell
        const centerX = (col + 0.5) * cellWidthPx;
        const centerY = (row + 0.5) * cellHeightPx;
        
        // Vẽ vòng tròn
        circlesSvg += `
  <circle
    cx="${centerX}"
    cy="${centerY}"
    r="${radiusPx}"
    fill="none"
    stroke="${strokeColor}"
    stroke-width="${lineWidthPx}"
  />`;
        
        // Vẽ số
        numbersSvg += `
  <text
    x="${centerX}"
    y="${centerY}"
    fill="${textColor}"
    font-family="${fontFamily}"
    font-size="${fontSizePx}"
    font-weight="${fontWeight}"
    text-anchor="middle"
    dominant-baseline="middle"
  >${escapeXml(numberText)}</text>`;
      }
    }
  }
  
  // Tạo SVG đầy đủ
  const svg = `
<svg width="${a4WidthPx}" height="${a4HeightPx}" viewBox="0 0 ${a4WidthPx} ${a4HeightPx}" xmlns="http://www.w3.org/2000/svg">
  <!-- Nền trắng -->
  <rect x="0" y="0" width="${a4WidthPx}" height="${a4HeightPx}" fill="#FFFFFF" stroke="#808080" stroke-width="2"/>
  <!-- Đường kẻ nét đứt để cắt -->
${cutLinesSvg}
  <!-- Vòng tròn -->
${circlesSvg}
  <!-- Số -->
${numbersSvg}
</svg>
`.trim();
  
  // Tạo thư mục output nếu chưa có
  const outputPngDir = path.join(__dirname, 'output', 'png');
  const outputPdfDir = path.join(__dirname, 'output', 'pdf');
  if (!fs.existsSync(outputPngDir)) {
    fs.mkdirSync(outputPngDir, { recursive: true });
  }
  if (!fs.existsSync(outputPdfDir)) {
    fs.mkdirSync(outputPdfDir, { recursive: true });
  }
  
  // Tên file
  const fileNamePng = `sheet_circle_${String(sheetIndex + 1).padStart(3, '0')}.png`;
  const fileNamePdf = `sheet_circle_${String(sheetIndex + 1).padStart(3, '0')}.pdf`;
  const outPathPng = path.join(outputPngDir, fileNamePng);
  const outPathPdf = path.join(outputPdfDir, fileNamePdf);
  
  // Tạo PNG từ SVG
  await sharp(Buffer.from(svg))
    .resize(a4WidthPx, a4HeightPx, { fit: 'fill' })
    .withMetadata({ density: DPI })
    .png()
    .toFile(outPathPng);
  
  // Tạo PDF từ PNG buffer
  const pdfWidth = a4WidthMm * 2.83465;  // ~595.28 points
  const pdfHeight = a4HeightMm * 2.83465; // ~841.89 points
  
  const imageBuffer = await sharp(Buffer.from(svg))
    .resize(a4WidthPx, a4HeightPx, { fit: 'fill' })
    .png()
    .toBuffer();
  
  const doc = new PDFDocument({
    size: [pdfWidth, pdfHeight],
    margin: 0
  });
  
  const stream = fs.createWriteStream(outPathPdf);
  doc.pipe(stream);
  
  // Thêm image vào PDF
  doc.image(imageBuffer, {
    fit: [pdfWidth, pdfHeight],
    align: 'center',
    valign: 'center'
  });
  
  doc.end();
  
  // Đợi stream hoàn thành
  await new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });
  
  const firstNumber = numbers.length > 0 ? numbers[0] : 0;
  const lastNumber = numbers.length > 0 ? numbers[Math.min(numbers.length - 1, totalCircles - 1)] : 0;
  console.log(`Đã tạo sheet vòng tròn và số: ${fileNamePng} và ${fileNamePdf} (${numbers.length} số: ${firstNumber} - ${lastNumber})`);
}

// ===== Main function =====
async function run() {
  console.log('=== Tạo sheet vòng tròn và số (3 cột x 5 dòng) ===\n');
  
  // Đọc CSV để lấy danh sách STT hợp lệ
  const candidatePaths = [
    path.join(__dirname, '..', 'resources', 'DataMain30-01.csv'),
    path.join(__dirname, '..', 'resources', 'DataTest.csv'),
    path.join(__dirname, '..', 'input.csv')
  ];

  const inputCsvPath = candidatePaths.find((p) => fs.existsSync(p));

  if (!inputCsvPath) {
    console.error('Không tìm thấy file CSV input.');
    console.error('- Thử các đường dẫn sau nhưng đều không có:');
    candidatePaths.forEach((p) => console.error('  - ' + p));
    console.error('Hãy đặt file input CSV vào resources/DataMain30-01.csv hoặc input.csv');
    process.exit(1);
  }

  console.log(`Đang đọc file CSV: ${inputCsvPath}`);
  const csvContent = fs.readFileSync(inputCsvPath, 'utf8');

  const parsed = Papa.parse(csvContent, {
    header: true,
    skipEmptyLines: true
  });

  if (parsed.errors && parsed.errors.length > 0) {
    console.error('Lỗi khi parse CSV:', parsed.errors[0]);
  }

  const rows = parsed.data || [];
  console.log(`Đã đọc ${rows.length} dòng từ CSV.`);

  // Tìm các key thực tế từ dòng đầu tiên
  const firstRow = rows[0] || {};
  const allKeys = Object.keys(firstRow);
  const codeKey = allKeys.find(k => 
    k.toLowerCase() === 'code' || 
    k.toLowerCase().includes('mã') || 
    k.toLowerCase().includes('ma nv')
  ) || 'Code';

  // Lọc các STT có Code hợp lệ (không trống và không phải "Khách mời")
  const validSttList = [];
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sttRaw = row.STT || row.stt || row.Stt || '';
    const codeRaw = row[codeKey] || row.Code || row.CODE || row.code || '';
    
    const codeRawTrimmed = String(codeRaw || '').trim();
    
    // Kiểm tra Code: nếu trống hoặc là "Khách mời" thì bỏ qua
    if (!codeRawTrimmed || codeRawTrimmed.toLowerCase() === 'khách mời') {
      continue;
    }
    
    // Lấy STT và chuyển sang số
    const stt = String(sttRaw).trim();
    if (stt) {
      const sttNumber = parseInt(stt, 10);
      if (!Number.isNaN(sttNumber)) {
        validSttList.push(sttNumber);
      }
    }
  }

  console.log(`\nĐã lọc được ${validSttList.length} STT hợp lệ (có Code và không phải "Khách mời")`);
  
  if (validSttList.length === 0) {
    console.error('Không có STT nào hợp lệ để tạo sheet!');
    process.exit(1);
  }

  // Sắp xếp STT theo thứ tự tăng dần
  validSttList.sort((a, b) => a - b);
  
  // Chia thành các sheet (mỗi sheet 15 vòng tròn)
  const circlesPerSheet = 15;
  const numSheets = Math.ceil(validSttList.length / circlesPerSheet);
  
  console.log(`Sẽ tạo ${numSheets} sheet (mỗi sheet ${circlesPerSheet} vòng tròn)\n`);
  
  // Tạo các sheet
  for (let i = 0; i < numSheets; i++) {
    const startIdx = i * circlesPerSheet;
    const endIdx = Math.min(startIdx + circlesPerSheet, validSttList.length);
    const sheetNumbers = validSttList.slice(startIdx, endIdx);
    
    await createCircleNumberSheet(sheetNumbers, i);
  }
  
  console.log(`\n=== Hoàn thành ===`);
  console.log(`- Đã tạo ${numSheets} sheet PNG trong thư mục circle-sheet-generator/output/png/`);
  console.log(`- Đã tạo ${numSheets} sheet PDF trong thư mục circle-sheet-generator/output/pdf/`);
  console.log(`- Tổng số STT đã in: ${validSttList.length}`);
}

// Chạy script
run().catch((err) => {
  console.error('Lỗi khi generate sheet:', err);
  process.exit(1);
});

