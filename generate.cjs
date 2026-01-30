// generate.cjs
// Tool generate label PNG từ CSV dùng papaparse + sharp (render qua SVG)

const fs = require('fs');
const path = require('path');
const Papa = require('papaparse');
const sharp = require('sharp');
const PDFDocument = require('pdfkit');

// ===== Helper: convert mm -> px theo DPI =====
function mmToPx(mm, dpi) {
  // 1 inch = 25.4 mm
  return Math.round((mm / 25.4) * dpi);
}

// ===== Load config =====
const configPath = path.join(__dirname, 'config.json');
if (!fs.existsSync(configPath)) {
  console.error('Không tìm thấy config.json');
  process.exit(1);
}
const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
const { page } = config;
const DPI = page.dpi || 300;

// Kích thước canvas bằng 1/4 A4 (hoặc theo config)
const widthPx = mmToPx(page.width_mm, DPI);
const heightPx = mmToPx(page.height_mm, DPI);

// ===== Đảm bảo thư mục output tồn tại =====
const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// ===== Helper: build thuộc tính text-anchor theo align =====
function getTextAnchor(align) {
  const a = (align || 'center').toLowerCase();
  if (a === 'left') return 'start';
  if (a === 'right') return 'end';
  return 'middle';
}

// ===== Helper: tính toạ độ X theo config (mm hoặc "center") =====
function getXPosition(fieldConfig) {
  const xConf = fieldConfig.x ?? 'center';
  if (xConf === 'center') {
    return widthPx / 2;
  }
  if (typeof xConf === 'number') {
    return mmToPx(xConf, DPI);
  }
  if (typeof xConf === 'string') {
    const num = Number(xConf);
    if (!Number.isNaN(num)) {
      return mmToPx(num, DPI);
    }
  }
  return widthPx / 2;
}

// ===== Helper: load font RobotoCondensed-Regular.ttf và convert sang base64 =====
const fontCache = {};
function loadFontAsBase64() {
  if (fontCache['RobotoCondensed']) {
    return fontCache['RobotoCondensed'];
  }

  const fontPath = path.join(__dirname, 'fonts', 'static', 'RobotoCondensed-Regular.ttf');

  if (!fs.existsSync(fontPath)) {
    console.warn(`Không tìm thấy font file: ${fontPath}, sử dụng font mặc định`);
    return null;
  }

  try {
    const fontBuffer = fs.readFileSync(fontPath);
    const base64 = fontBuffer.toString('base64');
    fontCache['RobotoCondensed'] = base64;
    return base64;
  } catch (error) {
    console.warn(`Lỗi khi đọc font file ${fontPath}:`, error.message);
    return null;
  }
}

// ===== Helper: map fontWeight sang CSS font-weight number =====
function getFontWeightNumber(fontWeight) {
  const weight = String(fontWeight || 'normal').toLowerCase();
  const weightMap = {
    'thin': '100',
    'extralight': '200',
    'light': '300',
    'normal': '400',
    'regular': '400',
    'medium': '500',
    'semibold': '600',
    'bold': '700',
    'extrabold': '800',
    'black': '900'
  };
  return weightMap[weight] || '400';
}

// ===== Helper: tạo @font-face cho SVG =====
function generateFontFace() {
  const base64 = loadFontAsBase64();
  if (!base64) return '';

  return `@font-face {
    font-family: 'RobotoCondensed';
    src: url(data:font/truetype;charset=utf-8;base64,${base64}) format('truetype');
  }`;
}

// ===== Helper: wrap text thành nhiều dòng dựa trên chiều rộng =====
function wrapText(text, maxWidthPx, fontSize, fontFamily) {
  if (!text) return [];
  
  // Ước tính chiều rộng của một ký tự (approximate)
  // Với font sans-serif, thường khoảng 0.6 * fontSize cho ký tự trung bình
  const avgCharWidth = fontSize * 0.6;
  const maxCharsPerLine = Math.floor(maxWidthPx / avgCharWidth);
  
  if (text.length <= maxCharsPerLine) {
    return [text];
  }
  
  const words = text.split(/\s+/);
  const lines = [];
  let currentLine = '';
  
  for (const word of words) {
    const testLine = currentLine ? `${currentLine} ${word}` : word;
    
    if (testLine.length <= maxCharsPerLine) {
      currentLine = testLine;
    } else {
      if (currentLine) {
        lines.push(currentLine);
      }
      // Nếu từ đơn lẻ dài hơn maxCharsPerLine, cắt nó
      if (word.length > maxCharsPerLine) {
        let remaining = word;
        while (remaining.length > maxCharsPerLine) {
          lines.push(remaining.substring(0, maxCharsPerLine));
          remaining = remaining.substring(maxCharsPerLine);
        }
        currentLine = remaining;
      } else {
        currentLine = word;
      }
    }
  }
  
  if (currentLine) {
    lines.push(currentLine);
  }
  
  return lines;
}

// ===== Tạo SVG label cho 1 dòng =====
function createLabelSvg({ stt, name, code, departmentCompany }) {
  const sttCfg = config.stt;
  const nameCfg = config.name;
  const codeCfg = config.code;
  const deptCfg = config.departmentCompany || {};
  const circleCfg = config.circle || {};

  const sttX = getXPosition(sttCfg);
  const sttY = mmToPx(sttCfg.y || 0, DPI);
  const nameX = getXPosition(nameCfg);
  const nameY = mmToPx(nameCfg.y || 0, DPI);
  const codeX = getXPosition(codeCfg);
  const codeY = mmToPx(codeCfg.y || 0, DPI);
  const deptX = getXPosition(deptCfg);
  const deptY = mmToPx(deptCfg.y || 0, DPI);

  const sttFontSize = sttCfg.fontSize || 80;
  // Vẽ vòng tròn lớn chiếm phần lớn chiều ngang, giống sample:
  // bán kính = 35% cạnh nhỏ hơn (trừ một chút margin)
  const radius =
    circleCfg.radiusPx ||
    Math.min(widthPx, heightPx) * (circleCfg.radiusRatio || 0.35);

  const sttAnchor = getTextAnchor(sttCfg.align);
  const nameAnchor = getTextAnchor(nameCfg.align);
  const codeAnchor = getTextAnchor(codeCfg.align);
  const deptAnchor = getTextAnchor(deptCfg.align);

  const sttColor = sttCfg.color || '#000000';
  const nameColor = nameCfg.color || '#000000';
  const codeColor = codeCfg.color || '#000000';
  const deptColor = deptCfg.color || '#000000';

  const strokeColor = circleCfg.strokeColor || '#000000';
  const lineWidth = circleCfg.lineWidth || 4;

  // Sử dụng fontFamily từ config (mặc định Arial)
  const sttFontFamily = sttCfg.fontFamily || 'Arial';
  const nameFontFamily = nameCfg.fontFamily || 'Arial';
  const codeFontFamily = codeCfg.fontFamily || 'Arial';
  const deptFontFamily = deptCfg.fontFamily || 'Arial';

  const sttFontWeight = sttCfg.fontWeight || 'bold';
  const nameFontWeight = nameCfg.fontWeight || 'normal';
  const codeFontWeight = codeCfg.fontWeight || 'normal';
  const deptFontWeight = deptCfg.fontWeight || 'normal';

  // Arial là font hệ thống, không cần embed font-face
  const fontFaces = '';

  const deptFontSize = deptCfg.fontSize || 20;
  const lineHeight = deptCfg.lineHeight || 1.2;

  // Wrap text cho Department - Company
  // Tính padding left và right (chuyển từ mm sang px)
  const paddingLeftPx = deptCfg.paddingLeft ? mmToPx(deptCfg.paddingLeft, DPI) : 0;
  const paddingRightPx = deptCfg.paddingRight ? mmToPx(deptCfg.paddingRight, DPI) : 0;
  const maxWidthPx = widthPx * 0.9 - paddingLeftPx - paddingRightPx; // 90% chiều rộng trừ padding
  const deptLines = departmentCompany
    ? wrapText(departmentCompany, maxWidthPx, deptFontSize, deptFontFamily)
    : [];

  // Tạo SVG cho Department - Company với nhiều dòng
  let deptTextSvg = '';
  if (deptLines.length > 0) {
    const lines = deptLines.map((line, idx) => {
      const yOffset = (idx - (deptLines.length - 1) / 2) * deptFontSize * lineHeight;
      return `<tspan x="${deptX}" y="${deptY + yOffset}" text-anchor="${deptAnchor}">${escapeXml(line)}</tspan>`;
    }).join('\n    ');
    deptTextSvg = `  <!-- Department - Company -->
  <text
    x="${deptX}"
    y="${deptY}"
    fill="${deptColor}"
    font-family="${deptFontFamily}"
    font-size="${deptFontSize}"
    font-weight="${deptFontWeight}"
    text-anchor="${deptAnchor}"
    dominant-baseline="middle"
  >
    ${lines}
  </text>`;
  }

  // Embed font vào SVG
  const svg = `
<svg width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <style>
  ${fontFaces}
    </style>
  </defs>
  <!-- Nền trắng -->
  <rect x="0" y="0" width="${widthPx}" height="${heightPx}" fill="#FFFFFF" />

  <!-- Vòng tròn quanh STT -->
  <circle
    cx="${sttX}"
    cy="${sttY}"
    r="${radius}"
    fill="none"
    stroke="${strokeColor}"
    stroke-width="${lineWidth}"
  />

  <!-- STT -->
  <text
    x="${sttX}"
    y="${sttY}"
    fill="${sttColor}"
    font-family="${sttFontFamily}"
    font-size="${sttFontSize}"
    font-weight="${sttFontWeight}"
    text-anchor="${sttAnchor}"
    dominant-baseline="middle"
  >${escapeXml(stt)}</text>

  <!-- Name -->
  <text
    x="${nameX}"
    y="${nameY}"
    fill="${nameColor}"
    font-family="${nameFontFamily}"
    font-size="${nameCfg.fontSize || 28}"
    font-weight="${nameFontWeight}"
    text-anchor="${nameAnchor}"
    dominant-baseline="middle"
  >${escapeXml(name)}</text>

  <!-- Code -->
  <text
    x="${codeX}"
    y="${codeY}"
    fill="${codeColor}"
    font-family="${codeFontFamily}"
    font-size="${codeCfg.fontSize || 22}"
    font-weight="${codeFontWeight}"
    text-anchor="${codeAnchor}"
    dominant-baseline="middle"
  >${escapeXml(code || '')}</text>

${deptTextSvg}
</svg>
`.trim();

  return svg;
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

// ===== Generate label cho 1 dòng CSV =====
async function generateLabelForRow(row, index, codeKey, nameKey, companyKey, departmentKey) {
  const sttRaw = row.STT || row.stt || row.Stt || '';
  // Hỗ trợ cả header tiếng Anh và tiếng Việt
  const nameRaw = row[nameKey] || row.Name || row.NAME || row.name || '';
  const codeRaw = row[codeKey] || row.Code || row.CODE || row.code || '';
  const departmentRaw = row[departmentKey] || row.Department || row.department || row.DEPARTMENT || '';
  const companyRaw = row[companyKey] || row.Company || row.company || row.COMPANY || '';

  const stt = String(sttRaw).trim();
  // Luôn viết hoa NAME và CODE
  const name = String(nameRaw).trim().toUpperCase();
  
  // Lấy Code từ CSV gốc và trim
  const codeRawTrimmed = String(codeRaw || '').trim();
  
  // Kiểm tra Code trong CSV gốc: nếu trống hoặc là "Khách mời" thì bỏ qua
  if (!codeRawTrimmed || codeRawTrimmed.toLowerCase() === 'khách mời') {
    console.warn(`Bỏ qua dòng ${index + 1} vì Code trống hoặc là "Khách mời".`);
    return;
  }
  
  const code = codeRawTrimmed.toUpperCase();
  
  // Kết hợp Department - Company
  const department = String(departmentRaw).trim();
  const company = String(companyRaw).trim();
  const departmentCompany = department && company
    ? `${department} - ${company}`
    : department || company;

  if (!name) {
    console.warn(`Bỏ qua dòng ${index + 1} vì Name trống.`);
    return;
  }

  // hiển thị STT dạng 3 chữ số: 1 -> "001"
  const sttDisplay = stt ? String(stt).padStart(3, '0') : '';

  const svg = createLabelSvg({ stt: sttDisplay, name, code, departmentCompany });
  const sttNumber = parseInt(stt, 10);
  const seq = Number.isNaN(sttNumber) ? index + 1 : sttNumber;
  const fileName = `label_${String(seq).padStart(3, '0')}.png`;
  
  // Lưu vào folder output/one/
  const outputOneDir = path.join(__dirname, 'output', 'one');
  if (!fs.existsSync(outputOneDir)) {
    fs.mkdirSync(outputOneDir, { recursive: true });
  }
  const outPath = path.join(outputOneDir, fileName);

  // Raster SVG ở kích thước pixel cố định (1/4 A4),
  // chỉ set metadata density để DPI = 300 mà không phóng to ảnh.
  await sharp(Buffer.from(svg))
    .resize(widthPx, heightPx, { fit: 'fill' })
    .png()
    .withMetadata({ density: DPI })
    .toFile(outPath);

  console.log(`Đã tạo: ${fileName}`);
}

// ===== Tạo sheet A4 ngang với 8 labels (2 hàng x 4 cột) =====
async function createA4Sheet(labels, sheetIndex, withCutLines = true) {
  // Kích thước A4 NGANG: 297mm x 210mm
  const a4WidthMm = 297;
  const a4HeightMm = 210;
  const a4WidthPx = mmToPx(a4WidthMm, DPI);
  const a4HeightPx = mmToPx(a4HeightMm, DPI);
  
  // Layout: 2 hàng x 4 cột
  const cols = 4;
  const rows = 2;
  
  // Kích thước mỗi label trong sheet (chia đều A4)
  const labelWidthInSheetMm = a4WidthMm / cols;   // 297/4 = 74.25mm
  const labelHeightInSheetMm = a4HeightMm / rows;  // 210/2 = 105mm
  const labelWidthInSheetPx = mmToPx(labelWidthInSheetMm, DPI);
  const labelHeightInSheetPx = mmToPx(labelHeightInSheetMm, DPI);
  
  // Tạo canvas A4 trắng
  let compositeArray = [];
  
  // Tạo từng label và đặt vào đúng vị trí
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      const index = row * cols + col;
      if (index < labels.length) {
        const label = labels[index];
        const x = col * labelWidthInSheetPx;
        const y = row * labelHeightInSheetPx;
        
        // Tạo SVG cho label này với kích thước gốc
        const labelSvg = createLabelSvg(label);
        
        // Convert SVG sang buffer và scale xuống kích thước trong sheet
        const labelBuffer = await sharp(Buffer.from(labelSvg))
          .resize(labelWidthInSheetPx, labelHeightInSheetPx, { fit: 'fill' })
          .png()
          .toBuffer();
        
        // Thêm vào mảng composite
        compositeArray.push({
          input: labelBuffer,
          top: y,
          left: x
        });
      }
    }
  }
  
  // Tạo đường kẻ đứt nếu cần
  let cutLinesSvg = '';
  if (withCutLines) {
    const dashArray = '10,10';
    const strokeWidth = 2;
    const strokeColor = '#CCCCCC';
    
    // Đường kẻ ngang (1 đường giữa 2 hàng)
    for (let i = 1; i < rows; i++) {
      const y = i * labelHeightInSheetPx;
      cutLinesSvg += `
  <line x1="0" y1="${y}" x2="${a4WidthPx}" y2="${y}" 
        stroke="${strokeColor}" stroke-width="${strokeWidth}" 
        stroke-dasharray="${dashArray}" />`;
    }
    
    // Đường kẻ dọc (3 đường giữa 4 cột)
    for (let i = 1; i < cols; i++) {
      const x = i * labelWidthInSheetPx;
      cutLinesSvg += `
  <line x1="${x}" y1="0" x2="${x}" y2="${a4HeightPx}" 
        stroke="${strokeColor}" stroke-width="${strokeWidth}" 
        stroke-dasharray="${dashArray}" />`;
    }
  }
  
  // Tạo SVG cho đường kẻ đứt
  const cutLinesSvgFull = `
<svg width="${a4WidthPx}" height="${a4HeightPx}" viewBox="0 0 ${a4WidthPx} ${a4HeightPx}" xmlns="http://www.w3.org/2000/svg">
  ${cutLinesSvg}
</svg>`.trim();
  
  const cutLinesBuffer = await sharp(Buffer.from(cutLinesSvgFull))
    .resize(a4WidthPx, a4HeightPx, { fit: 'fill' })
    .png()
    .toBuffer();
  
  // Thêm đường kẻ đứt vào composite array
  if (withCutLines) {
    compositeArray.push({
      input: cutLinesBuffer,
      top: 0,
      left: 0
    });
  }
  
  // Tạo canvas trắng A4 và composite tất cả
  const outputSheetDir = path.join(__dirname, 'output', 'sheet');
  const outputPdfDir = path.join(__dirname, 'output', 'sheet-pdf');
  if (!fs.existsSync(outputSheetDir)) {
    fs.mkdirSync(outputSheetDir, { recursive: true });
  }
  if (!fs.existsSync(outputPdfDir)) {
    fs.mkdirSync(outputPdfDir, { recursive: true });
  }
  
  // Tạo cả PNG và PDF
  const fileNamePng = `sheet_A4_${String(sheetIndex + 1).padStart(3, '0')}.png`;
  const fileNamePdf = `sheet_A4_${String(sheetIndex + 1).padStart(3, '0')}.pdf`;
  const outPathPng = path.join(outputSheetDir, fileNamePng);
  const outPathPdf = path.join(outputPdfDir, fileNamePdf);
  
  // Tạo image buffer từ composite
  const imageBuffer = await sharp({
    create: {
      width: a4WidthPx,
      height: a4HeightPx,
      channels: 4,
      background: { r: 255, g: 255, b: 255, alpha: 1 }
    }
  })
    .composite(compositeArray)
    .withMetadata({ density: DPI })
    .png()
    .toBuffer();
  
  // Tạo PNG
  await sharp(imageBuffer)
    .toFile(outPathPng);
  
  // Tạo PDF từ PNG buffer sử dụng pdfkit
  // A4 landscape: 297mm x 210mm (convert sang points: 1mm = 2.83465 points)
  const pdfWidth = a4WidthMm * 2.83465;  // ~841.89 points
  const pdfHeight = a4HeightMm * 2.83465; // ~595.28 points
  
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
  
  console.log(`Đã tạo sheet A4: ${fileNamePng} và ${fileNamePdf}`);
}

// ===== Đọc CSV bằng papaparse =====
async function run() {
  // Ưu tiên đọc từ resources/, sau đó fallback về file input.csv ở root
  const candidatePaths = [
    path.join(__dirname, 'resources', 'DataMain30-01.csv'),
    // path.join(__dirname, 'resources', 'DataTest.csv'),
    path.join(__dirname, 'input.csv')
  ];

  const inputCsvPath = candidatePaths.find((p) => fs.existsSync(p));

  if (!inputCsvPath) {
    console.error('Không tìm thấy file CSV input.');
    console.error('- Thử các đường dẫn sau nhưng đều không có:');
    candidatePaths.forEach((p) => console.error('  - ' + p));
    console.error('Hãy đặt file input CSV vào resources/input.csv hoặc resources/RegistrationYEP2025.csv hoặc input.csv (giữ header STT,Code,Name).');
    process.exit(1);
  }

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

  // Tạo mảng chứa thông tin labels
  const labelData = [];

  // Generate từng label riêng lẻ
  console.log('\n=== Tạo từng phiếu riêng lẻ ===');
  
  // Tìm các key thực tế từ dòng đầu tiên (hỗ trợ cả tiếng Anh và tiếng Việt)
  const firstRow = rows[0] || {};
  const allKeys = Object.keys(firstRow);
  const codeKey = allKeys.find(k => 
    k.toLowerCase() === 'code' || 
    k.toLowerCase().includes('mã') || 
    k.toLowerCase().includes('ma nv')
  ) || 'Code';
  const nameKey = allKeys.find(k => 
    k.toLowerCase() === 'name' || 
    k.toLowerCase().includes('họ tên') || 
    k.toLowerCase().includes('ho ten')
  ) || 'Name';
  const companyKey = allKeys.find(k => 
    k.toLowerCase() === 'company' || 
    k.toLowerCase().includes('công ty') || 
    k.toLowerCase().includes('cong ty')
  ) || 'Company';
  const departmentKey = allKeys.find(k => 
    k.toLowerCase() === 'department' || 
    k.toLowerCase().includes('khối') || 
    k.toLowerCase().includes('phòng')
  ) || 'Department';
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sttRaw = row.STT || row.stt || row.Stt || '';
    const nameRaw = row[nameKey] || row.Name || row.NAME || row.name || '';
    const codeRaw = row[codeKey] || row.Code || row.CODE || row.code || '';
    const departmentRaw = row[departmentKey] || row.Department || row.department || row.DEPARTMENT || '';
    const companyRaw = row[companyKey] || row.Company || row.company || row.COMPANY || '';

    const stt = String(sttRaw).trim();
    const name = String(nameRaw).trim().toUpperCase();
    
    // Lấy Code từ CSV gốc và trim
    const codeRawTrimmed = String(codeRaw || '').trim();
    
    // Kiểm tra Code trong CSV gốc: nếu trống hoặc là "Khách mời" thì bỏ qua
    if (!codeRawTrimmed || codeRawTrimmed.toLowerCase() === 'khách mời') {
      console.warn(`Bỏ qua dòng ${i + 1} vì Code trống hoặc là "Khách mời" (Code: "${codeRawTrimmed}").`);
      continue;
    }
    
    const code = codeRawTrimmed.toUpperCase();
    
    const department = String(departmentRaw).trim();
    const company = String(companyRaw).trim();
    const departmentCompany = department && company
      ? `${department} - ${company}`
      : department || company;

    if (!name) {
      console.warn(`Bỏ qua dòng ${i + 1} vì Name trống.`);
      continue;
    }

    const sttDisplay = stt ? String(stt).padStart(3, '0') : '';
    
    // Lưu vào mảng để tạo sheet A4 sau
    labelData.push({ stt: sttDisplay, name, code, departmentCompany });
    
    // eslint-disable-next-line no-await-in-loop
    await generateLabelForRow(row, i, codeKey, nameKey, companyKey, departmentKey);
  }

  // Generate các sheet A4
  console.log('\n=== Tạo các sheet A4 (8 phiếu/sheet) ===');
  const sheetsCount = Math.ceil(labelData.length / 8);
  for (let i = 0; i < sheetsCount; i++) {
    const startIdx = i * 8;
    const endIdx = Math.min(startIdx + 8, labelData.length);
    const sheetLabels = labelData.slice(startIdx, endIdx);
    
    // eslint-disable-next-line no-await-in-loop
    await createA4Sheet(sheetLabels, i, true);
  }
  
  console.log(`\n=== Hoàn thành ===`);
  console.log(`- Đã tạo ${labelData.length} phiếu riêng lẻ trong thư mục output/one/`);
  console.log(`- Đã tạo ${sheetsCount} sheet A4 PNG trong thư mục output/sheet/`);
  console.log(`- Đã tạo ${sheetsCount} sheet A4 PDF trong thư mục output/sheet-pdf/`);
}

run().catch((err) => {
  console.error('Lỗi khi generate label:', err);
  process.exit(1);
});


