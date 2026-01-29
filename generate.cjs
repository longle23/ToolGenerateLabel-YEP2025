// generate.cjs
// Tool generate label PNG từ CSV dùng papaparse + sharp (render qua SVG)

const fs = require('fs');
const path = require('path');
const Papa = require('papaparse');
const sharp = require('sharp');

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

  // Sử dụng RobotoCondensed cho tất cả text
  const sttFontFamily = 'RobotoCondensed';
  const nameFontFamily = 'RobotoCondensed';
  const codeFontFamily = 'RobotoCondensed';
  const deptFontFamily = 'RobotoCondensed';

  const sttFontWeight = sttCfg.fontWeight || 'bold';
  const nameFontWeight = nameCfg.fontWeight || 'normal';
  const codeFontWeight = codeCfg.fontWeight || 'normal';
  const deptFontWeight = deptCfg.fontWeight || 'normal';

  // Tạo @font-face cho font RobotoCondensed
  const fontFaces = generateFontFace();

  const deptFontSize = deptCfg.fontSize || 20;
  const lineHeight = deptCfg.lineHeight || 1.2;

  // Wrap text cho Department - Company
  const maxWidthPx = widthPx * 0.9; // 90% chiều rộng để có margin
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
async function generateLabelForRow(row, index) {
  const sttRaw = row.STT || row.stt || row.Stt || '';
  const nameRaw = row.Name || row.NAME || row.name || '';
  const codeRaw = row.Code || row.CODE || row.code || '';
  const departmentRaw = row.Department || row.department || row.DEPARTMENT || '';
  const companyRaw = row.Company || row.company || row.COMPANY || '';

  const stt = String(sttRaw).trim();
  // Luôn viết hoa NAME và CODE
  const name = String(nameRaw).trim().toUpperCase();
  const code = String(codeRaw).trim().toUpperCase();
  
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
  const outPath = path.join(outputDir, fileName);

  // Raster SVG ở kích thước pixel cố định (1/4 A4),
  // chỉ set metadata density để DPI = 300 mà không phóng to ảnh.
  await sharp(Buffer.from(svg))
    .resize(widthPx, heightPx, { fit: 'fill' })
    .png()
    .withMetadata({ density: DPI })
    .toFile(outPath);

  console.log(`Đã tạo: ${fileName}`);
}

// ===== Đọc CSV bằng papaparse =====
async function run() {
  // Ưu tiên đọc từ resources/, sau đó fallback về file input.csv ở root
  const candidatePaths = [
    path.join(__dirname, 'resources', 'DataTest.csv'),
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

  for (let i = 0; i < rows.length; i++) {
    // eslint-disable-next-line no-await-in-loop
    await generateLabelForRow(rows[i], i);
  }
}

run().catch((err) => {
  console.error('Lỗi khi generate label:', err);
  process.exit(1);
});


