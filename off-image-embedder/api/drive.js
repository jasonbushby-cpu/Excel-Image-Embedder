const { google } = require('googleapis');
const ExcelJS = require('exceljs');
const sharp = require('sharp');

const FOLDER_ID = '1B2js4ILgQkzYgPv9M65abw4M5-11k7_K';

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
}

async function findImageInDrive(drive, code) {
  const extensions = ['jpg', 'JPG', 'jpeg', 'JPEG', 'png', 'PNG'];
  for (const ext of extensions) {
    const response = await drive.files.list({
      q: `'${FOLDER_ID}' in parents and name = '${code}.${ext}' and trashed = false`,
      fields: 'files(id, name, mimeType)',
      pageSize: 1,
    });
    if (response.data.files && response.data.files.length > 0) {
      return response.data.files[0];
    }
  }
  return null;
}

async function getImageBuffer(drive, fileId) {
  const response = await drive.files.get(
    { fileId, alt: 'media' },
    { responseType: 'arraybuffer' }
  );
  return {
    buffer: Buffer.from(response.data),
    mimeType: response.headers['content-type'] || 'image/jpeg',
  };
}

const QUALITY_PRESETS = {
  email:        { size: 200, quality: 60 },
  standard:     { size: 300, quality: 72 },
  presentation: { size: 500, quality: 85 },
};

async function compressImage(buffer, preset) {
  const p = QUALITY_PRESETS[preset] || QUALITY_PRESETS.standard;
  return await sharp(buffer)
    .resize(p.size, p.size, { fit: 'inside', withoutEnlargement: true })
    .jpeg({ quality: p.quality })
    .toBuffer();
}

function columnHasData(rows, colIndex) {
  // Check data rows (skip header row 0) for any content in this column
  for (let i = 1; i < rows.length; i++) {
    const val = rows[i][colIndex];
    if (val !== null && val !== undefined && String(val).trim() !== '') {
      return true;
    }
  }
  return false;
}

function insertColumnIntoRows(rows, colIndex) {
  // Insert an empty column at colIndex, shifting everything right
  return rows.map(row => {
    const newRow = [...row];
    newRow.splice(colIndex, 0, '');
    return newRow;
  });
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { rows: inputRows, codeColIndex, imgColIndex, sheetName, imageMap, qualityPreset } = req.body;
    if (!inputRows || !inputRows.length) return res.status(400).json({ error: 'No rows provided' });

    const useDrive = !imageMap;
    let drive = null;
    if (useDrive) {
      const auth = getAuth();
      drive = google.drive({ version: 'v3', auth });
    }

    // Check if the chosen image column has data — if so, insert a blank column
    let rows = inputRows;
    let actualImgColIndex = imgColIndex;
    let columnInserted = false;

    if (columnHasData(rows, imgColIndex)) {
      rows = insertColumnIntoRows(rows, imgColIndex);
      columnInserted = true;
      // codeColIndex shifts right if it was at or after imgColIndex
      // (imgColIndex stays the same since we inserted there)
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(sheetName || 'Products');

    // Set image column width
    sheet.getColumn(actualImgColIndex + 1).width = 15;

    // Write all rows
    rows.forEach((row, rowIdx) => {
      const excelRow = sheet.getRow(rowIdx + 1);
      row.forEach((cellVal, colIdx) => {
        if (colIdx !== actualImgColIndex) {
          excelRow.getCell(colIdx + 1).value = cellVal;
        }
      });
      excelRow.commit();
    });

    let matched = 0;
    const missing = [];

    // Adjust codeColIndex if column was inserted before it
    const effectiveCodeCol = columnInserted && codeColIndex >= imgColIndex
      ? codeColIndex + 1
      : codeColIndex;

    for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) {
      const code = String(rows[rowIdx][effectiveCodeCol] || '').trim();
      // Skip empty cells and obvious header words
      if (!code) continue;
      const lc = code.toLowerCase();
      if (lc === 'code' || lc === 'product code' || lc === 'item no' || lc === 'item no.' || lc === 'sku') continue;

      try {
        let buffer, extension;

        if (useDrive) {
          const file = await findImageInDrive(drive, code);
          if (!file) { missing.push(code); continue; }
          const imgData = await getImageBuffer(drive, file.id);
          buffer = await compressImage(imgData.buffer, qualityPreset);
          extension = 'jpeg';
        } else {
          const b64 = imageMap[code.toUpperCase()];
          if (!b64) { missing.push(code); continue; }
          buffer = await compressImage(Buffer.from(b64, 'base64'), qualityPreset);
          extension = 'jpeg';
        }

        const imageId = workbook.addImage({ buffer, extension });
        sheet.getRow(rowIdx + 1).height = 60;
        sheet.addImage(imageId, {
          tl: { col: actualImgColIndex, row: rowIdx },
          br: { col: actualImgColIndex + 1, row: rowIdx + 1 },
          editAs: 'oneCell',
        });
        matched++;

      } catch (err) {
        console.error(`Error processing ${code}:`, err.message);
        missing.push(code);
      }
    }

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="output.xlsx"');
    res.setHeader('X-Matched', String(matched));
    res.setHeader('X-Missing', JSON.stringify(missing));
    res.setHeader('X-Column-Inserted', String(columnInserted));
    res.setHeader('Access-Control-Expose-Headers', 'X-Matched, X-Missing, X-Column-Inserted');

    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Handler error:', err.message);
    return res.status(500).json({ error: 'Server error', detail: err.message });
  }
};
