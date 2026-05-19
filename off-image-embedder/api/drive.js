const { google } = require('googleapis');
const ExcelJS = require('exceljs');

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

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { rows, codeColIndex, imgColIndex, sheetName, imageMap } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: 'No rows provided' });

    const useDrive = !imageMap;
    let drive = null;

    if (useDrive) {
      const auth = getAuth();
      drive = google.drive({ version: 'v3', auth });
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(sheetName || 'Products');

    // Set image column width
    sheet.getColumn(imgColIndex + 1).width = 15;

    // Write all rows
    rows.forEach((row, rowIdx) => {
      const excelRow = sheet.getRow(rowIdx + 1);
      row.forEach((cellVal, colIdx) => {
        if (colIdx !== imgColIndex) {
          excelRow.getCell(colIdx + 1).value = cellVal;
        }
      });
      excelRow.commit();
    });

    let matched = 0;
    const missing = [];

    // Process each data row
    for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) {
      const code = String(rows[rowIdx][codeColIndex] || '').trim();
      if (!code || code.toLowerCase() === 'code') continue;

      try {
        let buffer, extension;

        if (useDrive) {
          const file = await findImageInDrive(drive, code);
          if (!file) { missing.push(code); continue; }
          const imgData = await getImageBuffer(drive, file.id);
          buffer = imgData.buffer;
          extension = imgData.mimeType.includes('png') ? 'png' : 'jpeg';
        } else {
          const b64 = imageMap[code.toUpperCase()];
          if (!b64) { missing.push(code); continue; }
          buffer = Buffer.from(b64, 'base64');
          extension = 'jpeg';
        }

        const imageId = workbook.addImage({ buffer, extension });
        sheet.getRow(rowIdx + 1).height = 60;
        sheet.addImage(imageId, {
          tl: { col: imgColIndex, row: rowIdx },
          br: { col: imgColIndex + 1, row: rowIdx + 1 },
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
    res.setHeader('Access-Control-Expose-Headers', 'X-Matched, X-Missing');

    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Handler error:', err.message);
    return res.status(500).json({ error: 'Server error', detail: err.message });
  }
};
