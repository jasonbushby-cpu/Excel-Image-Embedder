const { google } = require('googleapis');
const ExcelJS = require('exceljs');
const sharp = require('sharp');

const FOLDER_ID = '1B2js4ILgQkzYgPv9M65abw4M5-11k7_K';
const CONCURRENCY = 8;

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

// Compress image buffer using sharp
// targetMB: total file budget per image in KB (we'll do proportional sizing)
async function compressImage(buffer, quality) {
  try {
    return await sharp(buffer)
      .resize({ width: 400, height: 400, fit: 'inside', withoutEnlargement: true })
      .jpeg({ quality })
      .toBuffer();
  } catch (err) {
    console.error('Compression error:', err.message);
    return buffer; // return original if compression fails
  }
}

async function pLimit(tasks, concurrency) {
  const results = new Array(tasks.length);
  let index = 0;
  async function worker() {
    while (index < tasks.length) {
      const i = index++;
      results[i] = await tasks[i]();
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, tasks.length) }, worker));
  return results;
}

function classifyError(msg) {
  if (!msg) return { userError: 'An unexpected error occurred.', hint: '' };
  if (msg.includes('ETIMEDOUT') || msg.includes('ECONNRESET') || msg.includes('socket hang up'))
    return { userError: 'Request timed out fetching images from Google Drive.', hint: 'Please try again.' };
  if (msg.includes('rateLimitExceeded') || msg.includes('429'))
    return { userError: 'Google Drive rate limit reached.', hint: 'Wait 60 seconds then try again.' };
  if (msg.includes('forbidden') || msg.includes('403'))
    return { userError: 'Access denied to Google Drive.', hint: 'Check the service account still has folder access.' };
  if (msg.includes('invalid_grant') || msg.includes('401'))
    return { userError: 'Google authentication failed.', hint: 'Contact your administrator to refresh the service account key.' };
  if (msg.includes('ENOMEM') || msg.includes('out of memory'))
    return { userError: 'Server ran out of memory.', hint: 'Try a smaller batch.' };
  if (msg.includes('timed out') || msg.includes('timeout'))
    return { userError: 'Server timeout — batch took too long.', hint: 'Please try again.' };
  return { userError: msg, hint: '' };
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { rows, codeColIndex, imgColIndex, sheetName, imageMap, compressTarget } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: 'No rows provided' });

    // Compression quality settings
    // We resize to max 400px and apply JPEG quality to hit target file size
    const compressQuality = compressTarget === '5' ? 50 : compressTarget === '10' ? 70 : null;

    const useDrive = !imageMap;
    let drive = null;
    if (useDrive) {
      const auth = getAuth();
      drive = google.drive({ version: 'v3', auth });
    }

    const dataRows = [];
    for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) {
      const code = String(rows[rowIdx][codeColIndex] || '').trim();
      if (code && /\d/.test(code)) dataRows.push({ rowIdx, code });
    }

    const imageResults = await pLimit(
      dataRows.map(({ rowIdx, code }) => async () => {
        try {
          let buffer, extension;

          if (useDrive) {
            const file = await findImageInDrive(drive, code);
            if (!file) return { rowIdx, code, found: false };
            const imgData = await getImageBuffer(drive, file.id);
            buffer = imgData.buffer;
            extension = imgData.mimeType.includes('png') ? 'png' : 'jpeg';
          } else {
            const entry = imageMap[code.toUpperCase()] || imageMap[code];
            if (!entry) return { rowIdx, code, found: false };
            const b64 = typeof entry === 'string' ? entry : entry.b64;
            const mime = typeof entry === 'string' ? 'image/jpeg' : (entry.mimeType || 'image/jpeg');
            buffer = Buffer.from(b64, 'base64');
            extension = mime.includes('png') ? 'png' : 'jpeg';
          }

          // Apply compression if requested
          if (compressQuality) {
            buffer = await compressImage(buffer, compressQuality);
            extension = 'jpeg'; // sharp always outputs jpeg when compressing
          }

          return { rowIdx, code, found: true, buffer, extension };
        } catch (err) {
          console.error(`Error fetching ${code}:`, err.message);
          return { rowIdx, code, found: false };
        }
      }),
      CONCURRENCY
    );

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(sheetName || 'Products');
    sheet.getColumn(imgColIndex + 1).width = 15;

    rows.forEach((row, rowIdx) => {
      const excelRow = sheet.getRow(rowIdx + 1);
      row.forEach((cellVal, colIdx) => {
        if (colIdx !== imgColIndex) excelRow.getCell(colIdx + 1).value = cellVal;
      });
      excelRow.commit();
    });

    let matched = 0;
    const missing = [];

    for (const result of imageResults) {
      if (!result.found) { missing.push(result.code); continue; }
      try {
        const imageId = workbook.addImage({ buffer: result.buffer, extension: result.extension });
        sheet.getRow(result.rowIdx + 1).height = 60;
        sheet.addImage(imageId, {
          tl: { col: imgColIndex, row: result.rowIdx },
          br: { col: imgColIndex + 1, row: result.rowIdx + 1 },
          editAs: 'oneCell',
        });
        matched++;
      } catch (err) {
        console.error(`Error embedding ${result.code}:`, err.message);
        missing.push(result.code);
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
    console.error('Handler error:', err.message, err.stack);
    const { userError, hint } = classifyError(err.message || '');
    return res.status(500).json({ error: userError, hint, detail: err.message });
  }
};
