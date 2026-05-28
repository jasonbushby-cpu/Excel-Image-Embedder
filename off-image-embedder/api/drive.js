const { google } = require('googleapis');
const ExcelJS = require('exceljs');

const FOLDER_ID = '1B2js4ILgQkzYgPv9M65abw4M5-11k7_K';
const CONCURRENCY = 8; // parallel Drive fetches per batch

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

// Run tasks with limited concurrency
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
  if (!msg) return { userError: 'An unexpected error occurred. Please try again.', hint: '' };
  if (msg.includes('ETIMEDOUT') || msg.includes('ECONNRESET') || msg.includes('socket hang up'))
    return { userError: 'The request timed out while fetching images from Google Drive.', hint: 'This is usually temporary — please try again.' };
  if (msg.includes('rateLimitExceeded') || msg.includes('userRateLimitExceeded') || msg.includes('429'))
    return { userError: 'Google Drive rate limit reached.', hint: 'Wait 60 seconds then try again.' };
  if (msg.includes('forbidden') || msg.includes('403') || msg.includes('accessNotConfigured'))
    return { userError: 'Access denied to Google Drive. The service account may have lost permission.', hint: 'Check that the service account still has access to the image library folder.' };
  if (msg.includes('invalid_grant') || msg.includes('unauthorized') || msg.includes('401'))
    return { userError: 'Google authentication failed — credentials may have expired.', hint: 'Contact your administrator to refresh the Google service account key.' };
  if (msg.includes('ENOMEM') || msg.includes('out of memory') || msg.includes('heap'))
    return { userError: 'The server ran out of memory.', hint: 'Contact support if this keeps happening.' };
  if (msg.includes('Function execution timed out') || msg.includes('execution timeout') || msg.includes('Task timed out'))
    return { userError: 'Server timeout — this request took too long.', hint: 'Please try again. If it keeps failing, contact support.' };
  if (msg.includes('ENOTFOUND') || msg.includes('getaddrinfo'))
    return { userError: 'Cannot reach Google Drive — network issue on the server.', hint: 'Wait a minute and try again.' };
  return { userError: msg, hint: '' };
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { rows, codeColIndex, imgColIndex, sheetName, imageMap, prefetchCodes } = req.body;

    // ── MODE 1: Prefetch — return image data as JSON, no Excel built ──
    if (prefetchCodes && Array.isArray(prefetchCodes)) {
      const auth = getAuth();
      const drive = google.drive({ version: 'v3', auth });

      const results = await pLimit(
        prefetchCodes.map(code => async () => {
          try {
            const file = await findImageInDrive(drive, code);
            if (!file) return { code, found: false };
            const imgData = await getImageBuffer(drive, file.id);
            const b64 = imgData.buffer.toString('base64');
            return { code, found: true, b64, mimeType: imgData.mimeType };
          } catch (err) {
            console.error(`Prefetch error for ${code}:`, err.message);
            return { code, found: false };
          }
        }),
        CONCURRENCY
      );

      // Return as { images: { CODE: { b64, mimeType } } }
      const images = {};
      for (const r of results) {
        if (r.found) images[r.code.toUpperCase()] = { b64: r.b64, mimeType: r.mimeType };
      }
      return res.status(200).json({ images, fetched: Object.keys(images).length, total: prefetchCodes.length });
    }

    // ── MODE 2: Build Excel — imageMap supplied from prefetch or upload ──
    if (!rows || !rows.length) return res.status(400).json({ error: 'No rows provided' });

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet(sheetName || 'Products');
    sheet.getColumn(imgColIndex + 1).width = 15;

    // Write all rows
    rows.forEach((row, rowIdx) => {
      const excelRow = sheet.getRow(rowIdx + 1);
      row.forEach((cellVal, colIdx) => {
        if (colIdx !== imgColIndex) excelRow.getCell(colIdx + 1).value = cellVal;
      });
      excelRow.commit();
    });

    let matched = 0;
    const missing = [];

    for (let rowIdx = 1; rowIdx < rows.length; rowIdx++) {
      const code = String(rows[rowIdx][codeColIndex] || '').trim();
      if (!code || code.toLowerCase() === 'code') continue;

      try {
        const entry = imageMap ? (imageMap[code.toUpperCase()] || imageMap[code]) : null;
        if (!entry) { missing.push(code); continue; }

        // Entry is either { b64, mimeType } (from prefetch) or a plain base64 string (from upload)
        const b64      = typeof entry === 'string' ? entry : entry.b64;
        const mimeType = typeof entry === 'string' ? 'image/jpeg' : (entry.mimeType || 'image/jpeg');
        const buffer   = Buffer.from(b64, 'base64');
        const extension = mimeType.includes('png') ? 'png' : 'jpeg';

        const imageId = workbook.addImage({ buffer, extension });
        sheet.getRow(rowIdx + 1).height = 60;
        sheet.addImage(imageId, {
          tl: { col: imgColIndex, row: rowIdx },
          br: { col: imgColIndex + 1, row: rowIdx + 1 },
          editAs: 'oneCell',
        });
        matched++;
      } catch (err) {
        console.error(`Error embedding ${code}:`, err.message);
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
    console.error('Handler error:', err.message, err.stack);
    const { userError, hint } = classifyError(err.message);
    return res.status(500).json({ error: userError, hint, detail: err.message });
  }
};
