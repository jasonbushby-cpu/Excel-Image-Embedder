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
    console.error('Handler error:', err.message, err.stack);

    // Classify the error into a user-friendly message
    const msg = err.message || '';
    let userError = 'An unexpected error occurred. Please try again.';
    let hint = '';

    if (msg.includes('ETIMEDOUT') || msg.includes('ECONNRESET') || msg.includes('socket hang up')) {
      userError = 'The request timed out while fetching images from Google Drive.';
      hint = 'Try splitting your spreadsheet into smaller batches of 200–300 rows and running each separately.';
    } else if (msg.includes('rateLimitExceeded') || msg.includes('userRateLimitExceeded') || msg.includes('429')) {
      userError = 'Google Drive rate limit reached — too many image requests at once.';
      hint = 'Wait 60 seconds then try again, or split into smaller batches.';
    } else if (msg.includes('quota') || msg.includes('storageQuota')) {
      userError = 'Google Drive storage quota exceeded.';
      hint = 'Contact your Google Workspace admin to check Drive quota.';
    } else if (msg.includes('forbidden') || msg.includes('403') || msg.includes('accessNotConfigured')) {
      userError = 'Access denied to Google Drive. The service account may have lost permission.';
      hint = 'Check that the service account still has access to the image library folder.';
    } else if (msg.includes('invalid_grant') || msg.includes('unauthorized') || msg.includes('401')) {
      userError = 'Google authentication failed — the service account credentials may have expired.';
      hint = 'Contact your administrator to refresh the Google service account key.';
    } else if (msg.includes('ENOMEM') || msg.includes('out of memory') || msg.includes('heap')) {
      userError = 'The server ran out of memory processing this many images.';
      hint = `Your spreadsheet has too many rows to process in one go. Split into batches of 200–300 rows.`;
    } else if (msg.includes('maxContentLength') || msg.includes('request entity too large') || msg.includes('413')) {
      userError = 'The uploaded spreadsheet or image data is too large to process.';
      hint = 'Try compressing your images or splitting the spreadsheet into smaller files.';
    } else if (msg.includes('ENOTFOUND') || msg.includes('getaddrinfo')) {
      userError = 'Cannot reach Google Drive — network connectivity issue on the server.';
      hint = 'This is likely a temporary issue. Wait a minute and try again.';
    } else if (msg.includes('Function execution timed out') || msg.includes('execution timeout') || msg.includes('Task timed out')) {
      userError = 'The request timed out — 1,759 images is too many to process in one go.';
      hint = 'Vercel serverless functions have a 10-second limit. Split your spreadsheet into batches of 200–300 rows.';
    } else if (msg.includes('No rows provided') || msg.includes('400')) {
      userError = 'No spreadsheet data was received by the server.';
      hint = 'Try re-uploading your spreadsheet and running again.';
    } else if (msg.includes('workbook') || msg.includes('ExcelJS') || msg.includes('worksheet')) {
      userError = 'Failed to build the Excel file — the spreadsheet data may be malformed.';
      hint = 'Check your spreadsheet for merged cells or unusual formatting, then try again.';
    } else if (msg) {
      userError = msg;
    }

    return res.status(500).json({
      error: userError,
      hint,
      detail: msg,
    });
  }
};
