const { google } = require('googleapis');

const FOLDER_ID = '1B2js4ILgQkzYgPv9M65abw4M5-11k7_K';

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/drive.readonly'],
  });
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') return res.status(200).end();

  const { code, action } = req.query;

  if (!code) return res.status(400).json({ error: 'No product code provided' });

  try {
    const auth = getAuth();
    const drive = google.drive({ version: 'v3', auth });

    if (action === 'list') {
      // List all files in folder for bulk matching
      const response = await drive.files.list({
        q: `'${FOLDER_ID}' in parents and trashed = false`,
        fields: 'files(id, name)',
        pageSize: 1000,
      });
      return res.status(200).json({ files: response.data.files });
    }

    // Find exact match for product code (no underscore variants)
    const extensions = ['jpg', 'JPG', 'jpeg', 'JPEG', 'png', 'PNG'];
    let fileId = null;

    for (const ext of extensions) {
      const response = await drive.files.list({
        q: `'${FOLDER_ID}' in parents and name = '${code}.${ext}' and trashed = false`,
        fields: 'files(id, name)',
        pageSize: 1,
      });
      if (response.data.files && response.data.files.length > 0) {
        fileId = response.data.files[0].id;
        break;
      }
    }

    if (!fileId) {
      return res.status(404).json({ error: 'Image not found', code });
    }

    // Stream the image back
    const imageResponse = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );

    const buffer = Buffer.from(imageResponse.data);
    res.setHeader('Content-Type', imageResponse.headers['content-type'] || 'image/jpeg');
    res.setHeader('Cache-Control', 'public, max-age=3600');
    return res.status(200).send(buffer);

  } catch (err) {
    console.error('Drive API error:', err.message);
    return res.status(500).json({ error: 'Drive API error', detail: err.message });
  }
};
