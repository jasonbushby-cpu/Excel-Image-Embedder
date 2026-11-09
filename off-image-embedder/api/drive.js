const { google } = require('googleapis');
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
  // One query covering every extension (was up to six separate queries per code)
  const safe = code.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const names = ['jpg', 'JPG', 'jpeg', 'JPEG', 'png', 'PNG']
    .map(ext => `name = '${safe}.${ext}'`).join(' or ');
  const response = await drive.files.list({
    q: `'${FOLDER_ID}' in parents and (${names}) and trashed = false`,
    fields: 'files(id, name, mimeType)',
    pageSize: 10,
  });
  const files = response.data.files || [];
  // Same preference order as before: jpg, then jpeg, then png
  const rank = f => ['jpg', 'jpeg', 'png'].indexOf(f.name.split('.').pop().toLowerCase());
  files.sort((a, b) => rank(a) - rank(b));
  return files[0] || null;
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

// Every image is resized before it's returned. Cells only show images at ~80px,
// and keeping them small is what keeps each response under Vercel's 4.5 MB limit.
const COMPRESS = {
  none: { size: 400, quality: 82 },
  '10': { size: 300, quality: 70 },
  '5':  { size: 200, quality: 65 },
};

async function compressImage(buffer, target) {
  const { size, quality } = COMPRESS[target] || COMPRESS.none;
  return sharp(buffer)
    .rotate()
    .resize({ width: size, height: size, fit: 'inside', withoutEnlargement: true })
    .flatten({ background: '#ffffff' })
    .jpeg({ quality, mozjpeg: true })
    .toBuffer();
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
    return { userError: 'Server ran out of memory.', hint: 'Try using the 5MB or 10MB compression option.' };
  if (msg.includes('timed out') || msg.includes('timeout'))
    return { userError: 'Server timeout — batch took too long.', hint: 'Please try again.' };
  return { userError: msg, hint: '' };
}

// Max base64 bytes per response — stays safely under Vercel's 4.5 MB response cap.
// Anything over is returned as "deferred" and the page simply asks for it again.
const MAX_RESPONSE_BYTES = 3.5 * 1024 * 1024;

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { codes, compressTarget } = req.body || {};
    if (!Array.isArray(codes) || !codes.length) return res.status(400).json({ error: 'No product codes provided' });
    if (codes.length > 100) return res.status(400).json({ error: 'Too many codes in one request (max 100)' });

    const drive = google.drive({ version: 'v3', auth: getAuth() });

    const results = await pLimit(
      codes.map(raw => async () => {
        const code = String(raw).trim();
        try {
          const file = await findImageInDrive(drive, code);
          if (!file) return { code, found: false };
          const { buffer } = await getImageBuffer(drive, file.id);
          const small = await compressImage(buffer, compressTarget);
          return { code, found: true, b64: small.toString('base64') };
        } catch (err) {
          console.error(`Error fetching ${code}:`, err.message);
          // Rate limits / timeouts: let the page retry this code rather than marking it missing
          const msg = err.message || '';
          const transient = /rateLimit|429|ETIMEDOUT|ECONNRESET|socket hang up|timeout/i.test(msg);
          return { code, found: false, retry: transient };
        }
      }),
      CONCURRENCY
    );

    const images = {};
    const missing = [];
    const deferred = [];
    let total = 0;
    for (const r of results) {
      if (r.retry) { deferred.push(r.code); continue; }
      if (!r.found) { missing.push(r.code); continue; }
      if (total + r.b64.length > MAX_RESPONSE_BYTES && Object.keys(images).length) { deferred.push(r.code); continue; }
      images[r.code] = r.b64;
      total += r.b64.length;
    }

    return res.status(200).json({ images, missing, deferred });
  } catch (err) {
    console.error('Handler error:', err.message, err.stack);
    const { userError, hint } = classifyError(err.message || '');
    return res.status(500).json({ error: userError, hint, detail: err.message });
  }
};
