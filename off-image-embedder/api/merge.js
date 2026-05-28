const ExcelJS = require('exceljs');

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    // batches: [{ xlsx: base64, matched: number, missing: string[] }]
    const { batches, sheetName } = req.body;
    if (!batches || !batches.length) return res.status(400).json({ error: 'No batches provided' });

    const mergedWb = new ExcelJS.Workbook();
    const mergedSheet = mergedWb.addWorksheet(sheetName || 'Products');

    let totalMatched = 0;
    let allMissing = [];
    let destRowNum = 1; // current write position in merged sheet (1-indexed)

    for (let bIdx = 0; bIdx < batches.length; bIdx++) {
      const { xlsx: b64, matched, missing } = batches[bIdx];
      totalMatched += matched || 0;
      allMissing = allMissing.concat(missing || []);

      const batchBuf = Buffer.from(b64, 'base64');
      const batchWb = new ExcelJS.Workbook();
      await batchWb.xlsx.load(batchBuf);
      const batchSheet = batchWb.worksheets[0];
      if (!batchSheet) continue;

      // Copy column widths from first batch only
      if (bIdx === 0) {
        batchSheet.columns.forEach((col, i) => {
          if (col.width) mergedSheet.getColumn(i + 1).width = col.width;
        });
      }

      // Rows to copy: include header (row 1) only from first batch
      const srcStartRow = bIdx === 0 ? 1 : 2;
      // Track the offset so we know where this batch's rows land in merged sheet
      const destStartRow = destRowNum;

      batchSheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
        if (rowNum < srcStartRow) return;
        const destRow = mergedSheet.getRow(destRowNum);
        if (row.height) destRow.height = row.height;
        row.eachCell({ includeEmpty: true }, (cell, colNum) => {
          const destCell = destRow.getCell(colNum);
          destCell.value = cell.value;
          try { if (cell.style) destCell.style = JSON.parse(JSON.stringify(cell.style)); } catch(_) {}
        });
        destRow.commit();
        destRowNum++;
      });

      // Copy images — using the public getImages() API
      const images = batchSheet.getImages();
      for (const img of images) {
        try {
          // Get raw image data from source workbook's media array
          const srcMedia = batchWb.media[img.imageId];
          if (!srcMedia || !srcMedia.buffer) continue;

          const newImgId = mergedWb.addImage({
            buffer: srcMedia.buffer,
            extension: srcMedia.extension || 'jpeg',
          });

          // Shift row position: src row offset relative to srcStartRow, then place at destStartRow
          const srcRow = img.range.tl.nativeRow; // 0-indexed
          const srcRowOffset = srcRow - (srcStartRow - 1); // how many rows from the start of copied block
          const newRow = (destStartRow - 1) + srcRowOffset; // 0-indexed in merged sheet

          mergedSheet.addImage(newImgId, {
            tl: { col: img.range.tl.nativeCol, row: newRow },
            br: { col: img.range.br.nativeCol, row: newRow + 1 },
            editAs: img.range.editAs || 'oneCell',
          });
        } catch (imgErr) {
          console.error('Image copy error:', imgErr.message);
        }
      }
    }

    const outputBuf = await mergedWb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="output.xlsx"');
    res.setHeader('X-Matched', String(totalMatched));
    res.setHeader('X-Missing', JSON.stringify(allMissing));
    res.setHeader('Access-Control-Expose-Headers', 'X-Matched, X-Missing');
    return res.status(200).send(outputBuf);

  } catch (err) {
    console.error('Merge error:', err.message, err.stack);
    return res.status(500).json({ error: 'Failed to merge batches: ' + err.message });
  }
};
