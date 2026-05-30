const fs = require('fs');
const XLSX = require('xlsx');
const { formidable } = require('formidable');
const { LokalmartImporter, ImportLog, toBool } = require('./_odoo');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }
  const log = new ImportLog();
  try {
    const form = formidable({ multiples: false, maxFileSize: 25 * 1024 * 1024 });
    const [fields, files] = await form.parse(req);
    const dryRunRaw = Array.isArray(fields.dryRun) ? fields.dryRun[0] : fields.dryRun;
    const dryRun = toBool(dryRunRaw, true);
    const uploaded = Array.isArray(files.file) ? files.file[0] : files.file;
    if (!uploaded) throw new Error('File XLSX belum diupload.');
    const buffer = fs.readFileSync(uploaded.filepath);
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: false });
    const importer = new LokalmartImporter(workbook, { dryRun });
    const result = await importer.run();
    res.status(200).json(result);
  } catch (e) {
    log.error('SYSTEM', e.message);
    res.status(400).json({ error: e.message, summary: log.summary(), logs: log.lines });
  }
};
