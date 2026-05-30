const { testConnection, ImportLog } = require('./_odoo');

module.exports = async function handler(req, res) {
  try {
    const result = await testConnection();
    res.status(200).json(result);
  } catch (e) {
    const log = new ImportLog();
    log.error('AUTH', e.message);
    res.status(400).json({ error: e.message, summary: log.summary(), logs: log.lines });
  }
};
