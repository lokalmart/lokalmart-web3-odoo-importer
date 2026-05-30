module.exports = async function handler(req, res) {
  const missing = [];
  ['ODOO_URL', 'ODOO_DB', 'ODOO_USERNAME', 'ODOO_API_KEY'].forEach((k) => { if (!process.env[k]) missing.push(k); });
  res.status(200).json({
    ok: missing.length === 0,
    missing,
    url: process.env.ODOO_URL || null,
    db: process.env.ODOO_DB || null,
    username: process.env.ODOO_USERNAME || null
  });
};
