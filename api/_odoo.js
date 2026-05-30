const XLSX = require('xlsx');

function isBlank(v) { return v === undefined || v === null || String(v).trim() === ''; }
function toBool(v, fallback = false) {
  if (v === undefined || v === null || v === '') return fallback;
  if (typeof v === 'boolean') return v;
  const s = String(v).trim().toLowerCase();
  return ['true', '1', 'yes', 'y', 'ya', 'iya'].includes(s);
}
function toNumberOrNull(v) { if (isBlank(v)) return null; const n = Number(String(v).replace(/,/g, '.')); return Number.isFinite(n) ? n : null; }
function compactObject(obj) { const out = {}; Object.entries(obj).forEach(([k, v]) => { if (v !== undefined && v !== null && v !== '') out[k] = v; }); return out; }
function parseListCell(value) { if (isBlank(value)) return []; return String(value).split(',').map(x => x.trim()).filter(Boolean); }
function splitExternalId(xmlid) { if (!xmlid || !String(xmlid).includes('.')) throw new Error(`External ID tidak valid: ${xmlid}`); const [module, ...rest] = String(xmlid).split('.'); return { module, name: rest.join('.') }; }
function normalizeHeader(h) { return String(h || '').trim(); }
function sheetRows(workbook, sheetName) {
  if (!workbook.Sheets[sheetName]) return [];
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '', raw: false }).map((r) => {
    const out = {}; Object.keys(r).forEach(k => { out[normalizeHeader(k)] = r[k]; }); return out;
  });
}

class ImportLog {
  constructor() { this.lines = []; }
  push(level, sheet, message, meta = {}) { this.lines.push({ time: new Date().toISOString(), level, sheet, message, meta }); }
  info(s, m, meta) { this.push('info', s, m, meta); }
  ok(s, m, meta) { this.push('ok', s, m, meta); }
  warn(s, m, meta) { this.push('warn', s, m, meta); }
  error(s, m, meta) { this.push('error', s, m, meta); }
  summary() { return this.lines.reduce((a, l) => { a[l.level] = (a[l.level] || 0) + 1; return a; }, {}); }
}

class OdooClient {
  constructor(config, log) { this.config = config; this.uid = null; this.rpcId = 1; this.log = log; }
  ensureConfig() {
    const missing = [];
    ['ODOO_URL', 'ODOO_DB', 'ODOO_USERNAME', 'ODOO_API_KEY'].forEach(k => { if (!process.env[k]) missing.push(k); });
    if (missing.length) throw new Error(`Environment belum lengkap: ${missing.join(', ')}`);
  }
  get url() { return String(process.env.ODOO_URL || '').replace(/\/$/, ''); }
  get db() { return process.env.ODOO_DB || ''; }
  get username() { return process.env.ODOO_USERNAME || ''; }
  get apiKey() { return process.env.ODOO_API_KEY || ''; }
  async jsonRpc(service, method, args) {
    this.ensureConfig();
    const payload = { jsonrpc: '2.0', method: 'call', params: { service, method, args }, id: this.rpcId++ };
    const res = await fetch(`${this.url}/jsonrpc`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const text = await res.text();
    let data; try { data = JSON.parse(text); } catch (_) { throw new Error(`Odoo response bukan JSON: HTTP ${res.status} ${text.slice(0, 300)}`); }
    if (!res.ok) throw new Error(`Odoo HTTP ${res.status}: ${text.slice(0, 500)}`);
    if (data.error) { const msg = data.error?.data?.message || data.error?.message || JSON.stringify(data.error); throw new Error(`Odoo RPC error: ${msg}`); }
    return data.result;
  }
  async authenticate() {
    if (this.uid) return this.uid;
    this.log.info('AUTH', `Login ke Odoo ${this.url} db=${this.db} user=${this.username}`);
    const uid = await this.jsonRpc('common', 'login', [this.db, this.username, this.apiKey]);
    if (!uid) throw new Error('Login Odoo gagal. Periksa ODOO_DB, ODOO_USERNAME, dan ODOO_API_KEY.');
    this.uid = uid;
    this.log.ok('AUTH', `Login sukses uid=${uid}`);
    return uid;
  }
  async executeKw(model, method, args = [], kwargs = {}) { await this.authenticate(); return this.jsonRpc('object', 'execute_kw', [this.db, this.uid, this.apiKey, model, method, args, kwargs]); }
  async searchRead(model, domain = [], fields = ['id'], limit = 0, order = '') { const kwargs = { fields }; if (limit) kwargs.limit = limit; if (order) kwargs.order = order; return this.executeKw(model, 'search_read', [domain], kwargs); }
  async create(model, values) { return this.executeKw(model, 'create', [values]); }
  async write(model, ids, values) { return this.executeKw(model, 'write', [Array.isArray(ids) ? ids : [ids], values]); }
}

class LokalmartImporter {
  constructor(workbook, { dryRun = true } = {}) {
    this.workbook = workbook; this.dryRun = dryRun; this.log = new ImportLog(); this.odoo = new OdooClient({}, this.log);
    this.cache = { model: new Map(), xmlid: new Map() };
  }
  async run() {
    if (!this.workbook?.SheetNames?.length) throw new Error('Workbook kosong/tidak valid.');
    this.log.info('SYSTEM', `Mode: ${this.dryRun ? 'DRY RUN' : 'IMPORT NOW'}`);
    await this.odoo.authenticate();
    await this.processModelsCheck();
    await this.processFields();
    await this.processSelections();
    await this.processPartners();
    await this.processProducts();
    await this.processStockLots();
    await this.processProjects();
    await this.processProjectStages();
    await this.processProjectTags();
    await this.processMilestones();
    await this.processTasks();
    await this.processWebsitePages();
    await this.processQrRegistry();
    this.log.ok('SYSTEM', 'Selesai. Baca log detail sebelum mengambil langkah berikutnya.');
    return { summary: this.log.summary(), logs: this.log.lines };
  }
  async getModelId(model) {
    if (this.cache.model.has(model)) return this.cache.model.get(model);
    const rows = await this.odoo.searchRead('ir.model', [['model', '=', model]], ['id', 'model', 'name'], 1);
    if (!rows.length) throw new Error(`Model tidak tersedia di Odoo: ${model}`);
    this.cache.model.set(model, rows[0].id); return rows[0].id;
  }
  async modelExists(model) { try { await this.getModelId(model); return true; } catch (_) { return false; } }
  async resolveXmlId(xmlid, expectedModel = null) {
    if (isBlank(xmlid)) return null;
    const key = expectedModel ? `${xmlid}|${expectedModel}` : String(xmlid); if (this.cache.xmlid.has(key)) return this.cache.xmlid.get(key);
    const { module, name } = splitExternalId(xmlid); const domain = [['module', '=', module], ['name', '=', name]]; if (expectedModel) domain.push(['model', '=', expectedModel]);
    const rows = await this.odoo.searchRead('ir.model.data', domain, ['id', 'model', 'res_id'], 1);
    if (!rows.length) return null; this.cache.xmlid.set(key, rows[0].res_id); return rows[0].res_id;
  }
  async ensureXmlId(xmlid, model, resId, sheet) {
    if (isBlank(xmlid) || !resId) return;
    const { module, name } = splitExternalId(xmlid); const rows = await this.odoo.searchRead('ir.model.data', [['module', '=', module], ['name', '=', name]], ['id', 'model', 'res_id'], 1);
    if (rows.length) return;
    if (this.dryRun) { this.log.info(sheet, `[dry-run] create XMLID ${xmlid} → ${model}:${resId}`); return; }
    await this.odoo.create('ir.model.data', { module, name, model, res_id: resId, noupdate: true }); this.log.ok(sheet, `External ID dibuat: ${xmlid}`);
  }
  async findByXmlIdOrDomain(xmlid, model, domain, fields = ['id']) {
    if (!isBlank(xmlid)) { const id = await this.resolveXmlId(xmlid, model); if (id) return { id, source: 'xmlid' }; }
    const rows = await this.odoo.searchRead(model, domain, fields, 1); return rows.length ? { id: rows[0].id, source: 'domain', row: rows[0] } : null;
  }
  async writeSafe(model, id, values, sheet, msg) {
    const cleaned = compactObject(values); if (!Object.keys(cleaned).length) return;
    try { await this.odoo.write(model, id, cleaned); this.log.ok(sheet, msg, { id }); }
    catch (e) { this.log.warn(sheet, `Write gagal untuk ${model}:${id}. Mungkin field belum ada/readonly.`, { error: e.message }); }
  }
  async m2o(xmlid, model, sheet, required = false) {
    if (isBlank(xmlid)) return null; const id = await this.resolveXmlId(xmlid, model);
    if (!id && required) throw new Error(`${sheet}: Relasi wajib tidak ditemukan: ${xmlid} (${model})`);
    if (!id) this.log.warn(sheet, `Relasi tidak ditemukan dan dikosongkan: ${xmlid}`); return id || null;
  }
  async m2m(cell, model, sheet) {
    const ids = []; for (const xmlid of parseListCell(cell)) { const id = await this.resolveXmlId(xmlid, model); if (id) ids.push(id); else this.log.warn(sheet, `XMLID many2many tidak ditemukan: ${xmlid}`); }
    return ids.length ? [[6, 0, ids]] : undefined;
  }
  async processModelsCheck() {
    const sheet = '01_MODELS_CHECK'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    for (const [i, row] of rows.entries()) { const model = row.model || row.technical_model || row.name; if (isBlank(model)) continue; const required = toBool(row.required, true); const exists = await this.modelExists(model); if (exists) this.log.ok(sheet, `Model tersedia: ${model}`); else if (required) throw new Error(`Model wajib tidak tersedia: ${model} row ${i + 2}`); else this.log.warn(sheet, `Model opsional tidak tersedia: ${model}`); }
  }
  async processFields() {
    const sheet = '02_FIELDS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    for (const [i, row] of rows.entries()) {
      const rowNum = i + 2; const xmlid = row.external_id || row.id; const model = row.model || row['model_id/model']; const name = row.name; const ttype = row.ttype || row.field_type;
      if (isBlank(model) || isBlank(name) || isBlank(ttype)) { this.log.warn(sheet, `Row ${rowNum} dilewati: model/name/ttype kosong.`); continue; }
      if (!String(name).startsWith('x_')) throw new Error(`Row ${rowNum}: custom field wajib diawali x_: ${name}`);
      const modelId = await this.getModelId(model); const existing = await this.odoo.searchRead('ir.model.fields', [['model', '=', model], ['name', '=', name]], ['id', 'ttype'], 1);
      if (existing.length) { if (existing[0].ttype !== ttype) throw new Error(`Field ${model}.${name} sudah ada tapi tipe ${existing[0].ttype}, bukan ${ttype}`); this.log.ok(sheet, `Field sudah ada: ${model}.${name}`); await this.ensureXmlId(xmlid, 'ir.model.fields', existing[0].id, sheet); continue; }
      const values = compactObject({ name, field_description: row.field_description || row.field_label || name, model_id: modelId, ttype, relation: row.relation, state: 'manual', store: toBool(row.store, true), required: toBool(row.required, false), readonly: toBool(row.readonly, false), index: toBool(row.index, false), copied: toBool(row.copied, true), help: row.help || row.notes });
      if (this.dryRun) { this.log.info(sheet, `[dry-run] create field ${model}.${name}`, values); continue; }
      const id = await this.odoo.create('ir.model.fields', values); this.log.ok(sheet, `Field dibuat: ${model}.${name}`, { id }); await this.ensureXmlId(xmlid, 'ir.model.fields', id, sheet);
    }
  }
  async getFieldId(fieldXmlid, model, fieldName) {
    if (!isBlank(fieldXmlid)) { const id = await this.resolveXmlId(fieldXmlid, 'ir.model.fields'); if (id) return id; }
    if (isBlank(model) || isBlank(fieldName)) throw new Error(`Field reference tidak lengkap: ${fieldXmlid || ''}`);
    const rows = await this.odoo.searchRead('ir.model.fields', [['model', '=', model], ['name', '=', fieldName]], ['id'], 1); if (!rows.length) throw new Error(`Field tidak ditemukan: ${model}.${fieldName}`); return rows[0].id;
  }
  async processSelections() {
    const sheet = '03_SELECTIONS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('ir.model.fields.selection'))) { this.log.warn(sheet, 'ir.model.fields.selection tidak tersedia.'); return; }
    for (const [i, row] of rows.entries()) {
      const rowNum = i + 2; const fieldXmlid = row.field_external_id || row['field_id/id']; const value = row.value; const label = row.label || row.name;
      if (isBlank(value) || isBlank(label)) { this.log.warn(sheet, `Row ${rowNum} dilewati: value/label kosong.`); continue; }
      let fieldId; try { fieldId = await this.getFieldId(fieldXmlid, row.model, row.field_name); } catch (e) { if (this.dryRun) { this.log.info(sheet, `[dry-run] selection ${value} menunggu field dibuat: ${fieldXmlid || row.field_name}`); continue; } throw e; }
      const exists = await this.odoo.searchRead('ir.model.fields.selection', [['field_id', '=', fieldId], ['value', '=', value]], ['id'], 1); if (exists.length) { this.log.ok(sheet, `Selection sudah ada: ${value}`); continue; }
      const values = { field_id: fieldId, value, name: label, sequence: toNumberOrNull(row.sequence) || 10 };
      if (this.dryRun) { this.log.info(sheet, `[dry-run] create selection ${value}`, values); continue; }
      const id = await this.odoo.create('ir.model.fields.selection', values); this.log.ok(sheet, `Selection dibuat: ${value}`, { id });
    }
  }
  async processPartners() {
    const sheet = '04_PARTNERS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    for (const [i, row] of rows.entries()) {
      const rowNum = i + 2; const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) { this.log.warn(sheet, `Row ${rowNum} dilewati: name kosong.`); continue; }
      const lokalId = row.x_lokal_id || row.lokal_id; const domain = !isBlank(lokalId) ? [['x_lokal_id', '=', lokalId]] : [['name', '=', name]]; const found = await this.findByXmlIdOrDomain(xmlid, 'res.partner', domain, ['id', 'name']);
      const values = compactObject({ name, phone: row.phone, mobile: row.mobile || row.whatsapp, email: row.email, street: row.street, city: row.city, zip: row.zip, x_lokal_id: lokalId, x_lokal_role: row.x_lokal_role || row.role, x_lokal_member_type: row.x_lokal_member_type, x_lokal_points: toNumberOrNull(row.x_lokal_points), x_lokal_area: row.x_lokal_area || row.area, x_lokal_verification_status: row.x_lokal_verification_status });
      if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update partner ${name}`); else await this.writeSafe('res.partner', found.id, values, sheet, `Partner diupdate: ${name}`); await this.ensureXmlId(xmlid, 'res.partner', found.id, sheet); }
      else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create partner ${name}`); continue; } const id = await this.odoo.create('res.partner', values); this.log.ok(sheet, `Partner dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'res.partner', id, sheet); }
    }
  }
  async processProducts() {
    const sheet = '05_PRODUCTS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    for (const [i, row] of rows.entries()) {
      const rowNum = i + 2; const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) { this.log.warn(sheet, `Row ${rowNum} dilewati: name kosong.`); continue; }
      const lokalId = row.x_lokal_id || row.x_lokal_product_id || row.lokal_product_id; const domain = !isBlank(lokalId) ? [['x_lokal_id', '=', lokalId]] : (!isBlank(row.default_code) ? [['default_code', '=', row.default_code]] : [['name', '=', name]]);
      const vendorId = await this.m2o(row['x_lokal_vendor_partner_id/id'] || row.vendor_external_id, 'res.partner', sheet, false);
      const values = compactObject({ name, default_code: row.default_code, barcode: row.barcode, list_price: toNumberOrNull(row.list_price || row.price), standard_price: toNumberOrNull(row.standard_price || row.cost), sale_ok: toBool(row.sale_ok, true), purchase_ok: toBool(row.purchase_ok, true), detailed_type: row.detailed_type || row.type || 'product', x_lokal_id: lokalId, x_lokal_tracking_level: row.x_lokal_tracking_level || row.tracking_level, x_lokal_passport_url: row.x_lokal_passport_url, x_lokal_vendor_partner_id: vendorId, x_lokal_origin_city: row.x_lokal_origin_city, x_lokal_origin_district: row.x_lokal_origin_district, x_lokal_verification_status: row.x_lokal_verification_status, x_lokal_story: row.x_lokal_story, x_lokal_proof_hash: row.x_lokal_proof_hash, x_lokal_public_visible: toBool(row.x_lokal_public_visible, true) });
      const found = await this.findByXmlIdOrDomain(xmlid, 'product.template', domain, ['id', 'name']);
      if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update product ${name}`); else await this.writeSafe('product.template', found.id, values, sheet, `Produk diupdate: ${name}`); await this.ensureXmlId(xmlid, 'product.template', found.id, sheet); }
      else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create product ${name}`); continue; } const id = await this.odoo.create('product.template', values); this.log.ok(sheet, `Produk dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'product.template', id, sheet); }
    }
  }
  async productVariantId(templateId) { const rows = await this.odoo.searchRead('product.product', [['product_tmpl_id', '=', templateId]], ['id'], 1); if (!rows.length) throw new Error(`product.product tidak ditemukan untuk template ${templateId}`); return rows[0].id; }
  async processStockLots() {
    const sheet = '06_STOCK_LOTS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('stock.lot'))) { this.log.warn(sheet, 'Inventory/stock.lot tidak tersedia.'); return; }
    for (const [i, row] of rows.entries()) {
      const rowNum = i + 2; const xmlid = row.external_id || row.id; const name = row.name || row.lot_name || row.x_lokal_batch_id || row.x_lokal_unit_id; if (isBlank(name)) { this.log.warn(sheet, `Row ${rowNum} dilewati: name kosong.`); continue; }
      const tmplId = await this.m2o(row['product_tmpl_id/id'] || row.product_template_external_id, 'product.template', sheet, true); const productId = await this.productVariantId(tmplId);
      const found = await this.findByXmlIdOrDomain(xmlid, 'stock.lot', [['name', '=', name], ['product_id', '=', productId]], ['id', 'name']);
      const values = compactObject({ name, product_id: productId, x_lokal_batch_id: row.x_lokal_batch_id, x_lokal_unit_id: row.x_lokal_unit_id, x_lokal_certificate_url: row.x_lokal_certificate_url, x_lokal_production_date: row.x_lokal_production_date, x_lokal_expiry_date: row.x_lokal_expiry_date, x_lokal_proof_hash: row.x_lokal_proof_hash, x_lokal_status: row.x_lokal_status });
      if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update stock.lot ${name}`); else await this.writeSafe('stock.lot', found.id, values, sheet, `Lot diupdate: ${name}`); await this.ensureXmlId(xmlid, 'stock.lot', found.id, sheet); }
      else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create stock.lot ${name}`); continue; } const id = await this.odoo.create('stock.lot', values); this.log.ok(sheet, `Lot dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'stock.lot', id, sheet); }
    }
  }
  async processProjects() {
    const sheet = '07_PROJECTS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('project.project'))) { this.log.warn(sheet, 'Project tidak tersedia.'); return; }
    for (const row of rows) { const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) continue; const found = await this.findByXmlIdOrDomain(xmlid, 'project.project', [['name', '=', name]], ['id', 'name']); const values = compactObject({ name, active: toBool(row.active, true), allow_milestones: toBool(row.allow_milestones, true), label_tasks: row.label_tasks, description: row.description }); if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update project ${name}`); else await this.writeSafe('project.project', found.id, values, sheet, `Project diupdate: ${name}`); await this.ensureXmlId(xmlid, 'project.project', found.id, sheet); } else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create project ${name}`); continue; } const id = await this.odoo.create('project.project', values); this.log.ok(sheet, `Project dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'project.project', id, sheet); } }
  }
  async processProjectStages() {
    const sheet = '08_PROJECT_STAGES'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('project.task.type'))) { this.log.warn(sheet, 'project.task.type tidak tersedia.'); return; }
    for (const row of rows) { const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) continue; const found = await this.findByXmlIdOrDomain(xmlid, 'project.task.type', [['name', '=', name]], ['id', 'name']); const values = compactObject({ name, sequence: toNumberOrNull(row.sequence) || 10, fold: toBool(row.fold, false) }); if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update stage ${name}`); else await this.writeSafe('project.task.type', found.id, values, sheet, `Stage diupdate: ${name}`); await this.ensureXmlId(xmlid, 'project.task.type', found.id, sheet); } else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create stage ${name}`); continue; } const id = await this.odoo.create('project.task.type', values); this.log.ok(sheet, `Stage dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'project.task.type', id, sheet); } }
  }
  async processProjectTags() {
    const sheet = '09_PROJECT_TAGS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('project.tags'))) { this.log.warn(sheet, 'project.tags tidak tersedia.'); return; }
    for (const row of rows) { const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) continue; const found = await this.findByXmlIdOrDomain(xmlid, 'project.tags', [['name', '=', name]], ['id', 'name']); if (found) { this.log.ok(sheet, `Tag sudah ada: ${name}`); await this.ensureXmlId(xmlid, 'project.tags', found.id, sheet); continue; } if (this.dryRun) { this.log.info(sheet, `[dry-run] create tag ${name}`); continue; } const id = await this.odoo.create('project.tags', { name }); this.log.ok(sheet, `Tag dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'project.tags', id, sheet); }
  }
  async processMilestones() {
    const sheet = '10_MILESTONES'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('project.milestone'))) { this.log.warn(sheet, 'project.milestone tidak tersedia.'); return; }
    for (const row of rows) { const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) continue; const projectId = await this.m2o(row['project_id/id'] || row.project_external_id, 'project.project', sheet, true); const found = await this.findByXmlIdOrDomain(xmlid, 'project.milestone', [['name', '=', name], ['project_id', '=', projectId]], ['id', 'name']); const values = compactObject({ name, project_id: projectId, deadline: row.deadline || row.date_deadline, is_reached: toBool(row.is_reached, false) }); if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update milestone ${name}`); else await this.writeSafe('project.milestone', found.id, values, sheet, `Milestone diupdate: ${name}`); await this.ensureXmlId(xmlid, 'project.milestone', found.id, sheet); } else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create milestone ${name}`); continue; } const id = await this.odoo.create('project.milestone', values); this.log.ok(sheet, `Milestone dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'project.milestone', id, sheet); } }
  }
  async processTasks() {
    const sheet = '11_TASKS'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; }
    if (!(await this.modelExists('project.task'))) { this.log.warn(sheet, 'project.task tidak tersedia.'); return; }
    for (const [i, row] of rows.entries()) {
      const rowNum = i + 2; const xmlid = row.external_id || row.id; const name = row.name; if (isBlank(name)) { this.log.warn(sheet, `Row ${rowNum} dilewati: name kosong.`); continue; }
      const projectId = await this.m2o(row['project_id/id'] || row.project_external_id, 'project.project', sheet, true); const stageId = await this.m2o(row['stage_id/id'], 'project.task.type', sheet, false); const milestoneId = await this.m2o(row['milestone_id/id'], 'project.milestone', sheet, false); const tagCmd = await this.m2m(row['tag_ids/id'], 'project.tags', sheet);
      const found = await this.findByXmlIdOrDomain(xmlid, 'project.task', [['name', '=', name], ['project_id', '=', projectId]], ['id', 'name']); const values = compactObject({ name, project_id: projectId, stage_id: stageId, milestone_id: milestoneId, date_deadline: row.date_deadline, allocated_hours: toNumberOrNull(row.allocated_hours), priority: row.priority, sequence: toNumberOrNull(row.sequence), description: row.description }); if (tagCmd) values.tag_ids = tagCmd;
      if (found) { if (this.dryRun) this.log.info(sheet, `[dry-run] update task ${name}`); else await this.writeSafe('project.task', found.id, values, sheet, `Task diupdate: ${name}`); await this.ensureXmlId(xmlid, 'project.task', found.id, sheet); }
      else { if (this.dryRun) { this.log.info(sheet, `[dry-run] create task ${name}`); continue; } const id = await this.odoo.create('project.task', values); this.log.ok(sheet, `Task dibuat: ${name}`, { id }); await this.ensureXmlId(xmlid, 'project.task', id, sheet); }
    }
  }
  async processWebsitePages() { const sheet = '12_WEBSITE_PAGES'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; } this.log.warn(sheet, 'Website page tidak ditulis otomatis agar tidak merusak view. Sheet ini dibaca sebagai rencana.'); }
  async processQrRegistry() { const sheet = '13_QR_ID_REGISTRY'; const rows = sheetRows(this.workbook, sheet); if (!rows.length) { this.log.info(sheet, 'Sheet kosong/tidak ada.'); return; } this.log.info(sheet, `QR registry terbaca ${rows.length} baris. Data utama tetap dari x_lokal_id pada model terkait.`); }
}

async function testConnection() {
  const log = new ImportLog(); const odoo = new OdooClient({}, log);
  await odoo.authenticate(); const sample = await odoo.searchRead('res.partner', [], ['id', 'name'], 1);
  log.ok('AUTH', 'Koneksi Odoo OK dan res.partner bisa dibaca.', { sample: sample[0] || null });
  return { summary: log.summary(), logs: log.lines };
}

module.exports = { ImportLog, OdooClient, LokalmartImporter, testConnection, toBool };
