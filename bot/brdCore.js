/**
 * brdCore.js — Synoptek BRD Excel generation
 * Exports generateExcel(brdData, outputDir) → returns saved file path
 *
 * Sheets:
 *   1. Cover
 *   2. LES BRD            (main requirements register, 13 cols)
 *   3. Scope Checklist Requirements
 *   4. BRD                (Scope | Sub-part)
 *   5. Out of Scope
 *   6. Scope              (Scope | Description)
 *   7. Sign Off - Acceptance
 *   8. LOV's
 *   9. Selections
 *
 * Colours:  red #F04D38 (headers)  |  dark blue #002344 (titles)
 *           mid blue #004677 (sections)  |  tan #E8E4E2 (alt rows)
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  red:       'F04D38',
  darkBlue:  '002344',
  midBlue:   '004677',
  tan:       'E8E4E2',
  white:     'FFFFFF',
  dark:      '1A1A1A',
  grey:      '595959',
  lgrey:     'F5F5F5',
  border:    'BFBFBF',
  priHigh:   'FFDDD9',
  priMed:    'FFF6CC',
  priLow:    'DFF5E3',
  hasYes:    'DFF5E3',
  hasNo:     'FFDDD9',
};
const FONT = 'Aptos Narrow';

// ── Core helpers ──────────────────────────────────────────────────────────────
function bdr(cell, color = C.border, style = 'thin') {
  const b = { style, color: { argb: color } };
  cell.border = { top: b, left: b, bottom: b, right: b };
}

function sc(ws, row, col, value = '', opts = {}) {
  const {
    bg = C.white, fg = C.dark, bold = false,
    align = 'left', vAlign = 'middle',
    size = 10, border = true, wrap = true,
  } = opts;
  const c = ws.getCell(row, col);
  c.value = value;
  c.font = { name: FONT, size, bold, color: { argb: fg } };
  c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  c.alignment = { vertical: vAlign, horizontal: align, wrapText: wrap };
  if (border) bdr(c);
  return c;
}

// Red header cell
function hc(ws, row, col, value) {
  return sc(ws, row, col, value, { bg: C.red, fg: C.white, bold: true, align: 'center', size: 10 });
}

// Dark-blue title bar (merged)
function titleBar(ws, row, fromCol, toCol, value, size = 13) {
  ws.mergeCells(row, fromCol, row, toCol);
  return sc(ws, row, fromCol, value, { bg: C.darkBlue, fg: C.white, bold: true, align: 'center', size });
}

// Mid-blue section header (merged)
function secBar(ws, row, fromCol, toCol, value) {
  ws.mergeCells(row, fromCol, row, toCol);
  return sc(ws, row, fromCol, value, { bg: C.midBlue, fg: C.white, bold: true, align: 'left', size: 11 });
}

// Label cell (tan, right-aligned, bold)
function lc(ws, row, col, value) {
  return sc(ws, row, col, value, { bg: C.tan, fg: C.dark, bold: true, align: 'right', size: 10 });
}

// Data cell (alt-row aware)
function dc(ws, row, col, value, alt = false, opts = {}) {
  return sc(ws, row, col, value, { bg: alt ? C.tan : C.white, fg: C.dark, size: 10, ...opts });
}

// ── Sheet 1: Cover ────────────────────────────────────────────────────────────
function buildCover(wb, d) {
  const ws = wb.addWorksheet('Cover', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 3 },   // A padding
    { width: 28 },  // B labels
    { width: 40 },  // C values
    { width: 20 },  // D extra
    { width: 3 },   // E padding
  ];

  // ── Title banner ──
  ws.mergeCells('B2:D4');
  const t = ws.getCell('B2');
  t.value = 'CRM Business Requirements Document (BRD)';
  t.font = { name: FONT, size: 18, bold: true, color: { argb: C.white } };
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.darkBlue } };
  t.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  bdr(t, C.darkBlue, 'medium');
  [2, 3, 4].forEach(r => { ws.getRow(r).height = 26; });

  // ── Project sub-banner ──
  ws.mergeCells('B5:D5');
  const p = ws.getCell('B5');
  p.value = d.project_name || 'Project Name';
  p.font = { name: FONT, size: 14, bold: true, color: { argb: C.white } };
  p.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.midBlue } };
  p.alignment = { vertical: 'middle', horizontal: 'center' };
  bdr(p, C.midBlue, 'medium');
  ws.getRow(5).height = 24;
  ws.getRow(6).height = 8;  // spacer

  // ── Document Overview ──
  secBar(ws, 7, 2, 4, '  Document Overview');
  ws.getRow(7).height = 22;

  const reqs = d.requirements || [];
  const overview = [
    ['Purpose',            (d.executive_summary || '').slice(0, 160)],
    ['Client',             d.client_name || d.project_name || ''],
    ['Platform',           d.platform || 'Microsoft Dynamics 365 CRM'],
    ['Phase',              d.phase || 'Phase 1'],
    ['Version',            d.document_version || 'v1.0'],
    ['Date',               d.document_date || ''],
    ['Prepared By',        d.prepared_by || ''],
    ['Status',             d.status || 'Draft'],
    ['Total Requirements', String(reqs.length)],
  ];

  let r = 8;
  for (const [label, val] of overview) {
    lc(ws, r, 2, label);
    ws.mergeCells(r, 3, r, 4);
    sc(ws, r, 3, val, { bg: C.white, fg: C.dark, align: 'left', size: 10 });
    ws.getRow(r).height = 20;
    r++;
  }

  // ── Key Stakeholders ──
  r++;
  secBar(ws, r, 2, 4, '  Key Stakeholders');
  ws.getRow(r++).height = 22;
  ['Name', 'Role', 'Organization'].forEach((h, i) => hc(ws, r, i + 2, h));
  ws.getRow(r++).height = 20;
  (d.key_stakeholders || []).forEach((s, i) => {
    const alt = i % 2 !== 0;
    dc(ws, r, 2, s.name || '', alt);
    dc(ws, r, 3, s.role || '', alt);
    dc(ws, r, 4, s.organization || '', alt);
    ws.getRow(r++).height = 18;
  });

  // ── Scope Summary by Module ──
  r++;
  secBar(ws, r, 2, 4, '  Scope Summary by Module');
  ws.getRow(r++).height = 22;
  ['Module', '# In Scope', 'Description'].forEach((h, i) => hc(ws, r, i + 2, h));
  ws.getRow(r++).height = 20;
  (d.scope_summary || []).forEach((s, i) => {
    const alt = i % 2 !== 0;
    dc(ws, r, 2, s.module || '', alt);
    sc(ws, r, 3, String(s.in_scope_count || 0), { bg: alt ? C.tan : C.white, fg: C.dark, align: 'center', size: 10 });
    dc(ws, r, 4, s.description || '', alt);
    ws.getRow(r++).height = 18;
  });

  // ── Requirement Statistics ──
  r++;
  secBar(ws, r, 2, 4, '  Requirement Statistics');
  ws.getRow(r++).height = 22;
  const stats = [
    ['Total Requirements', reqs.length],
    ['High Priority',      reqs.filter(x => x.priority === 'High').length],
    ['Medium Priority',    reqs.filter(x => x.priority === 'Medium').length],
    ['Low Priority',       reqs.filter(x => x.priority === 'Low').length],
    ['Draft',              reqs.filter(x => x.status === 'Draft').length],
    ['Approved',           reqs.filter(x => x.status === 'Approved').length],
  ];
  for (const [label, val] of stats) {
    sc(ws, r, 2, label, { bg: C.tan, fg: C.dark, bold: true, align: 'right', size: 10 });
    ws.mergeCells(r, 3, r, 4);
    sc(ws, r, 3, String(val), { bg: C.white, fg: C.dark, bold: true, align: 'center', size: 11 });
    ws.getRow(r++).height = 18;
  }

  // ── Document Structure ──
  r++;
  secBar(ws, r, 2, 4, '  Document Structure');
  ws.getRow(r++).height = 22;
  const sheets = [
    'LES BRD', 'Scope Checklist Requirements', 'BRD',
    'Out of Scope', 'Scope', 'Sign Off - Acceptance', "LOV's", 'Selections',
  ];
  sheets.forEach((name, i) => {
    ws.mergeCells(r, 2, r, 4);
    sc(ws, r, 2, `${i + 1}.  ${name}`, { bg: i % 2 !== 0 ? C.tan : C.white, fg: C.dark, align: 'left', size: 10 });
    ws.getRow(r++).height = 18;
  });
}

// ── Sheet 2: LES BRD ──────────────────────────────────────────────────────────
function buildLESBRD(wb, d) {
  const ws = wb.addWorksheet('LES BRD', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 10 },  // 1  ID
    { width: 12 },  // 2  Date
    { width: 18 },  // 3  Source
    { width: 18 },  // 4  Requester
    { width: 52 },  // 5  Requirement Description
    { width: 14 },  // 6  Have today?
    { width: 12 },  // 7  Priority
    { width: 18 },  // 8  Requirement Status
    { width: 18 },  // 9  Scope
    { width: 42 },  // 10 Context / Business Rules
    { width: 30 },  // 11 Technical Comments
    { width: 22 },  // 12 Remarks
    { width: 20 },  // 13 Requested By
  ];

  titleBar(ws, 1, 1, 13, 'LES BRD — Requirements Register', 13);
  ws.getRow(1).height = 28;

  const HEADERS = [
    'ID', 'Date', 'Source', 'Requester',
    'Requirement Description\n(high-level need / spec / user story)',
    'Have today?', 'Priority', 'Requirement Status', 'Scope',
    'Context / examples / clarifications / business rules',
    'Technical Comments', 'Remarks', 'Requested By',
  ];
  HEADERS.forEach((h, i) => hc(ws, 2, i + 1, h));
  ws.getRow(2).height = 38;

  let r = 3;
  (d.requirements || []).forEach((req, i) => {
    const alt = i % 2 !== 0;
    const rowH = Math.max(28, Math.ceil((req.description || '').length / 55) * 14 + 8);

    const priBg = req.priority === 'High' ? C.priHigh : req.priority === 'Low' ? C.priLow : C.priMed;
    const htBg  = req.have_today === 'Yes' ? C.hasYes : C.hasNo;
    const base  = alt ? C.tan : C.white;

    const vals = [
      [req.id || `LES-${String(i + 1).padStart(3, '0')}`, base,  'center'],
      [req.date || d.document_date || '',                  base,  'left'],
      [req.source || '',                                   base,  'left'],
      [req.requester || '',                                base,  'left'],
      [req.description || '',                              base,  'left'],
      [req.have_today || 'No',                             htBg,  'center'],
      [req.priority || 'Medium',                           priBg, 'center'],
      [req.status || 'Draft',                              base,  'center'],
      [req.scope || '',                                    base,  'left'],
      [req.context || '',                                  base,  'left'],
      [req.technical_comments || '',                       base,  'left'],
      [req.remarks || '',                                  base,  'left'],
      [req.requested_by || '',                             base,  'left'],
    ];

    vals.forEach(([val, bg, align], j) => {
      sc(ws, r, j + 1, val, { bg, fg: C.dark, align, size: 10 });
    });
    ws.getRow(r++).height = rowH;
  });
}

// ── Sheet 3: Scope Checklist Requirements ─────────────────────────────────────
function buildScopeChecklist(wb, d) {
  const ws = wb.addWorksheet('Scope Checklist Requirements', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 24 },  // Category
    { width: 28 },  // Module
    { width: 30 },  // Sub-Category
    { width: 62 },  // Requirement Description
  ];

  titleBar(ws, 1, 1, 4, 'Scope Checklist Requirements');
  ws.getRow(1).height = 28;
  ['Category', 'Module', 'Sub-Category', 'Requirement Description'].forEach((h, i) => hc(ws, 2, i + 1, h));
  ws.getRow(2).height = 22;

  let r = 3;
  (d.scope_checklist || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    dc(ws, r, 1, item.category || '', alt);
    dc(ws, r, 2, item.module || '', alt);
    dc(ws, r, 3, item.sub_category || '', alt);
    dc(ws, r, 4, item.description || '', alt);
    ws.getRow(r++).height = 20;
  });
}

// ── Sheet 4: BRD ─────────────────────────────────────────────────────────────
function buildBRD(wb, d) {
  const ws = wb.addWorksheet('BRD', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 30 },  // Scope
    { width: 70 },  // Sub-part
  ];

  titleBar(ws, 1, 1, 2, 'BRD — Scope & Sub-parts');
  ws.getRow(1).height = 28;
  ['Scope', 'Sub-part'].forEach((h, i) => hc(ws, 2, i + 1, h));
  ws.getRow(2).height = 22;

  let r = 3;
  (d.brd_scope || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    dc(ws, r, 1, item.scope || '', alt);
    dc(ws, r, 2, item.sub_part || '', alt);
    ws.getRow(r++).height = 20;
  });
}

// ── Sheet 5: Out of Scope ─────────────────────────────────────────────────────
function buildOutOfScope(wb, d) {
  const ws = wb.addWorksheet('Out of Scope', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 90 }];

  titleBar(ws, 1, 1, 1, 'Out of Scope Items');
  ws.getRow(1).height = 28;
  hc(ws, 2, 1, 'Out of Scope Items');
  ws.getRow(2).height = 22;

  let r = 3;
  (d.out_of_scope || []).forEach((item, i) => {
    dc(ws, r, 1, typeof item === 'string' ? item : (item.item || ''), i % 2 !== 0);
    ws.getRow(r++).height = 20;
  });
}

// ── Sheet 6: Scope ────────────────────────────────────────────────────────────
function buildScope(wb, d) {
  const ws = wb.addWorksheet('Scope', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 30 },  // Scope
    { width: 72 },  // Description
  ];

  titleBar(ws, 1, 1, 2, 'Scope Definitions');
  ws.getRow(1).height = 28;
  ['Scope', 'Description'].forEach((h, i) => hc(ws, 2, i + 1, h));
  ws.getRow(2).height = 22;

  let r = 3;
  (d.scope_definitions || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    dc(ws, r, 1, item.scope || '', alt);
    dc(ws, r, 2, item.description || '', alt);
    ws.getRow(r++).height = 20;
  });
}

// ── Sheet 7: Sign Off - Acceptance ───────────────────────────────────────────
function buildSignOff(wb, d) {
  const ws = wb.addWorksheet('Sign Off - Acceptance', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 14 },  // Version
    { width: 36 },  // Name and Role
    { width: 36 },  // Signature
    { width: 18 },  // Date
  ];

  titleBar(ws, 1, 1, 4, 'Sign Off — Acceptance');
  ws.getRow(1).height = 28;
  ['Version', 'Name and Role', 'Signature', 'Date'].forEach((h, i) => hc(ws, 2, i + 1, h));
  ws.getRow(2).height = 22;

  const rows = (d.sign_off && d.sign_off.length)
    ? d.sign_off
    : [
        { version: d.document_version || 'v1.0', name_and_role: d.prepared_by || '', signature: '', date: '' },
        { version: '', name_and_role: 'Client Representative', signature: '', date: '' },
        { version: '', name_and_role: '', signature: '', date: '' },
      ];

  rows.forEach((row, i) => {
    const alt = i % 2 !== 0;
    dc(ws, i + 3, 1, row.version || '', alt);
    dc(ws, i + 3, 2, row.name_and_role || '', alt);
    sc(ws, i + 3, 3, row.signature || '', { bg: alt ? C.tan : C.lgrey, fg: C.grey, size: 10 });
    dc(ws, i + 3, 4, row.date || '', alt);
    ws.getRow(i + 3).height = 30;
  });
}

// ── Sheet 8: LOV's ────────────────────────────────────────────────────────────
function buildLOVs(wb, d) {
  const ws = wb.addWorksheet("LOV's", { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 45 },  // Epic
    { width: 30 },  // Cycle
  ];

  titleBar(ws, 1, 1, 2, "LOV's — Lists of Values");
  ws.getRow(1).height = 28;
  ['Epic', 'Cycle'].forEach((h, i) => hc(ws, 2, i + 1, h));
  ws.getRow(2).height = 22;

  let r = 3;
  (d.lovs || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    dc(ws, r, 1, item.epic || '', alt);
    dc(ws, r, 2, item.cycle || '', alt);
    ws.getRow(r++).height = 20;
  });
}

// ── Sheet 9: Selections ───────────────────────────────────────────────────────
function buildSelections(wb, d) {
  const ws = wb.addWorksheet('Selections', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 42 },  // Covered by?
    { width: 42 },  // Have today?
  ];

  titleBar(ws, 1, 1, 2, 'Selections');
  ws.getRow(1).height = 28;
  hc(ws, 2, 1, 'Covered by? Selections');
  hc(ws, 2, 2, 'Have today? Selections');
  ws.getRow(2).height = 22;

  const coveredBy = (d.covered_by_selections && d.covered_by_selections.length)
    ? d.covered_by_selections
    : ['Yes', 'No', 'Partial', 'N/A', 'Future Phase'];
  const haveToday = (d.have_today_selections && d.have_today_selections.length)
    ? d.have_today_selections
    : ['Yes', 'No', 'Partial'];

  const maxR = Math.max(coveredBy.length, haveToday.length);
  for (let i = 0; i < maxR; i++) {
    const alt = i % 2 !== 0;
    dc(ws, 3 + i, 1, coveredBy[i] || '', alt);
    dc(ws, 3 + i, 2, haveToday[i] || '', alt);
    ws.getRow(3 + i).height = 18;
  }
}

// ── Public API ────────────────────────────────────────────────────────────────
async function generateExcel(brdData, outputDir) {
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const wb = new ExcelJS.Workbook();
  wb.creator  = 'BRD Generator — Synoptek Format';
  wb.created  = new Date();
  wb.modified = new Date();

  buildCover(wb, brdData);
  buildLESBRD(wb, brdData);
  buildScopeChecklist(wb, brdData);
  buildBRD(wb, brdData);
  buildOutOfScope(wb, brdData);
  buildScope(wb, brdData);
  buildSignOff(wb, brdData);
  buildLOVs(wb, brdData);
  buildSelections(wb, brdData);

  const safe  = (brdData.project_name || 'BRD').replace(/[^a-zA-Z0-9_\- ]/g, '').trim().replace(/ /g, '_');
  const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const name  = `${safe}_BRD_${stamp}.xlsx`;
  const file  = path.join(outputDir, name);
  await wb.xlsx.writeFile(file);
  return file;
}

module.exports = { generateExcel };
