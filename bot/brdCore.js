/**
 * brdCore.js — Synoptek BRD Excel generation
 * Unified format across all sheets:
 *   • Date header right-aligned (row 1)
 *   • Spacer row (row 2)
 *   • Bold red sheet title (rows 3–4)
 *   • Spacer row (row 5)
 *   • Light-gray (E7E6E6) section/column headers (row 6)
 *   • Alternating white / light-gray (F2F2F2) data rows from row 7
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const CV = {
  secHdr: 'E7E6E6',   // section / column header background (light gray)
  white:  'FFFFFF',
  red:    'C00000',   // sheet title text
  alt:    'F2F2F2',   // alternating data row
  dark:   '1A1A1A',   // primary text
  border: 'BFBFBF',
};
const FONT = 'Aptos Narrow';

// ── Project-name helpers ──────────────────────────────────────────────────────

/** Derive a 2–4 char uppercase prefix from the project name for requirement IDs */
function getBRDPrefix(projectName) {
  const name  = (projectName || 'BRD').trim();
  const words = name.split(/\s+/).filter(Boolean);
  if (words.length > 1) {
    const initials = words.map(w => (w.replace(/[^A-Za-z0-9]/g, '')[0] || '')).join('').toUpperCase();
    return initials.slice(0, 4) || 'BRD';
  }
  const clean = name.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  return clean.length <= 4 ? (clean || 'BRD') : clean.slice(0, 3);
}

/** Build a safe Excel sheet name: project name truncated to 27 chars + " BRD" */
function getBRDSheetName(projectName) {
  const safe = (projectName || 'BRD').replace(/[\[\]:*?/\\]/g, '').trim().slice(0, 27);
  return safe ? `${safe} BRD` : 'BRD';
}

// ── Shared helpers ────────────────────────────────────────────────────────────

function cvBdr(cell) {
  const b = { style: 'thin', color: { argb: CV.border } };
  cell.border = { top: b, left: b, bottom: b, right: b };
}

/**
 * Adds the standard 5-row header block used by every sheet:
 *   Row 1    date right-aligned (merged across all cols)
 *   Row 2    thin spacer
 *   Rows 3-4 red bold title (merged across all cols)
 *   Row 5    thin spacer
 */
function addSheetHeader(ws, d, title, numCols) {
  ws.mergeCells(1, 1, 1, numCols);
  const dateCell = ws.getCell(1, 1);
  dateCell.value = `Document Generated: ${d.document_date || ''}`;
  dateCell.font  = { name: FONT, size: 11, color: { argb: CV.dark } };
  dateCell.alignment = { vertical: 'middle', horizontal: 'right' };
  ws.getRow(1).height = 24;

  ws.getRow(2).height = 6;

  ws.mergeCells(3, 1, 4, numCols);
  const t = ws.getCell(3, 1);
  t.value = title;
  t.font  = { name: FONT, size: 16, bold: true, color: { argb: CV.red } };
  t.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.white } };
  t.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  cvBdr(t);
  ws.getRow(3).height = 24;
  ws.getRow(4).height = 24;
  ws.getRow(5).height = 6;
}

/** Light-gray column header cell (dark text) */
function nvHdr(ws, row, col, value, wrap = false) {
  const c = ws.getCell(row, col);
  c.value = value;
  c.font  = { name: FONT, size: 11, bold: true, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.secHdr } };
  c.alignment = { vertical: 'middle', horizontal: 'left', wrapText: wrap };
  cvBdr(c);
  return c;
}

/** Light-gray section sub-header (merged fromCol–toCol, dark text) */
function nvSecHdr(ws, row, fromCol, toCol, label) {
  if (fromCol < toCol) ws.mergeCells(row, fromCol, row, toCol);
  const c = ws.getCell(row, fromCol);
  c.value = label;
  c.font  = { name: FONT, size: 11, bold: true, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.secHdr } };
  c.alignment = { vertical: 'middle', horizontal: 'left', wrapText: false };
  cvBdr(c);
  ws.getRow(row).height = 22;
}

/** Data cell with alternating row background */
function nvDat(ws, row, col, value, { alt = false, wrap = false, align = 'left' } = {}) {
  const c = ws.getCell(row, col);
  c.value = value || '';
  c.font  = { name: FONT, size: 11, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: alt ? CV.alt : CV.white } };
  c.alignment = { vertical: 'top', horizontal: align, wrapText: wrap };
  cvBdr(c);
  return c;
}

// ── Cover-specific helpers (label / value two-column layout) ─────────────────

function cvSecHdr(ws, row, label) {
  ws.mergeCells(row, 2, row, 3);
  const c = ws.getCell(row, 2);
  c.value = label;
  c.font  = { name: FONT, size: 12, bold: true, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.secHdr } };
  c.alignment = { vertical: 'middle', horizontal: 'left', wrapText: false };
  cvBdr(c);
  ws.getRow(row).height = 22;
}

function cvLbl(ws, row, value) {
  const c = ws.getCell(row, 2);
  c.value = value;
  c.font  = { name: FONT, size: 11, bold: true, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.white } };
  c.alignment = { vertical: 'middle', horizontal: 'left', wrapText: false };
  cvBdr(c);
}

function cvVal(ws, row, value, alt = false, wrap = false) {
  const c = ws.getCell(row, 3);
  c.value = value || '';
  c.font  = { name: FONT, size: 11, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: alt ? CV.alt : CV.white } };
  c.alignment = { vertical: 'middle', horizontal: 'left', wrapText: wrap };
  cvBdr(c);
}

function cvFull(ws, row, value, alt = false) {
  ws.mergeCells(row, 2, row, 3);
  const c = ws.getCell(row, 2);
  c.value = value || '';
  c.font  = { name: FONT, size: 11, color: { argb: CV.dark } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: alt ? CV.alt : CV.white } };
  c.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  cvBdr(c);
}

// ── Sheet 1: Cover ────────────────────────────────────────────────────────────
function buildCover(wb, d) {
  const ws = wb.addWorksheet('Cover', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 3  },   // A  padding
    { width: 35 },   // B  labels
    { width: 62 },   // C  values
    { width: 3  },   // D  padding
  ];

  const title = `${d.project_name || 'BRD'} CRM Business Requirements Document (BRD)`;
  addSheetHeader(ws, d, title, 4);

  // ── Document Overview ──
  const reqs = d.requirements || [];
  let r = 6;
  cvSecHdr(ws, r++, '  📋  Document Overview');

  const overview = [
    ['Purpose',            (d.executive_summary || '').slice(0, 300)],
    ['Client',             d.client_name || d.project_name || ''],
    ['Platform',           d.platform || 'Microsoft Dynamics 365 CE (Sales + Customer Service)'],
    ['Release',            d.phase || 'Release 1 - CRM Implementation'],
    ['Total Requirements', String(reqs.length)],
  ];
  overview.forEach(([label, val], i) => {
    cvLbl(ws, r, label);
    cvVal(ws, r, val, i % 2 !== 0, label === 'Purpose');
    ws.getRow(r++).height = label === 'Purpose' ? 36 : 18;
  });

  // ── Key Stakeholders ──
  ws.getRow(r++).height = 6;
  cvSecHdr(ws, r++, '  👥  Key Stakeholders');
  (d.key_stakeholders || []).forEach((s, i) => {
    cvLbl(ws, r, s.name || '');
    cvVal(ws, r, s.role || '', i % 2 !== 0);
    ws.getRow(r++).height = 18;
  });

  // ── Scope Summary by Module ──
  ws.getRow(r++).height = 6;
  cvSecHdr(ws, r++, '  📂  Scope Summary by Module');

  const mHdr = ws.getCell(r, 2);
  mHdr.value = 'Module';
  mHdr.font  = { name: FONT, size: 11, bold: true, color: { argb: CV.dark } };
  mHdr.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.white } };
  mHdr.alignment = { vertical: 'middle', horizontal: 'left' };
  cvBdr(mHdr);

  const dHdr = ws.getCell(r, 3);
  dHdr.value = 'Description';
  dHdr.font  = { name: FONT, size: 11, bold: true, color: { argb: CV.dark } };
  dHdr.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: CV.white } };
  dHdr.alignment = { vertical: 'middle', horizontal: 'left' };
  cvBdr(dHdr);
  ws.getRow(r++).height = 18;

  (d.scope_summary || []).forEach((s, i) => {
    cvLbl(ws, r, s.module || '');
    cvVal(ws, r, s.description || '', i % 2 !== 0, true);
    ws.getRow(r++).height = 20;
  });

  // ── Requirement Statistics ──
  ws.getRow(r++).height = 6;
  cvSecHdr(ws, r++, '  📊  Requirement Statistics');
  const stats = [
    ['Total Requirements', reqs.length],
    ['Must Have',          reqs.filter(x => x.priority === 'Must Have').length],
    ['Should Have',        reqs.filter(x => x.priority === 'Should Have').length],
    ['Could Have',         reqs.filter(x => x.priority === 'Could Have').length],
    ["Won't Have",         reqs.filter(x => x.priority === "Won't Have").length],
    ['Draft / New',        reqs.filter(x => x.status === 'Draft' || x.status === 'New').length],
    ['Approved',           reqs.filter(x => x.status === 'Approved').length],
  ];
  stats.forEach(([label, val], i) => {
    cvLbl(ws, r, label);
    cvVal(ws, r, String(val), i % 2 !== 0);
    ws.getRow(r++).height = 18;
  });

  // ── Document Structure ──
  ws.getRow(r++).height = 6;
  cvSecHdr(ws, r++, '  📄  Document Structure');
  const sheets = [
    getBRDSheetName(d.project_name), 'Scope Checklist Requirements', 'Fit-Gap Analysis',
    'Out of Scope', 'Scope', 'Sign Off - Acceptance', "LOV's", 'Selections',
  ];
  sheets.forEach((name, i) => {
    cvFull(ws, r, `${i + 1}.  ${name}`, i % 2 !== 0);
    ws.getRow(r++).height = 18;
  });
}

// ── Sheet 2: BRD (sheet name derived from project name) ───────────────────────
function buildLESBRD(wb, d) {
  const ws = wb.addWorksheet(getBRDSheetName(d.project_name), {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [
    { width: 13.3 },  //  1  ID
    { width: 11.3 },  //  2  Date
    { width: 17.4 },  //  3  Source
    { width: 14.7 },  //  4  Scope
    { width: 38.3 },  //  5  Sub-part
    { width: 60.1 },  //  6  Requirement
    { width: 32.1 },  //  7  Have today?
    { width: 10.4 },  //  8  Priority
    { width: 15.7 },  //  9  Requirement Status
    { width: 27.1 },  // 10  Implementation Approach
    { width: 69.4 },  // 11  Context
    { width: 44.6 },  // 12  Technical Comments
    { width: 28.7 },  // 13  Remarks
    { width: 39.3 },  // 14  Lead Requested
    { width: 53.7 },  // 15  Meeting Participants
  ];

  const title = `${d.project_name || 'BRD'} — D365 CRM Business Requirements Document (BRD)`;
  addSheetHeader(ws, d, title, 15);

  // Row 6: column headers
  const HEADERS = [
    'ID', 'Date', 'Source', 'Scope', 'Sub-part',
    'Requirement as high-level need, spec or user story',
    'Have today?', 'Priority', 'Requirement Status', 'Implementation Approach',
    'Context, examples, clarifications, business rules',
    'Technical Comments', 'Remarks', 'Lead Requested', 'Meeting Participants',
  ];
  const wrapH = new Set([5, 10, 11, 12]);
  HEADERS.forEach((h, i) => nvHdr(ws, 6, i + 1, h, wrapH.has(i)));
  ws.getRow(6).height = 32;

  // Data rows from row 7
  let r = 7;
  (d.requirements || []).forEach((req, i) => {
    const id  = req.id || `${getBRDPrefix(d.project_name)}-${String(i + 1).padStart(3, '0')}`;
    const alt = i % 2 !== 0;
    const vals = [
      [id,                                                              false],
      [req.date      || '',                                            false],
      [req.source    || '',                                            false],
      [req.scope     || '',                                            false],
      [req.sub_part  || '',                                            false],
      [req.description || '',                                          true ],
      [req.have_today || "Don't have, gap/pain point - Need improvement", false],
      [req.priority  || 'Should Have',                                 false],
      [req.status    || 'New',                                         false],
      [req.scope2    || '',                                            false],
      [req.context   || '',                                            true ],
      [req.technical_comments || '',                                   true ],
      [req.remarks   || '',                                            true ],
      [req.requester || '',                                            false],
      [req.requested_by || '',                                         false],
    ];
    vals.forEach(([val, wrap], j) => nvDat(ws, r, j + 1, val, { alt, wrap }));
    ws.getRow(r++).height = 20;
  });
}

// ── Sheet 3: Scope Checklist Requirements ─────────────────────────────────────
function buildScopeChecklist(wb, d) {
  const ws = wb.addWorksheet('Scope Checklist Requirements', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [
    { width: 22.4  },  // Category
    { width: 22.0  },  // Module
    { width: 28.0  },  // Sub-Category
    { width: 131.6 },  // Requirement Description
  ];

  addSheetHeader(ws, d, 'Scope Checklist Requirements', 4);

  ['Category', 'Module', 'Sub-Category', 'Requirement Description'].forEach((h, i) =>
    nvHdr(ws, 6, i + 1, h, i === 3)
  );
  ws.getRow(6).height = 22;

  let r = 7;
  (d.scope_checklist || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    nvDat(ws, r, 1, item.category    || '', { alt });
    nvDat(ws, r, 2, item.module      || '', { alt });
    nvDat(ws, r, 3, item.sub_category || '', { alt });
    nvDat(ws, r, 4, item.description  || '', { alt, wrap: true });
    ws.getRow(r++).height = 21;
  });
}

// ── Sheet 4: Fit-Gap Analysis ─────────────────────────────────────────────────
function buildFitGap(wb, d) {
  const ws = wb.addWorksheet('Fit-Gap Analysis', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [
    { width: 13.3 },   // 1  Req ID
    { width: 55.0 },   // 2  Requirement
    { width: 50.0 },   // 3  Current State
    { width: 50.0 },   // 4  Gap
    { width: 60.0 },   // 5  Recommendation
    { width: 12.0 },   // 6  Priority
  ];

  addSheetHeader(ws, d, 'Fit-Gap Analysis', 6);

  ['Req ID', 'Requirement', 'Current State (Have Today)', 'Gap / Pain Point', 'Recommendation / D365 Approach', 'Priority'].forEach((h, i) =>
    nvHdr(ws, 6, i + 1, h, i > 0 && i < 5)
  );
  ws.getRow(6).height = 28;

  let r = 7;
  (d.fit_gap_analysis || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    nvDat(ws, r, 1, item.requirement_id || '', { alt });
    nvDat(ws, r, 2, item.requirement    || '', { alt, wrap: true });
    nvDat(ws, r, 3, item.current_state  || '', { alt, wrap: true });
    nvDat(ws, r, 4, item.gap            || '', { alt, wrap: true });
    nvDat(ws, r, 5, item.recommendation || '', { alt, wrap: true });
    nvDat(ws, r, 6, item.priority       || '', { alt });
    ws.getRow(r++).height = 28;
  });
}

// ── Sheet 5: Out of Scope ─────────────────────────────────────────────────────
function buildOutOfScope(wb, d) {
  const ws = wb.addWorksheet('Out of Scope', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [{ width: 222.9 }];

  addSheetHeader(ws, d, 'Out of Scope', 1);

  nvHdr(ws, 6, 1, 'Out of Scope Items');
  ws.getRow(6).height = 22;

  let r = 7;
  (d.out_of_scope || []).forEach((item, i) => {
    nvDat(ws, r, 1, typeof item === 'string' ? item : (item.item || ''), {
      alt: i % 2 !== 0, wrap: true,
    });
    ws.getRow(r++).height = 21;
  });
}

// ── Sheet 6: Scope ────────────────────────────────────────────────────────────
function buildScope(wb, d) {
  const ws = wb.addWorksheet('Scope', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [
    { width: 46.3  },  // Scope
    { width: 104.0 },  // Description
  ];

  addSheetHeader(ws, d, 'Scope', 2);

  ['Scope', 'Description'].forEach((h, i) => nvHdr(ws, 6, i + 1, h, i === 1));
  ws.getRow(6).height = 22;

  let r = 7;
  (d.scope_definitions || []).forEach((item, i) => {
    const alt = i % 2 !== 0;
    nvDat(ws, r, 1, item.scope       || '', { alt });
    nvDat(ws, r, 2, item.description || '', { alt, wrap: true });
    ws.getRow(r++).height = 31;
  });
}

// ── Sheet 7: Sign Off - Acceptance ───────────────────────────────────────────
function buildSignOff(wb, d) {
  const ws = wb.addWorksheet('Sign Off - Acceptance', {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [
    { width: 22.0 },  // Version
    { width: 35.0 },  // Name and Role
    { width: 35.0 },  // Signature
    { width: 18.0 },  // Date
  ];

  addSheetHeader(ws, d,
    'Signature & Acceptance — Signing indicates understanding and acceptance of this document.',
    4);

  ['Version', 'Name and Role', 'Signature', 'Date'].forEach((h, i) =>
    nvHdr(ws, 6, i + 1, h)
  );
  ws.getRow(6).height = 22;

  const rows = (d.sign_off && d.sign_off.length)
    ? d.sign_off
    : [
        { version: `${d.document_version || 'v1.0'} – Draft for Review`,  name_and_role: d.prepared_by || 'Synoptek – Project Manager' },
        { version: `${d.document_version || 'v1.0'} – MSP Approval`,      name_and_role: 'MSP – Project Manager' },
        { version: `${d.document_version || 'v1.0'} – Client Acceptance`, name_and_role: 'Client – Project Sponsor' },
        { version: '',                                                      name_and_role: '' },
      ];

  rows.forEach((row, i) => {
    const r   = 7 + i;
    const alt = i % 2 !== 0;
    nvDat(ws, r, 1, row.version       || '', { alt });
    nvDat(ws, r, 2, row.name_and_role || '', { alt });
    nvDat(ws, r, 3, row.signature     || '', { alt });
    nvDat(ws, r, 4, row.date          || '', { alt });
    ws.getRow(r).height = 25;
  });
}

// ── Sheet 8: LOV's ────────────────────────────────────────────────────────────
function buildLOVs(wb, d) {
  const ws = wb.addWorksheet("LOV's", {
    views: [{ showGridLines: false, state: 'frozen', ySplit: 6 }],
  });
  ws.columns = [
    { width: 25.0 },  // Epic
    { width: 30.0 },  // Cycle
    { width: 49.6 },  // Combined
    { width: 10.9 },  // Tag
  ];

  addSheetHeader(ws, d, 'Lists of Values', 4);

  ['Epic', 'Cycle', 'Column1', 'Column2'].forEach((h, i) => nvHdr(ws, 6, i + 1, h));
  ws.getRow(6).height = 22;

  let r = 7;
  (d.lovs || []).forEach((item, i) => {
    const alt   = i % 2 !== 0;
    const epic  = item.epic  || '';
    const cycle = item.cycle || '';
    nvDat(ws, r, 1, epic, { alt });
    nvDat(ws, r, 2, cycle, { alt });
    nvDat(ws, r, 3, epic && cycle ? `${epic}: ${cycle}` : (epic || cycle), { alt });
    nvDat(ws, r, 4, item.tag || '', { alt });
    ws.getRow(r++).height = 21;
  });
}

// ── Sheet 9: Selections ───────────────────────────────────────────────────────
function buildSelections(wb, d) {
  const ws = wb.addWorksheet('Selections', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 71.4 }];

  addSheetHeader(ws, d, 'Selections', 1);

  // Covered by? section
  nvSecHdr(ws, 6, 1, 1, '  Covered by? Selections');
  const coveredBy = (d.covered_by_selections && d.covered_by_selections.length)
    ? d.covered_by_selections
    : [
        'D365 OOB no config',
        'D365 OOB with config',
        'Customization/development',
        'Config & workflow automation',
        'Config & dashboard',
        '3rd party solution',
        'Config & integration',
        'Business procedure',
        'Other',
      ];
  coveredBy.forEach((v, i) => {
    nvDat(ws, 7 + i, 1, v, { alt: i % 2 !== 0 });
    ws.getRow(7 + i).height = 16;
  });

  // Have today? section
  const sep = 7 + coveredBy.length;
  ws.getRow(sep).height = 6;
  nvSecHdr(ws, sep + 1, 1, 1, '  Have today? Selections');
  const haveToday = (d.have_today_selections && d.have_today_selections.length)
    ? d.have_today_selections
    : [
        'Have and want to keep',
        'Have and want to change/improve',
        "Have and don't want to keep",
        "Don't have, gap/pain point - Need improvement",
        "Don't have and don't need/want",
      ];
  haveToday.forEach((v, i) => {
    nvDat(ws, sep + 2 + i, 1, v, { alt: i % 2 !== 0 });
    ws.getRow(sep + 2 + i).height = 16;
  });
}

// ── Public API ────────────────────────────────────────────────────────────────
async function generateExcel(brdData, outputDir) {
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const wb = new ExcelJS.Workbook();
  wb.creator  = 'BRD Generator — Synoptek';
  wb.created  = new Date();
  wb.modified = new Date();

  buildCover(wb, brdData);
  buildLESBRD(wb, brdData);
  buildScopeChecklist(wb, brdData);
  buildFitGap(wb, brdData);
  buildOutOfScope(wb, brdData);
  buildScope(wb, brdData);
  buildSignOff(wb, brdData);
  buildLOVs(wb, brdData);
  buildSelections(wb, brdData);

  const safe  = (brdData.project_name || 'BRD').replace(/[^a-zA-Z0-9_\- ]/g, '').trim().replace(/ /g, '_');
  const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const ver   = (brdData.document_version || 'v1.0').toUpperCase();
  const name  = `${safe}_BRD_${stamp}_${ver}.xlsx`;
  const file  = path.join(outputDir, name);
  await wb.xlsx.writeFile(file);
  return file;
}

module.exports = { generateExcel };
