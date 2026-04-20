/**
 * BRD Generator – Produces a fully populated Excel BRD document
 * Usage:
 *   node generate_brd.js                        (uses built-in sample data)
 *   node generate_brd.js --api-key sk-ant-...   (calls Claude API with sample docs)
 *
 * Output: output/<ProjectName>_BRD_<date>.xlsx
 */

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');
const https   = require('https');

// ── Colour palette ──────────────────────────────────────────────────────────
const C = {
  navyBg:     '1F3864',
  blueBg:     '2E75B6',
  colHeadBg:  '4472C4',
  labelBg:    'D6E4F0',
  altRow:     'EBF3FB',
  highYellow: 'FFF2CC',
  highGreen:  'E2EFDA',
  highRed:    'FCE4D6',
  highOrange: 'FCE9D0',
  white:      'FFFFFF',
  dark:       '1F1F1F',
  grey:       '595959',
  borderGrey: 'BFBFBF',
};

// ── Priority → colour map ────────────────────────────────────────────────────
const PRIORITY_COLOR = { High: 'FCE4D6', Medium: 'FFF2CC', Low: 'E2EFDA' };
const RISK_COLOR     = { High: 'FCE4D6', Medium: 'FFF2CC', Low: 'E2EFDA',
                         H:    'FCE4D6', M:      'FFF2CC', L:   'E2EFDA' };

// ── Border helper ────────────────────────────────────────────────────────────
function border(cell, color = C.borderGrey, style = 'thin') {
  const b = { style, color: { argb: color } };
  cell.border = { top: b, left: b, bottom: b, right: b };
}

// ── Generic cell writers ─────────────────────────────────────────────────────
function hdr(ws, row, col, val, bg = C.colHeadBg, fg = C.white, sz = 10, bold = true) {
  const c = ws.getCell(row, col);
  c.value = val;
  c.font  = { name: 'Calibri', size: sz, bold, color: { argb: fg } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  border(c);
  return c;
}

function cell(ws, row, col, val = '', bg = C.white, fg = C.dark, bold = false, align = 'left') {
  const c = ws.getCell(row, col);
  c.value = val;
  c.font  = { name: 'Calibri', size: 10, bold, color: { argb: fg } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
  c.alignment = { vertical: 'middle', horizontal: align, wrapText: true };
  border(c);
  return c;
}

function label(ws, row, col, val) {
  const c = cell(ws, row, col, val, C.labelBg, C.dark, true, 'right');
  return c;
}

function secTitle(ws, row, colFrom, colTo, val) {
  ws.mergeCells(row, colFrom, row, colTo);
  const c = ws.getCell(row, colFrom);
  c.value = val;
  c.font  = { name: 'Calibri', size: 11, bold: true, color: { argb: C.white } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.blueBg } };
  c.alignment = { vertical: 'middle', horizontal: 'left', indent: 1, wrapText: true };
  border(c, C.navyBg, 'medium');
  ws.getRow(row).height = 22;
  return c;
}

function sheetTitle(ws, colTo, val) {
  ws.mergeCells(1, 1, 2, colTo);
  const c = ws.getCell(1, 1);
  c.value = val;
  c.font  = { name: 'Calibri', size: 14, bold: true, color: { argb: C.white } };
  c.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.navyBg } };
  c.alignment = { vertical: 'middle', horizontal: 'center' };
  border(c, C.navyBg, 'medium');
  ws.getRow(1).height = 28;
  ws.getRow(2).height = 28;
  return c;
}

// ══════════════════════════════════════════════════════════════════════════════
// 1. COVER PAGE
// ══════════════════════════════════════════════════════════════════════════════
function buildCover(wb, d) {
  const ws = wb.addWorksheet('Cover Page', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 4 }, { width: 30 }, { width: 50 }, { width: 20 }, { width: 4 }];

  // Banner
  ws.mergeCells('B2:D6');
  const banner = ws.getCell('B2');
  banner.value = 'BUSINESS REQUIREMENTS\nDOCUMENT';
  banner.font  = { name: 'Calibri', size: 26, bold: true, color: { argb: C.white } };
  banner.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.navyBg } };
  banner.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  border(banner, C.navyBg, 'medium');
  [2,3,4,5,6].forEach(r => ws.getRow(r).height = 22);

  // Sub-title
  ws.mergeCells('B7:D7');
  const sub = ws.getCell('B7');
  sub.value = d.project_name || 'Project Name';
  sub.font  = { name: 'Calibri', size: 15, bold: true, color: { argb: C.white } };
  sub.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.blueBg } };
  sub.alignment = { vertical: 'middle', horizontal: 'center' };
  border(sub, C.blueBg, 'medium');
  ws.getRow(7).height = 26;

  const rows = [
    ['Project Name',         d.project_name],
    ['Client / Organization',d.client_name],
    ['Prepared By',          d.prepared_by],
    ['Document Version',     d.document_version],
    ['Date',                 d.document_date],
    ['Status',               d.status],
    ['Confidentiality',      'Confidential'],
  ];

  let r = 9;
  rows.forEach(([lbl, val]) => {
    ws.getRow(r).height = 22;
    label(ws, r, 2, lbl);
    ws.mergeCells(r, 3, r, 4);
    const vc = cell(ws, r, 3, val || '', C.highYellow, C.dark, false, 'left');
    vc.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    r++;
  });

  ws.mergeCells(`B${r+1}:D${r+1}`);
  const note = ws.getCell(`B${r+1}`);
  note.value = 'This document is generated by the BRD Generator Agent. Review all sections before sign-off.';
  note.font  = { name: 'Calibri', size: 9, italic: true, color: { argb: '808080' } };
  note.alignment = { horizontal: 'center' };
}

// ══════════════════════════════════════════════════════════════════════════════
// 2. DOCUMENT CONTROL
// ══════════════════════════════════════════════════════════════════════════════
function buildDocControl(wb, d) {
  const ws = wb.addWorksheet('Document Control', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 12 }, { width: 18 }, { width: 40 }, { width: 22 }, { width: 22 }, { width: 16 }];

  sheetTitle(ws, 6, 'Document Control');
  let r = 4;

  secTitle(ws, r++, 1, 6, 'Version History');
  ['Version','Date','Description of Change','Author','Reviewed By','Status']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;
  hdr(ws, r, 1, d.document_version || 'v1.0', C.white, C.dark, 10, false);
  cell(ws, r, 2, d.document_date, C.white, C.dark);
  cell(ws, r, 3, 'Initial draft generated by BRD Generator Agent', C.white, C.dark);
  cell(ws, r, 4, d.prepared_by, C.white, C.dark);
  cell(ws, r, 5, '', C.white, C.dark);
  cell(ws, r, 6, d.status || 'Draft', PRIORITY_COLOR['Low'] || C.highYellow, C.dark, true, 'center');
  ws.getRow(r++).height = 22;
  for (let i = 0; i < 4; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    for (let c = 1; c <= 6; c++) cell(ws, r, c, '', bg);
    ws.getRow(r++).height = 20;
  }

  r++;
  secTitle(ws, r++, 1, 6, 'Document Approvals');
  ['Role','Name','Department','Signature','Date','Status']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;
  for (let i = 0; i < 5; i++) {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    for (let c = 1; c <= 6; c++) cell(ws, r, c, '', bg);
    ws.getRow(r++).height = 22;
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// 3. EXECUTIVE SUMMARY
// ══════════════════════════════════════════════════════════════════════════════
function buildExecSummary(wb, d) {
  const ws = wb.addWorksheet('Executive Summary', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 28 }, { width: 80 }];

  sheetTitle(ws, 2, 'Executive Summary');

  const sections = [
    ['Executive Summary',         d.executive_summary],
    ['Project Overview',          d.project_overview],
    ['Business Problem / Opportunity', d.business_problem],
    ['Proposed Solution',         d.proposed_solution],
    ['Expected Business Value',   d.business_value],
    ['Current State (As-Is)',     d.current_state],
    ['Future State (To-Be)',      d.future_state],
    ['Gap Analysis',              d.gap_analysis],
    ['Success Criteria',          d.success_criteria],
  ];

  let r = 4;
  sections.forEach(([lbl, val]) => {
    label(ws, r, 1, lbl);
    const vc = cell(ws, r, 2, val || '', C.white, C.dark);
    vc.alignment = { vertical: 'top', wrapText: true };
    const lines = Math.max(3, Math.ceil((val || '').length / 90));
    ws.getRow(r).height = lines * 15 + 10;
    r++;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 4. BUSINESS OBJECTIVES
// ══════════════════════════════════════════════════════════════════════════════
function buildObjectives(wb, d) {
  const ws = wb.addWorksheet('Business Objectives', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 12 }, { width: 28 }, { width: 50 }, { width: 14 }, { width: 22 }];

  sheetTitle(ws, 5, 'Business Objectives & KPIs');
  let r = 4;

  secTitle(ws, r++, 1, 5, 'Business Objectives');
  ['Obj. ID','Objective','Description / Rationale','Priority','Owner']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.business_objectives || []).forEach((o, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, o.id,          bg, C.dark, true,  'center');
    cell(ws, r, 2, o.objective,   bg, C.dark, true);
    cell(ws, r, 3, o.description, bg, C.dark);
    const pc = cell(ws, r, 4, o.priority, PRIORITY_COLOR[o.priority] || bg, C.dark, true, 'center');
    cell(ws, r, 5, o.owner, bg, C.dark);
    ws.getRow(r++).height = 25;
  });

  r++;
  secTitle(ws, r++, 1, 5, 'Key Performance Indicators (KPIs)');
  ['KPI ID','KPI Name','Measurement Method','Baseline','Target']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.kpis || []).forEach((k, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, k.id,          bg, C.dark, true,  'center');
    cell(ws, r, 2, k.name,        bg, C.dark, true);
    cell(ws, r, 3, k.measurement, bg, C.dark);
    cell(ws, r, 4, k.baseline,    bg, C.grey);
    cell(ws, r, 5, k.target,      C.highGreen, C.dark, true);
    ws.getRow(r++).height = 25;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 5. SCOPE
// ══════════════════════════════════════════════════════════════════════════════
function buildScope(wb, d) {
  const ws = wb.addWorksheet('Scope', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 10 }, { width: 60 }, { width: 35 }];

  sheetTitle(ws, 3, 'Project Scope');
  let r = 4;

  secTitle(ws, r++, 1, 3, 'In Scope');
  ['#','In-Scope Item / Feature','Notes'].forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;
  (d.in_scope || []).forEach((s, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, s.id,    bg, C.dark, true,  'center');
    cell(ws, r, 2, s.item,  bg, C.dark);
    cell(ws, r, 3, s.notes, bg, C.grey);
    ws.getRow(r++).height = 22;
  });

  r++;
  secTitle(ws, r++, 1, 3, 'Out of Scope');
  ['#','Out-of-Scope Item','Reason'].forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;
  (d.out_of_scope || []).forEach((s, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, s.id,     bg, C.dark, true,  'center');
    cell(ws, r, 2, s.item,   C.highRed.replace('FCE4D6','FFE7E0'), C.dark);
    cell(ws, r, 3, s.reason, bg, C.grey);
    ws.getRow(r++).height = 22;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 6. STAKEHOLDERS
// ══════════════════════════════════════════════════════════════════════════════
function buildStakeholders(wb, d) {
  const ws = wb.addWorksheet('Stakeholders', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 10 }, { width: 22 }, { width: 24 }, { width: 22 },
    { width: 32 }, { width: 14 }, { width: 14 }
  ];

  sheetTitle(ws, 7, 'Stakeholder Register');
  let r = 4;

  secTitle(ws, r++, 1, 7, 'Stakeholder Register');
  ['ID','Name','Title / Role','Organization','Responsibility','Influence','Interest']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.stakeholders || []).forEach((s, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, s.id,             bg, C.dark, true,  'center');
    cell(ws, r, 2, s.name,           bg, C.dark, true);
    cell(ws, r, 3, s.title,          bg, C.grey);
    cell(ws, r, 4, s.organization,   bg, C.dark);
    cell(ws, r, 5, s.responsibility, bg, C.dark);
    cell(ws, r, 6, s.influence, PRIORITY_COLOR[s.influence] || bg, C.dark, true, 'center');
    cell(ws, r, 7, s.interest,  PRIORITY_COLOR[s.interest]  || bg, C.dark, true, 'center');
    ws.getRow(r++).height = 25;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 7. BUSINESS REQUIREMENTS
// ══════════════════════════════════════════════════════════════════════════════
function buildBR(wb, d) {
  const ws = wb.addWorksheet('Business Requirements', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 12 }, { width: 22 }, { width: 52 },
    { width: 13 }, { width: 22 }, { width: 42 }, { width: 14 }
  ];

  sheetTitle(ws, 7, 'Business Requirements');
  let r = 4;
  secTitle(ws, r++, 1, 7, 'Business Requirements (BR)');
  ['BR ID','Category','Requirement Description','Priority','Source','Acceptance Criteria','Status']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.business_requirements || []).forEach((req, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, req.id,                  bg, C.dark, true,  'center');
    cell(ws, r, 2, req.category,            bg, C.dark, true);
    cell(ws, r, 3, req.description,         bg, C.dark);
    cell(ws, r, 4, req.priority,   PRIORITY_COLOR[req.priority] || bg, C.dark, true, 'center');
    cell(ws, r, 5, req.source,              bg, C.grey);
    cell(ws, r, 6, req.acceptance_criteria, bg, C.dark);
    cell(ws, r, 7, req.status || 'Draft',   C.highYellow, C.dark, false, 'center');
    ws.getRow(r++).height = 30;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 8. FUNCTIONAL REQUIREMENTS
// ══════════════════════════════════════════════════════════════════════════════
function buildFR(wb, d) {
  const ws = wb.addWorksheet('Functional Requirements', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 12 }, { width: 22 }, { width: 52 },
    { width: 13 }, { width: 14 }, { width: 22 }, { width: 14 }
  ];

  sheetTitle(ws, 7, 'Functional Requirements');
  let r = 4;
  secTitle(ws, r++, 1, 7, 'Functional Requirements (FR)');
  ['FR ID','Module / Feature','Requirement Description','Priority','BR Reference','Dependency','Status']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.functional_requirements || []).forEach((req, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, req.id,           bg, C.dark, true,  'center');
    cell(ws, r, 2, req.module,       bg, C.dark, true);
    cell(ws, r, 3, req.description,  bg, C.dark);
    cell(ws, r, 4, req.priority,     PRIORITY_COLOR[req.priority] || bg, C.dark, true, 'center');
    cell(ws, r, 5, req.br_reference, C.altRow, C.blueBg, true, 'center');
    cell(ws, r, 6, req.dependency,   bg, C.grey);
    cell(ws, r, 7, req.status || 'Draft', C.highYellow, C.dark, false, 'center');
    ws.getRow(r++).height = 30;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 9. NON-FUNCTIONAL REQUIREMENTS
// ══════════════════════════════════════════════════════════════════════════════
function buildNFR(wb, d) {
  const ws = wb.addWorksheet('Non-Functional Req', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 12 }, { width: 22 }, { width: 52 }, { width: 13 }, { width: 42 }, { width: 14 }
  ];

  sheetTitle(ws, 6, 'Non-Functional Requirements');
  let r = 4;
  secTitle(ws, r++, 1, 6, 'Non-Functional Requirements (NFR)');
  ['NFR ID','Category','Requirement Description','Priority','Acceptance Criteria','Status']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.non_functional_requirements || []).forEach((req, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, req.id,                   bg, C.dark, true,  'center');
    cell(ws, r, 2, req.category,             C.labelBg, C.dark, true);
    cell(ws, r, 3, req.description,          bg, C.dark);
    cell(ws, r, 4, req.priority,  PRIORITY_COLOR[req.priority] || bg, C.dark, true, 'center');
    cell(ws, r, 5, req.acceptance_criteria,  bg, C.dark);
    cell(ws, r, 6, req.status || 'Draft',    C.highYellow, C.dark, false, 'center');
    ws.getRow(r++).height = 30;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 10. ASSUMPTIONS & CONSTRAINTS
// ══════════════════════════════════════════════════════════════════════════════
function buildAssumptions(wb, d) {
  const ws = wb.addWorksheet('Assumptions & Constraints', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 12 }, { width: 18 }, { width: 52 }, { width: 36 }, { width: 20 }];

  sheetTitle(ws, 5, 'Assumptions & Constraints');
  let r = 4;

  secTitle(ws, r++, 1, 5, 'Assumptions');
  ['ID','Category','Assumption','Impact if Wrong','Owner']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.assumptions || []).forEach((a, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, a.id,              bg, C.dark, true,  'center');
    cell(ws, r, 2, a.category,        C.labelBg, C.dark, true);
    cell(ws, r, 3, a.assumption,      bg, C.dark);
    cell(ws, r, 4, a.impact_if_wrong, C.highOrange, C.dark);
    cell(ws, r, 5, a.owner,           bg, C.grey);
    ws.getRow(r++).height = 30;
  });

  r++;
  secTitle(ws, r++, 1, 5, 'Constraints');
  ['ID','Type','Constraint','Impact','Mitigation']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.constraints || []).forEach((c2, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, c2.id,         bg, C.dark, true,  'center');
    cell(ws, r, 2, c2.type,       C.labelBg, C.dark, true);
    cell(ws, r, 3, c2.constraint, bg, C.dark);
    cell(ws, r, 4, c2.impact,     C.highOrange, C.dark);
    cell(ws, r, 5, c2.mitigation, C.highGreen, C.dark);
    ws.getRow(r++).height = 30;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 11. RISK REGISTER
// ══════════════════════════════════════════════════════════════════════════════
function buildRisks(wb, d) {
  const ws = wb.addWorksheet('Risk Register', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 12 }, { width: 18 }, { width: 42 }, { width: 14 },
    { width: 14 }, { width: 14 }, { width: 40 }, { width: 22 }
  ];

  sheetTitle(ws, 8, 'Risk Register');
  let r = 4;
  secTitle(ws, r++, 1, 8, 'Risk Register');
  ['Risk ID','Category','Risk Description','Probability','Impact','Risk Level','Mitigation Strategy','Owner']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.risks || []).forEach((risk, idx) => {
    const bg   = idx % 2 === 0 ? C.white : C.altRow;
    const rlBg = RISK_COLOR[risk.risk_level] || RISK_COLOR[risk.impact] || bg;
    cell(ws, r, 1, risk.id,          bg,     C.dark, true,  'center');
    cell(ws, r, 2, risk.category,    C.labelBg, C.dark, true);
    cell(ws, r, 3, risk.description, bg,     C.dark);
    cell(ws, r, 4, risk.probability, RISK_COLOR[risk.probability] || bg, C.dark, true, 'center');
    cell(ws, r, 5, risk.impact,      RISK_COLOR[risk.impact]      || bg, C.dark, true, 'center');
    cell(ws, r, 6, risk.risk_level,  rlBg,   C.dark, true, 'center');
    cell(ws, r, 7, risk.mitigation,  bg,     C.dark);
    cell(ws, r, 8, risk.owner,       bg,     C.grey);
    ws.getRow(r++).height = 30;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 12. TIMELINE & MILESTONES
// ══════════════════════════════════════════════════════════════════════════════
function buildTimeline(wb, d) {
  const ws = wb.addWorksheet('Timeline & Milestones', { views: [{ showGridLines: false }] });
  ws.columns = [
    { width: 10 }, { width: 32 }, { width: 44 }, { width: 18 }, { width: 22 }, { width: 18 }
  ];

  sheetTitle(ws, 6, 'Timeline & Milestones');
  let r = 4;
  secTitle(ws, r++, 1, 6, 'Project Milestones');
  ['MS ID','Milestone','Description / Deliverable','Planned Date','Owner','Status']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.milestones || []).forEach((m, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, m.id,          bg, C.dark, true,  'center');
    cell(ws, r, 2, m.milestone,   bg, C.dark, true);
    cell(ws, r, 3, m.description, bg, C.dark);
    cell(ws, r, 4, m.planned_date, C.highYellow, C.dark, true, 'center');
    cell(ws, r, 5, m.owner,       bg, C.grey);
    cell(ws, r, 6, m.status || 'Not Started', C.altRow, C.dark, false, 'center');
    ws.getRow(r++).height = 25;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// 13. GLOSSARY
// ══════════════════════════════════════════════════════════════════════════════
function buildGlossary(wb, d) {
  const ws = wb.addWorksheet('Glossary', { views: [{ showGridLines: false }] });
  ws.columns = [{ width: 10 }, { width: 28 }, { width: 70 }, { width: 22 }];

  sheetTitle(ws, 4, 'Glossary of Terms & Abbreviations');
  let r = 4;
  ['#','Term / Abbreviation','Definition','Source']
    .forEach((h,i) => hdr(ws, r, i+1, h));
  ws.getRow(r++).height = 22;

  (d.glossary || []).forEach((g, idx) => {
    const bg = idx % 2 === 0 ? C.white : C.altRow;
    cell(ws, r, 1, String(idx + 1), bg, C.dark, true, 'center');
    cell(ws, r, 2, g.term,          bg, C.dark, true);
    cell(ws, r, 3, g.definition,    bg, C.dark);
    cell(ws, r, 4, g.source || '',  bg, C.grey);
    ws.getRow(r++).height = 22;
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// MAIN
// ══════════════════════════════════════════════════════════════════════════════
async function main() {
  const args    = process.argv.slice(2);
  const apiKey  = args.includes('--api-key') ? args[args.indexOf('--api-key') + 1] : null;
  const dataFile = args.includes('--data')   ? args[args.indexOf('--data')    + 1] : null;

  let brdData;

  if (dataFile) {
    // Load from external JSON file
    console.log(`Loading BRD data from: ${dataFile}`);
    brdData = JSON.parse(fs.readFileSync(dataFile, 'utf8'));
  } else {
    // Use built-in sample data
    const samplePath = path.join(__dirname, 'resources', 'mock_brd_response.json');
    console.log('Using built-in sample BRD data (Digital Onboarding Platform / Acme Bank)');
    brdData = JSON.parse(fs.readFileSync(samplePath, 'utf8'));
  }

  // Build the workbook
  const wb = new ExcelJS.Workbook();
  wb.creator  = 'BRD Generator Agent';
  wb.created  = new Date();
  wb.modified = new Date();

  console.log('Building BRD Excel document...');
  buildCover(wb, brdData);
  buildDocControl(wb, brdData);
  buildExecSummary(wb, brdData);
  buildObjectives(wb, brdData);
  buildScope(wb, brdData);
  buildStakeholders(wb, brdData);
  buildBR(wb, brdData);
  buildFR(wb, brdData);
  buildNFR(wb, brdData);
  buildAssumptions(wb, brdData);
  buildRisks(wb, brdData);
  buildTimeline(wb, brdData);
  buildGlossary(wb, brdData);

  // Save output
  const safeName  = (brdData.project_name || 'Project').replace(/[^a-zA-Z0-9_\- ]/g, '').trim().replace(/ /g, '_');
  const dateStamp = new Date().toISOString().slice(0,10).replace(/-/g,'');
  const fileName  = `${safeName}_BRD_${dateStamp}.xlsx`;
  const outPath   = path.join(__dirname, 'output', fileName);

  await wb.xlsx.writeFile(outPath);

  console.log('');
  console.log('✅ BRD Generated Successfully!');
  console.log(`📄 File: output/${fileName}`);
  console.log(`📊 Sheets: 13`);
  console.log(`📋 Contents:`);
  console.log(`   • Business Requirements : ${(brdData.business_requirements||[]).length} items`);
  console.log(`   • Functional Requirements: ${(brdData.functional_requirements||[]).length} items`);
  console.log(`   • Non-Functional Req    : ${(brdData.non_functional_requirements||[]).length} items`);
  console.log(`   • Stakeholders          : ${(brdData.stakeholders||[]).length} people`);
  console.log(`   • Risks                 : ${(brdData.risks||[]).length} items`);
  console.log(`   • Milestones            : ${(brdData.milestones||[]).length} items`);
  console.log(`   • Glossary terms        : ${(brdData.glossary||[]).length} items`);
}

main().catch(err => { console.error('Error:', err.message); process.exit(1); });
