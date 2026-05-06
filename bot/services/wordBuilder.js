'use strict';

const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, WidthType, BorderStyle, AlignmentType,
  ShadingType,
} = require('docx');
const path = require('path');
const fs   = require('fs');

async function generateWord(brdData, outputDir) {
  const reqs = brdData.requirements || [];
  const date = brdData.document_date || new Date().toLocaleDateString();

  // ── Helpers ──────────────────────────────────────────────────────────────────
  const h1 = (text) => new Paragraph({
    text, heading: HeadingLevel.HEADING_1,
    spacing: { before: 300, after: 150 },
  });

  const h2 = (text) => new Paragraph({
    text, heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
  });

  const para = (text) => new Paragraph({
    children: [new TextRun({ text: String(text || ''), size: 22 })],
    spacing: { after: 100 },
  });

  const bold = (text) => new TextRun({ text: String(text || ''), bold: true, size: 22 });

  const cell = (text, isHeader = false) => new TableCell({
    children: [new Paragraph({
      children: [new TextRun({ text: String(text || ''), bold: isHeader, size: isHeader ? 22 : 20 })],
      alignment: AlignmentType.LEFT,
    })],
    shading: isHeader ? { type: ShadingType.SOLID, color: '002344', fill: '002344' } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
  });

  const headerRow = (cols) => new TableRow({
    children: cols.map(c => {
      const tc = cell(c, true);
      tc.options = tc.options || {};
      return new TableCell({
        children: [new Paragraph({
          children: [new TextRun({ text: String(c), bold: true, size: 22, color: 'FFFFFF' })],
        })],
        shading: { type: ShadingType.SOLID, color: '002344', fill: '002344' },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
      });
    }),
    tableHeader: true,
  });

  const dataRow = (cols) => new TableRow({
    children: cols.map(c => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: String(c || ''), size: 20 })],
      })],
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
    })),
  });

  // ── Cover page ────────────────────────────────────────────────────────────────
  const coverSection = [
    new Paragraph({ spacing: { before: 1000 } }),
    new Paragraph({
      children: [new TextRun({ text: 'CRM Business Requirements Document (BRD)', bold: true, size: 52, color: '002344' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: brdData.project_name || '', size: 36, color: '004080' })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({ spacing: { before: 400 } }),
  ];

  // ── Document overview ─────────────────────────────────────────────────────────
  const overviewTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      dataRow(['Client',   brdData.client_name     || '']),
      dataRow(['Platform', brdData.platform        || 'Microsoft Dynamics 365 CRM']),
      dataRow(['Phase',    brdData.phase           || '']),
      dataRow(['Version',  brdData.document_version || 'v1.0']),
      dataRow(['Date',     date]),
      dataRow(['Prepared By', brdData.prepared_by  || '']),
      dataRow(['Status',   brdData.status          || 'Draft']),
    ],
  });

  // ── Executive summary ─────────────────────────────────────────────────────────
  const execSection = [
    h1('Executive Summary'),
    para(brdData.executive_summary || ''),
  ];

  // ── Stakeholders table ────────────────────────────────────────────────────────
  const stakeholders = brdData.key_stakeholders || [];
  const stakeSection = [
    h1('Key Stakeholders'),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        headerRow(['Name', 'Role', 'Organization']),
        ...stakeholders.map(s => dataRow([s.name, s.role, s.organization])),
      ],
    }),
  ];

  // ── Requirements table ────────────────────────────────────────────────────────
  const reqSection = [
    h1('Requirements'),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        headerRow(['ID', 'Description', 'Priority', 'Status', 'Scope', 'Have Today']),
        ...reqs.map(r => dataRow([
          r.id, r.description, r.priority, r.status, r.scope || '', r.have_today || '',
        ])),
      ],
    }),
  ];

  // ── Out of scope ──────────────────────────────────────────────────────────────
  const outOfScope = brdData.out_of_scope || [];
  const oosSection = [
    h1('Out of Scope'),
    ...outOfScope.map(o => new Paragraph({
      children: [new TextRun({ text: `• ${o.item}`, size: 22 })],
      spacing: { after: 80 },
    })),
  ];

  // ── Sign-off ──────────────────────────────────────────────────────────────────
  const signOff = brdData.sign_off || [];
  const signSection = [
    h1('Sign Off'),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        headerRow(['Version', 'Name & Role', 'Signature', 'Date']),
        ...signOff.map(s => dataRow([s.version, s.name_and_role, s.signature || '', s.date || ''])),
      ],
    }),
  ];

  // ── Assemble document ─────────────────────────────────────────────────────────
  const doc = new Document({
    sections: [{
      children: [
        ...coverSection,
        h1('Document Overview'),
        overviewTable,
        new Paragraph({ spacing: { after: 200 } }),
        ...execSection,
        ...stakeSection,
        new Paragraph({ spacing: { after: 200 } }),
        ...reqSection,
        new Paragraph({ spacing: { after: 200 } }),
        ...oosSection,
        new Paragraph({ spacing: { after: 200 } }),
        ...signSection,
      ],
    }],
  });

  const safeProject = (brdData.project_name || 'BRD').replace(/[^a-zA-Z0-9_\- ]/g, '').trim().replace(/ /g, '_');
  const stamp       = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const ver         = (brdData.document_version || 'v1.0').toUpperCase();
  const fileName    = `${safeProject}_BRD_${stamp}_${ver}.docx`;
  const filePath    = path.join(outputDir, fileName);

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buffer);
  return filePath;
}

module.exports = { generateWord };
