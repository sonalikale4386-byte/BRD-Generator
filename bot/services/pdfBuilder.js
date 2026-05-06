'use strict';

const PDFDocument = require('pdfkit');
const path        = require('path');
const fs          = require('fs');

async function generatePDF(brdData, outputDir) {
  const reqs = brdData.requirements || [];

  return new Promise((resolve, reject) => {
    const safeProject = (brdData.project_name || 'BRD').replace(/[^a-zA-Z0-9_\- ]/g, '').trim().replace(/ /g, '_');
    const stamp       = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const ver         = (brdData.document_version || 'v1.0').toUpperCase();
    const fileName    = `${safeProject}_BRD_${stamp}_${ver}.pdf`;
    const filePath    = path.join(outputDir, fileName);

    const doc  = new PDFDocument({ margin: 50, size: 'A4' });
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);

    // ── Colours / helpers ───────────────────────────────────────────────────────
    const DARK_BLUE = '#002344';
    const MID_BLUE  = '#004080';
    const RED       = '#cc0000';
    const LIGHT_GRAY = '#f5f5f5';

    function heading(text, size = 16) {
      doc.moveDown(0.5)
         .fontSize(size).fillColor(DARK_BLUE).font('Helvetica-Bold')
         .text(text)
         .moveDown(0.3);
      doc.fillColor('black').font('Helvetica');
    }

    function subText(text, size = 10) {
      doc.fontSize(size).fillColor('#333333').font('Helvetica').text(String(text || ''), { lineGap: 3 });
    }

    function labelValue(label, value) {
      doc.fontSize(10).font('Helvetica-Bold').fillColor(DARK_BLUE)
         .text(label + ': ', { continued: true })
         .font('Helvetica').fillColor('black')
         .text(String(value || ''));
    }

    function tableHeader(cols, widths) {
      const x = doc.x;
      const y = doc.y;
      doc.rect(x, y, widths.reduce((a, b) => a + b, 0), 18).fill(DARK_BLUE);
      let cx = x;
      cols.forEach((col, i) => {
        doc.fillColor('white').font('Helvetica-Bold').fontSize(8)
           .text(col, cx + 3, y + 4, { width: widths[i] - 6, ellipsis: true });
        cx += widths[i];
      });
      doc.moveDown(0);
      doc.y = y + 20;
      doc.x = x;
      return y + 20;
    }

    function tableRow(cols, widths, shade = false) {
      const x = doc.x;
      const y = doc.y;
      const rowH = 18;
      if (shade) doc.rect(x, y, widths.reduce((a, b) => a + b, 0), rowH).fill(LIGHT_GRAY);
      let cx = x;
      cols.forEach((col, i) => {
        doc.fillColor('#222222').font('Helvetica').fontSize(8)
           .text(String(col || ''), cx + 3, y + 4, { width: widths[i] - 6, ellipsis: true });
        cx += widths[i];
      });
      doc.y = y + rowH;
      doc.x = x;
    }

    // ── Cover ───────────────────────────────────────────────────────────────────
    doc.rect(0, 0, doc.page.width, 120).fill(DARK_BLUE);
    doc.fontSize(22).fillColor('white').font('Helvetica-Bold')
       .text('CRM Business Requirements Document', 50, 35, { align: 'center' });
    doc.fontSize(14).fillColor('#aaccff')
       .text(brdData.project_name || '', 50, 70, { align: 'center' });
    doc.y = 140;

    // ── Document overview ───────────────────────────────────────────────────────
    heading('Document Overview');
    [
      ['Client',      brdData.client_name],
      ['Platform',    brdData.platform || 'Microsoft Dynamics 365 CRM'],
      ['Phase',       brdData.phase],
      ['Version',     brdData.document_version || 'v1.0'],
      ['Date',        brdData.document_date],
      ['Prepared By', brdData.prepared_by],
      ['Status',      brdData.status || 'Draft'],
    ].forEach(([l, v]) => labelValue(l, v));

    // ── Executive summary ───────────────────────────────────────────────────────
    heading('Executive Summary');
    subText(brdData.executive_summary || '');

    // ── Stakeholders ────────────────────────────────────────────────────────────
    heading('Key Stakeholders');
    const stakeW = [160, 170, 160];
    tableHeader(['Name', 'Role', 'Organization'], stakeW);
    (brdData.key_stakeholders || []).forEach((s, i) => {
      tableRow([s.name, s.role, s.organization], stakeW, i % 2 === 0);
    });

    // ── Requirements ────────────────────────────────────────────────────────────
    doc.addPage();
    heading('Requirements');
    const reqW = [45, 200, 55, 55, 80, 60];
    tableHeader(['ID', 'Description', 'Priority', 'Status', 'Scope', 'Have Today'], reqW);
    reqs.forEach((r, i) => {
      if (doc.y > doc.page.height - 80) doc.addPage();
      tableRow([r.id, r.description, r.priority, r.status, r.scope || '', r.have_today || ''], reqW, i % 2 === 0);
    });

    // ── Out of scope ────────────────────────────────────────────────────────────
    doc.addPage();
    heading('Out of Scope');
    (brdData.out_of_scope || []).forEach(o => subText(`• ${o.item}`));

    // ── Sign-off ────────────────────────────────────────────────────────────────
    heading('Sign Off');
    const signW = [60, 180, 120, 100];
    tableHeader(['Version', 'Name & Role', 'Signature', 'Date'], signW);
    (brdData.sign_off || []).forEach((s, i) => {
      tableRow([s.version, s.name_and_role, s.signature || '', s.date || ''], signW, i % 2 === 0);
    });

    doc.end();
    stream.on('finish', () => resolve(filePath));
    stream.on('error', reject);
  });
}

module.exports = { generatePDF };
