/**
 * templateExtractor.js
 * Extracts structure and content from a user-uploaded BRD reference template.
 * Supports: .xlsx (extracts sheet names + headers), .pdf, .docx, .txt
 */
const axios   = require('axios');
const path    = require('path');
const ExcelJS = require('exceljs');

class TemplateExtractor {
  /**
   * Download and extract structure from a template file.
   * Returns a { type, structure, rawText, summary } object.
   */
  async extract(url, fileName) {
    const ext    = path.extname(fileName).toLowerCase();
    const buffer = await this._download(url);

    if (ext === '.xlsx' || ext === '.xls') {
      return await this._extractExcel(buffer, fileName);
    }

    // For PDF / DOCX / TXT — extract raw text as structure guide
    const text = await this._extractText(buffer, ext);
    return {
      type: 'document',
      fileName,
      rawText: text.slice(0, 6000), // cap to avoid token overload
      summary: `Reference template (${ext}) with ${text.split('\n').length} lines of content.`,
      structure: null,
    };
  }

  // ── Excel template: extract sheet names + column headers ──────────────────
  async _extractExcel(buffer, fileName) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const sheets = [];
    wb.worksheets.forEach(ws => {
      const headers = [];
      // Scan first 5 rows to find header row (first non-empty row)
      for (let r = 1; r <= 5; r++) {
        const row = ws.getRow(r);
        const vals = [];
        row.eachCell({ includeEmpty: false }, cell => {
          const v = cell.value;
          if (v && typeof v === 'string' && v.trim()) vals.push(v.trim());
          else if (v && typeof v !== 'object') vals.push(String(v).trim());
        });
        if (vals.length > 1) { // found a row with multiple columns = header
          headers.push(...vals);
          break;
        }
      }

      // Collect sample data from first 3 data rows
      const sampleRows = [];
      for (let r = 2; r <= 4; r++) {
        const row = ws.getRow(r);
        const vals = [];
        row.eachCell({ includeEmpty: false }, cell => {
          const v = cell.value;
          if (v && typeof v !== 'object') vals.push(String(v).slice(0, 80));
        });
        if (vals.length) sampleRows.push(vals);
      }

      if (ws.name && headers.length) {
        sheets.push({ sheetName: ws.name, headers, sampleRows });
      }
    });

    const summary = sheets.map(s =>
      `Sheet "${s.sheetName}": columns → ${s.headers.join(' | ')}`
    ).join('\n');

    return {
      type: 'excel',
      fileName,
      sheets,
      summary,
      rawText: null,
      structure: sheets,
    };
  }

  // ── Text extraction for PDF / DOCX / TXT ──────────────────────────────────
  async _extractText(buffer, ext) {
    if (ext === '.txt') return buffer.toString('utf8');
    if (ext === '.pdf') {
      const pdfParse = require('pdf-parse');
      const data = await pdfParse(buffer);
      return data.text || '';
    }
    if (ext === '.docx') {
      const mammoth = require('mammoth');
      const result  = await mammoth.extractRawText({ buffer });
      return result.value || '';
    }
    return buffer.toString('utf8');
  }

  // ── Download file from URL ─────────────────────────────────────────────────
  async _download(url) {
    const response = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout: 30_000,
      validateStatus: s => s < 400,
    });
    return Buffer.from(response.data);
  }

  /**
   * Build a concise template instruction string for the Claude prompt.
   */
  buildPromptSection(templateInfo) {
    if (!templateInfo) return '';

    if (templateInfo.type === 'excel' && templateInfo.sheets) {
      const sheetDetails = templateInfo.sheets.map(s => {
        const sample = s.sampleRows.length
          ? `\n     Sample values: ${s.sampleRows[0].slice(0,3).join(' | ')}`
          : '';
        return `  - Sheet: "${s.sheetName}"\n    Columns: ${s.headers.join(' | ')}${sample}`;
      }).join('\n');

      return `
REFERENCE TEMPLATE (user uploaded Excel BRD template):
The user has provided their own BRD Excel template. Match the structure and terminology below when generating content.

Template sheets and columns:
${sheetDetails}

Instructions:
- Use the same section names, column names, and terminology from the template above
- Maintain the same level of detail and format as the template structure suggests
- If a template section does not map to a standard BRD field, include it as an additional item
`;
    }

    if (templateInfo.type === 'document' && templateInfo.rawText) {
      return `
REFERENCE TEMPLATE (user uploaded ${templateInfo.fileName}):
The user has provided a reference BRD document. Use its structure, sections, and terminology as a guide.

Template content:
${templateInfo.rawText}

Instructions:
- Follow the section structure and terminology from the reference document above
- Use the same level of detail and format as the reference
`;
    }

    return '';
  }
}

module.exports = { TemplateExtractor };
