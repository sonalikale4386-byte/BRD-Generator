/**
 * fileExtractor.js
 * Downloads a file from a Teams attachment URL and extracts plain text.
 * Supports: PDF, DOCX, TXT
 */
const axios = require('axios');
const path  = require('path');

class FileExtractor {
  /**
   * @param {string} url       - Direct download URL from Teams attachment
   * @param {string} fileName  - Original filename (used to detect type)
   * @returns {Promise<string>} Extracted plain text
   */
  async extractText(url, fileName) {
    const ext = path.extname(fileName).toLowerCase();

    // Download file as binary buffer
    const response = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout: 30_000,
      // Teams CDN URLs contain auth tokens — no extra headers needed
      validateStatus: s => s < 400,
    });

    const buffer = Buffer.from(response.data);

    if (ext === '.txt') {
      return buffer.toString('utf8');
    }

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

    // Fallback: try utf8 decode
    return buffer.toString('utf8');
  }

  /**
   * Extract text directly from a Buffer (used for OneDrive downloads).
   * @param {Buffer} buffer
   * @param {string} fileName  Original filename — used to detect file type
   */
  async extractFromBuffer(buffer, fileName) {
    const ext = path.extname(fileName).toLowerCase();

    if (ext === '.txt') return buffer.toString('utf8');

    if (ext === '.pdf') {
      const pdfParse = require('pdf-parse');
      const data = await pdfParse(buffer);
      return data.text || '';
    }

    if (ext === '.docx' || ext === '.doc') {
      const mammoth = require('mammoth');
      const result  = await mammoth.extractRawText({ buffer });
      return result.value || '';
    }

    if (ext === '.xlsx' || ext === '.xls') {
      const ExcelJS = require('exceljs');
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      let text = '';
      wb.eachSheet(ws => {
        text += `\n=== Sheet: ${ws.name} ===\n`;
        ws.eachRow(row => {
          const vals = (row.values || []).slice(1)
            .filter(v => v !== null && v !== undefined && String(v).trim() !== '')
            .map(v => {
              if (typeof v === 'object' && v !== null) {
                return v.text || v.result || v.formula || JSON.stringify(v);
              }
              return String(v);
            });
          if (vals.length) text += vals.join(' | ') + '\n';
        });
      });
      return text.trim();
    }

    return buffer.toString('utf8');
  }
}

module.exports = { FileExtractor };
