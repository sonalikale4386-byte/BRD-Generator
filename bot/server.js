/**
 * server.js — BRD Generator Standalone Web App
 *
 * No Azure Bot Service. No Teams. Pure Express + Claude AI + ExcelJS.
 *
 * Endpoints:
 *   GET  /                              → Web UI
 *   POST /api/generate-brd              → Upload files → generate BRD
 *   POST /api/generate-brd-sharepoint   → Fetch from SharePoint → generate BRD
 *   GET  /api/sharepoint/status         → Is SharePoint configured?
 *   POST /api/sharepoint/list-files     → List files in a SharePoint folder
 *   GET  /output/:file                  → Download generated Excel
 *   GET  /health                        → Health check
 */
// Load .env — prefer APP_ENV_PATH (set by Electron), fall back to project root
const _envPath = process.env.APP_ENV_PATH || require('path').join(__dirname, '..', '.env');
require('dotenv').config({ path: _envPath });

const express = require('express');
const multer  = require('multer');
const axios   = require('axios');
const path    = require('path');
const fs      = require('fs');

const { ClaudeService } = require('./services/claudeService');
const { FileExtractor } = require('./services/fileExtractor');
const { generateExcel } = require('./brdCore');
const { generateWord  } = require('./services/wordBuilder');
const { generatePDF   } = require('./services/pdfBuilder');

// ── Config ─────────────────────────────────────────────────────────────────────
const PORT    = process.env.PORT || 3978;
// OUTPUT_DIR set by Electron to a writable user path (Documents); fallback for dev
const OUTPUT  = process.env.OUTPUT_DIR || path.join(__dirname, '..', 'output');
const WEB_DIR = path.join(__dirname, '..', 'web');

// SharePoint / Azure AD (App-Only, client credentials)
const SP_TENANT_ID     = process.env.AZURE_TENANT_ID     || '';
const SP_CLIENT_ID     = process.env.AZURE_CLIENT_ID     || '';
const SP_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || '';
const SP_CONFIGURED    = !!(SP_TENANT_ID && SP_CLIENT_ID && SP_CLIENT_SECRET);

if (!fs.existsSync(OUTPUT)) fs.mkdirSync(OUTPUT, { recursive: true });

// ── App setup ──────────────────────────────────────────────────────────────────
const app = express();

// Allow cross-origin requests (needed when accessed via tunnel or cloud URL)
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

// Bypass localtunnel splash page for API calls
app.use((req, res, next) => {
  if (req.path.startsWith('/api') || req.path.startsWith('/output')) {
    res.setHeader('bypass-tunnel-reminder', 'true');
  }
  next();
});

app.use(express.json());

const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 20 * 1024 * 1024, files: 10 },
  fileFilter: (_req, file, cb) => {
    const allowed = ['.pdf', '.docx', '.txt', '.doc', '.xlsx', '.xls'];
    cb(null, allowed.includes(path.extname(file.originalname).toLowerCase()));
  },
});

app.use(express.static(WEB_DIR));
app.use('/output', express.static(OUTPUT));

// ── SharePoint helpers ─────────────────────────────────────────────────────────

/** Get a Graph API access token using Azure AD client credentials (app-only) */
async function getGraphToken() {
  try {
    const res = await axios.post(
      `https://login.microsoftonline.com/${SP_TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        grant_type:    'client_credentials',
        client_id:     SP_CLIENT_ID,
        client_secret: SP_CLIENT_SECRET,
        scope:         'https://graph.microsoft.com/.default',
      }).toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    return res.data.access_token;
  } catch (err) {
    // Extract the real Azure AD error message
    const aadErr  = err.response?.data;
    const errCode = aadErr?.error         || 'token_error';
    const errDesc = aadErr?.error_description?.split('\r\n')[0] || err.message;
    console.error('❌ Azure AD token error:', errCode, '—', errDesc);
    throw new Error(`Azure AD authentication failed: ${errCode} — ${errDesc}`);
  }
}

/** Create an axios instance authenticated with the Graph token */
function graphHttp(token) {
  return axios.create({
    baseURL: 'https://graph.microsoft.com/v1.0',
    headers: { Authorization: `Bearer ${token}` },
  });
}

/** Parse a SharePoint URL → { hostname, sitePath } */
function parseSPUrl(siteUrl) {
  const u = new URL(siteUrl.trim().replace(/\/$/, ''));
  return { hostname: u.hostname, sitePath: u.pathname };
}

/** Extract text from a Buffer using FileExtractor */
async function extractBuffer(extractor, buffer, fileName) {
  try {
    const text = await extractor.extractFromBuffer(buffer, fileName);
    // Allow up to 20 000 chars per file (≈ 5 000 words) — enough for full transcripts
    return text ? `\n\n=== ${fileName} ===\n${text.slice(0, 20000)}` : '';
  } catch (err) {
    console.warn(`  ⚠️  Extract failed for ${fileName}:`, err.message);
    return '';
  }
}

/** Generate output file in requested format */
async function generateOutput(brdData, outputDir, format) {
  if (format === 'word') return generateWord(brdData, outputDir);
  if (format === 'pdf')  return generatePDF(brdData, outputDir);
  return generateExcel(brdData, outputDir);
}

const LAST_BRD_PATH  = path.join(OUTPUT, 'last_brd.json');
const SESSIONS_DIR   = path.join(OUTPUT, 'sessions');
const USAGE_LOG_PATH = path.join(OUTPUT, 'usage_log.json');
if (!fs.existsSync(SESSIONS_DIR)) fs.mkdirSync(SESSIONS_DIR, { recursive: true });

// Claude Sonnet 4.6 pricing (USD per 1M tokens)
const PRICE_INPUT_PER_M  = 3.00;
const PRICE_OUTPUT_PER_M = 15.00;

function calcCost(input_tokens, output_tokens) {
  const input_cost  = (input_tokens  / 1_000_000) * PRICE_INPUT_PER_M;
  const output_cost = (output_tokens / 1_000_000) * PRICE_OUTPUT_PER_M;
  return parseFloat((input_cost + output_cost).toFixed(6));
}

/** Append one token-usage entry to usage_log.json */
function logUsage(projectName, usage) {
  try {
    let log = [];
    if (fs.existsSync(USAGE_LOG_PATH)) {
      try { log = JSON.parse(fs.readFileSync(USAGE_LOG_PATH, 'utf8')); } catch {}
    }
    const input_tokens  = usage.input_tokens  || 0;
    const output_tokens = usage.output_tokens || 0;
    const cost_usd      = calcCost(input_tokens, output_tokens);
    log.push({
      timestamp:     new Date().toISOString(),
      project:       projectName,
      input_tokens,
      output_tokens,
      total_tokens:  input_tokens + output_tokens,
      cost_usd,
    });
    fs.writeFileSync(USAGE_LOG_PATH, JSON.stringify(log, null, 2), 'utf8');
    const totals = log.reduce((a, e) => ({
      i: a.i + e.input_tokens, o: a.o + e.output_tokens, c: a.c + e.cost_usd,
    }), { i: 0, o: 0, c: 0 });
    console.log(`  📈 Cumulative — input: ${totals.i.toLocaleString()}, output: ${totals.o.toLocaleString()} tokens | cost: $${totals.c.toFixed(4)} across ${log.length} BRD(s)`);
  } catch (e) { console.warn('⚠️  Could not write usage_log:', e.message); }
}

function saveSession(data) {
  try {
    const id   = `sess_${Date.now()}`;
    const file = path.join(SESSIONS_DIR, `${id}.json`);
    fs.writeFileSync(file, JSON.stringify({ id, ...data }, null, 2), 'utf8');
    return id;
  } catch (e) { console.warn('⚠️  Could not save session:', e.message); return null; }
}

/**
 * Compute the next sequential version for a project.
 * New BRD always starts at V1.0.
 * Each update increments the major number: V1.0 → V2.0 → V3.0 …
 */
function getNextVersion(projectName, brdType) {
  if (brdType !== 'update') return 'V1.0';
  try {
    const allFiles = fs.readdirSync(SESSIONS_DIR)
      .filter(f => f.endsWith('.json')).sort().reverse();
    const projectKey = (projectName || '').toLowerCase().trim();
    let maxVer = 0;
    for (const f of allFiles) {
      try {
        const s = JSON.parse(fs.readFileSync(path.join(SESSIONS_DIR, f), 'utf8'));
        if ((s.projectName || '').toLowerCase().trim() !== projectKey) continue;
        const raw   = s.version || s.summary?.version || 'V1.0';
        const match = raw.match(/(\d+)/);
        if (match) maxVer = Math.max(maxVer, parseInt(match[1]));
      } catch { /* skip corrupt file */ }
    }
    return `V${maxVer + 1}.0`;
  } catch {
    return 'V2.0';
  }
}

/** Persist the latest BRD JSON so it can be reloaded for future updates */
function saveLastBRD(brdData) {
  try {
    fs.writeFileSync(LAST_BRD_PATH, JSON.stringify(brdData, null, 2), 'utf8');
  } catch (e) {
    console.warn('⚠️  Could not save last_brd.json:', e.message);
  }
}

/** Load the previously saved BRD JSON, or null if none exists */
function loadLastBRD() {
  try {
    if (fs.existsSync(LAST_BRD_PATH)) {
      return JSON.parse(fs.readFileSync(LAST_BRD_PATH, 'utf8'));
    }
  } catch (e) {
    console.warn('⚠️  Could not load last_brd.json:', e.message);
  }
  return null;
}

/** Build a BRD response object from brdData + filePath */
function buildResponse(brdData, filePath) {
  const fileName    = path.basename(filePath);
  const reqs        = brdData.requirements || [];
  const u           = brdData._usage || {};
  const inputTok    = u.input_tokens  || 0;
  const outputTok   = u.output_tokens || 0;
  return {
    success:     true,
    downloadUrl: `/output/${encodeURIComponent(fileName)}`,
    fileName,
    usage: {
      input_tokens:  inputTok,
      output_tokens: outputTok,
      total_tokens:  inputTok + outputTok,
      cost_usd:      calcCost(inputTok, outputTok),
    },
    summary: {
      projectName: brdData.project_name,
      date:        brdData.document_date,
      version:     brdData.document_version,
      status:      brdData.status,
      total:       reqs.length,
      mustHave:    reqs.filter(r => r.priority === 'Must Have').length,
      shouldHave:  reqs.filter(r => r.priority === 'Should Have').length,
      couldHave:   reqs.filter(r => r.priority === 'Could Have').length,
      wontHave:    reqs.filter(r => r.priority === "Won't Have").length,
    },
    requirements: reqs.map(r => ({
      id:         r.id,
      description:r.description,
      priority:   r.priority,
      status:     r.status,
      scope:      r.scope       || '',
      requester:  r.requester   || '',
      have_today: r.have_today  || '',
    })),
  };
}

// ── GET /api/sharepoint/status ─────────────────────────────────────────────────
app.get('/api/sharepoint/status', (_req, res) => {
  res.json({
    configured: SP_CONFIGURED,
    tenantId:   SP_TENANT_ID   ? SP_TENANT_ID.slice(0, 8)   + '...' : '(not set)',
    clientId:   SP_CLIENT_ID   ? SP_CLIENT_ID.slice(0, 8)   + '...' : '(not set)',
    message: SP_CONFIGURED
      ? 'SharePoint integration is ready.'
      : 'Not configured — set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in .env',
  });
});

// ── GET /api/sharepoint/test-connection ───────────────────────────────────────
// Validates Azure AD credentials by requesting a token (no Graph call needed)
app.get('/api/sharepoint/test-connection', async (_req, res) => {
  if (!SP_CONFIGURED) {
    return res.status(503).json({ ok: false, error: 'Credentials not set in .env' });
  }
  try {
    const token = await getGraphToken();
    // Decode token payload to confirm tenant
    const payload = JSON.parse(Buffer.from(token.split('.')[1], 'base64').toString());
    res.json({
      ok:       true,
      message:  'Azure AD token obtained successfully.',
      appId:    payload.appid || payload.azp || '(unknown)',
      tenantId: payload.tid  || '(unknown)',
      expires:  new Date(payload.exp * 1000).toISOString(),
    });
  } catch (err) {
    res.status(401).json({ ok: false, error: err.message });
  }
});

// ── POST /api/sharepoint/list-files ───────────────────────────────────────────
app.post('/api/sharepoint/list-files', async (req, res) => {
  if (!SP_CONFIGURED) {
    return res.status(503).json({
      error: 'SharePoint not configured. See .env for AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET.',
    });
  }

  const { siteUrl, folderPath = 'BRD Document' } = req.body;
  if (!siteUrl) return res.status(400).json({ error: 'siteUrl is required.' });

  try {
    const token = await getGraphToken();
    const http  = graphHttp(token);

    // Get SharePoint site ID
    const { hostname, sitePath } = parseSPUrl(siteUrl);
    const siteRes = await http.get(`/sites/${hostname}:${sitePath}`);
    const siteId  = siteRes.data.id;

    // Find the Documents library
    const drivesRes = await http.get(`/sites/${siteId}/drives`);
    const drive = drivesRes.data.value.find(
      d => ['Documents', 'Shared Documents', 'documents'].includes(d.name)
    ) || drivesRes.data.value[0];

    if (!drive) {
      return res.status(404).json({ error: 'No document library found on this site.' });
    }

    // List files in folder
    // Note: @microsoft.graph.downloadUrl must NOT be in $select — it's an OData annotation
    // that Graph returns automatically. Including it in $select causes generalException.
    let filesData;
    try {
      const encodedFolder = folderPath.split('/').map(p => encodeURIComponent(p)).join('/');
      const r = await http.get(
        `/drives/${drive.id}/root:/${encodedFolder}:/children`,
        { params: { $select: 'id,name,size,lastModifiedDateTime,file,folder' } }
      );
      filesData = r.data.value || [];
    } catch (err) {
      if (err.response?.status === 404) {
        return res.status(404).json({
          error: `Folder "${folderPath}" not found. Check the folder name or create it in SharePoint.`,
        });
      }
      const ge = err.response?.data?.error;
      if (ge) throw new Error(`${ge.code}: ${ge.message}`);
      throw err;
    }

    const SUPPORTED = ['.pdf', '.docx', '.txt', '.doc'];
    const files = filesData
      .filter(f => !f.folder)
      .filter(f => SUPPORTED.includes(path.extname(f.name || '').toLowerCase()))
      .map(f => ({
        id:           f.id,
        name:         f.name,
        size:         f.size,
        modified:     f.lastModifiedDateTime,
        downloadUrl:  f['@microsoft.graph.downloadUrl'],
        driveId:      drive.id,
        siteId,
      }));

    console.log(`📂 SharePoint: listed ${files.length} file(s) from "${folderPath}" @ ${siteUrl}`);
    res.json({ success: true, files, folderPath });

  } catch (err) {
    const graphErr = err.response?.data?.error;
    const msg = graphErr
      ? `${graphErr.code}: ${graphErr.message}`
      : err.message;
    console.error('❌ SharePoint list error:', msg);
    res.status(500).json({ error: msg });
  }
});

// ── POST /api/generate-brd  (Upload / Paste source) ──────────────────────────
app.post('/api/generate-brd', upload.array('files', 10), async (req, res) => {
  try {
    const projectName      = (req.body.projectName      || '').trim();
    const authorName       = (req.body.authorName       || 'BRD Generator').trim();
    const brdType          = (req.body.brdType          || 'new').trim();
    const updateMethod     = (req.body.updateMethod     || '').trim();
    const detailLevel      = (req.body.detailLevel      || 'elaborated').trim();
    const fitGap           = (req.body.fitGap           || 'no').trim();
    const sourceRef        = (req.body.sourceRef        || 'yes').trim();
    const moscow           = (req.body.moscow           || 'yes').trim();
    const additionalInputs = (req.body.additionalInputs || '').trim();
    const outputFormat     = (req.body.outputFormat     || 'excel').trim();
    const pasteContent     = (req.body.pasteContent     || '').trim();

    if (!projectName) return res.status(400).json({ error: 'Project name is required.' });

    const files = req.files || [];
    console.log(`\n📋 [Upload] Generating BRD — "${projectName}"  Files: ${files.length}  Format: ${outputFormat}`);

    const extractor = new FileExtractor();
    let extractedText = '';

    // Include pasted content as a document
    if (pasteContent) {
      extractedText += `\n\n=== Pasted Content ===\n${pasteContent.slice(0, 20000)}`;
      console.log(`  📋 Pasted content — ${pasteContent.length} chars`);
    }

    for (const file of files) {
      const chunk = await extractBuffer(extractor, file.buffer, file.originalname);
      extractedText += chunk;
      console.log(`  📄 ${file.originalname} — ${chunk.length} chars extracted`);
    }

    if (extractedText.length > 60000) {
      console.warn(`  ⚠️  Total extracted text ${extractedText.length} chars — truncating to 60 000`);
      extractedText = extractedText.slice(0, 60000);
    }
    console.log(`  📝 Total extracted: ${extractedText.length} chars`);

    // Load previous BRD for update scenarios
    let previousBRD = null;
    if (brdType === 'update' && updateMethod === 'latest') {
      previousBRD = loadLastBRD();
      if (previousBRD) {
        console.log(`  🔄 Loaded previous BRD: "${previousBRD.project_name}" (${previousBRD.requirements?.length || 0} requirements)`);
      } else {
        console.warn('  ⚠️  No previous BRD found — generating fresh');
      }
    }

    const claude  = new ClaudeService();
    const brdData = await claude.generateBRD({
      projectName, userName: authorName, extractedText,
      brdType, updateMethod, detailLevel, fitGap, sourceRef, moscow, additionalInputs,
      previousBRD,
    });

    // Override version with our sequential numbering (V1.0, V2.0, V3.0 …)
    brdData.document_version = getNextVersion(projectName, brdType);

    let chatLog = [];
    try { chatLog = JSON.parse(req.body.chatLog || '[]'); } catch { chatLog = []; }

    const usage = brdData._usage || {};
    logUsage(brdData.project_name, usage);

    saveLastBRD(brdData);
    const filePath  = await generateOutput(brdData, OUTPUT, outputFormat);
    const response  = buildResponse(brdData, filePath);
    saveLastBRD(brdData);
    saveSession({
      projectName: brdData.project_name, authorName, brdType, outputFormat,
      version:     brdData.document_version || 'v1.0',
      createdAt:   new Date().toISOString(),
      fileName:    response.fileName,
      downloadUrl: response.downloadUrl,
      summary:     response.summary,
      usage:       { input_tokens: usage.input_tokens || 0, output_tokens: usage.output_tokens || 0 },
      chatLog,
    });
    console.log(`  ✅ Done — ${path.basename(filePath)}`);
    res.json(response);

  } catch (err) {
    console.error('❌ Generate BRD (upload) error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/generate-brd-sharepoint  (SharePoint source) ────────────────────
app.post('/api/generate-brd-sharepoint', async (req, res) => {
  if (!SP_CONFIGURED) {
    return res.status(503).json({
      error: 'SharePoint not configured. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in .env',
    });
  }

  const {
    projectName, authorName = 'BRD Generator', files = [],
    brdType = 'new', updateMethod = '', detailLevel = 'elaborated',
    fitGap = 'no', sourceRef = 'yes', moscow = 'yes',
    additionalInputs = '', outputFormat = 'excel',
  } = req.body;
  if (!projectName) return res.status(400).json({ error: 'Project name is required.' });
  if (!files.length) return res.status(400).json({ error: 'No files selected from SharePoint.' });

  try {
    console.log(`\n📋 [SharePoint] Generating BRD — "${projectName}"  Files: ${files.length}  Format: ${outputFormat}`);

    const extractor = new FileExtractor();
    let extractedText = '';

    // Get a fresh Graph token once for all file downloads
    const dlToken = await getGraphToken();

    for (const file of files) {
      try {
        let resp;
        if (file.downloadUrl) {
          // Direct pre-auth URL (no token needed, short-lived)
          resp = await axios.get(file.downloadUrl, { responseType: 'arraybuffer', timeout: 30_000 });
        } else {
          // Fallback: download via Graph API using driveId + fileId
          resp = await axios.get(
            `https://graph.microsoft.com/v1.0/drives/${file.driveId}/items/${file.id}/content`,
            { headers: { Authorization: `Bearer ${dlToken}` }, responseType: 'arraybuffer', timeout: 30_000 }
          );
        }
        const buf   = Buffer.from(resp.data);
        const chunk = await extractBuffer(extractor, buf, file.name);
        extractedText += chunk;
        console.log(`  ✅ Fetched: ${file.name} — ${chunk.length} chars`);
      } catch (err) {
        console.warn(`  ⚠️  Fetch failed for ${file.name}:`, err.message);
      }
    }

    if (extractedText.length > 60000) {
      console.warn(`  ⚠️  Total extracted text ${extractedText.length} chars — truncating to 60 000`);
      extractedText = extractedText.slice(0, 60000);
    }
    console.log(`  📝 Total extracted: ${extractedText.length} chars across ${files.length} file(s)`);

    let previousBRD = null;
    if (brdType === 'update' && updateMethod === 'latest') {
      previousBRD = loadLastBRD();
      if (previousBRD) {
        console.log(`  🔄 Loaded previous BRD: "${previousBRD.project_name}" (${previousBRD.requirements?.length || 0} requirements)`);
      }
    }

    const claude   = new ClaudeService();
    const brdData  = await claude.generateBRD({
      projectName, userName: authorName, extractedText,
      brdType, updateMethod, detailLevel, fitGap, sourceRef, moscow, additionalInputs,
      previousBRD,
    });

    // Override version with our sequential numbering (V1.0, V2.0, V3.0 …)
    brdData.document_version = getNextVersion(projectName, brdType);

    const chatLog = Array.isArray(req.body.chatLog) ? req.body.chatLog : [];

    const spUsage = brdData._usage || {};
    logUsage(brdData.project_name, spUsage);

    saveLastBRD(brdData);
    const filePath  = await generateOutput(brdData, OUTPUT, outputFormat);
    const response  = buildResponse(brdData, filePath);
    saveLastBRD(brdData);
    saveSession({
      projectName: brdData.project_name, authorName, brdType, outputFormat,
      version:     brdData.document_version || 'v1.0',
      createdAt:   new Date().toISOString(),
      fileName:    response.fileName,
      downloadUrl: response.downloadUrl,
      summary:     response.summary,
      usage:       { input_tokens: spUsage.input_tokens || 0, output_tokens: spUsage.output_tokens || 0 },
      chatLog,
    });
    console.log(`  ✅ Done — ${path.basename(filePath)}`);
    res.json(response);

  } catch (err) {
    console.error('❌ Generate BRD (SharePoint) error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/sessions ──────────────────────────────────────────────────────────
app.get('/api/sessions', (_req, res) => {
  try {
    // Load newest-first, then deduplicate by project name keeping only the latest
    const all = fs.readdirSync(SESSIONS_DIR)
      .filter(f => f.endsWith('.json'))
      .sort().reverse().slice(0, 500)
      .map(f => { try { return JSON.parse(fs.readFileSync(path.join(SESSIONS_DIR, f), 'utf8')); } catch { return null; } })
      .filter(Boolean);

    const seen = new Set();
    const sessions = all.filter(s => {
      const key = (s.projectName || '').toLowerCase().trim();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    }).slice(0, 50);

    res.json({ sessions });
  } catch { res.json({ sessions: [] }); }
});

// ── GET /api/usage ─────────────────────────────────────────────────────────────
app.get('/api/usage', (_req, res) => {
  try {
    if (!fs.existsSync(USAGE_LOG_PATH)) {
      return res.json({ entries: [], totals: { input_tokens: 0, output_tokens: 0, total_tokens: 0, cost_usd: 0, brd_count: 0 } });
    }
    const entries = JSON.parse(fs.readFileSync(USAGE_LOG_PATH, 'utf8'));
    const totals  = entries.reduce((a, e) => ({
      input_tokens:  a.input_tokens  + e.input_tokens,
      output_tokens: a.output_tokens + e.output_tokens,
      total_tokens:  a.total_tokens  + e.total_tokens,
      cost_usd:      parseFloat((a.cost_usd + (e.cost_usd || calcCost(e.input_tokens, e.output_tokens))).toFixed(6)),
      brd_count:     a.brd_count + 1,
    }), { input_tokens: 0, output_tokens: 0, total_tokens: 0, cost_usd: 0, brd_count: 0 });
    res.json({ entries, totals, pricing: { model: 'claude-sonnet-4-6', input_per_1m: PRICE_INPUT_PER_M, output_per_1m: PRICE_OUTPUT_PER_M, currency: 'USD' } });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── POST /api/chat ─────────────────────────────────────────────────────────────
app.post('/api/chat', async (req, res) => {
  const { messages = [], brdContext = null } = req.body;
  if (!messages.length) return res.status(400).json({ error: 'messages is required.' });

  const claude = new ClaudeService();
  if (claude.useMock) {
    return res.json({ reply: 'AI chat is not available in mock mode. Please configure a Claude API key in the .env file.' });
  }

  // Build real-time context injected into every chat system prompt
  const now = new Date();
  const realTime = {
    date:       now.toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }),
    time:       now.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
    isoDate:    now.toISOString().slice(0, 10),
    timezone:   Intl.DateTimeFormat().resolvedOptions().timeZone,
    appVersion: require('../package.json').version,
  };

  // Attach recent BRD project names and usage totals
  try {
    const allFiles = fs.readdirSync(SESSIONS_DIR).filter(f => f.endsWith('.json')).sort().reverse().slice(0, 100);
    const allSessions = allFiles
      .map(f => { try { return JSON.parse(fs.readFileSync(path.join(SESSIONS_DIR, f), 'utf8')); } catch { return null; } })
      .filter(Boolean);

    const seen = new Set();
    realTime.recentProjects = allSessions
      .filter(s => { const k = (s.projectName || '').toLowerCase().trim(); if (seen.has(k)) return false; seen.add(k); return true; })
      .slice(0, 10)
      .map(s => `${s.projectName} (${new Date(s.createdAt).toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' })}, ${s.summary?.total || 0} reqs)`);

    realTime.totalBRDsGenerated = allSessions.length;
  } catch { /* non-fatal */ }

  // Attach cumulative token usage
  try {
    if (fs.existsSync(USAGE_LOG_PATH)) {
      const usageLog = JSON.parse(fs.readFileSync(USAGE_LOG_PATH, 'utf8'));
      const totals = usageLog.reduce((a, e) => ({
        input:  a.input  + (e.input_tokens  || 0),
        output: a.output + (e.output_tokens || 0),
        cost:   a.cost   + (e.cost_usd || 0),
      }), { input: 0, output: 0, cost: 0 });
      realTime.usage = {
        totalBRDs:     usageLog.length,
        inputTokens:   totals.input,
        outputTokens:  totals.output,
        totalTokens:   totals.input + totals.output,
        totalCostUSD:  parseFloat(totals.cost.toFixed(4)),
      };
    }
  } catch { /* non-fatal */ }

  try {
    const reply = await claude.chat(messages, brdContext, realTime);
    res.json({ reply });
  } catch (err) {
    console.error('❌ AI chat error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /health ────────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => {
  res.json({
    status:      'running',
    service:     'Synoptek CE BRD Generator',
    version:     require('../package.json').version,
    ai:          process.env.CLAUDE_API_KEY ? 'Claude API (live)' : 'Mock mode',
    sharepoint:  SP_CONFIGURED ? 'configured' : 'not configured',
    output:      OUTPUT,
  });
});

// ── Start ──────────────────────────────────────────────────────────────────────
app.listen(PORT, async () => {
  console.log('');
  console.log('╔══════════════════════════════════════════╗');
  console.log('║   BRD Generator — Standalone Web App     ║');
  console.log('╚══════════════════════════════════════════╝');
  console.log(`  Local URL    : http://localhost:${PORT}`);
  console.log(`  AI Mode      : ${process.env.CLAUDE_API_KEY ? '🟢 Claude API' : '🟡 Mock mode'}`);;
  console.log(`  SharePoint   : ${SP_CONFIGURED              ? '🟢 Configured' : '🔴 Not configured'}`);
  console.log(`  Output folder: ${OUTPUT}`);

  // ── Tunnel mode (npm run share) ─────────────────────────────────────────────
  if (process.env.SHARE === 'true') {
    console.log('');
    console.log('  🔄 Starting public tunnel...');
    try {
      const localtunnel = require('localtunnel');
      const tunnel = await localtunnel({
        port:      PORT,
        subdomain: `synoptek-brd-${Math.random().toString(36).slice(2, 7)}`,
      });

      console.log('');
      console.log('╔══════════════════════════════════════════════════════════╗');
      console.log('║  🌐  PUBLIC URL — Send this to your testers:             ║');
      console.log(`║  👉  ${tunnel.url.padEnd(52)} ║`);
      console.log('║                                                          ║');
      console.log('║  ⚠️  First visit: click "Continue" on the welcome page   ║');
      console.log('║  🔴  Tunnel stops when you close this terminal           ║');
      console.log('╚══════════════════════════════════════════════════════════╝');
      console.log('');

      tunnel.on('error', err => {
        console.error('⚠️  Tunnel error:', err.message);
      });
      tunnel.on('close', () => {
        console.log('⚠️  Tunnel closed. Run "npm run share" again to restart.');
      });

    } catch (err) {
      console.error('❌ Could not start tunnel:', err.message);
      console.log('   Try: npm install localtunnel --save');
    }
  } else {
    console.log('');
    console.log('  💡 To share with others: run  npm run share');
    console.log('');
  }
});
