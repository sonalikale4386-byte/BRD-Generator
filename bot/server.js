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
    const allowed = ['.pdf', '.docx', '.txt', '.doc'];
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

/** Build a BRD response object from brdData + filePath */
function buildResponse(brdData, filePath) {
  const fileName = path.basename(filePath);
  const reqs     = brdData.requirements || [];
  return {
    success:     true,
    downloadUrl: `/output/${encodeURIComponent(fileName)}`,
    fileName,
    summary: {
      projectName: brdData.project_name,
      date:        brdData.document_date,
      version:     brdData.document_version,
      status:      brdData.status,
      total:       reqs.length,
      high:        reqs.filter(r => r.priority === 'High').length,
      medium:      reqs.filter(r => r.priority === 'Medium').length,
      low:         reqs.filter(r => r.priority === 'Low').length,
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
    let filesData;
    try {
      const r = await http.get(
        `/drives/${drive.id}/root:/${folderPath}:/children`,
        { params: { $select: 'id,name,size,lastModifiedDateTime,@microsoft.graph.downloadUrl' } }
      );
      filesData = r.data.value || [];
    } catch (err) {
      if (err.response?.status === 404) {
        return res.status(404).json({
          error: `Folder "${folderPath}" not found. Check the folder name or create it in SharePoint.`,
        });
      }
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

// ── POST /api/generate-brd  (Upload source) ───────────────────────────────────
app.post('/api/generate-brd', upload.array('files', 10), async (req, res) => {
  try {
    const projectName = (req.body.projectName || '').trim();
    const authorName  = (req.body.authorName  || 'BRD Generator').trim();
    if (!projectName) return res.status(400).json({ error: 'Project name is required.' });

    const files = req.files || [];
    console.log(`\n📋 [Upload] Generating BRD — "${projectName}"  Files: ${files.length}`);

    const extractor = new FileExtractor();
    let extractedText = '';
    for (const file of files) {
      const chunk = await extractBuffer(extractor, file.buffer, file.originalname);
      extractedText += chunk;
      console.log(`  📄 ${file.originalname} — ${chunk.length} chars extracted`);
    }

    // Cap total context at 60 000 chars (~15 000 words) to stay within model limits
    if (extractedText.length > 60000) {
      console.warn(`  ⚠️  Total extracted text ${extractedText.length} chars — truncating to 60 000`);
      extractedText = extractedText.slice(0, 60000);
    }
    console.log(`  📝 Total extracted: ${extractedText.length} chars across ${files.length} file(s)`);

    const claude   = new ClaudeService();
    const brdData  = await claude.generateBRD({ projectName, userName: authorName, extractedText });
    const filePath = await generateExcel(brdData, OUTPUT);

    console.log(`  ✅ Done — ${path.basename(filePath)}`);
    res.json(buildResponse(brdData, filePath));

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

  const { projectName, authorName = 'BRD Generator', files = [] } = req.body;
  if (!projectName) return res.status(400).json({ error: 'Project name is required.' });
  if (!files.length) return res.status(400).json({ error: 'No files selected from SharePoint.' });

  try {
    console.log(`\n📋 [SharePoint] Generating BRD — "${projectName}"  Files: ${files.length}`);

    const extractor = new FileExtractor();
    let extractedText = '';

    for (const file of files) {
      try {
        const resp  = await axios.get(file.downloadUrl, { responseType: 'arraybuffer', timeout: 30_000 });
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

    const claude   = new ClaudeService();
    const brdData  = await claude.generateBRD({
      projectName,
      userName: authorName,
      extractedText,
    });
    const filePath = await generateExcel(brdData, OUTPUT);

    console.log(`  ✅ Done — ${path.basename(filePath)}`);
    res.json(buildResponse(brdData, filePath));

  } catch (err) {
    console.error('❌ Generate BRD (SharePoint) error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /health ────────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => {
  res.json({
    status:      'running',
    service:     'BRD Generator — Standalone Web App',
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
  console.log(`  AI Mode      : ${process.env.CLAUDE_API_KEY ? '🟢 Claude API' : '🟡 Mock mode'}`);
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
