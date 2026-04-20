/**
 * brdDialog.js — BRD Generation waterfall dialog
 *
 * Steps:
 *  0. askProjectName   — TextPrompt for project name
 *  1. acquireToken     — OAuthPrompt (Teams SSO) OR skipped in mock/local mode
 *  2. fetchDriveFiles  — fetch OneDrive "BRD Document" files + create subscription
 *                        first entry: fetch → show files → wait
 *                        re-entry:    advance to generate
 *  3. generate         — extract text → Claude → Excel → reply with link
 *
 * Modes:
 *   MOCK_ONEDRIVE=true  → reads from local test-docs/ folder (emulator testing)
 *   MICROSOFT_APP_ID=   → manual upload only (no SSO, no OneDrive)
 *   MICROSOFT_APP_ID set → full Teams SSO + real OneDrive
 */
const path = require('path');
const fs   = require('fs');

const {
  WaterfallDialog, TextPrompt, OAuthPrompt,
  ComponentDialog, DialogTurnStatus,
} = require('botbuilder-dialogs');
const { MessageFactory, TurnContext } = require('botbuilder');
const { FileExtractor }               = require('../services/fileExtractor');
const { ClaudeService }               = require('../services/claudeService');
const { GraphService, BRD_FOLDER }    = require('../services/graphService');
const { generateExcel }               = require('../brdCore');

const DIALOG_ID    = 'brdDialog';
const TEXT_PROMPT  = 'textPrompt';
const OAUTH_PROMPT = 'oAuthPrompt';

const TEST_DOCS_DIR = path.join(__dirname, '..', '..', 'test-docs');

class BRDDialog extends ComponentDialog {
  constructor(outputDir, baseUrl, waitingFilesAccessor, subStore) {
    super(DIALOG_ID);
    this.outputDir   = outputDir;
    this.baseUrl     = baseUrl;
    this.waitingFlag = waitingFilesAccessor;
    this.subStore    = subStore;
    this.fileEx      = new FileExtractor();
    this.claudeSvc   = new ClaudeService();

    this.addDialog(new TextPrompt(TEXT_PROMPT));
    this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
      connectionName: process.env.OAUTH_CONNECTION_NAME || 'GraphConnection',
      text:           'Please sign in to allow the BRD Agent to access your OneDrive.',
      title:          'Sign In to OneDrive',
      timeout:        300_000,
    }));
    this.addDialog(new WaterfallDialog(DIALOG_ID, [
      this.askProjectName.bind(this),   // 0
      this.acquireToken.bind(this),     // 1
      this.fetchDriveFiles.bind(this),  // 2  ← waits here
      this.generate.bind(this),         // 3
    ]));
    this.initialDialogId = DIALOG_ID;
  }

  // ── Step 0 ────────────────────────────────────────────────────────────────
  async askProjectName(step) {
    return step.prompt(TEXT_PROMPT, '📋 **What is the Project Name?**');
  }

  // ── Step 1 ────────────────────────────────────────────────────────────────
  async acquireToken(step) {
    step.values.projectName = step.result;

    // Mock mode or no credentials → skip OAuth
    if (process.env.MOCK_ONEDRIVE === 'true' || !process.env.MICROSOFT_APP_ID) {
      step.values.accessToken = null;
      return step.next();
    }

    return step.beginDialog(OAUTH_PROMPT);
  }

  // ── Step 2 ────────────────────────────────────────────────────────────────
  async fetchDriveFiles(step) {
    // Re-entry → advance to generate
    if (step.values._driveInit) return step.next();

    step.values._driveInit  = true;
    step.values.driveFiles  = [];
    step.values.manualFiles = [];

    const isMock     = process.env.MOCK_ONEDRIVE === 'true';
    const hasAppId   = !!process.env.MICROSOFT_APP_ID;
    const tokenResp  = step.result;  // OAuthPrompt result (or undefined)

    // ── MOCK MODE — read from test-docs/ folder ─────────────────────────────
    if (isMock) {
      const mockFiles = _readTestDocs();
      step.values.driveFiles = mockFiles;

      if (mockFiles.length > 0) {
        const list = mockFiles.map(f => `• 📄 ${f.name}  _(${_size(f.size)})_`).join('\n');
        await step.context.sendActivity(
          `✅ **[MOCK] Simulating OneDrive — BRD Document folder**\n\n` +
          `Found **${mockFiles.length} file(s)** in \`test-docs/\`:\n\n${list}\n\n` +
          `📎 Attach additional files, or type **"skip"** to proceed with these only.`
        );
      } else {
        await step.context.sendActivity(
          `📂 **[MOCK] \`test-docs/\` folder is empty.**\n\n` +
          `Add PDF / DOCX / TXT files to \`d:\\BRD Generator\\test-docs\\\` ` +
          `to simulate OneDrive documents.\n\n` +
          `📎 Attach files directly, or type **"skip"** to generate from project name only.`
        );
      }

    // ── NO CREDENTIALS — manual upload only ────────────────────────────────
    } else if (!hasAppId || !tokenResp?.token) {
      await step.context.sendActivity(
        `✅ Project: **${step.values.projectName}**\n\n` +
        `📎 Please **attach your documents** (PDF, DOCX, TXT), ` +
        `or type **"skip"** to generate from project name only.`
      );

    // ── REAL ONEDRIVE ───────────────────────────────────────────────────────
    } else {
      step.values.accessToken = tokenResp.token;
      const userId = step.context.activity.from.id;

      try {
        const graph = new GraphService(tokenResp.token);
        const files = await graph.listBRDFiles();
        step.values.driveFiles = files;

        // Create / renew Graph webhook subscription
        const notifUrl = `${this.baseUrl}/api/graph-notifications`;
        const existing = this.subStore.getSubscriptionByUser(userId);
        try {
          const sub = await graph.upsertSubscription(notifUrl, existing?.subId);
          this.subStore.setSubscription(sub.id, {
            userId,
            conversationRef: TurnContext.getConversationReference(step.context.activity),
            projectName:     step.values.projectName,
            subscribedAt:    new Date().toISOString(),
            expiresAt:       sub.expirationDateTime,
          });
        } catch (subErr) {
          console.warn('⚠️  Subscription creation failed:', subErr.message);
        }

        if (files.length > 0) {
          const list = files.map(f => `• 📄 ${f.name}  _(${_size(f.size)})_`).join('\n');
          await step.context.sendActivity(
            `✅ Found **${files.length} file(s)** in your **${BRD_FOLDER}** folder:\n\n${list}\n\n` +
            `📎 Attach additional files, or type **"skip"** to proceed with these only.`
          );
        } else {
          await step.context.sendActivity(
            `📂 Your **${BRD_FOLDER}** folder on OneDrive is empty.\n\n` +
            `📎 Attach documents directly, or type **"skip"** to generate from project name only.`
          );
        }
      } catch (err) {
        console.error('OneDrive fetch error:', err.message);
        await step.context.sendActivity(
          `⚠️ Could not read OneDrive: ${err.message}\n\n📎 Attach documents manually.`
        );
      }
    }

    await this.waitingFlag.set(step.context, true);
    return { status: DialogTurnStatus.waiting };
  }

  // ── Step 3 ────────────────────────────────────────────────────────────────
  async generate(step) {
    const { projectName, driveFiles, manualFiles, accessToken } = step.values;
    const userName  = step.context.activity.from?.name || 'BRD Agent';
    const isMock    = process.env.MOCK_ONEDRIVE === 'true';

    await step.context.sendActivity('⏳ **Generating your BRD...**\n\n_Step 1/3: Processing documents..._');

    let extractedText = '';

    // Extract from mock local files OR real OneDrive files
    if (driveFiles?.length) {
      const graph = (!isMock && accessToken) ? new GraphService(accessToken) : null;

      for (const file of driveFiles) {
        try {
          let buf;
          if (isMock || file._isMock) {
            // Read directly from disk
            buf = fs.readFileSync(file.localPath);
          } else {
            buf = await graph.downloadFile(file['@microsoft.graph.downloadUrl']);
          }
          const text = await this.fileEx.extractFromBuffer(buf, file.name);
          if (text) extractedText += `\n\n=== [OneDrive] ${file.name} ===\n${text.slice(0, 6000)}`;
        } catch (err) {
          console.warn(`⚠️  Extract failed for ${file.name}:`, err.message);
        }
      }
    }

    // Extract from manually uploaded files
    for (const file of (manualFiles || [])) {
      try {
        const text = await this.fileEx.extractText(file.url, file.name);
        if (text) extractedText += `\n\n=== [Uploaded] ${file.name} ===\n${text.slice(0, 6000)}`;
      } catch (err) {
        console.warn(`⚠️  Upload extract failed for ${file.name}:`, err.message);
      }
    }

    await step.context.sendActivity('_Step 2/3: Generating BRD content with AI..._');

    let brdData;
    try {
      brdData = await this.claudeSvc.generateBRD({ projectName, userName, extractedText });
    } catch (err) {
      await step.context.sendActivity(`❌ AI generation failed: ${err.message}\n\nType **generate BRD** to retry.`);
      return step.endDialog();
    }

    await step.context.sendActivity('_Step 3/3: Building Excel document..._');

    let filePath;
    try {
      filePath = await generateExcel(brdData, this.outputDir);
    } catch (err) {
      await step.context.sendActivity(`❌ Excel build failed: ${err.message}`);
      return step.endDialog();
    }

    await _sendResult(step.context, brdData, filePath, this.baseUrl);
    return step.endDialog();
  }

  // ── handleWaitingMessage ──────────────────────────────────────────────────
  async handleWaitingMessage(ctx, dialogState) {
    const text        = (ctx.activity.text || '').trim().toLowerCase();
    const attachments = ctx.activity.attachments || [];

    if (['generate', 'done', 'proceed'].includes(text)) return true;

    if (text === 'skip') {
      await ctx.sendActivity('⚠️ Skipping additional uploads — generating BRD now...');
      return true;
    }

    if (attachments.length > 0) {
      const values = dialogState.state.values;
      if (!values.manualFiles) values.manualFiles = [];

      for (const att of attachments) {
        const url  = att.contentType === 'application/vnd.microsoft.teams.file.download.info'
          ? att.content?.downloadUrl : att.contentUrl;
        const name = att.name || 'document';
        if (url) values.manualFiles.push({ url, name });
      }

      const dCount = (values.driveFiles  || []).length;
      const mCount = (values.manualFiles || []).length;
      const list   = values.manualFiles.map(f => `• 📄 ${f.name}`).join('\n');

      await ctx.sendActivity(
        `✅ **${attachments.length} file(s) received:**\n${list}\n\n` +
        `📊 Total: ${dCount} OneDrive + ${mCount} uploaded\n\n` +
        `⏳ Starting BRD generation...`
      );
      return true;
    }

    const dCount = (dialogState.state.values?.driveFiles || []).length;
    await ctx.sendActivity(
      dCount > 0
        ? `📎 Attach additional files, or type **"skip"** to proceed with ${dCount} OneDrive file(s).`
        : `📎 Attach your documents, or type **"skip"** to generate from project name only.`
    );
    return false;
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Read all supported files from test-docs/ and return mock driveFile objects */
function _readTestDocs() {
  const supported = ['.pdf', '.docx', '.txt'];
  if (!fs.existsSync(TEST_DOCS_DIR)) {
    fs.mkdirSync(TEST_DOCS_DIR, { recursive: true });
    return [];
  }
  return fs.readdirSync(TEST_DOCS_DIR)
    .filter(f => supported.includes(path.extname(f).toLowerCase()))
    .map(f => {
      const fullPath = path.join(TEST_DOCS_DIR, f);
      const stat     = fs.statSync(fullPath);
      return { name: f, localPath: fullPath, size: stat.size, _isMock: true };
    });
}

function _size(bytes) {
  if (!bytes) return '0 KB';
  return bytes < 1_048_576
    ? `${Math.round(bytes / 1024)} KB`
    : `${(bytes / 1_048_576).toFixed(1)} MB`;
}

async function _sendResult(ctx, brdData, filePath, baseUrl) {
  const fileName = path.basename(filePath);
  const fileUrl  = `${baseUrl}/output/${encodeURIComponent(fileName)}`;
  const reqs     = brdData.requirements || [];
  await ctx.sendActivity(MessageFactory.text(
    `✅ **BRD Generated Successfully!**\n\n` +
    `📊 **Project:** ${brdData.project_name}\n` +
    `📋 **Requirements:** ${reqs.length}` +
    `  |  🔴 High: ${reqs.filter(r => r.priority === 'High').length}` +
    `  |  🟡 Med: ${reqs.filter(r => r.priority === 'Medium').length}` +
    `  |  🟢 Low: ${reqs.filter(r => r.priority === 'Low').length}\n\n` +
    `📥 **[Download BRD: ${fileName}](${fileUrl})**\n\n` +
    `_Type **generate BRD** to create another._`
  ));
}

module.exports = { BRDDialog, DIALOG_ID };
