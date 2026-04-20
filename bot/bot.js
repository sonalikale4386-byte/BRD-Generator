/**
 * bot.js — BRDBot ActivityHandler
 *
 * Responsibilities:
 *  1. Route messages to the BRD waterfall dialog
 *  2. While waiting for files (step 2): delegate to dialog.handleWaitingMessage()
 *  3. Save conversation references for proactive messaging
 *  4. Handle pending "Regenerate BRD?" yes/no confirmations
 *  5. handleGraphNotification() — called by index.js when Graph webhook fires
 */
const { ActivityHandler, MessageFactory, TurnContext } = require('botbuilder');
const { DialogSet, DialogTurnStatus }                  = require('botbuilder-dialogs');
const { BRDDialog, DIALOG_ID }                         = require('./dialogues/brdDialog');
const { GraphService }                                 = require('./services/graphService');
const { ClaudeService }                                = require('./services/claudeService');
const { FileExtractor }                                = require('./services/fileExtractor');
const { generateExcel }                                = require('./brdCore');
const path                                             = require('path');
const fs                                               = require('fs');

const TRIGGERS = ['generate brd', 'create brd', 'new brd', 'start brd', 'brd', 'start', 'hello', 'hi'];

class BRDBot extends ActivityHandler {
  /**
   * @param {ConversationState} conversationState
   * @param {string}            outputDir
   * @param {string}            baseUrl
   * @param {SubscriptionStore} subStore
   * @param {CloudAdapter}      adapter   — needed for proactive messaging
   */
  constructor(conversationState, outputDir, baseUrl, subStore, adapter) {
    super();
    this.conversationState = conversationState;
    this.outputDir  = outputDir;
    this.baseUrl    = baseUrl;
    this.subStore   = subStore;
    this.adapter    = adapter;

    // Conversation state properties
    this.dialogStateProp  = conversationState.createProperty('DialogState');
    this.waitingFilesProp = conversationState.createProperty('WaitingFiles'); // bool

    // Dialog instance (share waitingFilesProp so dialog can set/clear it)
    this.brdDialog = new BRDDialog(outputDir, baseUrl, this.waitingFilesProp, subStore);

    // ── Message handler ─────────────────────────────────────────────────────
    this.onMessage(async (context, next) => {
      // Always save conversation reference + project name for proactive messaging
      this.subStore.saveConversationRef(
        context.activity.from.id,
        TurnContext.getConversationReference(context.activity)
      );

      // Save project name if dialog is active (for file-change notifications)
      const activeDialog = (await this._dc(context)).activeDialog;
      const projName     = activeDialog?.state?.values?.projectName;
      if (projName) this.subStore.saveLastProjectName(context.activity.from.id, projName);

      const dc      = await this._dc(context);
      const text    = (context.activity.text || '').trim().toLowerCase();
      const waiting = await this.waitingFilesProp.get(context, false);
      const userId  = context.activity.from.id;

      // ── Reset ────────────────────────────────────────────────────────────
      if (['reset', 'restart', 'cancel'].includes(text)) {
        await dc.cancelAllDialogs();
        await this.waitingFilesProp.set(context, false);
        this.subStore.clearPending(userId);
        await context.sendActivity('🔄 Reset. Type **generate BRD** to start again.');
        await this.conversationState.saveChanges(context, false);
        return next();
      }

      // ── Pending regeneration confirmation (yes / no) ─────────────────────
      const pending = this.subStore.getPending(userId);
      if (pending && ['yes', 'y'].includes(text)) {
        this.subStore.clearPending(userId);
        if (pending.isMock) {
          await this._regenerateFromLocal(context, pending);
        } else {
          await this._regenerateFromOneDrive(context, pending);
        }
        await this.conversationState.saveChanges(context, false);
        return next();
      }
      if (pending && ['no', 'n'].includes(text)) {
        this.subStore.clearPending(userId);
        await context.sendActivity('👍 OK — I\'ll include the new file next time you generate a BRD.');
        await this.conversationState.saveChanges(context, false);
        return next();
      }

      // ── Waiting for file uploads (dialog step 2) ─────────────────────────
      if (waiting && dc.activeDialog) {
        const done = await this.brdDialog.handleWaitingMessage(context, dc.activeDialog);
        if (done) {
          await this.waitingFilesProp.set(context, false);
          await dc.continueDialog();  // re-enters step 2 → _driveInit=true → step.next()
        }
        await this.conversationState.saveChanges(context, false);
        return next();
      }

      // ── Normal dialog flow ────────────────────────────────────────────────
      let result = await dc.continueDialog();

      if (result.status === DialogTurnStatus.empty) {
        if (TRIGGERS.some(t => text.includes(t))) {
          await dc.beginDialog(DIALOG_ID);
        } else {
          await context.sendActivity(this._welcome());
        }
      }

      await this.conversationState.saveChanges(context, false);
      return next();
    });

    // ── Welcome new members ─────────────────────────────────────────────────
    this.onMembersAdded(async (context, next) => {
      for (const m of context.activity.membersAdded || []) {
        if (m.id !== context.activity.recipient.id) {
          await context.sendActivity(this._welcome());
        }
      }
      return next();
    });
  }

  // ── Local file notification handler (mock OneDrive — called from index.js) ──
  /**
   * Called when fileWatcher detects a new file in test-docs/.
   * Sends a proactive "New file detected" message to the user.
   * @param {TurnContext} context   Proactive turn context
   * @param {string}      filename  Name of the new file
   * @param {string}      projectName  Last known project name (optional)
   */
  async handleLocalFileNotification(context, filename, projectName) {
    const userId = context.activity.from?.id || 'unknown';

    // Get current file list to show
    const TEST_DOCS = path.join(__dirname, '..', 'test-docs');
    const allFiles  = fs.existsSync(TEST_DOCS)
      ? fs.readdirSync(TEST_DOCS).filter(f => /\.(pdf|docx|txt)$/i.test(f))
      : [];

    const fileList = allFiles.map(f => `• 📄 ${f}`).join('\n');

    // Store pending confirmation
    this.subStore.setPending(userId, {
      isMock:      true,
      projectName: projectName || '',
      filename,
      fileList,
    });

    await context.sendActivity(MessageFactory.text(
      `📂 **New file detected in test-docs folder!**\n\n` +
      `➕ New file: **${filename}**\n\n` +
      `All files (${allFiles.length}):\n${fileList}\n\n` +
      `**Regenerate BRD with all documents?**\n` +
      `Reply **yes** or **no**`
    ));
  }

  // ── Graph notification handler (called from index.js) ─────────────────────
  /**
   * Called proactively when a new file is detected in the user's "BRD Document" folder.
   * Sends a Teams message asking if the user wants to regenerate the BRD.
   *
   * @param {TurnContext} context  Proactive turn context (from adapter.continueConversationAsync)
   * @param {object}      subEntry Subscription store entry for this subscription
   */
  async handleGraphNotification(context, subEntry) {
    const userId = subEntry.userId;

    // Try to get the list of current files to show the user
    let fileList = '';
    try {
      const tokenResp = await this.adapter.getUserToken(
        context,
        process.env.OAUTH_CONNECTION_NAME || 'GraphConnection'
      );
      if (tokenResp?.token) {
        const graph = new GraphService(tokenResp.token);
        const files = await graph.listBRDFiles();
        fileList = files.map(f => `• 📄 ${f.name}`).join('\n');
        // Update stored file list for the regeneration step
        subEntry.latestFiles = files;
      }
    } catch (_) { /* token not available — proceed without file list */ }

    // Store pending confirmation
    this.subStore.setPending(userId, {
      subId:       subEntry.subId,
      projectName: subEntry.projectName || '',
      fileList,
    });

    const msg = fileList
      ? `📂 **New file detected in your BRD Document folder!**\n\n${fileList}\n\n` +
        `**Regenerate BRD with all documents?**\nReply **yes** or **no**`
      : `📂 **New file detected in your BRD Document folder!**\n\n` +
        `**Regenerate BRD with all documents?**\nReply **yes** or **no**`;

    await context.sendActivity(MessageFactory.text(msg));
  }

  // ── Local regeneration (mock OneDrive) ────────────────────────────────────
  async _regenerateFromLocal(context, pending) {
    const TEST_DOCS = path.join(__dirname, '..', 'test-docs');
    const userName  = context.activity.from?.name || 'BRD Agent';

    await context.sendActivity('⏳ **Regenerating BRD with all documents...**\n\n_Step 1/3: Reading files..._');

    // Read all files from test-docs/
    const supported = ['.pdf', '.docx', '.txt'];
    const files     = fs.existsSync(TEST_DOCS)
      ? fs.readdirSync(TEST_DOCS).filter(f => supported.includes(path.extname(f).toLowerCase()))
      : [];

    if (!files.length) {
      await context.sendActivity('⚠️ No documents found in test-docs/ folder.');
      return;
    }

    // Extract text
    const extractor   = new FileExtractor();
    let extractedText = '';
    for (const filename of files) {
      try {
        const buf  = fs.readFileSync(path.join(TEST_DOCS, filename));
        const text = await extractor.extractFromBuffer(buf, filename);
        if (text) extractedText += `\n\n=== [OneDrive] ${filename} ===\n${text.slice(0, 6000)}`;
      } catch (err) {
        console.warn(`⚠️  Extract failed for ${filename}:`, err.message);
      }
    }

    await context.sendActivity('_Step 2/3: Generating BRD content with AI..._');

    let brdData;
    try {
      const claude = new ClaudeService();
      brdData = await claude.generateBRD({
        projectName: pending.projectName || 'BRD Project',
        userName,
        extractedText,
      });
    } catch (err) {
      await context.sendActivity(`❌ AI generation failed: ${err.message}`);
      return;
    }

    await context.sendActivity('_Step 3/3: Building Excel document..._');

    let filePath;
    try {
      filePath = await generateExcel(brdData, this.outputDir);
    } catch (err) {
      await context.sendActivity(`❌ Excel build failed: ${err.message}`);
      return;
    }

    const fileName = path.basename(filePath);
    const fileUrl  = `${this.baseUrl}/output/${encodeURIComponent(fileName)}`;
    const reqs     = brdData.requirements || [];

    await context.sendActivity(MessageFactory.text(
      `✅ **BRD Regenerated Successfully!**\n\n` +
      `📊 **Project:** ${brdData.project_name}\n` +
      `📋 **Requirements:** ${reqs.length}` +
      `  |  🔴 High: ${reqs.filter(r => r.priority === 'High').length}` +
      `  |  🟡 Med: ${reqs.filter(r => r.priority === 'Medium').length}` +
      `  |  🟢 Low: ${reqs.filter(r => r.priority === 'Low').length}\n\n` +
      `📥 **[Download BRD: ${fileName}](${fileUrl})**`
    ));
  }

  // ── Proactive regeneration ─────────────────────────────────────────────────
  async _regenerateFromOneDrive(context, pending) {
    const userName = context.activity.from?.name || 'BRD Agent';

    await context.sendActivity('⏳ **Regenerating BRD with all OneDrive documents...**\n\n_Step 1/3: Accessing OneDrive..._');

    // Get cached token
    let token;
    try {
      const resp = await this.adapter.getUserToken(
        context,
        process.env.OAUTH_CONNECTION_NAME || 'GraphConnection'
      );
      token = resp?.token;
    } catch (_) {}

    if (!token) {
      await context.sendActivity(
        '⚠️ Your session has expired. Please type **generate BRD** to sign in and regenerate.'
      );
      return;
    }

    // Fetch all current files from OneDrive
    let driveFiles = [];
    try {
      const graph = new GraphService(token);
      driveFiles  = await graph.listBRDFiles();
    } catch (err) {
      await context.sendActivity(`❌ Could not read OneDrive: ${err.message}`);
      return;
    }

    // Extract text
    await context.sendActivity('_Step 2/3: Processing documents..._');
    const graph = new GraphService(token);
    const fileEx = require('./services/fileExtractor');
    const extractor = new fileEx.FileExtractor();
    let extractedText = '';

    for (const file of driveFiles) {
      try {
        const buf  = await graph.downloadFile(file['@microsoft.graph.downloadUrl']);
        const text = await extractor.extractFromBuffer(buf, file.name);
        if (text) extractedText += `\n\n=== [OneDrive] ${file.name} ===\n${text.slice(0, 6000)}`;
      } catch (err) {
        console.warn(`⚠️ Extract failed: ${file.name}:`, err.message);
      }
    }

    await context.sendActivity('_Step 3/3: Generating BRD..._');

    let brdData;
    try {
      const claude = new ClaudeService();
      brdData = await claude.generateBRD({
        projectName: pending.projectName || 'BRD Project',
        userName,
        extractedText,
      });
    } catch (err) {
      await context.sendActivity(`❌ AI generation failed: ${err.message}`);
      return;
    }

    let filePath;
    try {
      filePath = await generateExcel(brdData, this.outputDir);
    } catch (err) {
      await context.sendActivity(`❌ Excel build failed: ${err.message}`);
      return;
    }

    const fileName = path.basename(filePath);
    const fileUrl  = `${this.baseUrl}/output/${encodeURIComponent(fileName)}`;
    const reqs     = brdData.requirements || [];

    await context.sendActivity(MessageFactory.text(
      `✅ **BRD Regenerated Successfully!**\n\n` +
      `📊 **Project:** ${brdData.project_name}\n` +
      `📋 **Requirements:** ${reqs.length}` +
      `  |  🔴 High: ${reqs.filter(r => r.priority === 'High').length}` +
      `  |  🟡 Med: ${reqs.filter(r => r.priority === 'Medium').length}` +
      `  |  🟢 Low: ${reqs.filter(r => r.priority === 'Low').length}\n\n` +
      `📥 **[Download BRD: ${fileName}](${fileUrl})**`
    ));
  }

  // ── Helpers ────────────────────────────────────────────────────────────────
  async _dc(context) {
    const dialogSet = new DialogSet(this.dialogStateProp);
    dialogSet.add(this.brdDialog);
    return dialogSet.createContext(context);
  }

  _welcome() {
    return MessageFactory.text(
      '👋 **BRD Generator Agent**\n\n' +
      'I generate Synoptek-format BRDs from your OneDrive documents and meeting files.\n\n' +
      '**How it works:**\n' +
      '1. Tell me the **project name**\n' +
      '2. Sign in so I can read your **OneDrive → BRD Document** folder\n' +
      '3. Optionally upload additional files\n' +
      '4. Get your **Excel BRD** instantly\n\n' +
      '📂 New files added to **BRD Document** are detected automatically — I\'ll ask before regenerating.\n\n' +
      '**To start:** type `generate BRD`\n' +
      '**To reset:** type `reset`'
    );
  }
}

module.exports = { BRDBot };
