/**
 * index.js — BRD Generator Agent entry point
 * M365 Agents SDK (Bot Framework v4) + Express
 *
 * Endpoints:
 *   POST /api/messages              — Bot Framework messages
 *   POST /api/graph-notifications   — Microsoft Graph webhook (OneDrive change events)
 *   GET  /output/:file              — Download generated Excel files
 *   GET  /                          — Health check
 */
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const express  = require('express');
const path     = require('path');
const fs       = require('fs');
const {
  CloudAdapter, ConfigurationBotFrameworkAuthentication,
  ConversationState, MemoryStorage,
} = require('botbuilder');

const { BRDBot }           = require('./bot');
const { SubscriptionStore } = require('./services/subscriptionStore');
const { WH_SECRET }        = require('./services/graphService');
const { FileWatcher }      = require('./services/fileWatcher');

// ── Config ────────────────────────────────────────────────────────────────────
const PORT     = process.env.PORT       || 3978;
const BASE_URL = process.env.BOT_BASE_URL || `http://localhost:${PORT}`;
const OUTPUT   = path.join(__dirname, '..', 'output');
if (!fs.existsSync(OUTPUT)) fs.mkdirSync(OUTPUT, { recursive: true });

// ── Bot Framework adapter ─────────────────────────────────────────────────────
const botAuth = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId:       process.env.MICROSOFT_APP_ID       || '',
  MicrosoftAppPassword: process.env.MICROSOFT_APP_PASSWORD || '',
  MicrosoftAppType:     process.env.MICROSOFT_APP_TYPE     || 'MultiTenant',
});
const adapter = new CloudAdapter(botAuth);

adapter.onTurnError = async (context, error) => {
  console.error('[BRD Agent] onTurnError:', error);
  try {
    await context.sendActivity('❌ Unexpected error. Type **reset** and try again.');
  } catch (_) {}
};

// ── Shared state ──────────────────────────────────────────────────────────────
const memory            = new MemoryStorage();
const conversationState = new ConversationState(memory);
const subStore          = new SubscriptionStore();

// ── Bot ───────────────────────────────────────────────────────────────────────
const bot = new BRDBot(conversationState, OUTPUT, BASE_URL, subStore, adapter);

// ── Express app ───────────────────────────────────────────────────────────────
const app = express();
app.use(express.json());

// ── Bot messages ──────────────────────────────────────────────────────────────
app.post('/api/messages', async (req, res) => {
  await adapter.process(req, res, async context => {
    await bot.run(context);
  });
});

// ── Microsoft Graph webhook (OneDrive change notifications) ───────────────────
app.post('/api/graph-notifications', express.text({ type: '*/*' }), async (req, res) => {
  // Step 1 — Subscription validation handshake
  if (req.query.validationToken) {
    console.log('✅ Graph subscription validated');
    return res.status(200)
      .set('Content-Type', 'text/plain')
      .send(req.query.validationToken);
  }

  // Step 2 — Acknowledge immediately (Graph requires < 3 s response)
  res.status(202).send();

  try {
    const body          = JSON.parse(typeof req.body === 'string' ? req.body : JSON.stringify(req.body));
    const notifications = body.value || [];
    console.log(`📬 Graph notification: ${notifications.length} event(s)`);

    for (const notif of notifications) {
      // Validate client state to prevent spoofing
      if (notif.clientState !== WH_SECRET) {
        console.warn('⚠️  Graph notification: invalid clientState, skipping');
        continue;
      }

      const subId    = notif.subscriptionId;
      const subEntry = subStore.getSubscription(subId);
      if (!subEntry) {
        console.warn(`⚠️  Graph notification: unknown subscription ${subId}`);
        continue;
      }

      const convRef = subEntry.conversationRef;
      if (!convRef) {
        console.warn(`⚠️  Graph notification: no conversation ref for sub ${subId}`);
        continue;
      }

      // Send proactive message to the user who owns this subscription
      console.log(`🔔 New file in OneDrive for user ${subEntry.userId} — sending notification`);
      await adapter.continueConversationAsync(
        process.env.MICROSOFT_APP_ID || '',
        convRef,
        async (context) => {
          await bot.handleGraphNotification(context, { ...subEntry, subId });
        }
      );
    }
  } catch (err) {
    console.error('Graph notification processing error:', err.message);
  }
});

// ── Serve generated Excel files ───────────────────────────────────────────────
app.use('/output', express.static(OUTPUT));

// ── Health check ──────────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({
    status:       'running',
    agent:        'BRD Generator — M365 Agents SDK',
    endpoint:     `${BASE_URL}/api/messages`,
    notifications:`${BASE_URL}/api/graph-notifications`,
    claudeApi:    process.env.CLAUDE_API_KEY ? 'configured' : 'mock mode',
    oauthConn:    process.env.OAUTH_CONNECTION_NAME || 'GraphConnection (default)',
  });
});

// ── Mock OneDrive: watch test-docs/ for new files ────────────────────────────
if (process.env.MOCK_ONEDRIVE === 'true') {
  const TEST_DOCS   = path.join(__dirname, '..', 'test-docs');
  const fileWatcher = new FileWatcher(TEST_DOCS);

  fileWatcher.onNewFile = async (filename) => {
    // Find the most recently active user's conversation reference
    const allRefs = subStore.getAllConversationRefs();
    if (!allRefs.length) {
      console.log('⚠️  New file detected but no active conversation to notify.');
      return;
    }

    // Notify the most recently active user
    const { userId, convRef } = allRefs[allRefs.length - 1];
    const lastProject = subStore.getLastProjectName(userId);

    try {
      await adapter.continueConversationAsync(
        process.env.MICROSOFT_APP_ID || '',
        convRef,
        async (context) => {
          await bot.handleLocalFileNotification(context, filename, lastProject);
        }
      );
    } catch (err) {
      console.error('⚠️  Could not send file notification:', err.message);
    }
  };

  fileWatcher.start();
}

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log('');
  console.log('╔══════════════════════════════════════════════════╗');
  console.log('║   BRD Generator Agent — M365 Agents SDK          ║');
  console.log('╚══════════════════════════════════════════════════╝');
  console.log(`  Bot endpoint    : ${BASE_URL}/api/messages`);
  console.log(`  Graph webhook   : ${BASE_URL}/api/graph-notifications`);
  console.log(`  Output folder   : ${OUTPUT}`);
  console.log(`  AI Mode         : ${process.env.CLAUDE_API_KEY ? '🟢 Claude API' : '🟡 Mock'}`);
  console.log(`  OAuth connection: ${process.env.OAUTH_CONNECTION_NAME || 'GraphConnection'}`);
  console.log('');
});
