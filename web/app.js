/**
 * app.js — BRD Generator Chatbot
 *
 * Conversation state machine:
 *   AWAIT_AUTHOR → AWAIT_PROJECT → AWAIT_BRD_TYPE
 *     ├─ update → AWAIT_UPDATE_METHOD ─┐
 *     └─ new    → AWAIT_DETAIL_LEVEL  ─┤
 *                                      ├─ AWAIT_FIT_GAP → AWAIT_SOURCE_REF → AWAIT_MOSCOW
 *                                      └─ AWAIT_SOURCE
 *                                           ├─ sharepoint → AWAIT_SP_URL → AWAIT_SP_FOLDER → AWAIT_SP_CONFIRM
 *                                           ├─ upload     → AWAIT_FILES
 *                                           └─ paste      → AWAIT_PASTE_CONTENT
 *                                      → AWAIT_ADDITIONAL_INPUTS → (AWAIT_ADDITIONAL_TEXT)
 *                                      → AWAIT_OUTPUT_FORMAT → GENERATING
 */

// ── DOM ───────────────────────────────────────────────────────────────────────
const chatMessages    = document.getElementById('chatMessages');
const chatInput       = document.getElementById('chatInput');
const sendBtn         = document.getElementById('sendBtn');
const attachBtn       = document.getElementById('attachBtn');
const fileInput       = document.getElementById('fileInput');
const attachPreview   = document.getElementById('attachPreview');
const attachPreviewIn = document.getElementById('attachPreviewInner');
const newChatBtn      = document.getElementById('newChatBtn');
const aiBadge         = document.getElementById('aiBadge');
const spBadge         = document.getElementById('spBadge');

// ── States ────────────────────────────────────────────────────────────────────
const S = {
  IDLE:                    'idle',
  AWAIT_AUTHOR:            'await_author',
  AWAIT_PROJECT:           'await_project',
  AWAIT_BRD_TYPE:          'await_brd_type',
  AWAIT_UPDATE_METHOD:     'await_update_method',
  AWAIT_DETAIL_LEVEL:      'await_detail_level',
  AWAIT_FIT_GAP:           'await_fit_gap',
  AWAIT_SOURCE_REF:        'await_source_ref',
  AWAIT_MOSCOW:            'await_moscow',
  AWAIT_SOURCE:            'await_source',
  AWAIT_PASTE_CONTENT:     'await_paste_content',
  AWAIT_FILES:             'await_files',
  AWAIT_SP_URL:            'await_sp_url',
  AWAIT_SP_FOLDER:         'await_sp_folder',
  AWAIT_SP_CONFIRM:        'await_sp_confirm',
  AWAIT_ADDITIONAL_INPUTS: 'await_additional_inputs',
  AWAIT_ADDITIONAL_TEXT:   'await_additional_text',
  AWAIT_OUTPUT_FORMAT:     'await_output_format',
  AWAIT_SOW:               'await_sow',
  AWAIT_SOW_FILES:         'await_sow_files',
  GENERATING:              'generating',
  AWAIT_UPDATE_DOCS:       'await_update_docs',
  AWAIT_UPDATE_FILES:      'await_update_files',
  DONE:                    'done',
};

// ── Session ───────────────────────────────────────────────────────────────────
let session = createSession();
let chatLog = [];          // [{role:'bot'|'user', type:'say'|'ask', text}]
let _chatLogActive = true; // false during generation + replay to avoid double-tracking
function createSession() {
  return {
    state:            S.IDLE,
    authorName:       '',
    projectName:      '',
    brdType:          '',   // 'new' | 'update'
    updateMethod:     '',   // 'latest' | 'upload_prev'
    detailLevel:      '',   // 'elaborated' | 'concise'
    fitGap:           '',   // 'yes' | 'no'
    sourceRef:        '',   // 'yes' | 'no'
    moscow:           '',   // 'yes' | 'no'
    source:           null, // 'sharepoint' | 'upload' | 'paste'
    pasteContent:     '',
    spSiteUrl:        '',
    spFolder:         'BRD Document',
    spFiles:          [],
    uploadFiles:      [],
    sowFiles:         [],
    additionalInputs: '',
    outputFormat:     '',   // 'excel' | 'word' | 'pdf'
    spConfigured:     false,
  };
}

// ── Init ──────────────────────────────────────────────────────────────────────
(async () => {
  try {
    const res  = await fetch('/health');
    const data = await res.json();

    if (data.ai?.includes('live')) {
      aiBadge.textContent = '🟢 AI: Claude';
      aiBadge.classList.add('live');
    } else {
      aiBadge.textContent = '🟡 AI: Mock';
      aiBadge.classList.add('mock');
    }

    session.spConfigured = data.sharepoint === 'configured';
    if (session.spConfigured) {
      spBadge.textContent = '🟢 SharePoint';
      spBadge.classList.add('ok');
    } else {
      spBadge.textContent = '🔴 SharePoint';
      spBadge.classList.add('off');
    }
  } catch { /* server offline */ }

  loadSessions();

  startConversation();
})();

// ── Start / Reset ─────────────────────────────────────────────────────────────
function startConversation() {
  chatMessages.innerHTML = '';
  session = createSession();
  session.spConfigured = spBadge.classList.contains('ok');
  chatLog = [];
  _chatLogActive = true;
  clearAttachments();

  addDateSep('Today');

  botSay(`Greetings! You are now interacting with the **Synoptek CE BRD Generator**.\n\nI'm here to assist you in developing a well-structured and professional Business Requirements Document in accordance with Synoptek CRM standards.`);

  setTimeout(() => {
    botAsk(`👤 Kindly specify the name of the BRD author.`, [
      { label: 'Skip', value: '__skip__' },
    ]);
    session.state = S.AWAIT_AUTHOR;
    enableInput('Type author name...');
  }, 600);
}

newChatBtn.addEventListener('click', () => {
  if (session.state === S.GENERATING) return;
  startConversation();
});

// ── Input handling ────────────────────────────────────────────────────────────
sendBtn.addEventListener('click', handleSend);
chatInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(); }
});

function handleSend() {
  const text = chatInput.value.trim();
  if (!text && session.uploadFiles.length === 0) return;
  if (session.state === S.GENERATING || session.state === S.IDLE) return;

  if (session.state === S.AWAIT_FILES && session.uploadFiles.length > 0 && !text) {
    handleUserMessage('(files attached)');
    return;
  }

  if (text) handleUserMessage(text);
}

function handleUserMessage(text) {
  addUserMsg(text);
  chatInput.value = '';
  processInput(text);
}

// ── State machine ─────────────────────────────────────────────────────────────
async function processInput(text) {
  const t = text.trim().toLowerCase();

  switch (session.state) {

    // ── Author name (FIRST question) ──────────────────────────────────────────
    case S.AWAIT_AUTHOR: {
      session.authorName = (text === '__skip__' || !text.trim()) ? 'BRD Agent' : text.trim();
      if (text !== '__skip__') botSay(`Author: **${session.authorName}** ✓`);
      session.state = S.AWAIT_PROJECT;
      setTimeout(() => {
        botAsk(`📋 To get started, please share the name of the project.`);
        enableInput('Type the project name...');
      }, 400);
      break;
    }

    // ── Project name ──────────────────────────────────────────────────────────
    case S.AWAIT_PROJECT: {
      if (!text.trim()) { botSay('Please enter a project name.'); return; }
      session.projectName = text.trim();
      session.state = S.AWAIT_BRD_TYPE;
      botSay(`Got it — **${session.projectName}** ✓`);
      setTimeout(() => {
        disableInput();
        botAsk(
          `Is this the first time you are generating a BRD for this project, or would you like to update an existing one?\n\n_(This helps me decide whether to start fresh or build on what already exists.)_`,
          [
            { label: '🆕 Create New BRD (first-time generation)',                          value: 'new'    },
            { label: '🔄 Update Existing BRD (use the latest version as a reference)',     value: 'update' },
          ]
        );
      }, 400);
      break;
    }

    // ── BRD type ──────────────────────────────────────────────────────────────
    case S.AWAIT_BRD_TYPE: {
      if (text !== 'new' && text !== 'update') { botSay(`Please select one of the options above.`); return; }
      session.brdType = text;
      botSay(`**${text === 'new' ? '🆕 Create New BRD' : '🔄 Update Existing BRD'}** selected ✓`);

      if (text === 'update') {
        session.state = S.AWAIT_UPDATE_METHOD;
        setTimeout(() => {
          disableInput();
          botAsk(
            `Please confirm how you would like to proceed:\n\n_(Choose the reference source for updating your BRD.)_`,
            [
              { label: '📄 Use the latest available BRD to refine and update the content based on new inputs', value: 'latest'      },
              { label: '📤 Upload a previous version of the BRD for reference',                                value: 'upload_prev' },
            ]
          );
        }, 400);
      } else {
        session.state = S.AWAIT_SOW;
        setTimeout(askSOW, 400);
      }
      break;
    }

    // ── Update method ─────────────────────────────────────────────────────────
    case S.AWAIT_UPDATE_METHOD: {
      if (text !== 'latest' && text !== 'upload_prev') { botSay(`Please select one of the options above.`); return; }
      session.updateMethod = text;
      if (text === 'latest') {
        botSay(`**📄 Latest BRD loaded as reference** ✓\n\nI will edit the previously generated BRD — keeping valid requirements, updating changed ones, and adding new ones based on your documents.`);
      } else {
        botSay(`**📤 Upload previous BRD version for reference** ✓`);
      }
      askFitGap();
      break;
    }

    // ── SOW question (new BRDs only) ──────────────────────────────────────────
    case S.AWAIT_SOW: {
      if (t === 'yes') {
        botSay(`📎 Please **attach your SOW document(s)** using the 📎 button, then type **done** to continue.`);
        session.sowFiles = [];
        session.state    = S.AWAIT_SOW_FILES;
        enableInput('Attach SOW file(s), then type done...');
        enableAttach(true);
      } else if (t === 'no') {
        botSay(`⏭ Continuing without SOW ✓`);
        askDetailLevel();
      } else {
        botSay(`Please select one of the options above.`);
      }
      break;
    }

    case S.AWAIT_SOW_FILES: {
      if (t === 'done' || text === '(files attached)') {
        if (!session.sowFiles.length) {
          botSay(`⚠️ Please attach your SOW document before proceeding.`);
          return;
        }
        const names = session.sowFiles.map(f => `• ${f.name}`).join('\n');
        botSay(`✅ **SOW uploaded (${session.sowFiles.length} file(s)):**\n${names}`);
        enableAttach(false);
        askDetailLevel();
      } else if (t === 'skip') {
        botSay(`⏭ Skipping SOW ✓`);
        enableAttach(false);
        askDetailLevel();
      } else {
        botSay(`Use the 📎 button to attach your SOW document, then type **done**.\nOr type **skip** to continue without it.`);
      }
      break;
    }

    // ── Detail level ──────────────────────────────────────────────────────────
    case S.AWAIT_DETAIL_LEVEL: {
      if (text !== 'elaborated' && text !== 'concise') { botSay(`Please select one of the options above.`); return; }
      session.detailLevel = text;
      botSay(`**${text === 'elaborated' ? '📘 Elaborated — comprehensive and detailed' : '📄 Concise — focused and high-level'}** ✓`);
      askFitGap();
      break;
    }

    // ── Fit-Gap Analysis ──────────────────────────────────────────────────────
    case S.AWAIT_FIT_GAP: {
      if (text !== 'yes' && text !== 'no') { botSay(`Please select one of the options above.`); return; }
      session.fitGap = text;
      botSay(`**${text === 'yes' ? '✅ Fit-Gap Analysis will be included' : '⏭ Fit-Gap Analysis excluded'}** ✓`);
      session.state = S.AWAIT_SOURCE_REF;
      setTimeout(() => {
        disableInput();
        botAsk(
          `Do you want to include the source for each requirement in the BRD?\n\n_(If enabled, every requirement will be tagged with the document or section it was extracted from — great for traceability and audits.)_`,
          [
            { label: '✅ Yes, include source references', value: 'yes' },
            { label: '⏭ No, not required',                value: 'no'  },
          ]
        );
      }, 400);
      break;
    }

    // ── Source references ─────────────────────────────────────────────────────
    case S.AWAIT_SOURCE_REF: {
      if (text !== 'yes' && text !== 'no') { botSay(`Please select one of the options above.`); return; }
      session.sourceRef = text;
      botSay(`**${text === 'yes' ? '✅ Source references will be included' : '⏭ Source references excluded'}** ✓`);
      session.state = S.AWAIT_MOSCOW;
      setTimeout(() => {
        disableInput();
        botAsk(
          `Should I include a requirement priority or MoSCoW classification?\n\n_(Helps stakeholders understand what is critical versus nice-to-have.)_`,
          [
            { label: '✅ Yes', value: 'yes' },
            { label: '⏭ No',  value: 'no'  },
          ]
        );
      }, 400);
      break;
    }

    // ── MoSCoW classification ─────────────────────────────────────────────────
    case S.AWAIT_MOSCOW: {
      if (text !== 'yes' && text !== 'no') { botSay(`Please select one of the options above.`); return; }
      session.moscow = text;
      botSay(`**${text === 'yes' ? '✅ MoSCoW classification will be included' : '⏭ MoSCoW classification excluded'}** ✓`);
      askDocumentSource();
      break;
    }

    // ── Document source ───────────────────────────────────────────────────────
    case S.AWAIT_SOURCE: {
      if (text === 'sharepoint') {
        if (!session.spConfigured) {
          botSay(`⚠️ SharePoint is not configured on this server.\n\nPlease add **AZURE_TENANT_ID**, **AZURE_CLIENT_ID**, **AZURE_CLIENT_SECRET** to the *.env* file and restart.`);
          setTimeout(() => askDocumentSource(), 600);
          return;
        }
        session.source = 'sharepoint';
        session.state  = S.AWAIT_SP_URL;
        botSay(`📁 **SharePoint selected** ✓`);
        setTimeout(() => {
          botAsk(`🔗 Please enter your **SharePoint Site URL**:\n_(e.g. https://yourcompany.sharepoint.com/sites/SiteName)_`);
          enableInput('Paste SharePoint URL...');
        }, 400);
      } else if (text === 'upload') {
        session.source = 'upload';
        session.state  = S.AWAIT_FILES;
        botSay(`📤 **File upload selected** ✓`);
        setTimeout(() => {
          botAsk(`📎 Please **attach your documents** (PDF, DOCX, TXT, XLSX).\n\nUse the 📎 button to attach files, then type **done** to proceed.\nOr type **skip** to generate without documents.`);
          enableInput('Type done when ready, or skip...');
          enableAttach(true);
        }, 400);
      } else if (text === 'paste') {
        session.source = 'paste';
        session.state  = S.AWAIT_PASTE_CONTENT;
        botSay(`📝 **Paste content selected** ✓`);
        setTimeout(() => {
          botAsk(`📋 Please **paste your content** directly into the input below and press Send.\n\n_(You can paste meeting notes, requirements, or any relevant text.)_`);
          enableInput('Paste your content here and press Enter...');
        }, 400);
      } else {
        askDocumentSource();
      }
      break;
    }

    // ── Paste content ─────────────────────────────────────────────────────────
    case S.AWAIT_PASTE_CONTENT: {
      if (!text.trim()) { botSay(`Please paste some content before proceeding.`); return; }
      session.pasteContent = text.trim();
      botSay(`✅ **Content received** (${text.length} characters) ✓`);
      askAdditionalInputs();
      break;
    }

    // ── SharePoint URL ────────────────────────────────────────────────────────
    case S.AWAIT_SP_URL: {
      try { new URL(text); } catch {
        botSay(`⚠️ That doesn't look like a valid URL. Please enter a full URL like:\n\`https://yourcompany.sharepoint.com/sites/SiteName\``);
        return;
      }
      session.spSiteUrl = text.trim();
      session.state = S.AWAIT_SP_FOLDER;
      botSay(`URL saved ✓`);
      setTimeout(() => {
        botAsk(`📂 Which **folder name** should I look in?\n_(Press Enter for default: **BRD Document**)_`);
        chatInput.placeholder = 'BRD Document';
      }, 400);
      break;
    }

    // ── SharePoint folder ─────────────────────────────────────────────────────
    case S.AWAIT_SP_FOLDER: {
      session.spFolder = text.trim() || 'BRD Document';
      session.state = S.AWAIT_SP_CONFIRM;
      botSay(`Folder: **${session.spFolder}** ✓\n\nFetching files...`);
      disableInput();
      await fetchSharePointFiles();
      break;
    }

    // ── SharePoint file confirmation ──────────────────────────────────────────
    case S.AWAIT_SP_CONFIRM: {
      if (text === '__sp_confirm__') {
        askAdditionalInputs();
      } else if (text === '__sp_retry__') {
        session.state = S.AWAIT_SP_URL;
        botAsk(`🔗 Please enter a different SharePoint Site URL:`);
        enableInput('Paste SharePoint URL...');
      }
      break;
    }

    // ── File upload ───────────────────────────────────────────────────────────
    case S.AWAIT_FILES: {
      if (t === 'skip') {
        session.uploadFiles = [];
        clearAttachments();
        botSay(`Skipping documents ✓`);
        askAdditionalInputs();
      } else if (t === 'done' || t === '(files attached)') {
        if (session.uploadFiles.length === 0) {
          botSay(`No files attached yet. Use the 📎 button to attach files, or type **skip** to continue without documents.`);
        } else {
          const names = session.uploadFiles.map(f => `• 📄 ${f.name}`).join('\n');
          botSay(`✅ **${session.uploadFiles.length} file(s) ready:**\n${names}`);
          clearAttachUI();
          askAdditionalInputs();
        }
      } else {
        botSay(`Use the 📎 button to attach files, then type **done**.\nOr type **skip** to proceed without documents.`);
      }
      break;
    }

    // ── Additional inputs ─────────────────────────────────────────────────────
    case S.AWAIT_ADDITIONAL_INPUTS: {
      if (text !== 'yes' && text !== 'no') { botSay(`Please select one of the options above.`); return; }
      if (text === 'yes') {
        session.state = S.AWAIT_ADDITIONAL_TEXT;
        setTimeout(() => {
          botAsk(`📝 Please share your additional inputs — assumptions, constraints, or formatting preferences:`);
          enableInput('Type your additional inputs...');
        }, 400);
      } else {
        askOutputFormat();
      }
      break;
    }

    // ── Additional text ───────────────────────────────────────────────────────
    case S.AWAIT_ADDITIONAL_TEXT: {
      if (!text.trim()) { botSay(`Please type your additional inputs.`); return; }
      session.additionalInputs = text.trim();
      botSay(`✅ **Additional inputs noted** ✓`);
      askOutputFormat();
      break;
    }

    // ── Output format ─────────────────────────────────────────────────────────
    case S.AWAIT_OUTPUT_FORMAT: {
      if (!['excel', 'word', 'pdf'].includes(text)) { botSay(`Please select one of the options above.`); return; }
      session.outputFormat = text;
      const fmtLabel = { excel: '📊 Excel', word: '📝 Word (.docx)', pdf: '📄 PDF' }[text];
      botSay(`**${fmtLabel}** selected ✓`);
      setTimeout(() => startGeneration(), 400);
      break;
    }

    // ── Additional documents after BRD generation ────────────────────────────
    case S.AWAIT_UPDATE_DOCS: {
      if (t === 'yes') {
        botSay(`📎 Please **attach your additional documents** using the 📎 button, then type **done** to update the BRD.`);
        session.uploadFiles = [];
        clearAttachUI();
        session.brdType      = 'update';
        session.updateMethod = 'latest';
        session.source       = 'upload';
        session.pasteContent = '';
        session.state        = S.AWAIT_UPDATE_FILES;
        enableInput('Attach files, then type done...');
        enableAttach(true);
      } else if (t === 'no') {
        botSay('Thank you, have a great day ahead! 😊');
        session.state = S.DONE;
        disableInput();
      }
      break;
    }

    case S.AWAIT_UPDATE_FILES: {
      if (t === 'done' || text === '(files attached)') {
        if (!session.uploadFiles.length) {
          botSay(`⚠️ Please attach at least one document before proceeding.`);
          return;
        }
        const names = session.uploadFiles.map(f => `• ${f.name}`).join('\n');
        botSay(`✅ **${session.uploadFiles.length} document(s) ready:**\n${names}`);
        setTimeout(() => startGeneration(), 400);
      } else if (t === 'skip') {
        botSay('Skipping additional documents. Thank you, have a great day ahead! 😊');
        session.state = S.DONE;
        disableInput();
      } else {
        botSay(`Use the 📎 button to attach documents, then type **done**.\nOr type **skip** to finish without updating.`);
      }
      break;
    }

    default: break;
  }
}

// ── Helper: ask SOW (only for new BRDs) ──────────────────────────────────────
function askSOW() {
  disableInput();
  botAsk(
    `📄 Do you have a **Statement of Work (SOW)** document for this project?\n\n_(Uploading the SOW helps generate more accurate and project-specific requirements.)_`,
    [
      { label: '📎 Yes, I have a SOW document', value: 'yes' },
      { label: '⏭ No, continue without SOW',   value: 'no'  },
    ]
  );
}

function askDetailLevel() {
  session.state = S.AWAIT_DETAIL_LEVEL;
  setTimeout(() => {
    disableInput();
    botAsk(
      `Please specify your preferred level of detail — how detailed should the BRD be?\n\n_(This controls the depth of each section — from a concise executive overview to a fully elaborated document.)_`,
      [
        { label: '📘 Elaborated (comprehensive and detailed documentation)', value: 'elaborated' },
        { label: '📄 Concise (focused and high-level documentation)',         value: 'concise'    },
      ]
    );
  }, 400);
}

// ── Helper: ask fit-gap (shared by update + new paths) ───────────────────────
function askFitGap() {
  session.state = S.AWAIT_FIT_GAP;
  setTimeout(() => {
    disableInput();
    botAsk(
      `Would you like to include a **Fit-Gap Analysis** as part of this BRD?\n\n_(This section compares current-state capabilities against the requirements and flags where gaps exist.)_`,
      [
        { label: '✅ Yes, include Fit-Gap Analysis', value: 'yes' },
        { label: '⏭ No, exclude Fit-Gap Analysis',  value: 'no'  },
      ]
    );
  }, 400);
}

// ── Helper: ask document source ───────────────────────────────────────────────
function askDocumentSource() {
  session.state = S.AWAIT_SOURCE;
  disableInput();
  botAsk(
    `📂 Please select your preferred approach for submitting source documents.\n\n_(You can upload multiple files at once — I'll extract requirements from all of them.)_`,
    [
      { label: '📁 From SharePoint',        value: 'sharepoint' },
      { label: '📤 Upload Files',            value: 'upload'     },
      { label: '📋 Paste Content Directly',  value: 'paste'      },
    ]
  );
}

// ── Helper: ask additional inputs ─────────────────────────────────────────────
function askAdditionalInputs() {
  session.state = S.AWAIT_ADDITIONAL_INPUTS;
  enableAttach(false);
  setTimeout(() => {
    disableInput();
    botAsk(
      `Would you like to include any additional inputs such as **assumptions**, **constraints**, or **specific formatting preferences**?`,
      [
        { label: '✅ Yes, add additional inputs', value: 'yes' },
        { label: '⏭ No, proceed as is',           value: 'no'  },
      ]
    );
  }, 400);
}

// ── Helper: ask output format ─────────────────────────────────────────────────
function askOutputFormat() {
  session.state = S.AWAIT_OUTPUT_FORMAT;
  setTimeout(() => {
    disableInput();
    botAsk(
      `What format would you like the final BRD delivered in?`,
      [
        { label: '📊 Excel',        value: 'excel' },
        { label: '📝 Word (.docx)', value: 'word'  },
        { label: '📄 PDF',          value: 'pdf'   },
      ]
    );
  }, 400);
}

// ── Fetch SharePoint files ────────────────────────────────────────────────────
async function fetchSharePointFiles() {
  const typingId = addTyping();

  try {
    const res  = await fetch('/api/sharepoint/list-files', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ siteUrl: session.spSiteUrl, folderPath: session.spFolder }),
    });
    const data = await res.json();
    removeTyping(typingId);

    if (!res.ok) {
      botSay(`❌ ${data.error || 'Could not connect to SharePoint.'}`);
      session.state = S.AWAIT_SOURCE;
      setTimeout(() => askDocumentSource(), 500);
      return;
    }

    session.spFiles = data.files || [];
    if (!session.spFiles.length) {
      botSay(`📂 No supported files (PDF, DOCX, TXT) found in **${session.spFolder}**.\n\nWould you like to try a different source?`);
      session.state = S.AWAIT_SOURCE;
      setTimeout(() => askDocumentSource(), 500);
      return;
    }

    addSpFilePicker(session.spFiles);
    session.state = S.AWAIT_SP_CONFIRM;

  } catch (err) {
    removeTyping(typingId);
    botSay(`❌ Network error: ${err.message}`);
    session.state = S.AWAIT_SOURCE;
    setTimeout(() => askDocumentSource(), 500);
  }
}

// ── Generation ────────────────────────────────────────────────────────────────
async function startGeneration() {
  session.state = S.GENERATING;
  const snapshotLog = [...chatLog]; // capture full chat before generation messages
  _chatLogActive = false;
  disableInput();
  enableAttach(false);

  const fmtLabel = { excel: 'Excel', word: 'Word document', pdf: 'PDF' }[session.outputFormat] || 'document';
  botSay(`⚡ Starting BRD generation for **${session.projectName}**...`);

  const progRowId = addProgressCard(fmtLabel);

  // Build shared metadata params
  const meta = {
    projectName:      session.projectName,
    authorName:       session.authorName,
    brdType:          session.brdType,
    updateMethod:     session.updateMethod,
    detailLevel:      session.detailLevel,
    fitGap:           session.fitGap,
    sourceRef:        session.sourceRef,
    moscow:           session.moscow,
    additionalInputs: session.additionalInputs,
    outputFormat:     session.outputFormat,
  };

  try {
    let data;

    if (session.source === 'sharepoint') {
      const selected = getSelectedSpFiles();
      if (!selected.length) {
        botSay(`⚠️ No files selected. Please select at least one file.`);
        session.state = S.AWAIT_SP_CONFIRM;
        return;
      }
      updateProgress(progRowId, 1, `Fetching ${selected.length} file(s) from SharePoint...`);
      await delay(800);
      updateProgress(progRowId, 2, 'AI generating BRD...');
      await delay(600);
      updateProgress(progRowId, 3, `Building ${fmtLabel}...`);

      const res = await fetch('/api/generate-brd-sharepoint', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ ...meta, files: selected, chatLog: snapshotLog }),
      });
      data = await res.json();
      if (!res.ok || !data.success) throw new Error(data.error || 'Generation failed');

    } else {
      const docCount = session.source === 'paste'
        ? (session.pasteContent ? 1 : 0)
        : session.uploadFiles.length;

      updateProgress(progRowId, 1, `Processing ${docCount} document(s)...`);
      await delay(800);
      updateProgress(progRowId, 2, 'AI generating BRD...');
      await delay(600);
      updateProgress(progRowId, 3, `Building ${fmtLabel}...`);

      const fd = new FormData();
      Object.entries(meta).forEach(([k, v]) => fd.append(k, v || ''));
      if (session.pasteContent) fd.append('pasteContent', session.pasteContent);
      fd.append('chatLog', JSON.stringify(snapshotLog));
      for (const f of [...(session.sowFiles || []), ...session.uploadFiles]) fd.append('files', f);

      const res = await fetch('/api/generate-brd', { method: 'POST', body: fd });
      data = await res.json();
      if (!res.ok || !data.success) throw new Error(data.error || 'Generation failed');
    }

    finishProgress(progRowId);
    await delay(400);
    showResult(data);
    refreshSidebar();
    session.state = S.AWAIT_UPDATE_DOCS;
    setTimeout(askUpdateDocs, 800);

  } catch (err) {
    finishProgress(progRowId, true);
    botSay(`❌ **Generation failed:** ${err.message}\n\nType **retry** to try again or **new** to start over.`);
    session.state = S.DONE;
    enableInput('Type retry or new...');
    chatInput.addEventListener('keydown', function retryHandler(e) {
      if (e.key === 'Enter') {
        const v = chatInput.value.trim().toLowerCase();
        if (v === 'retry') { chatInput.removeEventListener('keydown', retryHandler); startConversation(); }
        if (v === 'new')   { chatInput.removeEventListener('keydown', retryHandler); startConversation(); }
      }
    });
  }
}

// ── Show result ───────────────────────────────────────────────────────────────
function showResult(data) {
  const s    = data.summary;
  const reqs = data.requirements || [];

  // Store BRD context for AI chat
  aiBRDContext = { ...s, requirements: reqs };

  const fileCount   = (session.sowFiles || []).length + session.uploadFiles.length;
  const sowNote     = (session.sowFiles || []).length ? ` (incl. SOW)` : '';
  const sourceLabel = session.source === 'sharepoint' ? 'SharePoint' :
                      session.source === 'paste'       ? 'Pasted content' :
                      fileCount > 0                    ? `${fileCount} document${fileCount > 1 ? 's' : ''} uploaded${sowNote}`
                                                       : 'Project name only';

  const fmtIcon  = { excel: '📊', word: '📝', pdf: '📄' }[session.outputFormat] || '📊';
  const fmtLabel = { excel: 'Excel', word: 'Word (.docx)', pdf: 'PDF' }[session.outputFormat] || 'Excel';

  const u = data.usage || {};
  const usageHtml = u.total_tokens ? `
    <div class="result-usage">
      <span class="usage-item">📥 Input: <strong>${(u.input_tokens||0).toLocaleString()}</strong></span>
      <span class="usage-sep">·</span>
      <span class="usage-item">📤 Output: <strong>${(u.output_tokens||0).toLocaleString()}</strong></span>
      <span class="usage-sep">·</span>
      <span class="usage-item">🔢 Total: <strong>${(u.total_tokens||0).toLocaleString()}</strong> tokens</span>
      <span class="usage-sep">·</span>
      <span class="usage-item">💰 Cost: <strong>$${(u.cost_usd||0).toFixed(4)}</strong></span>
    </div>` : '';

  const card = el('div', 'result-card');
  card.innerHTML = `
    <div class="result-title">✅ BRD Generated Successfully!</div>
    <div class="result-stats">
      <div class="result-stat">
        <div class="result-stat-val">${s.total}</div>
        <div class="result-stat-lbl">Requirements</div>
      </div>
      <div class="result-stat high">
        <div class="result-stat-val">${s.mustHave ?? s.high ?? 0}</div>
        <div class="result-stat-lbl">Must Have</div>
      </div>
      <div class="result-stat med">
        <div class="result-stat-val">${s.shouldHave ?? s.medium ?? 0}</div>
        <div class="result-stat-lbl">Should Have</div>
      </div>
      <div class="result-stat low">
        <div class="result-stat-val">${s.couldHave ?? s.low ?? 0}</div>
        <div class="result-stat-lbl">Could Have</div>
      </div>
    </div>
    ${usageHtml}
    <div class="result-meta">
      <strong>Project:</strong> ${esc(s.projectName)} &nbsp;·&nbsp;
      <strong>Date:</strong> ${esc(s.date)} &nbsp;·&nbsp;
      <strong>Version:</strong> ${esc(s.version)} &nbsp;·&nbsp;
      <strong>Source:</strong> ${esc(sourceLabel)}
    </div>
    <a class="result-dl" href="${esc(data.downloadUrl)}" download="${esc(data.fileName)}">
      ${fmtIcon} Download BRD ${fmtLabel} — ${esc(data.fileName)}
    </a>
  `;

  addBotRow(card);
  scrollBottom();
  disableInput();
}

// ── Post-generation: ask if user has additional documents ─────────────────────
function askUpdateDocs() {
  disableInput();
  botAsk(
    `📎 Do you have **additional documents** to update this BRD?`,
    [
      { label: '📎 Yes, I have more documents', value: 'yes' },
      { label: '👍 No, I\'m done',              value: 'no'  },
    ]
  );
}

// ── Chip click handler (quick replies) ───────────────────────────────────────
document.addEventListener('click', (e) => {
  const chip = e.target.closest('.qr-chip');
  if (!chip || chip.classList.contains('chosen')) return;

  chip.closest('.quick-replies')?.querySelectorAll('.qr-chip').forEach(c => {
    c.classList.add('chosen');
  });

  const value = chip.dataset.value;

  if (value === '__new__')  { startConversation(); return; }
  if (value === '__done__') { botSay('Thanks! Have a great day. 👋'); disableInput(); return; }

  if (!value.startsWith('__')) {
    addUserMsg(chip.textContent.trim());
  } else {
    addUserMsg(chip.dataset.label || chip.textContent.trim());
  }

  processInput(value);
});

// ── File attachment ───────────────────────────────────────────────────────────
attachBtn.addEventListener('click', () => {
  if (session.state !== S.AWAIT_FILES &&
      session.state !== S.AWAIT_UPDATE_FILES &&
      session.state !== S.AWAIT_SOW_FILES) {
    botSay('📎 File attachment is available after you choose **Upload Files** as the document source.');
    return;
  }
  fileInput.click();
});

fileInput.addEventListener('change', () => {
  addFilesToSession(Array.from(fileInput.files));
  fileInput.value = '';
});

function addFilesToSession(files) {
  const allowed = ['.pdf', '.docx', '.txt', '.doc', '.xlsx', '.xls'];
  const isSow   = session.state === S.AWAIT_SOW_FILES;
  const bucket  = isSow ? session.sowFiles : session.uploadFiles;
  let added = 0;
  for (const f of files) {
    const ext = '.' + f.name.split('.').pop().toLowerCase();
    if (!allowed.includes(ext)) continue;
    if (bucket.some(x => x.name === f.name)) continue;
    bucket.push(f);
    added++;
  }
  if (added) renderAttachPreview();
  if (added) {
    if (isSow) {
      botSay(`📎 **${session.sowFiles.length} SOW file(s) attached.**\nAttach more, or type **done** to continue.`);
    } else if (session.state === S.AWAIT_FILES || session.state === S.AWAIT_UPDATE_FILES) {
      const label = session.state === S.AWAIT_UPDATE_FILES ? 'update the BRD' : 'start generation';
      botSay(`📎 **${session.uploadFiles.length} file(s) attached.**\nAttach more, or type **done** to ${label}.`);
    }
  }
}

function renderAttachPreview() {
  const isSow  = session.state === S.AWAIT_SOW_FILES;
  const bucket = isSow ? session.sowFiles : session.uploadFiles;
  attachPreviewIn.innerHTML = '';
  bucket.forEach((f, i) => {
    const chip = el('div', 'attach-chip');
    chip.innerHTML = `${fileIcon(f.name)} ${esc(f.name)} <button class="attach-chip-remove" data-idx="${i}" data-sow="${isSow ? '1' : '0'}">✕</button>`;
    attachPreviewIn.appendChild(chip);
  });
  attachPreview.classList.toggle('hidden', bucket.length === 0);
  attachBtn.classList.toggle('has-files', bucket.length > 0);
}

document.addEventListener('click', (e) => {
  const btn = e.target.closest('.attach-chip-remove');
  if (!btn) return;
  const bucket = btn.dataset.sow === '1' ? session.sowFiles : session.uploadFiles;
  bucket.splice(Number(btn.dataset.idx), 1);
  renderAttachPreview();
});

function clearAttachUI() {
  attachPreviewIn.innerHTML = '';
  attachPreview.classList.add('hidden');
  attachBtn.classList.remove('has-files');
  enableAttach(false);
}

function clearAttachments() {
  session.uploadFiles = [];
  clearAttachUI();
}

// ── SharePoint file picker ────────────────────────────────────────────────────
function addSpFilePicker(files) {
  const wrap = el('div', 'sp-picker');
  wrap.innerHTML = `
    <div class="sp-picker-header">
      <span>📁 ${files.length} file(s) found in <strong>${esc(session.spFolder)}</strong></span>
      <label class="sp-check-all"><input type="checkbox" id="spAll" checked /> Select all</label>
    </div>
    <ul class="sp-picker-list" id="spPickerList">
      ${files.map((f, i) => `
        <li class="sp-picker-item">
          <input type="checkbox" data-idx="${i}" checked />
          <span class="sp-picker-name">${fileIcon(f.name)} ${esc(f.name)}</span>
          <span class="sp-picker-size">${fmtSize(f.size)}</span>
        </li>
      `).join('')}
    </ul>
    <button class="sp-confirm-btn" id="spConfirmBtn">✅ Generate BRD with selected files</button>
  `;
  addBotRow(wrap);
  scrollBottom();

  document.getElementById('spAll')?.addEventListener('change', (e) => {
    wrap.querySelectorAll('.sp-picker-list input').forEach(cb => cb.checked = e.target.checked);
  });

  document.getElementById('spConfirmBtn')?.addEventListener('click', () => {
    document.getElementById('spConfirmBtn').disabled = true;
    document.getElementById('spConfirmBtn').textContent = 'Processing...';
    processInput('__sp_confirm__');
  });
}

function getSelectedSpFiles() {
  const list = document.getElementById('spPickerList');
  if (!list) return session.spFiles;
  const checked = Array.from(list.querySelectorAll('input:checked'));
  return checked.map(cb => session.spFiles[Number(cb.dataset.idx)]);
}

// ── Progress card ─────────────────────────────────────────────────────────────
function addProgressCard(fmtLabel = 'Document') {
  const id   = 'prog_' + Date.now();
  const card = el('div', 'prog-card');
  card.id = id;
  card.innerHTML = `
    <div class="prog-row active" id="${id}_1">
      <div class="prog-icon"><span class="prog-spinner"></span></div>
      <div class="prog-label" id="${id}_1_lbl">Initializing...</div>
    </div>
    <div class="prog-row wait" id="${id}_2">
      <div class="prog-icon">○</div>
      <div class="prog-label">AI Generation</div>
    </div>
    <div class="prog-row wait" id="${id}_3">
      <div class="prog-icon">○</div>
      <div class="prog-label">Building ${fmtLabel}</div>
    </div>
  `;
  addBotRow(card);
  scrollBottom();
  return id;
}

function updateProgress(id, step, label) {
  for (let i = 1; i <= 3; i++) {
    const row = document.getElementById(`${id}_${i}`);
    if (!row) continue;
    row.classList.remove('active', 'wait', 'done');
    if (i < step)  { row.classList.add('done');  row.querySelector('.prog-icon').innerHTML = '✓'; }
    if (i === step){ row.classList.add('active'); row.querySelector('.prog-icon').innerHTML = '<span class="prog-spinner"></span>'; }
    if (i > step)  { row.classList.add('wait');  row.querySelector('.prog-icon').innerHTML = '○'; }
  }
  const lbl = document.getElementById(`${id}_${step}_lbl`);
  if (lbl) lbl.textContent = label;
  scrollBottom();
}

function finishProgress(id, error = false) {
  for (let i = 1; i <= 3; i++) {
    const row = document.getElementById(`${id}_${i}`);
    if (!row) continue;
    row.classList.remove('active', 'wait');
    if (error) {
      row.classList.add('wait');
      row.querySelector('.prog-icon').innerHTML = '—';
    } else {
      row.classList.add('done');
      row.querySelector('.prog-icon').innerHTML = '✓';
    }
  }
  scrollBottom();
}

// ── Message rendering helpers ─────────────────────────────────────────────────
function botSay(text) {
  if (_chatLogActive) chatLog.push({ role: 'bot', type: 'say', text });
  const bubble = el('div', 'msg-bubble');
  bubble.innerHTML = mdToHtml(text);
  addBotRow(bubble);
  scrollBottom();
}

function botAsk(text, chips = []) {
  if (_chatLogActive) chatLog.push({ role: 'bot', type: 'ask', text });
  const wrap   = el('div', '');
  const bubble = el('div', 'msg-bubble');
  bubble.innerHTML = mdToHtml(text);
  wrap.appendChild(bubble);

  if (chips.length) {
    const qr = el('div', 'quick-replies');
    chips.forEach(c => {
      const btn = el('button', 'qr-chip');
      btn.textContent   = c.label;
      btn.dataset.value = c.value;
      btn.dataset.label = c.label;
      qr.appendChild(btn);
    });
    wrap.appendChild(qr);
  }

  addBotRow(wrap);
  scrollBottom();
}

function addBotRow(content) {
  const row = el('div', 'msg-row bot');
  const av  = el('div', 'msg-avatar');
  const img = document.createElement('img');
  img.src = 'synoptek-logo.png'; img.alt = 'S';
  av.appendChild(img);

  const col  = el('div', 'msg-col');
  col.appendChild(content);
  const time = el('div', 'msg-time');
  time.textContent = now();
  col.appendChild(time);

  row.appendChild(av);
  row.appendChild(col);
  chatMessages.appendChild(row);
}

function addUserMsg(text) {
  if (_chatLogActive) chatLog.push({ role: 'user', text });
  const row    = el('div', 'msg-row user');

  const col    = el('div', 'msg-col');
  const bubble = el('div', 'msg-bubble');
  bubble.textContent = text;
  const time   = el('div', 'msg-time');
  time.textContent = now();

  col.appendChild(bubble);
  col.appendChild(time);
  row.appendChild(col);
  chatMessages.appendChild(row);
  scrollBottom();
}

function addTyping() {
  const id  = 'typing_' + Date.now();
  const row = el('div', 'msg-row bot');
  row.id = id;

  const av  = el('div', 'msg-avatar');
  const img = document.createElement('img');
  img.src = 'synoptek-logo.png'; img.alt = 'S';
  av.appendChild(img);

  const bubble = el('div', 'typing-bubble');
  bubble.innerHTML = '<div class="typing-dot"></div><div class="typing-dot"></div><div class="typing-dot"></div>';

  row.appendChild(av);
  row.appendChild(bubble);
  chatMessages.appendChild(row);
  scrollBottom();
  return id;
}

function removeTyping(id) {
  document.getElementById(id)?.remove();
}

function addDateSep(label) {
  const sep = el('div', 'date-sep');
  sep.textContent = label;
  chatMessages.appendChild(sep);
}

// ── Input control ─────────────────────────────────────────────────────────────
function enableInput(placeholder = 'Type a message...') {
  chatInput.disabled    = false;
  sendBtn.disabled      = false;
  chatInput.placeholder = placeholder;
  chatInput.focus();
}
function disableInput() {
  chatInput.disabled    = true;
  sendBtn.disabled      = true;
  chatInput.placeholder = 'Please wait...';
}
function enableAttach(on) {
  fileInput.disabled = !on;
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function el(tag, cls) {
  const e = document.createElement(tag || 'div');
  if (cls) e.className = cls;
  return e;
}
function scrollBottom() {
  requestAnimationFrame(() => { chatMessages.scrollTop = chatMessages.scrollHeight; });
}
function delay(ms) { return new Promise(r => setTimeout(r, ms)); }
function now() {
  return new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
}
function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
function fileIcon(name) {
  const ext = (name || '').split('.').pop().toLowerCase();
  return { pdf: '📕', docx: '📘', doc: '📘', txt: '📄', xlsx: '📊', xls: '📊' }[ext] || '📄';
}
function fmtSize(bytes) {
  if (!bytes) return '—';
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' KB';
  return (bytes / 1024 / 1024).toFixed(1) + ' MB';
}

function mdToHtml(text) {
  return esc(text)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/_(.+?)_/g,       '<em>$1</em>')
    .replace(/`(.+?)`/g,       '<code>$1</code>')
    .replace(/\n/g,            '<br/>');
}

// ── Sidebar ───────────────────────────────────────────────────────────────────
const sidebar             = document.getElementById('sidebar');
const sidebarList         = document.getElementById('sidebarList');
const sidebarToggle       = document.getElementById('sidebarToggle');
const sidebarHeaderToggle = document.getElementById('sidebarHeaderToggle');

let activeSidebarId = null;

function toggleSidebar() {
  sidebar.classList.toggle('collapsed');
}
sidebarToggle.addEventListener('click', toggleSidebar);
sidebarHeaderToggle.addEventListener('click', toggleSidebar);

async function loadSessions() {
  try {
    const res  = await fetch('/api/sessions');
    const data = await res.json();
    renderSidebar(data.sessions || []);
  } catch { /* ignore */ }
}

function renderSidebar(sessions) {
  if (!sessions.length) {
    sidebarList.innerHTML = '<div class="sidebar-empty">No BRDs generated yet.<br/>Your history will appear here.</div>';
    return;
  }
  sidebarList.innerHTML = '';
  sessions.forEach(s => {
    const icon = { excel: '📊', word: '📝', pdf: '📄' }[s.outputFormat] || '📊';
    const date = new Date(s.createdAt).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    const item = document.createElement('div');
    item.className = 'sidebar-item' + (s.id === activeSidebarId ? ' active' : '');
    item.dataset.id = s.id;
    const ver = (s.version || s.summary?.version || 'v1.0').toUpperCase();
    item.innerHTML = `
      <div class="sidebar-item-icon">${icon}</div>
      <div class="sidebar-item-info">
        <div class="sidebar-item-name">${esc(s.projectName || 'Untitled')} <span style="font-size:10px;font-weight:600;background:#1d4ed8;color:#fff;border-radius:3px;padding:1px 5px;vertical-align:middle">${esc(ver)}</span></div>
        <div class="sidebar-item-meta">${date} · ${(s.summary?.total || 0)} reqs</div>
      </div>`;
    item.addEventListener('click', () => openSession(s));
    sidebarList.appendChild(item);
  });
}

function openSession(s) {
  activeSidebarId = s.id;
  sidebarList.querySelectorAll('.sidebar-item').forEach(el => {
    el.classList.toggle('active', el.dataset.id === s.id);
  });

  session = createSession();
  session.state = S.DONE;
  chatLog = [];
  _chatLogActive = false; // don't track replayed messages
  clearAttachments();
  chatMessages.innerHTML = '';

  const sessionDate = new Date(s.createdAt).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
  addDateSep(sessionDate);

  // Replay full conversation if saved
  if (s.chatLog && s.chatLog.length) {
    s.chatLog.forEach(msg => {
      if (msg.role === 'user') {
        addUserMsg(msg.text);
      } else {
        // Render bot messages as plain bubbles (no action chips in history)
        const bubble = el('div', 'msg-bubble');
        bubble.innerHTML = mdToHtml(msg.text);
        addBotRow(bubble);
      }
    });
  }

  const icon  = { excel: '📊', word: '📝', pdf: '📄' }[s.outputFormat] || '📊';
  const label = { excel: 'Excel', word: 'Word (.docx)', pdf: 'PDF' }[s.outputFormat] || 'Excel';
  const sm    = s.summary || {};

  const card = el('div', 'result-card');
  card.innerHTML = `
    <div class="result-title">✅ ${esc(s.projectName || 'BRD')} — Generated</div>
    <div class="result-stats">
      <div class="result-stat"><div class="result-stat-val">${sm.total || 0}</div><div class="result-stat-lbl">Requirements</div></div>
      <div class="result-stat high"><div class="result-stat-val">${sm.mustHave ?? sm.high ?? 0}</div><div class="result-stat-lbl">Must Have</div></div>
      <div class="result-stat med"><div class="result-stat-val">${sm.shouldHave ?? sm.medium ?? 0}</div><div class="result-stat-lbl">Should Have</div></div>
      <div class="result-stat low"><div class="result-stat-val">${sm.couldHave ?? sm.low ?? 0}</div><div class="result-stat-lbl">Could Have</div></div>
    </div>
    <div class="result-meta">
      <strong>Author:</strong> ${esc(s.authorName || '—')} &nbsp;·&nbsp;
      <strong>Version:</strong> ${esc(sm.version || 'v1.0')} &nbsp;·&nbsp;
      <strong>Date:</strong> ${esc(sm.date || '')}
    </div>
    <a class="result-dl" href="${esc(s.downloadUrl)}" download="${esc(s.fileName)}">
      ${icon} Download ${label} — ${esc(s.fileName)}
    </a>`;
  addBotRow(card);

  botAsk('Would you like to generate a new BRD?', [
    { label: '✅ Yes, new BRD', value: '__new__' },
  ]);
  disableInput();
  enableInput('Type your choice...');
  scrollBottom();
}

// Refresh sidebar after each generation
function refreshSidebar() { loadSessions(); }

// ── AI Chat Panel ─────────────────────────────────────────────────────────────
const aiPanel      = document.getElementById('aiPanel');
const aiMessagesEl = document.getElementById('aiMessages');
const aiInputEl    = document.getElementById('aiInput');
const aiSendBtnEl  = document.getElementById('aiSendBtn');
const aiPanelClose = document.getElementById('aiPanelClose');
const aiChatToggle = document.getElementById('aiChatToggle');

let aiHistory    = []; // { role: 'user'|'assistant', content: string }
let aiBRDContext = null; // set after BRD is generated

// Toggle panel open/close
aiChatToggle.addEventListener('click', () => {
  const isCollapsed = aiPanel.classList.toggle('collapsed');
  aiChatToggle.textContent = isCollapsed ? '💬 AI Chat' : '✕ AI Chat';
  if (!isCollapsed) aiInputEl.focus();
});
aiPanelClose.addEventListener('click', () => {
  aiPanel.classList.add('collapsed');
  aiChatToggle.textContent = '💬 AI Chat';
});

// Send on button click or Enter
aiSendBtnEl.addEventListener('click', sendAiMessage);
aiInputEl.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendAiMessage(); }
});

async function sendAiMessage() {
  const text = aiInputEl.value.trim();
  if (!text || aiSendBtnEl.disabled) return;
  aiInputEl.value = '';

  aiHistory.push({ role: 'user', content: text });
  appendAiMsg('user', text);

  const typingEl = appendAiTyping();
  aiSendBtnEl.disabled = true;
  aiInputEl.disabled   = true;

  try {
    const res  = await fetch('/api/chat', {
      method:  'POST',
      headers: { 'content-type': 'application/json' },
      body:    JSON.stringify({ messages: aiHistory, brdContext: aiBRDContext }),
    });
    const data = await res.json();
    if (data.error) throw new Error(data.error);

    typingEl.remove();
    aiHistory.push({ role: 'assistant', content: data.reply });
    appendAiMsg('assistant', data.reply);
  } catch (err) {
    typingEl.remove();
    appendAiMsg('assistant', `❌ ${err.message}`);
  } finally {
    aiSendBtnEl.disabled = false;
    aiInputEl.disabled   = false;
    aiInputEl.focus();
  }
}

function appendAiMsg(role, text) {
  // Remove welcome message on first real message
  const welcome = aiMessagesEl.querySelector('.ai-welcome');
  if (welcome) welcome.remove();

  const msg = document.createElement('div');
  msg.className = `ai-msg ${role}`;
  msg.innerHTML = mdToHtml(text);
  aiMessagesEl.appendChild(msg);
  aiMessagesEl.scrollTop = aiMessagesEl.scrollHeight;
  return msg;
}

function appendAiTyping() {
  const msg = document.createElement('div');
  msg.className = 'ai-msg assistant ai-typing';
  msg.innerHTML = '<span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span>';
  aiMessagesEl.appendChild(msg);
  aiMessagesEl.scrollTop = aiMessagesEl.scrollHeight;
  return msg;
}

// ── Settings & Usage Modal ────────────────────────────────────────────────────
const settingsOverlay = document.getElementById('settingsOverlay');
const settingsBtn     = document.getElementById('settingsBtn');
const settingsClose   = document.getElementById('settingsClose');

settingsBtn.addEventListener('click', () => { openSettings(); });
settingsClose.addEventListener('click', () => { settingsOverlay.classList.add('hidden'); });
settingsOverlay.addEventListener('click', e => { if (e.target === settingsOverlay) settingsOverlay.classList.add('hidden'); });

// Tab switching
document.querySelectorAll('.stab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.stab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.stab-content').forEach(c => c.classList.add('hidden'));
    btn.classList.add('active');
    const tab = document.getElementById('stab-' + btn.dataset.tab);
    if (tab) { tab.classList.remove('hidden'); tab.classList.add('active'); }
  });
});

async function openSettings() {
  settingsOverlay.classList.remove('hidden');
  await Promise.all([loadUsageTab(), loadConfigTab()]);
}

async function loadUsageTab() {
  try {
    const res  = await fetch('/api/usage');
    const data = await res.json();
    const t    = data.totals || {};
    document.getElementById('utBRDs').textContent   = (t.brd_count   || 0).toLocaleString();
    document.getElementById('utInput').textContent  = (t.input_tokens || 0).toLocaleString();
    document.getElementById('utOutput').textContent = (t.output_tokens|| 0).toLocaleString();
    document.getElementById('utCost').textContent   = '$' + (t.cost_usd || 0).toFixed(4);

    const tbody = document.getElementById('usageTbody');
    const entries = (data.entries || []).slice().reverse();
    if (!entries.length) {
      tbody.innerHTML = '<tr><td colspan="7" class="usage-empty">No usage data yet. Generate a BRD to start tracking.</td></tr>';
      return;
    }
    tbody.innerHTML = entries.map((e, i) => {
      const dt = new Date(e.timestamp).toLocaleString('en-GB', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
      return `<tr>
        <td>${entries.length - i}</td>
        <td>${esc(dt)}</td>
        <td>${esc(e.project || '—')}</td>
        <td>${(e.input_tokens  || 0).toLocaleString()}</td>
        <td>${(e.output_tokens || 0).toLocaleString()}</td>
        <td>${(e.total_tokens  || 0).toLocaleString()}</td>
        <td>$${(e.cost_usd    || 0).toFixed(4)}</td>
      </tr>`;
    }).join('');
  } catch { document.getElementById('usageTbody').innerHTML = '<tr><td colspan="7" class="usage-empty">Could not load usage data.</td></tr>'; }
}

async function loadConfigTab() {
  try {
    const res  = await fetch('/health');
    const data = await res.json();
    document.getElementById('cfgApiKey').textContent   = data.ai       || '—';
    document.getElementById('cfgSP').textContent       = data.sharepoint === 'configured' ? '✅ Configured' : '❌ Not configured';
    document.getElementById('cfgOutput').textContent   = data.output   || '—';
    document.getElementById('cfgVersion').textContent  = data.version || '2.0.36';
  } catch { document.getElementById('cfgVersion').textContent = '—'; }
}
