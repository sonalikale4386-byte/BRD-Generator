/**
 * app.js — BRD Generator Chatbot
 *
 * Conversation state machine:
 *   IDLE → AWAIT_PROJECT → AWAIT_AUTHOR → AWAIT_SOURCE
 *     ├─ sharepoint → AWAIT_SP_URL → AWAIT_SP_FOLDER → AWAIT_SP_CONFIRM → GENERATING
 *     ├─ upload     → AWAIT_FILES  → GENERATING
 *     └─ skip       → GENERATING
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
  IDLE:           'idle',
  AWAIT_PROJECT:  'await_project',
  AWAIT_AUTHOR:   'await_author',
  AWAIT_SOURCE:   'await_source',
  AWAIT_SP_URL:   'await_sp_url',
  AWAIT_SP_FOLDER:'await_sp_folder',
  AWAIT_SP_CONFIRM:'await_sp_confirm',
  AWAIT_FILES:    'await_files',
  GENERATING:     'generating',
  DONE:           'done',
};

// ── Session ───────────────────────────────────────────────────────────────────
let session = createSession();
function createSession() {
  return {
    state:       S.IDLE,
    projectName: '',
    authorName:  '',
    source:      null,
    spSiteUrl:   '',
    spFolder:    'BRD Document',
    spFiles:     [],    // fetched from SharePoint
    uploadFiles: [],    // File objects
    spConfigured: false,
  };
}

// ── Init ──────────────────────────────────────────────────────────────────────
(async () => {
  // Check health + SharePoint status
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

  // Start the conversation
  startConversation();
})();

// ── Start / Reset ─────────────────────────────────────────────────────────────
function startConversation() {
  chatMessages.innerHTML = '';
  session = createSession();
  session.spConfigured = spBadge.classList.contains('ok');
  clearAttachments();

  addDateSep('Today');

  botSay(`👋 Hi! I'm the **Synoptek BRD Generator**.\n\nI'll guide you through creating a professional Business Requirements Document in Synoptek CRM format.`);

  setTimeout(() => {
    botAsk(`📋 Let's start — what is the **Project Name**?`);
    session.state = S.AWAIT_PROJECT;
    enableInput('Type the project name...');
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

  // If waiting for files and user sends with attachments
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

    // ── Project name ─────────────────────────────────────────────────────────
    case S.AWAIT_PROJECT: {
      if (!text.trim()) { botSay('Please enter a project name.'); return; }
      session.projectName = text.trim();
      session.state = S.AWAIT_AUTHOR;
      botSay(`Got it — **${session.projectName}** ✓`);
      setTimeout(() => {
        botAsk(`👤 Who is preparing this BRD? _(Type your name or click Skip)_`, [
          { label: 'Skip', value: '__skip__' },
        ]);
        enableInput('Type your name...');
      }, 400);
      break;
    }

    // ── Author name ──────────────────────────────────────────────────────────
    case S.AWAIT_AUTHOR: {
      session.authorName = (text === '__skip__' || !text.trim()) ? 'BRD Agent' : text.trim();
      session.state = S.AWAIT_SOURCE;
      if (text !== '__skip__') botSay(`Prepared by **${session.authorName}** ✓`);
      setTimeout(() => askDocumentSource(), 400);
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
          botAsk(`📎 Please **attach your documents** (PDF, DOCX, TXT).\n\nUse the 📎 button to attach files, then type **done** to proceed.\nOr type **skip** to generate from project name only.`);
          enableInput('Type done when ready, or skip...');
          enableAttach(true);
        }, 400);
      } else if (text === 'skip') {
        session.source = 'skip';
        botSay(`📝 **No documents** — generating from project name only.`);
        setTimeout(() => startGeneration(), 400);
      } else {
        askDocumentSource();
      }
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
        startGeneration();
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
        setTimeout(() => startGeneration(), 400);
      } else if (t === 'done' || t === '(files attached)') {
        if (session.uploadFiles.length === 0) {
          botSay(`No files attached yet. Use the 📎 button to attach files, or type **skip** to continue without documents.`);
        } else {
          const names = session.uploadFiles.map(f => `• 📄 ${f.name}`).join('\n');
          botSay(`✅ **${session.uploadFiles.length} file(s) ready:**\n${names}`);
          clearAttachUI();
          setTimeout(() => startGeneration(), 400);
        }
      } else {
        botSay(`Use the 📎 button to attach files, then type **done**.\nOr type **skip** to proceed without documents.`);
      }
      break;
    }

    default: break;
  }
}

// ── Ask document source ───────────────────────────────────────────────────────
function askDocumentSource() {
  disableInput();
  const chips = [
    { label: '📁 From SharePoint', value: 'sharepoint' },
    { label: '📤 Upload Files',     value: 'upload' },
    { label: '⏭ Skip (no docs)',    value: 'skip' },
  ];
  botAsk(`📄 How would you like to provide **source documents**?`, chips);
  session.state = S.AWAIT_SOURCE;
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

    // Show file picker in chat
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
  disableInput();
  enableAttach(false);

  botSay(`⚡ Starting BRD generation for **${session.projectName}**...`);

  // Show progress card
  const progRowId = addProgressCard();

  try {
    let data;

    if (session.source === 'sharepoint') {
      // Get selected files from picker
      const selected = getSelectedSpFiles();
      if (!selected.length) {
        botSay(`⚠️ No files selected. Please select at least one file.`);
        session.state = S.AWAIT_SP_CONFIRM;
        return;
      }
      updateProgress(progRowId, 1, `Fetching ${selected.length} file(s) from SharePoint...`);
      await delay(800);
      updateProgress(progRowId, 2, 'Claude AI generating BRD...');
      await delay(600);
      updateProgress(progRowId, 3, 'Building Excel workbook...');

      const res = await fetch('/api/generate-brd-sharepoint', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          projectName: session.projectName,
          authorName:  session.authorName,
          files:       selected,
        }),
      });
      data = await res.json();
      if (!res.ok || !data.success) throw new Error(data.error || 'Generation failed');

    } else {
      // Upload or skip
      updateProgress(progRowId, 1, `Processing ${session.uploadFiles.length} document(s)...`);
      await delay(800);
      updateProgress(progRowId, 2, 'Claude AI generating BRD...');
      await delay(600);
      updateProgress(progRowId, 3, 'Building Excel workbook...');

      const fd = new FormData();
      fd.append('projectName', session.projectName);
      fd.append('authorName',  session.authorName);
      for (const f of session.uploadFiles) fd.append('files', f);

      const res = await fetch('/api/generate-brd', { method: 'POST', body: fd });
      data = await res.json();
      if (!res.ok || !data.success) throw new Error(data.error || 'Generation failed');
    }

    finishProgress(progRowId);
    await delay(400);
    showResult(data);
    session.state = S.DONE;

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

  const fileCount   = session.uploadFiles.length;
  const sourceLabel = session.source === 'sharepoint'  ? 'SharePoint' :
                      fileCount > 0                     ? `${fileCount} document${fileCount > 1 ? 's' : ''} uploaded` : 'Project name only';

  const card = el('div', 'result-card');
  card.innerHTML = `
    <div class="result-title">✅ BRD Generated Successfully!</div>
    <div class="result-stats">
      <div class="result-stat">
        <div class="result-stat-val">${s.total}</div>
        <div class="result-stat-lbl">Requirements</div>
      </div>
      <div class="result-stat high">
        <div class="result-stat-val">${s.high}</div>
        <div class="result-stat-lbl">High</div>
      </div>
      <div class="result-stat med">
        <div class="result-stat-val">${s.medium}</div>
        <div class="result-stat-lbl">Medium</div>
      </div>
      <div class="result-stat low">
        <div class="result-stat-val">${s.low}</div>
        <div class="result-stat-lbl">Low</div>
      </div>
    </div>
    <div class="result-meta">
      <strong>Project:</strong> ${esc(s.projectName)} &nbsp;·&nbsp;
      <strong>Date:</strong> ${esc(s.date)} &nbsp;·&nbsp;
      <strong>Version:</strong> ${esc(s.version)} &nbsp;·&nbsp;
      <strong>Source:</strong> ${esc(sourceLabel)}
    </div>
    <a class="result-dl" href="${esc(data.downloadUrl)}" download="${esc(data.fileName)}">
      📥 Download BRD Excel — ${esc(data.fileName)}
    </a>
  `;

  addBotRow(card);
  scrollBottom();

  // Follow-up message
  setTimeout(() => {
    botAsk(`🎉 Your BRD is ready! Click the button above to download.\n\nWould you like to generate another BRD?`, [
      { label: '✅ Yes, new BRD',  value: '__new__'  },
      { label: '👋 No, I\'m done', value: '__done__' },
    ]);
  }, 600);

  chatInput.addEventListener('keydown', function newHandler(e) {
    if (e.key !== 'Enter') return;
    const v = chatInput.value.trim().toLowerCase();
    chatInput.removeEventListener('keydown', newHandler);
    if (v === '__new__' || v === 'yes' || v === 'y') startConversation();
  });
}

// ── Chip click handler (quick replies) ───────────────────────────────────────
document.addEventListener('click', (e) => {
  const chip = e.target.closest('.qr-chip');
  if (!chip || chip.classList.contains('chosen')) return;

  // Mark all sibling chips as chosen
  chip.closest('.quick-replies')?.querySelectorAll('.qr-chip').forEach(c => {
    c.classList.add('chosen');
  });

  const value = chip.dataset.value;

  if (value === '__new__') { startConversation(); return; }
  if (value === '__done__') { botSay('Thanks! Have a great day. 👋'); disableInput(); return; }

  // Inject chip text as user message (except internal values)
  if (!value.startsWith('__')) {
    addUserMsg(chip.textContent.trim());
  } else {
    addUserMsg(chip.dataset.label || chip.textContent.trim());
  }

  processInput(value);
});

// ── File attachment ───────────────────────────────────────────────────────────
attachBtn.addEventListener('click', () => {
  if (session.state !== S.AWAIT_FILES) {
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
  const allowed = ['.pdf', '.docx', '.txt', '.doc'];
  let added = 0;
  for (const f of files) {
    const ext = '.' + f.name.split('.').pop().toLowerCase();
    if (!allowed.includes(ext)) continue;
    if (session.uploadFiles.some(x => x.name === f.name)) continue;
    session.uploadFiles.push(f);
    added++;
  }
  if (added) renderAttachPreview();
  if (session.state === S.AWAIT_FILES && added) {
    botSay(`📎 **${session.uploadFiles.length} file(s) attached.**\nAttach more, or type **done** to start generation.`);
  }
}

function renderAttachPreview() {
  attachPreviewIn.innerHTML = '';
  session.uploadFiles.forEach((f, i) => {
    const chip = el('div', 'attach-chip');
    chip.innerHTML = `${fileIcon(f.name)} ${esc(f.name)} <button class="attach-chip-remove" data-idx="${i}">✕</button>`;
    attachPreviewIn.appendChild(chip);
  });
  attachPreview.classList.toggle('hidden', session.uploadFiles.length === 0);
  attachBtn.classList.toggle('has-files', session.uploadFiles.length > 0);
}

document.addEventListener('click', (e) => {
  const btn = e.target.closest('.attach-chip-remove');
  if (!btn) return;
  session.uploadFiles.splice(Number(btn.dataset.idx), 1);
  renderAttachPreview();
});

// Clears only the UI strip (keeps session.uploadFiles intact for the API call)
function clearAttachUI() {
  attachPreviewIn.innerHTML = '';
  attachPreview.classList.add('hidden');
  attachBtn.classList.remove('has-files');
  enableAttach(false);
}

// Full reset — clears both UI and session file list (used on new conversation)
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

  // Select-all toggle
  document.getElementById('spAll')?.addEventListener('change', (e) => {
    wrap.querySelectorAll('.sp-picker-list input').forEach(cb => cb.checked = e.target.checked);
  });

  // Confirm button
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
function addProgressCard() {
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
      <div class="prog-label">Building Excel</div>
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
    if (i < step)  { row.classList.add('done');   row.querySelector('.prog-icon').innerHTML = '✓'; }
    if (i === step){ row.classList.add('active');  row.querySelector('.prog-icon').innerHTML = '<span class="prog-spinner"></span>'; }
    if (i > step)  { row.classList.add('wait');    row.querySelector('.prog-icon').innerHTML = '○'; }
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
  const bubble = el('div', 'msg-bubble');
  bubble.innerHTML = mdToHtml(text);
  addBotRow(bubble);
  scrollBottom();
}

function botAsk(text, chips = []) {
  const wrap  = el('div', '');
  const bubble= el('div', 'msg-bubble');
  bubble.innerHTML = mdToHtml(text);
  wrap.appendChild(bubble);

  if (chips.length) {
    const qr = el('div', 'quick-replies');
    chips.forEach(c => {
      const btn = el('button', 'qr-chip');
      btn.textContent  = c.label;
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

  const col = el('div', 'msg-col');
  col.appendChild(content);
  const time = el('div', 'msg-time');
  time.textContent = now();
  col.appendChild(time);

  row.appendChild(av);
  row.appendChild(col);
  chatMessages.appendChild(row);
}

function addUserMsg(text) {
  const row    = el('div', 'msg-row user');
  const av     = el('div', 'msg-avatar user-av');
  av.textContent = 'U';

  const col    = el('div', 'msg-col');
  const bubble = el('div', 'msg-bubble');
  bubble.textContent = text;
  const time   = el('div', 'msg-time');
  time.textContent = now();

  col.appendChild(bubble);
  col.appendChild(time);
  row.appendChild(col);
  row.appendChild(av);
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
  requestAnimationFrame(() => {
    chatMessages.scrollTop = chatMessages.scrollHeight;
  });
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
  return { pdf: '📕', docx: '📘', doc: '📘', txt: '📄' }[ext] || '📄';
}
function fmtSize(bytes) {
  if (!bytes) return '—';
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' KB';
  return (bytes / 1024 / 1024).toFixed(1) + ' MB';
}

/** Convert minimal markdown (**bold**, _italic_, \n, `code`) to safe HTML */
function mdToHtml(text) {
  return esc(text)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/_(.+?)_/g,       '<em>$1</em>')
    .replace(/`(.+?)`/g,       '<code>$1</code>')
    .replace(/\n/g,            '<br/>');
}
