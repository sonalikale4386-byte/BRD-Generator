'use strict';

const path  = require('path');
const fs    = require('fs');
const axios = require('axios');

class ClaudeService {
  constructor() {
    this.claudeKey  = process.env.CLAUDE_API_KEY || '';
    this.claudeModel = 'claude-sonnet-4-6';
  }

  get useMock() {
    return !this.claudeKey;
  }

  async generateBRD({ projectName, userName, extractedText,
    brdType = 'new', updateMethod = '', detailLevel = 'elaborated',
    fitGap = 'no', sourceRef = 'yes', moscow = 'yes', additionalInputs = '',
    previousBRD = null }) {

    const date      = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    const isConcise = detailLevel === 'concise';

    // Derive short prefix from project name (used in requirement IDs)
    const prefix = (() => {
      const name  = (projectName || 'BRD').trim();
      const words = name.split(/\s+/).filter(Boolean);
      if (words.length > 1) {
        const initials = words.map(w => (w.replace(/[^A-Za-z0-9]/g, '')[0] || '')).join('').toUpperCase();
        return initials.slice(0, 4) || 'BRD';
      }
      const clean = name.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
      return clean.length <= 4 ? (clean || 'BRD') : clean.slice(0, 3);
    })();

    // ── Mock fallback ────────────────────────────────────────────────────────
    if (this.useMock) {
      console.log('⚠️  No API key set — using mock BRD data');
      const mock = JSON.parse(
        fs.readFileSync(
          path.join(__dirname, '..', '..', 'resources', 'mock_brd_response.json'),
          'utf8'
        )
      );
      return {
        ...mock,
        project_name:  projectName || mock.project_name,
        prepared_by:   userName    || mock.prepared_by,
        document_date: date,
      };
    }

    const hasDocuments  = extractedText && extractedText.trim().length > 50;
    const isUpdate      = brdType === 'update' && previousBRD !== null;

    // ── System prompt ────────────────────────────────────────────────────────
    const systemPrompt = `You are a senior Business Analyst at Synoptek with 15+ years of experience \
creating enterprise Business Requirements Documents (BRDs) for Microsoft Dynamics 365 CRM implementations.

Your job is to read the provided source documents carefully and produce a project-specific BRD JSON. \
Every piece of content you generate must be grounded in what the source documents actually say — \
names, processes, systems, pain points, and goals.

CRITICAL RULES:
1. DO NOT generate generic or template content. Every requirement must trace to something real in the documents.
2. Requirements must describe ACTUAL business needs discovered in the source documents.
3. Stakeholder names/roles must come FROM the documents where possible.
4. The executive summary must describe THIS project specifically, not boilerplate text.
5. Vary MoSCoW priorities (Must Have/Should Have/Could Have/Won't Have) based on urgency and criticality signals in the documents — do not make everything "Must Have".
6. ${isConcise
  ? 'CONCISE MODE: Generate 8–12 requirements covering key functional areas only. Keep ALL text brief.'
  : 'ELABORATED MODE: Generate 12–20 requirements minimum; more if the documents are rich.'}
7. Return ONLY valid JSON — no markdown fences, no commentary outside the JSON.`;

    // ── User prompt ──────────────────────────────────────────────────────────
    const userPrompt = `${isUpdate ? `=== EXISTING BRD TO UPDATE ===
${JSON.stringify(previousBRD, null, 2)}
=== END EXISTING BRD ===

` : ''}${hasDocuments ? `=== SOURCE DOCUMENTS ===
${extractedText}
=== END SOURCE DOCUMENTS ===

` : ''}PROJECT DETAILS:
- Project Name: ${projectName}
- Prepared By: ${userName}
- Date: ${date}
- BRD Type: ${brdType === 'update' ? `Update Existing BRD (${updateMethod === 'latest' ? 'refine latest version' : 'upload previous for reference'})` : 'Create New BRD'}
- Detail Level: ${detailLevel === 'elaborated' ? 'Elaborated (comprehensive and detailed)' : 'Concise (focused and high-level)'}
- Include Fit-Gap Analysis: Yes (always required)
- Include Source References per Requirement: ${sourceRef === 'yes' ? 'Yes' : 'No'}
- Include MoSCoW Classification: ${moscow === 'yes' ? 'Yes' : 'No'}
${additionalInputs ? `- Additional Inputs: ${additionalInputs}` : ''}

${isUpdate
  ? `TASK: You are UPDATING the existing BRD provided above. Follow these rules strictly:
1. KEEP all existing requirements that remain valid — preserve their IDs (${prefix}-001, etc.), descriptions, and metadata.
2. MODIFY requirements where the source documents or additional inputs indicate changes are needed.
3. ADD new requirements for any new needs found in the source documents — continue numbering from the last existing ID.
4. REMOVE requirements that are explicitly superseded or no longer applicable.
5. UPDATE the executive summary, stakeholders, scope, and out-of-scope sections to reflect changes.
6. Increment document_version (e.g. v1.0 → v1.1, v1.1 → v1.2).
7. Set status of modified/new requirements to "In Review".
${hasDocuments ? 'Use the source documents above as the basis for what has changed.' : 'Use the additional inputs provided as the basis for changes.'}
Apply the detail level (${detailLevel}). Always populate fit_gap_analysis — update with any new or changed gaps.${sourceRef === 'no' ? ' Leave "source" blank.' : ''}Ensure MoSCoW priority is set in the "priority" field (Must Have/Should Have/Could Have/Won't Have).`
  : hasDocuments
    ? `TASK: Analyse the source documents above thoroughly. Extract the real requirements, \
stakeholders, processes, pain points, in-scope items, and out-of-scope items described. \
Build a complete BRD JSON that reflects what is actually described in those documents.
Apply the detail level (${detailLevel}) to control description depth. Always include a complete fit_gap_analysis section based on gaps found in the documents.${sourceRef === 'no' ? ' Leave "source" field blank in each requirement.' : ' Tag each requirement with its source document.'}Use the "priority" field for MoSCoW classification (Must Have/Should Have/Could Have/Won\'t Have).`
    : `TASK: No documents were uploaded. Generate a realistic BRD for a project named \
"${projectName}" that a Synoptek consultant would write for a typical CRM engagement. \
Use plausible but varied content — do not repeat the same values across fields.
Apply the detail level (${detailLevel}) to control description depth. Use MoSCoW classification in the "priority" field. Always include a complete fit_gap_analysis section with plausible gap findings for this type of CRM project.`
}

OUTPUT FORMAT — return a single JSON object matching this schema exactly:
{
  "project_name": string,
  "client_name": string,
  "platform": "Microsoft Dynamics 365 CRM",
  "phase": string,
  "prepared_by": string,
  "document_date": string,
  "document_version": "v1.0",
  "status": "Draft",
  "executive_summary": string (${isConcise ? '1–2 sentences, high-level only' : '2–4 sentences specific to this project'}),

  "key_stakeholders": [
    { "name": string, "role": string, "organization": string }
  ],

  "scope_summary": [
    { "module": string, "in_scope_count": number, "description": string }
  ],

  "requirements": [
    {
      "id": "${prefix}-001",
      "date": string (meeting/workshop date when this requirement was discussed — NOT the document generation date; use format DD-MMM-YYYY),
      "source": string,
      "scope": string,
      "sub_part": string,
      "description": string,
      "have_today": "Have and want to keep" | "Have and want to change/improve" | "Have and don't want to keep" | "Don't have, gap/pain point - Need improvement" | "Don't have and don't need/want",
      "priority": "Must Have" | "Should Have" | "Could Have" | "Won't Have",
      "status": "New" | "Draft" | "Approved" | "In Review" | "Deferred" | "Updated",
      "scope2": "D365 OOB no config" | "D365 OOB with config" | "Customization/development" | "Config & workflow automation" | "Config & dashboard" | "3rd party solution" | "Config & integration" | "Business procedure" | "Other",
      "context": string,
      "technical_comments": string,
      "remarks": string,
      "requester": string (displayed as "Lead Requested" — the senior stakeholder or lead who owns this requirement),
      "requested_by": string (displayed as "Meeting Participants" — all people present when this requirement was discussed)
    }
  ],

  "scope_checklist": [
    { "category": string, "module": string, "sub_category": string, "description": string }
  ],

  "brd_scope": [
    { "scope": string, "sub_part": string }
  ],

  "out_of_scope": [
    { "item": string }
  ],

  "scope_definitions": [
    { "scope": string, "description": string }
  ],

  "sign_off": [
    { "version": "v1.0", "name_and_role": "${userName}", "signature": "", "date": "" },
    { "version": "", "name_and_role": "Client Representative", "signature": "", "date": "" },
    { "version": "", "name_and_role": "Project Manager", "signature": "", "date": "" }
  ],

  "fit_gap_analysis": [
    {
      "requirement_id": string (matching ${prefix}-NNN),
      "requirement": string (brief requirement description),
      "current_state": string (what exists today — the "Have today" situation),
      "gap": string (what is missing or needs to change),
      "recommendation": string (D365 implementation approach to close this gap),
      "priority": "High" | "Medium" | "Low"
    }
  ],

  "lovs": [
    { "epic": string, "cycle": string }
  ],

  "covered_by_selections": ["D365 OOB no config","D365 OOB with config","Customization/development","Config & workflow automation","Config & dashboard","3rd party solution","Config & integration","Business procedure","Other"],
  "have_today_selections": ["Have and want to keep","Have and want to change/improve","Have and don't want to keep","Don't have, gap/pain point - Need improvement","Don't have and don't need/want"]
}

IMPORTANT REMINDERS:
- "requirements" must have ${isConcise ? '8–12' : 'AT LEAST 12'} entries, numbered ${prefix}-001, ${prefix}-002, etc.
- Each requirement description must be unique and specific.
- "executive_summary" must mention the actual client/project context.
- "date" per requirement: use the meeting/workshop/session date when this requirement was identified or discussed — NOT the document generation date. Extract dates from the source documents. Use DD-MMM-YYYY format (e.g. "15-Jan-2025").
- "scope" (column 4): the high-level CRM module or functional area (e.g. "Lead Management", "Opportunity Management").
- "sub_part" (column 5): the specific feature or sub-area within that scope module (e.g. "Lead Intake & Creation", "Pipeline Tracking"). Must be set for every requirement.
- "have_today" distribution: roughly 30% "Have and want to keep", 20% "Have and want to change/improve", 35% "Don't have, gap/pain point - Need improvement", 15% other values.
- "priority" distribution: roughly 40% "Must Have", 35% "Should Have", 15% "Could Have", 10% "Won't Have".
- "scope2" (displayed as "Implementation Approach") must be set for every requirement using the exact values listed above.
- "requester" (displayed as "Lead Requested"): the senior stakeholder / lead owner of the requirement.
- "requested_by" (displayed as "Meeting Participants"): all participants present when this was discussed.
- "fit_gap_analysis" must always be populated — include one entry per requirement that has a gap or change need. Minimum 5 entries.
- "scope_checklist" should have ${isConcise ? '4–6' : '8–15'} entries.
- "out_of_scope" should list items explicitly excluded or not mentioned.
${isConcise ? `
CONCISE MODE — STRICTLY ENFORCE these rules:
- "executive_summary": exactly 1–2 sentences, high-level summary only.
- "description" per requirement: ONE sentence maximum — state the business need plainly, no sub-bullets, no preamble.
- "context": leave EMPTY ("") for every requirement.
- "technical_comments": leave EMPTY ("") for every requirement.
- "remarks": leave EMPTY ("") for every requirement.
- "key_stakeholders": list top 3–5 only.
- "scope_definitions": 3–5 entries maximum, 1 sentence each.
- "scope_summary" descriptions: 1 sentence each.
- Total requirements: 8–12. Do NOT exceed 12.` : `
ELABORATED MODE — POPULATE EVERY FIELD FOR EVERY REQUIREMENT:
- "description": full paragraph (2–4 sentences) explaining the business need, current pain point, and expected outcome.
- "context": REQUIRED — never leave empty. Include specific examples, business rules, edge cases, volume estimates, and clarifications relevant to this requirement.
- "technical_comments": REQUIRED — never leave empty. Include the D365 implementation approach (OOB config, custom entity, workflow, plugin, integration, etc.), configuration notes, and any technical constraints.
- "remarks": REQUIRED — never leave empty. Include assumptions made, open questions, dependencies on other requirements, or stakeholder concerns.
- "requested_by": REQUIRED — name the specific stakeholder, team, or department who raised this requirement.
- "requester": REQUIRED — full name and/or role of the requesting party (e.g. "Sales Manager – John Smith").
- "sub_part": REQUIRED — the specific feature or sub-area within the scope module (e.g. "Lead Intake & Creation"). Never leave empty.
- "have_today": REQUIRED — assess current state accurately; distribute across all five options based on document evidence.
- "scope2": REQUIRED — classify the D365 implementation approach accurately for every requirement.
- "source": REQUIRED — reference the source document section or page where this requirement was identified.
- "key_stakeholders": list ALL stakeholders mentioned or implied in the source documents with name, role, and organization.
- "scope_checklist": 8–15 entries covering every functional area mentioned in the documents.
- "brd_scope": list all in-scope areas, each with scope and sub_part.
- "scope_definitions": 6–10 entries with full, detailed descriptions (2–3 sentences each).
- "out_of_scope": explicitly list items not addressed, with clear reasoning where possible.`}`;

    const raw = await this._callClaude(systemPrompt, userPrompt);

    const clean = raw
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/```\s*$/, '');

    let parsed;
    try {
      parsed = JSON.parse(clean);
    } catch (e) {
      // Try repairing truncated JSON (happens when response hits token limit)
      const repaired = this._repairJSON(clean);
      try {
        parsed = JSON.parse(repaired);
        console.warn('⚠️  BRD JSON was truncated — repaired and parsed partial response');
      } catch (e2) {
        const match = clean.match(/\{[\s\S]*\}/);
        if (match) {
          try {
            parsed = JSON.parse(this._repairJSON(match[0]));
          } catch {
            throw new Error(`Failed to parse AI response as JSON: ${e.message}\n\nRaw (first 500): ${clean.slice(0, 500)}`);
          }
        } else {
          throw new Error(`Failed to parse AI response as JSON: ${e.message}\n\nRaw (first 500): ${clean.slice(0, 500)}`);
        }
      }
    }

    parsed.project_name  = parsed.project_name  || projectName;
    parsed.prepared_by   = parsed.prepared_by   || userName;
    parsed.document_date = parsed.document_date || date;
    parsed.requirements  = parsed.requirements  || [];
    parsed._usage        = this.lastUsage || {};

    return parsed;
  }

  async chat(messages, brdContext = null, realTime = null) {
    // Fallback: build minimal real-time context if server didn't supply one
    const rt = realTime || (() => {
      const n = new Date();
      return {
        date:    n.toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' }),
        time:    n.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
        isoDate: n.toISOString().slice(0, 10),
        timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      };
    })();

    // ── Real-time context block ───────────────────────────────────────────────
    let rtBlock = `REAL-TIME CONTEXT (use this to answer date/time/app questions accurately):
- Today      : ${rt.date}
- Current time : ${rt.time}${rt.timezone ? ` (${rt.timezone})` : ''}
- ISO date   : ${rt.isoDate}
- App version: Synoptek CE BRD Generator v${rt.appVersion || 'latest'}`;

    if (rt.totalBRDsGenerated !== undefined) {
      rtBlock += `\n- BRDs generated (all time): ${rt.totalBRDsGenerated}`;
    }
    if (rt.recentProjects && rt.recentProjects.length) {
      rtBlock += `\n- Recent BRD projects:\n${rt.recentProjects.map(p => `    • ${p}`).join('\n')}`;
    }
    if (rt.usage) {
      const u = rt.usage;
      rtBlock += `\n- Cumulative API usage: ${u.totalBRDs} BRD(s), ${u.totalTokens.toLocaleString()} tokens total, $${u.totalCostUSD.toFixed(4)} spent`;
    }

    // ── System prompt ─────────────────────────────────────────────────────────
    let system = `You are a helpful AI assistant embedded in the Synoptek CE BRD Generator app.
You help users with business analysis, requirements engineering, Microsoft Dynamics 365 CRM, project management, and general questions.
Be concise and practical. Always use the real-time context below when answering questions about dates, times, or app data.

${rtBlock}`;

    // ── BRD context (if a BRD was just generated) ────────────────────────────
    if (brdContext) {
      const reqs      = brdContext.requirements || [];
      const mustHave  = reqs.filter(r => r.priority === 'Must Have').length;
      const shouldHave= reqs.filter(r => r.priority === 'Should Have').length;
      const couldHave = reqs.filter(r => r.priority === 'Could Have').length;
      system += `\n\nCURRENT BRD CONTEXT:
- Project    : ${brdContext.projectName}
- Date       : ${brdContext.date}   Version: ${brdContext.version}
- Requirements: ${brdContext.total} total (Must Have: ${mustHave}, Should Have: ${shouldHave}, Could Have: ${couldHave})
- Top requirements:
${reqs.slice(0, 10).map(r => `    [${r.id}] ${r.description} — ${r.priority}`).join('\n')}${reqs.length > 10 ? `\n    ...and ${reqs.length - 10} more` : ''}`;
    }

    try {
      const res = await axios.post(
        'https://api.anthropic.com/v1/messages',
        { model: this.claudeModel, max_tokens: 2048, system, messages },
        {
          headers: {
            'x-api-key':         this.claudeKey,
            'anthropic-version': '2023-06-01',
            'content-type':      'application/json',
          },
        }
      );
      return res.data.content[0].text.trim();
    } catch (err) {
      if (err.response) throw new Error(`Claude API ${err.response.status}: ${JSON.stringify(err.response.data)}`);
      throw new Error(`Claude API connection failed: ${err.message}`);
    }
  }

  _repairJSON(text) {
    // Walk the string tracking open braces/brackets; close whatever is still open
    const stack = [];
    let inStr = false, esc = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (esc)               { esc = false; continue; }
      if (ch === '\\' && inStr) { esc = true;  continue; }
      if (ch === '"')        { inStr = !inStr; continue; }
      if (inStr)             continue;
      if (ch === '{')        stack.push('}');
      else if (ch === '[')   stack.push(']');
      else if ((ch === '}' || ch === ']') && stack.length) stack.pop();
    }

    if (!stack.length) return text;   // already balanced

    let fixed = text;
    if (inStr) fixed += '"';                          // close open string
    fixed = fixed.replace(/,\s*$/, '');               // strip trailing comma
    fixed = fixed.replace(/:\s*$/, ': null');         // close dangling key
    fixed += stack.reverse().join('');                // close open brackets
    return fixed;
  }

  async _callClaude(systemPrompt, userPrompt) {
    console.log(`  🤖 Using Claude API (${this.claudeModel})`);
    try {
      const res = await axios.post(
        'https://api.anthropic.com/v1/messages',
        {
          model:      this.claudeModel,
          max_tokens: 32000,
          system:     systemPrompt,
          messages:   [{ role: 'user', content: userPrompt }],
        },
        {
          headers: {
            'x-api-key':         this.claudeKey,
            'anthropic-version': '2023-06-01',
            'content-type':      'application/json',
          },
        }
      );
      const usage = res.data.usage || {};
      this.lastUsage = usage;
      console.log(`  📊 Token usage — input: ${usage.input_tokens ?? '?'}, output: ${usage.output_tokens ?? '?'}, total: ${(usage.input_tokens || 0) + (usage.output_tokens || 0)}`);
      return res.data.content[0].text.trim();
    } catch (err) {
      if (err.response) {
        throw new Error(`Claude API ${err.response.status}: ${JSON.stringify(err.response.data)}`);
      }
      throw new Error(`Claude API connection failed: ${err.message}`);
    }
  }
}

module.exports = { ClaudeService };
