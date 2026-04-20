/**
 * claudeService.js
 * Calls Claude API to generate a Synoptek-format BRD JSON.
 * Falls back to mock data when CLAUDE_API_KEY is not set.
 */
const path = require('path');
const fs   = require('fs');

class ClaudeService {
  constructor() {
    this.apiKey = process.env.CLAUDE_API_KEY || '';
    this.model  = 'claude-sonnet-4-6';
  }

  async generateBRD({ projectName, userName, extractedText }) {
    const date = new Date().toLocaleDateString('en-GB', {
      day: '2-digit', month: 'short', year: 'numeric',
    });

    // ── Mock fallback (no API key) ───────────────────────────────────────────
    if (!this.apiKey) {
      console.log('⚠️  CLAUDE_API_KEY not set — using mock BRD data');
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

    const hasDocuments = extractedText && extractedText.trim().length > 50;

    // ── System prompt ────────────────────────────────────────────────────────
    const systemPrompt = `You are a senior Business Analyst at Synoptek with 15+ years of experience \
creating enterprise Business Requirements Documents (BRDs) for Microsoft Dynamics 365 CRM implementations.

Your job is to read the provided source documents carefully and produce a detailed, \
project-specific BRD JSON. Every piece of content you generate must be grounded in \
what the source documents actually say — names, processes, systems, pain points, and goals.

CRITICAL RULES:
1. DO NOT generate generic or template content. Every requirement must trace to something real in the documents.
2. Requirements must describe ACTUAL business needs discovered in the source documents.
3. Stakeholder names/roles must come FROM the documents where possible.
4. The executive summary must describe THIS project specifically, not boilerplate text.
5. Vary priorities (High/Medium/Low) based on urgency signals in the documents — do not make everything High.
6. Generate 12–20 requirements minimum; more if the documents are rich.
7. Return ONLY valid JSON — no markdown fences, no commentary outside the JSON.`;

    // ── User prompt ──────────────────────────────────────────────────────────
    const userPrompt = `${hasDocuments ? `=== SOURCE DOCUMENTS ===
${extractedText}
=== END SOURCE DOCUMENTS ===

` : ''}PROJECT DETAILS:
- Project Name: ${projectName}
- Prepared By: ${userName}
- Date: ${date}

${hasDocuments
  ? `TASK: Analyse the source documents above thoroughly. Extract the real requirements, \
stakeholders, processes, pain points, in-scope items, and out-of-scope items described. \
Build a complete BRD JSON that reflects what is actually described in those documents.`
  : `TASK: No documents were uploaded. Generate a realistic BRD for a project named \
"${projectName}" that a Synoptek consultant would write for a typical CRM engagement. \
Use plausible but varied content — do not repeat the same values across fields.`
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
  "executive_summary": string (2-4 sentences specific to this project),

  "key_stakeholders": [
    { "name": string, "role": string, "organization": string }
  ],

  "scope_summary": [
    { "module": string, "in_scope_count": number, "description": string }
  ],

  "requirements": [
    {
      "id": "LES-001",
      "date": string,
      "source": string (which document or meeting this came from),
      "requester": string (person or role who raised this),
      "description": string (specific, measurable functional requirement or user story),
      "have_today": "Yes" | "No" | "Partial",
      "priority": "High" | "Medium" | "Low",
      "status": "Draft" | "Approved" | "In Review" | "Deferred",
      "scope": string (CRM module/area),
      "context": string (business context or problem being solved),
      "technical_comments": string,
      "remarks": string,
      "requested_by": string
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

  "lovs": [
    { "epic": string, "cycle": string }
  ],

  "covered_by_selections": ["Yes","No","Partial","N/A","Future Phase"],
  "have_today_selections": ["Yes","No","Partial"]
}

IMPORTANT REMINDERS:
- "requirements" must have AT LEAST 12 entries, numbered LES-001, LES-002, etc.
- Each requirement description must be unique and specific — NO two requirements should read the same.
- "executive_summary" must mention the actual client/project context from the documents.
- "source" in each requirement should reference the specific document filename or meeting name.
- "have_today" distribution: roughly 40% No, 30% Partial, 30% Yes — vary based on the documents.
- "priority" distribution: roughly 40% High, 40% Medium, 20% Low — vary based on urgency signals.
- "scope_checklist" should have 8-15 entries covering relevant CRM modules found in the documents.
- "out_of_scope" should list items explicitly excluded or not mentioned in the source documents.`;

    // ── Call Claude API ──────────────────────────────────────────────────────
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key':         this.apiKey,
        'anthropic-version': '2023-06-01',
        'content-type':      'application/json',
      },
      body: JSON.stringify({
        model:      this.model,
        max_tokens: 16000,
        system:     systemPrompt,
        messages:   [{ role: 'user', content: userPrompt }],
      }),
    });

    if (!res.ok) {
      const err = await res.text();
      throw new Error(`Claude API ${res.status}: ${err}`);
    }

    const result = await res.json();
    const raw    = result.content[0].text.trim();

    // Strip accidental markdown fences if Claude adds them
    const clean = raw
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/```\s*$/,    '');

    let parsed;
    try {
      parsed = JSON.parse(clean);
    } catch (e) {
      // Attempt to extract JSON from response if there's surrounding text
      const match = clean.match(/\{[\s\S]*\}/);
      if (match) {
        parsed = JSON.parse(match[0]);
      } else {
        throw new Error(`Failed to parse Claude response as JSON: ${e.message}\n\nRaw (first 500): ${clean.slice(0, 500)}`);
      }
    }

    // Ensure required top-level fields are always present
    parsed.project_name  = parsed.project_name  || projectName;
    parsed.prepared_by   = parsed.prepared_by   || userName;
    parsed.document_date = parsed.document_date || date;
    parsed.requirements  = parsed.requirements  || [];

    return parsed;
  }
}

module.exports = { ClaudeService };
