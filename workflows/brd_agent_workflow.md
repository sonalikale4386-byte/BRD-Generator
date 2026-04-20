# BRD Generator Agent – End-to-End Workflow

## Overview
This agent generates a Business Requirements Document (BRD) in Excel format from:
- Project details provided by the user
- Uploaded supporting documents (meeting summaries, transcripts, notes)

Deployed in **Microsoft Teams** via **Copilot Studio** + **Power Automate** + **Claude API**.

---

## Architecture Diagram

```
User in Microsoft Teams
        │
        │  1. Types "generate BRD"
        │  2. Provides project details
        │  3. Uploads meeting docs (PDF/DOCX/TXT)
        ▼
┌─────────────────────────┐
│   Copilot Studio Agent  │  ← Conversation layer
│   "BRD Generator"       │    Collects inputs, validates
└──────────┬──────────────┘    Calls Power Automate
           │
           │  Trigger: When Copilot Studio calls a flow
           ▼
┌─────────────────────────┐
│   Power Automate Flow   │  ← Processing layer
│   "BRD_Generator_Flow"  │
│                         │
│  ① Download files       │
│  ② Extract text (AI     │
│     Builder)            │
│  ③ Build Claude prompt  │
│  ④ Call Claude API      │
│  ⑤ Parse BRD JSON       │
│  ⑥ Copy Excel template  │
│  ⑦ Populate Excel cells │
│  ⑧ Return file URL      │
└──────────┬──────────────┘
           │
           │  Calls
           ▼
┌─────────────────────────┐
│   Claude API            │  ← AI layer
│   claude-sonnet-4-6     │    Extracts requirements
│                         │    Returns structured JSON
└──────────┬──────────────┘
           │
           │  Reads/Writes
           ▼
┌─────────────────────────┐
│   SharePoint            │  ← Storage layer
│   BRD-Agent/Templates/  │    BRD_Template.xlsx (source)
│   BRD-Agent/Output/     │    ProjectName_BRD.xlsx (output)
└─────────────────────────┘
```

---

## Conversation Flow (in Teams)

```
User:  "generate BRD"
Bot:   "Hello! I'm the BRD Generator. What is the Project Name?"
User:  "Digital Onboarding Platform"
Bot:   "Got it. Please provide a brief Project Description."
User:  "Automate customer onboarding process for the bank..."
Bot:   "What is the Client / Organization name?"
User:  "Acme Bank"
Bot:   "Please upload your supporting documents (meeting summary,
        transcript, notes) — up to 3 files."
User:  [uploads MeetingSummary.pdf, Transcript.docx]
Bot:   "Confirming:
        • Project: Digital Onboarding Platform
        • Client: Acme Bank
        • Documents: 2 files uploaded
        Generate the BRD now?"
User:  "Yes"
Bot:   "⏳ Generating your BRD, please wait..."
       [Power Automate flow runs — ~30–60 seconds]
Bot:   "✅ BRD ready!
        📥 Download: https://sharepoint.../Digital_Onboarding_BRD_20260409.xlsx"
```

---

## BRD Excel Output Structure

The generated Excel file contains 13 sheets:

| # | Sheet Name               | Content                                      |
|---|--------------------------|----------------------------------------------|
| 1 | Cover Page               | Project title, client, date, version, status |
| 2 | Document Control         | Version history, approvals, distribution     |
| 3 | Executive Summary        | Overview, problem, solution, business value  |
| 4 | Business Objectives      | Goals, KPIs, strategic alignment             |
| 5 | Scope                    | In-scope, out-of-scope, boundaries           |
| 6 | Stakeholders             | Register, RACI matrix                        |
| 7 | Business Requirements    | BR table with IDs, priority, acceptance criteria |
| 8 | Functional Requirements  | FR table linked to BRs                       |
| 9 | Non-Functional Req       | Performance, security, compliance, etc.      |
|10 | Assumptions & Constraints| Assumptions, constraints, dependencies       |
|11 | Risk Register            | Risks, probability, impact, mitigation       |
|12 | Timeline & Milestones    | Milestones, phases, dates                    |
|13 | Glossary                 | Terms and definitions                        |

---

## File References

| File | Purpose |
|------|---------|
| `resources/generate_brd_template.js` | Generates the base BRD_Template.xlsx |
| `output/BRD_Template.xlsx` | Master template — upload to SharePoint |
| `resources/copilot_studio_setup.md` | Step-by-step Copilot Studio agent config |
| `resources/power_automate_flow.md` | Power Automate flow setup guide |
| `resources/claude_prompt_template.md` | Claude API prompt + expected JSON schema |
| `resources/teams_deployment.md` | Teams deployment and admin setup guide |

---

## Setup Sequence

Follow these steps in order:

1. **Generate Template** → run `node resources/generate_brd_template.js`
2. **Upload Template** → upload `output/BRD_Template.xlsx` to SharePoint `/BRD-Agent/Templates/`
3. **Add Named Tables** → open template in Excel, add named tables to each sheet (see `power_automate_flow.md`)
4. **Create PA Flow** → follow `power_automate_flow.md`
5. **Create CS Agent** → follow `copilot_studio_setup.md`
6. **Connect Flow to Agent** → link the PA flow to the CS agent topic
7. **Publish & Deploy** → follow `teams_deployment.md`
8. **Test** → run end-to-end test in Teams

---

## Supported Input Document Types

| Format | Processing Method        | Notes                      |
|--------|--------------------------|----------------------------|
| `.pdf` | AI Builder extraction    | Best for meeting summaries |
| `.docx`| AI Builder extraction    | Best for transcripts       |
| `.txt` | Direct text read         | Simple notes               |
| `.xlsx`| Not supported            | Convert to PDF first       |

---

## Limitations & Notes

- Maximum 3 files per BRD generation request (can be increased in PA flow)
- Maximum recommended file size: 10MB per file
- BRD generation time: ~30–90 seconds depending on document size
- AI-generated content should be reviewed and validated before client sign-off
- Claude API key must be kept secure — use Power Platform Environment Variables
