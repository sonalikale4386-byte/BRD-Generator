# Claude API Prompt Template – BRD Generator

## API Endpoint
```
POST https://api.anthropic.com/v1/messages
```

## Headers
```json
{
  "x-api-key": "{{CLAUDE_API_KEY}}",
  "anthropic-version": "2023-06-01",
  "content-type": "application/json"
}
```

---

## Full Request Body Template

```json
{
  "model": "claude-sonnet-4-6",
  "max_tokens": 8000,
  "system": "You are a senior Business Analyst with 15+ years of experience creating Business Requirement Documents (BRDs) for enterprise IT and digital transformation projects. You extract requirements from meeting transcripts, summaries, and project descriptions with high accuracy. You always produce complete, professional, and actionable BRDs. You output only valid JSON with no extra commentary.",
  "messages": [
    {
      "role": "user",
      "content": "Generate a complete Business Requirements Document (BRD) in structured JSON format.\n\n## PROJECT DETAILS\n- **Project Name:** {{ProjectName}}\n- **Client / Organization:** {{ClientName}}\n- **Project Description:** {{ProjectDesc}}\n- **Requested By:** {{UserName}}\n- **Date:** {{CurrentDate}}\n\n## SOURCE DOCUMENTS\nThe following text has been extracted from uploaded meeting documents:\n\n{{ExtractedDocumentText}}\n\n---\n\n## OUTPUT INSTRUCTIONS\nAnalyse all source documents above and extract every requirement, objective, risk, stakeholder, and constraint mentioned. Infer items that are clearly implied but not explicitly stated. Use professional business language.\n\nReturn a single JSON object with EXACTLY these keys and structures:\n\n```json\n{\n  \"project_name\": \"string\",\n  \"client_name\": \"string\",\n  \"prepared_by\": \"string\",\n  \"document_date\": \"string\",\n  \"document_version\": \"v1.0\",\n  \"status\": \"Draft\",\n\n  \"executive_summary\": \"2-3 paragraph executive overview\",\n  \"project_overview\": \"Concise project overview\",\n  \"business_problem\": \"The business problem or opportunity being addressed\",\n  \"proposed_solution\": \"The proposed solution approach\",\n  \"business_value\": \"Expected business value and benefits\",\n  \"current_state\": \"Description of the current state / as-is situation\",\n  \"future_state\": \"Description of the desired future state\",\n  \"gap_analysis\": \"Key gaps between current and future state\",\n  \"success_criteria\": \"How success will be measured\",\n\n  \"business_objectives\": [\n    {\"id\": \"OBJ-001\", \"objective\": \"string\", \"description\": \"string\", \"priority\": \"High|Medium|Low\", \"owner\": \"string\"}\n  ],\n\n  \"kpis\": [\n    {\"id\": \"KPI-001\", \"name\": \"string\", \"measurement\": \"string\", \"baseline\": \"string\", \"target\": \"string\"}\n  ],\n\n  \"in_scope\": [\n    {\"id\": \"SC-001\", \"item\": \"string\", \"notes\": \"string\"}\n  ],\n\n  \"out_of_scope\": [\n    {\"id\": \"OOS-001\", \"item\": \"string\", \"reason\": \"string\"}\n  ],\n\n  \"stakeholders\": [\n    {\"id\": \"STK-001\", \"name\": \"string\", \"title\": \"string\", \"organization\": \"string\", \"responsibility\": \"string\", \"influence\": \"High|Medium|Low\", \"interest\": \"High|Medium|Low\"}\n  ],\n\n  \"business_requirements\": [\n    {\"id\": \"BR-001\", \"category\": \"string\", \"description\": \"string\", \"priority\": \"High|Medium|Low\", \"source\": \"string\", \"acceptance_criteria\": \"string\", \"status\": \"Draft\"}\n  ],\n\n  \"functional_requirements\": [\n    {\"id\": \"FR-001\", \"module\": \"string\", \"description\": \"string\", \"priority\": \"High|Medium|Low\", \"br_reference\": \"BR-001\", \"dependency\": \"string\", \"status\": \"Draft\"}\n  ],\n\n  \"non_functional_requirements\": [\n    {\"id\": \"NFR-001\", \"category\": \"Performance|Security|Scalability|Availability|Usability|Compliance|Integration|Maintainability\", \"description\": \"string\", \"priority\": \"High|Medium|Low\", \"acceptance_criteria\": \"string\", \"status\": \"Draft\"}\n  ],\n\n  \"assumptions\": [\n    {\"id\": \"ASM-001\", \"category\": \"string\", \"assumption\": \"string\", \"impact_if_wrong\": \"string\", \"owner\": \"string\"}\n  ],\n\n  \"constraints\": [\n    {\"id\": \"CON-001\", \"type\": \"Technical|Budget|Time|Resource|Regulatory\", \"constraint\": \"string\", \"impact\": \"string\", \"mitigation\": \"string\"}\n  ],\n\n  \"risks\": [\n    {\"id\": \"RSK-001\", \"category\": \"string\", \"description\": \"string\", \"probability\": \"H|M|L\", \"impact\": \"H|M|L\", \"risk_level\": \"High|Medium|Low\", \"mitigation\": \"string\", \"owner\": \"string\"}\n  ],\n\n  \"milestones\": [\n    {\"id\": \"MS-01\", \"milestone\": \"string\", \"description\": \"string\", \"planned_date\": \"string\", \"owner\": \"string\", \"status\": \"Not Started\"}\n  ],\n\n  \"glossary\": [\n    {\"term\": \"string\", \"definition\": \"string\", \"source\": \"string\"}\n  ]\n}\n```\n\nRules:\n- IDs must be sequential (BR-001, BR-002, etc.)\n- Minimum items: 5 business requirements, 5 functional requirements, 3 NFRs, 3 risks, 3 milestones\n- Priority must be exactly: High, Medium, or Low\n- Return ONLY the JSON object — no markdown fences, no explanation text"
    }
  ]
}
```

---

## Variable Substitutions for Power Automate

When building this prompt in the Power Automate **Compose** action, replace placeholders:

| Placeholder              | Power Automate Expression                                      |
|--------------------------|----------------------------------------------------------------|
| `{{ProjectName}}`        | `@{triggerBody()['text_ProjectName']}`                        |
| `{{ClientName}}`         | `@{triggerBody()['text_ClientName']}`                         |
| `{{ProjectDesc}}`        | `@{triggerBody()['text_ProjectDesc']}`                        |
| `{{UserName}}`           | `@{triggerBody()['text_UserName']}`                           |
| `{{CurrentDate}}`        | `@{formatDateTime(utcNow(), 'dd MMM yyyy')}`                  |
| `{{ExtractedDocumentText}}` | `@{variables('varExtractedText')}`                         |

---

## Expected Response Structure

Claude will return a JSON string like:

```json
{
  "project_name": "Digital Onboarding Platform",
  "client_name": "Acme Corp",
  "executive_summary": "Acme Corp is initiating a digital transformation...",
  "business_requirements": [
    {
      "id": "BR-001",
      "category": "User Management",
      "description": "The system shall allow new customers to register online without visiting a branch",
      "priority": "High",
      "acceptance_criteria": "User can complete registration in under 5 minutes",
      "status": "Draft"
    }
  ]
}
```

---

## Token Usage Estimate

| Document Size       | Input Tokens (est.) | Output Tokens (est.) | Cost (Sonnet) |
|---------------------|---------------------|----------------------|---------------|
| Small (1-2 pages)   | ~2,000              | ~3,000               | ~$0.04        |
| Medium (5-10 pages) | ~6,000              | ~5,000               | ~$0.09        |
| Large (15-20 pages) | ~12,000             | ~7,000               | ~$0.17        |

> Pricing based on claude-sonnet-4-6: $3/M input, $15/M output tokens (verify current pricing at anthropic.com)
