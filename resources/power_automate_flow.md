# Power Automate Flow – BRD Generator

## Flow Name: `BRD_Generator_Flow`
**Type:** Instant cloud flow (called from Copilot Studio)

---

## Flow Overview

```
[Copilot Studio calls flow]
        │
        ▼
[Receive inputs: project details + file URLs]
        │
        ▼
[Loop: Download each uploaded file → Extract text]
        │
        ▼
[Build Claude API prompt]
        │
        ▼
[HTTP POST → Claude API]
        │
        ▼
[Parse Claude JSON response]
        │
        ▼
[Copy BRD Template in SharePoint]
        │
        ▼
[Populate Excel via Excel Online connector]
        │
        ▼
[Generate shareable file link]
        │
        ▼
[Return file URL to Copilot Studio]
```

---

## Step-by-Step Flow Setup

### Step 0 – Create the Flow

1. Go to https://make.powerautomate.com
2. Click **Create** → **Instant cloud flow**
3. Name: `BRD_Generator_Flow`
4. Trigger: **When Copilot Studio calls a flow**
5. Click **Create**

---

### Step 1 – Define Inputs (from Copilot Studio)

In the trigger node, add these **Input parameters**:

| Parameter Name  | Type   | Description                         |
|-----------------|--------|-------------------------------------|
| `ProjectName`   | Text   | Project name from agent variable    |
| `ProjectDesc`   | Text   | Project description                 |
| `ClientName`    | Text   | Client / organization name          |
| `UserName`      | Text   | Name of the user requesting BRD     |
| `File1URL`      | Text   | URL of uploaded file 1 (or empty)   |
| `File2URL`      | Text   | URL of uploaded file 2 (or empty)   |
| `File3URL`      | Text   | URL of uploaded file 3 (or empty)   |

---

### Step 2 – Initialize Variables

Add **Initialize variable** actions:

| Variable Name       | Type   | Initial Value |
|---------------------|--------|---------------|
| `varExtractedText`  | String | (empty)       |
| `varBRDContent`     | String | (empty)       |
| `varFileURLs`       | Array  | `["@{triggerBody()['text_File1URL']}","@{triggerBody()['text_File2URL']}","@{triggerBody()['text_File3URL']}"]` |
| `varOutputFileName` | String | `@{triggerBody()['text_ProjectName']}_BRD_@{formatDateTime(utcNow(),'yyyyMMdd')}.xlsx` |

---

### Step 3 – Extract Text from Uploaded Files

Add **Apply to each** → loop over `varFileURLs`

Inside the loop:

#### 3a – Condition: Check if URL is not empty
```
Condition: Current item is not equal to (empty string)
```

**If yes:**

#### 3b – HTTP GET – Download File
```
Action: HTTP
Method: GET
URI: @{items('Apply_to_each')}
Headers:
  Authorization: Bearer @{body('Get_Teams_token')}
```

> **Note:** For files uploaded via Teams/Copilot Studio, the URL includes an auth token.
> Use the "Get file content using path" action from the SharePoint or OneDrive connector if files are stored there.

#### 3c – AI Builder – Extract Text (for PDF/DOCX)
```
Action: AI Builder → Extract information from documents
Document type: General document
Document: @{body('HTTP_Download_File')}
```

**OR** for plain text files:
```
Action: Compose
Inputs: @{base64ToString(body('HTTP_Download_File')['$content'])}
```

#### 3d – Append to variable
```
Action: Append to string variable
Name: varExtractedText
Value: 

=== Document @{iterationIndexes('Apply_to_each')} ===
@{body('Extract_information_from_documents')?['text']}

```

---

### Step 4 – Build Claude API Prompt

Add **Compose** action:

```
Name: Compose_ClaudePrompt
Inputs:
{
  "model": "claude-sonnet-4-6",
  "max_tokens": 8000,
  "messages": [
    {
      "role": "user",
      "content": "You are a Business Analyst expert. Generate a complete, structured Business Requirements Document (BRD) in JSON format based on the following inputs.\n\nPROJECT DETAILS:\n- Project Name: @{triggerBody()['text_ProjectName']}\n- Client: @{triggerBody()['text_ClientName']}\n- Description: @{triggerBody()['text_ProjectDesc']}\n- Requested By: @{triggerBody()['text_UserName']}\n- Date: @{formatDateTime(utcNow(), 'dd MMM yyyy')}\n\nSOURCE DOCUMENTS (extracted text):\n@{variables('varExtractedText')}\n\nGenerate a JSON object with exactly these keys:\n{\n  \"executive_summary\": \"...\",\n  \"project_overview\": \"...\",\n  \"business_problem\": \"...\",\n  \"proposed_solution\": \"...\",\n  \"business_value\": \"...\",\n  \"current_state\": \"...\",\n  \"future_state\": \"...\",\n  \"success_criteria\": \"...\",\n  \"kpis\": \"...\",\n  \"in_scope\": [\"item1\", \"item2\", ...],\n  \"out_of_scope\": [\"item1\", \"item2\", ...],\n  \"stakeholders\": [{\"name\":\"\",\"role\":\"\",\"responsibility\":\"\"},...],\n  \"business_requirements\": [{\"id\":\"BR-001\",\"category\":\"\",\"description\":\"\",\"priority\":\"High/Medium/Low\",\"acceptance_criteria\":\"\"},...],\n  \"functional_requirements\": [{\"id\":\"FR-001\",\"module\":\"\",\"description\":\"\",\"priority\":\"High/Medium/Low\",\"br_ref\":\"BR-001\"},...],\n  \"non_functional_requirements\": [{\"id\":\"NFR-001\",\"category\":\"\",\"description\":\"\",\"priority\":\"High/Medium/Low\"},...],\n  \"assumptions\": [{\"id\":\"A-001\",\"assumption\":\"\",\"impact\":\"\"},...],\n  \"constraints\": [{\"id\":\"C-001\",\"constraint\":\"\",\"impact\":\"\"},...],\n  \"risks\": [{\"id\":\"RSK-001\",\"description\":\"\",\"probability\":\"H/M/L\",\"impact\":\"H/M/L\",\"mitigation\":\"\"},...],\n  \"milestones\": [{\"id\":\"MS-01\",\"milestone\":\"\",\"description\":\"\",\"planned_date\":\"\"},...],\n  \"glossary\": [{\"term\":\"\",\"definition\":\"\"},...]\n}\n\nReturn ONLY the JSON object, no markdown, no explanation."
    }
  ]
}
```

---

### Step 5 – Call Claude API

```
Action: HTTP
Method: POST
URI: https://api.anthropic.com/v1/messages
Headers:
  x-api-key:         YOUR_CLAUDE_API_KEY
  anthropic-version: 2023-06-01
  content-type:      application/json
Body: @{outputs('Compose_ClaudePrompt')}
```

> **Security tip:** Store the Claude API key in **Azure Key Vault** or Power Platform **Environment Variables**, not hardcoded.

---

### Step 6 – Parse Claude Response

#### 6a – Parse JSON
```
Action: Parse JSON
Content: @{body('HTTP_Claude_API')}
Schema: (generate from sample response)
```

#### 6b – Set variable
```
Action: Set variable
Name: varBRDContent
Value: @{body('Parse_JSON')?['content'][0]['text']}
```

#### 6c – Parse BRD JSON
```
Action: Parse JSON
Content: @{variables('varBRDContent')}
Schema: (generate from sample BRD JSON)
```

---

### Step 7 – Copy BRD Template in SharePoint

```
Action: SharePoint → Copy file
Current Site Address: https://yourtenant.sharepoint.com/sites/YourSite
File to Copy: /BRD-Agent/Templates/BRD_Template.xlsx
Destination Site: https://yourtenant.sharepoint.com/sites/YourSite
Destination Folder: /BRD-Agent/Output
New Name: @{variables('varOutputFileName')}
```

---

### Step 8 – Populate Excel Sections

Use **Excel Online (Business)** connector actions to populate each sheet.

#### Cover Page
```
Action: Excel Online (Business) → Update a row
Location: SharePoint
Document Library: Documents
File: /BRD-Agent/Output/@{variables('varOutputFileName')}
Table: CoverTable   ← (pre-define named tables in the template)
Row ID: 1
Columns:
  ProjectName: @{triggerBody()['text_ProjectName']}
  ClientName:  @{triggerBody()['text_ClientName']}
  PreparedBy:  @{triggerBody()['text_UserName']}
  Date:        @{formatDateTime(utcNow(), 'dd MMM yyyy')}
  Status:      Draft
```

#### Executive Summary
```
Action: Excel Online (Business) → Update a row
Table: ExecutiveSummaryTable
Columns:
  ProjectOverview:  @{body('Parse_BRD_JSON')?['project_overview']}
  BusinessProblem:  @{body('Parse_BRD_JSON')?['business_problem']}
  ProposedSolution: @{body('Parse_BRD_JSON')?['proposed_solution']}
  BusinessValue:    @{body('Parse_BRD_JSON')?['business_value']}
  CurrentState:     @{body('Parse_BRD_JSON')?['current_state']}
  FutureState:      @{body('Parse_BRD_JSON')?['future_state']}
  SuccessCriteria:  @{body('Parse_BRD_JSON')?['success_criteria']}
```

#### Business Requirements (loop)
```
Action: Apply to each → loop over body('Parse_BRD_JSON')?['business_requirements']
  Action: Excel Online → Add a row into a table
  Table: BusinessRequirementsTable
  Columns:
    BRID:        @{items()?['id']}
    Category:    @{items()?['category']}
    Description: @{items()?['description']}
    Priority:    @{items()?['priority']}
    Acceptance:  @{items()?['acceptance_criteria']}
    Status:      Draft
```

> Repeat the same loop pattern for:
> - Functional Requirements → FunctionalRequirementsTable
> - Non-Functional Requirements → NFRTable
> - Stakeholders → StakeholdersTable
> - Risks → RiskTable
> - In Scope → ScopeInTable
> - Milestones → MilestonesTable
> - Glossary → GlossaryTable

---

### Step 9 – Get Shareable File Link

```
Action: SharePoint → Get file properties
Site Address: https://yourtenant.sharepoint.com/sites/YourSite
Library: Documents
File Identifier: /BRD-Agent/Output/@{variables('varOutputFileName')}

Action: Compose
Inputs: @{body('Get_file_properties')?['{Link}']?['AbsoluteUrl']}
```

---

### Step 10 – Return Output to Copilot Studio

```
Action: Return value(s) to Power Virtual Agents
Outputs:
  BRDFileURL (Text): @{outputs('Compose_FileURL')}
  StatusMessage (Text): "BRD generated successfully for @{triggerBody()['text_ProjectName']}"
```

---

## Named Excel Tables Required in Template

Before the flow can populate Excel, add named tables to `BRD_Template.xlsx`:

| Sheet                   | Table Name                    | Start Cell |
|-------------------------|-------------------------------|------------|
| Cover Page              | CoverTable                    | B8         |
| Executive Summary       | ExecutiveSummaryTable         | A2         |
| Business Objectives     | ObjectivesTable               | A3         |
| Scope                   | ScopeInTable / ScopeOutTable  | A3 / A15   |
| Stakeholders            | StakeholdersTable             | A3         |
| Business Requirements   | BusinessRequirementsTable     | A3         |
| Functional Requirements | FunctionalRequirementsTable   | A3         |
| Non-Functional Req      | NFRTable                      | A3         |
| Risk Register           | RiskTable                     | A3         |
| Timeline & Milestones   | MilestonesTable               | A3         |
| Glossary                | GlossaryTable                 | A3         |

> To add a table in Excel: Select the data range → Insert → Table → Name the table in the Table Design tab.

---

## Environment Variables (Power Platform)

Store secrets safely:

| Variable Name         | Value                          |
|-----------------------|--------------------------------|
| `ClaudeAPIKey`        | Your Anthropic API key         |
| `SharePointSiteURL`   | Your SharePoint site URL       |
| `BRDTemplatePath`     | /BRD-Agent/Templates/BRD_Template.xlsx |
| `BRDOutputPath`       | /BRD-Agent/Output/             |
