# End-to-End Testing Guide – BRD Generator Agent
## Testing Without Claude API Key (Mock Mode)

This guide walks through the complete setup and test from scratch.
Total estimated time: **90–120 minutes**

---

## Phase 1 – SharePoint Setup (15 minutes)

### 1.1 Create Folder Structure in SharePoint

1. Go to your SharePoint site (e.g., `https://yourtenant.sharepoint.com/sites/YourSite`)
2. Open **Documents** library
3. Create this folder structure:
   ```
   Documents/
   └── BRD-Agent/
       ├── Templates/
       └── Output/
   ```

### 1.2 Upload the BRD Template

1. Go to `Documents/BRD-Agent/Templates/`
2. Upload `output/BRD_Template.xlsx` from your local machine
3. Confirm the file appears in SharePoint

### 1.3 Add Named Tables to the Template

> This step is critical — Power Automate uses table names to find where to write data.

1. Open `BRD_Template.xlsx` from SharePoint (click to open in Excel Online)
2. For each sheet below, select the data range and create a named table:

   | Sheet | Select Range | Table Name |
   |-------|-------------|------------|
   | Business Requirements | A2:G22 | `BusinessRequirementsTable` |
   | Functional Requirements | A2:G22 | `FunctionalRequirementsTable` |
   | Non-Functional Req | A2:F17 | `NFRTable` |
   | Stakeholders | A2:H12 | `StakeholdersTable` |
   | Risk Register | A2:H17 | `RiskTable` |
   | Timeline & Milestones | A2:G10 | `MilestonesTable` |
   | Glossary | A2:D17 | `GlossaryTable` |

   **How to create a named table in Excel Online:**
   - Select the range (e.g., A2:G22)
   - Click **Insert** → **Table**
   - Check "My table has headers" → **OK**
   - Click the table → **Table Design** tab → rename in "Table Name" box

3. Save the file (Ctrl+S)

---

## Phase 2 – Power Automate Flow (Mock Mode) (30 minutes)

> In Mock Mode, we skip the Claude API call and use the pre-built JSON from
> `resources/mock_brd_response.json`. This tests the entire flow EXCEPT the AI step.

### 2.1 Create the Flow

1. Go to https://make.powerautomate.com
2. Click **Create** → **Instant cloud flow**
3. Name: `BRD_Generator_Flow_TEST`
4. Trigger: **Manually trigger a flow** (for now — we'll change to Copilot Studio trigger later)
5. Click **Create**

### 2.2 Add Manual Test Inputs

In the trigger, add these inputs:

| Input Name | Type | Sample Value |
|---|---|---|
| `ProjectName` | Text | `Digital Onboarding Platform` |
| `ClientName` | Text | `Acme Bank` |
| `ProjectDesc` | Text | `Automate digital customer onboarding for retail banking` |
| `UserName` | Text | `Test User` |

### 2.3 Add Initialize Variable – OutputFileName

```
Action: Initialize variable
Name: varOutputFileName
Type: String
Value: @{triggerBody()['text']}_BRD_@{formatDateTime(utcNow(),'yyyyMMdd')}.xlsx
```

> Replace `@{triggerBody()['text']}` with `@{triggerBody()['ProjectName']}`

### 2.4 Add the Mock BRD JSON (Skip Claude API for now)

Instead of calling Claude API, we'll hardcode the mock response:

```
Action: Initialize variable
Name: varBRDContent
Type: String
Value: (paste the ENTIRE contents of resources/mock_brd_response.json here)
```

### 2.5 Parse the BRD JSON

```
Action: Parse JSON
Content: @{variables('varBRDContent')}
Schema: (click "Generate from sample" and paste the mock_brd_response.json content)
```

### 2.6 Copy Template File in SharePoint

```
Action: SharePoint – Copy file
Current Site Address: [Your SharePoint URL]
File to Copy: /sites/YourSite/Shared Documents/BRD-Agent/Templates/BRD_Template.xlsx
Destination Site: [Your SharePoint URL]
Destination Folder: /sites/YourSite/Shared Documents/BRD-Agent/Output
New Name: @{variables('varOutputFileName')}
```

### 2.7 Populate Business Requirements Table

```
Action: Apply to each
Select Output: body('Parse_JSON')?['business_requirements']

  Inside loop:
  Action: Excel Online (Business) – Add a row into a table
  Location: SharePoint
  Document Library: Shared Documents
  File: /BRD-Agent/Output/@{variables('varOutputFileName')}
  Table: BusinessRequirementsTable
  Columns:
    BR ID:        @{items()?['id']}
    Category:     @{items()?['category']}
    Requirement Description: @{items()?['description']}
    Priority:     @{items()?['priority']}
    Source:       @{items()?['source']}
    Acceptance Criteria: @{items()?['acceptance_criteria']}
    Status:       @{items()?['status']}
```

### 2.8 Populate Stakeholders Table

```
Action: Apply to each
Select Output: body('Parse_JSON')?['stakeholders']

  Inside loop:
  Action: Excel Online (Business) – Add a row into a table
  Table: StakeholdersTable
  Columns:
    ID:           @{items()?['id']}
    Name:         @{items()?['name']}
    Title / Role: @{items()?['title']}
    Organization: @{items()?['organization']}
    Responsibility: @{items()?['responsibility']}
    Influence:    @{items()?['influence']}
    Interest:     @{items()?['interest']}
```

### 2.9 Populate Risk Table

```
Action: Apply to each
Select Output: body('Parse_JSON')?['risks']

  Inside loop:
  Action: Excel Online (Business) – Add a row into a table
  Table: RiskTable
  Columns:
    Risk ID:      @{items()?['id']}
    Category:     @{items()?['category']}
    Risk Description: @{items()?['description']}
    Probability:  @{items()?['probability']}
    Impact:       @{items()?['impact']}
    Risk Level:   @{items()?['risk_level']}
    Mitigation Strategy: @{items()?['mitigation']}
    Owner:        @{items()?['owner']}
```

### 2.10 Populate Milestones Table

```
Action: Apply to each
Select Output: body('Parse_JSON')?['milestones']

  Inside loop:
  Action: Excel Online (Business) – Add a row into a table
  Table: MilestonesTable
  Columns:
    MS ID:        @{items()?['id']}
    Milestone:    @{items()?['milestone']}
    Description / Deliverable: @{items()?['description']}
    Planned Date: @{items()?['planned_date']}
    Owner:        @{items()?['owner']}
    Status:       @{items()?['status']}
```

### 2.11 Populate Functional Requirements Table

```
Action: Apply to each
Select Output: body('Parse_JSON')?['functional_requirements']

  Inside loop:
  Action: Excel Online (Business) – Add a row into a table
  Table: FunctionalRequirementsTable
  Columns:
    FR ID:        @{items()?['id']}
    Module / Feature: @{items()?['module']}
    Requirement Description: @{items()?['description']}
    Priority:     @{items()?['priority']}
    BR Reference: @{items()?['br_reference']}
    Dependency:   @{items()?['dependency']}
    Status:       @{items()?['status']}
```

### 2.12 Populate NFR Table

```
Action: Apply to each
Select Output: body('Parse_JSON')?['non_functional_requirements']

  Inside loop:
  Action: Excel Online (Business) – Add a row into a table
  Table: NFRTable
  Columns:
    NFR ID:       @{items()?['id']}
    Category:     @{items()?['category']}
    Requirement Description: @{items()?['description']}
    Priority:     @{items()?['priority']}
    Acceptance Criteria: @{items()?['acceptance_criteria']}
    Status:       @{items()?['status']}
```

### 2.13 Get File Share Link

```
Action: SharePoint – Get file properties
Site Address: [Your SharePoint URL]
Library Name: Shared Documents
File Identifier: /BRD-Agent/Output/@{variables('varOutputFileName')}

Action: Compose
Inputs: @{body('Get_file_properties')?['{Link}']}
```

### 2.14 Save Flow and Run Test

1. Click **Save**
2. Click **Test** → **Manually** → **Run flow**
3. Enter the test input values
4. Watch the flow run step by step
5. Check the **Output** folder in SharePoint for the generated Excel file
6. Open the file and verify all sheets are populated

---

## Phase 3 – Copilot Studio Agent Setup (30 minutes)

### 3.1 Create the Agent

1. Go to https://copilotstudio.microsoft.com
2. Click **Create** → **New agent**
3. Fill in:
   - **Name:** `BRD Generator`
   - **Description:** `Generates Business Requirement Documents from project details and meeting documents`
4. In the **Instructions** box, paste:

```
You are a BRD Generator Agent for Microsoft Teams. Collect the following from the user:
1. Project Name
2. Project Description  
3. Client / Organization Name
4. Supporting documents (meeting summary, transcript) — ask them to upload files

After collecting all inputs, confirm and call the Power Automate flow to generate the BRD Excel document. Return the download link to the user.
```

5. Click **Create**

### 3.2 Update the Flow Trigger

1. Go back to your Power Automate flow
2. Change the trigger from **Manually trigger a flow** to:
   **When Copilot Studio calls a flow**
3. Add all input parameters (ProjectName, ClientName, ProjectDesc, UserName, File1URL, File2URL)
4. Save the flow

### 3.3 Connect Flow to Agent Topic

1. In Copilot Studio → **Topics** → create a new topic named **Generate BRD**
2. Trigger phrases: `generate BRD`, `create BRD`, `new BRD`, `start`
3. Add conversation nodes to collect:
   - Project Name → save to variable `varProjectName`
   - Project Description → save to `varProjectDesc`
   - Client Name → save to `varClientName`
   - File upload → save URL to `varFile1URL`
4. Add **Call an action** node → select `BRD_Generator_Flow_TEST`
5. Map variables to flow inputs
6. Add message node: `✅ BRD ready! Download: {varBRDFileURL}`
7. Save topic

### 3.4 Test in Copilot Studio

1. Click **Test your agent** (right panel)
2. Type: `generate BRD`
3. Walk through the conversation
4. Confirm the flow is triggered
5. Verify the output link is returned

---

## Phase 4 – Deploy to Teams (15 minutes)

1. In Copilot Studio → **Publish** → **Publish**
2. Go to **Channels** → **Microsoft Teams** → **Turn on Teams**
3. Click **Open agent in Teams**
4. In Teams dialog → **Add**
5. Start chatting with the bot in Teams

---

## Phase 5 – Upgrade to Live Claude API

Once you have your Claude API key from https://console.anthropic.com:

1. In Power Automate flow, after the file text extraction step, **delete** the mock variable initialization
2. Add these steps instead:
   ```
   Action: Compose – Build Claude Prompt
   (Use the template from resources/claude_prompt_template.md)

   Action: HTTP POST
   URI: https://api.anthropic.com/v1/messages
   Headers:
     x-api-key: [your key]
     anthropic-version: 2023-06-01
     content-type: application/json
   Body: @{outputs('Compose_ClaudePrompt')}
   ```
3. Parse the response: `body('HTTP_Claude')?['content'][0]['text']`
4. Continue with the Excel population steps as before

---

## Test Checklist

| Test Item | Expected Result | Status |
|-----------|----------------|--------|
| Power Automate flow runs without errors | All steps green ✅ | ☐ |
| Excel file created in SharePoint Output folder | File exists with correct name | ☐ |
| Business Requirements sheet populated | 7 rows of BR data | ☐ |
| Stakeholders sheet populated | 6 rows of stakeholder data | ☐ |
| Risk Register sheet populated | 5 rows of risk data | ☐ |
| Milestones sheet populated | 8 rows of milestone data | ☐ |
| Copilot Studio agent responds in test panel | Conversation flows correctly | ☐ |
| Agent deployed to Teams | Bot visible in Teams chat | ☐ |
| Teams end-to-end test | Bot responds + returns file link | ☐ |
| File opens correctly in Excel | All sheets readable, formatting intact | ☐ |
