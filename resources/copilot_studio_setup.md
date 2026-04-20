# Copilot Studio Agent Setup – BRD Generator

## Prerequisites
- Microsoft 365 license with Copilot Studio access
- Power Automate (included in M365)
- SharePoint site (for storing Excel template + output files)
- Claude API key from https://console.anthropic.com

---

## Step 1 – Upload BRD Template to SharePoint

1. Go to your SharePoint site → **Documents** library
2. Create a folder: `BRD-Agent/Templates/`
3. Upload `output/BRD_Template.xlsx` to that folder
4. Note the full file path (you'll need it in the Power Automate flow)

---

## Step 2 – Create the Agent in Copilot Studio

### 2.1 Create New Agent
1. Go to https://copilotstudio.microsoft.com
2. Click **Create** → **New agent**
3. Name: `BRD Generator Agent`
4. Description: `Generates Business Requirement Documents from project details and uploaded meeting documents`
5. Instructions (paste into the Instructions field):

```
You are a BRD Generator Agent. Your role is to help users create Business Requirement Documents.

When a user starts a conversation:
1. Greet them and explain what you can do
2. Ask for the Project Name
3. Ask for a brief Project Description
4. Ask for the Client / Organization name
5. Ask them to upload supporting documents (meeting summary, transcript, notes) — up to 5 files (PDF, DOCX, or TXT)
6. Confirm all details and trigger the BRD generation flow
7. When the flow completes, share the download link for the generated Excel BRD

Always be professional and concise. If any required information is missing, ask again.
```

6. Click **Create**

---

## Step 3 – Configure Agent Variables

In the agent, go to **Variables** and create these:

| Variable Name        | Type   | Scope  |
|----------------------|--------|--------|
| `varProjectName`     | String | Global |
| `varProjectDesc`     | String | Global |
| `varClientName`      | String | Global |
| `varFile1URL`        | String | Global |
| `varFile2URL`        | String | Global |
| `varFile3URL`        | String | Global |
| `varBRDFileURL`      | String | Global |
| `varUserDisplayName` | String | Global |

---

## Step 4 – Create Topics

### Topic 1: Greeting / Start
**Trigger phrases:**
- "generate BRD"
- "create BRD"
- "new BRD"
- "start"
- "hello"
- "hi"

**Conversation nodes:**
```
Message: "Hello! I'm the BRD Generator Agent. I can create a Business Requirements Document from your project details and meeting documents.

Let's get started! I'll need:
✅ Project details (name, description, client)
📎 Supporting documents (meeting summary, transcript)"

→ [Go to Topic: Collect Project Details]
```

---

### Topic 2: Collect Project Details
**Nodes:**

```
Ask Question: "What is the Project Name?"
  → Save response to: varProjectName

Ask Question: "Please provide a brief Project Description (what problem does this project solve?)"
  → Save response to: varProjectDesc

Ask Question: "What is the Client / Organization name?"
  → Save response to: varClientName

Message: "Great! Now please upload your supporting documents.
You can upload:
• Meeting summary (PDF, DOCX, or TXT)
• Meeting transcript
• Any other relevant notes

Upload up to 3 files."

Ask Question (File Upload): "Please upload your documents"
  → Save attachment URL to: varFile1URL
  → (repeat for varFile2URL, varFile3URL if needed)

→ [Go to Topic: Confirm and Generate]
```

---

### Topic 3: Confirm and Generate
**Nodes:**

```
Message: "Here's a summary of what I have:
• Project: {varProjectName}
• Client: {varClientName}
• Description: {varProjectDesc}
• Documents: Uploaded ✅

Shall I generate the BRD now?"

Condition: User confirms → Yes
  → Call Power Automate flow: "BRD_Generator_Flow"
       Input: varProjectName, varProjectDesc, varClientName, varFile1URL, varFile2URL, varFile3URL
       Output: varBRDFileURL

Message: "✅ Your BRD has been generated!
📥 Download here: {varBRDFileURL}

The document has been saved to SharePoint. You can also find it at:
BRD-Agent/Output/{varProjectName}_BRD.xlsx"

Condition: User confirms → No
  → Message: "No problem! Let me know when you're ready or if you'd like to change anything."
```

---

### Topic 4: Error Handling
**Trigger:** On flow failure

```
Message: "I encountered an issue generating the BRD. Please check:
• Your uploaded files are PDF, DOCX, or TXT format
• Files are not password protected
• File size is under 10MB each

Please try again or contact your admin."
```

---

## Step 5 – Configure File Upload in Teams

By default, Copilot Studio captures file attachments in Teams as URLs. To enable:

1. In your agent settings → **Channels** → **Microsoft Teams**
2. Ensure **File attachments** is enabled
3. In Power Automate flow, use the attachment URL to download and process the file

---

## Step 6 – Test the Agent

1. In Copilot Studio, click **Test your agent** (right panel)
2. Type: "generate BRD"
3. Walk through the conversation flow
4. Upload a sample PDF or DOCX file
5. Confirm the flow triggers correctly

---

## Step 7 – Publish to Microsoft Teams

See `teams_deployment.md` for full deployment steps.
