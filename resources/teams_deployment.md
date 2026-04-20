# Microsoft Teams Deployment Guide – BRD Generator Agent

## Prerequisites Checklist
- [ ] Copilot Studio agent created and tested
- [ ] Power Automate flow `BRD_Generator_Flow` created and tested
- [ ] BRD_Template.xlsx uploaded to SharePoint
- [ ] Claude API key stored in Environment Variables
- [ ] Microsoft Teams admin access (or Teams app sideload permission)

---

## Step 1 – Publish the Agent in Copilot Studio

1. Open your agent in **Copilot Studio** (https://copilotstudio.microsoft.com)
2. Click **Publish** in the top right
3. Click **Publish** again in the confirmation dialog
4. Wait for "Published successfully" confirmation

> Publishing makes the agent available to add to Teams. You must publish before deploying.

---

## Step 2 – Add the Teams Channel

1. In your agent → left sidebar → **Channels**
2. Click **Microsoft Teams**
3. Click **Turn on Teams**
4. Configure the following:

   **Bot name:** `BRD Generator`
   **Short description:** `Generate Business Requirement Documents from meeting transcripts and project details`
   **Long description:** `Upload your meeting summaries, transcripts, and project notes. The BRD Generator Agent uses AI to extract requirements and produces a professional Excel BRD document automatically.`
   **Developer name:** `[Your Organization Name]`
   **Developer website:** `[Your intranet or SharePoint URL]`

5. Click **Save**

---

## Step 3 – Configure Teams App Settings

1. Still in the **Microsoft Teams** channel → click **Edit details**
2. Upload icons:
   - **Color icon:** 192×192 px PNG (your org / project logo)
   - **Outline icon:** 32×32 px white PNG on transparent background
3. Set **App ID** — copy this, you'll need it later
4. Click **Save**

---

## Step 4 – Option A: Add to Teams (Personal / Direct Install)

This lets you test quickly by installing directly for yourself.

1. In Copilot Studio → **Channels** → **Microsoft Teams** → **Open agent in Teams**
2. Your browser will open a Teams deep link
3. Click **Add** in the Teams dialog
4. The bot will appear in your Teams chat list under **Apps**
5. Start a conversation with it and test end-to-end

---

## Step 5 – Option B: Deploy to Your Whole Organization (Admin)

For company-wide rollout via Teams Admin Center:

### 5a – Download the App Package
1. In Copilot Studio → **Channels** → **Microsoft Teams**
2. Click **Download app** — this downloads a `.zip` app package

### 5b – Upload to Teams Admin Center
1. Go to **Microsoft Teams Admin Center**: https://admin.teams.microsoft.com
2. Navigate to **Teams apps** → **Manage apps**
3. Click **Upload new app** → **Upload**
4. Select the `.zip` file downloaded in 5a
5. The app will appear in your org's app catalog

### 5c – Create App Setup Policy (optional but recommended)
1. In Teams Admin Center → **Teams apps** → **Setup policies**
2. Click **Add** → Name: `BRD Generator Policy`
3. Under **Installed apps** → **Add apps** → search `BRD Generator` → **Add**
4. Optionally pin it to the Teams sidebar: under **Pinned apps** → **Add apps**
5. Click **Save**

### 5d – Assign Policy to Users/Groups
1. Go to the policy you just created
2. Click **Manage users** → add specific users or security groups
3. Or go to **Users** → select users → **Assign policies** → select the policy

> Roll-out time: policy changes take up to 24 hours to propagate to users.

---

## Step 6 – Option C: Add to a Specific Teams Channel

To deploy the bot inside a Teams channel (so a team can use it):

1. Go to the Teams channel where you want the bot
2. Click **+** (Add a tab / connector) → search `BRD Generator`
3. Click **Add** → **Save**
4. Users in that channel can now @mention the bot: `@BRD Generator generate BRD`

---

## Step 7 – Test End-to-End in Teams

Run this test script in Teams:

| Step | Action | Expected Result |
|------|--------|-----------------|
| 1 | Open a chat with BRD Generator bot | Bot sends welcome greeting |
| 2 | Type "generate BRD" | Bot asks for project name |
| 3 | Enter project name | Bot asks for description |
| 4 | Enter description | Bot asks for client name |
| 5 | Enter client name | Bot asks to upload documents |
| 6 | Upload a .docx or .pdf meeting file | Bot confirms upload |
| 7 | Confirm generation | Power Automate flow triggers |
| 8 | Wait ~30–60 seconds | Bot returns a SharePoint link |
| 9 | Click the link | Excel BRD opens with populated content |

---

## Step 8 – Configure Proactive Notifications (Optional)

To have the bot notify users when the BRD is ready (async):

1. In Power Automate flow → add action at the end:
   ```
   Action: Microsoft Teams – Post message in a chat or channel
   Post in: Chat with bot
   Recipient: @{triggerBody()['text_UserEmail']}
   Message: "✅ Your BRD for **@{triggerBody()['text_ProjectName']}** is ready!
   📥 Download: @{outputs('Compose_FileURL')}"
   ```

2. Add `UserEmail` as an input variable to the flow
3. Pass the user's email from Copilot Studio agent variable

---

## Troubleshooting

| Issue | Cause | Fix |
|-------|-------|-----|
| Bot not appearing in Teams | Not published | Re-publish in Copilot Studio |
| Flow not triggering | Connection not authorized | Re-authenticate connectors in PA |
| File upload fails | File too large or wrong format | Limit to 10MB, PDF/DOCX/TXT only |
| Excel not populated | Table names wrong | Check named tables match exactly |
| Claude API error 401 | Invalid API key | Update Environment Variable |
| Claude API error 429 | Rate limit exceeded | Add retry logic in PA HTTP action |
| SharePoint copy fails | Path doesn't exist | Create `/BRD-Agent/Output/` folder |

---

## Recommended Flow Timeout Settings

In Power Automate flow settings:
- **Timeout:** 10 minutes (BRD generation can take 30–90 seconds)
- **Retry policy:** Fixed interval, 3 retries, 30-second interval (for Claude API calls)
- **Concurrency:** 10 (allows 10 simultaneous BRD generations)

---

## Security Considerations

- Store Claude API key in **Power Platform Environment Variables** (not hardcoded)
- Restrict SharePoint Output folder to authenticated users only
- Enable **DLP policies** in Power Platform to control data flow
- Review generated BRDs before sharing externally — AI output should be validated
