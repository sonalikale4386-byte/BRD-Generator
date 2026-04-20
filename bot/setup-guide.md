# Deployment Guide — BRD Generator Agent (M365 Agents SDK)

## Prerequisites
- Node.js installed ✅ (already confirmed)
- Microsoft 365 account with Teams access
- Azure account (free tier is sufficient — only for bot registration)
- ngrok installed (for local testing)

---

## Step 1 — Register the Bot in Azure (10 min)

> This gives you the App ID and Password the bot needs to talk to Teams.
> The bot itself runs on YOUR machine — Azure just routes the messages.

1. Go to **https://portal.azure.com**
2. Search for **"Azure Bot"** → click **Create**
3. Fill in:
   - **Bot handle:** `BRDGeneratorAgent`
   - **Subscription:** your subscription
   - **Resource group:** create new → `rg-brd-generator`
   - **Pricing tier:** F0 (Free)
   - **Type of App:** Multi Tenant
4. Click **Review + Create** → **Create**
5. Once deployed, go to the resource → **Configuration**
6. Note down the **Microsoft App ID**
7. Click **Manage Password** → **New client secret**
   - Description: `brd-bot-secret`
   - Expiry: 24 months
   - Copy the **Value** (shown only once!)

---

## Step 2 — Update .env File (2 min)

Open `.env` in the project root and fill in:

```
MICROSOFT_APP_ID=paste-your-app-id-here
MICROSOFT_APP_PASSWORD=paste-your-client-secret-here
CLAUDE_API_KEY=sk-ant-...   (leave blank to use mock mode)
```

---

## Step 3 — Start ngrok (2 min)

> ngrok creates a public HTTPS URL that forwards to your local bot server.

1. Download ngrok from **https://ngrok.com/download** (free account)
2. Install and authenticate:
   ```
   ngrok config add-authtoken YOUR_NGROK_TOKEN
   ```
3. Start the tunnel:
   ```
   ngrok http 3978
   ```
4. Copy the **Forwarding HTTPS URL**, e.g.:
   ```
   https://abc123.ngrok-free.app
   ```

---

## Step 4 — Configure the Bot Endpoint in Azure (2 min)

1. Go back to your Azure Bot resource → **Configuration**
2. Set **Messaging endpoint**:
   ```
   https://abc123.ngrok-free.app/api/messages
   ```
   _(use YOUR ngrok URL)_
3. Click **Apply**

---

## Step 5 — Update .env with ngrok URL (1 min)

```
BOT_BASE_URL=https://abc123.ngrok-free.app
```

This is used to build the Excel download link returned to the user.

---

## Step 6 — Update Teams Manifest (2 min)

Open `teams-manifest/manifest.json` and replace ALL placeholders:

| Placeholder | Replace with |
|---|---|
| `YOUR_MICROSOFT_APP_ID` | Your Azure App ID (appears twice) |
| `YOUR_NGROK_URL` | Your ngrok HTTPS URL |
| `YOUR_NGROK_DOMAIN` | Your ngrok domain only (e.g. `abc123.ngrok-free.app`) |

---

## Step 7 — Add Placeholder Icons (1 min)

The Teams manifest requires two icon files in `teams-manifest/`:
- `color.png` — 192×192 px PNG (your logo or any square image)
- `outline.png` — 32×32 px PNG (white icon on transparent background)

For quick testing, create simple placeholder PNGs using any image editor or online tool.

---

## Step 8 — Start the Bot (1 min)

```bash
cd "d:/BRD Generator"
node bot/index.js
```

You should see:
```
╔══════════════════════════════════════════════════╗
║   BRD Generator Agent — M365 Agents SDK          ║
╚══════════════════════════════════════════════════╝
  Server  : http://localhost:3978
  Endpoint: http://localhost:3978/api/messages
  AI Mode : 🟡 Mock (no API key)
```

---

## Step 9 — Sideload to Teams (5 min)

1. Zip the `teams-manifest/` folder contents:
   - `manifest.json`
   - `color.png`
   - `outline.png`
   - Name the zip: `BRDGeneratorAgent.zip`

2. Open **Microsoft Teams**
3. Click **Apps** (left sidebar) → **Manage your apps**
4. Click **Upload an app** → **Upload a custom app**
5. Select `BRDGeneratorAgent.zip`
6. Click **Add** in the Teams dialog

---

## Step 10 — Test in Teams

1. Find **BRD Generator** in your Teams chat list
2. Send: `generate BRD`
3. Answer the 3 questions
4. Attach a PDF or DOCX file using the paperclip icon
5. Type `generate`
6. Bot generates the BRD and returns a download link ✅

---

## Verify Everything is Working

| Check | Command / URL |
|---|---|
| Bot server running | `http://localhost:3978` → should show JSON status |
| ngrok tunnel active | ngrok dashboard at `http://localhost:4040` |
| Azure endpoint set | Azure Portal → Bot → Configuration |
| Teams app loaded | Teams → Apps → BRD Generator appears |

---

## Upgrade to Claude API

Once you have a key from **https://console.anthropic.com**:

1. Add to `.env`:
   ```
   CLAUDE_API_KEY=sk-ant-xxxxxxxxxxxxx
   ```
2. Restart the bot: `node bot/index.js`
3. Bot will now show: `🟢 Claude API`
4. Uploaded documents will be analysed by Claude to generate real BRD content

---

## For Production Deployment (instead of ngrok)

Deploy `bot/index.js` to any cloud host:

| Option | Command |
|---|---|
| **Railway** | `railway up` (free tier available) |
| **Render** | Connect GitHub repo → auto-deploy |
| **Azure App Service** | `az webapp up --name brd-generator-bot` |

Update `BOT_BASE_URL` and Azure messaging endpoint with the production URL.
