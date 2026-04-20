/**
 * graphService.js
 * Microsoft Graph API client for OneDrive "BRD Document" folder.
 * Uses a Bearer token obtained via Teams SSO (OAuthPrompt / GraphConnection).
 *
 * Requires Azure AD app permissions (delegated):
 *   Files.Read  — read user's OneDrive files
 */
const axios = require('axios');

const GRAPH   = 'https://graph.microsoft.com/v1.0';
const FOLDER  = 'BRD Document';                      // fixed OneDrive folder name
const WH_SECRET = 'BRDGenerator-WebhookSecret';       // clientState for subscription validation

class GraphService {
  constructor(accessToken) {
    this._http = axios.create({
      baseURL: GRAPH,
      headers: { Authorization: `Bearer ${accessToken}` },
    });
  }

  // ── Files ──────────────────────────────────────────────────────────────────

  /**
   * List all files (not folders) inside OneDrive "BRD Document" folder.
   * Returns [] if the folder does not exist yet.
   */
  async listBRDFiles() {
    try {
      const res = await this._http.get(
        `/me/drive/root:/${FOLDER}:/children`,
        { params: { $select: 'id,name,size,lastModifiedDateTime,@microsoft.graph.downloadUrl' } }
      );
      return (res.data.value || []).filter(item => !item.folder);
    } catch (err) {
      if (err.response?.status === 404) return [];   // folder not found — return empty
      throw err;
    }
  }

  /**
   * Download a file from its Graph download URL and return as Buffer.
   * The @microsoft.graph.downloadUrl is pre-authenticated, so no auth header needed.
   */
  async downloadFile(downloadUrl) {
    const res = await axios.get(downloadUrl, { responseType: 'arraybuffer', timeout: 30_000 });
    return Buffer.from(res.data);
  }

  // ── Webhook Subscriptions ──────────────────────────────────────────────────

  /**
   * Create a new subscription (or renew an existing one) for file creation
   * events in the "BRD Document" folder.
   *
   * Graph limits OneDrive subscription expiry to 4230 minutes (~3 days).
   * Call this on every BRD generation to keep the subscription alive.
   *
   * @param {string}  notificationUrl  Public HTTPS endpoint (e.g. ngrok URL + /api/graph-notifications)
   * @param {string?} existingSubId    Existing subscription ID to renew (or null to create new)
   * @returns {object} Graph subscription object
   */
  async upsertSubscription(notificationUrl, existingSubId = null) {
    const expiry = new Date(Date.now() + 4229 * 60 * 1000).toISOString();

    if (existingSubId) {
      try {
        const res = await this._http.patch(`/subscriptions/${existingSubId}`, {
          expirationDateTime: expiry,
        });
        console.log(`🔄 Renewed Graph subscription ${existingSubId}`);
        return res.data;
      } catch (err) {
        console.warn(`⚠️  Could not renew subscription ${existingSubId}: ${err.message} — creating new`);
      }
    }

    const res = await this._http.post('/subscriptions', {
      changeType:         'created',
      notificationUrl,
      resource:           `me/drive/root:/${FOLDER}:/children`,
      expirationDateTime: expiry,
      clientState:        WH_SECRET,
    });
    console.log(`✅ Created Graph subscription ${res.data.id}`);
    return res.data;
  }
}

module.exports = { GraphService, BRD_FOLDER: FOLDER, WH_SECRET };
