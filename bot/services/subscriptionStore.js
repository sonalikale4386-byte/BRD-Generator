/**
 * subscriptionStore.js
 * Persists Graph webhook subscriptions, conversation references,
 * and pending regeneration confirmations to resources/subscriptions.json.
 *
 * Keys used in the JSON file:
 *   sub_{subId}       — subscription entry
 *   ref_{userId}      — conversation reference for proactive messaging
 *   pending_{userId}  — pending yes/no regeneration confirmation
 */
const fs   = require('fs');
const path = require('path');

const STORE = path.join(__dirname, '..', '..', 'resources', 'subscriptions.json');

class SubscriptionStore {
  // ── Internal ───────────────────────────────────────────────────────────────
  _read() {
    try { return JSON.parse(fs.readFileSync(STORE, 'utf8')); }
    catch { return {}; }
  }
  _write(data) {
    fs.writeFileSync(STORE, JSON.stringify(data, null, 2));
  }

  // ── Subscription entries ───────────────────────────────────────────────────

  /** Persist a subscription entry keyed by Graph subscription ID */
  setSubscription(subId, entry) {
    const d = this._read();
    d[`sub_${subId}`] = entry;
    this._write(d);
  }

  /** Get a subscription entry by Graph subscription ID */
  getSubscription(subId) {
    return this._read()[`sub_${subId}`] || null;
  }

  /** Get the most recent subscription for a given userId (or null) */
  getSubscriptionByUser(userId) {
    const d = this._read();
    const entries = Object.entries(d)
      .filter(([k, v]) => k.startsWith('sub_') && v.userId === userId)
      .map(([k, v]) => ({ subId: k.replace('sub_', ''), ...v }));
    // Return the most recently subscribed
    entries.sort((a, b) => new Date(b.subscribedAt) - new Date(a.subscribedAt));
    return entries[0] || null;
  }

  /** Remove a subscription by ID */
  deleteSubscription(subId) {
    const d = this._read();
    delete d[`sub_${subId}`];
    this._write(d);
  }

  /** Return all subscription entries */
  getAllSubscriptions() {
    const d = this._read();
    return Object.entries(d)
      .filter(([k]) => k.startsWith('sub_'))
      .map(([k, v]) => ({ subId: k.replace('sub_', ''), ...v }));
  }

  // ── Conversation references (for proactive messaging) ─────────────────────

  /** Save the Teams conversation reference for a user */
  saveConversationRef(userId, convRef) {
    const d = this._read();
    d[`ref_${userId}`] = convRef;
    this._write(d);
  }

  /** Get the saved conversation reference for a user */
  getConversationRef(userId) {
    return this._read()[`ref_${userId}`] || null;
  }

  /** Get all saved conversation references (for proactive messaging) */
  getAllConversationRefs() {
    const d = this._read();
    return Object.entries(d)
      .filter(([k]) => k.startsWith('ref_'))
      .map(([k, v]) => ({ userId: k.replace('ref_', ''), convRef: v }));
  }

  /** Save last used project name for a user (for proactive regeneration) */
  saveLastProjectName(userId, projectName) {
    const d = this._read();
    d[`proj_${userId}`] = projectName;
    this._write(d);
  }

  /** Get last used project name for a user */
  getLastProjectName(userId) {
    return this._read()[`proj_${userId}`] || '';
  }

  // ── Pending regeneration confirmations ─────────────────────────────────────

  /**
   * Store a pending "Regenerate BRD?" confirmation for a user.
   * @param {string} userId
   * @param {{ subId: string, projectName: string, fileNames: string[] }} data
   */
  setPending(userId, data) {
    const d = this._read();
    d[`pending_${userId}`] = { ...data, askedAt: new Date().toISOString() };
    this._write(d);
  }

  /** Get pending confirmation for a user (or null) */
  getPending(userId) {
    return this._read()[`pending_${userId}`] || null;
  }

  /** Clear pending confirmation for a user */
  clearPending(userId) {
    const d = this._read();
    delete d[`pending_${userId}`];
    this._write(d);
  }
}

module.exports = { SubscriptionStore };
