/**
 * share.js — Start server + open a public tunnel for team testing
 * Usage: npm run share
 *
 * Creates a public HTTPS URL via localtunnel.
 * Anyone with the URL can access the BRD Generator without any setup.
 */
process.env.SHARE = 'true';
require('./server');
