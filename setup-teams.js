/**
 * setup-teams.js
 * Updates manifest validDomains with your bot URL and rebuilds the Teams ZIP.
 *
 * Usage:
 *   node setup-teams.js https://abc123.ngrok-free.app
 *   node setup-teams.js https://mybot.azurewebsites.net
 */

const fs   = require('fs');
const path = require('path');

const MANIFEST_DIR = path.join(__dirname, 'teams-manifest');
const MANIFEST     = path.join(MANIFEST_DIR, 'manifest.json');
const ZIP_PATH     = path.join(MANIFEST_DIR, 'BRDGeneratorAgent.zip');
const ENV_FILE     = path.join(__dirname, '.env');

// ── Get URL from args ─────────────────────────────────────────────────────────
let botUrl = process.argv[2];
if (!botUrl) {
  // Try reading from .env BOT_BASE_URL
  try {
    const env = fs.readFileSync(ENV_FILE, 'utf8');
    const match = env.match(/^BOT_BASE_URL=(.+)$/m);
    if (match) botUrl = match[1].trim();
  } catch (_) {}
}

if (!botUrl || botUrl.startsWith('http://localhost')) {
  console.error('');
  console.error('❌  No public URL provided.');
  console.error('');
  console.error('    Usage:  node setup-teams.js https://YOUR-NGROK-URL.ngrok-free.app');
  console.error('');
  console.error('    Steps to get a URL:');
  console.error('    1. Run:  ngrok http 3978');
  console.error('    2. Copy the https://... URL shown');
  console.error('    3. Re-run this script with that URL');
  process.exit(1);
}

// Strip trailing slash and protocol for domain extraction
const domain = botUrl.replace(/^https?:\/\//, '').replace(/\/$/, '');
const https  = `https://${domain}`;

console.log('');
console.log('🔧 Teams Manifest Setup');
console.log('─'.repeat(50));
console.log(`Bot URL : ${https}`);
console.log(`Domain  : ${domain}`);
console.log('');

// ── Update manifest ───────────────────────────────────────────────────────────
const manifest = JSON.parse(fs.readFileSync(MANIFEST, 'utf8'));
manifest.validDomains = ['token.botframework.com', domain];
fs.writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2));
console.log('✅ manifest.json updated');

// ── Update .env BOT_BASE_URL ──────────────────────────────────────────────────
try {
  let env = fs.readFileSync(ENV_FILE, 'utf8');
  env = env.replace(/^BOT_BASE_URL=.*$/m, `BOT_BASE_URL=${https}`);
  fs.writeFileSync(ENV_FILE, env);
  console.log('✅ .env BOT_BASE_URL updated');
} catch (_) {}

// ── Rebuild ZIP ───────────────────────────────────────────────────────────────
// Use JSZip-free approach: write ZIP manually using Node.js buffers
// We'll use the built-in child_process to call PowerShell
const { execSync } = require('child_process');

const ps = `
Add-Type -AssemblyName System.IO.Compression.FileSystem
$folder  = '${MANIFEST_DIR.replace(/\\/g, '\\\\')}'
$zipPath = '${ZIP_PATH.replace(/\\/g, '\\\\')}'
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
$zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')
foreach ($file in @('manifest.json', 'color.png', 'outline.png')) {
  $full = Join-Path $folder $file
  if (Test-Path $full) {
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $full, $file, 'Optimal') | Out-Null
  }
}
$zip.Dispose()
Write-Host "OK"
`.trim();

const result = execSync(`powershell -NoProfile -Command "${ps.replace(/"/g, '\\"').replace(/\n/g, '; ')}"`, { encoding: 'utf8' });
if (result.includes('OK')) {
  const size = fs.statSync(ZIP_PATH).size;
  console.log(`✅ BRDGeneratorAgent.zip rebuilt (${size} bytes)`);
} else {
  console.error('❌ ZIP rebuild failed:', result);
  process.exit(1);
}

// ── Print next steps ──────────────────────────────────────────────────────────
console.log('');
console.log('─'.repeat(50));
console.log('📋 Next Steps:');
console.log('');
console.log('  1. Azure Bot → Configuration → Messaging endpoint:');
console.log(`     ${https}/api/messages`);
console.log('');
console.log('  2. Start your bot:');
console.log('     node bot/index.js');
console.log('');
console.log('  3. Upload to Teams:');
console.log('     teams-manifest/BRDGeneratorAgent.zip');
console.log('');
console.log('  ⚠️  Other person\'s tenant: their Teams admin must enable');
console.log('     "Upload custom apps" in Teams Admin Center.');
console.log('─'.repeat(50));
console.log('');
