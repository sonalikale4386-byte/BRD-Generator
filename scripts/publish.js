'use strict';
// Loads .env so GH_TOKEN is available, then runs electron-builder publish.
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env') });

process.env.ELECTRON_BUILDER_RCEDIT_PATH = './electron/rcedit';

const { execSync } = require('child_process');
const builder = path.join(__dirname, '..', 'node_modules', '.bin', 'electron-builder');
execSync(`"${builder}" --win --x64 --publish always`, { stdio: 'inherit', env: process.env });
