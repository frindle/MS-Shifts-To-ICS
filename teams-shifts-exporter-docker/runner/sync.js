'use strict';

const puppeteer = require('puppeteer-core');
const fs        = require('fs');

const CHROME_PATH    = process.env.CHROME_PATH || '/usr/bin/google-chrome-stable';
const EXTENSION_PATH = '/app/extension';
const PATCHED_EXTENSION_PATH = '/data/extension-run';
const PROFILE_PATH   = '/data/chrome-profile';
const SETTINGS_PATH  = '/data/settings.json';
const PUSHOVER_URL   = 'https://api.pushover.net/1/messages.json';

function loadSettings() {
  const defaults = {
    importToOutlook: false, importToiCloud: false, icloudEmail: '', icloudAppPassword: '', includeOpenShifts: true,
    importToNextcloud: false, nextcloudUrl: '', nextcloudUser: '', nextcloudAppPassword: '', nextcloudCalendarName: '',
  };
  let saved = {};
  try { saved = JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8')); } catch {}
  const s = { ...defaults, ...saved };
  // Env vars fill in anything not set via the dashboard (settings.json wins).
  s.nextcloudUrl          = s.nextcloudUrl          || process.env.NEXTCLOUD_URL || '';
  s.nextcloudUser         = s.nextcloudUser         || process.env.NEXTCLOUD_USER || '';
  s.nextcloudAppPassword  = s.nextcloudAppPassword  || process.env.NEXTCLOUD_APP_PASSWORD || '';
  s.nextcloudCalendarName = s.nextcloudCalendarName || process.env.NEXTCLOUD_CALENDAR_NAME || '';
  if (saved.importToNextcloud === undefined && s.nextcloudUrl && s.nextcloudUser && s.nextcloudAppPassword) {
    s.importToNextcloud = true;
  }
  return s;
}

// The Nextcloud server URL is user-entered, so the extension declares it as an
// optional host permission that a headless run cannot prompt for. Instead, load a
// copy of the extension whose manifest has the server origin added to host_permissions.
function prepareExtension(settings) {
  if (!settings.importToNextcloud || !settings.nextcloudUrl) return EXTENSION_PATH;
  let serverOrigin;
  try { serverOrigin = new URL(settings.nextcloudUrl).origin; }
  catch { console.error('[sync] Invalid Nextcloud URL — loading extension unpatched'); return EXTENSION_PATH; }
  fs.rmSync(PATCHED_EXTENSION_PATH, { recursive: true, force: true });
  fs.cpSync(EXTENSION_PATH, PATCHED_EXTENSION_PATH, { recursive: true });
  const manifestPath = `${PATCHED_EXTENSION_PATH}/manifest.json`;
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  manifest.host_permissions = [...(manifest.host_permissions || []), `${serverOrigin}/*`];
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
  return PATCHED_EXTENSION_PATH;
}

// How long to wait for Teams to fire a webRequest and let the SW capture auth headers.
const AUTH_TIMEOUT_MS = 90_000;
// How long to let the export run before giving up.
const SYNC_TIMEOUT_MS = 6 * 60_000;

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function sendPushover(message) {
  const token = process.env.PUSHOVER_TOKEN;
  const user  = process.env.PUSHOVER_USER;
  if (!token || !user) {
    console.error('[pushover] PUSHOVER_TOKEN / PUSHOVER_USER not set — skipping alert');
    return;
  }
  const body = new URLSearchParams({ token, user, message, title: 'Teams Shifts Sync' });
  try {
    const resp = await fetch(PUSHOVER_URL, { method: 'POST', body });
    if (resp.ok) console.log('[pushover] Alert sent');
    else console.error('[pushover] HTTP', resp.status, await resp.text());
  } catch (err) {
    console.error('[pushover] Request failed:', err.message);
  }
}

// Wait until a service worker target belonging to our extension appears.
async function findExtensionSwTarget(browser, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const sw = browser.targets().find(
      t => t.type() === 'service_worker' && t.url().startsWith('chrome-extension://')
    );
    if (sw) return sw;
    await sleep(500);
  }
  throw new Error('Extension service worker target not found');
}

// Evaluate an expression in the extension SW context via CDP.
async function swEval(session, expression, awaitPromise = false) {
  const result = await session.send('Runtime.evaluate', {
    expression,
    awaitPromise,
    returnByValue: true,
  });
  if (result.exceptionDetails) {
    throw new Error(result.exceptionDetails.text || JSON.stringify(result.exceptionDetails));
  }
  return result.result.value;
}

async function runSync() {
  console.log(`[sync] ${new Date().toISOString()} — starting`);
  let browser;

  try {
    const settings = loadSettings();
    const extensionPath = prepareExtension(settings);
    browser = await puppeteer.launch({
      headless: false,
      executablePath: CHROME_PATH,
      args: [
        `--load-extension=${extensionPath}`,
        `--user-data-dir=${PROFILE_PATH}`,
        '--no-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--disable-blink-features=AutomationControlled',
        '--window-size=1280,900',
      ],
    });

    const swTarget = await findExtensionSwTarget(browser);
    const session  = await swTarget.createCDPSession();
    console.log('[sync] Extension SW found:', swTarget.url());

    // Open Teams so the webRequest listener captures auth headers.
    const page = await browser.newPage();
    await page.goto('https://teams.cloud.microsoft/', {
      waitUntil: 'domcontentloaded',
      timeout: 30_000,
    });
    console.log('[sync] Teams tab opened — waiting for auth capture...');

    // Poll until background.js's capturedShiftsHeaders is non-null.
    const authDeadline = Date.now() + AUTH_TIMEOUT_MS;
    let authCaptured = false;
    while (Date.now() < authDeadline) {
      try {
        authCaptured = await swEval(session, 'capturedShiftsHeaders !== null');
      } catch {}
      if (authCaptured) break;
      await sleep(2_000);
    }

    if (!authCaptured) {
      throw new Error(
        'Teams auth not captured within 90s — connect to VNC (port 5900) and log in to Teams'
      );
    }
    console.log('[sync] Auth captured — applying settings and triggering export');

    // Push settings from /data/settings.json into chrome.storage.local so the
    // extension picks them up (iCloud creds, toggles, etc.) before running.
    await swEval(
      session,
      `(async () => {
        await chrome.storage.local.set({
          syncRunning:       false,
          lastError:         null,
          lastExport:        null,
          importToOutlook:   ${JSON.stringify(settings.importToOutlook)},
          importToiCloud:    ${JSON.stringify(settings.importToiCloud)},
          includeOpenShifts: ${JSON.stringify(settings.includeOpenShifts)},
          icloudEmail:       ${JSON.stringify(settings.icloudEmail)},
          icloudAppPassword: ${JSON.stringify(settings.icloudAppPassword)},
          importToNextcloud:     ${JSON.stringify(settings.importToNextcloud)},
          nextcloudUrl:          ${JSON.stringify(settings.nextcloudUrl)},
          nextcloudUser:         ${JSON.stringify(settings.nextcloudUser)},
          nextcloudAppPassword:  ${JSON.stringify(settings.nextcloudAppPassword)},
          nextcloudCalendarName: ${JSON.stringify(settings.nextcloudCalendarName)},
        });
      })()`,
      true
    );

    // Fire runExport — don't await so the SW keeps running independently.
    session.send('Runtime.evaluate', {
      expression: 'runExport({ auto: true })',
      awaitPromise: false,
    }).catch(() => {});

    // Poll storage until syncRunning goes false and lastExport is set.
    const syncDeadline = Date.now() + SYNC_TIMEOUT_MS;
    let state = null;
    while (Date.now() < syncDeadline) {
      await sleep(4_000);
      try {
        state = await swEval(
          session,
          `(async () => {
            return new Promise(resolve =>
              chrome.storage.local.get(
                ['syncRunning', 'lastError', 'lastCount', 'lastExport'],
                resolve
              )
            );
          })()`,
          true
        );
      } catch {}
      if (state && !state.syncRunning && state.lastExport) break;
      state = null;
    }

    if (!state) throw new Error('Sync timed out after 6 minutes');
    if (state.lastError) throw new Error(state.lastError);

    console.log(`[sync] Done — ${state.lastCount} events synced`);

  } catch (err) {
    console.error('[sync] FAILED:', err.message);
    await sendPushover(`Teams Shifts sync failed: ${err.message}`);
    process.exitCode = 1;
  } finally {
    if (browser) await browser.close().catch(() => {});
  }
}

runSync();
