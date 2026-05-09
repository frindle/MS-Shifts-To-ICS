'use strict';

const puppeteer = require('puppeteer-core');

const CHROME_PATH = process.env.CHROME_PATH || '/usr/bin/google-chrome-stable';
const EXTENSION_PATH = '/app/extension';
const PROFILE_PATH = '/data/chrome-profile';
const PUSHOVER_URL = 'https://api.pushover.net/1/messages.json';

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
    browser = await puppeteer.launch({
      headless: false,
      executablePath: CHROME_PATH,
      args: [
        `--load-extension=${EXTENSION_PATH}`,
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
    console.log('[sync] Auth captured — triggering export');

    // Reset storage state so we can detect a fresh completion.
    await swEval(
      session,
      `(async () => {
        await chrome.storage.local.set({ syncRunning: false, lastError: null, lastExport: null });
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
