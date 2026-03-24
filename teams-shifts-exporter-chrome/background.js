// background.js — service worker

const ALARM_NAME = 'daily-shifts-export';
const TEAMS_SHIFTS_URL = 'https://teams.cloud.microsoft/';
const OUTLOOK_CALENDAR_URL = 'https://outlook.office.com/calendar/view/month';

// ─── Alarm Setup ─────────────────────────────────────────────────────────────

chrome.runtime.onInstalled.addListener(() => {
  scheduleDailyAlarm();
});

chrome.runtime.onStartup.addListener(() => {
  scheduleDailyAlarm();
});

function scheduleDailyAlarm() {
  chrome.alarms.get(ALARM_NAME, (existing) => {
    if (!existing) {
      chrome.alarms.create(ALARM_NAME, {
        periodInMinutes: 24 * 60, // every 24 hours
        when: Date.now() + 60 * 1000, // first run in 1 minute (to allow Teams to load)
      });
    }
  });
}

// ─── Alarm Handler ───────────────────────────────────────────────────────────

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name === ALARM_NAME) {
    await runExport({ auto: true });
  }
});

// ─── Progress Helpers ─────────────────────────────────────────────────────────

function setProgress(step, percent) {
  chrome.storage.local.set({ syncRunning: true, syncStep: step, syncPercent: percent }).catch(() => {});
}

function clearProgress() {
  chrome.storage.local.set({ syncRunning: false, syncStep: '', syncPercent: 0 }).catch(() => {});
}

// ─── Export Logic ─────────────────────────────────────────────────────────────

async function runExport({ auto = false, skipICloud = false } = {}) {
  let scrapeWinId = null;
  try {
    // Always open a fresh Teams tab in a minimized window so the scraper
    // never touches the user's existing tabs or blocks their screen.
    setProgress('Opening Teams...', 2);
    const win = await chrome.windows.create({ url: TEAMS_SHIFTS_URL, state: 'minimized' });
    scrapeWinId = win.id;
    const tab = win.tabs[0];

    // Step 1: wait until Teams sidebar is interactive, then navigate to Shifts
    await waitForTeamsReady(tab.id);
    setProgress('Navigating to Shifts...', 8);
    await chrome.scripting.executeScript({ target: { tabId: tab.id, frameIds: [0] }, files: ['content.js'] });
    await chrome.tabs.sendMessage(tab.id, { action: 'NAVIGATE_TO_SHIFTS' }, { frameId: 0 });

    // Step 2: wait for the Shifts iframe to appear and get its frameId
    const shiftsFrame = await waitForShiftsFrame(tab.id);
    setProgress('Loading Shifts...', 14);

    // Step 3: inject content script into the iframe and wait for Shifts UI to render
    await chrome.scripting.executeScript({ target: { tabId: tab.id, frameIds: [shiftsFrame.frameId] }, files: ['content.js'] });
    await waitForShiftsReady(tab.id, shiftsFrame.frameId);
    setProgress('Starting scrape...', 18);

    // Auto-detect user name from Teams top frame ("Account manager for ...")
    let { userName } = await chrome.storage.local.get('userName');
    try {
      const nameResults = await chrome.scripting.executeScript({
        target: { tabId: tab.id, frameIds: [0] },
        func: () => {
          const btn = document.querySelector('button[aria-label^="Account manager for"]');
          if (btn) {
            const match = btn.getAttribute('aria-label').match(/Account manager for (.+)/);
            return match ? match[1] : null;
          }
          return null;
        },
      });
      const detectedName = nameResults?.[0]?.result;
      if (detectedName) {
        userName = detectedName;
        await chrome.storage.local.set({ userName: detectedName, userNameAutoDetected: true });
        console.info('[ShiftsExport] Auto-detected user name:', detectedName);
      }
    } catch (e) {
      console.info('[ShiftsExport] Name auto-detect failed, using saved name:', userName);
    }

    const response = await chrome.tabs.sendMessage(
      tab.id,
      { action: 'SCRAPE_AND_EXPORT', userName: userName || null },
      { frameId: shiftsFrame.frameId }
    );

    if (!response || !response.success) {
      throw new Error(response?.error || 'Scrape failed');
    }
    setProgress('Processing shifts...', 70);

    // Filter out open shifts if the user disabled them
    const { includeOpenShifts } = await chrome.storage.local.get('includeOpenShifts');
    let events = response.events || [];
    if (includeOpenShifts === false) {
      events = events.filter((e) => !e.isOpenShift);
    }

    // Merge freshly scraped events with stored history, then rebuild ICS
    const mergedEvents = await mergeWithHistory(events);
    const mergedICS = generateICS(mergedEvents);

    // Import to Outlook Web if the setting is enabled
    const { importToOutlook } = await chrome.storage.local.get('importToOutlook');
    let outlookResult = null;
    if (importToOutlook) {
      outlookResult = await importToOutlookWeb(mergedICS, auto);
    }

    // Sync to iCloud Calendar via CalDAV if enabled (skipped when CLEAR_AND_REIMPORT handles it)
    const { importToiCloud } = await chrome.storage.local.get('importToiCloud');
    let icloudResult = null;
    if (importToiCloud && !skipICloud) {
      icloudResult = await syncToiCloud(mergedEvents);
    }

    // Update last export time and store ICS for clear & re-import
    await chrome.storage.local.set({ lastExport: Date.now(), lastCount: mergedEvents.length, lastICS: mergedICS, lastEvents: mergedEvents });

    clearProgress();
    return { success: true, count: mergedEvents.length, outlookResult, icloudResult };
  } catch (err) {
    console.error('[ShiftsExport] Export error:', err);
    if (auto) {
      chrome.notifications.create('sync-failed', {
        type: 'basic',
        iconUrl: 'icon.png',
        title: 'Teams Shifts — Sync Failed',
        message: err.message || 'The daily sync encountered an error. Open the extension to retry.',
      });
    }
    return { success: false, error: err.message };
  } finally {
    clearProgress();
    // Always close the scraping window when done
    if (scrapeWinId) {
      try { await chrome.windows.remove(scrapeWinId); } catch {}
    }
  }
}

// ─── Outlook Web Import ───────────────────────────────────────────────────────

async function importToOutlookWeb(icsContent, autoMode) {
  try {
    let outlookTab = await getOrOpenOutlookTab(autoMode);
    if (!outlookTab) {
      return { success: false, error: 'Outlook Web not open' };
    }

    // Give the page a moment to be ready
    await sleep(2000);

    // Inject Outlook content script in case the tab was open before the extension loaded
    await chrome.scripting.executeScript({ target: { tabId: outlookTab.id }, files: ['outlook_content.js'] });
    await sleep(300);

    const response = await chrome.tabs.sendMessage(outlookTab.id, {
      action: 'IMPORT_ICS_TO_OUTLOOK',
      icsContent,
    });

    return response;
  } catch (err) {
    console.error('[ShiftsExport] Outlook import error:', err);
    return { success: false, error: err.message };
  }
}

async function getOrOpenOutlookTab(autoMode) {
  const patterns = [
    'https://outlook.office.com/*',
    'https://outlook.office365.com/*',
  ];

  for (const pattern of patterns) {
    const tabs = await chrome.tabs.query({ url: pattern });
    if (tabs.length > 0) return tabs[0];
  }

  // In auto mode, don't open a new Outlook tab unexpectedly
  if (autoMode) return null;

  // Open Outlook in a minimized window so it doesn't interrupt the user
  const win = await chrome.windows.create({ url: OUTLOOK_CALENDAR_URL, state: 'minimized' });
  await sleep(5000); // wait for Outlook Web to load
  return win.tabs[0];
}

// ─── History Merging ──────────────────────────────────────────────────────────

function deduplicateEvents(events) {
  const seen = new Set();
  return events.filter((e) => {
    const key = `${e.summary}|${e.startMs}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// Loads past events from storage, merges with freshly scraped events,
// deduplicates, saves back, and returns the full merged list.
async function mergeWithHistory(newEvents) {
  const now = Date.now();
  const { storedEvents = [] } = await chrome.storage.local.get('storedEvents');

  // Only keep stored events that are already in the past
  const pastStored = storedEvents.filter((e) => e.endMs < now);

  // Merge past history + all newly scraped events, then deduplicate
  const merged = deduplicateEvents([...pastStored, ...newEvents]);

  await chrome.storage.local.set({ storedEvents: merged });
  return merged;
}

// Generate an ICS string from a flat array of { summary, startMs, endMs }
function generateICS(events) {
  const pad = (n) => String(n).padStart(2, '0');

  function toICSDate(ms) {
    const d = new Date(ms);
    return (
      d.getFullYear() +
      pad(d.getMonth() + 1) +
      pad(d.getDate()) +
      'T' +
      pad(d.getHours()) +
      pad(d.getMinutes()) +
      '00'
    );
  }

  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Teams Shifts Export//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'X-WR-CALNAME:Teams Shifts',
  ];

  events.forEach((ev, i) => {
    const summaryText = ev.isOpenShift ? `OPEN: ${ev.summary}` : ev.summary;
    lines.push('BEGIN:VEVENT');
    lines.push(`UID:teams-shift-${ev.startMs}-${i}@shifts-export`);
    lines.push(`DTSTAMP:${toICSDate(Date.now())}`);
    lines.push(`DTSTART:${toICSDate(ev.startMs)}`);
    lines.push(`DTEND:${toICSDate(ev.endMs)}`);
    lines.push(`SUMMARY:${summaryText.replace(/,/g, '\\,').replace(/\n/g, '\\n')}`);
    if (ev.notes) {
      lines.push(`DESCRIPTION:${ev.notes.replace(/\\/g, '\\\\').replace(/,/g, '\\,').replace(/;/g, '\\;').replace(/\n/g, '\\n')}`);
    }
    if (ev.isOpenShift) {
      lines.push('CATEGORIES:Open Shift');
    }
    lines.push('END:VEVENT');
  });

  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}

function buildFilename() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `teams-shifts-${y}${m}${d}.ics`;
}

async function downloadICS(icsContent, filename) {
  const encoded = encodeURIComponent(icsContent);
  const dataUrl = `data:text/calendar;charset=utf-8,${encoded}`;

  await chrome.downloads.download({
    url: dataUrl,
    filename,
    saveAs: false,
    conflictAction: 'overwrite',
  });
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// ─── Load Polling Helpers ─────────────────────────────────────────────────────

// Poll until Teams sidebar navigation elements are available in the top frame
async function waitForTeamsReady(tabId, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const [r] = await chrome.scripting.executeScript({
        target: { tabId, frameIds: [0] },
        func: () => !!(
          document.querySelector('[aria-label*="Shifts" i][role="button"]') ||
          document.querySelector('[aria-label*="more" i][role="button"]') ||
          document.querySelector('[aria-label*="Apps" i][role="button"]') ||
          document.querySelector('button[aria-label^="Account manager for"]')
        ),
      });
      if (r?.result) return;
    } catch {}
    await sleep(500);
  }
  throw new Error('Timed out waiting for Teams to load');
}

// Poll until the Shifts iframe (flw.teams.cloud.microsoft) appears, return its frame record
async function waitForShiftsFrame(tabId, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const frames = await chrome.scripting.executeScript({
        target: { tabId, allFrames: true },
        func: () => window.location.href,
      });
      const frame = frames.find((f) => f.result && f.result.includes('flw.teams.cloud.microsoft'));
      if (frame) return frame;
    } catch {}
    await sleep(500);
  }
  throw new Error('Timed out waiting for Shifts iframe');
}

// Poll until the Shifts week-navigation UI is visible inside the iframe
async function waitForShiftsReady(tabId, frameId, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const [r] = await chrome.scripting.executeScript({
        target: { tabId, frameIds: [frameId] },
        func: () => !!(
          document.querySelector('button[aria-label="Go to next week"]') ||
          document.querySelector('button[aria-label*="Pick a date"]') ||
          document.querySelector('button[aria-label*="Your shifts" i]') ||
          document.querySelector('[data-tid="your-shifts-tab"]') ||
          document.querySelector('[data-tid="yourShifts-tab"]')
        ),
      });
      if (r?.result) return;
    } catch {}
    await sleep(500);
  }
  throw new Error('Timed out waiting for Shifts UI to load');
}

// ─── iCloud CalDAV Sync ───────────────────────────────────────────────────────

// Shared ICS date formatter used by both generateICS and the CalDAV client
function toICSDateStr(ms) {
  const pad = (n) => String(n).padStart(2, '0');
  const d = new Date(ms);
  return (
    d.getFullYear() +
    pad(d.getMonth() + 1) +
    pad(d.getDate()) +
    'T' +
    pad(d.getHours()) +
    pad(d.getMinutes()) +
    '00'
  );
}

// Build a single-event VCALENDAR block for CalDAV PUT
function buildSingleEventICS(event, uid) {
  const summaryText = event.isOpenShift ? `OPEN: ${event.summary}` : event.summary;
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Teams Shifts Export//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    `UID:${uid}@teams-shifts-export`,
    `DTSTAMP:${toICSDateStr(Date.now())}`,
    `DTSTART:${toICSDateStr(event.startMs)}`,
    `DTEND:${toICSDateStr(event.endMs)}`,
    `SUMMARY:${summaryText.replace(/,/g, '\\,').replace(/\n/g, '\\n')}`,
  ];
  if (event.notes) {
    lines.push(`DESCRIPTION:${event.notes.replace(/\\/g, '\\\\').replace(/,/g, '\\,').replace(/;/g, '\\;').replace(/\n/g, '\\n')}`);
  }
  if (event.isOpenShift) lines.push('CATEGORIES:Open Shift');
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return lines.join('\r\n');
}

class iCloudCalDAVClient {
  constructor(email, appPassword) {
    this.authHeader = 'Basic ' + btoa(`${email}:${appPassword}`);
    this.calendarHomeUrl = null;
  }

  // fetch wrapper that manually follows redirects to preserve HTTP method
  async request(url, method, body = null, extraHeaders = {}, timeoutMs = 20000) {
    const headers = {
      'Authorization': this.authHeader,
      'Content-Type': 'application/xml; charset=utf-8',
      ...extraHeaders,
    };

    const fetchWithTimeout = (fetchUrl) => {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      return fetch(fetchUrl, { method, headers, body, redirect: 'manual', signal: controller.signal })
        .finally(() => clearTimeout(timer));
    };

    let response = await fetchWithTimeout(url);

    let hops = 0;
    while (hops < 5) {
      const status = response.status;
      if (status === 301 || status === 302 || status === 307 || status === 308) {
        const location = response.headers.get('Location');
        if (!location) break;
        response = await fetchWithTimeout(location);
        hops++;
      } else {
        break;
      }
    }

    return response;
  }

  // Discover the calendar-home-set URL for this account
  async connect() {
    const principalBody = `<?xml version="1.0" encoding="utf-8"?><D:propfind xmlns:D="DAV:"><D:prop><D:current-user-principal/></D:prop></D:propfind>`;
    let res = await this.request('https://caldav.icloud.com/', 'PROPFIND', principalBody, { 'Depth': '0' });
    if (!res.ok) throw new Error(`iCloud principal discovery failed (HTTP ${res.status}). Check your Apple ID and app-specific password.`);

    const principalXml = await res.text();
    const principalPath = this._hrefInsideTag(principalXml, 'current-user-principal');
    if (!principalPath) throw new Error('iCloud did not return a principal URL. Ensure you are using an app-specific password.');

    const principalUrl = principalPath.startsWith('http')
      ? principalPath
      : `https://caldav.icloud.com${principalPath}`;

    const homeBody = `<?xml version="1.0" encoding="utf-8"?><D:propfind xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav"><D:prop><C:calendar-home-set/></D:prop></D:propfind>`;
    res = await this.request(principalUrl, 'PROPFIND', homeBody, { 'Depth': '0' });
    if (!res.ok) throw new Error(`iCloud calendar-home-set discovery failed (HTTP ${res.status})`);

    const homeXml = await res.text();
    const homePath = this._hrefInsideTag(homeXml, 'calendar-home-set');
    if (!homePath) throw new Error('iCloud did not return a calendar-home-set URL');

    this.calendarHomeUrl = homePath.startsWith('http')
      ? homePath
      : `https://caldav.icloud.com${homePath}`;
    if (!this.calendarHomeUrl.endsWith('/')) this.calendarHomeUrl += '/';
  }

  // Find the "Work Shifts" calendar URL, or create it if missing.
  // Matches by display name only — skipping resource-type checks that are
  // fragile across iCloud's varying namespace prefixes.
  async findOrCreateCalendar(displayName) {
    const listBody = `<?xml version="1.0" encoding="utf-8"?><D:propfind xmlns:D="DAV:"><D:prop><D:displayname/><D:resourcetype/></D:prop></D:propfind>`;
    const res = await this.request(this.calendarHomeUrl, 'PROPFIND', listBody, { 'Depth': '1' });
    if (!res.ok) throw new Error(`iCloud calendar listing failed (HTTP ${res.status})`);

    const xml = await res.text();
    console.info('[iCloud] Calendar list XML:', xml);

    for (const block of this._responseBlocks(xml)) {
      const name = this._tagText(block, 'displayname');
      if (name && name.trim() === displayName) {
        const href = this._tagText(block, 'href');
        if (href) return href.startsWith('http') ? href : `https://caldav.icloud.com${href}`;
      }
    }

    // Calendar not found — create it
    const uid = crypto.randomUUID();
    const newCalUrl = `${this.calendarHomeUrl}${uid}/`;
    const mkBody = `<?xml version="1.0" encoding="utf-8"?><C:mkcalendar xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav"><D:set><D:prop><D:displayname>${displayName}</D:displayname></D:prop></D:set></C:mkcalendar>`;
    const mkRes = await this.request(newCalUrl, 'MKCALENDAR', mkBody);
    if (mkRes.status !== 201) throw new Error(`iCloud MKCALENDAR failed (HTTP ${mkRes.status})`);
    console.info('[iCloud] Created calendar:', displayName, newCalUrl);
    return newCalUrl;
  }

  // REPORT the calendar to get a map of { uid -> { url, etag } } for our events
  async getOurEvents(calendarUrl) {
    const reportBody = `<?xml version="1.0" encoding="utf-8"?><C:calendar-query xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav"><D:prop><D:getetag/></D:prop><C:filter><C:comp-filter name="VCALENDAR"/></C:filter></C:calendar-query>`;
    const res = await this.request(calendarUrl, 'REPORT', reportBody, { 'Depth': '1' });
    if (res.status !== 207 && !res.ok) throw new Error(`iCloud REPORT failed (HTTP ${res.status})`);

    const xml = await res.text();
    const result = new Map(); // uid -> { url, etag }
    for (const block of this._responseBlocks(xml)) {
      const href = this._tagText(block, 'href');
      if (!href || !href.endsWith('.ics')) continue;
      const filename = href.split('/').pop().replace('.ics', '');
      if (filename.startsWith('teams-shift-')) {
        const rawEtag = this._tagText(block, 'getetag') || '';
        result.set(filename, {
          url: href.startsWith('http') ? href : `https://caldav.icloud.com${href}`,
          etag: rawEtag.replace(/^"|"$/g, ''), // strip surrounding quotes
        });
      }
    }
    return result;
  }

  // PUT a single event. Tries create-only first (If-None-Match: *); if the event
  // already exists (412), retries as an unconditional update. This avoids needing
  // to track ETags while still satisfying iCloud's precondition requirements.
  async putEvent(calendarUrl, uid, icsContent) {
    const base = calendarUrl.endsWith('/') ? calendarUrl : calendarUrl + '/';
    const url = `${base}${uid}.ics`;
    const headers = { 'Content-Type': 'text/calendar; charset=utf-8' };

    let res = await this.request(url, 'PUT', icsContent, { ...headers, 'If-None-Match': '*' });
    if (res.status === 412) {
      // Event already exists — update it unconditionally
      res = await this.request(url, 'PUT', icsContent, headers);
    }
    if (res.status < 200 || res.status >= 300) {
      throw new Error(`iCloud PUT failed (HTTP ${res.status}) for ${uid}`);
    }
  }

  async deleteEvent(eventUrl) {
    const res = await this.request(eventUrl, 'DELETE');
    if (res.status !== 204 && res.status !== 200) {
      console.warn('[iCloud] DELETE returned', res.status, 'for', eventUrl);
    }
  }

  // Delete every teams-shift-* event from the calendar regardless of date
  async clearAllOurEvents(calendarUrl) {
    const existingOurEvents = await this.getOurEvents(calendarUrl);
    for (const [uid, { url }] of existingOurEvents) {
      await this.deleteEvent(url);
      console.info('[iCloud] Cleared event:', uid);
    }
  }

  // Sync events to iCloud:
  // - Scheduled shifts: always upsert (add new, update existing)
  // - Open shifts: only add if never synced before; if user deleted one from
  //   iCloud, leave it deleted (don't re-add on subsequent syncs)
  // - Stale future events no longer in Teams: delete from iCloud
  async syncEvents(calendarUrl, events, onProgress = null) {
    const currentUids = new Set();
    const now = Date.now();
    let uploaded = 0;

    // Load the set of open shift UIDs we've already pushed to iCloud
    const { syncedOpenShiftUids: storedUids = [] } = await chrome.storage.local.get('syncedOpenShiftUids');
    const syncedOpenShiftUids = new Set(storedUids);

    // Pre-calculate how many events will actually be uploaded so the fraction
    // fills 0→1 over real uploads, not skipped open shifts.
    const total = events.filter((e) => {
      if (!e.isOpenShift) return true;
      const uid = `teams-shift-${e.startMs}-${e.summary.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()}`;
      return !syncedOpenShiftUids.has(uid);
    }).length || 1;

    for (const event of events) {
      const uid = `teams-shift-${event.startMs}-${event.summary.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()}`;
      currentUids.add(uid);

      if (event.isOpenShift) {
        // Only add open shifts we haven't pushed before
        if (!syncedOpenShiftUids.has(uid)) {
          if (onProgress) onProgress(`Uploading shift ${++uploaded} of ${total}…`, uploaded / total);
          await this.putEvent(calendarUrl, uid, buildSingleEventICS(event, uid));
          syncedOpenShiftUids.add(uid);
        }
      } else {
        // Scheduled shifts: always upsert
        if (onProgress) onProgress(`Uploading shift ${++uploaded} of ${total}…`, uploaded / total);
        await this.putEvent(calendarUrl, uid, buildSingleEventICS(event, uid));
      }
    }

    // Delete stale future events that are no longer in Teams
    if (onProgress) onProgress('Removing old shifts…', 0.97);
    const existingOurEvents = await this.getOurEvents(calendarUrl);
    for (const [uid, { url }] of existingOurEvents) {
      if (currentUids.has(uid)) continue;
      const startMs = parseInt(uid.replace('teams-shift-', ''), 10);
      if (!isNaN(startMs) && startMs > now) {
        await this.deleteEvent(url);
        syncedOpenShiftUids.delete(uid);
        console.info('[iCloud] Deleted stale event:', uid);
      }
    }

    // Persist updated open shift tracking set
    await chrome.storage.local.set({ syncedOpenShiftUids: [...syncedOpenShiftUids] });
  }

  // ── Regex-based XML helpers (DOMParser unavailable in MV3 service workers) ──

  // Namespace prefix is optional — iCloud may omit it in some responses
  _responseBlocks(xml) {
    const blocks = [];
    const re = /<(?:[a-zA-Z]+:)?response[\s>][\s\S]*?<\/(?:[a-zA-Z]+:)?response>/gi;
    let m;
    while ((m = re.exec(xml)) !== null) blocks.push(m[0]);
    return blocks;
  }

  _tagText(xml, localName) {
    const re = new RegExp(`<[a-zA-Z]*:?${localName}[^>]*>([^<]*)<`, 'i');
    const m = re.exec(xml);
    return m ? m[1].trim() : null;
  }

  _hasTag(xml, localName) {
    return new RegExp(`<[a-zA-Z]*:?${localName}[\\s/>]`, 'i').test(xml);
  }

  _hrefInsideTag(xml, parentLocalName) {
    const re = new RegExp(`<[a-zA-Z]*:?${parentLocalName}[^>]*>([\\s\\S]*?)<\\/[a-zA-Z]*:?${parentLocalName}>`, 'i');
    const m = re.exec(xml);
    if (!m) return null;
    return this._tagText(m[1], 'href');
  }
}

async function clearAndResyncToiCloud() {
  // No overall Promise.race timeout here — per-request 20s timeouts in request()
  // are sufficient. Using Promise.race causes a goroutine leak where syncEvents
  // continues running after timeout and saves syncedOpenShiftUids to storage,
  // marking open shifts as "already synced" even though they never reached iCloud.
  try {
    const { icloudEmail, icloudAppPassword, lastEvents } = await chrome.storage.local.get(['icloudEmail', 'icloudAppPassword', 'lastEvents']);
    if (!icloudEmail || !icloudAppPassword) {
      return { success: false, error: 'iCloud credentials not configured — open the popup to save them.' };
    }
    if (!lastEvents || !lastEvents.length) {
      return { success: false, error: 'No shift data available — run a sync first.' };
    }
    const client = new iCloudCalDAVClient(icloudEmail, icloudAppPassword);
    setProgress('Connecting to iCloud…', 68);
    await client.connect();
    setProgress('Loading Work Shifts calendar…', 71);
    const calendarUrl = await client.findOrCreateCalendar('Work Shifts');
    setProgress('Clearing iCloud calendar…', 74);
    await client.clearAllOurEvents(calendarUrl);
    // Reset open shift tracking BEFORE re-syncing so all open shifts are re-added
    await chrome.storage.local.set({ syncedOpenShiftUids: [] });
    await client.syncEvents(calendarUrl, lastEvents, (step, fraction) => {
      setProgress(step, 80 + Math.round(fraction * 18)); // 80–98%
    });
    console.info('[ShiftsExport] iCloud clear & resync complete —', lastEvents.length, 'events');
    return { success: true };
  } catch (err) {
    console.error('[ShiftsExport] iCloud clear & resync error:', err);
    // Reset tracking set so the next regular sync re-adds any open shifts that
    // didn't make it to iCloud due to this failure.
    await chrome.storage.local.set({ syncedOpenShiftUids: [] }).catch(() => {});
    return { success: false, error: err.message };
  }
}

async function syncToiCloud(events) {
  // No overall Promise.race timeout — per-request 20s timeouts in request() prevent
  // hanging. Promise.race leaks the inner async continuation, which can save
  // syncedOpenShiftUids to storage after a timeout, marking open shifts as synced
  // even when they never reached iCloud.
  try {
    const { icloudEmail, icloudAppPassword } = await chrome.storage.local.get(['icloudEmail', 'icloudAppPassword']);
    if (!icloudEmail || !icloudAppPassword) {
      return { success: false, error: 'iCloud credentials not configured — open the popup to save them.' };
    }
    const client = new iCloudCalDAVClient(icloudEmail, icloudAppPassword);
    setProgress('Connecting to iCloud…', 75);
    await client.connect();
    setProgress('Loading Work Shifts calendar…', 79);
    const calendarUrl = await client.findOrCreateCalendar('Work Shifts');
    await client.syncEvents(calendarUrl, events, (step, fraction) => {
      setProgress(step, 83 + Math.round(fraction * 14)); // 83–97%
    });
    console.info('[ShiftsExport] iCloud sync complete —', events.length, 'events');
    return { success: true };
  } catch (err) {
    console.error('[ShiftsExport] iCloud CalDAV error:', err);
    return { success: false, error: err.message };
  }
}

// ─── Message Handler (from popup) ────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === 'SYNC_PROGRESS') {
    chrome.storage.local.set({ syncRunning: true, syncStep: msg.step, syncPercent: msg.percent }).catch(() => {});
    return false;
  }

  if (msg.action === 'EXPORT_NOW') {
    runExport({ auto: false }).then((r) => {
      try { sendResponse(r); } catch {}
    });
    return true;
  }

  if (msg.action === 'GET_STATUS') {
    chrome.storage.local.get(['lastExport', 'lastCount', 'userName', 'importToOutlook'], (data) => {
      try { sendResponse(data); } catch {}
    });
    return true;
  }

  if (msg.action === 'SET_IMPORT_TO_OUTLOOK') {
    chrome.storage.local.set({ importToOutlook: msg.value });
    return false;
  }

  if (msg.action === 'SET_INCLUDE_OPEN_SHIFTS') {
    chrome.storage.local.set({ includeOpenShifts: msg.value });
    return false;
  }

  if (msg.action === 'DOWNLOAD_ICS') {
    chrome.storage.local.get('lastICS', (data) => {
      if (!data.lastICS) {
        try { sendResponse({ success: false, error: 'No ICS available — run an export first.' }); } catch {}
        return;
      }
      const filename = buildFilename();
      downloadICS(data.lastICS, filename).then(() => {
        try { sendResponse({ success: true }); } catch {}
      }).catch((err) => {
        try { sendResponse({ success: false, error: err.message }); } catch {}
      });
    });
    return true;
  }

  if (msg.action === 'CLEAR_AND_REIMPORT') {
    (async () => {
      try {
        const { importToOutlook, importToiCloud } = await chrome.storage.local.get(['importToOutlook', 'importToiCloud']);

        // Scrape fresh shifts — skip iCloud sync here; we handle it below with a full clear+resync
        const exportResult = await runExport({ auto: false, skipICloud: true });
        if (!exportResult.success) {
          try { sendResponse({ success: false, error: exportResult.error }); } catch {}
          return;
        }

        const { lastICS, lastCount } = await chrome.storage.local.get(['lastICS', 'lastCount']);
        let outlookResult = null;
        let icloudResult = null;

        // Clear + reimport to Outlook
        if (importToOutlook) {
          if (!lastICS) {
            outlookResult = { success: false, error: 'No ICS content available' };
          } else {
            const outlookTab = await getOrOpenOutlookTab();
            await chrome.scripting.executeScript({
              target: { tabId: outlookTab.id },
              files: ['outlook_content.js'],
            });
            await sleep(500);
            outlookResult = await chrome.tabs.sendMessage(outlookTab.id, {
              action: 'CLEAR_AND_IMPORT_ICS',
              icsContent: lastICS,
            });
          }
        }

        // Clear + reimport to iCloud
        if (importToiCloud) {
          icloudResult = await clearAndResyncToiCloud();
        }

        const success = (!importToOutlook || outlookResult?.success) &&
                        (!importToiCloud || icloudResult?.success);
        try { sendResponse({ success, count: lastCount, outlookResult, icloudResult }); } catch {}
      } catch (err) {
        try { sendResponse({ success: false, error: err.message }); } catch {}
      }
    })();
    return true;
  }
});
