// background.js — Firefox MV2 background page (non-persistent / event page)
// Uses browser.* promises natively; chrome.* also works via Firefox's compatibility shim.

const ALARM_NAME = 'daily-shifts-export';

// ─── Capture Teams auth headers via webRequest ────────────────────────────────

let capturedShiftsHeaders = null;
let capturedApiBase = null;
let capturedTenantId = null;

browser.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    if (!details.requestHeaders) return;
    const hasAuth = details.requestHeaders.some((h) => h.name.toLowerCase() === 'authorization');
    if (hasAuth && details.url.includes('/api/')) {
      capturedShiftsHeaders = [...details.requestHeaders];
      try {
        const url = new URL(details.url);
        const region = url.pathname.split('/').filter(Boolean)[0];
        if (region) capturedApiBase = `${url.origin}/${region}/api`;
        const tenantMatch = url.pathname.match(/\/tenants\/([^/]+)/);
        if (tenantMatch) capturedTenantId = tenantMatch[1];
      } catch (e) {
        console.error('[ShiftsExport] webRequest parse error:', e);
      }
    }
  },
  { urls: ['https://flw.teams.cloud.microsoft/*'] },
  ['requestHeaders']
);
const TEAMS_SHIFTS_URL = 'https://teams.cloud.microsoft/';
const OUTLOOK_CALENDAR_URL = 'https://outlook.office.com/calendar/view/month';

// ─── Alarm Setup ─────────────────────────────────────────────────────────────

browser.runtime.onInstalled.addListener(() => {
  clearProgress();
  scheduleDailyAlarm();
});

browser.runtime.onStartup.addListener(() => {
  clearProgress();
  scheduleDailyAlarm();
});

function scheduleDailyAlarm() {
  browser.alarms.get(ALARM_NAME).then((existing) => {
    if (!existing) {
      browser.alarms.create(ALARM_NAME, {
        periodInMinutes: 24 * 60,
        when: Date.now() + 60 * 1000,
      });
    }
  });
}

// ─── Alarm Handler ───────────────────────────────────────────────────────────

browser.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name === ALARM_NAME) {
    await runExport({ auto: true });
  }
});

// ─── Progress Helpers ─────────────────────────────────────────────────────────

function setProgress(step, percent) {
  browser.storage.local.set({ syncRunning: true, syncStep: step, syncPercent: percent }).catch(() => {});
}

function clearProgress() {
  browser.storage.local.set({ syncRunning: false, syncStep: '', syncPercent: 0, syncCancelled: false }).catch(() => {});
}

async function checkCancelled() {
  const data = await browser.storage.local.get('syncCancelled');
  if (data.syncCancelled) throw new Error('Sync cancelled');
}

// Watchdog: if syncPercent doesn't change for idleMs, cancel the sync automatically
function startWatchdog(idleMs = 90000) {
  let lastPercent = -1;
  let lastChange = Date.now();
  const interval = setInterval(async () => {
    try {
      const data = await browser.storage.local.get(['syncRunning', 'syncPercent']);
      if (!data.syncRunning) { clearInterval(interval); return; }
      if (data.syncPercent !== lastPercent) {
        lastPercent = data.syncPercent;
        lastChange = Date.now();
      } else if (Date.now() - lastChange >= idleMs) {
        clearInterval(interval);
        await browser.storage.local.set({ syncCancelled: true });
      }
    } catch {}
  }, 5000);
  return () => clearInterval(interval);
}

// ─── Export Logic ─────────────────────────────────────────────────────────────

function getTargetEndDate() {
  const now = new Date();
  const twoWeeksFromNow = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
  const year = now.getFullYear();
  const cands = [
    new Date(year, 1, 28), new Date(year, 7, 31),
    new Date(year+1, 1, 28), new Date(year+1, 7, 31),
    new Date(year+2, 1, 28),
  ];
  cands.forEach((d, i) => {
    if (d.getMonth() === 1) {
      const ly = d.getFullYear();
      if ((ly % 4 === 0 && ly % 100 !== 0) || ly % 400 === 0) cands[i] = new Date(ly, 1, 29);
    }
  });
  const future = cands.filter((d) => d > now).sort((a, b) => a - b);
  return (future.length > 1 && future[0] <= twoWeeksFromNow) ? future[1] : future[0];
}

async function hasShiftsFrame(tabId) {
  try {
    const frames = await browser.webNavigation.getAllFrames({ tabId });
    return frames.some((f) => f.url && f.url.includes('flw.teams.cloud.microsoft'));
  } catch {
    return false;
  }
}

async function ensureShiftsFrameLoaded() {
  let tabs = await browser.tabs.query({ url: 'https://teams.cloud.microsoft/*' });
  if (!tabs.length) {
    setProgress('Opening Teams...', 3);
    const tab = await browser.tabs.create({ url: 'https://teams.cloud.microsoft/', active: true });
    // Wait for tab to fully load (status: complete) before navigating
    const deadline = Date.now() + 30000;
    while (Date.now() < deadline) {
      await sleep(1000);
      try {
        const t = await browser.tabs.get(tab.id);
        if (t.status === 'complete') break;
      } catch { break; }
    }
    // Give Teams JS a moment to initialize after load
    await sleep(2000);
    tabs = [tab];
  }
  for (const tab of tabs) {
    // Try navigating up to 2 times — the "Almost there" dialog causes a reload
    // which destroys the content script; we wait for reload then re-navigate
    for (let nav = 0; nav < 2; nav++) {
      // Wait for tab to be fully loaded before navigating
      const loadDeadline = Date.now() + 15000;
      while (Date.now() < loadDeadline) {
        await sleep(500);
        try {
          const t = await browser.tabs.get(tab.id);
          if (t.status === 'complete') break;
        } catch { break; }
      }
      await sleep(1500);

      try {
        await browser.tabs.sendMessage(tab.id, { action: 'NAVIGATE_TO_SHIFTS' });
      } catch {}

      // If clicking "Continue" on the Teams permission dialog caused a page reload,
      // wait for the reload to finish then navigate to Shifts on the reloaded page.
      await sleep(500);
      const tabAfterNav = await browser.tabs.get(tab.id).catch(() => null);
      if (tabAfterNav?.status === 'loading') {
        const reloadDeadline = Date.now() + 20000;
        while (Date.now() < reloadDeadline) {
          await sleep(500);
          const t2 = await browser.tabs.get(tab.id).catch(() => null);
          if (!t2 || t2.status === 'complete') break;
        }
        await sleep(2000);
        try {
          await browser.tabs.sendMessage(tab.id, { action: 'NAVIGATE_TO_SHIFTS' });
        } catch {}
        await sleep(2000);
      }

      const frameDeadline = Date.now() + 20000;
      while (Date.now() < frameDeadline) {
        await sleep(500);
        if (await hasShiftsFrame(tab.id)) return true;
      }

      if (nav === 0) setProgress('Waiting for Teams to finish loading...', 6);
    }
  }
  return false;
}

async function fetchShifts() {
  for (let attempt = 0; attempt < 2; attempt++) {
    if (!capturedShiftsHeaders) {
      await ensureShiftsFrameLoaded();
      const deadline = Date.now() + 10000;
      while (Date.now() < deadline) {
        if (capturedShiftsHeaders) break;
        await sleep(500);
      }
    }

    if (!capturedShiftsHeaders || !capturedApiBase) {
      if (attempt === 0) { capturedShiftsHeaders = null; continue; }
      throw new Error('Could not capture Teams auth — open Teams and navigate to Shifts first.');
    }

    const headers = {};
    for (const h of capturedShiftsHeaders) headers[h.name] = h.value;

    const teamsResp = await fetch(`${capturedApiBase}/users/me/teams`, { headers });
    const teamsText = await teamsResp.text();
    if (!teamsResp.ok || teamsText.trimStart().startsWith('<')) {
      capturedShiftsHeaders = null;
      if (attempt === 0) continue;
      throw new Error(`Teams API ${teamsResp.status}: ${teamsText.slice(0, 200)}`);
    }
  const teamsData = JSON.parse(teamsText);
  const teamList = teamsData.teams || teamsData.value || (Array.isArray(teamsData) ? teamsData : []);
  const teamIds = teamList.map((t) => t.team?.id || t.id || t.teamId).filter(Boolean);
  if (!teamIds.length) {
      console.warn('[ShiftsExport] No teams found. Keys:', Object.keys(teamsData||{}), 'Sample:', JSON.stringify(teamsData).slice(0,400));
      throw new Error('No teams found in Shifts. Make sure you have a schedule set up in Teams Shifts.');
    }

  const startTime = new Date();
  startTime.setHours(0, 0, 0, 0);
  const endTime = getTargetEndDate();
  const params = new URLSearchParams({
    teamIds: teamIds.join(','),
    startTime: startTime.toISOString(),
    endTime: endTime.toISOString(),
    includeShifts: 'true',
    includeNotes: 'true',
    includeOpenShifts: 'true',
    includeDraft: 'true',
  });

    // Fetch open shift sign-up requests to know which availability slots the user has signed up for
    const signedUpOpenShiftIds = new Set();
    const tenantId = capturedTenantId || getTenantIdFromHeaders(capturedShiftsHeaders);
    if (tenantId) {
      for (const teamId of teamIds) {
        try {
          const reqUrl = `${capturedApiBase}/tenants/${tenantId}/teams/${teamId}/shifts/open/requests?startTime=${startTime.toISOString()}&endTime=${endTime.toISOString()}`;
          const reqResp = await fetch(reqUrl, { headers });
          if (reqResp.ok) {
            const reqText = await reqResp.text();
            if (!reqText.trimStart().startsWith('<')) {
              const reqData = JSON.parse(reqText);
              const requests = reqData.requests || reqData.openShiftChangeRequests || reqData.value || (Array.isArray(reqData) ? reqData : []);
              for (const req of requests) {
                const state = req.state || req.requestState;
                const shiftId = req.shiftId || req.openShiftId;
                if (shiftId && (state === 'WaitingOnManager' || state === 'ManagerApproved' || state === 'pending' || state === 'approved')) {
                  signedUpOpenShiftIds.add(shiftId);
                }
              }
            }
          }
        } catch (e) {
          console.warn('[ShiftsExport] Could not fetch availability sign-up requests:', e.message);
        }
      }
    }

    const parseShiftItem = (item) => {
      const start = new Date(item.startTime || item.startDateTime || item.StartDateTime || item.start);
      const end = new Date(item.endTime || item.endDateTime || item.EndDateTime || item.end);
      if (isNaN(start) || isNaN(end)) return null;
      const notes = item.notes || item.Notes || '';
      const rawType = item.shiftType || item.theme || 'Shift';
      const summary = item.title || item.displayName || (rawType === 'Absence' ? 'Time Off' : rawType);
      const isAllDay = rawType === 'Absence';
      return { startMs: start.getTime(), endMs: end.getTime(), summary, notes, isAllDay };
    };

    const events = [];
    let nextUrl = `${capturedApiBase}/users/me/dataindaterange?${params}`;
    let pageCount = 0;

    while (nextUrl && pageCount < 20) {
      const resp = await fetch(nextUrl, { headers });
      const text = await resp.text();
      if (!resp.ok || text.trimStart().startsWith('<')) {
        capturedShiftsHeaders = null;
        if (attempt === 0) { nextUrl = null; break; }
        throw new Error(`Shifts API ${resp.status}: ${text.slice(0, 200)}`);
      }
      const data = JSON.parse(text);

      for (const shift of (data.shifts || data.Shifts || [])) {
        const parsed = parseShiftItem(shift);
        if (parsed) events.push({ ...parsed, isOpenShift: false });
      }
      for (const shift of (data.openShifts || data.OpenShifts || [])) {
        const parsed = parseShiftItem(shift);
        if (parsed) events.push({ ...parsed, isOpenShift: true, isAvailabilitySignup: signedUpOpenShiftIds.has(shift.id) });
      }
      for (const tdo of (data.timesOff || data.TimesOff || data.timeOffRequests || [])) {
        const startStr = tdo.startTime || tdo.startDateTime || tdo.StartDateTime;
        if (!startStr) continue;
        const start = new Date(startStr);
        const endStr = tdo.endTime || tdo.endDateTime || tdo.EndDateTime;
        const end = endStr ? new Date(endStr) : new Date(start.getTime() + 86400000);
        const summary = tdo.title || tdo.reason?.displayName || tdo.displayName || 'Time Off';
        events.push({ summary, notes: tdo.notes || '', startMs: start.getTime(), endMs: end.getTime(), isOpenShift: false, isAllDay: true });
      }

      nextUrl = data.nextLink || data['@odata.nextLink'] || null;
      pageCount++;
    }

    if (attempt === 0 && events.length === 0) { capturedShiftsHeaders = null; continue; }
    return events;
  }
}

async function runExport({ auto = false, skipICloud = false } = {}) {
  const { syncRunning } = await browser.storage.local.get('syncRunning');
  if (syncRunning) return { success: false, error: 'Sync already in progress' };

  const stopWatchdog = startWatchdog(120000);
  try {
    browser.storage.local.set({ lastError: null });
    setProgress('Fetching shifts...', 10);
    await checkCancelled();

    let events = await fetchShifts();
    await checkCancelled();
    setProgress('Processing shifts...', 70);

    const { includeOpenShifts, includeSignedUpSlots } = await browser.storage.local.get(['includeOpenShifts', 'includeSignedUpSlots']);
    const scheduled = events.filter((e) => !e.isOpenShift);
    events = events.filter((e) => {
      if (!e.isOpenShift) return true;
      if (e.isAvailabilitySignup) return includeSignedUpSlots === true;
      if (includeOpenShifts !== true) return false;
      if (/^\d{4}\b/.test(e.summary)) return false;
      return isEligibleOpenShift(e, scheduled);
    });

    const mergedEvents = await mergeWithHistory(events);
    const mergedICS = generateICS(mergedEvents);


    const { importToOutlook } = await browser.storage.local.get('importToOutlook');
    let outlookResult = null;
    if (importToOutlook) outlookResult = await importToOutlookWeb(mergedICS, auto);

    const { importToiCloud } = await browser.storage.local.get('importToiCloud');
    let icloudResult = null;
    if (importToiCloud && !skipICloud) icloudResult = await syncToiCloud(mergedEvents);

    await browser.storage.local.set({ lastExport: Date.now(), lastCount: mergedEvents.length, lastICS: mergedICS, lastEvents: mergedEvents, lastExportSuccess: Date.now() });
    clearProgress();
    browser.storage.local.set({ lastError: null });
    return { success: true, count: mergedEvents.length, outlookResult, icloudResult };
  } catch (err) {
    console.error('[ShiftsExport] Export error:', err);
    const errMsg = err?.message || (typeof err === 'string' ? err : '') || err?.name || 'Unknown error';
    browser.storage.local.set({ lastError: errMsg, lastErrorTime: Date.now() });
    clearProgress();
    return { success: false, error: errMsg };
  } finally {
    stopWatchdog();
  }
}

async function getOrOpenTeamsShiftsTab(autoMode) {
  const tabs = await browser.tabs.query({ url: 'https://teams.cloud.microsoft/*' });

  if (tabs.length > 0) {
    const shiftsTab = tabs.find(
      (t) => t.url && (t.url.includes('scheduling') || t.url.includes('shifts'))
    );
    const tab = shiftsTab || tabs[0];

    if (!tab.url.includes('scheduling') && !tab.url.includes('shifts')) {
      await browser.tabs.update(tab.id, { url: TEAMS_SHIFTS_URL });
      await sleep(3000);
    }
    return tab;
  }

  if (autoMode) return null;

  const newTab = await browser.tabs.create({ url: TEAMS_SHIFTS_URL, active: true });
  await sleep(4000);
  return newTab;
}

// ─── Outlook Web Import ───────────────────────────────────────────────────────

async function importToOutlookWeb(icsContent, autoMode) {
  try {
    const outlookTab = await getOrOpenOutlookTab(autoMode);
    if (!outlookTab) {
      return { success: false, error: 'Outlook Web not open' };
    }

    await sleep(2000);

    // Inject Outlook content script in case the tab was open before the extension loaded
    await browser.tabs.executeScript(outlookTab.id, { file: 'outlook_content.js' });
    await sleep(300);

    const response = await browser.tabs.sendMessage(outlookTab.id, {
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
  for (const pattern of ['https://outlook.office.com/*', 'https://outlook.office365.com/*']) {
    const tabs = await browser.tabs.query({ url: pattern });
    if (tabs.length > 0) return tabs[0];
  }

  if (autoMode) return null;

  const newTab = await browser.tabs.create({ url: OUTLOOK_CALENDAR_URL, active: true });
  await sleep(5000);
  return newTab;
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
  const { storedEvents = [], excludedStartMs = [] } = await browser.storage.local.get(['storedEvents', 'excludedStartMs']);
  const excluded = new Set(excludedStartMs);

  // Validate storedEvents is an array with expected shape
  if (!Array.isArray(storedEvents)) {
    console.warn('[ShiftsExport] storedEvents corrupted, resetting');
    await browser.storage.local.set({ storedEvents: [] });
    return newEvents.filter((e) => !excluded.has(e.startMs));
  }

  const pastStored = storedEvents.filter((e) => {
    if (!e || typeof e !== 'object') return false;
    if (typeof e.startMs !== 'number' || typeof e.endMs !== 'number') return false;
    return e.endMs < now && !excluded.has(e.startMs);
  });
  const filtered = newEvents.filter((e) => !excluded.has(e.startMs));
  const merged = deduplicateEvents([...pastStored, ...filtered]);

  await browser.storage.local.set({ storedEvents: merged });
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
    const summaryText = ev.isAvailabilitySignup ? `Signed Up: ${ev.summary}` : ev.isOpenShift ? `OPEN: ${ev.summary}` : ev.summary;
    lines.push('BEGIN:VEVENT');
    lines.push(`UID:teams-shift-${ev.startMs}-${i}@shifts-export`);
    lines.push(`DTSTAMP:${toICSDate(Date.now())}`);
    if (ev.isAllDay) {
      lines.push(`DTSTART;VALUE=DATE:${toICSDateOnlyStr(ev.startMs)}`);
      lines.push(`DTEND;VALUE=DATE:${toICSDateOnlyStr(ev.endMs)}`);
    } else {
      lines.push(`DTSTART:${toICSDate(ev.startMs)}`);
      lines.push(`DTEND:${toICSDate(ev.endMs)}`);
    }
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

// ─── Download ─────────────────────────────────────────────────────────────────

function isEligibleOpenShift(openShift, scheduledShifts) {
  const minGapMs = 8 * 60 * 60 * 1000;
  for (const s of scheduledShifts) {
    if (openShift.startMs < s.endMs && openShift.endMs > s.startMs) return false;
    const gap = openShift.startMs >= s.endMs
      ? openShift.startMs - s.endMs
      : s.startMs - openShift.endMs;
    if (gap < minGapMs) return false;
  }
  return true;
}

function buildFilename() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `teams-shifts-${y}${m}${d}.ics`;
}

async function downloadICS(icsContent, filename) {
  const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  await browser.downloads.download({ url, filename, saveAs: false, conflictAction: 'overwrite' });
  URL.revokeObjectURL(url);
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function getTenantIdFromHeaders(headers) {
  const authHeader = headers.find((h) => h.name.toLowerCase() === 'authorization');
  if (!authHeader) return null;
  try {
    const parts = authHeader.value.replace('Bearer ', '').split('.');
    if (parts.length < 2) return null;
    const payload = JSON.parse(atob(parts[1].replace(/-/g, '+').replace(/_/g, '/') + '=='));
    return payload.tid || null;
  } catch { return null; }
}

// ─── Load Polling Helpers ─────────────────────────────────────────────────────

// Wait for the tab to reach "complete" status (handles post-reload settling).
async function waitForTabComplete(tabId, timeoutMs = 10000) {
  await sleep(600); // let any in-flight navigation begin before we start polling
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const tab = await browser.tabs.get(tabId);
      if (tab.status === 'complete') return;
    } catch {}
    await sleep(400);
  }
}

// Poll until Teams sidebar navigation elements are available in the top frame
async function waitForTeamsReady(tabId, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const results = await browser.tabs.executeScript(tabId, {
        frameId: 0,
        code: `(function() {
          return !!(
            document.querySelector('[aria-label*="Shifts" i][role="button"]') ||
            document.querySelector('[aria-label*="more" i][role="button"]') ||
            document.querySelector('[aria-label*="Apps" i][role="button"]') ||
            document.querySelector('button[aria-label^="Account manager for"]')
          );
        })();`,
      });
      if (results?.[0]) return;
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
      const frames = await browser.webNavigation.getAllFrames({ tabId });
      const frame = frames.find((f) => f.url && f.url.includes('flw.teams.cloud.microsoft'));
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
      const results = await browser.tabs.executeScript(tabId, {
        frameId,
        code: `(function() {
          return !!(
            document.querySelector('button[aria-label="Go to next week"]') ||
            document.querySelector('button[aria-label*="Pick a date"]') ||
            document.querySelector('[data-tid="your-shifts-tab"]') ||
            document.querySelector('[data-tid="yourShifts-tab"]')
          );
        })();`,
      });
      if (results?.[0]) return;
    } catch {}
    await sleep(500);
  }
  throw new Error('Timed out waiting for Shifts UI to load');
}

// ─── iCloud CalDAV Sync ───────────────────────────────────────────────────────

// Date-only ICS formatter for all-day events (no time component)
function toICSDateOnlyStr(ms) {
  const pad = (n) => String(n).padStart(2, '0');
  const d = new Date(ms);
  return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
}

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
  const summaryText = event.isAvailabilitySignup ? `Signed Up: ${event.summary}` : event.isOpenShift ? `OPEN: ${event.summary}` : event.summary;
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Teams Shifts Export//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    `UID:${uid}@teams-shifts-export`,
    `DTSTAMP:${toICSDateStr(Date.now())}`,
    ...(event.isAllDay
      ? [`DTSTART;VALUE=DATE:${toICSDateOnlyStr(event.startMs)}`, `DTEND;VALUE=DATE:${toICSDateOnlyStr(event.endMs)}`]
      : [`DTSTART:${toICSDateStr(event.startMs)}`, `DTEND:${toICSDateStr(event.endMs)}`]),
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
      if (status === 0 || response.type === 'opaqueredirect') {
        throw new Error(`iCloud request got opaque redirect — credential or URL issue`);
      }
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
  // Retries up to 2 times on timeout/network errors to handle iCloud stalls.
  async putEvent(calendarUrl, uid, icsContent) {
    const base = calendarUrl.endsWith('/') ? calendarUrl : calendarUrl + '/';
    const url = `${base}${uid}.ics`;
    const headers = { 'Content-Type': 'text/calendar; charset=utf-8' };

    let lastErr;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        if (attempt > 0) await sleep(2000 * attempt);
        let res = await this.request(url, 'PUT', icsContent, { ...headers, 'If-None-Match': '*' });
        if (res.status === 412) {
          res = await this.request(url, 'PUT', icsContent, headers);
        }
        if (res.status < 200 || res.status >= 300) {
          throw new Error(`iCloud PUT failed (HTTP ${res.status}) for ${uid}`);
        }
        return; // success
      } catch (err) {
        lastErr = err;
        console.warn(`[iCloud] putEvent attempt ${attempt + 1} failed for ${uid}:`, err.message);
      }
    }
    throw lastErr;
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
    const { syncedOpenShiftUids: storedUids = [] } = await browser.storage.local.get('syncedOpenShiftUids');
    const syncedOpenShiftUids = new Set(storedUids);

    // Pre-calculate how many events will actually be uploaded so the fraction
    // fills 0→1 over real uploads, not skipped open shifts.
    const total = events.filter((e) => {
      if (!e.isOpenShift || e.isAvailabilitySignup) return true;
      const uid = `teams-shift-${e.startMs}-${e.summary.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()}`;
      return !syncedOpenShiftUids.has(uid);
    }).length || 1;

    for (const event of events) {
      const uid = `teams-shift-${event.startMs}-${event.summary.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()}`;
      currentUids.add(uid);

      if (event.isOpenShift && !event.isAvailabilitySignup) {
        // Only add open shifts we haven't pushed before
        if (!syncedOpenShiftUids.has(uid)) {
          await checkCancelled();
          if (onProgress) onProgress(`Uploading shift ${++uploaded} of ${total}…`, uploaded / total);
          await this.putEvent(calendarUrl, uid, buildSingleEventICS(event, uid));
          syncedOpenShiftUids.add(uid);
        }
      } else {
        // Scheduled shifts: always upsert
        await checkCancelled();
        if (onProgress) onProgress(`Uploading shift ${++uploaded} of ${total}…`, uploaded / total);
        await this.putEvent(calendarUrl, uid, buildSingleEventICS(event, uid));
        await sleep(250); // pace requests to avoid iCloud rate-limiting
      }
    }

    // Delete stale future events that are no longer in Teams
    if (onProgress) onProgress('Checking iCloud calendar…', 0.97);
    await checkCancelled();
    const existingOurEvents = await this.getOurEvents(calendarUrl);
    const stale = [...existingOurEvents.entries()].filter(([uid]) => {
      if (currentUids.has(uid)) return false;
      const startMs = parseInt(uid.replace('teams-shift-', ''), 10);
      return !isNaN(startMs) && startMs > now;
    });
    for (let i = 0; i < stale.length; i++) {
      const [uid, { url }] = stale[i];
      await checkCancelled();
      if (onProgress) onProgress(`Removing old shift ${i + 1} of ${stale.length}…`, 0.97 + (i / stale.length) * 0.02);
      await this.deleteEvent(url);
      syncedOpenShiftUids.delete(uid);
      console.info('[iCloud] Deleted stale event:', uid);
    }

    // Persist updated open shift tracking set
    await browser.storage.local.set({ syncedOpenShiftUids: [...syncedOpenShiftUids] });
  }

  // ── Regex-based XML helpers (DOMParser unavailable in MV2 background pages) ──

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
  const { syncRunning } = await browser.storage.local.get('syncRunning');
  if (syncRunning) return { success: false, error: 'Sync already in progress' };

  try {
    const { icloudEmail, icloudAppPassword, lastEvents } = await browser.storage.local.get(['icloudEmail', 'icloudAppPassword', 'lastEvents']);
    if (!icloudEmail || !icloudAppPassword) {
      return { success: false, error: 'iCloud credentials not configured — open the popup to save them.' };
    }
    if (!lastEvents || !lastEvents.length) {
      return { success: false, error: 'No shift data available — run a sync first.' };
    }
    const client = new iCloudCalDAVClient(icloudEmail, icloudAppPassword);
    setProgress('Connecting to iCloud…', 68);
    await client.connect();
    const regularEvents = lastEvents.filter(e => !e.isAvailabilitySignup);
    const signupEvents = lastEvents.filter(e => e.isAvailabilitySignup);

    setProgress('Loading Work Shifts calendar…', 71);
    const calendarUrl = await client.findOrCreateCalendar('Work Shifts');
    setProgress('Clearing Work Shifts calendar…', 73);
    await client.clearAllOurEvents(calendarUrl);

    const signupCalendarUrl = await client.findOrCreateCalendar('Open Time Signup');
    setProgress('Clearing Open Time Signup calendar…', 75);
    await client.clearAllOurEvents(signupCalendarUrl);

    // Reset open shift tracking BEFORE re-syncing so all open shifts are re-added
    await browser.storage.local.set({ syncedOpenShiftUids: [] });
    await client.syncEvents(calendarUrl, regularEvents, (step, fraction) => {
      setProgress(step, 80 + Math.round(fraction * (signupEvents.length ? 9 : 18)));
    });
    if (signupEvents.length) {
      await client.syncEvents(signupCalendarUrl, signupEvents, (step, fraction) => {
        setProgress(step, 89 + Math.round(fraction * 9));
      });
    }
    console.info('[ShiftsExport] iCloud clear & resync complete —', lastEvents.length, 'events');
    return { success: true };
  } catch (err) {
    console.error('[ShiftsExport] iCloud clear & resync error:', err);
    await browser.storage.local.set({ syncedOpenShiftUids: [] }).catch(() => {});
    return { success: false, error: err.message };
  }
}

async function syncToiCloud(events) {
  try {
    const { icloudEmail, icloudAppPassword } = await browser.storage.local.get(['icloudEmail', 'icloudAppPassword']);
    if (!icloudEmail || !icloudAppPassword) {
      return { success: false, error: 'iCloud credentials not configured — open the popup to save them.' };
    }
    const client = new iCloudCalDAVClient(icloudEmail, icloudAppPassword);
    setProgress('Connecting to iCloud…', 75);
    await client.connect();

    const regularEvents = events.filter(e => !e.isAvailabilitySignup);
    const signupEvents = events.filter(e => e.isAvailabilitySignup);

    setProgress('Loading Work Shifts calendar…', 79);
    const calendarUrl = await client.findOrCreateCalendar('Work Shifts');
    await client.syncEvents(calendarUrl, regularEvents, (step, fraction) => {
      setProgress(step, 83 + Math.round(fraction * (signupEvents.length ? 7 : 14)));
    });

    if (signupEvents.length) {
      setProgress('Loading Open Time Signup calendar…', 90);
      const signupCalendarUrl = await client.findOrCreateCalendar('Open Time Signup');
      await client.syncEvents(signupCalendarUrl, signupEvents, (step, fraction) => {
        setProgress(step, 90 + Math.round(fraction * 7));
      });
    }

    console.info('[ShiftsExport] iCloud sync complete —', events.length, 'events');
    return { success: true };
  } catch (err) {
    console.error('[ShiftsExport] iCloud CalDAV error:', err);
    return { success: false, error: err.message };
  }
}

// ─── Message Handler (from popup) ────────────────────────────────────────────

browser.runtime.onMessage.addListener((msg) => {
  // Firefox MV2: return a Promise from the listener for async responses
  if (msg.action === 'SYNC_PROGRESS') {
    browser.storage.local.set({ syncRunning: true, syncStep: msg.step, syncPercent: msg.percent }).catch(() => {});
    return false;
  }

  if (msg.action === 'CANCEL_SYNC') {
    clearProgress();
    browser.storage.local.set({ syncCancelled: true, lastError: null }).catch(() => {});
    return false;
  }

  if (msg.action === 'EXPORT_NOW') {
    return runExport({ auto: false });
  }

  if (msg.action === 'GET_STATUS') {
    return browser.storage.local.get(['lastExport', 'lastCount', 'userName', 'importToOutlook']);
  }

  if (msg.action === 'SET_IMPORT_TO_OUTLOOK') {
    return browser.storage.local.set({ importToOutlook: msg.value });
  }

  if (msg.action === 'SET_IMPORT_TO_ICLOUD') {
    return browser.storage.local.set({ importToiCloud: msg.value });
  }

  if (msg.action === 'SET_INCLUDE_OPEN_SHIFTS') {
    return browser.storage.local.set({ includeOpenShifts: msg.value });
  }

  if (msg.action === 'SET_INCLUDE_SIGNED_UP_SLOTS') {
    return browser.storage.local.set({ includeSignedUpSlots: msg.value });
  }

  if (msg.action === 'DOWNLOAD_ICS') {
    return (async () => {
      const { lastICS } = await browser.storage.local.get('lastICS');
      if (!lastICS) return { success: false, error: 'No ICS available — run a sync first.' };
      await downloadICS(lastICS, buildFilename());
      return { success: true };
    })();
  }

  if (msg.action === 'CLEAR_AND_REIMPORT') {
    return (async () => {
      try {
        const { importToOutlook, importToiCloud } = await browser.storage.local.get(['importToOutlook', 'importToiCloud']);

        // Scrape fresh shifts — skip iCloud sync here; we handle it below with a full clear+resync
        const exportResult = await runExport({ auto: false, skipICloud: true });
        if (!exportResult.success) {
          return { success: false, error: exportResult.error };
        }

        const { lastICS, lastCount } = await browser.storage.local.get(['lastICS', 'lastCount']);
        let outlookResult = null;
        let icloudResult = null;

        // Clear + reimport to Outlook
        if (importToOutlook) {
          if (!lastICS) {
            outlookResult = { success: false, error: 'No ICS content available' };
          } else {
            const outlookTab = await getOrOpenOutlookTab();
            await browser.tabs.executeScript(outlookTab.id, { file: 'outlook_content.js' });
            await sleep(500);
            outlookResult = await browser.tabs.sendMessage(outlookTab.id, {
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
        return { success, count: lastCount, outlookResult, icloudResult };
      } catch (err) {
        return { success: false, error: err.message };
      }
    })();
  }
});
