// background.js — Firefox MV2 background page (non-persistent / event page)
// Uses browser.* promises natively; chrome.* also works via Firefox's compatibility shim.

const ALARM_NAME = 'daily-shifts-export';
const TEAMS_SHIFTS_URL = 'https://teams.cloud.microsoft/';
const OUTLOOK_CALENDAR_URL = 'https://outlook.office.com/calendar/view/month';

// ─── Alarm Setup ─────────────────────────────────────────────────────────────

browser.runtime.onInstalled.addListener(() => {
  scheduleDailyAlarm();
});

browser.runtime.onStartup.addListener(() => {
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

// ─── Export Logic ─────────────────────────────────────────────────────────────

async function runExport({ auto = false } = {}) {
  try {
    const tab = await getOrOpenTeamsShiftsTab(auto);
    if (!tab) {
      console.warn('[ShiftsExport] No Teams Shifts tab available.');
      if (auto) {
        browser.notifications.create({
          type: 'basic',
          iconUrl: 'icon.png',
          title: 'Teams Shifts Export',
          message: 'Auto-export skipped — Teams is not open. Open Teams and click Export.',
        });
      }
      return { success: false, error: 'No Teams tab found' };
    }

    await sleep(2000);

    // Step 1: inject into top frame and navigate to Shifts
    await browser.tabs.executeScript(tab.id, { file: 'content.js', frameId: 0 });
    await sleep(300);
    await browser.tabs.sendMessage(tab.id, { action: 'NAVIGATE_TO_SHIFTS' }, { frameId: 0 });
    await sleep(3000);

    // Step 2: find the Shifts iframe frame ID
    const frames = await browser.webNavigation.getAllFrames({ tabId: tab.id });
    const shiftsFrame = frames.find(
      (f) => f.url && f.url.includes('flw.teams.cloud.microsoft')
    );
    if (!shiftsFrame) throw new Error('Shifts iframe not found');

    // Step 3: inject content script into the iframe
    await browser.tabs.executeScript(tab.id, { file: 'content.js', frameId: shiftsFrame.frameId });
    await sleep(500);

    // Auto-detect user name from Teams top frame ("Account manager for ...")
    let { userName } = await browser.storage.local.get('userName');
    try {
      const nameResults = await browser.tabs.executeScript(tab.id, {
        frameId: 0,
        code: `(function() {
          const btn = document.querySelector('button[aria-label^="Account manager for"]');
          if (btn) {
            const match = btn.getAttribute('aria-label').match(/Account manager for (.+)/);
            return match ? match[1] : null;
          }
          return null;
        })();`,
      });
      const detectedName = nameResults?.[0];
      if (detectedName) {
        userName = detectedName;
        await browser.storage.local.set({ userName: detectedName, userNameAutoDetected: true });
        console.info('[ShiftsExport] Auto-detected user name:', detectedName);
      }
    } catch (e) {
      console.info('[ShiftsExport] Name auto-detect failed, using saved name:', userName);
    }

    const response = await browser.tabs.sendMessage(
      tab.id,
      { action: 'SCRAPE_AND_EXPORT', userName: userName || null },
      { frameId: shiftsFrame.frameId }
    );

    if (!response || !response.success) {
      throw new Error(response?.error || 'Scrape failed');
    }

    // Filter out open shifts if the user disabled them
    const { includeOpenShifts } = await browser.storage.local.get('includeOpenShifts');
    let events = response.events || [];
    if (includeOpenShifts === false) {
      events = events.filter((e) => !e.isOpenShift);
    }

    // Merge freshly scraped events with stored history, then rebuild ICS
    const mergedEvents = await mergeWithHistory(events);
    const mergedICS = generateICS(mergedEvents);

    const filename = buildFilename();
    await downloadICS(mergedICS, filename);

    const { importToOutlook } = await browser.storage.local.get('importToOutlook');
    let outlookResult = null;
    if (importToOutlook) {
      outlookResult = await importToOutlookWeb(mergedICS, auto);
    }

    await browser.storage.local.set({ lastExport: Date.now(), lastCount: mergedEvents.length, lastICS: mergedICS });

    return { success: true, count: mergedEvents.length, outlookResult };
  } catch (err) {
    console.error('[ShiftsExport] Export error:', err);
    return { success: false, error: err.message };
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
  const { storedEvents = [] } = await browser.storage.local.get('storedEvents');

  // Only keep stored events that are already in the past
  const pastStored = storedEvents.filter((e) => e.endMs < now);

  // Merge past history + all newly scraped events, then deduplicate
  const merged = deduplicateEvents([...pastStored, ...newEvents]);

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

// ─── Download ─────────────────────────────────────────────────────────────────

function buildFilename() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `teams-shifts-${y}${m}${d}.ics`;
}

async function downloadICS(icsContent, filename) {
  // Firefox requires a data URL or object URL for downloads.download()
  // Using a data URL avoids the need to revoke it from a service worker context.
  const encoded = encodeURIComponent(icsContent);
  const dataUrl = `data:text/calendar;charset=utf-8,${encoded}`;

  await browser.downloads.download({
    url: dataUrl,
    filename,
    saveAs: false,
    conflictAction: 'overwrite',
  });
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// ─── Message Handler (from popup) ────────────────────────────────────────────

browser.runtime.onMessage.addListener((msg) => {
  // Firefox MV2: return a Promise from the listener for async responses
  if (msg.action === 'EXPORT_NOW') {
    return runExport({ auto: false });
  }

  if (msg.action === 'GET_STATUS') {
    return browser.storage.local.get(['lastExport', 'lastCount', 'userName', 'importToOutlook']);
  }

  if (msg.action === 'SET_IMPORT_TO_OUTLOOK') {
    return browser.storage.local.set({ importToOutlook: msg.value });
  }

  if (msg.action === 'SET_INCLUDE_OPEN_SHIFTS') {
    return browser.storage.local.set({ includeOpenShifts: msg.value });
  }

  if (msg.action === 'CLEAR_AND_REIMPORT') {
    return (async () => {
      try {
        // First run the export to get fresh ICS
        const exportResult = await runExport({ auto: false });
        if (!exportResult.success) {
          return { success: false, error: exportResult.error };
        }

        // Get the stored ICS content
        const { lastICS } = await browser.storage.local.get('lastICS');
        if (!lastICS) {
          return { success: false, error: 'No ICS content available' };
        }

        // Send clear + import to Outlook
        const outlookTab = await getOrOpenOutlookTab();
        await browser.tabs.executeScript(outlookTab.id, { file: 'outlook_content.js' });
        await sleep(500);

        const result = await browser.tabs.sendMessage(outlookTab.id, {
          action: 'CLEAR_AND_IMPORT_ICS',
          icsContent: lastICS,
        });
        return result;
      } catch (err) {
        return { success: false, error: err.message };
      }
    })();
  }
});
