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

// ─── Export Logic ─────────────────────────────────────────────────────────────

async function runExport({ auto = false } = {}) {
  try {
    const tab = await getOrOpenTeamsShiftsTab(auto);
    if (!tab) {
      console.warn('[ShiftsExport] No Teams Shifts tab available.');
      // Notify user if auto (they probably have Teams closed)
      if (auto) {
        chrome.notifications.create({
          type: 'basic',
          iconUrl: 'icon.png',
          title: 'Teams Shifts Export',
          message: 'Auto-export skipped — Teams is not open. Open Teams and click Export.',
        });
      }
      return { success: false, error: 'No Teams tab found' };
    }

    // Wait briefly for the page to be ready
    await sleep(2000);

    // Step 1: inject into top frame and navigate to Shifts
    await chrome.scripting.executeScript({ target: { tabId: tab.id, frameIds: [0] }, files: ['content.js'] });
    await sleep(300);
    await chrome.tabs.sendMessage(tab.id, { action: 'NAVIGATE_TO_SHIFTS' }, { frameId: 0 });
    await sleep(3000); // wait for Shifts iframe to load

    // Step 2: find the Shifts iframe frame ID
    const frameResults = await chrome.scripting.executeScript({
      target: { tabId: tab.id, allFrames: true },
      func: () => window.location.href,
    });
    const shiftsFrame = frameResults.find(
      (r) => r.result && r.result.includes('flw.teams.cloud.microsoft')
    );
    if (!shiftsFrame) throw new Error('Shifts iframe not found — make sure you are on teams.cloud.microsoft');

    // Step 3: inject content script into the iframe
    await chrome.scripting.executeScript({ target: { tabId: tab.id, frameIds: [shiftsFrame.frameId] }, files: ['content.js'] });
    await sleep(500);

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

    // Filter out open shifts if the user disabled them
    const { includeOpenShifts } = await chrome.storage.local.get('includeOpenShifts');
    let events = response.events || [];
    if (includeOpenShifts === false) {
      events = events.filter((e) => !e.isOpenShift);
    }

    // Merge freshly scraped events with stored history, then rebuild ICS
    const mergedEvents = await mergeWithHistory(events);
    const mergedICS = generateICS(mergedEvents);

    const filename = buildFilename();
    await downloadICS(mergedICS, filename);

    // Import to Outlook Web if the setting is enabled
    const { importToOutlook } = await chrome.storage.local.get('importToOutlook');
    let outlookResult = null;
    if (importToOutlook) {
      outlookResult = await importToOutlookWeb(mergedICS, auto);
    }

    // Update last export time and store ICS for clear & re-import
    await chrome.storage.local.set({ lastExport: Date.now(), lastCount: mergedEvents.length, lastICS: mergedICS });

    return { success: true, count: mergedEvents.length, outlookResult };
  } catch (err) {
    console.error('[ShiftsExport] Export error:', err);
    return { success: false, error: err.message };
  }
}

async function getOrOpenTeamsShiftsTab(autoMode) {
  // Look for an existing Teams tab
  const tabs = await chrome.tabs.query({ url: 'https://teams.cloud.microsoft/*' });

  if (tabs.length > 0) {
    // Prefer a tab that already has Shifts open
    const shiftsTab = tabs.find(
      (t) => t.url && (t.url.includes('scheduling') || t.url.includes('shifts'))
    );
    const tab = shiftsTab || tabs[0];

    return tab;
  }

  // In auto mode, don't open a new tab without user interaction
  if (autoMode) return null;

  // In manual mode, open a new Teams Shifts tab
  const newTab = await chrome.tabs.create({ url: TEAMS_SHIFTS_URL, active: true });
  await sleep(4000); // wait for Teams to load
  return newTab;
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

  const newTab = await chrome.tabs.create({ url: OUTLOOK_CALENDAR_URL, active: true });
  await sleep(5000); // wait for Outlook Web to load
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

// ─── Message Handler (from popup) ────────────────────────────────────────────

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
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

  if (msg.action === 'CLEAR_AND_REIMPORT') {
    (async () => {
      try {
        // First run the export to get fresh ICS
        const exportResult = await runExport({ auto: false });
        if (!exportResult.success) {
          try { sendResponse({ success: false, error: exportResult.error }); } catch {}
          return;
        }

        // Get the stored ICS content
        const { lastICS } = await chrome.storage.local.get('lastICS');
        if (!lastICS) {
          try { sendResponse({ success: false, error: 'No ICS content available' }); } catch {}
          return;
        }

        // Send clear + import to Outlook
        const outlookTab = await getOrOpenOutlookTab();
        await chrome.scripting.executeScript({
          target: { tabId: outlookTab.id },
          files: ['outlook_content.js'],
        });
        await sleep(500);

        const result = await chrome.tabs.sendMessage(outlookTab.id, {
          action: 'CLEAR_AND_IMPORT_ICS',
          icsContent: lastICS,
        });
        try { sendResponse(result); } catch {}
      } catch (err) {
        try { sendResponse({ success: false, error: err.message }); } catch {}
      }
    })();
    return true;
  }
});
