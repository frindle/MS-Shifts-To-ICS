// outlook_content.js — injected into outlook.office.com (Firefox)
// Handles automated ICS import into Outlook Web calendar

(function () {
  if (window.__outlookImportInitialized) return;
  window.__outlookImportInitialized = true;

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  const TARGET_CALENDAR = 'Work Shifts';

  // ─── Auto-focus Work Shifts calendar after import reload ───────────────
  async function autoFocusWorkShifts() {
    const { focusWorkShifts } = await browser.storage.local.get('focusWorkShifts');
    if (!focusWorkShifts) return;
    browser.storage.local.remove('focusWorkShifts');

    // Wait for the calendar sidebar to load
    for (let attempt = 0; attempt < 20; attempt++) {
      await sleep(500);
      const calItems = document.querySelectorAll(
        '[role="checkbox"][aria-label], [role="menuitemcheckbox"][aria-label]'
      );
      if (calItems.length === 0) continue;

      for (const item of calItems) {
        const label = (item.getAttribute('aria-label') || '').toLowerCase();
        const checked = item.getAttribute('aria-checked') === 'true';

        if (label.includes(TARGET_CALENDAR.toLowerCase())) {
          if (!checked) item.click();
        } else if (label.includes('calendar')) {
          if (checked) item.click();
        }
      }
      console.info('[ShiftsExport] Auto-focused Work Shifts calendar');
      return;
    }
    console.info('[ShiftsExport] Could not find calendar checkboxes to auto-focus');
  }
  autoFocusWorkShifts();

  // ─── Storage Helpers ───────────────────────────────────────────────────
  function storageGet(keys) {
    return browser.storage.local.get(keys);
  }
  function storageSet(data) {
    return browser.storage.local.set(data);
  }

  // ─── UI Overlay ──────────────────────────────────────────────────────────

  function showOverlay(message) {
    const old = document.getElementById('shifts-export-overlay');
    if (old) old.remove();

    const overlay = document.createElement('div');
    overlay.id = 'shifts-export-overlay';
    Object.assign(overlay.style, {
      position: 'fixed',
      top: '0',
      left: '0',
      width: '100vw',
      height: '100vh',
      zIndex: '999999',
      background: 'rgba(0,0,0,0.7)',
      color: '#fff',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontSize: '28px',
      fontFamily: 'sans-serif',
      pointerEvents: 'auto',
    });
    overlay.textContent = message;

    const style = document.createElement('style');
    style.id = 'shifts-export-overlay-style';
    style.textContent = 'body > *:not(#shifts-export-overlay) { pointer-events: none !important; }';
    document.head.appendChild(style);
    document.body.appendChild(overlay);

    return {
      update(msg) { overlay.textContent = msg; },
      remove() {
        overlay.remove();
        const s = document.getElementById('shifts-export-overlay-style');
        if (s) s.remove();
      },
    };
  }

  // API bases to try in order: OWA REST API first (matches LokiAuthToken audience),
  // then Microsoft Graph as fallback.
  const API_BASES = [
    'https://outlook.office.com/api/v2.0',
    'https://graph.microsoft.com/v1.0',
  ];

  // ─── Token Discovery ──────────────────────────────────────────────────────
  // OWA stores its access token directly as a raw JWT string under 'LokiAuthToken'
  // in sessionStorage — NOT under a JSON wrapper like MSAL.js normally does.

  function findMSALTokens() {
    const tokens = [];
    const seen   = new Set();

    const push = (t) => { if (t && t.startsWith('eyJ') && !seen.has(t)) { seen.add(t); tokens.push(t); } };

    // ① OWA primary token — plain JWT under 'LokiAuthToken'
    push(sessionStorage.getItem('LokiAuthToken'));

    // ② Account-specific LokiAuthToken variants (exclude EXP* expiry keys)
    for (let i = 0; i < sessionStorage.length; i++) {
      const key = sessionStorage.key(i);
      if (key.startsWith('LokiAuthToken') && !key.startsWith('EXPLokiAuthToken')) {
        push(sessionStorage.getItem(key));
      }
    }

    // ③ Use MSAL token-keys index to find exact access-token cache keys
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key.includes('token.keys.')) continue;
      try {
        const idx = JSON.parse(localStorage.getItem(key));
        const atKeys = idx?.accessToken || [];
        for (const atKey of atKeys) {
          try {
            const val = JSON.parse(localStorage.getItem(atKey));
            push(val?.secret);
          } catch {}
        }
      } catch {}
    }

    // ④ Fallback: case-insensitive scan for 'accesstoken'
    for (const store of [sessionStorage, localStorage]) {
      for (let i = 0; i < store.length; i++) {
        const key = store.key(i);
        if (!key.toLowerCase().includes('accesstoken')) continue;
        try {
          const val = JSON.parse(store.getItem(key));
          push(val?.secret || val?.accessToken || val?.access_token);
        } catch {}
      }
    }

    return tokens;
  }

  // Normalise Graph API (lowercase) and Outlook REST API (PascalCase) property names.
  const norm = (obj, ...keys) => {
    for (const k of keys) {
      if (obj?.[k] !== undefined) return obj[k];
      const lc = k.charAt(0).toLowerCase() + k.slice(1);
      if (obj?.[lc] !== undefined) return obj[lc];
    }
    return undefined;
  };

  // ─── ICS Parser ───────────────────────────────────────────────────────────

  function parseICS(icsString) {
    const events = [];
    const blocks = icsString.match(/BEGIN:VEVENT[\s\S]*?END:VEVENT/g) || [];
    for (const block of blocks) {
      const get = (field) => {
        const m = block.match(new RegExp(`^${field}[^:]*:(.+)$`, 'm'));
        return m ? m[1].trim() : '';
      };
      const dtstart = get('DTSTART');
      const dtend   = get('DTEND');
      if (!dtstart || !dtend) continue;
      const categories = get('CATEGORIES');
      const isOpenShift = categories.includes('Open Shift') || (get('SUMMARY') || '').startsWith('OPEN: ');
      events.push({
        dtstart,
        dtend,
        summary: get('SUMMARY') || 'Shift',
        desc:    get('DESCRIPTION'),
        categories,
        isOpenShift,
      });
    }
    return events;
  }

  function icsDateToISO(d) {
    const m = d.match(/^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})(Z?)$/);
    if (!m) return d;
    return `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}${m[7] ? 'Z' : ''}`;
  }

  // ─── Graph API Import ─────────────────────────────────────────────────────

  async function importViaGraphAPI(icsContent, onProgress) {
    const tokens = findMSALTokens();
    if (tokens.length === 0) throw new Error('No access tokens found in sessionStorage/localStorage');

    let workingToken  = null;
    let apiBase       = null;
    let calendarId    = null;
    let isPascal      = false;
    let calendarIsNew = false; // true when we just created the calendar from scratch

    outer:
    for (const token of tokens) {
      const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
      for (const base of API_BASES) {
        try {
          const resp = await fetch(`${base}/me/calendars`, { headers });
          if (!resp.ok) { console.info(`[ShiftsExport] ${base} → ${resp.status}`); continue; }
          const data = await resp.json();

          const pascal  = base.includes('outlook.office.com');
          const nameKey = pascal ? 'Name' : 'name';
          const idKey   = pascal ? 'Id'   : 'id';

          let cal = (data.value || []).find((c) => c[nameKey] === TARGET_CALENDAR);
          if (!cal) {
            const body = pascal ? { Name: TARGET_CALENDAR } : { name: TARGET_CALENDAR };
            const cr   = await fetch(`${base}/me/calendars`, {
              method: 'POST', headers, body: JSON.stringify(body),
            });
            if (!cr.ok) continue;
            cal = await cr.json();
            calendarIsNew = true;
            console.info(`[ShiftsExport] Created "${TARGET_CALENDAR}" via ${base}`);
          }

          if (cal?.[idKey]) {
            workingToken = token;
            apiBase      = base;
            calendarId   = cal[idKey];
            isPascal     = pascal;
            break outer;
          }
        } catch (e) { console.info(`[ShiftsExport] ${base} error:`, e.message); }
      }
    }

    if (!workingToken || !calendarId) {
      throw new Error('No token had Calendars access across all tried API bases');
    }

    console.info(`[ShiftsExport] Using API base: ${apiBase}`);

    const headers = { Authorization: `Bearer ${workingToken}`, 'Content-Type': 'application/json' };

    const events  = parseICS(icsContent);
    if (events.length === 0) throw new Error('No events parsed from ICS content');

    // ── Deduplication + sync ─────────────────────────────────────────────────
    const minStart = events.reduce((m, e) => (e.dtstart < m ? e.dtstart : m), events[0].dtstart);
    const maxEnd   = events.reduce((m, e) => (e.dtend   > m ? e.dtend   : m), events[0].dtend);
    const existingKeys = new Set();
    const existingOpenShiftKeys = new Set();
    // Track existing regular (non-open) shift event IDs for removal sync
    const existingRegularShifts = []; // { id, key }

    try {
      const subjectKey  = isPascal ? 'Subject' : 'subject';
      const startKey    = isPascal ? 'Start'   : 'start';
      const catKey      = isPascal ? 'Categories' : 'categories';
      const idKey       = isPascal ? 'Id' : 'id';
      const selectParam = isPascal ? 'Id,Subject,Start,Categories' : 'id,subject,start,categories';
      // Request times in the user's local timezone so they match ICS local times
      const userTZ = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
      const viewHeaders = { ...headers, 'Prefer': `outlook.timezone="${userTZ}"` };
      const viewResp = await fetch(
        `${apiBase}/me/calendars/${calendarId}/calendarView` +
        `?startDateTime=${icsDateToISO(minStart)}&endDateTime=${icsDateToISO(maxEnd)}` +
        `&$select=${selectParam}&$top=1000`,
        { headers: viewHeaders }
      );
      if (viewResp.ok) {
        for (const ev of ((await viewResp.json()).value || [])) {
          const dt = (ev[startKey]?.DateTime || ev[startKey]?.dateTime || '').substring(0, 16);
          const key = `${ev[subjectKey]}|${dt}`;
          existingKeys.add(key);
          const cats = ev[catKey] || [];
          if (cats.includes('Open Shift')) {
            existingOpenShiftKeys.add(key);
          } else {
            existingRegularShifts.push({ id: ev[idKey], key });
          }
        }
      }
    } catch {}

    // ── Remove regular shifts no longer in Teams (today forward only) ───
    // Historical shifts (before today) are preserved regardless of Teams state.
    const todayStr = new Date().toISOString().substring(0, 10); // "YYYY-MM-DD"

    const scrapedRegularKeys = new Set();
    for (const ev of events) {
      if (!ev.isOpenShift) {
        scrapedRegularKeys.add(`${ev.summary}|${icsDateToISO(ev.dtstart).substring(0, 16)}`);
      }
    }

    let deleted = 0;
    for (const existing of existingRegularShifts) {
      // Only delete future/current shifts — leave historical data alone
      const eventDate = existing.key.split('|')[1]?.substring(0, 10) || '';
      if (eventDate < todayStr) continue;

      if (!scrapedRegularKeys.has(existing.key)) {
        try {
          const delResp = await fetch(`${apiBase}/me/events/${existing.id}`, {
            method: 'DELETE', headers,
          });
          if (delResp.ok || delResp.status === 204) {
            deleted++;
            console.info(`[ShiftsExport] Removed stale shift: ${existing.key}`);
          }
        } catch {}
      }
    }
    if (deleted > 0) {
      console.info(`[ShiftsExport] Removed ${deleted} stale regular shifts from Outlook`);
    }

    // ── Fresh start: calendar was just created ──────────────────────────
    // Reset all tracking so everything (including historical data) imports cleanly
    if (calendarIsNew) {
      console.info('[ShiftsExport] Calendar is new — resetting open shift tracking for fresh import');
      await storageSet({ dismissedOpenShifts: [], trackedOpenShifts: [] });
    }

    // ── Dismissed open shifts ──────────────────────────────────────────────
    // If user deleted an open shift from the calendar, don't re-import it
    const stored = await storageGet(['dismissedOpenShifts', 'trackedOpenShifts']);
    const dismissedSet = new Set(stored.dismissedOpenShifts || []);
    const previouslyTracked = stored.trackedOpenShifts || [];

    // Open shifts we previously imported that are no longer in calendar = user dismissed them
    for (const key of previouslyTracked) {
      if (!existingOpenShiftKeys.has(key)) {
        dismissedSet.add(key);
        console.info(`[ShiftsExport] Open shift dismissed by user: ${key}`);
      }
    }

    const newEvents = events.filter((ev) => {
      const key = `${ev.summary}|${icsDateToISO(ev.dtstart).substring(0, 16)}`;
      if (existingKeys.has(key)) return false;
      if (ev.isOpenShift && dismissedSet.has(key)) {
        console.info(`[ShiftsExport] Skipping dismissed open shift: ${key}`);
        return false;
      }
      return true;
    });

    console.info(
      `[ShiftsExport] ${events.length} total, ${newEvents.length} new, ${deleted} removed ` +
      `(${events.length - newEvents.length} already exist or dismissed)`
    );
    if (newEvents.length === 0) {
      await storageSet({
        dismissedOpenShifts: [...dismissedSet],
        trackedOpenShifts: [...existingOpenShiftKeys],
      });
      return 0;
    }

    // ── Create events ─────────────────────────────────────────────────────────
    const userTZ = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    let created = 0;
    for (let i = 0; i < newEvents.length; i += 20) {
      const chunk    = newEvents.slice(i, i + 20);
      const requests = chunk.map((ev, idx) => {
        const body = isPascal
          ? {
              Subject: ev.summary,
              Start:   { DateTime: icsDateToISO(ev.dtstart), TimeZone: userTZ },
              End:     { DateTime: icsDateToISO(ev.dtend),   TimeZone: userTZ },
              IsReminderOn: false,
              ...(ev.desc ? { Body: { ContentType: 'Text', Content: ev.desc } } : {}),
              ...(ev.isOpenShift ? { Categories: ['Red category', 'Open Shift'] } : {}),
            }
          : {
              subject: ev.summary,
              start:   { dateTime: icsDateToISO(ev.dtstart), timeZone: userTZ },
              end:     { dateTime: icsDateToISO(ev.dtend),   timeZone: userTZ },
              isReminderOn: false,
              ...(ev.desc ? { body: { contentType: 'text', content: ev.desc } } : {}),
              ...(ev.isOpenShift ? { categories: ['Red category', 'Open Shift'] } : {}),
            };
        return {
          id:      String(i + idx + 1),
          method:  'POST',
          url:     `/me/calendars/${calendarId}/events`,
          headers: { 'Content-Type': 'application/json' },
          body,
        };
      });

      // Outlook REST API doesn't support $batch — POST events individually
      if (isPascal) {
        for (const req of requests) {
          const r = await fetch(`${apiBase}${req.url}`, {
            method: 'POST', headers, body: JSON.stringify(req.body),
          });
          if (r.status === 201) created++;
          if (onProgress) onProgress(created, newEvents.length);
        }
      } else {
        const batchResp = await fetch(`https://graph.microsoft.com/v1.0/$batch`, {
          method: 'POST', headers,
          body: JSON.stringify({ requests }),
        });
        if (!batchResp.ok) throw new Error(`$batch failed: ${batchResp.status}`);
        created += ((await batchResp.json()).responses || []).filter((r) => r.status === 201).length;
        if (onProgress) onProgress(created, newEvents.length);
      }
    }

    // Update tracked open shifts = existing + newly created open shifts
    const updatedTracked = new Set(existingOpenShiftKeys);
    for (const ev of newEvents) {
      if (ev.isOpenShift) {
        updatedTracked.add(`${ev.summary}|${icsDateToISO(ev.dtstart).substring(0, 16)}`);
      }
    }
    await storageSet({
      dismissedOpenShifts: [...dismissedSet],
      trackedOpenShifts: [...updatedTracked],
    });

    console.info(`[ShiftsExport] ${created} events created in "${TARGET_CALENDAR}" via ${apiBase}`);
    return created;
  }

  // ─── Fallback UI Helpers ──────────────────────────────────────────────────

  async function ensureNavigationPane() {
    const showBtn = document.querySelector('button[aria-label="Show navigation pane"]');
    if (showBtn) {
      showBtn.click();
      await sleep(600);
    }
  }

  function findAddCalendarBtn() {
    return (
      document.querySelector('button[title="Add a new calendar"]') ||
      document.querySelector('button[aria-label*="Add calendar" i]') ||
      Array.from(document.querySelectorAll('button')).find((b) =>
        /add\s+(a\s+new\s+)?calendar/i.test(b.textContent + b.title)
      )
    );
  }

  async function ensureWorkShiftsCalendar() {
    await ensureNavigationPane();

    const showAllBtn = Array.from(document.querySelectorAll('button')).find((b) =>
      /show all/i.test(b.title + b.textContent)
    );
    if (showAllBtn) { showAllBtn.click(); await sleep(600); }

    const existing =
      document.querySelector(`[title*="${TARGET_CALENDAR}"]`) ||
      document.querySelector(`[aria-label*="${TARGET_CALENDAR}"]`) ||
      Array.from(document.querySelectorAll('button, [role="option"], [role="treeitem"]')).find(
        (el) =>
          (el.textContent + (el.getAttribute('title') || '')).toLowerCase()
            .includes(TARGET_CALENDAR.toLowerCase())
      );

    if (existing) {
      console.info('[ShiftsExport] "Work Shifts" calendar already exists — skipping creation.');
      return true;
    }

    console.info('[ShiftsExport] Creating "Work Shifts" calendar via UI…');
    const addBtn = findAddCalendarBtn();
    if (!addBtn) return false;

    addBtn.click();
    await sleep(1000);

    const createTab =
      document.querySelector('[role="tab"][title="Create blank calendar"]') ||
      Array.from(document.querySelectorAll('[role="tab"], button')).find((el) =>
        /create.*blank|blank.*calendar/i.test(el.textContent + el.title)
      );
    if (createTab) { createTab.click(); await sleep(600); }

    const nameInput =
      document.querySelector('input[placeholder*="calendar name" i]') ||
      document.querySelector('input[placeholder*="name" i]');
    if (!nameInput) {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      return false;
    }

    const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    nameInput.focus();
    nativeSetter.call(nameInput, TARGET_CALENDAR);
    nameInput.dispatchEvent(new Event('input',  { bubbles: true }));
    nameInput.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(400);

    let saveBtn = null;
    for (let i = 0; i < 10; i++) {
      saveBtn = Array.from(document.querySelectorAll('button')).find(
        (b) => /^save$/i.test(b.textContent.trim()) && !b.disabled
      );
      if (saveBtn) break;
      await sleep(300);
    }
    saveBtn = saveBtn || Array.from(document.querySelectorAll('button')).find(
      (b) => /^save$/i.test(b.textContent.trim())
    );
    if (saveBtn) { saveBtn.click(); await sleep(1200); return true; }

    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    return false;
  }

  async function injectFile(fileInput, icsContent) {
    const icsFile = new File([icsContent], 'shifts.ics', { type: 'text/calendar' });
    const dt = new DataTransfer();
    dt.items.add(icsFile);
    Object.defineProperty(fileInput, 'files', { configurable: true, get: () => dt.files });
    fileInput.dispatchEvent(new Event('change', { bubbles: true }));
    fileInput.dispatchEvent(new InputEvent('input', { bubbles: true }));
    await sleep(500);
  }

  async function tryAddCalendarDialog(icsContent) {
    await ensureNavigationPane();
    const addBtn = findAddCalendarBtn();
    if (!addBtn) throw new Error('"Add calendar" button not found');
    addBtn.click();
    await sleep(1200);

    const uploadTab =
      Array.from(document.querySelectorAll('[role="tab"], button')).find((el) =>
        /^upload from file$/i.test(el.textContent.trim())
      ) ||
      document.querySelector('[role="tab"][title="Upload from file"]');
    if (!uploadTab) throw new Error('"Upload from file" tab not found');

    uploadTab.click();
    await sleep(1000);

    const fileInput = document.querySelector('input[type="file"]');
    if (!fileInput) throw new Error('File input not found');

    await injectFile(fileInput, icsContent);

    let submitBtn = null;
    for (let i = 0; i < 10; i++) {
      submitBtn = Array.from(document.querySelectorAll('button')).find(
        (b) => /^import$/i.test(b.textContent.trim()) && !b.disabled
      );
      if (submitBtn) break;
      await sleep(300);
    }
    submitBtn = submitBtn || Array.from(document.querySelectorAll('button')).find(
      (b) => /^import$/i.test(b.textContent.trim())
    );
    if (!submitBtn) throw new Error('"Import" button not found');

    submitBtn.click();
    await sleep(1500);
    return true;
  }

  // ─── Main Import Orchestrator ─────────────────────────────────────────────

  async function importICS(icsContent, overlay) {
    // Primary strategy: Microsoft Graph API — direct event creation, no UI needed
    try {
      console.info('[ShiftsExport] Trying: Graph API');
      const progressCb = overlay
        ? (created, total) => overlay.update(`Creating event ${created} of ${total}...`)
        : undefined;
      const count = await importViaGraphAPI(icsContent, progressCb);
      console.info(`[ShiftsExport] Import succeeded via Graph API (${count} new events)`);
      return { success: true, method: 'Graph API', count };
    } catch (err) {
      console.warn('[ShiftsExport] Graph API failed:', err.message, '— falling back to UI automation');
    }

    // Fallback: UI automation
    if (overlay) overlay.update('Importing shifts via UI...');
    await ensureWorkShiftsCalendar();
    try {
      console.info('[ShiftsExport] Trying: Add Calendar dialog');
      await tryAddCalendarDialog(icsContent);
      console.info('[ShiftsExport] Import succeeded via: Add Calendar dialog');
      return { success: true, method: 'Add Calendar dialog' };
    } catch (err) {
      console.warn('[ShiftsExport] Add Calendar dialog failed:', err.message);
    }

    throw new Error(
      'All import strategies failed. Open Outlook Calendar, go to ' +
      'Add calendar → Upload from file and import the .ics from your Downloads.'
    );
  }

  // ─── Clear all events from a calendar via Graph API ──────────────────────

  async function clearCalendarEvents() {
    const tokens = findMSALTokens();
    if (tokens.length === 0) throw new Error('No access tokens found');

    for (const token of tokens) {
      const headers = { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' };
      for (const base of API_BASES) {
        try {
          const resp = await fetch(`${base}/me/calendars`, { headers });
          if (!resp.ok) continue;
          const data = await resp.json();

          const pascal  = base.includes('outlook.office.com');
          const nameKey = pascal ? 'Name' : 'name';
          const idKey   = pascal ? 'Id'   : 'id';

          const cal = (data.value || []).find((c) => c[nameKey] === TARGET_CALENDAR);
          if (!cal?.[idKey]) {
            console.info(`[ShiftsExport] "${TARGET_CALENDAR}" not found — nothing to clear.`);
            return 0;
          }

          const calId = cal[idKey];

          // Fetch ALL events in the calendar (paginate up to 500)
          const selectParam = pascal ? 'Id' : 'id';
          let eventsUrl = `${base}/me/calendars/${calId}/events?$select=${selectParam}&$top=500`;
          let deleted = 0;

          while (eventsUrl) {
            const evResp = await fetch(eventsUrl, { headers });
            if (!evResp.ok) break;
            const evData = await evResp.json();
            const items  = evData.value || [];

            for (const ev of items) {
              const evId = ev[pascal ? 'Id' : 'id'];
              if (!evId) continue;
              const delResp = await fetch(`${base}/me/events/${evId}`, {
                method: 'DELETE', headers,
              });
              if (delResp.ok || delResp.status === 204) deleted++;
            }

            eventsUrl = evData['@odata.nextLink'] || evData['@odata.nextlink'] || null;
          }

          console.info(`[ShiftsExport] Cleared ${deleted} events from "${TARGET_CALENDAR}"`);
          return deleted;
        } catch (e) {
          console.info(`[ShiftsExport] Clear failed on ${base}:`, e.message);
        }
      }
    }
    throw new Error('Could not clear calendar — no working API token');
  }

  // ─── Message Listener ─────────────────────────────────────────────────────

  browser.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.action === 'IMPORT_ICS_TO_OUTLOOK') {
      const overlay = showOverlay('Importing shifts to Outlook...');
      (async () => {
        try {
          const result = await importICS(msg.icsContent, overlay);
          overlay.update('Import complete! Reloading calendar...');
          try { sendResponse({ success: true, ...result }); } catch {}
          await storageSet({ focusWorkShifts: true });
          await sleep(1500);
          window.location.reload();
        } catch (err) {
          try { sendResponse({ success: false, error: err.message }); } catch {}
          overlay.remove();
        }
      })();
      return true;
    }

    if (msg.action === 'CLEAR_AND_IMPORT_ICS') {
      const overlay = showOverlay('Clearing existing events...');
      (async () => {
        try {
          // Reset dismissed open shifts on full re-import
          await storageSet({ dismissedOpenShifts: [], trackedOpenShifts: [] });
          const deleted = await clearCalendarEvents();
          console.info(`[ShiftsExport] Cleared ${deleted} old events, now importing fresh...`);
          overlay.update('Importing shifts to Outlook...');
          const result = await importICS(msg.icsContent, overlay);
          overlay.update('Import complete! Reloading calendar...');
          try { sendResponse({ success: true, deleted, ...result }); } catch {}
          await storageSet({ focusWorkShifts: true });
          await sleep(1500);
          window.location.reload();
        } catch (err) {
          try { sendResponse({ success: false, error: err.message }); } catch {}
          overlay.remove();
        }
      })();
      return true;
    }
  });
})();
