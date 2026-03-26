// content.js — injected into teams.microsoft.com
// Exposes window.__shiftsExport.scrape() for use by the background service worker

(function () {
  if (window.__shiftsExportInitialized) return;
  window.__shiftsExportInitialized = true;

  // ─── Utilities ────────────────────────────────────────────────────────────

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  // Calculate the end date of the nearest future August or February.
  // If we are within 2 weeks of that boundary, return the NEXT boundary so
  // newly-posted shifts for the upcoming period are also captured.
  function getTargetEndDate() {
    const now = new Date();
    const twoWeeksFromNow = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
    const year = now.getFullYear();

    // Candidates: end of Feb/Aug for current year, next year, and year after
    // (extra candidate needed when we skip one boundary due to the 2-week rule)
    const candidates = [
      new Date(year,     1, 28), // Feb this year
      new Date(year,     7, 31), // Aug this year
      new Date(year + 1, 1, 28), // Feb next year
      new Date(year + 1, 7, 31), // Aug next year
      new Date(year + 2, 1, 28), // Feb year after (fallback)
    ];

    // Adjust Feb dates for leap years
    candidates.forEach((d, i) => {
      if (d.getMonth() === 1) {
        const ly = d.getFullYear();
        if ((ly % 4 === 0 && ly % 100 !== 0) || ly % 400 === 0) {
          candidates[i] = new Date(ly, 1, 29);
        }
      }
    });

    // Sort future candidates ascending
    const future = candidates.filter((d) => d > now).sort((a, b) => a - b);

    // If the nearest boundary is within 2 weeks, jump to the one after it
    if (future.length > 1 && future[0] <= twoWeeksFromNow) {
      return future[1];
    }

    return future[0];
  }

  function weeksUntil(targetDate) {
    const now = new Date();
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    return Math.ceil((targetDate - now) / msPerWeek);
  }

  // Parse a time string like "9:00 AM" into { hour, minute } (24h)
  function parseTime(str) {
    const m = str.trim().match(/(\d{1,2})(?::(\d{2}))?\s*([AP]M)/i);
    if (!m) return null;
    let hour = parseInt(m[1], 10);
    const minute = m[2] ? parseInt(m[2], 10) : 0;
    const ampm = m[3].toUpperCase();
    if (ampm === 'AM' && hour === 12) hour = 0;
    if (ampm === 'PM' && hour !== 12) hour += 12;
    return { hour, minute };
  }

  // Build a Date object from a date string (e.g. "Mon 3/18" or "March 18") and time parts
  function buildDate(dateStr, timeStr, referenceYear) {
    const timeParts = parseTime(timeStr);
    if (!timeParts) return null;

    let month, day;

    // Format: "3/18" or "Mon 3/18"
    let m = dateStr.match(/(\d{1,2})\/(\d{1,2})/);
    if (m) {
      month = parseInt(m[1], 10) - 1;
      day = parseInt(m[2], 10);
    } else {
      // Format: "March 18" or "Mar 18"
      const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
      m = dateStr.match(/([A-Za-z]+)\s+(\d{1,2})/);
      if (m) {
        month = months.indexOf(m[1].toLowerCase().slice(0, 3));
        day = parseInt(m[2], 10);
      }
    }

    if (month === undefined || day === undefined || month < 0) return null;

    const d = new Date(referenceYear, month, day, timeParts.hour, timeParts.minute, 0);

    // Handle year boundary: if the month is behind the current month by more than 6, bump year
    const now = new Date();
    if (d < now) {
      const diffMs = now - d;
      const diffMonths = diffMs / (1000 * 60 * 60 * 24 * 30);
      if (diffMonths > 6) d.setFullYear(d.getFullYear() + 1);
    }

    return d;
  }

  // Format a Date to ICS DTSTART/DTEND format (local time with TZID approach, or UTC)
  function toICSDate(date) {
    const pad = (n) => String(n).padStart(2, '0');
    return (
      date.getFullYear().toString() +
      pad(date.getMonth() + 1) +
      pad(date.getDate()) +
      'T' +
      pad(date.getHours()) +
      pad(date.getMinutes()) +
      '00'
    );
  }

  // ─── UI Overlay ──────────────────────────────────────────────────────────

  function showOverlay(message) {
    // Scraping runs in a minimized background window, so this overlay is not
    // visible to the user — it just logs progress to the console.
    console.info('[ShiftsExport]', message);
    return {
      update(msg) { console.info('[ShiftsExport]', msg); },
      remove() {},
    };
  }

  // ─── Week Change Detection ──────────────────────────────────────────────

  function getCurrentWeekLabel() {
    // Try the date picker button first: aria-label like "August 3 - August 9, 2026. Pick a date"
    const datePickerBtn = document.querySelector('button[aria-label*="Pick a date"]');
    if (datePickerBtn) {
      return datePickerBtn.getAttribute('aria-label') || '';
    }
    // Fallback: first day column header with aria-label like "Monday, August 3, 2026"
    const headers = Array.from(
      document.querySelectorAll('div[aria-label]')
    ).filter((el) => /\w+day,\s+\w+\s+\d+,\s+\d{4}/.test(el.getAttribute('aria-label') || ''));
    if (headers.length > 0) return headers[0].getAttribute('aria-label');
    return '';
  }

  function waitForWeekChange(previousDateLabel, maxWaitMs = 5000) {
    return new Promise((resolve) => {
      const interval = 300;
      let elapsed = 0;
      const timer = setInterval(() => {
        elapsed += interval;
        const current = getCurrentWeekLabel();
        if (current && current !== previousDateLabel) {
          clearInterval(timer);
          resolve(true);
        } else if (elapsed >= maxWaitMs) {
          clearInterval(timer);
          resolve(false); // timed out
        }
      }, interval);
    });
  }

  // Wait until the number of shift cards on the page stops changing.
  // This is more reliable than a fixed delay — some weeks render faster,
  // some slower, and weeks with no shifts would still wait unnecessarily.
  async function waitForShiftsStable(maxWaitMs = 6000) {
    const pollMs = 300;
    const stableThresholdMs = 1200; // card count must be unchanged for this long
    // Brief initial pause so Teams has time to begin fetching data for the new
    // week before we start checking card count. Without this, a 0-card state
    // during the network request looks like a stable empty week.
    await sleep(500);
    let prevCount = -1;
    let stableFor = 0;
    const deadline = Date.now() + maxWaitMs;

    while (Date.now() < deadline) {
      const count =
        document.querySelectorAll('div[aria-label^="Shift."], div[aria-label^="Open shift"]').length;

      if (count === prevCount) {
        stableFor += pollMs;
        if (stableFor >= stableThresholdMs) return;
      } else {
        stableFor = 0;
        prevCount = count;
      }
      await sleep(pollMs);
    }
    // Timed out — proceed anyway with whatever cards are present
  }

  // Scroll the shifts grid container to the bottom and back to the top so that
  // virtualised rows outside the viewport get rendered before we scrape.
  async function scrollShiftsIntoView() {
    const container =
      document.querySelector('[data-tid="shifts-grid"]') ||
      document.querySelector('[class*="scheduleGrid"]') ||
      document.querySelector('[role="grid"]') ||
      document.querySelector('[class*="shiftsBody"]') ||
      // Last resort: find the deepest scrollable element inside the shifts iframe
      Array.from(document.querySelectorAll('*')).find(
        (el) => el.scrollHeight > el.clientHeight + 10 && getComputedStyle(el).overflowY !== 'visible'
      );

    if (!container) return;

    const original = container.scrollTop;
    container.scrollTop = container.scrollHeight;
    await sleep(200);
    container.scrollTop = 0;
    await sleep(200);
    // Restore original position so the UI looks unchanged
    container.scrollTop = original;
    await sleep(100);
  }

  // ─── DOM Scraping ─────────────────────────────────────────────────────────

  // Find the "next week" navigation button
  // Confirmed selector from live DOM inspection
  function findNextButton() {
    return (
      document.querySelector('button[aria-label="Go to next week"]') ||
      document.querySelector('button[title="Go to next week"]')
    );
  }

  // Get the year currently displayed from the date column headers.
  // Headers have aria-label like "Monday, August 3, 2026"
  function getDisplayedYear() {
    const headers = Array.from(
      document.querySelectorAll('div[aria-label*="_fcContentWrapper"], div[aria-label]')
    ).filter((el) => /\w+day,\s+\w+\s+\d+,\s+\d{4}/.test(el.getAttribute('aria-label') || ''));

    if (headers.length > 0) {
      const m = headers[0].getAttribute('aria-label').match(/\d{4}/);
      if (m) return parseInt(m[0]);
    }
    return new Date().getFullYear();
  }

  // Scrape shift cards from the current week view.
  // Shift aria-label format (confirmed from live DOM):
  // "Shift. thursday, Aug 6, 12:45 PM - 10:45 PM D6P 1245. themeDarkBlue. . Press Enter"
  function scrapeCurrentWeek(userName) {
    const shifts = [];
    const year = getDisplayedYear();

    // Find all shift cards — confirmed selector from live DOM
    let shiftCards = Array.from(document.querySelectorAll('div[aria-label^="Shift."]'));

    // If userName set, try to scope to that member's row
    if (userName) {
      const memberCell = Array.from(document.querySelectorAll('div[aria-label]')).find((el) =>
        el.getAttribute('aria-label').toLowerCase().includes(`member name: ${userName.toLowerCase()}`)
      );
      if (memberCell) {
        // Walk up to find the row container, then search within it
        let row = memberCell.parentElement;
        for (let i = 0; i < 4; i++) {
          const rowShifts = row ? row.querySelectorAll('div[aria-label^="Shift."]') : [];
          if (rowShifts.length > 0) {
            shiftCards = Array.from(rowShifts);
            break;
          }
          row = row?.parentElement;
        }
      }
    }

    for (const card of shiftCards) {
      const ariaLabel = card.getAttribute('aria-label') || '';

        // Full format:
      // "Shift. {weekday}, {Month} {day}, {start} - {end} {name}. {theme}. {notes}. Press Enter key..."
      // Notes segment is empty ("  .") when there are no notes.
      const match = ariaLabel.match(
        /Shift\.\s+\w+,\s+(\w+\s+\d+),\s+(\d{1,2}(?::\d{2})?\s*[AP]M)\s*-\s*(?:\w+\s+\d+,\s+)?(\d{1,2}(?::\d{2})?\s*[AP]M)\s+(.*?)\.\s*\w+\.\s*(.*?)\.\s*Press Enter/i
      );
      if (!match) continue;

      const dateStr  = match[1].trim(); // "Aug 6"
      const startStr = match[2].trim(); // "12:45 PM"
      const endStr   = match[3].trim(); // "10:45 PM"
      const summary  = match[4].trim() || 'Shift'; // "D6P 1245"
      const notes    = match[5].trim(); // "Picked up from Steve Z" or ""

      shifts.push({ summary, notes, dateStr, startStr, endStr, referenceYear: year, isOpenShift: false });
    }

    // Also scrape open shift cards
    const openShiftCards = Array.from(document.querySelectorAll('div[aria-label^="Open shift"]'));
    for (const card of openShiftCards) {
      const ariaLabel = card.getAttribute('aria-label') || '';

      // Open shift format (similar to regular shifts):
      // "Open shift. {weekday}, {Month} {day}, {start} - {end} {name}. {theme}. {notes}. Press Enter key..."
      const match = ariaLabel.match(
        /Open\s+shift\.\s+\w+,\s+(\w+\s+\d+),\s+(\d{1,2}(?::\d{2})?\s*[AP]M)\s*-\s*(?:\w+\s+\d+,\s+)?(\d{1,2}(?::\d{2})?\s*[AP]M)\s+(.*?)\.\s*\w+\.\s*(.*?)\.\s*Press Enter/i
      );
      if (!match) continue;

      const dateStr  = match[1].trim();
      const startStr = match[2].trim();
      const endStr   = match[3].trim();
      const summary  = match[4].trim() || 'Open Shift';
      const notes    = match[5].trim();

      shifts.push({ summary, notes, dateStr, startStr, endStr, referenceYear: year, isOpenShift: true });
    }

    return shifts;
  }

  // ─── Teams Permission Dialog ──────────────────────────────────────────────
  // Teams sometimes shows an "Almost there!" consent dialog when first opening
  // Shifts. Detect it and click Continue so the scrape can proceed.

  async function dismissTeamsPermissionDialog(maxWaitMs = 8000) {
    const deadline = Date.now() + maxWaitMs;
    while (Date.now() < deadline) {
      const hasDialog = Array.from(document.querySelectorAll('*')).some(
        (el) => el.childElementCount === 0 && /almost there/i.test(el.textContent)
      );
      if (hasDialog) {
        const continueBtn = Array.from(document.querySelectorAll('button')).find(
          (btn) => /^continue$/i.test(btn.textContent.trim())
        );
        if (continueBtn) {
          continueBtn.click();
          await sleep(2500); // wait long enough for Teams to begin the post-auth reload
          return true;
        }
      }
      await sleep(400);
    }
    return false;
  }

  // ─── Navigate to Shifts ───────────────────────────────────────────────────
  // New Teams (teams.cloud.microsoft) doesn't have a direct URL for Shifts.
  // We click: left sidebar "..." (more apps) → Shifts.

  async function navigateToShifts() {
    // If Shifts is already visible (e.g. pinned in sidebar), just click it directly
    const pinnedShifts =
      document.querySelector('[aria-label*="Shifts" i][role="button"]') ||
      Array.from(document.querySelectorAll('[role="button"], [role="tab"], a')).find(
        (el) => el.textContent.trim() === 'Shifts' || el.getAttribute('aria-label') === 'Shifts'
      );

    if (pinnedShifts) {
      pinnedShifts.click();
      await sleep(2000);
      await dismissTeamsPermissionDialog();
      return true;
    }

    // Otherwise open the "more apps" / three-dots menu in the far-left sidebar.
    // Be specific — Teams has many "more" buttons (e.g. "More options" near chat).
    const moreBtn =
      document.querySelector('[aria-label="More apps"]') ||
      document.querySelector('[aria-label="More added apps"]') ||
      document.querySelector('[aria-label*="more apps" i]') ||
      // Fallback: a "..." button that is a direct child of a nav/sidebar element
      Array.from(document.querySelectorAll('nav [role="button"], [data-tid*="sidebar"] [role="button"], [data-tid*="rail"] [role="button"]')).find((el) =>
        /^(…|\.\.\.)$/.test(el.textContent.trim())
      );

    if (!moreBtn) {
      console.warn('[ShiftsExport] Could not find "more apps" button — ensure you are on Teams');
      return false;
    }

    moreBtn.click();
    await sleep(800);

    // Find Shifts in the flyout menu
    const shiftsItem = Array.from(document.querySelectorAll('[role="menuitem"], [role="option"], [role="button"], li')).find(
      (el) => /^shifts$/i.test(el.textContent.trim()) || /^shifts$/i.test(el.getAttribute('aria-label') || '')
    );

    if (!shiftsItem) {
      console.warn('[ShiftsExport] Shifts not found in more-apps menu');
      // Dismiss the menu
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      return false;
    }

    shiftsItem.click();
    await sleep(2500); // wait for Shifts to load
    await dismissTeamsPermissionDialog();
    return true;
  }

  // ─── "My Schedule" Navigation ─────────────────────────────────────────────

  // Try to click the "Your shifts" (or legacy "My schedule") tab so we only
  // see the current user's shifts.  Returns true if successfully switched.
  async function switchToMySchedule() {
    const selectors = [
      // New Teams Shifts uses "Your shifts"
      'button[aria-label*="Your shifts" i]',
      '[data-tid="your-shifts-tab"]',
      '[data-tid="yourShifts-tab"]',
      // Legacy "My schedule" selectors
      '[data-tid="my-schedule-tab"]',
      '[data-tid="mySchedule-tab"]',
      '[data-tid="shifts-my-schedule"]',
      'button[aria-label*="my schedule" i]',
      'button[role="tab"][aria-label*="my" i]',
      '[role="tab"]',
    ];

    for (const sel of selectors) {
      const els = Array.from(document.querySelectorAll(sel));
      const match = els.find((el) => {
        const text = el.textContent || '';
        const aria = el.getAttribute('aria-label') || '';
        return /your\s+shifts/i.test(text) || /your\s+shifts/i.test(aria) ||
               /my\s+schedule/i.test(text) || /my\s+schedule/i.test(aria);
      });
      if (match) {
        match.click();
        await sleep(1500);
        return true;
      }
    }
    return false;
  }

  // ─── Ensure Weekly View ──────────────────────────────────────────────────
  // The scraper relies on week-by-week navigation.  If Shifts is in month or
  // day view the "Go to next week" button won't exist and shifts outside the
  // visible viewport will be missed.  Force weekly view before scraping.

  async function ensureWeeklyView() {
    const calTypeBtn = document.querySelector('button[aria-label="Calendar type"]');
    if (!calTypeBtn) return; // button not found — likely already week view or different UI

    // If the button already says "Week" we're good
    if (/week/i.test(calTypeBtn.title || calTypeBtn.textContent)) return;

    console.info('[ShiftsExport] Switching to weekly view…');
    calTypeBtn.click();
    await sleep(600);

    // Find "Week" option in the dropdown menu
    const weekOption = Array.from(
      document.querySelectorAll('[role="option"], [role="menuitem"], [role="menuitemradio"], button')
    ).find((el) => /^week$/i.test(el.textContent.trim()));

    if (weekOption) {
      weekOption.click();
      await sleep(1000);
    } else {
      // Dismiss the menu if we couldn't find the option
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      console.warn('[ShiftsExport] Could not find "Week" option in calendar type menu');
    }
  }

  // ─── Name-based Filter ────────────────────────────────────────────────────

  // Filter raw shifts to only those belonging to `userName`.
  // Looks for the name in the card's text content or aria-label.
  function filterByUser(cards, userName) {
    if (!userName) return cards;
    const name = userName.trim().toLowerCase();
    return cards.filter((card) => {
      const text = (card.textContent || '').toLowerCase();
      const aria = (card.getAttribute('aria-label') || '').toLowerCase();
      return text.includes(name) || aria.includes(name);
    });
  }

  // ─── Main Scrape Function ─────────────────────────────────────────────────

  async function scrape(options = {}) {
    const { userName } = options; // optional name filter fallback
    const targetDate = getTargetEndDate();
    const totalWeeks = weeksUntil(targetDate);
    const allRawShifts = [];

    const overlay = showOverlay('Scraping shifts...');
    try {
      // Step 1: navigate to Shifts (only needed in top frame; skip inside iframe)
      if (window === window.top) await navigateToShifts();

      // Step 2: try to switch to "Your shifts" (or legacy "My schedule") view
      const switched = await switchToMySchedule();
      if (!switched && userName) {
        console.info('[ShiftsExport] Could not find "Your shifts" tab — will filter by name:', userName);
      }

      // Step 2b: ensure we're in weekly view (scraper relies on week-by-week navigation)
      await ensureWeeklyView();

      // Step 3: navigate to "today" to reset position
      const todayBtn = document.querySelector('button[title="Today"]');
      if (todayBtn) {
        todayBtn.click();
        await sleep(1500);
      }

      for (let week = 0; week < totalWeeks; week++) {
        overlay.update(`Scraping week ${week + 1} of ${totalWeeks}...`);

        // Send per-week progress to background (popup polls this)
        chrome.runtime.sendMessage({
          action: 'SYNC_PROGRESS',
          step: `Scraping week ${week + 1} of ${totalWeeks}…`,
          percent: 18 + Math.round((week / totalWeeks) * 50),
        }).catch(() => {});

        // Bail out if the user cancelled
        const { syncCancelled } = await chrome.storage.local.get('syncCancelled');
        if (syncCancelled) throw new Error('Sync cancelled');

        // Scroll the grid so virtualised rows outside the viewport get rendered
        await scrollShiftsIntoView();

        const weekShifts = scrapeCurrentWeek(userName);
        allRawShifts.push(...weekShifts);

        const previousLabel = getCurrentWeekLabel();

        const nextBtn = findNextButton();
        if (!nextBtn) {
          console.warn('[ShiftsExport] Could not find "next" button — stopping at week', week);
          break;
        }
        nextBtn.click();

        let changed = await waitForWeekChange(previousLabel);
        if (!changed) {
          // Retry the click once
          console.warn('[ShiftsExport] Week change not detected — retrying click');
          nextBtn.click();
          changed = await waitForWeekChange(previousLabel);
          if (!changed) {
            console.warn('[ShiftsExport] Week change still not detected after retry — continuing');
            await sleep(2000); // final fallback delay
          }
        }

        // Wait for shift cards to finish rendering before scraping
        await waitForShiftsStable();
      }

      // Build ICS events
      const events = [];
      const referenceYear = new Date().getFullYear();

      for (const raw of allRawShifts) {
        const start = buildDate(raw.dateStr, raw.startStr, referenceYear);
        const end = buildDate(raw.dateStr, raw.endStr, referenceYear);
        if (!start || !end) continue;

        // If end is before start (midnight crossing), add one day to end
        if (end <= start) end.setDate(end.getDate() + 1);

        events.push({ summary: raw.summary, notes: raw.notes || '', start, end, isOpenShift: !!raw.isOpenShift });
      }

      return events;
    } finally {
      overlay.remove();
    }
  }

  // ─── ICS Generation ──────────────────────────────────────────────────────

  function generateICS(events) {
    const lines = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Teams Shifts Export//EN',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      'X-WR-CALNAME:Teams Shifts',
    ];

    events.forEach((ev, i) => {
      const uid = `teams-shift-${ev.start.getTime()}-${i}@shifts-export`;
      const summaryText = ev.isOpenShift ? `OPEN: ${ev.summary}` : ev.summary;
      lines.push('BEGIN:VEVENT');
      lines.push(`UID:${uid}`);
      lines.push(`DTSTAMP:${toICSDate(new Date())}`);
      lines.push(`DTSTART:${toICSDate(ev.start)}`);
      lines.push(`DTEND:${toICSDate(ev.end)}`);
      lines.push(`SUMMARY:${summaryText.replace(/,/g, '\\,').replace(/\n/g, '\\n')}`);
      if (ev.isOpenShift) {
        lines.push('CATEGORIES:Open Shift');
      }
      lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
  }

  // ─── Expose API ───────────────────────────────────────────────────────────

  window.__shiftsExport = {
    scrape,
    generateICS,
    getTargetEndDate,
  };

  // Listen for messages from the background service worker
  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.action === 'NAVIGATE_TO_SHIFTS') {
      navigateToShifts()
        .then(() => sendResponse({ success: true }))
        .catch((err) => sendResponse({ success: false, error: err.message }));
      return true;
    }

    if (msg.action === 'SCRAPE_AND_EXPORT') {
      scrape({ userName: msg.userName || null })
        .then((events) => {
          const ics = generateICS(events);
          // Also return raw serialisable events so the background can merge with history
          const serializable = events.map((e) => ({
            summary: e.summary,
            notes:   e.notes || '',
            startMs: e.start.getTime(),
            endMs:   e.end.getTime(),
            isOpenShift: !!e.isOpenShift,
          }));
          sendResponse({ success: true, ics, count: events.length, events: serializable });
        })
        .catch((err) => {
          sendResponse({ success: false, error: err.message });
        });
      return true; // keep channel open for async response
    }
  });
})();
