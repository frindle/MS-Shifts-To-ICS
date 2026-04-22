// content.js — injected into teams.microsoft.com / flw.teams.cloud.microsoft

(function () {
  if (window.__shiftsExportInitialized) return;
  window.__shiftsExportInitialized = true;

  // Store API base URL directly when running inside the Shifts iframe
  if (window.location.hostname.includes('flw.teams.cloud.microsoft')) {
    const region = window.location.pathname.split('/').filter(Boolean)[0];
    if (region) {
      chrome.storage.local.set({ shiftsApiBase: `${window.location.origin}/${region}/api` }).catch(() => {});
    }
  }

  // ─── Utilities ────────────────────────────────────────────────────────────

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function getTargetEndDate() {
    const now = new Date();
    const twoWeeksFromNow = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
    const year = now.getFullYear();

    const candidates = [
      new Date(year,     1, 28),
      new Date(year,     7, 31),
      new Date(year + 1, 1, 28),
      new Date(year + 1, 7, 31),
      new Date(year + 2, 1, 28),
    ];

    candidates.forEach((d, i) => {
      if (d.getMonth() === 1) {
        const ly = d.getFullYear();
        if ((ly % 4 === 0 && ly % 100 !== 0) || ly % 400 === 0) {
          candidates[i] = new Date(ly, 1, 29);
        }
      }
    });

    const future = candidates.filter((d) => d > now).sort((a, b) => a - b);
    if (future.length > 1 && future[0] <= twoWeeksFromNow) return future[1];
    return future[0];
  }

  function toICSDate(date) {
    const pad = (n) => String(n).padStart(2, '0');
    return (
      date.getFullYear() +
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
    const old = document.getElementById('shifts-export-overlay');
    if (old) old.remove();

    const overlay = document.createElement('div');
    overlay.id = 'shifts-export-overlay';
    Object.assign(overlay.style, {
      position: 'fixed', top: '0', left: '0',
      width: '100vw', height: '100vh', zIndex: '999999',
      background: 'rgba(0,0,0,0.7)', color: '#fff',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      fontSize: '28px', fontFamily: 'sans-serif', pointerEvents: 'auto',
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

  // ─── Teams Navigation ─────────────────────────────────────────────────────

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
          await sleep(2500);
          return true;
        }
      }
      await sleep(400);
    }
    return false;
  }

  async function navigateToShifts() {
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

    const moreBtn =
      document.querySelector('[aria-label="More apps"]') ||
      document.querySelector('[aria-label="More added apps"]') ||
      document.querySelector('[aria-label*="more apps" i]') ||
      Array.from(document.querySelectorAll('nav [role="button"], [data-tid*="sidebar"] [role="button"], [data-tid*="rail"] [role="button"]')).find((el) =>
        /^(…|\.\.\.)$/.test(el.textContent.trim())
      );

    if (!moreBtn) throw new Error('Could not find Teams sidebar');

    moreBtn.click();

    const flyoutDeadline = Date.now() + 2000;
    let shiftsItem = null;
    while (Date.now() < flyoutDeadline) {
      shiftsItem = Array.from(document.querySelectorAll('[role="menuitem"], [role="option"], [role="button"], li')).find((el) => {
        const text = el.textContent.trim();
        const aria = el.getAttribute('aria-label') || '';
        return /\bshifts\b/i.test(aria) || /^shifts$/i.test(text) || /^shifts\b/i.test(text);
      });
      if (shiftsItem) break;
      await sleep(200);
    }

    if (!shiftsItem) {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
      throw new Error('Shifts not found in Teams menu. Please pin Shifts to your sidebar.');
    }

    shiftsItem.click();
    await sleep(2500);
    await dismissTeamsPermissionDialog();
    return true;
  }

  // ─── API-based Scrape ─────────────────────────────────────────────────────

  async function scrape(options = {}) {
    const overlay = showOverlay('Fetching shifts...');
    try {
      if (window === window.top) await navigateToShifts();

      const region = window.location.pathname.split('/').filter(Boolean)[0];
      const apiBase = `${window.location.origin}/${region}/api`;
      console.info('[ShiftsExport] API base:', apiBase);

      // Use page's fetch so Teams' auth tokens (Bearer etc.) are included
      const apiFetch = typeof unsafeWindow !== 'undefined' && unsafeWindow.fetch
        ? unsafeWindow.fetch.bind(unsafeWindow)
        : fetch;

      overlay.update('Getting teams...');
      const teamsResp = await apiFetch(`${apiBase}/users/me/teams`);
      const teamsText = await teamsResp.text();
      if (!teamsResp.ok) throw new Error(`Teams API error: ${teamsResp.status}: ${teamsText.slice(0, 120)}`);
      const teamsData = JSON.parse(teamsText);
      console.info('[ShiftsExport] Teams raw:', JSON.stringify(teamsData).slice(0, 500));

      const teamList = teamsData.teams || teamsData.value || (Array.isArray(teamsData) ? teamsData : []);
      const teamIds = teamList.map((t) => t.id || t.teamId).filter(Boolean);
      if (!teamIds.length) throw new Error('No teams found');

      console.info('[ShiftsExport] Team IDs:', teamIds);

      overlay.update('Fetching shifts...');
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

      const url = `${apiBase}/users/me/dataindaterange?${params}`;
      console.info('[ShiftsExport] Fetching:', url);

      const shiftsResp = await apiFetch(url);
      const shiftsText = await shiftsResp.text();
      if (!shiftsResp.ok) throw new Error(`Shifts API error: ${shiftsResp.status}: ${shiftsText.slice(0, 120)}`);
      const data = JSON.parse(shiftsText);

      console.info('[ShiftsExport] Response keys:', Object.keys(data));
      console.info('[ShiftsExport] Raw sample:', JSON.stringify(data).slice(0, 3000));

      const events = [];

      const parseShiftItem = (item) => {
        const start = new Date(item.startDateTime || item.StartDateTime || item.start);
        const end   = new Date(item.endDateTime   || item.EndDateTime   || item.end);
        if (isNaN(start) || isNaN(end)) return null;
        const summary = item.displayName || item.DisplayName || item.theme || item.Theme || 'Shift';
        const notes   = item.notes || item.Notes || '';
        return { start, end, summary, notes };
      };

      for (const shift of (data.shifts || data.Shifts || [])) {
        const item = shift.sharedShift || shift.shiftItem || shift.draftShift || shift;
        const parsed = parseShiftItem(item);
        if (parsed) events.push({ ...parsed, isOpenShift: false, isAllDay: false });
      }

      for (const shift of (data.openShifts || data.OpenShifts || [])) {
        const item = shift.sharedOpenShift || shift.openShiftItem || shift.draftOpenShift || shift;
        const parsed = parseShiftItem(item);
        if (parsed) events.push({ ...parsed, isOpenShift: true, isAllDay: false });
      }

      for (const tdo of (data.timesOff || data.TimesOff || data.timeOffRequests || [])) {
        const item = tdo.sharedTimeOff || tdo.timeOffItem || tdo.draftTimeOff || tdo;
        const startStr = item.startDateTime || item.StartDateTime;
        if (!startStr) continue;
        const start = new Date(startStr);
        const endStr = item.endDateTime || item.EndDateTime;
        const end = endStr ? new Date(endStr) : new Date(start.getTime() + 86400000);
        const summary = item.reason?.displayName || item.displayName || 'TDO';
        events.push({ summary, notes: '', start, end, isOpenShift: false, isAllDay: true });
      }

      console.info('[ShiftsExport] Parsed', events.length, 'events');
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
      const pad = (n) => String(n).padStart(2, '0');
      const dateOnly = (d) => d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
      lines.push('BEGIN:VEVENT');
      lines.push(`UID:${uid}`);
      lines.push(`DTSTAMP:${toICSDate(new Date())}`);
      if (ev.isAllDay) {
        lines.push(`DTSTART;VALUE=DATE:${dateOnly(ev.start)}`);
        lines.push(`DTEND;VALUE=DATE:${dateOnly(ev.end)}`);
      } else {
        lines.push(`DTSTART:${toICSDate(ev.start)}`);
        lines.push(`DTEND:${toICSDate(ev.end)}`);
      }
      lines.push(`SUMMARY:${summaryText.replace(/,/g, '\\,').replace(/\n/g, '\\n')}`);
      if (ev.notes) {
        lines.push(`DESCRIPTION:${ev.notes.replace(/\\/g, '\\\\').replace(/,/g, '\\,').replace(/;/g, '\\;').replace(/\n/g, '\\n')}`);
      }
      if (ev.isOpenShift) lines.push('CATEGORIES:Open Shift');
      lines.push('END:VEVENT');
    });

    lines.push('END:VCALENDAR');
    return lines.join('\r\n');
  }

  // ─── Expose API ───────────────────────────────────────────────────────────

  window.__shiftsExport = { scrape, generateICS, getTargetEndDate };

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.action === 'NAVIGATE_TO_SHIFTS') {
      navigateToShifts()
        .then(() => sendResponse({ success: true }))
        .catch((err) => sendResponse({ success: false, error: err.message }));
      return true;
    }

    if (msg.action === 'SCRAPE_AND_EXPORT') {
      if (!window.location.hostname.includes('flw.teams.cloud.microsoft')) return;
      scrape({ userName: msg.userName || null })
        .then((events) => {
          const ics = generateICS(events);
          const serializable = events.map((e) => ({
            summary:     e.summary,
            notes:       e.notes || '',
            startMs:     e.start.getTime(),
            endMs:       e.end.getTime(),
            isOpenShift: !!e.isOpenShift,
            isAllDay:    !!e.isAllDay,
          }));
          sendResponse({ success: true, ics, count: events.length, events: serializable });
        })
        .catch((err) => sendResponse({ success: false, error: err.message }));
      return true;
    }
  });
})();
