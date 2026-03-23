// popup.js

function getTargetEndDate() {
  const now = new Date();
  const year = now.getFullYear();

  const candidates = [
    new Date(year, 1, 28),
    new Date(year, 7, 31),
    new Date(year + 1, 1, 28),
    new Date(year + 1, 7, 31),
  ];

  candidates.forEach((d, i) => {
    if (d.getMonth() === 1) {
      const y = d.getFullYear();
      if ((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0) {
        candidates[i] = new Date(y, 1, 29);
      }
    }
  });

  return candidates.filter((d) => d > now).sort((a, b) => a - b)[0];
}

function formatDate(ts) {
  if (!ts) return null;
  const d = new Date(ts);
  return d.toLocaleDateString(undefined, {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
}

// ─── Init ─────────────────────────────────────────────────────────────────────

const exportBtn = document.getElementById('exportBtn');
const downloadICSBtn = document.getElementById('downloadICSBtn');
const logEl = document.getElementById('log');
const lastExportEl = document.getElementById('lastExport');
const targetDateEl = document.getElementById('targetDate');
const importToOutlookEl = document.getElementById('importToOutlook');

// Show target date
const target = getTargetEndDate();
targetDateEl.textContent = target.toLocaleDateString(undefined, {
  month: 'long', day: 'numeric', year: 'numeric',
});

// Load last export status
const includeOpenShiftsEl = document.getElementById('includeOpenShifts');

chrome.storage.local.get(['lastExport', 'lastCount', 'importToOutlook', 'includeOpenShifts'], (data) => {
  if (data.lastExport) {
    lastExportEl.textContent =
      `${formatDate(data.lastExport)} — ${data.lastCount ?? '?'} shifts`;
    lastExportEl.classList.remove('none');
  }
  if (data.importToOutlook) {
    importToOutlookEl.checked = true;
  }
  // Default to true if never set
  includeOpenShiftsEl.checked = data.includeOpenShifts !== false;
});

// Save Outlook toggle
importToOutlookEl.addEventListener('change', () => {
  chrome.runtime.sendMessage({ action: 'SET_IMPORT_TO_OUTLOOK', value: importToOutlookEl.checked });
});

// Save open shifts toggle
includeOpenShiftsEl.addEventListener('change', () => {
  chrome.runtime.sendMessage({ action: 'SET_INCLUDE_OPEN_SHIFTS', value: includeOpenShiftsEl.checked });
});

// ─── Export Button ────────────────────────────────────────────────────────────

// ─── Clear & Re-import Button ────────────────────────────────────────────────

const clearReimportBtn = document.getElementById('clearReimportBtn');

clearReimportBtn.addEventListener('click', () => {
  clearReimportBtn.disabled = true;
  clearReimportBtn.textContent = 'Clearing & re-importing...';
  logEl.textContent = '';
  logEl.className = '';

  chrome.runtime.sendMessage({ action: 'CLEAR_AND_REIMPORT' }, (response) => {
    clearReimportBtn.disabled = false;
    clearReimportBtn.textContent = 'Clear & Re-import to Outlook';

    if (response && response.success) {
      logEl.textContent = `Done — cleared old events, imported ${response.count ?? '?'} shifts.`;
      logEl.className = 'ok';
    } else {
      logEl.textContent = `Error: ${response?.error || 'Unknown error'}`;
      logEl.className = '';
    }
  });
});

// ─── Export Button ────────────────────────────────────────────────────────────

downloadICSBtn.addEventListener('click', () => {
  downloadICSBtn.disabled = true;
  downloadICSBtn.textContent = 'Downloading...';
  logEl.textContent = '';
  logEl.className = '';

  chrome.runtime.sendMessage({ action: 'DOWNLOAD_ICS' }, (response) => {
    downloadICSBtn.disabled = false;
    downloadICSBtn.textContent = 'Download ICS File';

    if (response && response.success) {
      logEl.textContent = 'ICS file saved to Downloads.';
      logEl.className = 'ok';
    } else {
      logEl.textContent = `Error: ${response?.error || 'Unknown error'}`;
      logEl.className = '';
    }
  });
});

exportBtn.addEventListener('click', () => {
  exportBtn.disabled = true;
  exportBtn.textContent = 'Syncing...';
  logEl.textContent = '';
  logEl.className = '';

  chrome.runtime.sendMessage({ action: 'EXPORT_NOW' }, (response) => {
    exportBtn.disabled = false;
    exportBtn.textContent = 'Sync Shifts';

    if (response && response.success) {
      logEl.textContent = `Done — ${response.count} shifts synced.`;
      logEl.className = 'ok';

      // Refresh status
      chrome.runtime.sendMessage({ action: 'GET_STATUS' }, (data) => {
        if (data && data.lastExport) {
          lastExportEl.textContent =
            `${formatDate(data.lastExport)} — ${data.lastCount ?? '?'} shifts`;
          lastExportEl.classList.remove('none');
        }
      });
    } else {
      logEl.textContent = `Error: ${response?.error || 'Unknown error'}`;
      logEl.className = '';
    }
  });
});
