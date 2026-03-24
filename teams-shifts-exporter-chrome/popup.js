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
const importToiCloudEl = document.getElementById('importToiCloud');
const icloudCredsSectionEl = document.getElementById('icloudCredsSection');
const icloudCredsFieldsEl = document.getElementById('icloudCredsFields');
const icloudCredsChevronEl = document.getElementById('icloudCredsChevron');
const icloudCredsSummaryEl = document.getElementById('icloudCredsSummary');
const icloudEmailEl = document.getElementById('icloudEmail');
const icloudAppPasswordEl = document.getElementById('icloudAppPassword');
const saveICloudCredsBtn = document.getElementById('saveICloudCredsBtn');
const icloudCredsStatusEl = document.getElementById('icloudCredsStatus');

// ─── Progress Polling ─────────────────────────────────────────────────────────

const progressSectionEl = document.getElementById('progressSection');
const progressLabelEl = document.getElementById('progressLabel');
const progressFillEl = document.getElementById('progressFill');
const cancelSyncBtn = document.getElementById('cancelSyncBtn');

let progressInterval = null;

function updateProgressUI(step, percent) {
  progressSectionEl.style.display = 'block';
  progressLabelEl.textContent = step || 'Syncing...';
  progressFillEl.style.width = `${percent || 0}%`;
}

function hideProgress() {
  progressSectionEl.style.display = 'none';
  progressFillEl.style.width = '0%';
  cancelSyncBtn.disabled = false;
  cancelSyncBtn.textContent = 'Cancel';
  if (progressInterval) {
    clearInterval(progressInterval);
    progressInterval = null;
  }
}

cancelSyncBtn.addEventListener('click', () => {
  cancelSyncBtn.disabled = true;
  cancelSyncBtn.textContent = 'Cancelling…';
  chrome.runtime.sendMessage({ action: 'CANCEL_SYNC' });
});

function startProgressPolling() {
  if (progressInterval) return;
  progressInterval = setInterval(() => {
    chrome.storage.local.get(['syncRunning', 'syncStep', 'syncPercent'], (data) => {
      if (data.syncRunning) {
        updateProgressUI(data.syncStep, data.syncPercent);
      } else {
        hideProgress();
      }
    });
  }, 500);
}

// On popup open, check if a sync is already running
chrome.storage.local.get(['syncRunning', 'syncStep', 'syncPercent'], (data) => {
  if (data.syncRunning) {
    updateProgressUI(data.syncStep, data.syncPercent);
    startProgressPolling();
  }
});

// Show target date
const target = getTargetEndDate();
targetDateEl.textContent = target.toLocaleDateString(undefined, {
  month: 'long', day: 'numeric', year: 'numeric',
});

// Load last export status
const includeOpenShiftsEl = document.getElementById('includeOpenShifts');

function setICloudCredsCollapsed(collapsed, email) {
  icloudCredsFieldsEl.style.display = collapsed ? 'none' : 'block';
  icloudCredsChevronEl.classList.toggle('open', !collapsed);
  icloudCredsSummaryEl.textContent = collapsed && email ? email : '';
  icloudCredsSummaryEl.style.display = collapsed && email ? 'block' : 'none';
}

chrome.storage.local.get(
  ['lastExport', 'lastCount', 'importToOutlook', 'includeOpenShifts', 'importToiCloud', 'icloudEmail', 'icloudCredsSet'],
  (data) => {
    if (data.lastExport) {
      lastExportEl.textContent =
        `${formatDate(data.lastExport)} — ${data.lastCount ?? '?'} shifts`;
      lastExportEl.classList.remove('none');
    }
    if (data.importToOutlook) {
      importToOutlookEl.checked = true;
    }
    includeOpenShiftsEl.checked = data.includeOpenShifts !== false;

    if (data.importToiCloud) {
      importToiCloudEl.checked = true;
      icloudCredsSectionEl.style.display = 'block';
      icloudCredsChevronEl.style.display = 'inline';
    }
    if (data.icloudEmail) {
      icloudEmailEl.value = data.icloudEmail;
    }
    // Collapse the credentials fields if already saved
    if (data.icloudCredsSet) {
      setICloudCredsCollapsed(true, data.icloudEmail);
    }
  }
);

// Save Outlook toggle
importToOutlookEl.addEventListener('change', () => {
  chrome.runtime.sendMessage({ action: 'SET_IMPORT_TO_OUTLOOK', value: importToOutlookEl.checked });
});

// iCloud toggle — show/hide credential section and chevron
importToiCloudEl.addEventListener('change', () => {
  const on = importToiCloudEl.checked;
  icloudCredsSectionEl.style.display = on ? 'block' : 'none';
  icloudCredsChevronEl.style.display = on ? 'inline' : 'none';
  // When turning on, expand fields if no credentials saved yet
  if (on) {
    chrome.storage.local.get(['icloudCredsSet', 'icloudEmail'], (data) => {
      setICloudCredsCollapsed(!!data.icloudCredsSet, data.icloudEmail);
    });
  }
  chrome.storage.local.set({ importToiCloud: on });
});

// Chevron click — expand/collapse credential fields
icloudCredsChevronEl.addEventListener('click', (e) => {
  e.stopPropagation();
  const isCollapsed = icloudCredsFieldsEl.style.display === 'none';
  chrome.storage.local.get('icloudEmail', (data) => {
    setICloudCredsCollapsed(!isCollapsed, data.icloudEmail);
  });
});

// Save iCloud credentials and auto-collapse
saveICloudCredsBtn.addEventListener('click', () => {
  const email = icloudEmailEl.value.trim();
  const password = icloudAppPasswordEl.value.trim();
  if (!email || !password) {
    icloudCredsStatusEl.textContent = 'Enter both Apple ID and app-specific password.';
    icloudCredsStatusEl.style.color = '#F48120';
    return;
  }
  chrome.storage.local.set({ icloudEmail: email, icloudAppPassword: password, icloudCredsSet: true }, () => {
    icloudAppPasswordEl.value = '';
    icloudCredsStatusEl.textContent = '';
    setICloudCredsCollapsed(true, email);
  });
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
  startProgressPolling();
  logEl.textContent = '';
  logEl.className = '';

  chrome.runtime.sendMessage({ action: 'CLEAR_AND_REIMPORT' }, (response) => {
    clearReimportBtn.disabled = false;
    clearReimportBtn.textContent = 'Clear & Re-import Selected';

    if (response && response.success) {
      let msg = `Done — cleared old events, imported ${response.count ?? '?'} shifts.`;
      if (response.icloudResult) {
        msg += response.icloudResult.success
          ? ' iCloud cleared & re-synced.'
          : ` iCloud error: ${response.icloudResult.error}`;
      }
      logEl.textContent = msg;
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
  startProgressPolling();
  logEl.textContent = '';
  logEl.className = '';

  chrome.runtime.sendMessage({ action: 'EXPORT_NOW' }, (response) => {
    exportBtn.disabled = false;
    exportBtn.textContent = 'Sync Shifts';

    if (response && response.success) {
      let msg = `Done — ${response.count} shifts synced.`;
      if (response.icloudResult) {
        msg += response.icloudResult.success
          ? ' iCloud updated.'
          : ` iCloud error: ${response.icloudResult.error}`;
      }
      logEl.textContent = msg;
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
