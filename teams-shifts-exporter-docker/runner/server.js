'use strict';

const http = require('http');
const fs   = require('fs');
const { spawn } = require('child_process');

const PORT          = 8080;
const SETTINGS_PATH = '/data/settings.json';

const DEFAULT_SETTINGS = {
  importToOutlook:   false,
  importToiCloud:    false,
  icloudEmail:       '',
  icloudAppPassword: '',
  includeOpenShifts: true,
};

// ─── Settings helpers ─────────────────────────────────────────────────────────

function readSettings() {
  try {
    return { ...DEFAULT_SETTINGS, ...JSON.parse(fs.readFileSync(SETTINGS_PATH, 'utf8')) };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function writeSettings(patch) {
  const current = readSettings();
  const next = { ...current, ...patch };
  fs.writeFileSync(SETTINGS_PATH, JSON.stringify(next, null, 2));
  return next;
}

// ─── Sync runner ──────────────────────────────────────────────────────────────

let syncRunning = false;
let lastResult  = null;

function runSync() {
  if (syncRunning) return { started: false, reason: 'Sync already in progress' };
  syncRunning = true;
  console.log(`[server] Manual sync triggered at ${new Date().toISOString()}`);
  const child = spawn('node', ['/app/runner/sync.js'], { stdio: 'inherit' });
  child.on('close', (code) => {
    syncRunning = false;
    lastResult  = {
      success: code === 0,
      time:    new Date().toISOString(),
      message: code === 0 ? 'Sync completed successfully' : `Sync failed (exit code ${code})`,
    };
    console.log(`[server] Sync finished — ${lastResult.message}`);
  });
  return { started: true };
}

// ─── Request body helper ──────────────────────────────────────────────────────

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => { data += chunk; });
    req.on('end',  () => { try { resolve(JSON.parse(data)); } catch { reject(new Error('Invalid JSON')); } });
    req.on('error', reject);
  });
}

// ─── HTML dashboard ───────────────────────────────────────────────────────────

function renderDashboard() {
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Teams Shifts Sync</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:13px;
       background:#012a4a;color:#e8edf2;min-height:100vh}
  .header{background:#01579B;padding:14px 16px;border-bottom:2px solid #F48120}
  .header h1{font-size:15px;font-weight:600;color:#FBCE20}
  .header p{font-size:11px;color:#b8d4ea;margin-top:3px}
  .body{padding:14px 16px;max-width:340px;margin:0 auto}
  .status-box{background:#01466e;border-radius:6px;padding:10px 12px;
               margin-bottom:14px;border:1px solid #0277a8}
  .status-label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;
                 color:#7bacc8;margin-bottom:4px}
  .status-value{font-size:12px;color:#FBCE20}
  .status-value.none{color:#7bacc8}
  .sync-group{background:#01466e;border:1px solid #0277a8;border-radius:6px;
               margin-bottom:14px;overflow:hidden}
  .sync-group-label{font-size:10px;text-transform:uppercase;letter-spacing:.5px;
                     color:#7bacc8;padding:7px 12px 5px}
  .sync-row{display:flex;align-items:center;justify-content:space-between;
             padding:9px 12px;font-size:12px;color:#e8edf2;
             border-top:1px solid #0277a8;cursor:pointer}
  .sync-row:first-of-type{border-top:none}
  .sync-row input[type=checkbox]{width:16px;height:16px;accent-color:#F48120;cursor:pointer}
  .toggle-row{display:flex;align-items:center;justify-content:space-between;
               background:#01466e;border:1px solid #0277a8;border-radius:6px;
               padding:9px 12px;margin-bottom:14px;cursor:pointer;
               font-size:12px;color:#e8edf2}
  .toggle-row input[type=checkbox]{width:16px;height:16px;accent-color:#F48120;cursor:pointer}
  .creds-section{display:none;border-top:1px solid #0277a8;
                  background:#012040;padding:10px 12px 12px}
  .chevron{font-size:9px;color:#7bacc8;padding:2px 4px;border-radius:3px;
            transition:transform .15s;display:none;line-height:1}
  .chevron.open{transform:rotate(180deg)}
  .field{margin-bottom:10px}
  .field label{display:block;font-size:10px;text-transform:uppercase;
                letter-spacing:.5px;color:#7bacc8;margin-bottom:4px}
  .field input{width:100%;background:#01466e;border:1px solid #0277a8;
                border-radius:5px;padding:7px 10px;font-size:12px;
                color:#e8edf2;outline:none}
  .field input:focus{border-color:#F48120}
  .field-hint{font-size:10px;color:#5a8aab;margin-top:4px}
  button{width:100%;padding:10px;border:none;border-radius:6px;font-size:13px;
          font-weight:600;cursor:pointer;transition:opacity .15s}
  button:hover{opacity:.88}
  button:disabled{opacity:.4;cursor:not-allowed}
  #syncBtn{background:#F48120;color:#fff;margin-bottom:8px}
  #saveBtn{background:#1a6a9a;color:#e8edf2}
  #saveStatus{font-size:11px;margin-top:8px;min-height:16px;color:#FBCE20}
  #syncStatus{font-size:11px;margin-top:8px;min-height:16px;color:#F48120}
  #syncStatus.ok{color:#FBCE20}
  .section-gap{margin-bottom:14px}
  .divider{border:none;border-top:1px solid #01466e;margin:16px 0}
</style>
</head>
<body>
<div class="header">
  <h1>Teams Shifts Sync</h1>
  <p>Docker container dashboard</p>
</div>
<div class="body">

  <div class="status-box">
    <div class="status-label">Sync status</div>
    <div class="status-value none" id="syncStatusBox">Loading...</div>
  </div>

  <button id="syncBtn" onclick="triggerSync()">Force Sync Now</button>
  <div id="syncStatus"></div>

  <hr class="divider" />

  <!-- Settings -->
  <div class="sync-group">
    <div class="sync-group-label">Sync to</div>
    <label class="sync-row">
      <span>Outlook Calendar</span>
      <input type="checkbox" id="importToOutlook" />
    </label>
    <div class="sync-row" id="icloudRow">
      <label for="importToiCloud" style="flex:1;cursor:pointer">iCloud Calendar</label>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="chevron" id="icloudChevron">&#9660;</span>
        <input type="checkbox" id="importToiCloud" />
      </div>
    </div>
    <div class="creds-section" id="icloudCreds">
      <div class="field">
        <label>Apple ID (email)</label>
        <input type="email" id="icloudEmail" placeholder="you@icloud.com" autocomplete="off" />
      </div>
      <div class="field">
        <label>App-Specific Password</label>
        <input type="password" id="icloudAppPassword" placeholder="xxxx-xxxx-xxxx-xxxx" autocomplete="off" />
        <div class="field-hint">Generate at <strong>appleid.apple.com</strong> &rarr; Sign-In &amp; Security &rarr; App-Specific Passwords</div>
      </div>
    </div>
  </div>

  <label class="toggle-row">
    <span>Include open shifts</span>
    <input type="checkbox" id="includeOpenShifts" />
  </label>

  <button id="saveBtn" onclick="saveSettings()">Save Settings</button>
  <div id="saveStatus"></div>

</div>
<script>
// ─── Settings load ────────────────────────────────────────────────────────────
async function loadSettings() {
  const r = await fetch('/api/settings');
  const s = await r.json();
  document.getElementById('importToOutlook').checked  = !!s.importToOutlook;
  document.getElementById('importToiCloud').checked   = !!s.importToiCloud;
  document.getElementById('includeOpenShifts').checked = s.includeOpenShifts !== false;
  document.getElementById('icloudEmail').value        = s.icloudEmail || '';
  // Never pre-fill the password field — leave blank; only written if user types something
  toggleIcloudCreds(!!s.importToiCloud, /* expand */ !s.icloudEmail);
}

function toggleIcloudCreds(on, expand) {
  const chevron = document.getElementById('icloudChevron');
  const creds   = document.getElementById('icloudCreds');
  chevron.style.display = on ? 'inline' : 'none';
  if (!on) { creds.style.display = 'none'; chevron.classList.remove('open'); return; }
  if (expand) { creds.style.display = 'block'; chevron.classList.add('open'); }
}

document.getElementById('importToiCloud').addEventListener('change', function() {
  toggleIcloudCreds(this.checked, true);
});

document.getElementById('icloudChevron').addEventListener('click', (e) => {
  e.stopPropagation();
  const creds = document.getElementById('icloudCreds');
  const open  = creds.style.display !== 'none';
  creds.style.display = open ? 'none' : 'block';
  document.getElementById('icloudChevron').classList.toggle('open', !open);
});

async function saveSettings() {
  const saveBtn    = document.getElementById('saveBtn');
  const saveStatus = document.getElementById('saveStatus');
  saveBtn.disabled = true;
  saveBtn.textContent = 'Saving...';

  const body = {
    importToOutlook:   document.getElementById('importToOutlook').checked,
    importToiCloud:    document.getElementById('importToiCloud').checked,
    includeOpenShifts: document.getElementById('includeOpenShifts').checked,
    icloudEmail:       document.getElementById('icloudEmail').value.trim(),
  };
  const pw = document.getElementById('icloudAppPassword').value.trim();
  if (pw) body.icloudAppPassword = pw;

  const r = await fetch('/api/settings', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

  saveBtn.disabled = false;
  saveBtn.textContent = 'Save Settings';

  if (r.ok) {
    saveStatus.textContent = 'Settings saved.';
    saveStatus.style.color = '#FBCE20';
    document.getElementById('icloudAppPassword').value = '';
    setTimeout(() => { saveStatus.textContent = ''; }, 3000);
  } else {
    saveStatus.textContent = 'Save failed.';
    saveStatus.style.color = '#F48120';
  }
}

// ─── Sync controls ────────────────────────────────────────────────────────────
async function triggerSync() {
  const btn = document.getElementById('syncBtn');
  const el  = document.getElementById('syncStatus');
  btn.disabled = true;
  el.textContent = 'Starting...';
  el.className = '';
  const r = await fetch('/sync', { method: 'POST' });
  const d = await r.json();
  if (d.started) {
    el.textContent = 'Sync started — check container logs for progress.';
    pollStatus();
  } else {
    el.textContent = d.reason;
    btn.disabled = false;
  }
}

async function pollStatus() {
  const r = await fetch('/status');
  const d = await r.json();
  const box = document.getElementById('syncStatusBox');
  const el  = document.getElementById('syncStatus');
  const btn = document.getElementById('syncBtn');

  if (d.syncRunning) {
    box.textContent = 'Sync in progress...';
    box.className   = 'status-value';
    btn.disabled    = true;
    setTimeout(pollStatus, 4000);
  } else {
    btn.disabled = false;
    if (d.lastResult) {
      box.textContent = (d.lastResult.success ? '✓ ' : '✗ ') + d.lastResult.message;
      box.className   = 'status-value' + (d.lastResult.success ? '' : ' none');
      el.textContent  = 'Last run: ' + new Date(d.lastResult.time).toLocaleString();
      el.className    = d.lastResult.success ? 'ok' : '';
    } else {
      box.textContent = 'No sync run yet';
      box.className   = 'status-value none';
    }
  }
}

loadSettings();
pollStatus();
</script>
</body>
</html>`;
}

// ─── HTTP server ──────────────────────────────────────────────────────────────

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

  if (url.pathname === '/api/settings' && req.method === 'GET') {
    const s = readSettings();
    // Never send the password back to the browser
    const safe = { ...s, icloudAppPassword: s.icloudAppPassword ? '••••••••' : '' };
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(safe));
    return;
  }

  if (url.pathname === '/api/settings' && req.method === 'POST') {
    try {
      const body = await readBody(req);
      // If password comes back as the placeholder, don't overwrite the stored one
      if (body.icloudAppPassword === '••••••••') delete body.icloudAppPassword;
      const next = writeSettings(body);
      console.log('[server] Settings updated');
      const safe = { ...next, icloudAppPassword: next.icloudAppPassword ? '••••••••' : '' };
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(safe));
    } catch (err) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  if (url.pathname === '/sync' && (req.method === 'GET' || req.method === 'POST')) {
    const result = runSync();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(result));
    return;
  }

  if (url.pathname === '/status') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ syncRunning, lastResult }));
    return;
  }

  res.writeHead(200, { 'Content-Type': 'text/html' });
  res.end(renderDashboard());
});

server.listen(PORT, () => {
  console.log(`[server] Listening on port ${PORT} — http://<container-ip>:${PORT}`);
});

module.exports = { runSync, readSettings };
