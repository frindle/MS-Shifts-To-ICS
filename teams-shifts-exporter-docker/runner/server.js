'use strict';

const http = require('http');
const { spawn } = require('child_process');

const PORT = 8080;

let syncRunning = false;
let lastResult = null; // { success, time, message }

function runSync() {
  if (syncRunning) return { started: false, reason: 'Sync already in progress' };

  syncRunning = true;
  console.log(`[server] Manual sync triggered at ${new Date().toISOString()}`);

  const child = spawn('node', ['/app/runner/sync.js'], { stdio: 'inherit' });
  child.on('close', (code) => {
    syncRunning = false;
    lastResult = {
      success: code === 0,
      time: new Date().toISOString(),
      message: code === 0 ? 'Sync completed successfully' : `Sync failed (exit code ${code})`,
    };
    console.log(`[server] Sync finished — ${lastResult.message}`);
  });

  return { started: true };
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);

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

  // Root — simple dashboard
  res.writeHead(200, { 'Content-Type': 'text/html' });
  res.end(`<!DOCTYPE html>
<html>
<head><title>Teams Shifts Sync</title>
<style>body{font-family:sans-serif;max-width:480px;margin:60px auto;padding:0 20px}
button{padding:12px 24px;font-size:16px;cursor:pointer;border:none;border-radius:6px;
background:#0277a8;color:#fff}button:hover{background:#01579B}
#status{margin-top:20px;color:#555;font-size:14px}</style></head>
<body>
<h2>Teams Shifts Sync</h2>
<button onclick="triggerSync()">Force Sync Now</button>
<div id="status">Loading...</div>
<script>
async function triggerSync() {
  document.getElementById('status').textContent = 'Starting sync...';
  const r = await fetch('/sync', { method: 'POST' });
  const d = await r.json();
  document.getElementById('status').textContent = d.started
    ? 'Sync started — check container logs for progress.'
    : 'Already running: ' + d.reason;
  setTimeout(pollStatus, 3000);
}
async function pollStatus() {
  const r = await fetch('/status');
  const d = await r.json();
  const last = d.lastResult;
  document.getElementById('status').textContent = d.syncRunning
    ? 'Sync in progress...'
    : last ? (last.success ? '✓ ' : '✗ ') + last.message + ' at ' + last.time : 'No sync run yet.';
  if (d.syncRunning) setTimeout(pollStatus, 5000);
}
pollStatus();
</script>
</body></html>`);
});

server.listen(PORT, () => {
  console.log(`[server] Listening on port ${PORT} — http://<container-ip>:${PORT}`);
});

module.exports = { runSync };
