// offscreen.js — runs in an offscreen document to create blob URLs for downloads
// chrome.downloads is not available here; use anchor click on the blob URL instead.

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === 'OFFSCREEN_DOWNLOAD') {
    const blob = new Blob([msg.content], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = msg.filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    sendResponse({ ok: true });
  }
});
